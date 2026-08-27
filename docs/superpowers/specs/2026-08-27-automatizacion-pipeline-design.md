# Diseño: Automatización del Pipeline de Conciliación SRI

**Fecha:** 2026-08-27
**Proyecto:** sri_conciliacion
**Estado:** Aprobado

---

## Contexto

El pipeline actual (5 scripts Python) corre manualmente en la máquina local de quien hace el análisis. Se quiere convertirlo en un servicio web desplegado en un VPS para que cualquier persona autorizada pueda:

1. Subir los dos archivos de entrada (`Recibidos.txt` y `reten_odoo.xlsx`)
2. Esperar mientras el servidor procesa
3. Descargar el Excel de conciliación resultante

Ambos archivos de entrada siguen siendo descargados manualmente (SRI requiere login humano; Odoo export es manual por decisión).

---

## Arquitectura

```
Internet → nginx (443, TLS Let's Encrypt)
               └── proxy_pass → gunicorn :8000 (localhost)
                                    └── FastAPI app.py
                                         └── hilos Python (background jobs)
                                              ├── parse_recibidos_txt.py
                                              ├── descargar_comprobantes_sri.py
                                              ├── build_workbook_final.py
                                              └── conciliar_odoo.py
```

### Estructura de carpetas en el VPS

```
/opt/sri_conciliacion/
├── app.py                    # FastAPI — único archivo de servidor
├── users.yaml                # credenciales (bcrypt hashes)
├── sri_session.py            # HTTP adapter con TLS correcto para el SRI
├── parse_recibidos_txt.py
├── descargar_comprobantes_sri.py
├── parse_comprobante_xml.py
├── build_workbook_final.py
├── conciliar_odoo.py
├── templates/
│   └── index.html            # UI completa (HTML + CSS + JS vanilla)
├── cache/
│   └── xml/                  # XMLs por clave — persisten entre períodos
└── jobs/
    └── {job_id}/
        ├── Recibidos.txt
        ├── reten_odoo.xlsx
        ├── claves.txt
        └── resultado.xlsx
```

---

## Autenticación

- **Mecanismo:** HTTP Basic Auth validado en cada request contra `users.yaml`
- **Contraseñas:** almacenadas como hashes bcrypt (nunca en texto plano)
- **Administración:** editar `users.yaml` + `systemctl restart sri-conciliacion`
- **HTTPS:** obligatorio — nginx termina TLS con Let's Encrypt (certbot)

```yaml
# users.yaml
users:
  - username: ana
    password_hash: "$2b$12$..."
  - username: carlos
    password_hash: "$2b$12$..."
```

Sin sesiones, sin cookies, sin tokens que rotar. Agregar o revocar usuarios es editar el archivo.

---

## Endpoints FastAPI

| Método | Ruta | Auth | Descripción |
|---|---|---|---|
| `GET /` | Página principal | ✅ | Devuelve `index.html` |
| `POST /upload` | Sube archivos | ✅ | Guarda archivos, lanza job → `{"job_id": "abc123"}` |
| `GET /status/{job_id}` | Estado del job | ✅ | `{estado, progreso, resumen}` |
| `GET /download/{job_id}` | Descarga Excel | ✅ | Stream del archivo resultado |

---

## Pipeline y caché de XMLs

### Flujo de un job

| Fase | Estado reportado |
|---|---|
| Archivos recibidos, hilo lanzado | `pending` |
| Leyendo TXT | `running: leyendo TXT (N claves)` |
| Descargando XMLs | `running: descargando XMLs (K/N)` — actualiza cada 10 |
| Procesando XMLs → Excel | `running: procesando XMLs` |
| Conciliando con Odoo | `running: conciliando con Odoo` |
| Terminado | `done` + resumen numérico |
| Error | `error: <mensaje>` |

### Caché de XMLs

- Ubicación: `cache/xml/<clave_acceso>.xml`
- Compartida entre todos los jobs y períodos
- Los XMLs son inmutables (clave autorizada = contenido fijo)
- La caché nunca se invalida, solo crece
- **Impacto en tiempo de ejecución:**
  - Primera corrida del período: ~5-8 minutos
  - Corridas subsiguientes del mismo período: ~30 segundos

### Estado de jobs en disco

El estado de cada job se persiste en `jobs/{job_id}/status.json` en lugar de memoria.
Esto es necesario porque gunicorn corre 2 workers (procesos separados): el `/upload`
puede llegar al worker 1 y el `/status` al worker 2 — si el estado fuera en memoria,
worker 2 no sabría nada del job.

```json
{
  "estado": "running",
  "mensaje": "Descargando XMLs del SRI (143/578)",
  "progreso": 0.25,
  "resumen": null,
  "error": null
}
```

- El hilo de pipeline escribe `status.json` al avanzar cada fase
- `/status/{id}` lee el archivo y lo devuelve — sin necesidad de compartir memoria entre workers
- Los jobs se eliminan de disco 2 horas después de completarse
- Si el servidor se reinicia, los jobs activos se marcan como `error` al arrancar (el hilo murió)
- Decisión aceptada: la corrida tarda minutos, no horas — perder un job en progreso es raro

---

## Interfaz de usuario

Una sola página HTML estática servida por FastAPI. Sin frameworks frontend (HTML + CSS + JavaScript vanilla).

### Estados de la UI

**1. Formulario inicial**
```
[ Recibidos.txt      ] [Seleccionar archivo]
[ reten_odoo.xlsx    ] [Seleccionar archivo]

           [ Procesar ]
```

**2. Procesando** (polling cada 5 segundos a `/status/{id}`)
```
⏳ Procesando...

Descargando XMLs del SRI  (143/578)
████████░░░░░░░░░  25%

Tiempo transcurrido: 1m 23s
```

**3. Listo**
```
✅ Listo — Agosto 2026

578  retenciones SRI
620  líneas coinciden
317  solo en SRI  ·  $17,694 pendientes
  8  solo en Odoo  ·  revisar
  1  factura incorrecta en Odoo

[ Descargar Excel ]
```

**4. Error**
```
❌ Error al procesar

<mensaje de error>

[ Intentar de nuevo ]
```

### Comportamiento del polling

- JavaScript llama a `GET /status/{id}` cada 5 segundos
- Al recibir `done`: muestra resumen + botón de descarga
- Al recibir `error`: muestra mensaje + botón de reintentar (vuelve al formulario)
- Barra de progreso calculada por fase: TXT=5%, descarga XMLs=5-80%, procesamiento=80-95%, conciliación=95-100%

---

## Despliegue en VPS

### Requisitos

- Ubuntu 22.04 LTS
- Python 3.10+
- nginx + certbot
- Dominio DNS apuntando al VPS

### Dependencias Python adicionales

```bash
pip install fastapi uvicorn gunicorn bcrypt pyyaml python-multipart
# (pandas, openpyxl, requests ya existían)
```

### systemd service

```ini
# /etc/systemd/system/sri-conciliacion.service
[Unit]
Description=SRI Conciliación Web App
After=network.target

[Service]
WorkingDirectory=/opt/sri_conciliacion
ExecStart=/opt/sri_conciliacion/venv/bin/gunicorn app:app \
    -w 2 -k uvicorn.workers.UvicornWorker --bind 127.0.0.1:8000
Restart=on-failure
User=sri

[Install]
WantedBy=multi-user.target
```

2 workers gunicorn = hasta 2 jobs concurrentes. El caché de XMLs en disco es compartido entre workers de forma segura (escrituras son atómicas por archivo).

### nginx config (fragmento)

```nginx
server {
    listen 443 ssl;
    server_name conciliacion.tuempresa.com;

    ssl_certificate     /etc/letsencrypt/live/conciliacion.tuempresa.com/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/conciliacion.tuempresa.com/privkey.pem;

    client_max_body_size 20M;   # para los archivos subidos

    location / {
        proxy_pass         http://127.0.0.1:8000;
        proxy_read_timeout 30s;
    }
}
server {
    listen 80;
    server_name conciliacion.tuempresa.com;
    return 301 https://$host$request_uri;
}
```

### Proceso de actualización de scripts

```bash
# Copiar scripts actualizados al VPS
scp *.py sri@vps:/opt/sri_conciliacion/
systemctl restart sri-conciliacion
```

---

## Lo que queda fuera del alcance

- Automatización del export de Odoo (se mantiene manual)
- Historial de análisis anteriores (sin base de datos)
- Notificaciones por correo o Slack al terminar
- Panel de administración de usuarios (se edita `users.yaml` a mano)
- Múltiples empresas / RUCs (diseñado para `2390000831001`)

Cualquiera de estos puede añadirse en el futuro sin cambiar la arquitectura base.

---

## Nuevos archivos a crear

| Archivo | Descripción |
|---|---|
| `app.py` | Servidor FastAPI completo |
| `templates/index.html` | UI single-page |
| `users.yaml` | Credenciales (a generar con script de setup) |
| `setup_usuario.py` | Script CLI para añadir/cambiar usuarios (genera bcrypt hash) |
| `deploy/sri-conciliacion.service` | Plantilla systemd |
| `deploy/nginx.conf` | Plantilla nginx |
| `deploy/install.sh` | Script de instalación en VPS limpio |
