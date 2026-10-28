# Conciliación SRI vs Odoo

Aplicación web para contabilidad que concilia los comprobantes electrónicos del portal del SRI
(Ecuador) contra Odoo 14 y muestra el resultado en un tablero en vivo.

- **Compras**: facturas, notas de crédito, notas de débito y retenciones **recibidas**.
- **Ventas**: facturas, notas de crédito, notas de débito, liquidaciones y retenciones **emitidas**.

Odoo se lee por XML-RPC en **solo lectura**; la app nunca escribe en Odoo.

## Cómo se usa

1. Entrar a la app con usuario y contraseña y elegir **Compras** o **Ventas**.
2. Soltar los TXT que exporta el portal del SRI ("Exportar Txt"). Compras admite hasta 5 TXT por
   sesión y Ventas hasta 400. El tipo de cada TXT se detecta por su contenido, no por el nombre.
3. Mientras se sube, la pantalla recalcula contra Odoo y muestra:
   - una tarjeta por tipo de comprobante ("N de M comprobantes del SRI coinciden");
   - la tabla de excepciones, con filtros por tipo y estado y búsqueda por número de documento;
   - las advertencias (falta un tipo, período largo, XML no disponible, etc.).
4. Descargar:
   - **Observaciones (Excel)**: solo lo que hay que corregir, una hoja por estado y una columna
     "Qué corregir en Odoo". Es el archivo que se reparte a los demás departamentos.
   - **Completo (Excel)**: el detalle con fórmulas, todas las hojas.
5. Cuando se corrija en Odoo, **volver a subir los TXT**. Cada resultado es una fotografía del momento:
   no se guarda historial y lo corregido aparece como "Coincide".

Las sesiones se borran solas 2 horas después de la última actividad.

## Estados

| Estado | Qué significa | ¿Requiere acción? |
|---|---|---|
| Coincide | Mismos valores y datos en SRI y Odoo | No |
| Solo en SRI | Autorizado en el SRI y sin registro en Odoo | Sí |
| Solo en Odoo | En Odoo y sin autorización en el SRI | Sí |
| Diferencia de valor | Total o base distintos | Sí |
| Diferencia de datos | Mismos valores, pero clave, fecha o marca electrónica distintas | Sí |
| Factura incorrecta en Odoo | Retención con la factura sustento equivocada | Sí |
| Odoo anulado, SRI autorizado | Cancelado en Odoo y vigente en el SRI | Sí |
| Duplicado en Odoo | Dos registros de Odoo para el mismo documento | Sí |
| XML no disponible | No se pudo bajar el XML de la retención; reintentar más tarde | Sí |
| Informativo | Borrador, cancelado sin autorización o no electrónico | No |

El detalle de reglas de cruce está en [`CLAUDE.md`](CLAUDE.md).

## Despliegue (Easypanel / Docker)

La imagen se construye con el `Dockerfile` del repositorio y arranca con:

```
gunicorn app.web.main:app -w 2 -k uvicorn.workers.UvicornWorker --bind 0.0.0.0:8000 --timeout 30
```

**Variables de entorno** (en el servicio, nunca en la imagen):

| Variable | Uso |
|---|---|
| `ODOO_URL`, `ODOO_DB`, `ODOO_USER`, `ODOO_PASSWORD` | Acceso de lectura a Odoo |
| `SECRET_KEY` | Firma de la cookie de sesión. **Obligatoria**: si falta, cada worker genera una distinta y el inicio de sesión falla de forma intermitente |
| `SESSION_HOURS` | Duración de la sesión web (por defecto 8) |

**Volúmenes persistentes** (montar en `/app/...`):

| Ruta | Contenido |
|---|---|
| `/app/data` | `users.yaml` (usuarios) |
| `/app/cache` | `xml/`: caché compartido de XML de retenciones |
| `/app/jobs` | Sesiones de trabajo (temporales) |

**Primer usuario** (consola del servicio en Easypanel):

```bash
python scripts/setup_usuario.py add <usuario>      # también: remove, list
```

**Después de desplegar**, comprobar que carga `/login`, entrar, subir un TXT y ver aparecer las
tarjetas. Si el servicio no arranca, revisar los logs: un `ModuleNotFoundError` o un error con
`app:app` indica un comando de inicio antiguo; debe ser `app.web.main:app`.

Notas:

- La app rechaza por sí misma lo que pase de 20 MB por archivo o 200 MB por envío (HTTP 413). Solo si
  se pone un nginx propio delante hace falta `client_max_body_size` de unos 210 MB.
- `deploy/install.sh` es la instalación manual en un VPS Ubuntu. Hoy **no** incluye los archivos que
  menciona (`deploy/nginx.conf`, `deploy/sri-conciliacion.service`), así que no sirve tal cual; en
  Easypanel no se usa.

## Desarrollo local

```bash
pip install -r requirements.txt
python -m pytest tests -q
# .env local (ignorado por git) con las variables ODOO_*; opcional SECRET_KEY
uvicorn app.web.main:app --reload
```

LibreOffice es opcional: si está instalado, recalcula las fórmulas del Excel completo; si no, la
corrida lo avisa y el Excel se calcula al abrirlo.

## Estructura

```
app/
  web/          FastAPI, login, plantillas y estáticos del tablero
  sesion.py     sesión de trabajo: TXT subidos, recálculo cancelable, salidas atómicas
  job_manager.py  carpetas jobs/<id>/ y limpieza
  sri/          lectura de TXT/XML del SRI, descarga de XML, caché
  odoo/         cliente XML-RPC de solo lectura
  compras/      conciliación de recibidos
  ventas/       conciliación de emitidos
  retenciones/  cruce de retenciones (compartido)
  reportes/     cifras del tablero y Excel de observaciones
scripts/        herramientas de línea de comandos (usuarios, conciliaciones sueltas)
tests/          espeja app/
docs/superpowers/  diseños y planes de cada funcionalidad
```

## Seguridad

- Odoo solo se lee; los errores nunca incluyen credenciales.
- Todas las rutas de datos exigen sesión (401 sin ella). Los Excel protegen el texto contra inyección
  de fórmulas.
- Límites de subida: 20 MB por archivo, 200 MB por envío, solo `.txt`.
- Limitación conocida: cualquier usuario con sesión que conozca el id de una sesión de trabajo (8
  caracteres hexadecimales) puede ver su tablero y descargar sus Excel. Los usuarios son del equipo de
  contabilidad; si eso deja de ser cierto, hay que asociar cada sesión a su dueño.

## Pendientes conocidos

- Analizar las ~198 retenciones "Solo en SRI" de agosto (¿faltan en Odoo o son de otro período?).
- Aislar las sesiones por usuario (ver Seguridad).
- Menores: las filas "Solo en Odoo" de Compras muestran el nombre del asiento en vez del número del
  proveedor; las retenciones de Ventas muestran el RUC propio en la columna RUC; un TXT de retenciones
  sin fechas válidas hace fallar todo el recálculo sin nombrar el archivo; accesibilidad por teclado de
  los encabezados ordenables.
