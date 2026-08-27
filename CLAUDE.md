# Conciliación de Comprobantes de Retención SRI (Ecuador)

## Qué hace este proyecto

Automatiza el proceso mensual que hoy es manual: descargar comprobantes de
retención del portal SRI en línea uno por uno, y cotejarlos contra el ERP
(Odoo) para detectar inconsistencias.

RUC de la empresa (receptor de las retenciones): `2390000831001`

Flujo completo:

```
Portal SRI (manual, 1 clic)          Automático (este proyecto)
──────────────────────────           ─────────────────────────────────────
"Comprobantes electrónicos    →      1. parse_recibidos_txt.py
recibidos" → filtrar período →          lee el TXT, genera claves.txt
Exportar Txt
                                      2. descargar_comprobantes_sri.py
                                         baja el XML de cada comprobante
                                         por su clave de acceso, vía el
                                         web service PÚBLICO del SRI
                                         (sin login, sin captcha)
                                      3. parse_comprobante_xml.py
                                         extrae cabecera + detalle de
                                         cada XML (incl. numDocSustento)
                                      4. build_workbook_final.py
                                         arma el Excel de conciliación

Exportar retenciones de Odoo  →      5. conciliar_odoo.py
(reten_odoo.xlsx)                        cruza SRI vs Odoo y añade
                                         4 hojas de resultado al Excel
```

## Archivos del proyecto

| Archivo | Rol |
|---|---|
| `parse_recibidos_txt.py` | Lee el TXT oficial exportado del portal (`RUC_Recibidos.txt`). Codificación **ISO-8859-1**, delimitado por **tabulador**. Genera la lista de claves de acceso del período. |
| `descargar_comprobantes_sri.py` | Descarga el XML de cada comprobante por clave de acceso, usando el WS público `AutorizacionComprobantesOffline`. Usa `sri_session()` para TLS completo. |
| `sri_session.py` | `HTTPAdapter` personalizado que maneja la redirección del SRI a la IP `181.113.227.222` con TLS completo (sin `verify=False`). Importar con `from sri_session import sri_session`. |
| `parse_comprobante_xml.py` | Parsea cada XML descargado: cabecera + una fila por cada línea de retención. Soporta **v1.0.0** (`<impuestos>/<impuesto>`) y **v2.0.0** (`<docsSustento>/<docSustento>/<retenciones>/<retencion>`). |
| `build_workbook_final.py` | Genera `Comprobantes_Retencion_<periodo>_Final.xlsx` con 5 hojas: `SRI_Listado`, `SRI_Detalle_Retenciones` (incluye columna `N_Retencion` y `Num_Doc_Sustento` formateados con guiones), `SRI_Por_Doc_Sustento`, `ERP_Compras`, `Conciliacion`. |
| `conciliar_odoo.py` | Cruza `SRI_Detalle_Retenciones` contra `reten_odoo.xlsx` y añade 4 hojas al Excel: `Conciliacion_Odoo` (resumen), `Solo_en_SRI`, `Solo_en_Odoo`, `Diferencias_Valor`, `Factura_Incorrecta_Odoo`. |
| `reten_odoo.xlsx` | Reporte exportado desde Odoo (retenciones de compras del período). Se reemplaza cada vez que se corra el análisis. |
| `xml_samples/` | Carpeta donde caen los XML descargados (nombrados `<clave_acceso>.xml`). El archivo `Comprobante_de_Reten.xml` es una muestra manual; el glob de parseo usa `[0-9]*.xml` para excluirlo. |
| `claves_<periodo>.txt` | Lista de claves de acceso del período, generada por el paso 1. |
| `claves_pendientes.txt` | Generado automáticamente con las claves que fallaron en la primera pasada, para reintento. |

## Comandos

```bash
# 1) Listar comprobantes del mes y generar claves.txt
python parse_recibidos_txt.py 2390000831001_Recibidos.txt claves_agosto2026.txt

# 2) Descargar los XMLs
python descargar_comprobantes_sri.py claves_agosto2026.txt xml_samples/

# 2b) Si quedaron claves fallidas, reintentarlas con verify=False
#     (el SRI a veces redirige a una IP sin cert válido)
python -c "
import requests, re, os, time, urllib3
urllib3.disable_warnings()
claves = [l.strip() for l in open('claves_pendientes.txt') if l.strip()]
ENDPOINT = 'https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline'
SOAP = open('...').read()  # usar el template de descargar_comprobantes_sri.py
# ... (ver descargar_comprobantes_sri.py, añadir verify=False al requests.post)
"

# 3) Armar el Excel final
python build_workbook_final.py

# 4) Conciliar contra Odoo (requiere reten_odoo.xlsx en la misma carpeta)
python conciliar_odoo.py
```

## Web service SRI — endpoints confirmados (agosto 2026)

| Endpoint | Estado |
|---|---|
| `https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantes` | ❌ Cierra conexión sin respuesta |
| `https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline` | ✅ Funciona — usar este |

Header obligatorio: `Connection: close` (evita el error `RemoteDisconnected`).

Algunos requests son redirigidos a la IP `181.113.227.222` y fallan por SSL
(`certificate is not valid for IP address`). Es un defecto de infraestructura
del SRI: esa IP sirve el cert de `srienlinea.sri.gob.ec` (válido, firmado por
DigiCert), pero la IP no aparece en las SANs del cert.

**Solución implementada** (`sri_session.py`): `HTTPAdapter` personalizado que al
detectar esa IP usa `assert_hostname='srienlinea.sri.gob.ec'` y
`server_hostname='srienlinea.sri.gob.ec'` en urllib3. Resultado: la cadena de CA
**y** el hostname se verifican completamente — no se deshabilita nada.

`descargar_comprobantes_sri.py` importa `sri_session()` en lugar de usar
`requests` directamente; no existe `verify=False` en el codebase.

## Formato de los XML del SRI — dos versiones del esquema

El SRI usa dos versiones del `<comprobanteRetencion>`:

| Versión | Estructura de impuestos |
|---|---|
| **v1.0.0** | `<impuestos>/<impuesto>` — `numDocSustento` dentro de cada `<impuesto>` |
| **v2.0.0** | `<docsSustento>/<docSustento>/<retenciones>/<retencion>` — `numDocSustento` en `<docSustento>`, retenciones anidadas |

En el lote de agosto 2026: **112 comprobantes v1.0** y **466 comprobantes v2.0**.
`parse_comprobante_xml.py` detecta la versión por el atributo `version` del nodo raíz.

## Formato de campos clave

- **`N_Retencion`**: `Establecimiento-PuntoEmision-Secuencial` → `001-200-000004340`
- **`Num_Doc_Sustento`**: `primeros3-sig3-resto` → `001-003-000088766` (split 3-3-9, estándar SRI)
- Estos mismos formatos usa Odoo en `No. de retención` y `Factura`.

## Conciliación SRI vs Odoo — lógica y estados

Clave de cruce: `(N° Retencion, Factura, Porcentaje)` — identifica una línea de retención única.

| Estado | Significado |
|---|---|
| **Coincide** | Línea presente en SRI y Odoo con mismo valor |
| **Solo en SRI** | Retención autorizada en SRI pero no registrada en Odoo |
| **Solo en Odoo** | Registrada en Odoo pero no encontrada en SRI (N° incorrecto, otro período, o no autorizada) |
| **Factura incorrecta en Odoo** | N° de retención correcto pero el número de factura sustento difiere — típicamente error en los primeros 3 dígitos (establecimiento) |
| **Diferencia de valor** | Cruza por clave pero el valor retenido difiere en más de $0.02 |

## Datos ya confirmados sobre el formato del SRI

- El TXT de "Exportar Txt" trae `VALOR_SIN_IMPUESTOS`, `IVA` e
  `IMPORTE_TOTAL` **siempre vacíos** para comprobantes de tipo
  "Comprobante de Retención" (confirmado sobre 578 filas reales de
  agosto/2026).
- `NUMERO_DOCUMENTO_MODIFICADO` del TXT es un ID interno del SRI, **no**
  es el número de la factura sustento — no sirve para conciliar.
- Los montos reales y `numDocSustento` solo existen dentro del XML individual.
- Un mismo `numDocSustento` puede tener varias líneas de retención
  (p.ej. IVA + Renta) — se suman para el total por factura.

## Pendiente / a validar

- [x] Confirmar que `descargar_comprobantes_sri.py` funciona en producción
      → **Resuelto**: usar `AutorizacionComprobantesOffline` con `Connection: close`.
- [x] Confirmar formato del número de documento del ERP
      → **Resuelto**: Odoo usa el mismo formato `NNN-NNN-NNNNNNNNN` que el SRI.
- [ ] Definir si el pipeline se deja como tarea programada (cron /
      Task Scheduler / n8n) y con qué frecuencia.
- [ ] Analizar las ~198 retenciones "Solo en SRI" de agosto para determinar
      si son registros pendientes de ingresar en Odoo o pertenecen a otra empresa/período.

## Dependencias

```bash
pip install pandas openpyxl requests
```

## Convenciones al editar el Excel

- Fórmulas siempre (`INDEX`/`MATCH`, `SUMIFS`, `IFERROR`), nunca valores
  calculados a mano en Python volcados como texto.
- Después de cualquier cambio al workbook con openpyxl, recalcular con
  LibreOffice antes de entregarlo (evita `#NAME?`/celdas en blanco por
  fórmulas sin caché).
- Fuente Arial en todo el workbook; celdas amarillas = a llenar por el
  usuario; nunca sobreescribir fórmulas existentes en `Conciliacion`.
- Colores de las hojas de conciliación: azul = resumen, naranja = Solo SRI,
  rojo = Solo Odoo, amarillo = diferencias de valor, violeta = factura incorrecta.
