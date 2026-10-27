# Conciliación SRI vs Odoo (Ecuador)

## Qué hace este proyecto

Aplicación web (FastAPI, desplegada en Easypanel/Docker) que concilia los comprobantes electrónicos del
portal SRI contra Odoo 14 (leído por XML-RPC en **solo lectura**) y muestra el resultado en un **tablero
en vivo**. El usuario sube los TXT que exporta el portal ("Exportar Txt") y la app cruza, clasifica cada
documento y entrega las excepciones que hay que corregir en Odoo.

RUC de la empresa: `2390000831001`

Dos pestañas independientes (cada una abre su propia sesión de trabajo):

- **Compras**: comprobantes **recibidos** (facturas 01, notas de crédito 04, notas de débito 05 y retenciones recibidas).
- **Ventas**: comprobantes **emitidos** (FAC, NC, ND, LIQ y retenciones emitidas).

## Estructura del repositorio

```
app/
  web/            main.py (FastAPI), auth.py, templates/ (index, login), static/ (tablero.js, tablero.css)
  sesion.py       sesión de trabajo: TXT, recálculo cancelable, publicación atómica de salidas
  job_manager.py  jobs/<id>/: crear, status.json, limpieza automática
  rutas.py        RAIZ, JOBS_DIR, CACHE_XML_DIR, DATA_DIR
  sri/            lectores de TXT/XML del SRI, descarga por WS público, caché de XML
  odoo/           client.py (OdooReader, solo lectura)
  compras/        conciliación de recibidos (motor de cruce, clasificación, Excel, pipeline)
  ventas/         conciliación de emitidos (estados, Excel, pipeline)
  retenciones/    conciliación de retenciones por línea y por documento (comparte Compras/Ventas)
  reportes/       tablero.py (cifras del tablero), observaciones.py (Excel de observaciones), estilos.py
scripts/          CLI sueltos: conciliar_emitidas.py, conciliar_recibidas.py, mover_emitidos.py, setup_usuario.py
tests/            espeja app/ (tests/web, tests/sri, tests/compras, ...) más tests/scripts
deploy/           install.sh (instalación en VPS Ubuntu)
docs/superpowers/ specs y planes de cada funcionalidad (fuente de diseño)
```

## Comandos

```bash
python -m pytest tests -q                              # pruebas
gunicorn app.web.main:app -w 2 -k uvicorn.workers.UvicornWorker --bind 0.0.0.0:8000   # como en el Dockerfile
python scripts/setup_usuario.py add <usuario>          # alta de usuario (data/users.yaml); también remove / list
python scripts/conciliar_emitidas.py odoo.xlsx retenciones_emitidas salida.xlsx [desde] [hasta]   # CLI, ver abajo
```

Dependencias: `pip install -r requirements.txt` (pandas, openpyxl, requests, fastapi, uvicorn, gunicorn, bcrypt, pyyaml, itsdangerous, python-multipart, pytest, httpx).
`Dockerfile` copia solo `app/` y `scripts/`; `cache/xml`, `jobs` y `data` son volúmenes persistentes.

## Aplicación web (FastAPI)

`app/web/main.py`. Auth por cookie de sesión firmada (`SECRET_KEY`, `SESSION_HOURS`, default 8 h); usuarios en
`data/users.yaml` (volumen). Las rutas de API responden **401 JSON** sin sesión; `/` redirige a `/login`.
`/static` (CSS/JS del tablero) es público y sin datos.

| Endpoint | Función |
|---|---|
| `GET /login`, `POST /login`, `GET /logout` | Autenticación |
| `GET /` | Tablero (`templates/index.html`, JS en `static/tablero.js`) |
| `POST /sesion` `{modo: "compras"\|"ventas"}` | Crea la sesión; devuelve `{sesion_id}` |
| `POST /sesion/{id}/txt` (campo `archivos`, uno o varios `.txt`) | Guarda y valida cada TXT; devuelve `{archivos, generacion}`. **No recalcula** |
| `DELETE /sesion/{id}/txt/{nombre}` | Quita un TXT; devuelve `{archivos}` |
| `POST /sesion/{id}/conciliar` | Incrementa la generación y lanza el recálculo completo en un hilo; devuelve `{generacion}` |
| `GET /status/{id}` | `status.json` del último recálculo (`estado`, `mensaje`, `progreso`, `resumen`) |
| `GET /tablero/{id}` | `tablero.json` publicado (404 si aún no hay) |
| `GET /download/{id}?tipo=observaciones\|completo` | Excel (`Observaciones_SRI.xlsx` / `Conciliacion_SRI.xlsx`); default `completo` |

- Los ids (`[0-9a-f]{8}`, de `create_job`) se validan antes de tocar el disco; sesión inexistente o vencida → 404; sesión ocupada (bloqueo > 10 s) → 503.
- **Límites de subida**: Compras 5 TXT y Ventas 400 TXT por sesión (`sesion.LIMITES`; el mismo nombre = reemplazo y no cuenta), **20 MB por archivo** y **200 MB por solicitud** (`MAX_BYTES_TXT`, `MAX_BYTES_LOTE`) → 413 / 400. El middleware ASGI `LimiteSubida` corta con 413 (por `Content-Length` o contando bytes si va chunked) antes de analizar/guardar el cuerpo. Solo `.txt`; los nombres se reducen a basename y la extensión se normaliza a `.txt` minúscula (los pipelines usan `glob`, sensible a mayúsculas en Linux).
- El cliente (`tablero.js`) sube los TXT, y tras ~700 ms sin cambios pide `POST .../conciliar` y sondea `/status` cada 2 s hasta ver su generación; entonces lee `/tablero`.

### Sesión de trabajo (`app/sesion.py`)

Todo el estado vive en disco (gunicorn corre varios workers):

```
jobs/<id>/txt/<nombre>.txt     TXT subidos (mismo nombre = reemplazo)
jobs/<id>/sesion.json          {modo, generacion, archivos:[{nombre, estado ok|error, tipos, filas, error}], ultima_actividad}
jobs/<id>/tablero.json         resultado para la pantalla (incluye "generacion")
jobs/<id>/resultado.xlsx       Excel completo (no existe si no hay ningún TXT válido)
jobs/<id>/observaciones.xlsx   Excel de observaciones
jobs/<id>/status.json          estado del último recálculo (job_manager)
jobs/<id>/.lock                bloqueo por sesión (O_EXCL; serializa hilos y workers)
```

- **Recálculo completo y cancelable**: cada `solicitar_recalculo` sube `generacion` y lanza un hilo. El hilo compara su generación con la de disco en los puntos de control del pipeline y justo antes de publicar; si lo superó termina sin escribir nada. Las salidas se escriben con nombre temporal y se renombran juntas bajo el bloqueo, así que `GET /tablero` y `/download` ven siempre el conjunto de una sola generación (`sesion.leer_publicado`).
- Cada TXT se **valida al subirlo** con los lectores reales; un TXT ilegible o sin tipos válidos queda `estado: error` y se excluye del recálculo con una advertencia. Sin ningún TXT `ok`: tablero vacío, sin consultar Odoo, sin `resultado.xlsx`.
- Limpieza: el job se borra 2 h después de la última actividad. Al arrancar, los jobs `pending`/`running` huérfanos pasan a `error`.
- Errores del recálculo quedan en `status.json` (`estado: error`); nunca incluyen credenciales de Odoo.

### Tablero y Excel de observaciones (`app/reportes/`)

- `tablero.construir_tablero` produce `tablero.json`: `modo`, `periodo`, `tarjetas` (una por tipo FAC/NC/ND/LIQ/RET con `total`, `coinciden`, `por_estado`), `excepciones` (todo lo que no es `Coincide`, con `tipo, estado, documento, fecha, contraparte, ruc, valor_sri, valor_odoo, diferencia, detalle, que_corregir, accionable`) y `advertencias`.
- `observaciones.QUE_CORREGIR` mapea cada estado **accionable** a la instrucción de corrección en Odoo; `ACCIONABLES` es su conjunto: Solo en SRI, Solo en Odoo, Diferencia de valor, Diferencia de datos, Factura incorrecta en Odoo, Odoo anulado/SRI autorizado, Duplicado en Odoo, XML no disponible. Los Informativos no son accionables.
- `escribir_excel_observaciones`: **solo filas accionables**, una hoja por estado, columnas fijas con "Qué corregir en Odoo" (`que_corregir`); sin observaciones, una hoja `Sin_observaciones`. Los textos se protegen contra inyección de fórmulas.
- El Excel **completo** (`resultado.xlsx`) es el de cada pipeline (`excel_compras`, `excel_emitidos`) y se ofrece aparte. Ambos Excel y el tablero salen de las mismas cifras.

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

**Solución implementada** (`app/sri/sri_session.py`, `from app.sri.sri_session import sri_session`):
`HTTPAdapter` personalizado que al detectar esa IP usa `assert_hostname='srienlinea.sri.gob.ec'` y
`server_hostname='srienlinea.sri.gob.ec'` en urllib3. Resultado: la cadena de CA **y** el hostname se
verifican completamente — no se deshabilita nada. `app/sri/descargar_comprobantes_sri.py` usa
`sri_session()`; no existe `verify=False` en el codebase.

## Formato de los XML y TXT del SRI

El TXT del portal ("Exportar Txt") es **ISO-8859-1**, delimitado por **tabulador**.

El SRI usa dos versiones del `<comprobanteRetencion>` (`app/sri/parse_comprobante_xml.py` detecta la versión por el atributo `version` del nodo raíz):

| Versión | Estructura de impuestos |
|---|---|
| **v1.0.0** | `<impuestos>/<impuesto>` — `numDocSustento` dentro de cada `<impuesto>` |
| **v2.0.0** | `<docsSustento>/<docSustento>/<retenciones>/<retencion>` — `numDocSustento` en `<docSustento>`, retenciones anidadas |

Formatos clave (los mismos que usa Odoo en `No. de retención` y `Factura`):

- **`N_Retencion`**: `Establecimiento-PuntoEmision-Secuencial` → `001-200-000004340`
- **`Num_Doc_Sustento`**: `primeros3-sig3-resto` → `001-003-000088766` (split 3-3-9, estándar SRI)

Datos confirmados del TXT:

- Para comprobantes de tipo "Comprobante de Retención", `VALOR_SIN_IMPUESTOS`, `IVA` e `IMPORTE_TOTAL` vienen **siempre vacíos** (578 filas reales de agosto/2026).
- `NUMERO_DOCUMENTO_MODIFICADO` es un ID interno del SRI, **no** el número de la factura sustento — no sirve para conciliar.
- Los montos reales y `numDocSustento` de las retenciones solo existen dentro del XML individual.
- Un mismo `numDocSustento` puede tener varias líneas de retención (p. ej. IVA + Renta) — se suman para el total por factura.

### Caché de XML compartido (importante)

`cache/xml/<clave>.xml` es **compartido entre todos los jobs y períodos** (volumen persistente): evita
re-descargar comprobantes ya bajados. Consecuencia: contiene XML de **todos** los meses procesados.
Nunca alimentes el parseo ni la conciliación con el glob del caché sin filtrar por las claves del TXT del
período: `xml_cache.lineas_sri(claves, ...)` ya filtra por clave (sin ese filtro septiembre mostraba
retenciones de agosto).

## Compras (comprobantes recibidos SRI vs Odoo por API)

Sesión `modo: "compras"` (1 a 5 TXT: facturas, NC, ND y/o retenciones). Concilia facturas (01), notas de
crédito (04), notas de débito (05) y **retenciones recibidas** contra Odoo 14 leído por XML-RPC en **solo lectura**.

| Archivo | Rol |
|---|---|
| `app/odoo/client.py` | `OdooReader.from_env()`: `authenticate` + `search_read` de `account.move` (`in_invoice`/`in_refund`, `posted`, `invoice_date` del período ± 5 días). Calcula `elegible` por fila. Los errores nunca incluyen credenciales. |
| `app/compras/conciliar_compras.py` | `leer_sri_txts`, `tipos_faltantes`, `conciliar` (cruce por niveles: clave de acceso → N° documento + importe → N° documento → NC por factura afectada + importe) y `conciliar_y_excel`. |
| `app/compras/clasificar.py` | `clasificar_compras`: estado por documento a partir de los pares del cruce (mismos estados que Ventas). |
| `app/compras/excel_compras.py` | `construir_excel`: libro con fórmulas (resumen por tipo + detalle). |
| `app/compras/pipeline_compras.py` | `procesar_compras(rutas, salida, avisar, cancelado, advertencias)` → `(resumen, tablero)`; lo invoca `sesion.recalcular`. |

- Variables de entorno del servicio (Easypanel, nunca en la imagen): `ODOO_URL`, `ODOO_DB`, `ODOO_USER`, `ODOO_PASSWORD`. `os.environ` tiene prioridad; el `.env` local (ignorado por git) solo completa las que falten.
- **Regla de elegibilidad** (una fila de Odoo se cruza solo si): tipo de comprobante `x_receipt_type_id` 01/04/05 **y** `x_authorization_number` numérico de 45–49 dígitos **y** `x_is_edocument` falso (no emitido por la propia empresa). Todo lo demás (tipo `OTRO`, sin clave, documentos propios) queda como "No aplica (informativo)".
- Las NC se comparan por valor absoluto y se enlazan a la factura afectada (`reversed_entry_id`).
- Facturas/NC/ND **no usan** el caché `cache/xml/` ni descargan XML: trabajan solo con el TXT (solo las retenciones descargan XML).
- **Fuera de alcance**: liquidaciones de compra y vista diaria.
- Período de reporte = **mes dominante** de las fechas de emisión del SRI (el que tiene más filas; empate: el más reciente). La lectura de Odoo usa la ventana min..max de fechas del SRI (más el margen de 5 días) para que un comprobante atípico (p. ej. una ND de julio dentro del TXT de septiembre) aún encuentre su par; los asientos de Odoo sin par fuera del período de reporte no se listan.

### Retenciones recibidas dentro de Compras

Si entre los TXT viene el de "Comprobante de Retención", la misma sesión concilia también las retenciones
(hojas `Ret_*` en el Excel completo). Cualquier subconjunto de TXT es válido. A diferencia de
facturas/NC/ND, las retenciones **sí** necesitan el XML (montos y `numDocSustento` solo existen ahí) y
usan el caché compartido.

| Archivo | Rol |
|---|---|
| `app/sri/xml_cache.py` | `descargar_xmls` (WS público, con caché; devuelve los fallidos) y `lineas_sri` (líneas de retención de los XML del caché, filtradas por las claves del TXT). |
| `app/retenciones/conciliar_retenciones.py` | `leer_sri_retenciones_txt`, `conciliar_retenciones` (cruce **por línea**: `(N° retención, factura, porcentaje)` y estados) y `resumir_por_documento` (agrupa las líneas en **documentos** con un estado de documento). |
| `app/retenciones/excel_retenciones.py` | Hojas `Ret_Documentos`, `Ret_Conciliacion`, `Ret_Solo_SRI`, `Ret_Solo_Odoo`, `Ret_Diferencia_Valor`, `Ret_Factura_Incorrecta`, `Ret_Informativos` y la fila "Retenciones" del Resumen (fórmulas `COUNTIFS` sobre `Ret_Documentos`). |
| `app/odoo/client.py` | `fetch_retenciones`: `out_retention` recibidas por XML-RPC, solo lectura, todos los estados. |

- **Un documento = una fila**: un comprobante de retención (N° de retención) suele tener varias líneas (Renta + IVA, hasta 4). El cruce sigue siendo por línea, pero el resultado se presenta por **documento**:
  - `Ret_Documentos` (primera hoja `Ret_*`): una fila por documento con Estado, líneas y valores SRI/Odoo (fórmulas `COUNTIFS`/`SUMIFS` sobre `Ret_Conciliacion` por la columna `Documento`), Diferencia y Detalle (p. ej. "1 de 2 líneas con diferencia (valor: 1)").
  - `Ret_Conciliacion` es el **detalle por línea**, con las líneas de cada documento contiguas y dos columnas al final: `Documento` (clave) y `Estado_Documento` (fórmula `INDEX`/`MATCH` sobre `Ret_Documentos`).
  - Estado del documento: se ignoran las líneas Informativo si hay otras; todas Informativo → Informativo; todas XML no disponible → XML no disponible; alguna Factura incorrecta → Factura incorrecta; todas Coincide / Solo en SRI / Solo en Odoo → ese estado; cualquier otra mezcla (p. ej. Renta coincide e IVA falta en Odoo) → Diferencia de valor.
  - Clave de documento: N° de retención; si el mismo N° aparece con dos RUC distintos son dos documentos (`N° | RUC`); líneas sin N° son documentos propios (`(sin N°) id X`).
  - `Ret_Solo_SRI`, `Ret_Solo_Odoo`, `Ret_Factura_Incorrecta`, `Ret_Informativos`: una fila por documento. `Ret_Diferencia_Valor`: todas las líneas de los documentos con diferencia, con `Estado_Linea`.
  - Los contadores (`resumen.por_tipo.retenciones`, la fila "Retenciones" del Resumen y la tarjeta RET del tablero) cuentan **documentos**; `documentos` y `lineas` dan los totales; `total_sri`/`total_odoo` siguen sumando valores.
- Clave de cruce por línea: `(N° Retencion, Factura, Porcentaje)`. Estados de línea: **Coincide** (mismo valor), **Solo en SRI**, **Solo en Odoo** (N° incorrecto, otro período o no autorizada), **Factura incorrecta en Odoo** (N° de retención correcto pero distinto número de factura sustento, típicamente error en el establecimiento), **Diferencia de valor** (valor retenido difiere en más de $0.02).
- **Porcentaje distinto**: mismo N° de retención y misma factura pero distinto porcentaje se empareja uno a uno por valor más cercano y se reporta como *Diferencia de valor*, no como Solo SRI + Solo Odoo.
- **XML no disponible**: comprobantes del TXT cuyo XML no se pudo descargar o leer (o está vacío) no se declaran "Solo en SRI"; se listan como "XML no disponible" (se detecta por `(RUC emisor, N° retención)`). El conteo va en `resumen.retenciones_fallidos` y en `advertencias`.
- Odoo cancelada/borrador queda "Informativo"; Odoo publicado solo cuenta como "Solo en Odoo" si su fecha cae en el período.
- El período de las retenciones se calcula con sus propias fechas (mes dominante) y su propia ventana de lectura en Odoo, independiente de la de compras.
- Si el TXT de retenciones no tiene ninguna fecha de emisión válida la corrida termina en `error`.

## Ventas (comprobantes emitidos SRI vs Odoo por API)

Concilia lo que la **empresa emite**: facturas (FAC), notas de crédito (NC), notas de débito (ND),
liquidaciones de compra (LIQ) y retenciones emitidas (RET). Uso semanal: el usuario sube **todos los TXT de
la semana** del portal SRI (hasta 400, tipos mezclados, con cualquier nombre de archivo) en una sesión
`modo: "ventas"`; Odoo 14 se lee por XML-RPC en solo lectura (mismas variables `ODOO_*` que Compras). No se
sube ningún Excel de Odoo.

| Archivo | Rol |
|---|---|
| `app/sri/emitidos_sri.py` | `RUC_EMPRESA`, `TIPOS_DOC`, `leer_emitidos_txts` (devuelve `docs, ret, advertencias`) y `periodo_emitidos`. |
| `app/ventas/conciliar_emitidos.py` | `clasificar_documentos` (estado por documento FAC/NC/ND/LIQ, reutiliza el motor `conciliar` de Compras) y `conciliar_retenciones_emitidas` (envuelve `conciliar_retenciones`). |
| `app/ventas/excel_emitidos.py` | `construir_excel_emitidos`: libro con fórmulas; las retenciones reutilizan las hojas `Ret_*` de `excel_retenciones`. |
| `app/ventas/pipeline_ventas.py` | `procesar_ventas(rutas, salida, avisar, cancelado, advertencias)` → `(resumen, tablero)`: TXT → período → Odoo → XML de retenciones (si hay) → estados → Excel. |
| `app/odoo/client.py` | `fetch_emitidos` (documentos) y `fetch_retenciones_emitidas` (`in_retention`, comparte código con `fetch_retenciones`); ambos con margen de 5 días y errores sin credenciales. |
| `app/sri/xml_cache.py` | `lineas_sri(claves, cache_dir, contraparte=True)`: `proveedor` pasa a ser la razón social del sujeto retenido. |

- **Tipos**: se detectan por la columna `COMPROBANTE` del TXT, nunca por el nombre del archivo. Las guías de remisión se ignoran con una advertencia (igual que los tipos desconocidos). Claves repetidas se deduplican (se conserva la primera fila; advertencia si los datos difieren). Las retenciones sin clave de acceso de 45–49 dígitos se descartan con la advertencia "N retenciones del TXT se ignoraron por no tener clave de acceso válida".
- **`RUC_EMPRESA`** (`app/sri/emitidos_sri.py`): en retenciones **emitidas** el `ruc` de SRI (XML) y de Odoo es siempre el de la empresa (es el emisor); el proveedor retenido va en `proveedor` (SRI, razón social del sujeto retenido) / `socio` (Odoo). No lo uses para identificar al proveedor.
- **Mapeo en Odoo** (`fetch_emitidos`; estados `posted`, `cancel` y `draft`): `out_invoice` con tipo de comprobante 01 → FAC; `out_refund` con 04 → NC; `out_invoice` con 05 → ND; `in_invoice` con `account.receipt.type` de código `03` (ids 3 Liquidación de compra y 15, ambos con ese código) → LIQ. Otros tipos o códigos incoherentes con el `move_type` se descartan. `elegible` = publicado o cancelado. Retenciones: `account.move.retention` con `type = in_retention`, clave desde `auth_retention_number`; una retención publicada cuyo `edocument_state` no es `authorized` sale con `estado = "no_autorizada"` y por tanto queda como Informativo (fuera del cruce, nunca "Solo en Odoo").
- **Estados de documento** (FAC/NC/ND/LIQ, un registro por pareja o lado sin pareja):

| Estado | Significado |
|---|---|
| **Coincide** | Total y base iguales (±$0.02), misma clave, misma fecha y electrónico en Odoo |
| **Solo en SRI** | Autorizado en el SRI sin registro en Odoo |
| **Solo en Odoo** | Publicado, electrónico y sin autorización en el SRI (dentro del período) |
| **Diferencia de valor** | Total o base distintos |
| **Diferencia de datos** | Mismos valores pero clave, fecha o marca electrónica distintas |
| **Odoo anulado, SRI autorizado** | Cancelado en Odoo y autorizado en el SRI (hoja `Anulados_Odoo`) |
| **Duplicado en Odoo** | Otro registro de Odoo ya cruza con ese documento; un publicado prevalece sobre el cancelado |
| **Informativo** | Borrador, cancelado sin autorización, no electrónico o duplicado anulado en Odoo |

  Las **retenciones** usan los estados de Compras por documento (`Coincide`, `Solo en SRI`, `Solo en Odoo`, `Diferencia de valor`, `Factura incorrecta en Odoo`, `XML no disponible`, `Informativo (Odoo cancelada/borrador)`) más `Odoo anulado, SRI autorizado` (una retención cancelada en Odoo y autorizada en el SRI).
- **Período** = min..max de las fechas de emisión de todos los TXT (sin mes dominante); si `max - min > 35` días se agrega una advertencia. Odoo se lee con esa ventana ± 5 días; los registros de Odoo sin pareja fuera de min..max no se listan.
- **Límite de 30 días del XML**: los XML de retenciones solo se pueden bajar del WS público durante ~30 días; las más antiguas quedan como `XML no disponible` (no `Solo en SRI`) salvo que ya estén en `cache/xml/` (caché compartido).
- **Falta un tipo**: si la carga trae documentos pero ninguna fila de un tipo (FAC/NC/ND/LIQ) y Odoo tiene registros publicados o cancelados de ese tipo dentro de min..max, se advierte "Falta el TXT de {facturas|notas de crédito|notas de débito|liquidaciones}: los N registros de Odoo de ese tipo aparecerán como 'Solo en Odoo'" (Odoo no se filtra por tipo).
- **Solo retenciones**: si la carga no trae documentos (FAC/NC/ND/LIQ) no se leen documentos de Odoo, solo `in_retention`. Si trae retenciones sin ninguna fecha válida, la corrida termina en `error`; sin ningún documento ni retención también.
- **Excel completo**: `Resumen` (una fila por FAC/NC/ND/LIQ, fila "Retenciones" por documentos, notas y advertencias), `Conciliacion` (una fila por documento), `Solo_en_SRI`, `Solo_en_Odoo`, `Diferencias` (valor y datos), `Anulados_Odoo` y, si hubo retenciones, `Ret_Documentos`, `Ret_Conciliacion`, `Ret_Solo_SRI`, `Ret_Solo_Odoo`, `Ret_Diferencia_Valor`, `Ret_Factura_Incorrecta`, `Ret_Informativos`.
- **`resumen` del job** (`/status/{id}`): `por_tipo` (por FAC/NC/ND/LIQ y, solo si se subió un TXT con retenciones, `RET`: `total` y conteo por estado), `comprobantes_sri` (filas del TXT por tipo, incl. `RET`), `periodo` (`[desde, hasta]` ISO), `advertencias`, `retenciones_fallidos` (XML no descargados o ilegibles) y `recalculado` (si LibreOffice recalculó las fórmulas; si no, hay advertencia).

## Scripts CLI (`scripts/`)

- `conciliar_emitidas.py`: concilia retenciones **EMITIDAS** (la empresa retiene a proveedores): TXT diarios `retenciones_emitidas/emitidos_AAAA-MM-DD.txt` vs Excel de Odoo. Cruza por N° de retención a nivel de comprobante (el TXT no trae montos y el WS público solo guarda XML de ~30 días). Genera `Conciliacion_Retenciones_Emitidas_*.xlsx` con `Resumen_Mensual`, `Conciliacion`, `Solo_en_SRI`, `Solo_en_Odoo`, `Fecha_distinta`. Uso: `python scripts/conciliar_emitidas.py odoo.xlsx retenciones_emitidas salida.xlsx [desde] [hasta]`. **Independiente de la app web.** No confundir con `app/ventas/conciliar_emitidos.py` (con **o**), que es la lógica de la pestaña Ventas.
- `conciliar_recibidas.py`: concilia facturas recibidas (TXT del SRI con montos vs Excel de Odoo), cruce en tres niveles; uso en su docstring. Independiente de la app.
- `mover_emitidos.py`: mueve los TXT descargados del SRI a `emitidos_AAAA-MM-DD.txt` según su fecha de emisión (insumo del CLI de arriba).
- `setup_usuario.py`: alta/baja/listado de usuarios de la app en `data/users.yaml`.

## Convenciones

- Excel: fórmulas siempre (`INDEX`/`MATCH`, `SUMIFS`, `IFERROR`, `COUNTIFS`), nunca valores calculados a mano en Python volcados como texto; fuente Arial; tras generar el workbook se recalcula con LibreOffice cuando está disponible (evita celdas en blanco por fórmulas sin caché; si no, `recalculado: false` y advertencia).
- Colores de hojas de conciliación: azul = resumen, naranja = Solo SRI, rojo = Solo Odoo, amarillo = diferencias de valor, violeta = factura incorrecta.
- Datos del negocio (`*.xlsx`, `*.txt`, `*.xml`, `jobs/`, `cache/`, `users.yaml`, `.env`) nunca se versionan (ver `.gitignore`).
- Diseño y planes de cada funcionalidad: `docs/superpowers/specs/` y `docs/superpowers/plans/`.

## Pendiente / a validar

- [ ] Analizar las ~198 retenciones "Solo en SRI" de agosto para determinar si son registros pendientes de ingresar en Odoo o pertenecen a otra empresa/período.
