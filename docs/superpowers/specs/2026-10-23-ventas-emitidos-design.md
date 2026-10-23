# Pestaña Ventas: conciliación de comprobantes electrónicos emitidos (SRI vs Odoo)

Fecha: 2026-10-23. Estado: borrador para aprobación.

## Intención

Hoy la empresa (RUC `2390000831001`) revisa a mano si lo que emitió y está
autorizado en el SRI coincide con lo registrado en Odoo. Son miles de
documentos al mes (septiembre 2026: 6.265 facturas, 523 notas de crédito,
239 retenciones, 74 liquidaciones). La pestaña **Compras** ya hace esta
revisión para los comprobantes **recibidos**. Se quiere lo mismo para los
**emitidos**, automatizado, para tener constancia de que Odoo y el SRI dicen
lo mismo.

La revisión se hará **semana por semana**: el usuario descarga del portal SRI
todos los TXT de emitidos de la semana, los sube juntos en un solo envío y
recibe un Excel de conciliación.

## Alcance

Documentos cubiertos:

1. Facturas (01)
2. Notas de crédito (04)
3. Notas de débito (05) (en septiembre 2026 no hubo ninguna; debe funcionar igual)
4. Retenciones emitidas (07), **cruzadas por línea y con valores** (renta, IVA, cada porcentaje)
5. Liquidaciones de compra (03)

La pestaña **Ventas actual** (solo facturas, Excel de Odoo subido a mano, TXT
`emitidos_AAAA-MM-DD.txt`) se **reemplaza** por la nueva. La ruta
`/upload-ventas` y su pipeline pasan a ser la nueva funcionalidad.

### No-objetivos

- **Guías de remisión**: fuera de alcance (decisión del usuario).
- Escribir en Odoo o en el SRI: todo es solo lectura.
- Tarea programada / ejecución sin intervención: el usuario sube los TXT cada semana a mano.
- Conciliar valores de retenciones cuya XML ya no esté disponible en el SRI (ver límite abajo).

### Anti-objetivos (fallaría aunque "funcione")

- Que una retención o factura autorizada en el SRI y ausente en Odoo pase sin avisar.
- Que se oculte una discrepancia por tolerancia mal aplicada (tolerancia fija $0,02).
- Que errores o logs expongan credenciales de Odoo.
- Deshabilitar verificación TLS (no existe `verify=False` en el código; se usa `sri_session()`).

## Entradas

### TXT del SRI (portal "Comprobantes electrónicos emitidos" → exportar)

- Un archivo por día o por consulta, ISO-8859-1, separado por tabulador, nombre real del portal
  (p. ej. `2390000831001_EmitidosAutorizados (1).txt`) o `emitidos_AAAA-MM-DD.txt`.
  **No se debe depender del nombre del archivo.**
- Columnas (todas las clases de documento): `COMPROBANTE`, `SERIE_COMPROBANTE`,
  `CLAVE_ACCESO`, `FECHA_AUTORIZACION`, `FECHA_EMISION`, `VALOR_SIN_IMPUESTOS`, `IVA`, `IMPORTE_TOTAL`.
- El **tipo** se determina por el texto de `COMPROBANTE` (Factura, Nota de Crédito, Nota de Débito,
  Liquidación de Compra, Comprobante de Retención), normalizado sin acentos ni mayúsculas.
  Un envío puede mezclar tipos y varias fechas.
- `FECHA_EMISION` puede traer hora (`01/09/2026 00:00:00`): se usa solo la fecha (primeros 10 caracteres, `dd/mm/aaaa`).
- Hechos confirmados con archivos reales:
  - **Notas de crédito**: `IMPORTE_TOTAL` viene vacío en todas las filas → total = base + IVA.
  - **Liquidaciones y facturas**: `IMPORTE_TOTAL` viene lleno.
  - **Retenciones**: `VALOR_SIN_IMPUESTOS`, `IVA`, `IMPORTE_TOTAL` vacíos; los montos solo existen en el XML.
- Filas de tipo "Guía de Remisión" u otros no soportados se ignoran y se informa en `advertencias`.
- Tipos faltantes en el envío no son error: se concilia lo que venga (como Compras).

### Odoo 14 (XML-RPC, solo lectura)

Variables de entorno `ODOO_URL`, `ODOO_DB`, `ODOO_USER`, `ODOO_PASSWORD`, igual que Compras
(`OdooReader.from_env()`, errores sin credenciales).

Confirmado en septiembre 2026 (documentos publicados):

| Documento | Modelo / filtro en Odoo |
|---|---|
| Factura | `account.move`, `move_type = out_invoice`, tipo de comprobante Factura (01), `x_is_edocument = true` |
| Nota de crédito | `account.move`, `out_refund`, tipo 04, `x_is_edocument = true`; factura afectada por `reversed_entry_id` |
| Nota de débito | `account.move`, `out_invoice` con `x_receipt_type_id.code = '05'` (se identifica solo por tipo de comprobante). No existe ninguna ND publicada en el historial de Odoo, así que no hay ejemplo real: el código debe soportarla y probarse con datos sintéticos. |
| Liquidación de compra | `account.move`, `in_invoice`, `x_is_edocument = true`, tipos de comprobante **3** y **15** ("Comprobante de venta emitido por reembolso"). Tienen la misma serie `001-002` y numeración continua. **Confirmado contra datos:** el 01/09/2026 el SRI lista 7 liquidaciones y Odoo tiene 7 documentos tipo 15 con los mismos números, claves e importes. |
| Retención emitida | `account.move.retention`, `type = in_retention`. Campos confirmados: `retention_number` (N° `NNN-NNN-NNNNNNNNN`), `auth_retention_number` (clave de acceso), `edocument_state` (p. ej. `authorized`), `state`, `retention_date`, `partner_id`, `invoice_id` (enlace a `account.move`, factura sustento). Líneas en `account.move.retention.line` (`amount_untaxed`, `amount_retention`, `tax_percent`, `tax_id`), igual que `fetch_retenciones` hace para `out_retention`. Contraparte = proveedor retenido. |

- Campos útiles ya usados por Compras: `x_number` (N° `NNN-NNN-NNNNNNNNN`),
  `x_authorization_number` (clave de acceso, 45–49 dígitos), `x_receipt_type_id`,
  `x_is_edocument`, `invoice_date`, `amount_untaxed`, `amount_tax`, `amount_total`, `state`.
- Los nombres de campo de arriba se verificaron en Odoo; quien construya debe leer también `edocument_state` para retenciones y tratar como no electrónicas las que no estén `authorized`.
- A diferencia de Compras, para emitidos se leen **estados publicado y cancelado** (se necesita el caso "anulado en Odoo, autorizado en el SRI"); borrador queda como informativo.
- Ventana de lectura: rango min..max de fechas de emisión de los TXT, más margen de 5 días.
- Retenciones: fecha por `retention_date`.

## Qué se concilia y cómo

### Facturas, NC, ND y liquidaciones (por documento)

Cruce por niveles, reutilizando el motor de `conciliar_compras.py`:
1. Clave de acceso.
2. N° de documento (+ tipo) con importe.
3. N° de documento.
4. NC: por factura afectada + importe.

Importes: base y total con tolerancia $0,02; NC en valor absoluto.

Estados (propuestos y aceptados en conversación como base; el usuario no pidió cambios):

| Estado | Significado |
|---|---|
| Coincide | Mismo documento, clave, número e importes |
| Solo en SRI | Autorizado en el SRI, sin registro en Odoo |
| Solo en Odoo | Publicado en Odoo como electrónico, con fecha dentro del período, sin autorización en el SRI |
| Diferencia de valor | Base o total difieren más de $0,02 |
| Diferencia de datos | Coincide el documento pero difiere fecha de emisión o clave de acceso |
| Odoo anulado, SRI autorizado | Cancelado en Odoo, vigente en el SRI (separado por su acción correctiva) |
| Informativo | Borrador/no aplica (p. ej. documento no electrónico) |

### Retenciones emitidas (por línea, presentadas por documento)

- Requieren el **XML** del comprobante (montos y `numDocSustento` solo existen ahí).
  Se descarga por clave de acceso del servicio público del SRI (`AutorizacionComprobantesOffline`,
  `Connection: close`, `sri_session()`), con el caché compartido `cache/xml/` y `xml_cache.py`.
  Filtrar siempre el caché por las claves del TXT cargado (nunca el glob completo).
- Parser: `parse_comprobante_xml.py` ya soporta v1.0.0 y v2.0.0. Para emitidas la contraparte
  es el sujeto retenido (proveedor), no el emisor.
- Cruce por línea `(N° retención, factura sustento, porcentaje)` y presentación por documento
  (un documento = una fila con estado), con los mismos estados que Compras:
  Coincide, Solo en SRI, Solo en Odoo, Diferencia de valor, Factura incorrecta,
  XML no disponible, Informativo.
- Porcentaje distinto con mismo N° y factura → *Diferencia de valor*.
- **Límite**: el servicio público solo conserva el XML ~30 días. Con ejecución semanal no es problema;
  semanas antiguas sin XML en caché terminan como "XML no disponible" y se cruzan solo por documento.
  El conteo va en `resumen.retenciones_fallidos` y en `advertencias`.
- Si el TXT de retenciones no tiene fechas válidas, el job termina en `error`.

### Período

La revisión es semanal, no mensual. Período de reporte = rango min..max de fechas de emisión de los TXT
cargados (**decisión mía**; Compras usa el mes dominante). Documentos de Odoo con fecha fuera del período
(dentro del margen de lectura) que no tengan par no se listan como "Solo en Odoo".

## Salidas

Un Excel (`resultado.xlsx`) con las convenciones del proyecto: Arial, fórmulas
(`COUNTIFS`/`SUMIFS`/`INDEX`/`MATCH`) en lugar de valores calculados, recalculado con LibreOffice,
colores azul = resumen, naranja = Solo SRI, rojo = Solo Odoo, amarillo = diferencias de valor, violeta = factura incorrecta.

- Hoja `Resumen` con una fila por tipo de documento (Facturas, NC, ND, Liquidaciones, Retenciones) y conteos por estado.
- Detalle por tipo de documento (misma estructura que `excel_compras.py`).
- Hojas `Ret_*` para retenciones (misma estructura que `excel_retenciones.py`: `Ret_Documentos`, `Ret_Conciliacion`, etc.).
- `advertencias` en el resumen (tipos ignorados, XML no disponible, etc.).

## Interfaz y arquitectura

- Misma pestaña "Ventas": un solo campo de carga multi-archivo (`emitidos`, uno o muchos `.txt`);
  ya **no** se pide `odoo_venta.xlsx`.
- Ruta `POST /upload-ventas` (reemplaza la actual); estado y descarga por `/status/{job_id}` y `/download/{job_id}`.
- Pipeline en hilo background con progreso, mismo patrón que `pipeline_compras.py`
  (Odoo se lee primero, falla rápido; luego descarga de XML solo si hay retenciones).
- Reutilizar al máximo: `OdooReader`, `xml_cache.py`, `conciliar_retenciones.py`, `excel_retenciones.py`,
  motor de cruce de `conciliar_compras.py`. Lo nuevo es la lectura de Odoo de emitidos, la lectura/normalización
  de TXT por `COMPROBANTE`, las reglas de estado propias de emitidos y el armado del Excel.

## Restricciones

- Odoo 14, solo lectura; credenciales solo por variables de entorno (Easypanel).
- Sin `verify=False` hacia el SRI; usar `sri_session()`.
- Tests con `python -m pytest tests -q`; los tests de Ventas actuales (`test_conciliar_ventas.py`, `test_pipeline_ventas.py`)
  se reemplazan por los de la nueva funcionalidad.

## Decisiones tomadas por mí (a revisar)

1. Período = min..max de las fechas de emisión de los TXT (no el mes dominante). Si el rango supera 35 días se agrega una advertencia (probable comprobante atípico o TXT equivocado) y se sigue.
2. Liquidaciones = `in_invoice` electrónicas de tipos 3 y 15 (verificado con datos del 01/09/2026).
3. Se leen documentos publicados y cancelados de Odoo para detectar "anulado en Odoo, autorizado en el SRI"; el estado aplica a **todos** los tipos (incluidas retenciones, donde sustituye al "Informativo" de Compras para cancelada).
3b. **Duplicados:** si la misma clave de acceso aparece en más de un TXT (consultas que se solapan), se deduplica por clave; si los datos difieren se conserva la primera fila y se agrega advertencia.
3c. **SRI sin Odoo elegible:** un documento del SRI sin par por clave pero con el mismo N° en Odoo (sin clave o no electrónico) se reporta como "Diferencia de datos"; sin ningún par es "Solo en SRI".
3d. **Sin compatibilidad hacia atrás:** `/upload-ventas` con `odoo_venta.xlsx` desaparece por completo; la app es de uso interno, sin otros consumidores.
4. Guías de remisión y tipos desconocidos se ignoran con advertencia, sin error.
5. Los scripts viejos de Ventas (`conciliar_ventas_sri.py`, `conciliar_facturas_emitidas.py`, `pipeline_ventas.py`)
   se retiran si nada más los referencia; `conciliar_emitidas.py` (CLI de retenciones) se deja intacto.

## Queda a criterio del constructor

- Estructura interna de módulos y nombres de archivos nuevos.
- Cómo particionar el trabajo en tareas y el orden (sugerido: TXT y tipos → Odoo emitidos → cruce → Excel → pipeline/UI → retenciones emitidas).
- Detalles de las hojas de detalle del Excel que no contradigan lo anterior.
- Nombre exacto de los estados cuando no esté definido arriba.
