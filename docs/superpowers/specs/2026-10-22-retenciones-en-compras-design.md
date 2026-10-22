# Retenciones recibidas dentro de la pestaña Compras (SRI vs Odoo por API)

Fecha: 2026-10-22 · Estado: borrador para aprobación

## Intención y por qué

La pestaña Compras ya concilia facturas (01), notas de crédito (04) y notas de
débito (05) recibidas: TXT del SRI contra Odoo 14 leído por API. Las retenciones
recibidas siguen en otra pestaña, con otro flujo: el personal debe exportar a mano
`reten_odoo.xlsx` de Odoo y subirlo junto al TXT. El objetivo es que **todos los
comprobantes recibidos se concilien en un solo lugar, con un solo envío de TXT y
un solo Excel**, sin exportar nada de Odoo a mano.

Éxito: el usuario sube los TXT del período (facturas, NC, ND y retenciones) en la
pestaña Compras y recibe un único Excel que dice, para cada tipo, qué está solo en
el SRI, qué está solo en Odoo y **dónde el valor difiere** entre ambos.

## Decisiones del usuario

- Un solo Excel con todo (compras y retenciones). Confirmado el 2026-10-22.
- Odoo se lee por API, en solo lectura (misma conexión y credenciales por variables
  de entorno que Compras: `ODOO_URL`, `ODOO_DB`, `ODOO_USER`, `ODOO_PASSWORD`).
- Lo que más importa, además de "existe en SRI / existe en Odoo", es **detectar
  diferencias de valor**: el SRI dice un valor retenido y Odoo tiene otro.
- La lógica de conciliación vigente de retenciones (`conciliar_odoo.py`) es la
  referencia del comportamiento esperado: el usuario dijo que ahí está todo lo que
  necesita detectar.

## Alcance

Se agrega el cuarto tipo, "Comprobante de Retención", al flujo de Compras:

1. El usuario sube hasta 5 TXT (ya admitido) y uno puede ser el de retenciones.
2. Las claves de acceso del TXT de retenciones se descargan como XML con el web
   service público del SRI, usando el caché compartido `cache/xml/` filtrado
   siempre por las claves del TXT cargado (el TXT de retenciones no trae montos;
   los montos y `numDocSustento` solo existen en el XML).
3. Se leen de Odoo por API las retenciones recibidas del período.
4. Se concilia y las hojas de retenciones se agregan al Excel de Compras.

## Lo ya verificado en Odoo 14 (solo lectura, 2026-10-22)

- Modelo `account.move.retention` (17.233 registros). Retención recibida =
  `type = 'out_retention'` (etiqueta "Retención venta": la retiene el cliente).
  `in_retention` son las que emite la empresa y **no** entran.
- Campos útiles: `retention_number` (`001-001-000007460`), `retention_date`,
  `state` (`draft`, `posted`, `cancel`), `partner_id`, `invoice_id` (factura
  sustento, `account.move`), `retention_total`, `line_ids`,
  `edocument_access_key`, `is_edocument`.
- Modelo `account.move.retention.line`: `amount_untaxed` (base), `amount_retention`
  (valor retenido), `tax_percent` (porcentaje), `tax_id` (nombre tipo
  `312 - Transferencia de bienes ... (2%)`, empieza con el código SRI).
- Septiembre 2026, por fecha de retención: 589 `out_retention` publicadas, 7
  canceladas o en borrador; 239 `in_retention` publicadas (no aplican).
- El número de la factura sustento tiene formato `NNN-NNN-NNNNNNNNN` como en el
  SRI (`invoice_id` en el ejemplo leído mostró `006-001-000003385`).
- El usuario de API puede leer estos modelos. El modelo de reporte RIDE no es
  accesible (no hace falta).

## Reglas de conciliación (mismas que `conciliar_odoo.py`)

Clave de cruce a nivel de línea: `(N° retención, Factura sustento, Porcentaje)`.
Tolerancia de $0,02 en el valor retenido.

| Estado | Significado |
|---|---|
| Coincide | Línea en SRI y Odoo, valor retenido igual (±0,02) |
| Solo en SRI | Autorizada en el SRI y no registrada en Odoo |
| Solo en Odoo | Registrada en Odoo (publicada) y no encontrada en el SRI |
| Diferencia de valor | Cruza por clave, valor retenido distinto en más de 0,02 |
| Factura incorrecta en Odoo | Mismo N° de retención, secuencial de factura (últimos 9 dígitos) y porcentaje, pero distinto establecimiento/punto de emisión |

- La hoja de diferencias muestra también la base imponible de ambos lados, para
  que se vea si el error está en la base o en el valor.
- Las retenciones de Odoo canceladas o en borrador se listan como informativas,
  nunca como "Solo en Odoo".
- El período de reporte y la ventana de lectura de Odoo siguen la lógica ya
  vigente en Compras (mes dominante de las fechas del SRI; ventana min..max ± 5
  días), aplicada sobre `retention_date`.

## Entrega

El Excel de Compras gana las hojas de retenciones: resumen (incluido en la hoja
Resumen existente, una fila por tipo), Conciliación de retenciones, Solo en SRI,
Solo en Odoo, Diferencia de valor, Factura incorrecta, Informativos. Reglas
vigentes: Arial, fórmulas (no valores pegados), colores por hoja (azul resumen,
naranja Solo SRI, rojo Solo Odoo, amarillo diferencias, violeta factura
incorrecta), protección contra inyección de fórmulas en texto, `fullCalcOnLoad`.

## Restricciones

- Solo lectura sobre Odoo. Credenciales solo por variables de entorno; nunca en
  logs, errores, Excel ni repositorio.
- Caché de XML compartido entre períodos: filtrar siempre por las claves del TXT.
- Datos de negocio (TXT/XLSX/XML, `jobs/`, `cache/`) nunca versionados.
- Misma autenticación de la app (cookie firmada). No romper las pestañas
  existentes, incluida la actual de Retenciones, que queda como está hasta que el
  usuario confirme que la nueva funciona.
- Si algún XML no se puede descargar: advertencia en el resumen con la cantidad y
  las claves; la conciliación continúa con lo disponible. No se elude ningún
  control del SRI.

## No objetivos y antiobjetivos

- No eliminar ni modificar la pestaña Retenciones antigua en esta fase.
- No escribir nada en Odoo.
- No automatizar login ni captcha del portal del SRI.
- No incluir retenciones emitidas (`in_retention`), que tienen su propio flujo.
- Antiobjetivo: marcar como error retenciones canceladas o fuera del período.

## Llamados míos tras la revisión (pendientes de que los revises)

- **XML no descargado**: no debe generar falsos "Solo en Odoo". Se cruza a nivel
  de comprobante por N° de retención (el TXT lo trae en la serie del comprobante,
  y Odoo en `retention_number`): las retenciones de Odoo cuyo N° está en el TXT
  pero cuyo XML falló se listan como informativas "XML no disponible", nunca como
  error. En Odoo `edocument_access_key` viene vacío en las recibidas, por eso no
  se usa. El web service guarda XML solo unos 30 días: para períodos viejos
  conviene que el caché ya tenga los XML.
- **Porcentaje distinto**: una segunda pasada empareja las filas "Solo en SRI" y
  "Solo en Odoo" con el mismo N° de retención y factura pero distinto porcentaje,
  y las muestra como "Diferencia de valor" con ambos porcentajes, bases y valores.
  Es una regla nueva respecto a `conciliar_odoo.py`, que las dejaba como dos filas
  sueltas; responde a tu prioridad de ver dónde difiere el valor.
- **Ventana de fechas**: cada tipo usa la ventana de su propio TXT (mes dominante
  y min..max ± 5 días), porque las retenciones suelen emitirse en otro mes que la
  factura que sustentan.
- **Cualquier subconjunto de TXT es válido**, incluido solo retenciones. Los tipos
  sin TXT se omiten con advertencia. Como retenciones descarga XML, el estado del
  trabajo informa el avance de la descarga y las fallas parciales no abortan.

## Lo que queda al constructor

- Cómo reutilizar el código existente: extraer la lógica de cruce de
  `conciliar_odoo.py` a una función pura sobre DataFrames y/o reusar
  `parse_comprobante_xml.py` / descarga con caché, sin romper la pestaña antigua.
- Obtener el número de factura sustento de `invoice_id` (campo exacto de
  `account.move`; en facturas de venta puede estar en `name` y no en `x_number`).
- Extracción del código y del porcentaje de retención desde `tax_id` /
  `tax_percent`, y cómo cruzar con `Porcentaje Retener` / `Codigo Retencion` del XML.
- Cómo tratar varias líneas con el mismo (retención, factura, porcentaje).
- Mensajes cuando el TXT de retenciones falta: advertencia y se omite el tipo
  (igual que ya ocurre con los otros tipos).
- Pruebas con pytest siguiendo `tests/`.
