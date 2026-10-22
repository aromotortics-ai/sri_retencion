# Conciliación de todos los comprobantes electrónicos recibidos (SRI vs Odoo 14)

Fecha: 2026-10-22 · Estado: borrador para aprobación

## Intención y por qué

Hoy el personal baja del portal del SRI, uno por uno, los comprobantes recibidos
(facturas, notas de crédito, notas de débito, liquidaciones de compra) y los
contrasta contra Odoo a mano. Las retenciones recibidas ya están automatizadas
en la app web (FastAPI, Easypanel). Este trabajo completa el resto de tipos y
quita dos pasos manuales: ya no se exporta Excel desde Odoo y el resultado sale
como un solo Excel de observaciones.

Éxito: con un clic por período, el usuario sube los TXT del SRI (uno por tipo
de comprobante) y recibe un Excel con las diferencias SRI vs Odoo de todos los tipos.

## Decisiones del usuario

- **Opción 1, semi-manual**: el usuario descarga del portal SRI el TXT de cada
  tipo ("Exportar Txt") y los sube. No se automatiza el login al SRI: la pantalla de
  consulta tiene reCAPTCHA y no se va a eludir. Por eso **no se guardan credenciales del SRI**.
- **Odoo por API**: la app se conecta a Odoo 14 CE (`https://aromotor.com`, base
  `aromotor`) y lee los documentos del período; ya existe usuario de API y el usuario
  lo usa en otras aplicaciones. Solo lectura.
- **Solo documentos electrónicos** se comparan contra el SRI.
- La vista diaria en pantalla (conciliar día a día) queda **fuera de alcance** de esta
  fase; se retoma después.

## Alcance

Tipos de comprobante del portal: Factura (1), Liquidación de compra (2),
Nota de crédito (3), Nota de débito (4), Comprobante de retención (6, ya existente).

Flujo nuevo:
1. El usuario elige el período y sube los TXT del SRI (hasta 5, detectando el tipo por la
   columna `TIPO_COMPROBANTE`).
2. Facturas: el TXT ya trae base, IVA y total, **no hace falta bajar XML**. Para notas de
   crédito/débito y liquidaciones, bajar el XML (con el caché compartido existente) solo
   si el TXT no trae los montos (en retenciones llegan vacíos; hay que verificar en los
   otros tipos).
3. Leer de Odoo por API los documentos de compra del período.
4. Conciliar y entregar un Excel.

## Lo ya verificado con datos reales (septiembre 2026)

- TXT de facturas: 448 filas, con `VALOR_SIN_IMPUESTOS`, `IVA`, `IMPORTE_TOTAL` llenos.
- Excel de Odoo (16 columnas: número, número de factura, socio, sucursal, fecha,
  subtotales, importe sin impuestos con signo, total con signo, estado, número de
  autorización): 409 filas de septiembre. No trae el RUC del proveedor.
- El número de autorización de Odoo equivale a la clave de acceso del SRI (49 dígitos).
  46 filas lo tienen con 48 dígitos porque Excel borró el cero inicial: rellenar con ceros.
- Por clave de acceso exacta coinciden solo 241 de 409. Una parte son documentos que
  nunca estarán en "recibidos" (ver siguiente punto) y otra son claves mal digitadas en Odoo
  (47 facturas del SRI existen en Odoo por número de factura con otra autorización).
- Prefijos de Odoo (por el número interno): `CE` compras normales; `LEAR/LECO/LETC` liquidaciones
  de compra **emitidas por la empresa**; `CEXAR` importaciones; `CFAR/CFCO/CFTE/VBICO` recibos y
  facturas sin autorización electrónica; `NTAR` notas de crédito (solo 18 de 360 con autorización).

Ya existe `conciliar_recibidas.py` (sin versionar) que concilia facturas recibidas
desde el Excel de Odoo con tres niveles de cruce: clave de acceso, N° factura + importe,
N° factura con importe distinto. Esta fase lo toma como base.

## Reglas de conciliación

- Solo entran al cruce los documentos de Odoo electrónicos: los que tienen clave de
  acceso de 45 a 49 dígitos (tras completar ceros) y no son liquidaciones propias.
  Mi llamado: las liquidaciones propias (`LEAR/LECO/LETC`) se tratan como "no aplica"
  porque las emite la empresa y no aparecen en recibidos.
- Los no electrónicos (`CEXAR`, `CFAR`, etc. sin clave) se listan como informativos,
  no como error.
- Estados de resultado: Coincide, Solo en SRI, Solo en Odoo, Diferencia de valor
  (tolerancia $0,02), Clave de acceso distinta en Odoo (mismo número, otra clave),
  Duplicado en Odoo. Por tipo de comprobante.
- Las notas de crédito se comparan por valor absoluto (Odoo las guarda con signo).

## Entrega

Un Excel con: resumen por tipo y por mes, conciliación completa, Solo en SRI,
Solo en Odoo, Diferencia de valor, Duplicados. Reglas ya vigentes del proyecto:
Arial, fórmulas (no valores pegados), colores por hoja, recalcular con LibreOffice
antes de entregar.

## Restricciones

- Solo lectura sobre Odoo, usuario de API dedicado, credenciales por variables de
  entorno del contenedor, nunca en el repositorio ni en logs.
- El caché de XML es compartido entre períodos: filtrar siempre por las claves del TXT
  cargado (bug ya corregido para retenciones).
- Misma autenticación de la app actual (cookie firmada).
- No romper las pestañas existentes: Retenciones y Ventas.

## No objetivos y antiobjetivos

- No automatizar el login ni el captcha del portal del SRI.
- No escribir nada en Odoo.
- No vista diaria en esta fase.
- Antiobjetivo: que la conciliación marque como error documentos que por naturaleza
  no están en el SRI (liquidaciones propias, importaciones, recibos).

## Llamados míos tras la revisión (pendientes de que los revises)

- **Cruce de facturas**: se mantienen los tres niveles de `conciliar_recibidas.py`
  (clave de acceso; N° factura + importe; N° factura con importe distinto). Con la API se
  añade el RUC del proveedor para desempatar. "Mismo número" = establecimiento-punto-secuencial.
- **Electrónico = tiene clave de acceso** de 45 a 49 dígitos (completada con ceros). En los datos reales,
  de 360 `NTAR` de septiembre solo 18 tienen clave; las demás son ajustes internos con Aduana,
  agentes de carga o la propia empresa, no notas de crédito de proveedores. Quedan como informativas.
- **Qué se lee de Odoo**: solo asientos publicados (`posted`), por fecha de factura; los emitidos
  el último día del mes o el primero del siguiente se marcan para revisar, no como error.
- **TXT faltante**: advertencia; el tipo sin TXT se omite y el resumen lo indica. No bloquea.
- **Subida**: un solo envío con hasta 5 archivos.

## Preguntas resueltas el 2026-10-22 (verificado en datos reales, solo lectura)

- **Liquidaciones de compra: fuera de alcance** por decisión del usuario. Los tipos son Factura,
  Nota de crédito, Nota de débito y Retención (esta ya existe).
- **TXT de notas de crédito (30 filas) y de débito (1 fila, septiembre 2026)**: traen
  `VALOR_SIN_IMPUESTOS` e `IVA`; `IMPORTE_TOTAL` viene vacío (total = base + IVA) y
  `NUMERO_DOCUMENTO_MODIFICADO` trae el número real de la factura afectada. **No se necesita XML**.
  TXT en `documentos_electronicos/recibidas/{facturas,NC,ND,retenciones}_recibidas/`.
- **API de Odoo 14.0 verificada** (credenciales en `.env`: `ODOO_URL`, `ODOO_DB`, `ODOO_USER`,
  `ODOO_PASSWORD`, ignorado por git). El usuario puede leer `account.move` y `res.partner.vat`.
  Septiembre 2026, asientos publicados: 409 `in_invoice` (coincide con el Excel) y 25 `in_refund`.
- **Campos de `account.move`**: clave de acceso de documentos recibidos = `x_authorization_number`
  (49 dígitos; `x_access_key` solo se llena en documentos propios); número de documento = `x_number`
  (a confirmar); tipo de documento = `x_receipt_type_id` (códigos SRI: 01 factura, 04 nota de crédito,
  05 nota de débito, 03 liquidación, 15 reembolso, 02 nota de venta, 12 exterior, 16 DAU, 00 no tributable);
  `x_is_edocument` = **emitido electrónicamente por la empresa** (liquidaciones y reembolsos), no
  "recibido electrónico": no sirve para filtrar recibidos.
- **Notas de crédito en Odoo** son `in_refund` con prefijo `RCE/`, `ref` = "Reversa de: <factura>"
  y `reversed_entry_id` apunta a la factura afectada; sirve para cruzar contra
  `NUMERO_DOCUMENTO_MODIFICADO` del SRI. Los `NTAR` del Excel son otra cosa (ajustes internos).
- **Regla de elegibilidad en Odoo** (mi llamado): comparan contra el SRI los asientos con tipo
  de comprobante 01, 04 o 05 y clave de 45 a 49 dígitos; los tipos 02, 12, 16, 00 y 03/15 propios
  se listan como informativos.

## Lo que queda al constructor

- Integración exacta: nueva pestaña "Compras" en la app, o ampliar la existente.
- Modelo de Odoo y campos para leer (en Odoo 14 el tipo es `move_type`; la localización
  ecuatoriana tiene campos propios, p.ej. número de autorización y `account.move.retention`).
  Traer el RUC del proveedor (`partner_id.vat`) para reforzar el cruce.
- Si la API falla o no está configurada: respaldo subiendo el Excel de Odoo (mi llamado).
- Forma de detectar el tipo de cada TXT subido y de reportar TXT faltantes del período.
- Pruebas: pytest, siguiendo `tests/` actuales.
