# Ventas: conciliación de comprobantes emitidos (SRI vs Odoo) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reemplazar la pestaña Ventas por una conciliación semanal de facturas, NC, ND, liquidaciones y retenciones emitidas: TXT del SRI (mezclados) contra Odoo leído por API en solo lectura.

**Architecture:** Mismo patrón que Compras. Un lector de TXT separa por tipo; `OdooReader` gana `fetch_emitidos` y `fetch_retenciones_emitidas`; el cruce de documentos reutiliza `conciliar_compras.conciliar` y se clasifica en estados propios de emitidos; las retenciones reutilizan `conciliar_retenciones` y `excel_retenciones` (con XML del caché compartido). Un `excel_emitidos.py` arma el libro y `pipeline_ventas.py` lo orquesta en un hilo.

**Tech Stack:** Python, pandas, openpyxl, FastAPI, pytest, XML-RPC (Odoo 14), LibreOffice (recálculo).

**Spec:** `docs/superpowers/specs/2026-10-23-ventas-emitidos-design.md`

## Global Constraints

- Odoo 14 **solo lectura**; credenciales solo por `ODOO_URL`, `ODOO_DB`, `ODOO_USER`, `ODOO_PASSWORD`; los errores nunca incluyen credenciales (usar `OdooError` con mensaje genérico, como `fetch_compras`).
- Sin `verify=False` hacia el SRI; descarga de XML solo vía `xml_cache.descargar_xmls` (usa `sri_session()`).
- TXT del SRI: ISO-8859-1, tabulador; **no depender del nombre del archivo**; `FECHA_EMISION` puede traer hora → usar los primeros 10 caracteres (`dd/mm/aaaa`).
- Tolerancia de importes: $0,02 (`TOL_CENTAVOS = 2`). NC en valor absoluto. NC: `IMPORTE_TOTAL` vacío → total = base + IVA.
- RUC de la empresa: `2390000831001` (constante `RUC_EMPRESA`).
- Excel: Arial, fórmulas (`COUNTIFS`/`SUMIFS`/`INDEX`/`MATCH`) para conteos, sumas y diferencias; el **estado** de cada documento es un valor calculado en Python y testeado; recalcular con LibreOffice (`conciliar_compras._recalcular`); proteger contra inyección de fórmulas con los helpers de `excel_compras.py`. Colores: azul = resumen, naranja = Solo SRI, rojo = Solo Odoo, amarillo = diferencias de valor, violeta = factura incorrecta.
- Caché `cache/xml/` compartido: siempre filtrar por las claves del TXT cargado, nunca el glob completo.
- Guías de remisión y tipos desconocidos: se ignoran con advertencia, nunca error.
- Código sigue el estilo existente: cabecera `# -*- coding: utf-8 -*-` donde ya se usa, comentarios y textos en español, clases/funciones en snake_case.
- Tests: `python -m pytest tests -q` (se corre desde la raíz del repo).

## Review Focus

1. Un envío con un solo tipo (p. ej. solo retenciones, o solo NC): no es error; sin filas de un tipo el resumen muestra 0 y no revienta.
2. La misma clave en dos TXT que se solapan: se deduplica y los conteos no se inflan.
3. TXT con una sola fila de fecha atípica (meses atrás): el rango > 35 días da advertencia, no error.
4. Nombre de cliente de Odoo que empieza con `=`, `+`, `-` o `@`: no se escribe como fórmula en el Excel.
5. Ningún XML de retenciones descargable (SRI caído): las retenciones salen como "XML no disponible" y el job termina `done` con advertencia, no `error`.

---

## File Structure

- Create `emitidos_sri.py` — lectura/normalización de TXT del SRI emitidos y período.
- Modify `odoo_client.py` — `fetch_emitidos`, `fetch_retenciones_emitidas`.
- Modify `xml_cache.py`, `conciliar_retenciones.py` — contraparte (sujeto retenido) en emitidas.
- Create `conciliar_emitidos.py` — estados de documentos y retenciones emitidas.
- Create `excel_emitidos.py` — libro de resultado.
- Modify `pipeline_ventas.py` (reescribir), `app.py`, `templates/index.html`.
- Delete `conciliar_ventas_sri.py`, `conciliar_facturas_emitidas.py` (si nada más los referencia), `tests/test_conciliar_ventas.py`; reemplazar `tests/test_pipeline_ventas.py`.
- Tests nuevos: `tests/test_emitidos_sri.py`, `tests/test_odoo_emitidos.py`, `tests/test_conciliar_emitidos.py`, `tests/test_excel_emitidos.py`.
- Modify `CLAUDE.md` (sección Ventas).

---

### Task 1: Lector de TXT de emitidos

**Files:**
- Create: `emitidos_sri.py`
- Test: `tests/test_emitidos_sri.py`

**Interfaces:**
- Produces:
  - `RUC_EMPRESA = "2390000831001"`, `TIPOS_DOC = ("FAC", "NC", "ND", "LIQ")`
  - `leer_emitidos_txts(rutas: list[str]) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]` → `(docs, ret, advertencias)`.
    - `docs` columnas: `clave, serie, fecha (Timestamp), tipo (FAC|NC|ND|LIQ), base, iva, total, ruc (NA), proveedor (NA), afectada (NA)` (mismo contrato que `conciliar_compras.leer_sri_txts`, para poder llamar a `conciliar`).
    - `ret` columnas `clave, serie, ruc (=RUC_EMPRESA), proveedor (NA), fecha` (contrato `conciliar_retenciones.COLUMNAS_TXT`), solo con clave de 45–49 dígitos completada a 49.
  - `periodo_emitidos(fechas: pd.Series) -> tuple[date, date, list[str]]` → `(desde, hasta, advertencias)`; `desde/hasta` = min/max de fechas; advertencia si el rango supera 35 días.

- [ ] **Step 1: Write the failing tests** en `tests/test_emitidos_sri.py` (archivos escritos con `encoding="iso-8859-1"`, nombres arbitrarios como `2390000831001_EmitidosAutorizados (1).txt`):
  - `test_separa_tipos_por_columna_comprobante`: un TXT mezclando Factura, "Nota de Crédito", "Nota de Débito", "Liquidación de Compra" y "Comprobante de Retención" → `docs.tipo` contiene FAC/NC/ND/LIQ, `ret` tiene 1 fila; no depende del nombre del archivo.
  - `test_nc_total_vacio_es_base_mas_iva`: NC con base `7.72`, IVA `1.16`, total vacío → `total == 8.88`.
  - `test_fecha_con_hora`: `01/09/2026 00:00:00` → `Timestamp(2026, 9, 1)`.
  - `test_duplicados_entre_txt`: la misma clave en dos archivos → una sola fila; si los datos difieren, una advertencia que menciona la clave.
  - `test_guia_remision_se_ignora_con_advertencia`: fila "Guía de Remisión" → no está en `docs` y la advertencia dice cuántas se ignoraron.
  - `test_faltan_columnas`: TXT sin `CLAVE_ACCESO` → `ValueError` que nombra el archivo y las columnas.
  - `test_solo_retenciones`: envío con solo retenciones → `docs` vacío (con las columnas), `ret` con filas, sin error.
  - `test_periodo_rango_largo`: fechas del 1 al 15 de septiembre y una de julio → advertencia; `desde`/`hasta` son min/max.

- [ ] **Step 2: Run** `python -m pytest tests/test_emitidos_sri.py -q` → FAIL (módulo inexistente).

- [ ] **Step 3: Implement** `leer_emitidos_txts` y `periodo_emitidos` en `emitidos_sri.py`. Detectar el tipo con prefijos normalizados sin acentos/mayúsculas (`factura`, `nota de cr`, `nota de d`, `liquidaci`, `comprobante de reten`, `gu` → guía); el "�" de lecturas mal codificadas no debe romper el prefijo (comparar solo hasta antes del carácter acentuado). Dedupe por `CLAVE_ACCESO` conservando la primera fila.

- [ ] **Step 4: Run** `python -m pytest tests/test_emitidos_sri.py -q` → PASS.

- [ ] **Step 5: Commit** `feat: lector de TXT de comprobantes emitidos`.

---

### Task 2: Odoo — documentos emitidos

**Files:**
- Modify: `odoo_client.py` (nuevas constantes y método en `OdooReader`)
- Test: `tests/test_odoo_emitidos.py` (usar el mismo estilo de doble de XML-RPC que `tests/test_odoo_client.py`)

**Interfaces:**
- Consumes: `OdooReader._search_read_todo`, `_call`, `_conectar`, `_txt`, `_clave`, `OdooError`.
- Produces: `COLUMNAS_EMITIDOS = ["id","num","nfac","socio","ruc","fecha","tipo","clave","base","iva","total","afectada","estado","electronico","elegible"]` y `OdooReader.fetch_emitidos(desde: date, hasta: date, margen_dias: int = 5) -> pd.DataFrame`.
  - `tipo` ∈ FAC/NC/ND/LIQ. `estado` = `state` de Odoo. `electronico` = `x_is_edocument`. `elegible` = `estado in ("posted","cancel")`.
  - Importes en valor absoluto; `afectada` = `x_number` de `reversed_entry_id` solo para NC.

- [ ] **Step 1: Failing tests:**
  - `test_fetch_emitidos_tipos`: `out_invoice` código 01 → FAC; `out_refund` 04 → NC; `out_invoice` 05 → ND; `in_invoice` con código de liquidación → LIQ; `in_invoice` tipo factura (compra normal) → **excluido**.
  - `test_fetch_emitidos_estados`: se leen `posted` y `cancel` y `draft`; `draft` tiene `elegible False`; `cancel` tiene `elegible True`.
  - `test_fetch_emitidos_incluye_no_electronicos`: un `out_invoice` con `x_is_edocument False` aparece con `electronico False`.
  - `test_fetch_emitidos_error_sin_credenciales`: una excepción del servidor produce `OdooError` cuyo texto no contiene la URL ni la contraseña.
  - `test_fetch_emitidos_vacio`: sin movimientos → DataFrame vacío con `COLUMNAS_EMITIDOS`.

- [ ] **Step 2: Run** `python -m pytest tests/test_odoo_emitidos.py -q` → FAIL.

- [ ] **Step 3: Implement.** Dominio: `move_type in (out_invoice, out_refund, in_invoice)`, `state in (posted, cancel, draft)`, `invoice_date` en `[desde-margen, hasta+margen]`; filtrar en Python los `in_invoice` cuyo código de `account.receipt.type` no esté en `CODIGOS_LIQUIDACION`. **Antes de fijar `CODIGOS_LIQUIDACION`, consultar en solo lectura el campo `code` de `account.receipt.type` para los ids 3 y 15** (ya verificado: ambos son liquidaciones) y dejar la constante con esos códigos; dejar el mapeo código→tipo `{"01":"FAC","04":"NC","05":"ND"}` + liquidación en una sola constante.

- [ ] **Step 4: Run** tests → PASS; correr también `python -m pytest tests/test_odoo_client.py -q` (sin regresiones).

- [ ] **Step 5: Commit** `feat: lectura de documentos emitidos desde Odoo`.

---

### Task 3: Retenciones emitidas — Odoo, XML y contraparte

**Files:**
- Modify: `odoo_client.py` (generalizar `_fetch_retenciones` por tipo), `xml_cache.py` (`lineas_sri`), `conciliar_retenciones.py` (nada, salvo si un test lo exige)
- Test: ampliar `tests/test_odoo_emitidos.py` y `tests/test_xml_cache.py`

**Interfaces:**
- Consumes: `OdooReader._fetch_retenciones` (hoy fijo a `out_retention`), `xml_cache.lineas_sri(claves, cache_dir)`, `parse_comprobante_xml`.
- Produces:
  - `OdooReader.fetch_retenciones_emitidas(desde, hasta, margen_dias=5) -> pd.DataFrame` con `COLUMNAS_RET` + `clave` (de `auth_retention_number`); filtro `type = in_retention`; **`ruc` siempre `RUC_EMPRESA`** y `socio` = nombre del proveedor retenido; todos los estados.
  - `lineas_sri(claves, cache_dir, contraparte: bool = False)`; con `contraparte=True`, `proveedor` = razón social del **sujeto retenido** (el `ruc` sigue siendo el del emisor = empresa).
  - `fetch_retenciones` (recibidas) mantiene exactamente su salida actual.

- [ ] **Step 1: Failing tests:**
  - `test_fetch_retenciones_emitidas_usa_in_retention`: el dominio enviado contiene `("type","=","in_retention")`; el resultado trae `clave` y `ruc == "2390000831001"` aunque el proveedor tenga otro RUC.
  - `test_fetch_retenciones_recibidas_sin_cambios`: la salida de `fetch_retenciones` tiene exactamente `COLUMNAS_RET` (sin `clave`).
  - `test_lineas_sri_contraparte`: con un XML v2.0 de muestra (usar `tests/_helpers._autorizacion`), `contraparte=True` deja `proveedor` = razón social del sujeto retenido; con `False` conserva el comportamiento actual.

- [ ] **Step 2: Run** los tests nuevos → FAIL.

- [ ] **Step 3: Implement.** Refactor mínimo: `_fetch_retenciones(desde, hasta, margen_dias, tipo, emitidas=False)`; si `emitidas`, pedir además `auth_retention_number`. En `lineas_sri`, la razón social del sujeto retenido está en la cabecera que devuelve `parse_comprobante_xml` (`Razon_Social_Sujeto_Retenido`); si el `detalle` no la trae, tomarla de `cab`.

- [ ] **Step 4: Run** `python -m pytest tests/test_odoo_emitidos.py tests/test_odoo_client.py tests/test_xml_cache.py -q` → PASS.

- [ ] **Step 5: Commit** `feat: retenciones emitidas desde Odoo y contraparte en líneas SRI`.

---

### Task 4: Clasificación de estados (documentos y retenciones)

**Files:**
- Create: `conciliar_emitidos.py`
- Test: `tests/test_conciliar_emitidos.py`

**Interfaces:**
- Consumes: `conciliar_compras.conciliar(sri, odoo, desde, hasta)` (columnas `idx_sri, id_odoo, tipo, metodo, cat`); `conciliar_compras.CAT_DUP`, `_mismo_total`; `conciliar_retenciones.conciliar_retenciones(...)` y constantes `EST_*`.
- Produces (constantes de estado): `EST_OK="Coincide"`, `EST_SOLO_SRI="Solo en SRI"`, `EST_SOLO_ODOO="Solo en Odoo"`, `EST_DIF_VALOR="Diferencia de valor"`, `EST_DIF_DATOS="Diferencia de datos"`, `EST_ANULADO="Odoo anulado, SRI autorizado"`, `EST_DUP="Duplicado en Odoo"`, `EST_INFO="Informativo"`.
  - `clasificar_documentos(sri: pd.DataFrame, odoo: pd.DataFrame, desde: date, hasta: date) -> pd.DataFrame` — una fila por pareja o lado sin pareja, columnas `estado, tipo, idx_sri, id_odoo, metodo, detalle` (texto corto con la causa, p. ej. "fecha SRI 01/09 ≠ Odoo 02/09").
  - `conciliar_retenciones_emitidas(sri_lineas, odoo_ret, sri_txt, desde, hasta) -> pd.DataFrame` — igual a `conciliar_retenciones` pero con `EST_ANULADO` donde corresponde.

Reglas de `clasificar_documentos` (spec §"Qué se concilia"):
  - Pareja con Odoo `estado == "cancel"` → `EST_ANULADO` (sin importar importes).
  - Pareja con totales que difieren > $0,02 → `EST_DIF_VALOR`.
  - Pareja con importes iguales pero (clave distinta o ausente en Odoo, o fecha distinta, o `electronico False`) → `EST_DIF_DATOS`.
  - Pareja con `cat == CAT_DUP` → `EST_DUP`.
  - SRI sin pareja → `EST_SOLO_SRI`.
  - Odoo sin pareja: `cancel` → `EST_INFO` (anulado correctamente); `electronico False` o `draft` → `EST_INFO`; publicado y electrónico con fecha en `[desde, hasta]` → `EST_SOLO_ODOO`; fuera del período → no se lista.
  - Para `conciliar()`, el SRI debe llevar `ruc` NA (no hay RUC de cliente en el TXT; `_ruc_compatible` ya lo tolera).

- [ ] **Step 1: Failing tests** (DataFrames mínimos construidos a mano):
  - `test_coincide_por_clave`, `test_diferencia_de_valor_mayor_a_2_centavos`, `test_diferencia_dentro_de_tolerancia_coincide` (diferencia 0,02 → Coincide), `test_diferencia_de_fecha_es_diferencia_de_datos`, `test_clave_ausente_en_odoo_cruza_por_numero_y_es_diferencia_de_datos`, `test_odoo_cancelado_con_sri_es_anulado`, `test_odoo_cancelado_sin_sri_es_informativo`, `test_solo_en_sri`, `test_solo_en_odoo_dentro_y_fuera_del_periodo`, `test_nc_compara_valor_absoluto`, `test_nota_debito_sintetica` (ND con tipo "ND" coincide igual que una factura), `test_duplicado_en_odoo`.
  - `test_retencion_cancelada_en_odoo_con_sri_es_anulada`: SRI con líneas de la retención `001-002-000000001`; Odoo con esa retención `cancel` → un solo documento con `EST_ANULADO` (no "Solo en SRI" más "Informativo").
  - `test_retencion_sin_xml`: clave del TXT sin líneas → `EST_XML` (comportamiento heredado) y no "Solo en SRI".

- [ ] **Step 2: Run** `python -m pytest tests/test_conciliar_emitidos.py -q` → FAIL.

- [ ] **Step 3: Implement** ambas funciones. Para las retenciones: llamar a `conciliar_retenciones` y, sobre su resultado, reetiquetar como `EST_ANULADO` las filas "Solo en SRI" cuyo `n_ret` exista en `odoo_ret` con `estado == "cancel"`, eliminando la fila `EST_INFO` duplicada de esa misma retención.

- [ ] **Step 4: Run** tests → PASS, más `python -m pytest tests/test_conciliar_compras.py tests/test_conciliar_retenciones.py -q` sin regresiones.

- [ ] **Step 5: Commit** `feat: estados de conciliación para comprobantes emitidos`.

---

### Task 5: Excel de resultado

**Files:**
- Create: `excel_emitidos.py`
- Test: `tests/test_excel_emitidos.py`

**Interfaces:**
- Consumes: `clasificar_documentos` (Task 4), helpers de `excel_compras.py` (`_hoja`, `_proteger`, `_nan`, `_fecha`, `COLORES`), `excel_retenciones.agregar_hojas_retenciones(wb, ret)` y `fila_resumen_retenciones`, `conciliar_compras._recalcular`.
- Produces: `construir_excel_emitidos(sri, odoo, docs, ret, desde, hasta, salida: str) -> dict` que escribe `salida` y devuelve `{"documentos": {...}}`: por cada tipo `FAC/NC/ND/LIQ/RET` un dict `{"total", "<estado>": n, ...}` (RET cuenta documentos, no líneas). `ret` puede ser `None`.

Hojas: `Resumen` (una fila por tipo, columnas por estado con `COUNTIFS` sobre `Conciliacion`; la fila de retenciones con `COUNTIFS` sobre `Ret_Documentos` vía `fila_resumen_retenciones`), `Conciliacion` (una fila por documento: Tipo, Estado, Número, Fecha_SRI, Fecha_Odoo, Cliente, Total_SRI, Total_Odoo, Diferencia = fórmula, Detalle), `Solo_en_SRI`, `Solo_en_Odoo`, `Diferencias`, `Anulados_Odoo`, y las hojas `Ret_*` si `ret` no es `None`.

- [ ] **Step 1: Failing tests** (abrir con `openpyxl`, `data_only=False`):
  - `test_hojas_y_resumen`: libro contiene las hojas listadas; `Resumen` tiene celdas con `COUNTIFS`.
  - `test_diferencia_es_formula`: la columna Diferencia es fórmula (`=` …), no número.
  - `test_nombre_cliente_con_signo_igual_no_es_formula`: un cliente `"=1+1"` se guarda como texto (no `data_type == "f"`).
  - `test_sin_retenciones`: `ret=None` → no existen hojas `Ret_*`; el resumen no incluye la fila RET.
  - `test_envio_con_un_solo_tipo`: solo NC → el resumen devuelve 0 para los demás tipos sin error.
  - `test_fuente_arial`: la celda A1 de `Resumen` usa Arial.

- [ ] **Step 2: Run** → FAIL.
- [ ] **Step 3: Implement** `construir_excel_emitidos` siguiendo la estructura de `excel_compras.construir_excel`; las hojas filtradas son copias de las filas de `Conciliacion` por estado (valores), con el color de la convención.
- [ ] **Step 4: Run** `python -m pytest tests/test_excel_emitidos.py -q` → PASS.
- [ ] **Step 5: Commit** `feat: Excel de conciliación de emitidos`.

---

### Task 6: Pipeline de Ventas

**Files:**
- Modify (reescribir): `pipeline_ventas.py`
- Replace: `tests/test_pipeline_ventas.py`

**Interfaces:**
- Consumes: Tasks 1–5; `OdooReader.from_env()`, `xml_cache.descargar_xmls/lineas_sri`, `job_manager.write_status/schedule_cleanup/get_job_dir`; patrón de `pipeline_compras.py` (monkeypatch de `schedule_cleanup`, `CACHE_XML_DIR`).
- Produces: `run_pipeline_ventas(job_id: str) -> None` — lee `jobs/<id>/emitidos/*.txt`, escribe `jobs/<id>/resultado.xlsx` y `write_status(..., resumen={"por_tipo": ..., "periodo": [desde, hasta], "advertencias": [...], "retenciones_fallidos": n})`.

Orden (falla rápido): leer TXT → período (`periodo_emitidos`) → `OdooReader.from_env()` → leer Odoo de documentos y retenciones → si hay retenciones, descargar XML (progreso 0,15–0,65) → clasificar → Excel → `done`. Cualquier excepción → `error` con traceback y `schedule_cleanup`, como Compras. Advertencias: tipos ignorados, rango > 35 días, XML no disponibles, LibreOffice no disponible.

- [ ] **Step 1: Failing tests** (lector de Odoo y descarga de XML sustituidos con dobles, como en `tests/test_pipeline_compras.py`):
  - `test_pipeline_done_con_todos_los_tipos`: estado `done`, `resultado.xlsx` existe, `resumen.por_tipo` incluye FAC/NC/LIQ/RET.
  - `test_pipeline_sin_txt_validos`: TXT sin ningún tipo soportado → estado `error` con mensaje claro.
  - `test_pipeline_odoo_falla_antes_de_descargar_xml`: `OdooError` → `error` y `descargar_xmls` no se llamó.
  - `test_pipeline_xml_todos_fallidos`: `descargar_xmls` devuelve todo fallido → `done`, `resumen.retenciones_fallidos == n` y advertencia.
  - `test_pipeline_solo_documentos_no_descarga_xml`: sin retenciones en los TXT no se llama a `descargar_xmls`.

- [ ] **Step 2: Run** `python -m pytest tests/test_pipeline_ventas.py -q` → FAIL.
- [ ] **Step 3: Implement** `run_pipeline_ventas`.
- [ ] **Step 4: Run** → PASS.
- [ ] **Step 5: Commit** `feat: pipeline de Ventas con Odoo por API`.

---

### Task 7: Ruta, interfaz y limpieza del flujo antiguo

**Files:**
- Modify: `app.py` (ruta `/upload-ventas`, importaciones), `templates/index.html` (formulario y resultado de Ventas)
- Modify: `tests/test_app.py` (pruebas de `/upload-ventas`)
- Delete: `conciliar_ventas_sri.py`, `conciliar_facturas_emitidas.py` (solo si `grep` confirma que nada más los importa), `tests/test_conciliar_ventas.py`

**Interfaces:**
- Consumes: `run_pipeline_ventas` (Task 6).
- Produces: `POST /upload-ventas` con campo multi-archivo `emitidos` (1 o más `.txt`, cualquier nombre, máximo 60 archivos); sin campo `odoo_venta`. Responde `{"job_id": ...}`. El nombre de descarga de Ventas se mantiene como hoy (`Conciliacion_Ventas.xlsx`; verificar en `test_download_ventas_name`).

- [ ] **Step 1: Failing tests** en `tests/test_app.py`:
  - `test_upload_ventas_acepta_txt_con_nombre_del_portal`: sube `2390000831001_EmitidosAutorizados (1).txt` → 200 y el archivo queda en `jobs/<id>/emitidos/`.
  - `test_upload_ventas_rechaza_no_txt`: archivo `.xlsx` → 400.
  - `test_upload_ventas_ya_no_pide_excel_de_odoo`: POST solo con `emitidos` → 200.
  - Mantener `...returns_401_without_session`.
- [ ] **Step 2: Run** `python -m pytest tests/test_app.py -q` → FAIL.
- [ ] **Step 3: Implement** la ruta (basename del nombre para evitar path traversal; prefijo numérico si hay nombres repetidos; sin renombrar a `emitidos_*`). En `index.html`: formulario Ventas con un único selector de TXT y texto "Sube todos los TXT de emitidos de la semana (facturas, notas de crédito, notas de débito, liquidaciones y retenciones). Odoo se lee automáticamente."; el resultado muestra una fila por tipo con sus estados y las advertencias, reutilizando el estilo de `showResultCompras`.
- [ ] **Step 4: Run** `python -m pytest tests -q` → PASS completo.
- [ ] **Step 5: Commit** `feat: pestaña Ventas con TXT mezclados y Odoo por API; retira flujo antiguo`.

---

### Task 8: Documentación y verificación manual

**Files:**
- Modify: `CLAUDE.md` (sección de Ventas/emitidos: tabla de archivos, reglas de estados, límite de 30 días del XML, constante `RUC_EMPRESA`, que Ventas ya no usa `reten_odoo`/`odoo_venta.xlsx`).

- [ ] **Step 1:** Actualizar `CLAUDE.md`; marcar `conciliar_emitidas.py` como CLI independiente que se conserva.
- [ ] **Step 2:** Correr la app local con los TXT reales de `documentos_electronicos/emitidas/` (facturas, NC, LIQ de septiembre) y las credenciales del `.env`; abrir el Excel y verificar contra el recuento conocido: 01/09/2026 → 7 liquidaciones SRI = 7 en Odoo (todas `Coincide`).
- [ ] **Step 3:** `python -m pytest tests -q` → todo verde.
- [ ] **Step 4: Commit** `docs: Ventas emitidos en CLAUDE.md`.

---

## Self-Review

- **Spec coverage:** tipos y entradas (T1), Odoo documentos/retenciones (T2–T3), estados y decisiones 3/3b/3c (T1 dedupe, T4), período + advertencia 35 días (T1, T6), salida Excel (T5), ruta/UI/reemplazo y retiro de scripts (T7), documentación (T8). Guías ignoradas (T1). ND sintética (T4). Límite 30 días XML (T6 test).
- **Consistencia de tipos:** `docs`/`sri` de T1 cumple el contrato de `conciliar`; `COLUMNAS_EMITIDOS` (T2) incluye `id, tipo, clave, nfac, ruc, total, afectada, fecha, elegible` que `conciliar` lee; `ret` de T1 cumple `COLUMNAS_TXT`; `RUC_EMPRESA` se usa igual en T1 y T3.
- **Riesgos conocidos para el ejecutor:** (a) código de `account.receipt.type` de los ids 3/15 se confirma en T2; (b) `clave` de retenciones en Odoo viene de `auth_retention_number`; (c) si `parse_comprobante_xml` no entrega la razón social del sujeto retenido en el detalle, tomarla de la cabecera (T3).
