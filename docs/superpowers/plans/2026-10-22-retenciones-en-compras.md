# Retenciones recibidas en la pestaña Compras — Plan de implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Que la pestaña Compras concilie también las retenciones recibidas (TXT del SRI + XML por clave contra `account.move.retention` de Odoo por API) y entregue un único Excel con compras y retenciones.

**Architecture:** Módulos nuevos y pequeños que no tocan los de la pestaña Retenciones antigua (`pipeline.py`, `conciliar_odoo.py`, `build_workbook_final.py` quedan intactos): `odoo_client.fetch_retenciones` (lectura), `xml_cache.py` (descarga con caché), `conciliar_retenciones.py` (lógica pura sobre DataFrames), `excel_retenciones.py` (hojas `Ret_*` agregadas al libro de Compras). `pipeline_compras.py` orquesta; `construir_excel` recibe el resultado de retenciones opcional.

**Tech Stack:** Python, pandas, openpyxl, XML-RPC (`xmlrpc.client`), requests vía `sri_session`, FastAPI, pytest.

**Spec:** `docs/superpowers/specs/2026-10-22-retenciones-en-compras-design.md` (el plan argumenta desde ella; ante conflicto manda la spec).

## Global Constraints

- Odoo **solo lectura**: únicamente `authenticate`, `fields_get`, `search_read`, `read`, `search_count`. Nunca `create/write/unlink` ni métodos de negocio.
- Credenciales solo por `ODOO_URL`, `ODOO_DB`, `ODOO_USER`, `ODOO_PASSWORD` (vía `OdooReader.from_env()`); nunca en logs, errores, Excel ni repositorio. Errores de Odoo con mensaje genérico `from None`.
- Retenciones recibidas = `account.move.retention` con `type = 'out_retention'`. `in_retention` (emitidas) nunca entra.
- Clave de cruce por línea: `(N° retención, Factura sustento, Porcentaje)`; tolerancia `0.02` en valor retenido.
- Caché `cache/xml/` compartido: filtrar **siempre** por las claves del TXT cargado; nunca parsear el glob completo.
- Datos de negocio (TXT/XLSX/XML, `jobs/`, `cache/`, `documentos_electronicos/`) nunca versionados; los tests usan datos sintéticos.
- Excel: Arial, fórmulas (no valores calculados pegados), colores por hoja (azul resumen, naranja Solo SRI, rojo Solo Odoo, amarillo diferencias, violeta factura incorrecta, gris informativos), `_proteger` contra inyección de fórmulas en texto externo, `fullCalcOnLoad`.
- No modificar la pestaña/pipeline antiguos de Retenciones. No eludir captchas ni controles del SRI.
- Tests: `python -m pytest tests -q`. Fallan 4 tests ya conocidos en `main` limpio (`tests/test_app.py::test_upload_creates_job` y 3 en `tests/test_build_workbook.py`, por archivos de datos no versionados): no son de este trabajo; no los "arregles".
- Commits con el trailer `Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>`.

## Review Focus

Entradas/condiciones que la spec implica y que ningún test obvio cubre; cada una tiene su test en la tarea dueña:

1. Retención con XML no descargado: no debe aparecer como "Solo en Odoo" (Tarea 3).
2. Retención de Odoo cancelada o en borrador: informativa, nunca error (Tarea 3).
3. Misma `(N° retención, factura, %)` repetida en Odoo o SRI: se suman base y valor, sin filas duplicadas (Tarea 3).
4. Porcentaje distinto en el mismo N° de retención y factura: una sola fila "Diferencia de valor" con ambos porcentajes (Tarea 3).
5. Subida con solo TXT de retenciones (sin FAC/NC/ND), y subida sin TXT de retenciones: ambos válidos (Tarea 5).
6. Texto de proveedor que empieza con `=`/`+`/`-`/`@` en las hojas `Ret_*` (Tarea 4).
7. Caché con XML de otros meses: no se mezclan (Tarea 2).

---

### Task 1: Lectura de retenciones recibidas de Odoo

**Files:**
- Modify: `odoo_client.py` (añadir constantes y `OdooReader.fetch_retenciones`)
- Test: `tests/test_odoo_client.py` (añadir; usa el `proxy_factory` falso que ya usan los tests de este archivo)

**Interfaces:**
- Consumes: `OdooReader._conectar`, `_call`, `_search_read_todo`, `_txt`, `OdooError` (ya existen).
- Produces: `OdooReader.fetch_retenciones(desde: date, hasta: date, margen_dias: int = 5) -> pd.DataFrame` y la constante `COLUMNAS_RET`. Una fila por **línea de retención**.

`COLUMNAS_RET = ["id_ret", "n_ret", "fecha", "estado", "socio", "ruc", "factura", "pct", "base", "valor", "cod"]` con:
`id_ret` (int, id de `account.move.retention`), `n_ret` (`retention_number`), `fecha` (Timestamp de `retention_date`), `estado` (`posted`/`draft`/`cancel`), `socio`, `ruc` (`res.partner.vat` o NA), `factura` (número `NNN-NNN-NNNNNNNNN` de la factura sustento o NA), `pct` (`tax_percent`, float), `base` (`amount_untaxed`), `valor` (`amount_retention`), `cod` (código SRI = texto antes de ` - ` en el nombre de `tax_id`, ej. `"312"`, o NA).

- [ ] **Step 1: Write the failing tests** en `tests/test_odoo_client.py` con un proxy falso que devuelva, para `account.move.retention`, dos retenciones (una `posted` con 2 líneas, una `cancel` con 1), para `account.move.retention.line` las líneas, para `res.partner` el `vat`, para `account.move` la factura sustento (`name`, `x_number`):
  - `test_fetch_retenciones_una_fila_por_linea`: 3 filas; columnas == `COLUMNAS_RET`.
  - `test_fetch_retenciones_solo_recibidas`: el dominio enviado contiene `("type", "=", "out_retention")` y fechas `retention_date` con el margen de 5 días; no contiene `in_retention`.
  - `test_fetch_retenciones_codigo_y_pct`: `cod == "312"` para `"312 - Transferencia ... (2%)"`; `pct == 2.0`.
  - `test_fetch_retenciones_factura_formato_sri`: `factura == "006-001-000003385"` tanto si el número viene en `x_number` como si solo viene en `name` (con o sin prefijo de texto como `"FAC 006-001-000003385"`; se extrae el patrón 3-3-9).
  - `test_fetch_retenciones_estado_cancelada_se_conserva`: la fila de la cancelada trae `estado == "cancel"` (no se filtra aquí; el filtrado es de la conciliación).
  - `test_fetch_retenciones_error_generico`: si el proxy lanza una excepción cuyo texto contiene la contraseña, el `OdooError` resultante **no** la contiene.
  - `test_fetch_retenciones_vacio`: sin retenciones devuelve DataFrame vacío con `COLUMNAS_RET`.
- [ ] **Step 2: Run** `python -m pytest tests/test_odoo_client.py -q` → Expected: FAIL (`fetch_retenciones` no existe).
- [ ] **Step 3: Implement `OdooReader.fetch_retenciones(self, desde, hasta, margen_dias=5)`** con el mismo envoltorio try/except que `fetch_compras` (mensaje genérico `from None`). Lee `account.move.retention` (`type = out_retention`, `retention_date` en [desde−margen, hasta+margen], **todos los estados**), luego sus líneas por `retention_id in ids`, `res.partner.vat` y las facturas `invoice_id` (`name`, `x_number`) en lotes con `read`/`search_read`; `tax_id` viene como `[id, "312 - ..."]`, el nombre da el código. Sin consultas por fila (N+1).
- [ ] **Step 4: Run** `python -m pytest tests/test_odoo_client.py -q` → Expected: PASS.
- [ ] **Step 5: Verificación contra datos reales (solo lectura, sin imprimir credenciales ni datos de clientes completos):** con un script en el scratchpad que use `OdooReader.from_env().fetch_retenciones(date(2026,9,1), date(2026,9,30))`, imprimir solo conteos: filas, retenciones únicas (`n_ret.nunique()`), filas con `factura` NA, filas con `cod` NA, conteo por `estado`. Expected: ~589 retenciones publicadas en septiembre; `factura` NA cercano a 0. Si `factura` sale NA o con otro formato, ajustar la extracción (el nombre de factura de venta puede estar en `name`).
- [ ] **Step 6: Commit** `git add odoo_client.py tests/test_odoo_client.py` y `git commit -m "feat: lectura de retenciones recibidas de Odoo por API"`.

---

### Task 2: Descarga de XML con caché y líneas del SRI

**Files:**
- Create: `xml_cache.py`
- Test: `tests/test_xml_cache.py`

**Interfaces:**
- Consumes: `sri_session.sri_session()`, `descargar_comprobantes_sri.evaluar_respuesta`, `parse_comprobante_xml.parse_comprobante_xml`, el formateo 3-3-9 que usa `build_workbook_final.py` para `N_Retencion` y `Num_Doc_Sustento` (leerlo allí y reutilizarlo; si es una función privada, extraerla a una función pública mínima en `xml_cache.py` copiando su comportamiento sin editar `build_workbook_final.py`).
- Produces:
  - `descargar_xmls(claves: list[str], cache_dir: Path, on_progress=None, sesion=None, pausa_seg: float = 0.5) -> list[dict]` → descarga solo las claves sin XML en `cache_dir`; devuelve fallidos `[{"clave", "motivo"}]`. `on_progress(hechas: int, total: int, fallidos: int)` cada 10 descargas y al final.
  - `lineas_sri(claves: list[str], cache_dir: Path) -> pd.DataFrame` con columnas `["clave", "n_ret", "fecha", "ruc", "proveedor", "factura", "cod", "pct", "base", "valor"]`: una fila por línea de retención, **solo** de los XML cuya clave está en `claves` y existe en el caché. `n_ret` = `Estab-Pto-Secuencial` (3-3-9), `factura` = `Num_Doc_Sustento` con guiones 3-3-9, `fecha` Timestamp de `Fecha_Emision`.

- [ ] **Step 1: Write the failing tests** (XML sintético de retención v1.0.0 y v2.0.0 escritos en `tmp_path`; `sesion` falsa):
  - `test_lineas_sri_v1_y_v2`: ambos formatos producen filas con `n_ret == "001-200-000004340"`, `factura == "001-003-000088766"`, `pct`, `base`, `valor` correctos.
  - `test_lineas_sri_filtra_por_claves`: un XML de otro mes presente en el caché pero fuera de `claves` no aparece (Review Focus 7).
  - `test_lineas_sri_un_xml_corrupto_no_aborta`: un archivo ilegible se omite y los demás se leen.
  - `test_descargar_xmls_usa_cache`: una clave ya en el caché no llama a `sesion.post`.
  - `test_descargar_xmls_guarda_autorizadas_y_reporta_fallidas`: con una respuesta autorizada se escribe `<clave>.xml`; con una rechazada/vacía queda en `fallidos` con `motivo`; una excepción de red también queda en `fallidos` sin abortar.
  - `test_descargar_xmls_progreso`: `on_progress` se invoca con `total` correcto.
- [ ] **Step 2: Run** `python -m pytest tests/test_xml_cache.py -q` → Expected: FAIL (módulo no existe).
- [ ] **Step 3: Implement** las dos funciones. `descargar_xmls` replica la lógica de `pipeline.py` (SOAP `AutorizacionComprobantesOffline`, header `Connection: close`, `timeout=20`, `evaluar_respuesta`, escritura con la cabecera `<?xml ... standalone="yes"?>`), pero parametrizada; **no editar `pipeline.py`**. `lineas_sri` llama a `parse_comprobante_xml` por archivo dentro de try/except.
- [ ] **Step 4: Run** `python -m pytest tests/test_xml_cache.py -q` → Expected: PASS.
- [ ] **Step 5: Commit** `git add xml_cache.py tests/test_xml_cache.py` y `git commit -m "feat: descarga de XML con caché y líneas de retención del SRI"`.

---

### Task 3: Conciliación de retenciones (lógica pura)

**Files:**
- Create: `conciliar_retenciones.py`
- Test: `tests/test_conciliar_retenciones.py`

**Interfaces:**
- Consumes: DataFrame del SRI con columnas de `xml_cache.lineas_sri`; DataFrame de Odoo con `COLUMNAS_RET` (Tarea 1); `sri_txt` = DataFrame del TXT de retenciones con al menos `clave`, `serie` (N° de retención), `fecha`, `ruc`, `proveedor` (para detectar XML no disponible).
- Produces:
  - `leer_sri_retenciones_txt(rutas: list[str]) -> pd.DataFrame` → filas del/los TXT cuyo `TIPO_COMPROBANTE` empieza con `"Comprobante de Reten"`; columnas `["clave", "serie", "ruc", "proveedor", "fecha"]`; sin duplicados por clave; mismas columnas requeridas y lectura (ISO-8859-1, tab) que `conciliar_compras.leer_sri_txts`.
  - Constantes de estado: `EST_OK = "Coincide"`, `EST_SOLO_SRI = "Solo en SRI"`, `EST_SOLO_ODOO = "Solo en Odoo"`, `EST_DIF = "Diferencia de valor"`, `EST_FAC = "Factura incorrecta en Odoo"`, `EST_XML = "XML no disponible"`, `EST_INFO = "Informativo (Odoo cancelada/borrador)"`.
  - `conciliar_retenciones(sri_lineas, odoo_lineas, sri_txt, desde, hasta) -> pd.DataFrame` con una fila por línea conciliada y columnas `["estado", "n_ret", "fecha", "ruc", "proveedor", "factura_sri", "factura_odoo", "pct_sri", "pct_odoo", "base_sri", "base_odoo", "valor_sri", "valor_odoo", "dif_valor", "id_ret", "clave"]` (NA donde no aplique).

Reglas (spec): (1) agregar primero por clave `(n_ret, factura, pct)` sumando `base` y `valor` en cada lado, solo para filas `posted` de Odoo; (2) cruce exacto: `Coincide` si `|valor_sri − valor_odoo| ≤ 0.02`, si no `Diferencia de valor`; (3) restantes: pasada de **factura incorrecta** (mismo `n_ret`, últimos 9 dígitos de factura y `pct`, distinta factura) → `EST_FAC` una sola fila con ambas facturas; (4) restantes: pasada de **porcentaje distinto** (mismo `n_ret` y `factura`, distinto `pct`; empareja uno a uno por orden de `valor` cercano) → `Diferencia de valor` con ambos `pct`; (5) lo que quede del SRI → `Solo en SRI`; lo que quede de Odoo `posted` → `Solo en Odoo`, **salvo** si su `n_ret` está en `sri_txt["serie"]` pero no tiene ninguna línea SRI (XML faltante) → `EST_XML`; (6) Odoo `cancel`/`draft` → `EST_INFO` siempre; (7) filas de Odoo `posted` cuya `fecha` esté fuera de `[desde, hasta]` y sin par en el SRI no generan fila; las del SRI sí.

- [ ] **Step 1: Write the failing tests** (DataFrames sintéticos mínimos):
  - `test_coincide_dentro_de_tolerancia` (diferencia 0.02 → `Coincide`) y `test_diferencia_de_valor_fuera_de_tolerancia` (0.03 → `Diferencia de valor`, `dif_valor == 0.03`).
  - `test_solo_en_sri` / `test_solo_en_odoo`.
  - `test_factura_incorrecta_una_sola_fila`: `001-003-000088766` vs `002-003-000088766`, mismo n_ret y pct → una fila `EST_FAC` con `factura_sri` y `factura_odoo` distintas.
  - `test_porcentaje_distinto_es_diferencia_de_valor`: SRI 2% valor 11.60, Odoo 1% valor 5.80 → una fila `Diferencia de valor` con `pct_sri == 2.0` y `pct_odoo == 1.0` (Review Focus 4).
  - `test_xml_no_disponible_no_es_solo_odoo`: retención de Odoo cuyo `n_ret` está en `sri_txt` sin líneas SRI → `EST_XML` (Review Focus 1).
  - `test_odoo_cancelada_es_informativa`: nunca `Solo en Odoo` (Review Focus 2).
  - `test_lineas_repetidas_se_suman`: dos líneas iguales en Odoo (mismo n_ret/factura/pct) se concilian como una con valor sumado (Review Focus 3).
  - `test_odoo_fuera_de_periodo_sin_par_no_genera_fila`.
  - `test_entradas_vacias`: SRI vacío y/o Odoo vacío no lanzan excepción y devuelven DataFrame con las columnas esperadas.
  - `test_leer_sri_retenciones_txt`: un TXT sintético ISO-8859-1 con 2 filas de retención y 1 de factura devuelve solo las 2; `serie` conserva el formato `001-158-003714396`.
- [ ] **Step 2: Run** `python -m pytest tests/test_conciliar_retenciones.py -q` → Expected: FAIL.
- [ ] **Step 3: Implement** las dos funciones y constantes en `conciliar_retenciones.py`. Redondear a centavos enteros para la tolerancia (como `_centavos` en `conciliar_compras.py`), sin comparar floats directamente.
- [ ] **Step 4: Run** `python -m pytest tests/test_conciliar_retenciones.py -q` → Expected: PASS.
- [ ] **Step 5: Commit** `git add conciliar_retenciones.py tests/test_conciliar_retenciones.py` y `git commit -m "feat: conciliación de retenciones recibidas SRI vs Odoo"`.

---

### Task 4: Hojas de retenciones en el Excel de Compras

**Files:**
- Create: `excel_retenciones.py`
- Modify: `excel_compras.py` (`construir_excel` acepta `ret=None` y la hoja Resumen suma una fila por retenciones; no cambiar hojas ni fórmulas existentes)
- Test: `tests/test_excel_retenciones.py`; ajustar `tests/test_excel_compras.py` solo si hace falta para la firma nueva

**Interfaces:**
- Consumes: DataFrame de `conciliar_retenciones.conciliar_retenciones` (Tarea 3); `excel_compras._hoja`, `_proteger`, `COLORES`, `_fecha`, `_nan`; `conciliar_recibidas._enc/_fuente`.
- Produces:
  - `excel_retenciones.agregar_hojas_retenciones(wb, ret: pd.DataFrame) -> dict` → crea hojas `Ret_Conciliacion` (azul), `Ret_Solo_SRI` (naranja), `Ret_Solo_Odoo` (rojo), `Ret_Diferencia_Valor` (amarillo, con base y porcentaje de ambos lados), `Ret_Factura_Incorrecta` (violeta), `Ret_Informativos` (gris: `EST_XML` y `EST_INFO`); devuelve contadores `{"coinciden", "diferencia", "solo_sri", "solo_odoo", "factura_incorrecta", "informativos", "total_sri", "total_odoo"}` y escribe en `Ret_Conciliacion` el estado como **fórmula** a partir de `valor_sri`/`valor_odoo` (como `Conciliacion` de compras: `IF(ABS(dif)>0.02, "Diferencia de valor", "Coincide")`).
  - `construir_excel(sri, odoo, pares, desde, hasta, salida, ret=None) -> dict`: si `ret` no es `None`, llama a `agregar_hojas_retenciones` antes de guardar, añade en `Resumen` una fila "Retenciones" con `COUNTIFS` sobre `Ret_Conciliacion` y agrega `"retenciones": contadores` al dict devuelto. Debe seguir funcionando con `sri` vacío (subida solo de retenciones) y con `ret=None` (comportamiento actual intacto).

- [ ] **Step 1: Write the failing tests** (patrón de `tests/test_excel_compras.py`):
  - `test_hojas_ret_y_colores`: existen las 6 hojas `Ret_*` con los colores de pestaña de la spec.
  - `test_ret_estado_es_formula`: la columna de estado de `Ret_Conciliacion` empieza con `=`.
  - `test_ret_resumen_en_hoja_resumen`: `Resumen` contiene la fila de retenciones con fórmulas `COUNTIFS` y los contadores devueltos coinciden con el DataFrame de entrada.
  - `test_ret_texto_peligroso_no_es_formula`: proveedor `"=HYPERLINK(...)"` queda como texto (`data_type == "s"`) en las hojas `Ret_*` (Review Focus 6).
  - `test_ret_fuente_arial`.
  - `test_ret_diferencia_muestra_ambos_porcentajes`.
  - `test_construir_excel_sin_ret_no_cambia`: con `ret=None` no hay hojas `Ret_*` y las hojas previas son las 9 de antes.
  - `test_construir_excel_solo_retenciones`: `sri` vacío + `ret` con filas genera el libro sin error ni fórmulas circulares.
  - Si es posible sin `soffice`: validar fórmulas con el paquete `formulas` (si está instalado; si no, omitir con `pytest.importorskip`) → 0 errores.
- [ ] **Step 2: Run** `python -m pytest tests/test_excel_retenciones.py tests/test_excel_compras.py -q` → Expected: FAIL en los nuevos.
- [ ] **Step 3: Implement** `excel_retenciones.py` y el parámetro `ret` en `construir_excel`. Reusar `_hoja` para encabezado, filtros, anchos, formatos y `_proteger`. Incluir `fullCalcOnLoad` (ya lo pone `construir_excel`).
- [ ] **Step 4: Run** `python -m pytest tests/test_excel_retenciones.py tests/test_excel_compras.py -q` → Expected: PASS (todos, incluidos los previos).
- [ ] **Step 5: Commit** `git add excel_retenciones.py excel_compras.py tests/` y `git commit -m "feat: hojas de retenciones en el Excel de Compras"`.

---

### Task 5: Orquestación en `pipeline_compras` y `conciliar_compras`

**Files:**
- Modify: `conciliar_compras.py` (`conciliar_compras` acepta `ret=None` y la lista de advertencias incluye retenciones), `pipeline_compras.py`
- Test: `tests/test_pipeline_compras.py` (crear si no existe; usar `monkeypatch` de `OdooReader`, de `xml_cache.descargar_xmls` y de `job_manager.JOBS_DIR`/`get_job_dir` como hacen los tests de pipeline existentes)

**Interfaces:**
- Consumes: Tareas 1–4.
- Produces:
  - `conciliar_compras(rutas_txt, odoo, desde, hasta, salida, ret=None) -> {"por_tipo", "advertencias"}`: pasa `ret` a `construir_excel`; añade `por_tipo["retenciones"]` cuando `ret` existe.
  - `run_pipeline_compras(job_id)`: lee los TXT; separa compras (FAC/NC/ND, flujo actual) de retenciones (`leer_sri_retenciones_txt`); **si no hay ni compras ni retenciones** → error "Los TXT no contienen facturas, notas de crédito, notas de débito ni retenciones"; si hay retenciones: calcula su ventana con la misma lógica de `periodo_sri` aplicada al `fecha` del TXT de retenciones (mes dominante + min..max ± 5 días, **independiente** de la de compras), descarga XML con `descargar_xmls(..., on_progress=...)` que actualiza `write_status` (descarga entre 0.15 y 0.65, mensaje `"Descargando XML de retenciones (n/total, k fallidos)"`), lee Odoo con `fetch_retenciones`, concilia y pasa el resultado al Excel. Las fallas de descarga no abortan: se agregan a `advertencias` ("N XML de retenciones no se pudieron descargar; sus retenciones se listan como 'XML no disponible'") y al resumen como `resumen["retenciones_fallidos"]` (cantidad). Sin TXT de retenciones → advertencia "Falta el TXT de retenciones" y flujo de compras igual que hoy. Sin TXT de compras pero con retenciones → se omiten Odoo-compras y se emiten solo las hojas de retenciones y las de compras vacías.
  - Un solo `OdooReader.from_env()` para ambas lecturas; errores como hoy (`write_status(... "error" ...)`, `schedule_cleanup`).

- [ ] **Step 1: Write the failing tests:**
  - `test_pipeline_con_compras_y_retenciones`: TXT sintéticos de FAC y retenciones; Odoo y descarga falsos; el job termina `done`; el Excel tiene hojas de compras y `Ret_*`; `resumen["por_tipo"]["retenciones"]` presente.
  - `test_pipeline_solo_retenciones` (Review Focus 5): sin FAC/NC/ND termina `done`; advertencias incluyen los TXT de compras faltantes.
  - `test_pipeline_sin_retenciones_no_cambia`: igual que antes, con advertencia "Falta el TXT de retenciones" y **sin** llamar a `descargar_xmls` ni `fetch_retenciones`.
  - `test_pipeline_xml_fallidos_no_abortan`: `descargar_xmls` devuelve fallidos; el job termina `done`, las retenciones afectadas quedan en `Ret_Informativos` y la advertencia aparece.
  - `test_pipeline_ventana_de_retenciones_es_propia`: `fetch_retenciones` recibe fechas derivadas del TXT de retenciones, no del de facturas.
  - `test_pipeline_sin_ningun_tipo_da_error`.
  - `test_pipeline_error_de_odoo_no_filtra_credenciales`: el `error`/mensaje del status no contiene el valor de las variables `ODOO_*` del entorno de prueba.
- [ ] **Step 2: Run** `python -m pytest tests/test_pipeline_compras.py -q` → Expected: FAIL.
- [ ] **Step 3: Implement** los cambios en `pipeline_compras.py` y `conciliar_compras.py`. Mantener `schedule_cleanup` en éxito y error.
- [ ] **Step 4: Run** `python -m pytest tests -q` → Expected: todo pasa excepto los 4 fallos conocidos de Global Constraints.
- [ ] **Step 5: Commit** `git add pipeline_compras.py conciliar_compras.py tests/` y `git commit -m "feat: pipeline de Compras incluye retenciones recibidas"`.

---

### Task 6: Pantalla, documentación y verificación con datos reales

**Files:**
- Modify: `templates/index.html` (pestaña Compras: texto de ayuda que mencione retenciones y render del bloque `por_tipo.retenciones` y de `advertencias`; escapar todo lo que venga del resumen, como ya hace el render de Compras), `CLAUDE.md` (sección Pestaña Compras: añadir retenciones, nuevos módulos, hojas `Ret_*`, regla de porcentaje distinto y de "XML no disponible", y quitar "retenciones" de "Fuera de alcance"), `app.py` solo si el texto de error de la ruta lo requiere (el límite de 5 archivos no cambia).
- Test: `tests/test_app.py` (añadir)

**Interfaces:**
- Consumes: forma del resumen de la Tarea 5 (`por_tipo.retenciones` con los contadores de la Tarea 4, `advertencias`, `retenciones_fallidos`).
- Produces: UI y documentación; ningún cambio de API.

- [ ] **Step 1: Write the failing test** `test_upload_compras_acepta_txt_de_retenciones`: con sesión y `run_pipeline_compras` falso, subir un `.txt` llamado `retenciones.txt` devuelve `job_id` (la ruta no depende del tipo). Y `test_index_menciona_retenciones_en_compras`: el HTML servido en `/` contiene el texto de ayuda actualizado.
- [ ] **Step 2: Run** `python -m pytest tests/test_app.py -q` → Expected: FAIL solo en el segundo test (y los 4 fallos conocidos sin cambio).
- [ ] **Step 3: Implement** los cambios de plantilla y `CLAUDE.md`.
- [ ] **Step 4: Run** `python -m pytest tests -q` → Expected: todo pasa excepto los 4 fallos conocidos.
- [ ] **Step 5: Verificación con datos reales (solo lectura; no imprimir credenciales ni datos de clientes):** con los TXT de `documentos_electronicos/recibidas/` de septiembre 2026 (facturas, NC, ND y retenciones) ejecutar el pipeline en un script de scratchpad contra Odoo y el web service del SRI, y reportar solo conteos: por estado de retenciones (Coincide / Diferencia / Solo SRI / Solo Odoo / Factura incorrecta / XML no disponible / Informativo), XML fallidos, y que el Excel abre con `openpyxl` y las fórmulas pasan la validación con el paquete `formulas` (0 errores). Comparar los conteos de Solo SRI / Solo Odoo / Diferencia con la pestaña Retenciones antigua para el mismo mes cuando exista un `reten_odoo.xlsx` de referencia; explicar cualquier diferencia entre ambas.
- [ ] **Step 6: Commit** `git add templates/index.html CLAUDE.md tests/test_app.py app.py` y `git commit -m "feat: pantalla y documentación de retenciones en Compras"`.
