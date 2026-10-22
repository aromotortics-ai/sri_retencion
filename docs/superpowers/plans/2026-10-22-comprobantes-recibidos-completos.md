# Conciliación de comprobantes electrónicos recibidos — Plan de implementación

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Desde la app web, subir los TXT del SRI (facturas, notas de crédito, notas de débito) y recibir un Excel de conciliación contra Odoo 14, leído directo por API en solo lectura.

**Architecture:** Cuatro módulos pequeños y un runner. `odoo_client` lee y normaliza asientos de compra de Odoo por XML-RPC. `conciliar_compras` lee los TXT, cruza por niveles y devuelve pares. `excel_compras` arma el libro con fórmulas. `pipeline_compras` y la ruta `/upload-compras` copian el patrón ya existente de Ventas (hilo background, `status.json`, `/download`).

**Tech Stack:** Python 3, pandas, openpyxl, `xmlrpc.client` (stdlib), FastAPI, pytest. Sin dependencias nuevas.

**Spec:** `docs/superpowers/specs/2026-10-22-comprobantes-recibidos-completos-design.md`

## Global Constraints

- Odoo 14.0 en `https://aromotor.com`, base `aromotor`; **solo lectura** (únicamente `authenticate`, `fields_get`, `search_read`, `read`, `search_count`).
- Credenciales solo por variables de entorno `ODOO_URL`, `ODOO_DB`, `ODOO_USER`, `ODOO_PASSWORD` (en local, desde `.env`, que git ignora). Nunca en el repositorio, ni en logs, ni en mensajes de error.
- TXT del SRI: ISO-8859-1, separador tabulador. `IMPORTE_TOTAL` viene vacío en notas de crédito y débito: total = base + IVA.
- Tolerancia de valor: `0.02`. Notas de crédito se comparan por valor absoluto.
- Clave de acceso: 49 dígitos; las de Odoo con 45–48 dígitos se rellenan con ceros a la izquierda (Excel borraba el cero inicial).
- Tipos de comparación: Factura, Nota de crédito, Nota de débito. Las liquidaciones y las retenciones quedan fuera de esta fase (retenciones ya existen en `/upload`).
- Solo se compara contra el SRI lo que en Odoo tiene `x_receipt_type_id.code` en `01`, `04` o `05`, clave de 45–49 dígitos y `x_is_edocument = False`. Lo demás es "informativo", nunca error.
- Excel: Arial; fórmulas (`INDEX/MATCH`, `COUNTIFS`, `SUMIFS`, `IFERROR`), nunca valores calculados pegados; recalcular con LibreOffice antes de entregar; colores de hoja: azul resumen, naranja Solo SRI, rojo Solo Odoo, amarillo diferencias, violeta duplicados.
- No romper las pestañas Retenciones y Ventas ni sus tests. Spanish para etiquetas de usuario; funciones y módulos nuevos en snake_case.

## Review Focus

- Clave de Odoo con 48 dígitos por cero inicial perdido → debe cruzar por clave, no caer a "Solo en Odoo" (Task 1 y 3).
- Nota de crédito/débito con `IMPORTE_TOTAL` vacío → total = base + IVA, no 0 (Task 2).
- Documento de Odoo fechado fuera del período pero emparejado con el SRI → "Coincide"; sin pareja y fuera de período → no aparece como "Solo en Odoo" (Task 3).
- Misma clave en dos TXT de meses solapados, o dos asientos Odoo con la misma clave → una sola fila SRI; el segundo Odoo marcado "Duplicado en Odoo" (Task 2 y 3).
- Odoo inalcanzable o credenciales ausentes → job en `error` con mensaje claro y sin credenciales en el texto (Task 1 y 5).

---

### Task 1: Cliente de Odoo de solo lectura

**Files:**
- Create: `odoo_client.py`
- Test: `tests/test_odoo_client.py`

**Interfaces:**
- Produces:
  - `class OdooError(Exception)`
  - `def cargar_env(ruta_env: str | None = None) -> dict[str, str]` — toma `os.environ` y, para claves faltantes, el `.env` (por defecto junto a este archivo); formato `CLAVE=valor`, ignora comentarios y comillas.
  - `class OdooReader` con `__init__(self, url, db, user, password, proxy_factory=xmlrpc.client.ServerProxy)`, `@classmethod from_env(cls) -> OdooReader` (lanza `OdooError("Falta configurar ODOO_URL/ODOO_DB/ODOO_USER/ODOO_PASSWORD")` listando solo los **nombres** faltantes) y `fetch_compras(self, desde: date, hasta: date, margen_dias: int = 5) -> pd.DataFrame`.
  - Columnas de `fetch_compras`: `id:int`, `num:str`, `nfac:str|NA`, `socio:str`, `ruc:str|NA`, `fecha:Timestamp`, `tipo:str` en `{"FAC","NC","ND","OTRO"}`, `cod_comp:str`, `clave:str|NA`, `base:float`, `iva:float`, `total:float` (valor absoluto), `afectada:str|NA`, `elegible:bool`.

- [ ] **Step 1: Verificar el campo del número de documento (solo lectura)**

Con el `.env` real, leer 3 asientos `in_invoice` de septiembre 2026 (`name`, `x_number`, `ref`) y confirmar que `x_number` trae el número del proveedor con formato `NNN-NNN-NNNNNNNNN`, igual que "Número de factura" del Excel de Odoo. Si el formato estuviera en otro campo, usar ese y actualizar los nombres de este plan y del spec.

- [ ] **Step 2: Escribir los tests con un `proxy_factory` falso**

Casos (con filas de `search_read` inventadas, sin red): `test_clave_48_digitos_se_rellena_a_49`; `test_tipo_por_codigo_comprobante` (`01→FAC`, `04→NC`, `05→ND`, `02→OTRO`); `test_elegible_exige_clave_y_no_ser_documento_propio` (`x_is_edocument=True` ⇒ `elegible=False`); `test_total_es_valor_absoluto`; `test_ruc_viene_del_partner`; `test_afectada_de_reversed_entry_para_nc`; `test_consulta_usa_solo_posted_y_margen_de_fechas` (el dominio contiene `state=posted`, `invoice_date >= desde-margen`, `<= hasta+margen`, `move_type in [in_invoice, in_refund]`); `test_from_env_lista_nombres_faltantes_sin_valores`; `test_fallo_de_conexion_lanza_odooerror_sin_credenciales` (el texto del error no contiene la contraseña).

- [ ] **Step 3: Ejecutar los tests y verificar que fallan**

Run: `python -m pytest tests/test_odoo_client.py -q` — Expected: FAIL (módulo inexistente).

- [ ] **Step 4: Implementar `odoo_client.py`**

Campos a leer de `account.move`: `name, x_number, partner_id, invoice_date, move_type, state, x_receipt_type_id, x_authorization_number, x_is_edocument, amount_untaxed, amount_tax, amount_total, reversed_entry_id`. El `code` del tipo de comprobante se obtiene con una sola lectura de `account.receipt.type` y el `vat` con una sola lectura de `res.partner` por lote de ids. Para `NC`, `afectada` = `x_number` del asiento de `reversed_entry_id` (lectura en lote). Cualquier excepción de red/`xmlrpc.client.Fault` se envuelve en `OdooError` con mensaje genérico sin URL con usuario ni contraseña. Paginar `search_read` en bloques de 500.

- [ ] **Step 5: Ejecutar los tests y verificar que pasan**

Run: `python -m pytest tests/test_odoo_client.py -q` — Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add odoo_client.py tests/test_odoo_client.py
git commit -m "feat: cliente de solo lectura de Odoo 14 para compras"
```

---

### Task 2: Lector de los TXT del SRI por tipo

**Files:**
- Create: `conciliar_compras.py` (solo la parte de lectura en esta tarea)
- Test: `tests/test_conciliar_compras.py`

**Interfaces:**
- Produces:
  - `TIPOS = ("FAC", "NC", "ND")`
  - `def leer_sri_txts(rutas: list[str]) -> pd.DataFrame` — columnas `clave:str`, `serie:str`, `ruc:str`, `proveedor:str`, `fecha:Timestamp`, `tipo` (`FAC|NC|ND`), `base`, `iva`, `total` (floats; si `IMPORTE_TOTAL` vacío ⇒ `base + iva`), `afectada:str|NA`. Detecta el tipo por `TIPO_COMPROBANTE` (empieza con "Factura", "Nota de Cr", "Nota de D"); filas de otros tipos (retenciones, liquidaciones) se ignoran. Elimina duplicados por `clave`. Índice 0..n-1 ordenado por fecha y serie.
  - `def tipos_faltantes(sri: pd.DataFrame) -> list[str]` — tipos de `TIPOS` sin filas.

- [ ] **Step 1: Escribir los tests con TXT generados en `tmp_path`**

`test_total_vacio_en_nc_es_base_mas_iva` (300.26 + 45.04 ⇒ 345.30); `test_tipo_detectado_por_texto_con_acentos_iso8859`; `test_ignora_retenciones_y_liquidaciones`; `test_misma_clave_en_dos_archivos_se_deduplica`; `test_afectada_viene_de_numero_documento_modificado`; `test_tipos_faltantes_reporta_nd_si_no_hay_archivo`.

- [ ] **Step 2: Ejecutar y verificar que fallan** — `python -m pytest tests/test_conciliar_compras.py -q` → FAIL.

- [ ] **Step 3: Implementar `leer_sri_txts` y `tipos_faltantes`**

Leer con `pd.read_csv(sep="\t", encoding="iso-8859-1", dtype=str)`; error claro (`ValueError`) si faltan columnas esperadas del portal.

- [ ] **Step 4: Ejecutar y verificar que pasan** — mismo comando → PASS.

- [ ] **Step 5: Commit**

```bash
git add conciliar_compras.py tests/test_conciliar_compras.py
git commit -m "feat: lector de TXT del SRI para facturas, notas de crédito y débito"
```

---

### Task 3: Motor de cruce SRI vs Odoo

**Files:**
- Modify: `conciliar_compras.py`
- Test: `tests/test_conciliar_compras.py`

**Interfaces:**
- Consumes: `leer_sri_txts` (Task 2) y el DataFrame de `OdooReader.fetch_compras` (Task 1).
- Produces:
  - Constantes de categoría: `CAT_INFO = "No aplica (informativo)"`, `CAT_DUP = "Duplicado en Odoo"`, `CAT_SOLO_ODOO = "Solo en Odoo"`.
  - `def conciliar(sri: pd.DataFrame, odoo: pd.DataFrame, desde: date, hasta: date) -> pd.DataFrame` — columnas `idx_sri` (Int, NaN si no hay), `id_odoo` (Int, NaN si no hay), `tipo`, `metodo`, `cat`. Una fila por pareja, y una por cada lado sin pareja.

Reglas de cruce, en este orden por documento de Odoo elegible y del mismo `tipo`: (1) clave de acceso; (2) `nfac + total` con el mismo RUC cuando ambos lados lo tienen; (3) `nfac` único en el SRI con total distinto → `metodo="N° documento (importe distinto)"`; (4) solo NC: `(ruc, afectada, total)`. Documentos de Odoo no elegibles ⇒ `cat=CAT_INFO`. Segundo asiento Odoo con la misma pareja SRI ⇒ `cat=CAT_DUP`. Un documento Odoo sin pareja con `fecha` fuera de `[desde, hasta]` **no genera fila** (era solo margen de fechas); uno emparejado fuera de período sí.

- [ ] **Step 1: Escribir los tests con DataFrames mínimos construidos a mano**

`test_cruza_por_clave_con_48_digitos_rellenada`; `test_cruza_por_numero_y_total_si_la_clave_de_odoo_esta_mal`; `test_numero_con_total_distinto_se_empareja_pero_marca_metodo`; `test_nc_cruza_por_factura_afectada`; `test_no_elegible_es_informativo_no_error`; `test_odoo_fuera_de_periodo_sin_pareja_se_omite`; `test_odoo_fuera_de_periodo_con_pareja_si_cruza`; `test_duplicado_en_odoo`; `test_no_cruza_entre_tipos` (misma serie en FAC y NC no se mezcla); `test_solo_sri_genera_fila_sin_id_odoo`.

- [ ] **Step 2: Ejecutar y verificar que fallan.**

- [ ] **Step 3: Implementar `conciliar`**

Tomar como base la lógica de `conciliar` de `conciliar_recibidas.py` (diccionarios por clave, `k2`, serie única) agregando el filtro por `tipo`, el RUC y el nivel (4).

- [ ] **Step 4: Ejecutar y verificar que pasan.**

- [ ] **Step 5: Commit**

```bash
git add conciliar_compras.py tests/test_conciliar_compras.py
git commit -m "feat: cruce SRI vs Odoo por tipo de comprobante"
```

---

### Task 4: Libro Excel de resultados

**Files:**
- Create: `excel_compras.py`
- Test: `tests/test_excel_compras.py`
- Commit (sin cambios): `conciliar_recibidas.py`, del que se importan `_enc`, `_hoja`, `_fuente`, `COLORES`, `FONT`; hoy está sin versionar y esta tarea lo deja en git.

**Interfaces:**
- Consumes: `sri`, `odoo`, `pares` (Tasks 1–3).
- Produces: `def construir_excel(sri, odoo, pares, desde: date, hasta: date, salida: str) -> dict` — escribe el `.xlsx` y devuelve el resumen `{tipo: {"coinciden": int, "solo_sri": int, "solo_odoo": int, "diferencias": int, "informativos": int}}`.

Hojas: `Resumen` (por tipo y mes con `COUNTIFS` sobre `Conciliacion`), `Conciliacion` (una fila por par; columna `Resultado` por fórmula: Coincide / Diferencia de valor / Solo en SRI / `cat`), `Solo_en_SRI`, `Solo_en_Odoo`, `Diferencia_Valor`, `Duplicados_Odoo`, `Informativos`, `SRI_Listado`, `Odoo_Lineas`. Las hojas de datos usan las columnas `Tipo`, `Clave`, `Serie/Número`, `RUC`, `Proveedor`, `Fecha`, `Base`, `IVA`, `Total`. Marca "Fecha fuera del período" en filas de Odoo fuera de rango emparejadas.

- [ ] **Step 1: Escribir los tests**

Con datos mínimos, abrir el libro con openpyxl: `test_hojas_esperadas_y_colores`; `test_resultado_es_formula_no_valor` (la celda `Resultado` empieza con `=`); `test_resumen_cuenta_por_tipo` (dos FAC coinciden, una NC solo SRI ⇒ valores correctos en el dict devuelto); `test_texto_que_empieza_con_igual_no_se_guarda_como_formula` (regresión del commit `96f5249`: proveedor `=ABC`); `test_fuente_arial`.

- [ ] **Step 2: Ejecutar y verificar que fallan.**

- [ ] **Step 3: Implementar `construir_excel`**

Seguir `conciliar_recibidas.construir_excel` (índices `_fila_SRI`/`_fila_Odoo` con `MATCH` y `INDEX`) añadiendo `Tipo` y `RUC`. Después de escribir, el runner (Task 5) invoca la recalculación con LibreOffice; en este módulo solo se fija `CalcProperties(fullCalcOnLoad=True)`.

- [ ] **Step 4: Ejecutar y verificar que pasan.**

- [ ] **Step 5: Commit**

```bash
git add conciliar_recibidas.py excel_compras.py tests/test_excel_compras.py
git commit -m "feat: libro Excel de conciliación de compras por tipo"
```

---

### Task 5: Orquestación, pipeline, ruta web y pestaña "Compras"

**Files:**
- Modify: `conciliar_compras.py`, `app.py`, `templates/index.html`
- Create: `pipeline_compras.py`
- Test: `tests/test_pipeline_compras.py`, `tests/test_app.py` (casos nuevos)

**Interfaces:**
- Consumes: todo lo anterior; `job_manager.create_job/get_job_dir/write_status/schedule_cleanup` como en `pipeline_ventas.py`.
- Produces:
  - `def conciliar_compras(rutas_txt: list[str], odoo: pd.DataFrame, desde: date, hasta: date, salida: str) -> dict` en `conciliar_compras.py`: lee SRI, calcula `desde/hasta` con el mes de las fechas del SRI si no vienen, concilia, escribe el Excel y devuelve `{"por_tipo": {...}, "advertencias": ["Falta el TXT de notas de débito", ...]}`.
  - `def run_pipeline_compras(job_id: str) -> None` en `pipeline_compras.py`: lee `jobs/<id>/sri/*.txt`, instancia `OdooReader.from_env()`, escribe progreso (0.1 leyendo SRI, 0.4 leyendo Odoo, 0.8 conciliando), deja `resultado.xlsx`, estado `done` con `resumen`, `schedule_cleanup(job_id, delay_seconds=7200)`. Ante `OdooError` o cualquier excepción: `write_status(..., "error", mensaje)` con el mensaje de la excepción (ya sin credenciales).
  - `POST /upload-compras` (`sri: list[UploadFile] = File(...)`, hasta 5): mismo manejo de nombres que `/upload-ventas` (`Path(...).name`, evitar choques); responde `{"job_id": ...}`; requiere sesión.
  - Pestaña "Compras" junto a Retenciones y Ventas, con subida múltiple de TXT; usa `/status` y `/download` existentes.

- [ ] **Step 1: Escribir los tests**

`test_pipeline_compras_ok` (parchea `OdooReader.from_env` y `conciliar_compras`, como en `test_pipeline_ventas.py`); `test_pipeline_compras_sin_credenciales_error_claro`; `test_upload_compras_requiere_sesion` y `test_upload_compras_devuelve_job_id` (patrón de `test_app.py`); `test_upload_compras_nombre_con_ruta_no_escapa_del_job`; y en `test_conciliar_compras.py`: `test_advertencia_por_tipo_faltante` y `test_resumen_por_tipo`.

- [ ] **Step 2: Ejecutar y verificar que fallan.**

- [ ] **Step 3: Implementar** `conciliar_compras()`, `pipeline_compras.py`, la ruta y la pestaña.

La recalculación con LibreOffice (`soffice --headless --convert-to xlsx`) ocurre en `conciliar_compras()` tras escribir el libro; si LibreOffice no existe (p. ej. en el contenedor), se deja `fullCalcOnLoad` y se añade una advertencia al resumen en vez de fallar.

- [ ] **Step 4: Ejecutar toda la suite**

Run: `python -m pytest tests -q` — Expected: PASS, incluidas las pruebas de Retenciones y Ventas.

- [ ] **Step 5: Commit**

```bash
git add conciliar_compras.py pipeline_compras.py app.py templates/index.html tests
git commit -m "feat: pestaña Compras que concilia TXT del SRI contra Odoo por API"
```

---

### Task 6: Verificación con datos reales, despliegue y documentación

**Files:**
- Modify: `CLAUDE.md`, `Dockerfile`/`deploy/` solo si hace falta, `.gitignore` solo si `.env` no estuviera ignorado.

- [ ] **Step 1: Prueba de humo con septiembre 2026**

Ejecutar `conciliar_compras` con los TXT de `documentos_electronicos/recibidas/{facturas,NC,ND}_recibidas/` de septiembre y el lector de Odoo real. Esperado: 448 facturas, 30 notas de crédito y 1 nota de débito del SRI leídas; Odoo con 409 `in_invoice` y 25 `in_refund` publicados en el período. Revisar a mano 10 filas de "Solo en SRI" y 10 de "Solo en Odoo" y que ninguna sea un documento no electrónico.

- [ ] **Step 2: Abrir el Excel en LibreOffice sin errores**

Verificar que no hay `#NAME?`, `#REF!` ni celdas en blanco por fórmulas sin caché.

- [ ] **Step 3: Configurar variables en Easypanel**

Agregar `ODOO_URL`, `ODOO_DB`, `ODOO_USER`, `ODOO_PASSWORD` como variables de entorno del servicio (nunca en la imagen). Probar `/upload-compras` en el despliegue con los TXT de septiembre.

- [ ] **Step 4: Documentar en `CLAUDE.md`**

Añadir a "Aplicación web" los archivos nuevos, las cuatro variables de entorno, la regla de elegibilidad y que el caché de XML no se usa en Compras.

- [ ] **Step 5: Commit**

```bash
git add CLAUDE.md
git commit -m "docs: documentar conciliación de compras por API de Odoo"
```

---

## Cobertura del spec

Alcance por tipo → Tasks 2 y 3 · API de Odoo solo lectura y credenciales → Task 1 · reglas de elegibilidad e informativos → Tasks 1 y 3 · cruce en niveles con RUC → Task 3 · notas de crédito por valor absoluto y afectada → Tasks 1 y 3 · TXT faltante como advertencia → Tasks 2 y 5 · subida múltiple → Task 5 · entrega en Excel con fórmulas → Task 4 · pestañas existentes intactas → Task 5 · vista diaria, login al SRI, liquidaciones y respaldo con Excel de Odoo: **fuera de alcance** (el respaldo con Excel, que el spec listaba como mi llamado, se omite por YAGNI).
