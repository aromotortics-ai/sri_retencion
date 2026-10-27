# Tablero de conciliación en vivo y reorganización — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Dos pantallas (Compras y Ventas) con tablero en vivo y Excel de observaciones, sobre un repositorio ordenado en el paquete `app/`, sin el flujo antiguo de Retenciones.

**Architecture:** Una "sesión de trabajo" es un job (`jobs/<id>/`) que guarda los TXT subidos; cada alta/baja de TXT lanza un recálculo completo en un hilo, con cancelación por contador de generación guardado en disco (gunicorn corre 2 workers: nada de estado en memoria). Cada recálculo produce una única estructura `tablero` (JSON) de la que salen la pantalla y el Excel de observaciones; el Excel completo actual se sigue generando por separado.

**Tech Stack:** Python 3.12, FastAPI, pandas, openpyxl, pytest, HTML+JS simple (sin framework).

**Spec:** `docs/superpowers/specs/2026-10-27-tablero-conciliacion-design.md` (incluye la sección "Decisiones tomadas por mí tras la revisión").

## Global Constraints

- Lectura de Odoo solo por XML-RPC de solo lectura; los errores nunca incluyen credenciales.
- No cambia la lógica de cruce, estados ni reglas de elegibilidad de `CLAUDE.md`.
- Límites de TXT: Compras 5, Ventas 400. Extensión de los TXT guardados normalizada a `.txt` minúscula.
- Estados y colores de `CLAUDE.md`; Excel completo conserva fórmulas, Arial y recálculo con LibreOffice.
- `cache/`, `jobs/`, `data/` quedan en la raíz del repo (volúmenes de Easypanel). Limpieza de jobs a las 2 h.
- Conservar `conciliar_emitidas.py` y `mover_emitidos.py` como CLI.
- Texto de interfaz, columnas y estados en español.
- `python -m pytest tests -q` debe quedar en verde al terminar cada tarea (los 4 fallos previos eran del flujo antiguo y desaparecen con la Tarea 1).
- Commits terminan con `Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>`.

## Review Focus

1. TXT ilegible o de tipo desconocido en la sesión: se marca error en ese archivo, se excluye con advertencia y el resto se concilia (Tarea 5).
2. Quitar el último TXT: el tablero queda vacío y sin error (Tarea 5).
3. Dos recálculos solapados: el superado nunca pisa el resultado del nuevo (Tarea 5).
4. Todo coincide: el Excel de observaciones es un archivo válido con una hoja que lo dice, no un libro vacío (Tarea 4).
5. Re-subir un TXT con el mismo nombre (flujo de corrección) lo reemplaza, no lo duplica ni lo renombra con prefijo (Tarea 5).

---

## Mapa de archivos (destino)

| Hoy | Destino |
|---|---|
| `app.py` | `app/web/main.py` (entrada: `app.web.main:app`) |
| `auth.py`, `templates/` | `app/web/auth.py`, `app/web/templates/` |
| `job_manager.py` | `app/job_manager.py` |
| (nuevo) | `app/rutas.py`: `RAIZ`, `JOBS_DIR`, `CACHE_XML_DIR`, `DATA_DIR` (la raíz del repo) |
| `sri_session.py`, `descargar_comprobantes_sri.py`, `parse_comprobante_xml.py`, `xml_cache.py`, `emitidos_sri.py` | `app/sri/` (mismos nombres) |
| `odoo_client.py` | `app/odoo/client.py` |
| `conciliar_compras.py`, `excel_compras.py`, `pipeline_compras.py` | `app/compras/` |
| `conciliar_emitidos.py`, `excel_emitidos.py`, `pipeline_ventas.py` | `app/ventas/` |
| `conciliar_retenciones.py`, `excel_retenciones.py` | `app/retenciones/` |
| `FONT`, `_enc`, `_fuente` de `conciliar_recibidas.py` | `app/reportes/estilos.py`; el resto de `conciliar_recibidas.py` → `scripts/` |
| `conciliar_emitidas.py`, `mover_emitidos.py`, `setup_usuario.py` | `scripts/` |
| `tests/test_*.py` | `tests/<paquete>/test_*.py`, imports `from app.…` |

---

### Task 1: Retirar el flujo antiguo de Retenciones

**Files:**
- Delete: `pipeline.py`, `build_workbook_final.py`, `conciliar_odoo.py`, `parse_recibidos_txt.py`, `tests/test_build_workbook.py`, `tests/test_conciliar.py` (si solo cubre `conciliar_odoo`), la prueba de `/upload` en `tests/test_app.py`
- Modify: `app.py` (quitar import de `run_pipeline` y la ruta `POST /upload`), `templates/index.html` (quitar pestaña, `form-retenciones` y la rama `retenciones` del JS)

**Interfaces:**
- Produces: la app solo expone `/upload-ventas` y `/upload-compras` (se reemplazan en la Tarea 6).

- [ ] **Step 1: Confirmar por imports** que nada de Compras/Ventas importa los módulos a borrar: `grep -rn "pipeline import\|build_workbook\|conciliar_odoo\|parse_recibidos_txt" --include=*.py .` solo debe listar los archivos a borrar y `app.py`. Conservar `descargar_comprobantes_sri.py`, `parse_comprobante_xml.py` y `conciliar_recibidas.py` (los usan `xml_cache.py` y los `excel_*`).
- [ ] **Step 2: Borrar y editar** lo listado; en `index.html` la pestaña inicial pasa a ser Compras (`mode = "compras"`).
- [ ] **Step 3: Run** `python -m pytest tests -q` — Expected: PASS completo.
- [ ] **Step 4: Commit** `refactor: retira el flujo antiguo de Retenciones`.

---

### Task 2: Reorganizar en el paquete `app/`

**Files:** según el "Mapa de archivos". Create: `app/__init__.py` y un `__init__.py` por subpaquete, `app/rutas.py`, `tests/<paquete>/`. Modify: `Dockerfile`, `.gitignore`, `.dockerignore` si existe.

**Interfaces:**
- Produces: `app.rutas.RAIZ: Path`, `JOBS_DIR`, `CACHE_XML_DIR`, `DATA_DIR` (todos `RAIZ / "jobs"`, `"cache"/"xml"`, `"data"`); estos reemplazan cada `Path(__file__).parent / …` actual (`job_manager.JOBS_DIR`, `CACHE_XML_DIR` de ambos pipelines, `USERS_FILE`, `TEMPLATE`).
- Produces: `app.reportes.estilos.FONT`, `_enc`, `_fuente` (extraídos de `conciliar_recibidas.py`; los `excel_*` los importan de aquí).

- [ ] **Step 1: Test de humo** `tests/test_rutas.py::test_rutas_apuntan_a_la_raiz`: `RAIZ / "Dockerfile"` existe; `JOBS_DIR == RAIZ / "jobs"`; `CACHE_XML_DIR == RAIZ / "cache" / "xml"`. Run → FAIL (módulo no existe).
- [ ] **Step 2: Mover con `git mv`** (conserva historial), crear `app/rutas.py`, reescribir imports a `from app.<paquete>.<módulo> import …` y los `monkeypatch`/rutas de cadena de los tests (`"pipeline_ventas.OdooReader"` → `"app.ventas.pipeline_ventas.OdooReader"`). Las plantillas se leen desde `app/web/templates/`.
- [ ] **Step 3: Dockerfile**: `COPY app/ app/`, `COPY scripts/ scripts/`; `CMD` con `app.web.main:app`. `.gitignore`: añadir `*.xlsx` de prueba de la raíz; mover/eliminar `test_final.xlsx`, `test_output.xlsx`, `verify_output.xlsx`, `claves_agosto2026.txt`, `Comprobantes_Retencion_Agosto2026_Final.xlsx`, `reten_odoo.xlsx`, `xml_samples/` según sean basura (si hay duda sobre alguno, moverlo a `docs/ejemplos/` en vez de borrarlo).
- [ ] **Step 4: Run** `python -m pytest tests -q` y `python -c "import app.web.main"` — Expected: PASS; mismo número de tests que antes de la tarea.
- [ ] **Step 5: Commit** `refactor: reorganiza el código en el paquete app/`.

---

### Task 3: Clasificación de Compras en Python (estados por documento)

Compras hoy calcula estados con fórmulas del Excel; el tablero necesita los estados en Python, con las mismas reglas.

**Files:**
- Create: `app/compras/clasificar.py`
- Test: `tests/compras/test_clasificar.py`

**Interfaces:**
- Consumes: `conciliar(sri, odoo, desde, hasta) -> DataFrame[idx_sri, id_odoo, tipo, metodo, cat]` (`app/compras/conciliar_compras.py`), `_mismo_total`, `CAT_DUP`, `CAT_INFO`, `CAT_SOLO_ODOO`.
- Produces: `clasificar_compras(sri, odoo, pares) -> DataFrame` con las columnas de `app.ventas.conciliar_emitidos.COLUMNAS` (`estado, tipo, idx_sri, id_odoo, metodo, detalle`) y los estados `Coincide`, `Diferencia de valor`, `Solo en SRI`, `Solo en Odoo`, `Duplicado en Odoo`, `Informativo`.

- [ ] **Step 1: Test** `test_clasificar_compras_replica_el_excel`: con SRI/Odoo sintéticos (una pareja igual, una con total distinto en $0.50, una con $0.02 de diferencia, un solo SRI, un solo Odoo elegible, uno no elegible, un duplicado) los estados son, en ese orden, `Coincide`, `Diferencia de valor`, `Coincide`, `Solo en SRI`, `Solo en Odoo`, `Informativo`, `Duplicado en Odoo`. Y `test_conteos_igualan_construir_excel`: los conteos por estado y tipo de `clasificar_compras` son iguales a `construir_excel(...)["FAC"]` (claves `coinciden`, `diferencias`, `solo_sri`, `solo_odoo`, `informativos`).
- [ ] **Step 2: Run** → FAIL.
- [ ] **Step 3: Implement** `clasificar_compras` en `app/compras/clasificar.py`: misma regla que la fórmula de `excel_compras.py:125` (dup → dup; ambos lados → |dif| > `TOL` ⇒ Diferencia, si no Coincide; solo SRI; si no la `cat` de Odoo). `detalle` para diferencia: `"total SRI X ≠ Odoo Y"` como en Ventas.
- [ ] **Step 4: Run** → PASS. **Step 5: Commit** `feat: clasificación de Compras en Python`.

---

### Task 4: Tablero y Excel de observaciones

**Files:**
- Create: `app/reportes/tablero.py`, `app/reportes/observaciones.py`
- Test: `tests/reportes/test_tablero.py`, `tests/reportes/test_observaciones.py`

**Interfaces:**
- Consumes: `clasificar_compras` (Tarea 3), `clasificar_documentos`, `resumir_por_documento(ret) -> DataFrame[COLUMNAS_DOC]`, `TIPOS_DOC`.
- Produces:
  - `QUE_CORREGIR: dict[str, str]` — texto fijo por estado: `Solo en SRI` → "Registrar el comprobante en Odoo"; `Solo en Odoo` → "Verificar el número/clave del documento en Odoo o anularlo si no existe en el SRI"; `Diferencia de valor` → "Corregir el valor en Odoo (ver detalle)"; `Diferencia de datos` → "Corregir clave de acceso, fecha o marca electrónica en Odoo (ver detalle)"; `Factura incorrecta en Odoo` → "Corregir el número de la factura sustento en la retención"; `Odoo anulado, SRI autorizado` → "Reactivar el documento en Odoo o anular el comprobante en el SRI"; `Duplicado en Odoo` → "Anular el registro duplicado en Odoo"; `XML no disponible` → "Reintentar: no se pudo leer el XML del SRI".
  - `ACCIONABLES: frozenset[str]` = las claves de `QUE_CORREGIR`. `Coincide` e `Informativo*` no son accionables.
  - `COLUMNAS_OBS = ["tipo","estado","documento","fecha","contraparte","ruc","valor_sri","valor_odoo","diferencia","detalle","que_corregir","accionable"]`.
  - `construir_tablero(modo, docs, retdocs, periodo, advertencias) -> dict` con `modo: "compras"|"ventas"`, `periodo: [iso, iso] | None`, `tarjetas: list[{tipo, etiqueta, total, coinciden, por_estado: dict[str,int]}]` (retenciones cuentan **documentos**), `excepciones: list[dict]` (todas las filas que no son `Coincide`, con las llaves de `COLUMNAS_OBS`), `advertencias: list[str]`. Valores JSON-serializables (fechas ISO, NaN → `None`).
  - `escribir_excel_observaciones(tablero: dict, ruta: str) -> int` — escribe solo `accionable == True`, una hoja por estado (nombre ≤ 31 caracteres, sin caracteres inválidos), Arial, filtros y encabezado fijo; devuelve el número de observaciones.

- [ ] **Step 1: Tests** `test_tarjetas_cuentan_documentos_de_retenciones` (retención de 2 líneas = 1 documento); `test_excepciones_excluyen_coincide_e_incluyen_informativo` (Informativo está en `excepciones` con `accionable=False`); `test_excel_observaciones_solo_accionables` (el libro no contiene ninguna fila `Coincide` ni `Informativo`; `devuelve == sum(accionable)`); `test_excel_sin_observaciones_es_valido` (todo `Coincide` → abre con openpyxl, tiene la hoja `Sin_observaciones` con el texto "No hay observaciones"); `test_cifras_tablero_igualan_excel_completo` (conteos de `tarjetas` iguales a `res["documentos"]` de `construir_excel_emitidos` y a `por_tipo` de `construir_excel`).
- [ ] **Step 2: Run** → FAIL.
- [ ] **Step 3: Implement** los tres símbolos; `que_corregir` sale de `QUE_CORREGIR[estado]` y, si hay `detalle`, no se repite (el detalle va en su columna).
- [ ] **Step 4: Run** → PASS. **Step 5: Commit** `feat: tablero y Excel de observaciones`.

---

### Task 5: Sesión de trabajo con recálculo cancelable

**Files:**
- Create: `app/sesion.py`
- Modify: `app/compras/pipeline_compras.py`, `app/ventas/pipeline_ventas.py` (devuelven además el `tablero`; reciben un `cancelado()` opcional), `app/job_manager.py`
- Test: `tests/test_sesion.py`, ajustar `tests/compras/test_pipeline_compras.py` y `tests/ventas/test_pipeline_ventas.py`

**Interfaces:**
- Consumes: `construir_tablero`, `escribir_excel_observaciones`, `create_job`, `get_job_dir`, `write_status`.
- Produces (`app/sesion.py`), todo basado en archivos de `jobs/<id>/` (`txt/` y `sesion.json`):
  - `crear_sesion(modo: str) -> str` — `modo ∈ {"compras","ventas"}`; devuelve el id.
  - `agregar_txt(sesion_id: str, nombre: str, contenido: bytes) -> dict` — guarda en `txt/` con extensión `.txt`; **mismo nombre reemplaza**; valida lectura (TXT del portal con sus columnas) y devuelve `{"nombre", "estado": "ok"|"error", "tipos": {"FAC": n,…}, "filas": int, "error": str|None}`. Lanza `ValueError` si se excede el límite del modo (5/400).
  - `quitar_txt(sesion_id: str, nombre: str) -> None`.
  - `listar_txt(sesion_id: str) -> list[dict]` (mismo dict que `agregar_txt`).
  - `solicitar_recalculo(sesion_id: str) -> int` — incrementa `generacion` en `sesion.json` (escritura atómica) y lanza el hilo; devuelve la generación.
  - `recalcular(sesion_id: str, generacion: int) -> None` — corre el pipeline del modo con solo los TXT en estado `ok`; los `error` se excluyen y suman una advertencia; en puntos de control (tras leer TXT, tras Odoo, tras XML, antes de escribir) compara `generacion` con la de disco y, si fue superada, termina sin escribir. Al terminar escribe `tablero.json`, `resultado.xlsx`, `observaciones.xlsx` y `write_status(..., "done", resumen=…)`. Sin ningún TXT `ok` escribe un `tablero.json` vacío (sin tarjetas ni excepciones) y estado `done`, sin llamar a Odoo.
- Debounce: lo hace el cliente (Tarea 7); el servidor solo cancela lo superado.

- [ ] **Step 1: Tests** `test_mismo_nombre_reemplaza` (dos `agregar_txt` con "facturas.txt" → `listar_txt` tiene 1 y sin prefijo numérico); `test_txt_ilegible_se_excluye_y_el_resto_concilia` (con Odoo falso: un TXT roto + uno válido → `tablero.json` tiene las tarjetas del válido y una advertencia que nombra al roto); `test_quitar_ultimo_txt_deja_tablero_vacio` (estado `done`, `tarjetas == []`, Odoo no se invoca); `test_recalculo_superado_no_pisa_el_nuevo` (se lanza `recalcular(id, 1)` con un Odoo falso que bloquea en un `threading.Event`; se pide la generación 2 y se completa; liberar el evento; `tablero.json` es el de la generación 2); `test_limite_de_txt_por_modo` (6.º TXT en compras → `ValueError`).
- [ ] **Step 2: Run** → FAIL.
- [ ] **Step 3: Implement** `app/sesion.py` y refactorizar los pipelines: separar el cuerpo de `run_pipeline_*` en una función que recibe rutas de TXT y devuelve `(resumen, tablero)`; `run_pipeline_*` actuales dejan de existir (los reemplaza `recalcular`). El hilo usa `threading.Thread(daemon=True)`; `schedule_cleanup` se reprograma en cada alta (2 h desde la última actividad).
- [ ] **Step 4: Run** `python -m pytest tests -q` → PASS. **Step 5: Commit** `feat: sesión de trabajo con recálculo cancelable`.

---

### Task 6: Endpoints de la sesión

**Files:**
- Modify: `app/web/main.py` (quitar `/upload-ventas`, `/upload-compras`, `MAX_TXT_VENTAS`; agregar rutas)
- Test: `tests/web/test_app.py`

**Interfaces:**
- Consumes: `app.sesion.*` (Tarea 5). Todas con `Depends(require_session)`.
- Produces:
  - `POST /sesion` body JSON `{"modo": "compras"|"ventas"}` → `{"sesion_id"}`.
  - `POST /sesion/{id}/txt` (multipart, campo `archivos`, 1..N) → `{"archivos": [dict de agregar_txt], "generacion": int}`; 400 si el nombre no es `.txt` o excede el límite; **no** recalcula (lo pide el cliente).
  - `DELETE /sesion/{id}/txt/{nombre}` → `{"archivos": [...]}`.
  - `POST /sesion/{id}/conciliar` → `{"generacion": int}`.
  - `GET /status/{id}` (existente) y `GET /tablero/{id}` → contenido de `tablero.json` (404 si aún no existe; incluye `generacion`).
  - `GET /download/{id}?tipo=observaciones|completo` (por defecto `completo`) con nombres `Observaciones_SRI.xlsx` y `Conciliacion_SRI.xlsx`.
  - 404 si la sesión no existe; el nombre del TXT se sanea con `Path(nombre).name`.

- [ ] **Step 1: Tests** (con `TestClient` y sesión autenticada como los de `test_app.py`): `test_flujo_crear_subir_conciliar_tablero` (Odoo falso; tras `conciliar` y esperar `done`, `GET /tablero/{id}` trae `tarjetas`); `test_subir_sin_sesion_da_401`; `test_nombre_con_ruta_se_sanea` (`../../x.txt` queda como `x.txt`); `test_descargar_observaciones_y_completo` (ambos 200 con content-type xlsx); `test_sesion_inexistente_404`.
- [ ] **Step 2: Run** → FAIL. **Step 3: Implement** las rutas. **Step 4: Run** `python -m pytest tests -q` → PASS.
- [ ] **Step 5: Commit** `feat: API de sesión para el tablero en vivo`.

---

### Task 7: Pantalla del tablero (opción A)

**Files:**
- Modify: `app/web/templates/index.html` (reescritura; se separan `app/web/static/tablero.js` y `tablero.css` si el archivo supera ~300 líneas; servir con `StaticFiles` en `/static`)
- Modify: `app/web/main.py` (montar `/static`), `Dockerfile` si hace falta
- Test: `tests/web/test_app.py::test_index_trae_las_dos_pestanas`

**Interfaces:**
- Consumes: endpoints de la Tarea 6; forma de `tablero` de la Tarea 4.

Comportamiento (según el boceto A aprobado):
- Pestañas **Compras** y **Ventas**; cada una crea su sesión al primer TXT (`POST /sesion`) y la conserva al cambiar de pestaña; "Nueva conciliación" la descarta.
- Zona para soltar/seleccionar TXT arriba; lista de archivos con estado (en cola, listo con `filas`, error con el motivo) y botón para quitar cada uno.
- Tras subir/quitar, **debounce de 700 ms** y un solo `POST /conciliar`; se hace sondeo de `/status` cada 2 s; mientras recalcula se muestra "Actualizando…" y se conserva el tablero anterior; se ignoran respuestas de generaciones anteriores a la última pedida.
- Advertencias del tablero en un recuadro sobre las tarjetas.
- Tarjetas por tipo: "N de M coinciden" + barra; clic en una tarjeta filtra la tabla por ese tipo.
- Tabla de excepciones: filtros por estado (con contador) y por tipo, búsqueda por documento/contraparte, orden por columna; las filas `Informativo*` se muestran atenuadas y con el filtro "Informativos" (apagado por defecto).
- Botones **Descargar observaciones (Excel)** (`?tipo=observaciones`, deshabilitado si no hay accionables) y **Descargar Excel completo**.
- Todo texto dinámico pasa por `esc()` (como hoy); 401 → `/login`.

- [ ] **Step 1: Test** `test_index_trae_las_dos_pestanas`: `GET /` autenticado contiene `id="tab-compras"` y `id="tab-ventas"` y no contiene `form-retenciones`. Run → FAIL.
- [ ] **Step 2: Implement** la pantalla.
- [ ] **Step 3: Run** `python -m pytest tests -q` → PASS.
- [ ] **Step 4: Verificación manual en navegador** (skill `run`/Playwright): arrancar la app con un script temporal fuera del repo que sustituya `OdooReader.from_env` por un lector falso (sin agregar ningún gancho de prueba al código de producción); subir dos TXT sintéticos de Ventas uno tras otro y comprobar que las tarjetas cambian sin recargar, que quitar un TXT recalcula, que re-subir el mismo nombre no lo duplica y que ambos Excel se descargan y abren. Capturar pantalla.
- [ ] **Step 5: Commit** `feat: pantalla de tablero en vivo`.

---

### Task 8: Documentación y cierre

**Files:** Modify: `CLAUDE.md`, `docs/superpowers/specs/…` (sin cambios), `README` si existe.

- [ ] **Step 1: Actualizar `CLAUDE.md`**: nueva estructura de carpetas, tabla de archivos con rutas nuevas, endpoints de sesión, tablero y Excel de observaciones; eliminar las secciones de la pestaña Retenciones antigua, `pipeline.py`, `build_workbook_final.py`, `conciliar_odoo.py`; comandos (`gunicorn app.web.main:app`, `python -m pytest tests -q`).
- [ ] **Step 2: Run** `python -m pytest tests -q`; `docker build .` si Docker está disponible — Expected: PASS / imagen construida y `python -c "import app.web.main"` dentro de ella sin error.
- [ ] **Step 3: Commit** `docs: CLAUDE.md con la nueva estructura y el tablero`.

---

## Autorevisión

- **Cobertura del spec:** pantalla A, dos pestañas, subida incremental, recálculo completo, advertencias en pantalla (T5/T7); Excel de observaciones solo accionables con "Qué corregir" y Excel completo aparte (T4); cifras idénticas (T4 test); retiro del flujo antiguo (T1); carpetas (T2); Docker (T2/T8); límites y saneo de nombres (T5/T6); documentación (T8). Sin tareas huérfanas.
- **Consistencia de nombres:** `construir_tablero`, `escribir_excel_observaciones`, `clasificar_compras`, `QUE_CORREGIR`, `ACCIONABLES`, `agregar_txt`, `recalcular`, `solicitar_recalculo` se usan igual en todas las tareas.
- **Proporción:** el plan especifica contratos y pruebas; los cuerpos los escribe quien construye.
