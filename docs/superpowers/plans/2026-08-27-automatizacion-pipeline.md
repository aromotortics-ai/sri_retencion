# Automatización Pipeline SRI — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Convertir el pipeline local de conciliación SRI en una aplicación web desplegable en VPS donde usuarios autorizados suben dos archivos y descargan el Excel resultado.

**Architecture:** FastAPI + gunicorn + nginx en VPS. Jobs de pipeline corren en hilos background; su estado se persiste en `jobs/{id}/status.json` para ser compartido entre workers. XMLs del SRI se cachean en `cache/xml/` para reutilización entre períodos.

**Tech Stack:** FastAPI, uvicorn, gunicorn, bcrypt, pyyaml, python-multipart, pytest, httpx. Los scripts de pipeline existentes se refactorizan para exponer funciones llamables (en vez de ejecutar lógica a nivel de módulo).

---

## File Map

| Archivo | Acción | Responsabilidad |
|---|---|---|
| `build_workbook_final.py` | **Modificar** | Envolver lógica en `build_workbook(txt_path, xml_glob, output_path) → dict` |
| `conciliar_odoo.py` | **Modificar** | Envolver lógica en `conciliar(excel_path, reten_odoo_path) → dict` |
| `auth.py` | **Crear** | Cargar `users.yaml`, verificar bcrypt; función `verify_user(username, password) → bool` |
| `job_manager.py` | **Crear** | Crear jobs, escribir/leer `status.json`, limpiar jobs expirados |
| `pipeline.py` | **Crear** | Función `run_pipeline(job_id)` que orquesta los 4 pasos con caché y progreso |
| `app.py` | **Crear** | FastAPI: endpoints `/`, `/upload`, `/status/{id}`, `/download/{id}` |
| `templates/index.html` | **Crear** | UI single-page: formulario, barra progreso, resumen, descarga |
| `setup_usuario.py` | **Crear** | CLI para añadir/cambiar/listar usuarios en `users.yaml` |
| `users.yaml` | **Crear** | Archivo de credenciales (generado por `setup_usuario.py`) |
| `deploy/sri-conciliacion.service` | **Crear** | Plantilla systemd |
| `deploy/nginx.conf` | **Crear** | Plantilla nginx con TLS |
| `deploy/install.sh` | **Crear** | Script de instalación en VPS limpio |
| `tests/test_auth.py` | **Crear** | Tests de autenticación |
| `tests/test_job_manager.py` | **Crear** | Tests del gestor de jobs |
| `tests/test_app.py` | **Crear** | Tests de endpoints FastAPI |

---

## Task 1: Instalar dependencias y estructura de carpetas

**Files:**
- Modify: `requirements.txt` (crear si no existe)

- [ ] **Step 1: Instalar dependencias nuevas**

```bash
cd C:/Users/User/Documents/2026/sri_conciliacion
pip install fastapi uvicorn gunicorn bcrypt pyyaml python-multipart pytest httpx
```

- [ ] **Step 2: Crear requirements.txt**

```
pandas
openpyxl
requests
certifi
fastapi
uvicorn[standard]
gunicorn
bcrypt
pyyaml
python-multipart
pytest
httpx
```

Guardar en `requirements.txt`.

- [ ] **Step 3: Crear carpetas necesarias**

```bash
mkdir -p "C:/Users/User/Documents/2026/sri_conciliacion/templates"
mkdir -p "C:/Users/User/Documents/2026/sri_conciliacion/cache/xml"
mkdir -p "C:/Users/User/Documents/2026/sri_conciliacion/jobs"
mkdir -p "C:/Users/User/Documents/2026/sri_conciliacion/tests"
mkdir -p "C:/Users/User/Documents/2026/sri_conciliacion/deploy"
```

- [ ] **Step 4: Commit**

```bash
git add requirements.txt
git commit -m "chore: add web app dependencies"
```

---

## Task 2: Refactorizar `build_workbook_final.py`

La lógica actual corre a nivel de módulo con rutas hardcodeadas. La envolveremos en una función llamable que acepta rutas como parámetros.

**Files:**
- Modify: `build_workbook_final.py`
- Create: `tests/test_build_workbook.py`

- [ ] **Step 1: Escribir el test**

Crear `tests/test_build_workbook.py`:

```python
import pytest
from pathlib import Path
from build_workbook_final import build_workbook

FIXTURES = Path(__file__).parent.parent

def test_build_workbook_returns_stats(tmp_path):
    """build_workbook debe retornar dict con n_listado y n_xml."""
    result = build_workbook(
        txt_path=str(FIXTURES / "2390000831001_Recibidos.txt"),
        xml_glob=str(FIXTURES / "xml_samples/[0-9]*.xml"),
        output_path=str(tmp_path / "out.xlsx"),
    )
    assert "n_listado" in result
    assert "n_xml" in result
    assert result["n_listado"] > 0
    assert Path(tmp_path / "out.xlsx").exists()

def test_build_workbook_file_created(tmp_path):
    """El archivo Excel de salida debe crearse en la ruta indicada."""
    out = tmp_path / "subdir" / "resultado.xlsx"
    out.parent.mkdir()
    build_workbook(
        txt_path=str(FIXTURES / "2390000831001_Recibidos.txt"),
        xml_glob=str(FIXTURES / "xml_samples/[0-9]*.xml"),
        output_path=str(out),
    )
    assert out.exists()
    assert out.stat().st_size > 0
```

- [ ] **Step 2: Verificar que falla**

```bash
cd C:/Users/User/Documents/2026/sri_conciliacion
python -m pytest tests/test_build_workbook.py -v
```

Expected: `ImportError: cannot import name 'build_workbook'`

- [ ] **Step 3: Refactorizar `build_workbook_final.py`**

Envolver toda la lógica actual en una función. Reemplazar el contenido completo por:

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from parse_recibidos_txt import parse_recibidos_txt
from parse_comprobante_xml import parse_folder

FONT = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
WARN_FILL = PatternFill("solid", fgColor="C00000")
HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
BODY_FONT = Font(name=FONT, size=10)
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def escribir_encabezado(ws, headers):
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=j, value=h.replace("_", " "))
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER


def build_workbook(txt_path: str, xml_glob: str, output_path: str) -> dict:
    """
    Genera el Excel de conciliación SRI.

    Args:
        txt_path:    ruta al archivo RUC_Recibidos.txt
        xml_glob:    patrón glob para los XML descargados (ej. 'cache/xml/[0-9]*.xml')
        output_path: ruta donde guardar el Excel resultante

    Returns:
        dict con claves: n_listado, n_xml, n_pendientes
    """
    df_listado = parse_recibidos_txt(txt_path)
    df_cab, df_det = parse_folder(xml_glob)

    claves_con_xml = set(df_cab["Clave_Acceso"]) if len(df_cab) else set()
    df_listado["XML_Descargado"] = df_listado["CLAVE_ACCESO"].isin(claves_con_xml)

    if len(df_det):
        # N° Retencion: Establecimiento-Punto_Emision-Secuencial
        df_det.insert(
            df_det.columns.get_loc("Secuencial") + 1,
            "N_Retencion",
            df_det["Establecimiento"] + "-" + df_det["Punto_Emision"] + "-" + df_det["Secuencial"],
        )
        # Num_Doc_Sustento con guiones: 3-3-resto
        df_det["Num_Doc_Sustento"] = df_det["Num_Doc_Sustento"].apply(
            lambda v: f"{v[:3]}-{v[3:6]}-{v[6:]}" if isinstance(v, str) and len(v) >= 6 else v
        )
        df_agg = (
            df_det.groupby(
                ["Num_Doc_Sustento", "RUC_Emisor", "Razon_Social_Emisor", "Fecha_Emision_Doc_Sustento"],
                as_index=False,
            ).agg(Valor_Total_Retenido=("Valor_Retenido", "sum"))
        )
    else:
        df_agg = df_det

    wb = Workbook()

    # ---------- Hoja 1: SRI_Listado ----------
    ws1 = wb.active
    ws1.title = "SRI_Listado"
    headers1 = list(df_listado.columns)
    escribir_encabezado(ws1, headers1)
    col_xml_desc = headers1.index("XML_Descargado") + 1
    for i, row in enumerate(df_listado.itertuples(index=False), start=2):
        for j, val in enumerate(row, start=1):
            c = ws1.cell(row=i, column=j, value=val)
            c.font = BODY_FONT
            c.border = BORDER
        c_flag = ws1.cell(row=i, column=col_xml_desc)
        if not c_flag.value:
            c_flag.fill = WARN_FILL
            c_flag.font = Font(name=FONT, size=10, color="FFFFFF", bold=True)
    for j in range(1, len(headers1) + 1):
        ws1.column_dimensions[get_column_letter(j)].width = 20
    ws1.freeze_panes = "A2"
    n_listado = len(df_listado)

    resumen_row = n_listado + 3
    ws1.cell(row=resumen_row, column=1, value="Total comprobantes del período:").font = Font(name=FONT, bold=True, size=10)
    ws1.cell(row=resumen_row, column=2, value=f"=COUNTA(A2:A{1+n_listado})").font = Font(name=FONT, bold=True, size=10)
    ws1.cell(row=resumen_row + 1, column=1, value="Con XML descargado:").font = Font(name=FONT, bold=True, size=10)
    ws1.cell(row=resumen_row + 1, column=2, value=f"=COUNTIF({get_column_letter(col_xml_desc)}2:{get_column_letter(col_xml_desc)}{1+n_listado},TRUE)").font = Font(name=FONT, bold=True, size=10)
    ws1.cell(row=resumen_row + 2, column=1, value="Pendientes de descargar:").font = Font(name=FONT, bold=True, size=10)
    ws1.cell(row=resumen_row + 2, column=2, value=f"=COUNTIF({get_column_letter(col_xml_desc)}2:{get_column_letter(col_xml_desc)}{1+n_listado},FALSE)").font = Font(name=FONT, bold=True, size=10)

    # ---------- Hoja 2: SRI_Detalle_Retenciones ----------
    ws2 = wb.create_sheet("SRI_Detalle_Retenciones")
    if len(df_det):
        headers2 = list(df_det.columns)
        escribir_encabezado(ws2, headers2)
        for i, row in enumerate(df_det.itertuples(index=False), start=2):
            for j, val in enumerate(row, start=1):
                c = ws2.cell(row=i, column=j, value=val)
                c.font = BODY_FONT
                c.border = BORDER
                if headers2[j - 1] in ("Base_Imponible", "Valor_Retenido"):
                    c.number_format = "#,##0.00"
        for j in range(1, len(headers2) + 1):
            ws2.column_dimensions[get_column_letter(j)].width = 20
        ws2.freeze_panes = "A2"
    else:
        ws2["A1"] = "Aún no hay XML descargados."

    # ---------- Hoja 3: SRI_Por_Doc_Sustento ----------
    ws3 = wb.create_sheet("SRI_Por_Doc_Sustento")
    n_agg = len(df_agg)
    if n_agg:
        headers3 = list(df_agg.columns)
        escribir_encabezado(ws3, headers3)
        for i, row in enumerate(df_agg.itertuples(index=False), start=2):
            for j, val in enumerate(row, start=1):
                c = ws3.cell(row=i, column=j, value=val)
                c.font = BODY_FONT
                c.border = BORDER
                if headers3[j - 1] == "Valor_Total_Retenido":
                    c.number_format = "#,##0.00"
        for j in range(1, len(headers3) + 1):
            ws3.column_dimensions[get_column_letter(j)].width = 22
        ws3.freeze_panes = "A2"
    else:
        ws3["A1"] = "Sin datos todavía."

    # ---------- Hoja 4: ERP_Compras (plantilla) ----------
    ws4 = wb.create_sheet("ERP_Compras")
    headers4 = ["Num_Documento_Compra", "RUC_Proveedor", "Razon_Social",
                "Fecha_Comprobante", "Base_Imponible", "Total_Retenido_ERP"]
    escribir_encabezado(ws4, headers4)
    ejemplo = ["001003000088744", "1001789088001", "EJEMPLO PROVEEDOR S.A.",
               "2026-07-31", 109.36, 2.24]
    for j, val in enumerate(ejemplo, start=1):
        c = ws4.cell(row=2, column=j, value=val)
        c.font = Font(name=FONT, size=10, italic=True, color="808080")
        c.border = BORDER
    N_ERP = 600
    for r in range(3, 3 + N_ERP):
        for j in range(1, len(headers4) + 1):
            c = ws4.cell(row=r, column=j)
            c.fill = INPUT_FILL
            c.border = BORDER
            c.font = BODY_FONT
    for j, w in enumerate([22, 16, 32, 16, 14, 16], start=1):
        ws4.column_dimensions[get_column_letter(j)].width = w
    ws4.freeze_panes = "A2"
    n_erp_start, n_erp_end = 3, 3 + N_ERP - 1

    # ---------- Hoja 5: Conciliacion ----------
    ws5 = wb.create_sheet("Conciliacion")
    headers5 = ["Num_Doc_Sustento", "RUC_Emisor", "Razon_Social_Emisor",
                "Valor_Retenido_SRI", "Total_Retenido_ERP", "Diferencia", "Estado"]
    escribir_encabezado(ws5, headers5)
    if n_agg:
        for i in range(2, 2 + n_agg):
            ws5.cell(row=i, column=1, value=f"=SRI_Por_Doc_Sustento!A{i}")
            ws5.cell(row=i, column=2, value=f"=SRI_Por_Doc_Sustento!B{i}")
            ws5.cell(row=i, column=3, value=f"=SRI_Por_Doc_Sustento!C{i}")
            ws5.cell(row=i, column=4, value=f"=SRI_Por_Doc_Sustento!E{i}")
            ws5.cell(
                row=i, column=5,
                value=(f'=IFERROR(INDEX(ERP_Compras!$F${n_erp_start}:$F${n_erp_end},'
                       f'MATCH(A{i},ERP_Compras!$A${n_erp_start}:$A${n_erp_end},0)),"")')
            )
            ws5.cell(row=i, column=6, value=f'=IF(E{i}="","",D{i}-E{i})')
            ws5.cell(
                row=i, column=7,
                value=(f'=IF(E{i}="","No encontrado en ERP",'
                       f'IF(ABS(D{i}-E{i})<0.01,"Coincide","Diferencia de valor"))')
            )
            for col in range(1, 8):
                c = ws5.cell(row=i, column=col)
                c.font = BODY_FONT
                c.border = BORDER
                if headers5[col - 1] in ("Valor_Retenido_SRI", "Total_Retenido_ERP", "Diferencia"):
                    c.number_format = "#,##0.00"
    for j, w in enumerate([22, 16, 32, 18, 18, 12, 20], start=1):
        ws5.column_dimensions[get_column_letter(j)].width = w
    ws5.freeze_panes = "A2"

    wb.save(output_path)

    return {
        "n_listado": n_listado,
        "n_xml": len(df_cab),
        "n_pendientes": n_listado - len(df_cab),
    }


if __name__ == "__main__":
    stats = build_workbook(
        txt_path="2390000831001_Recibidos.txt",
        xml_glob="xml_samples/[0-9]*.xml",
        output_path="Comprobantes_Retencion_Agosto2026_Final.xlsx",
    )
    print(f"OK. Listado: {stats['n_listado']} | XML: {stats['n_xml']} | Pendientes: {stats['n_pendientes']}")
```

- [ ] **Step 4: Correr los tests**

```bash
python -m pytest tests/test_build_workbook.py -v
```

Expected: 2 tests PASSED

- [ ] **Step 5: Commit**

```bash
git add build_workbook_final.py tests/test_build_workbook.py
git commit -m "refactor: expose build_workbook() function with path parameters"
```

---

## Task 3: Refactorizar `conciliar_odoo.py`

**Files:**
- Modify: `conciliar_odoo.py`
- Create: `tests/test_conciliar.py`

- [ ] **Step 1: Escribir el test**

Crear `tests/test_conciliar.py`:

```python
import pytest
from pathlib import Path
from conciliar_odoo import conciliar

FIXTURES = Path(__file__).parent.parent

def test_conciliar_returns_summary(tmp_path):
    """conciliar debe retornar dict con conteos y agregar hojas al Excel."""
    import shutil
    excel = tmp_path / "resultado.xlsx"
    shutil.copy(FIXTURES / "Comprobantes_Retencion_Agosto2026_Final.xlsx", excel)

    result = conciliar(
        excel_path=str(excel),
        reten_odoo_path=str(FIXTURES / "reten_odoo.xlsx"),
    )

    assert "n_sri" in result
    assert "n_odoo" in result
    assert "coinciden" in result
    assert "solo_sri" in result
    assert "solo_odoo" in result
    assert "fac_incorrecta" in result
    assert result["n_sri"] > 0

def test_conciliar_adds_sheets(tmp_path):
    """conciliar debe añadir las 4 hojas de resultado al Excel."""
    import shutil, openpyxl
    excel = tmp_path / "resultado.xlsx"
    shutil.copy(FIXTURES / "Comprobantes_Retencion_Agosto2026_Final.xlsx", excel)
    conciliar(str(excel), str(FIXTURES / "reten_odoo.xlsx"))

    wb = openpyxl.load_workbook(str(excel))
    assert "Conciliacion_Odoo" in wb.sheetnames
    assert "Solo_en_SRI" in wb.sheetnames
    assert "Solo_en_Odoo" in wb.sheetnames
    assert "Diferencias_Valor" in wb.sheetnames
```

- [ ] **Step 2: Verificar que falla**

```bash
python -m pytest tests/test_conciliar.py -v
```

Expected: `ImportError: cannot import name 'conciliar'`

- [ ] **Step 3: Refactorizar `conciliar_odoo.py`**

Envolver toda la lógica actual en una función `conciliar(excel_path, reten_odoo_path) -> dict`. Reemplazar el contenido manteniendo todas las funciones auxiliares (`_hdr`, `_write_df`) y añadiendo:

```python
def conciliar(excel_path: str, reten_odoo_path: str) -> dict:
    """
    Cruza SRI_Detalle_Retenciones contra el reporte de Odoo.
    Añade 4 hojas de resultado al Excel en excel_path.

    Returns:
        dict con claves: n_sri, n_odoo, coinciden, solo_sri, solo_odoo,
                         fac_incorrecta, diff_valor
    """
    global EXCEL_FINAL, EXCEL_ODOO
    EXCEL_FINAL = excel_path
    EXCEL_ODOO  = reten_odoo_path
    # ... todo el código actual de las constantes hacia abajo, hasta wb_main.save() ...
    # Al final, en lugar de solo print(), también retornar:
    return {
        "n_sri":        len(df_sri),
        "n_odoo":       len(df_odoo),
        "coinciden":    len(df_ok),
        "solo_sri":     len(df_solo_sri),
        "solo_odoo":    len(df_solo_odoo),
        "fac_incorrecta": len(df_fac_error) // 2,
        "diff_valor":   len(df_diff_val),
    }


if __name__ == "__main__":
    import sys
    excel = sys.argv[1] if len(sys.argv) > 1 else EXCEL_FINAL
    odoo  = sys.argv[2] if len(sys.argv) > 2 else EXCEL_ODOO
    resumen = conciliar(excel, odoo)
    print("=" * 60)
    for k, v in resumen.items():
        print(f"  {k:30s}: {v}")
```

**Nota:** Todo el bloque de código que actualmente está a nivel de módulo (desde `wb_main = openpyxl.load_workbook(EXCEL_FINAL)` hasta `wb_main.save(EXCEL_FINAL)`) se mueve dentro de `conciliar()`. Las constantes `EXCEL_FINAL` y `EXCEL_ODOO` pasan a ser parámetros de la función.

- [ ] **Step 4: Correr los tests**

```bash
python -m pytest tests/test_conciliar.py -v
```

Expected: 2 tests PASSED

- [ ] **Step 5: Commit**

```bash
git add conciliar_odoo.py tests/test_conciliar.py
git commit -m "refactor: expose conciliar() function with path parameters"
```

---

## Task 4: Módulo de autenticación (`auth.py`)

**Files:**
- Create: `auth.py`
- Create: `tests/test_auth.py`
- Create: `users.yaml` (archivo de ejemplo para tests)

- [ ] **Step 1: Escribir los tests**

Crear `tests/test_auth.py`:

```python
import pytest
import bcrypt
import yaml
from pathlib import Path
from unittest.mock import patch


def _make_users_yaml(tmp_path, users: list) -> Path:
    """Crea un users.yaml temporal con los usuarios dados."""
    data = {"users": []}
    for username, password in users:
        hashed = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
        data["users"].append({"username": username, "password_hash": hashed})
    f = tmp_path / "users.yaml"
    f.write_text(yaml.dump(data))
    return f


def test_verify_user_valid(tmp_path):
    from auth import verify_user
    users_file = _make_users_yaml(tmp_path, [("ana", "secreta123")])
    assert verify_user("ana", "secreta123", users_file=str(users_file)) is True


def test_verify_user_wrong_password(tmp_path):
    from auth import verify_user
    users_file = _make_users_yaml(tmp_path, [("ana", "secreta123")])
    assert verify_user("ana", "wrongpass", users_file=str(users_file)) is False


def test_verify_user_unknown_user(tmp_path):
    from auth import verify_user
    users_file = _make_users_yaml(tmp_path, [("ana", "secreta123")])
    assert verify_user("hacker", "secreta123", users_file=str(users_file)) is False


def test_verify_user_empty_file(tmp_path):
    from auth import verify_user
    f = tmp_path / "users.yaml"
    f.write_text("users: []")
    assert verify_user("anyone", "pass", users_file=str(f)) is False
```

- [ ] **Step 2: Verificar que falla**

```bash
python -m pytest tests/test_auth.py -v
```

Expected: `ModuleNotFoundError: No module named 'auth'`

- [ ] **Step 3: Crear `auth.py`**

```python
"""
Autenticación HTTP Basic Auth contra users.yaml con contraseñas bcrypt.
"""
from pathlib import Path
import bcrypt
import yaml

DEFAULT_USERS_FILE = str(Path(__file__).parent / "users.yaml")


def verify_user(username: str, password: str, users_file: str = DEFAULT_USERS_FILE) -> bool:
    """
    Verifica que username/password coinciden con una entrada en users_file.

    Args:
        username:   nombre de usuario
        password:   contraseña en texto plano
        users_file: ruta al archivo YAML de usuarios (default: users.yaml junto a auth.py)

    Returns:
        True si las credenciales son válidas, False en cualquier otro caso.
    """
    try:
        data = yaml.safe_load(Path(users_file).read_text(encoding="utf-8"))
        users = data.get("users", [])
    except (FileNotFoundError, yaml.YAMLError):
        return False

    for user in users:
        if user.get("username") == username:
            stored_hash = user.get("password_hash", "").encode()
            try:
                return bcrypt.checkpw(password.encode(), stored_hash)
            except Exception:
                return False
    return False
```

- [ ] **Step 4: Correr los tests**

```bash
python -m pytest tests/test_auth.py -v
```

Expected: 4 tests PASSED

- [ ] **Step 5: Crear `users.yaml` de ejemplo**

```yaml
# users.yaml — editar con setup_usuario.py
# NUNCA commitear este archivo con contraseñas reales
users: []
```

Crear `.gitignore` si no existe:

```
users.yaml
jobs/
cache/
__pycache__/
*.pyc
.env
```

- [ ] **Step 6: Commit**

```bash
git add auth.py tests/test_auth.py .gitignore
git commit -m "feat: add HTTP Basic Auth module with bcrypt verification"
```

---

## Task 5: Gestor de jobs (`job_manager.py`)

**Files:**
- Create: `job_manager.py`
- Create: `tests/test_job_manager.py`

- [ ] **Step 1: Escribir los tests**

Crear `tests/test_job_manager.py`:

```python
import pytest
import json
import time
from pathlib import Path
from unittest.mock import patch


def test_create_job_returns_id(tmp_path):
    from job_manager import create_job
    with patch("job_manager.JOBS_DIR", tmp_path):
        job_id = create_job()
    assert isinstance(job_id, str)
    assert len(job_id) == 8
    assert (tmp_path / job_id).is_dir()


def test_write_and_read_status(tmp_path):
    from job_manager import create_job, write_status, read_status
    with patch("job_manager.JOBS_DIR", tmp_path):
        job_id = create_job()
        write_status(job_id, "running", "Procesando...", 0.5)
        status = read_status(job_id)

    assert status["estado"] == "running"
    assert status["mensaje"] == "Procesando..."
    assert status["progreso"] == pytest.approx(0.5)


def test_read_status_missing_job(tmp_path):
    from job_manager import read_status
    with patch("job_manager.JOBS_DIR", tmp_path):
        result = read_status("noexiste")
    assert result is None


def test_write_status_done_with_resumen(tmp_path):
    from job_manager import create_job, write_status, read_status
    resumen = {"coinciden": 620, "solo_sri": 317}
    with patch("job_manager.JOBS_DIR", tmp_path):
        job_id = create_job()
        write_status(job_id, "done", "Completado", 1.0, resumen=resumen)
        status = read_status(job_id)

    assert status["estado"] == "done"
    assert status["resumen"]["coinciden"] == 620


def test_get_job_dir(tmp_path):
    from job_manager import create_job, get_job_dir
    with patch("job_manager.JOBS_DIR", tmp_path):
        job_id = create_job()
        d = get_job_dir(job_id)
    assert d == tmp_path / job_id
```

- [ ] **Step 2: Verificar que falla**

```bash
python -m pytest tests/test_job_manager.py -v
```

Expected: `ModuleNotFoundError: No module named 'job_manager'`

- [ ] **Step 3: Crear `job_manager.py`**

```python
"""
Gestión de jobs de pipeline: crear, actualizar estado, leer estado, limpiar.
El estado se persiste en jobs/{job_id}/status.json para ser visible entre
los múltiples workers de gunicorn (no usar estado en memoria).
"""
import json
import uuid
import shutil
import threading
from datetime import datetime
from pathlib import Path

JOBS_DIR = Path(__file__).parent / "jobs"


def create_job() -> str:
    """Crea la carpeta del job y devuelve su ID (8 hex chars)."""
    job_id = uuid.uuid4().hex[:8]
    get_job_dir(job_id).mkdir(parents=True, exist_ok=True)
    write_status(job_id, "pending", "En cola...", 0.0)
    return job_id


def get_job_dir(job_id: str) -> Path:
    return JOBS_DIR / job_id


def write_status(
    job_id: str,
    estado: str,
    mensaje: str,
    progreso: float = 0.0,
    resumen: dict = None,
    error: str = None,
) -> None:
    """
    Escribe el estado del job en jobs/{job_id}/status.json.

    Args:
        job_id:   identificador del job
        estado:   'pending' | 'running' | 'done' | 'error'
        mensaje:  texto de progreso para mostrar al usuario
        progreso: float 0.0-1.0 para la barra de progreso
        resumen:  dict con estadísticas finales (solo cuando estado='done')
        error:    mensaje de error (solo cuando estado='error')
    """
    status = {
        "estado":   estado,
        "mensaje":  mensaje,
        "progreso": round(progreso, 3),
        "resumen":  resumen,
        "error":    error,
        "ts":       datetime.utcnow().isoformat(),
    }
    status_file = get_job_dir(job_id) / "status.json"
    # Escritura atómica: escribir a .tmp y renombrar
    tmp = status_file.with_suffix(".tmp")
    tmp.write_text(json.dumps(status, ensure_ascii=False), encoding="utf-8")
    tmp.replace(status_file)


def read_status(job_id: str) -> dict | None:
    """Lee el status.json del job. Retorna None si el job no existe."""
    status_file = get_job_dir(job_id) / "status.json"
    if not status_file.exists():
        return None
    try:
        return json.loads(status_file.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None


def schedule_cleanup(job_id: str, delay_seconds: int = 7200) -> None:
    """Programa la eliminación del job tras delay_seconds (default 2 horas)."""
    def _cleanup():
        job_dir = get_job_dir(job_id)
        if job_dir.exists():
            shutil.rmtree(job_dir, ignore_errors=True)

    t = threading.Timer(delay_seconds, _cleanup)
    t.daemon = True
    t.start()


def mark_stale_jobs_as_error() -> None:
    """
    Marca como error todos los jobs 'pending' o 'running' encontrados al iniciar.
    Llamar una vez al arrancar la app (los hilos de esos jobs murieron al reiniciar).
    """
    if not JOBS_DIR.exists():
        return
    for job_dir in JOBS_DIR.iterdir():
        if not job_dir.is_dir():
            continue
        status = read_status(job_dir.name)
        if status and status.get("estado") in ("pending", "running"):
            write_status(
                job_dir.name,
                "error",
                "El servidor se reinició mientras el job estaba en progreso. Vuelve a subir los archivos.",
            )
```

- [ ] **Step 4: Correr los tests**

```bash
python -m pytest tests/test_job_manager.py -v
```

Expected: 5 tests PASSED

- [ ] **Step 5: Commit**

```bash
git add job_manager.py tests/test_job_manager.py
git commit -m "feat: add file-based job manager for multi-worker gunicorn compatibility"
```

---

## Task 6: Runner del pipeline (`pipeline.py`)

**Files:**
- Create: `pipeline.py`

No hay test unitario para este módulo porque orquesta I/O real (red, disco). Se verifica con la prueba de integración de la app completa en Task 8.

- [ ] **Step 1: Crear `pipeline.py`**

```python
"""
Runner del pipeline de conciliación SRI.
Se ejecuta en un hilo background lanzado por app.py.
"""
import re
import time
import threading
from pathlib import Path

from job_manager import write_status, schedule_cleanup, get_job_dir
from sri_session import sri_session

CACHE_XML_DIR = Path(__file__).parent / "cache" / "xml"

ENDPOINT = "https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline"
SOAP_TEMPLATE = """<?xml version="1.0" encoding="UTF-8"?>
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ec="http://ec.gob.sri.ws.autorizacion">
  <soapenv:Header/>
  <soapenv:Body>
    <ec:autorizacionComprobante>
      <claveAccesoComprobante>{clave}</claveAccesoComprobante>
    </ec:autorizacionComprobante>
  </soapenv:Body>
</soapenv:Envelope>"""
HEADERS = {"Content-Type": "text/xml; charset=utf-8", "Connection": "close"}


def run_pipeline(job_id: str) -> None:
    """
    Orquesta los 4 pasos del pipeline para un job dado.
    Escribe progreso en status.json cada vez que avanza.
    Llamar en un hilo daemon.
    """
    job_dir = get_job_dir(job_id)
    try:
        # ── Paso 1: Leer TXT ─────────────────────────────────────────────────
        write_status(job_id, "running", "Leyendo TXT del SRI...", 0.02)
        from parse_recibidos_txt import parse_recibidos_txt
        df = parse_recibidos_txt(str(job_dir / "Recibidos.txt"))
        claves = df["CLAVE_ACCESO"].dropna().tolist()
        write_status(job_id, "running", f"TXT leído: {len(claves)} claves de acceso", 0.05)

        # ── Paso 2: Descargar XMLs (con caché) ───────────────────────────────
        CACHE_XML_DIR.mkdir(parents=True, exist_ok=True)
        pendientes = [c for c in claves if not (CACHE_XML_DIR / f"{c}.xml").exists()]
        ya_cache = len(claves) - len(pendientes)

        session = sri_session()
        ok = ya_cache
        fallidos = []

        for i, clave in enumerate(pendientes):
            body = SOAP_TEMPLATE.format(clave=clave).encode("utf-8")
            try:
                r = session.post(ENDPOINT, data=body, headers=HEADERS, timeout=20)
                r.raise_for_status()
                match = re.search(r"<autorizacion>.*?</autorizacion>", r.text, re.DOTALL)
                if match:
                    dest = CACHE_XML_DIR / f"{clave}.xml"
                    dest.write_text(
                        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + match.group(0),
                        encoding="utf-8",
                    )
                    ok += 1
                else:
                    fallidos.append(clave)
            except Exception as e:
                fallidos.append(clave)
            time.sleep(0.5)

            # Actualizar progreso cada 10 descargas
            if (i + 1) % 10 == 0 or i == len(pendientes) - 1:
                progreso = 0.05 + (ok / max(len(claves), 1)) * 0.70
                msg = f"Descargando XMLs del SRI ({ok}/{len(claves)}"
                if fallidos:
                    msg += f", {len(fallidos)} fallidos"
                if ya_cache:
                    msg += f", {ya_cache} desde caché"
                msg += ")"
                write_status(job_id, "running", msg, progreso)

        # ── Paso 3: Armar Excel ───────────────────────────────────────────────
        write_status(job_id, "running", "Armando Excel de conciliación...", 0.80)
        xml_glob = str(CACHE_XML_DIR / "[0-9]*.xml")
        output_path = str(job_dir / "resultado.xlsx")
        from build_workbook_final import build_workbook
        stats = build_workbook(
            txt_path=str(job_dir / "Recibidos.txt"),
            xml_glob=xml_glob,
            output_path=output_path,
        )

        # ── Paso 4: Conciliar con Odoo ────────────────────────────────────────
        write_status(job_id, "running", "Conciliando con Odoo...", 0.92)
        from conciliar_odoo import conciliar
        resumen = conciliar(
            excel_path=output_path,
            reten_odoo_path=str(job_dir / "reten_odoo.xlsx"),
        )
        resumen["n_listado"] = stats["n_listado"]
        resumen["n_xml"]     = stats["n_xml"]
        resumen["fallidos"]  = len(fallidos)

        # ── Completado ────────────────────────────────────────────────────────
        write_status(job_id, "done", "Completado", 1.0, resumen=resumen)
        schedule_cleanup(job_id, delay_seconds=7200)

    except Exception as exc:
        import traceback
        write_status(
            job_id,
            "error",
            f"{type(exc).__name__}: {exc}",
            error=traceback.format_exc(),
        )
```

- [ ] **Step 2: Verificar que importa sin error**

```bash
python -c "from pipeline import run_pipeline; print('OK')"
```

Expected: `OK`

- [ ] **Step 3: Commit**

```bash
git add pipeline.py
git commit -m "feat: add pipeline runner with XML cache and per-step progress reporting"
```

---

## Task 7: Aplicación FastAPI (`app.py`)

**Files:**
- Create: `app.py`
- Create: `tests/test_app.py`

- [ ] **Step 1: Escribir los tests**

Crear `tests/test_app.py`:

```python
import pytest
import bcrypt
import yaml
from pathlib import Path
from fastapi.testclient import TestClient
import base64

FIXTURES = Path(__file__).parent.parent


def _auth(username="testuser", password="testpass"):
    token = base64.b64encode(f"{username}:{password}".encode()).decode()
    return {"Authorization": f"Basic {token}"}


@pytest.fixture
def users_yaml(tmp_path):
    hashed = bcrypt.hashpw(b"testpass", bcrypt.gensalt()).decode()
    data = {"users": [{"username": "testuser", "password_hash": hashed}]}
    f = tmp_path / "users.yaml"
    f.write_text(yaml.dump(data))
    return f


@pytest.fixture
def client(users_yaml, tmp_path, monkeypatch):
    monkeypatch.setenv("USERS_FILE", str(users_yaml))
    monkeypatch.setattr("job_manager.JOBS_DIR", tmp_path / "jobs")
    import app as app_module
    monkeypatch.setattr(app_module, "USERS_FILE", str(users_yaml))
    return TestClient(app_module.app, raise_server_exceptions=False)


def test_root_requires_auth(client):
    r = client.get("/", auth=None)
    assert r.status_code == 401


def test_root_returns_html(client):
    r = client.get("/", auth=("testuser", "testpass"))
    assert r.status_code == 200
    assert "text/html" in r.headers["content-type"]


def test_upload_requires_auth(client):
    r = client.post("/upload")
    assert r.status_code == 401


def test_upload_missing_files(client):
    r = client.post("/upload", auth=("testuser", "testpass"))
    assert r.status_code == 422   # FastAPI validation error


def test_status_unknown_job(client):
    r = client.get("/status/xxxxxxxx", auth=("testuser", "testpass"))
    assert r.status_code == 404


def test_download_unknown_job(client):
    r = client.get("/download/xxxxxxxx", auth=("testuser", "testpass"))
    assert r.status_code == 404


def test_upload_creates_job(client, tmp_path, monkeypatch):
    """Upload con archivos válidos debe retornar job_id y estado pending."""
    import app as app_module
    # Mock run_pipeline para no ejecutar el pipeline real
    monkeypatch.setattr(app_module, "run_pipeline", lambda job_id: None)

    txt_content = (FIXTURES / "2390000831001_Recibidos.txt").read_bytes()
    xlsx_content = (FIXTURES / "reten_odoo.xlsx").read_bytes()

    r = client.post(
        "/upload",
        auth=("testuser", "testpass"),
        files={
            "recibidos": ("Recibidos.txt", txt_content, "text/plain"),
            "reten_odoo": ("reten_odoo.xlsx", xlsx_content,
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        },
    )
    assert r.status_code == 200
    data = r.json()
    assert "job_id" in data
    assert len(data["job_id"]) == 8

    # Status debe existir
    job_id = data["job_id"]
    r2 = client.get(f"/status/{job_id}", auth=("testuser", "testpass"))
    assert r2.status_code == 200
    assert r2.json()["estado"] in ("pending", "running", "done")
```

- [ ] **Step 2: Verificar que falla**

```bash
python -m pytest tests/test_app.py -v
```

Expected: `ModuleNotFoundError: No module named 'app'`

- [ ] **Step 3: Crear `app.py`**

```python
"""
Aplicación FastAPI para el pipeline de conciliación SRI.
Endpoints: GET /, POST /upload, GET /status/{job_id}, GET /download/{job_id}
Auth: HTTP Basic Auth contra users.yaml (bcrypt)
"""
import threading
from pathlib import Path

from fastapi import FastAPI, Depends, HTTPException, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.security import HTTPBasic, HTTPBasicCredentials

from auth import verify_user
from job_manager import create_job, read_status, get_job_dir, mark_stale_jobs_as_error
from pipeline import run_pipeline

BASE_DIR    = Path(__file__).parent
USERS_FILE  = str(BASE_DIR / "users.yaml")
TEMPLATE    = BASE_DIR / "templates" / "index.html"

app = FastAPI(title="Conciliación SRI", docs_url=None, redoc_url=None)
security = HTTPBasic()


def get_user(credentials: HTTPBasicCredentials = Depends(security)) -> str:
    if not verify_user(credentials.username, credentials.password, users_file=USERS_FILE):
        raise HTTPException(
            status_code=401,
            detail="Credenciales inválidas",
            headers={"WWW-Authenticate": "Basic"},
        )
    return credentials.username


@app.on_event("startup")
def startup():
    mark_stale_jobs_as_error()


@app.get("/", response_class=HTMLResponse)
def index(_user: str = Depends(get_user)):
    return TEMPLATE.read_text(encoding="utf-8")


@app.post("/upload")
def upload(
    recibidos: UploadFile = File(...),
    reten_odoo: UploadFile = File(...),
    _user: str = Depends(get_user),
):
    job_id  = create_job()
    job_dir = get_job_dir(job_id)

    (job_dir / "Recibidos.txt").write_bytes(recibidos.file.read())
    (job_dir / "reten_odoo.xlsx").write_bytes(reten_odoo.file.read())

    t = threading.Thread(target=run_pipeline, args=(job_id,), daemon=True)
    t.start()

    return {"job_id": job_id}


@app.get("/status/{job_id}")
def status(job_id: str, _user: str = Depends(get_user)):
    s = read_status(job_id)
    if s is None:
        raise HTTPException(status_code=404, detail="Job no encontrado")
    return s


@app.get("/download/{job_id}")
def download(job_id: str, _user: str = Depends(get_user)):
    excel = get_job_dir(job_id) / "resultado.xlsx"
    if not excel.exists():
        raise HTTPException(status_code=404, detail="Resultado no disponible aún")
    return FileResponse(
        path=str(excel),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="Conciliacion_SRI.xlsx",
    )
```

- [ ] **Step 4: Correr los tests**

```bash
python -m pytest tests/test_app.py -v
```

Expected: 7 tests PASSED

- [ ] **Step 5: Commit**

```bash
git add app.py tests/test_app.py
git commit -m "feat: add FastAPI web app with Basic Auth, upload, status and download endpoints"
```

---

## Task 8: Frontend (`templates/index.html`)

**Files:**
- Create: `templates/index.html`

No hay test unitario — se valida manualmente arrancando el servidor.

- [ ] **Step 1: Crear `templates/index.html`**

```html
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Conciliación SRI · Retenciones</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: Arial, sans-serif; background: #f4f6f9; display: flex;
           justify-content: center; align-items: center; min-height: 100vh; padding: 1rem; }
    .card { background: white; border-radius: 8px; box-shadow: 0 2px 12px rgba(0,0,0,.1);
            padding: 2rem; width: 100%; max-width: 480px; }
    h1 { color: #1F4E78; font-size: 1.3rem; margin-bottom: 0.25rem; }
    .subtitle { color: #666; font-size: 0.85rem; margin-bottom: 1.5rem; }
    .field { margin-bottom: 1rem; }
    label { display: block; font-size: 0.85rem; color: #444; margin-bottom: 0.4rem; font-weight: bold; }
    input[type=file] { width: 100%; border: 1px solid #ccc; border-radius: 4px;
                       padding: 0.5rem; font-size: 0.85rem; cursor: pointer; }
    button { width: 100%; padding: 0.75rem; background: #1F4E78; color: white;
             border: none; border-radius: 4px; font-size: 1rem; cursor: pointer;
             margin-top: 0.5rem; transition: background .2s; }
    button:hover { background: #163d5f; }
    button:disabled { background: #999; cursor: not-allowed; }
    .progress-section { display: none; }
    .status-msg { color: #444; font-size: 0.9rem; margin-bottom: 0.75rem; min-height: 1.2em; }
    .bar-bg { background: #e0e7ef; border-radius: 4px; height: 12px; overflow: hidden; }
    .bar-fill { background: #1F4E78; height: 100%; width: 0%; transition: width .4s ease;
                border-radius: 4px; }
    .elapsed { font-size: 0.78rem; color: #888; margin-top: 0.4rem; }
    .result-section { display: none; }
    .result-title { font-size: 1.1rem; color: #1a7a3c; font-weight: bold; margin-bottom: 1rem; }
    .stat-grid { display: grid; grid-template-columns: 1fr auto; gap: 0.3rem 1rem;
                 font-size: 0.88rem; margin-bottom: 1.25rem; }
    .stat-label { color: #555; }
    .stat-value { text-align: right; font-weight: bold; }
    .stat-value.warn { color: #c0392b; }
    .stat-value.ok { color: #1a7a3c; }
    .btn-download { background: #1a7a3c; }
    .btn-download:hover { background: #145c2d; }
    .error-section { display: none; }
    .error-title { color: #c0392b; font-weight: bold; margin-bottom: 0.75rem; }
    .error-msg { background: #fdecea; border-radius: 4px; padding: 0.75rem;
                 font-size: 0.82rem; color: #7b1c1c; white-space: pre-wrap;
                 max-height: 120px; overflow-y: auto; margin-bottom: 1rem; }
    .btn-retry { background: #888; }
    .btn-retry:hover { background: #666; }
  </style>
</head>
<body>
<div class="card">
  <h1>Conciliación SRI · Retenciones</h1>
  <p class="subtitle">Sube los dos archivos y descarga el Excel de conciliación.</p>

  <!-- Formulario -->
  <div id="form-section">
    <div class="field">
      <label>1. Archivo TXT del portal SRI</label>
      <input type="file" id="txt-file" accept=".txt">
    </div>
    <div class="field">
      <label>2. Reporte de retenciones de Odoo (.xlsx)</label>
      <input type="file" id="xlsx-file" accept=".xlsx">
    </div>
    <button id="btn-procesar" onclick="startJob()">Procesar</button>
  </div>

  <!-- Progreso -->
  <div class="progress-section" id="progress-section">
    <p class="status-msg" id="status-msg">Iniciando...</p>
    <div class="bar-bg"><div class="bar-fill" id="bar-fill"></div></div>
    <p class="elapsed" id="elapsed"></p>
  </div>

  <!-- Resultado -->
  <div class="result-section" id="result-section">
    <p class="result-title">✅ Listo</p>
    <div class="stat-grid" id="stat-grid"></div>
    <button class="btn-download" id="btn-download">⬇ Descargar Excel</button>
    <button class="btn-retry" onclick="reset()" style="margin-top:.5rem;">Procesar otro período</button>
  </div>

  <!-- Error -->
  <div class="error-section" id="error-section">
    <p class="error-title">❌ Error al procesar</p>
    <div class="error-msg" id="error-msg"></div>
    <button class="btn-retry" onclick="reset()">Intentar de nuevo</button>
  </div>
</div>

<script>
  let pollInterval = null;
  let startTime   = null;
  let currentJobId = null;

  function show(id) {
    ["form-section","progress-section","result-section","error-section"]
      .forEach(s => document.getElementById(s).style.display = s === id ? "block" : "none");
  }

  async function startJob() {
    const txt  = document.getElementById("txt-file").files[0];
    const xlsx = document.getElementById("xlsx-file").files[0];
    if (!txt || !xlsx) { alert("Selecciona ambos archivos."); return; }

    document.getElementById("btn-procesar").disabled = true;
    const fd = new FormData();
    fd.append("recibidos", txt);
    fd.append("reten_odoo", xlsx);

    try {
      const r = await fetch("/upload", { method: "POST", body: fd });
      if (!r.ok) { throw new Error(`HTTP ${r.status}`); }
      const { job_id } = await r.json();
      currentJobId = job_id;
      startTime = Date.now();
      show("progress-section");
      pollInterval = setInterval(() => poll(job_id), 5000);
      poll(job_id);
    } catch(e) {
      alert("Error al subir archivos: " + e.message);
      document.getElementById("btn-procesar").disabled = false;
    }
  }

  async function poll(job_id) {
    try {
      const r = await fetch(`/status/${job_id}`);
      if (!r.ok) return;
      const data = await r.json();

      document.getElementById("status-msg").textContent = data.mensaje || "";
      document.getElementById("bar-fill").style.width = ((data.progreso || 0) * 100) + "%";

      const elapsed = Math.round((Date.now() - startTime) / 1000);
      const m = Math.floor(elapsed / 60), s = elapsed % 60;
      document.getElementById("elapsed").textContent =
        `Tiempo transcurrido: ${m}m ${String(s).padStart(2,"0")}s`;

      if (data.estado === "done") {
        clearInterval(pollInterval);
        showResult(job_id, data.resumen);
      } else if (data.estado === "error") {
        clearInterval(pollInterval);
        showError(data.mensaje || "Error desconocido");
      }
    } catch(e) { /* red temporalmente caída — reintentar en el siguiente tick */ }
  }

  function showResult(job_id, resumen) {
    show("result-section");
    const grid = document.getElementById("stat-grid");
    const rows = [
      ["Retenciones SRI",           resumen.n_listado,       ""],
      ["XMLs descargados",          resumen.n_xml,           resumen.n_xml === resumen.n_listado ? "ok" : "warn"],
      ["Líneas coinciden",          resumen.coinciden,       "ok"],
      ["Solo en SRI · pendientes",  resumen.solo_sri,        resumen.solo_sri > 0 ? "warn" : "ok"],
      ["Solo en Odoo · revisar",    resumen.solo_odoo,       resumen.solo_odoo > 0 ? "warn" : "ok"],
      ["Factura incorrecta en Odoo",resumen.fac_incorrecta,  resumen.fac_incorrecta > 0 ? "warn" : "ok"],
      ["Diferencias de valor",      resumen.diff_valor,      resumen.diff_valor > 0 ? "warn" : "ok"],
    ];
    grid.innerHTML = rows.map(([label, val, cls]) =>
      `<span class="stat-label">${label}</span>` +
      `<span class="stat-value ${cls}">${val ?? "—"}</span>`
    ).join("");
    document.getElementById("btn-download").onclick = () => {
      window.location = `/download/${job_id}`;
    };
  }

  function showError(msg) {
    show("error-section");
    document.getElementById("error-msg").textContent = msg;
  }

  function reset() {
    clearInterval(pollInterval);
    currentJobId = null;
    document.getElementById("txt-file").value = "";
    document.getElementById("xlsx-file").value = "";
    document.getElementById("btn-procesar").disabled = false;
    document.getElementById("bar-fill").style.width = "0%";
    show("form-section");
  }
</script>
</body>
</html>
```

- [ ] **Step 2: Verificar arranque local**

```bash
cd C:/Users/User/Documents/2026/sri_conciliacion
uvicorn app:app --reload --port 8000
```

Abrir `http://localhost:8000` en el navegador. El browser debe mostrar el diálogo de usuario/contraseña. Verificar que la página carga con el formulario de dos campos.

Detener con `Ctrl+C`.

- [ ] **Step 3: Commit**

```bash
git add templates/index.html
git commit -m "feat: add single-page upload/progress/download UI"
```

---

## Task 9: Utilidad de gestión de usuarios (`setup_usuario.py`)

**Files:**
- Create: `setup_usuario.py`

- [ ] **Step 1: Crear `setup_usuario.py`**

```python
#!/usr/bin/env python3
"""
CLI para administrar usuarios en users.yaml.

Uso:
    python setup_usuario.py add <username>          # añadir o actualizar usuario
    python setup_usuario.py remove <username>       # eliminar usuario
    python setup_usuario.py list                    # listar usuarios
"""
import sys
import getpass
import bcrypt
import yaml
from pathlib import Path

USERS_FILE = Path(__file__).parent / "users.yaml"


def load_users():
    if not USERS_FILE.exists():
        return []
    data = yaml.safe_load(USERS_FILE.read_text(encoding="utf-8")) or {}
    return data.get("users", [])


def save_users(users):
    USERS_FILE.write_text(
        yaml.dump({"users": users}, allow_unicode=True, default_flow_style=False),
        encoding="utf-8",
    )


def cmd_add(username):
    password = getpass.getpass(f"Contraseña para '{username}': ")
    if len(password) < 8:
        print("Error: la contraseña debe tener al menos 8 caracteres.")
        sys.exit(1)
    confirm = getpass.getpass("Confirmar contraseña: ")
    if password != confirm:
        print("Error: las contraseñas no coinciden.")
        sys.exit(1)

    hashed = bcrypt.hashpw(password.encode(), bcrypt.gensalt(rounds=12)).decode()
    users = load_users()
    existing = next((u for u in users if u["username"] == username), None)
    if existing:
        existing["password_hash"] = hashed
        print(f"✓ Contraseña de '{username}' actualizada.")
    else:
        users.append({"username": username, "password_hash": hashed})
        print(f"✓ Usuario '{username}' añadido.")
    save_users(users)


def cmd_remove(username):
    users = load_users()
    new_users = [u for u in users if u["username"] != username]
    if len(new_users) == len(users):
        print(f"Usuario '{username}' no encontrado.")
        sys.exit(1)
    save_users(new_users)
    print(f"✓ Usuario '{username}' eliminado.")


def cmd_list():
    users = load_users()
    if not users:
        print("No hay usuarios registrados.")
        return
    print(f"{'Usuario':<20} {'Hash (primeros 20 chars)'}")
    print("-" * 50)
    for u in users:
        print(f"{u['username']:<20} {u.get('password_hash','')[:20]}...")


def main():
    if len(sys.argv) < 2 or sys.argv[1] not in ("add", "remove", "list"):
        print(__doc__)
        sys.exit(1)

    cmd = sys.argv[1]
    if cmd == "list":
        cmd_list()
    elif cmd in ("add", "remove"):
        if len(sys.argv) < 3:
            print(f"Uso: python setup_usuario.py {cmd} <username>")
            sys.exit(1)
        if cmd == "add":
            cmd_add(sys.argv[2])
        else:
            cmd_remove(sys.argv[2])


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Probar la utilidad**

```bash
python setup_usuario.py add testuser
# Ingresar contraseña cuando lo pida
python setup_usuario.py list
```

Expected: usuario aparece en la lista con hash bcrypt.

- [ ] **Step 3: Commit**

```bash
git add setup_usuario.py
git commit -m "feat: add user management CLI for users.yaml"
```

---

## Task 10: Archivos de despliegue (`deploy/`)

**Files:**
- Create: `deploy/sri-conciliacion.service`
- Create: `deploy/nginx.conf`
- Create: `deploy/install.sh`

- [ ] **Step 1: Crear `deploy/sri-conciliacion.service`**

```ini
[Unit]
Description=SRI Conciliacion Web App
After=network.target

[Service]
WorkingDirectory=/opt/sri_conciliacion
ExecStart=/opt/sri_conciliacion/venv/bin/gunicorn app:app \
    -w 2 \
    -k uvicorn.workers.UvicornWorker \
    --bind 127.0.0.1:8000 \
    --timeout 30 \
    --access-logfile /var/log/sri_conciliacion/access.log \
    --error-logfile /var/log/sri_conciliacion/error.log
Restart=on-failure
RestartSec=5
User=sri
Group=sri

[Install]
WantedBy=multi-user.target
```

- [ ] **Step 2: Crear `deploy/nginx.conf`**

```nginx
# Reemplazar TU_DOMINIO con el dominio real (ej: conciliacion.tuempresa.com)
# Después de modificar: sudo nginx -t && sudo systemctl reload nginx

server {
    listen 80;
    server_name TU_DOMINIO;
    return 301 https://$host$request_uri;
}

server {
    listen 443 ssl;
    server_name TU_DOMINIO;

    ssl_certificate     /etc/letsencrypt/live/TU_DOMINIO/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/TU_DOMINIO/privkey.pem;
    ssl_protocols       TLSv1.2 TLSv1.3;
    ssl_ciphers         HIGH:!aNULL:!MD5;

    # Archivos subidos pueden ser hasta ~20 MB
    client_max_body_size 20M;

    location / {
        proxy_pass         http://127.0.0.1:8000;
        proxy_set_header   Host $host;
        proxy_set_header   X-Real-IP $remote_addr;
        proxy_read_timeout 30s;
    }
}
```

- [ ] **Step 3: Crear `deploy/install.sh`**

```bash
#!/bin/bash
# install.sh — instala la app en un VPS Ubuntu 22.04 limpio
# Uso: sudo bash install.sh TU_DOMINIO
# Requisito: el DNS del dominio ya debe apuntar a este servidor

set -e
DOMAIN="${1:?Uso: sudo bash install.sh TU_DOMINIO}"
APP_DIR="/opt/sri_conciliacion"
APP_USER="sri"

echo "=== [1/7] Actualizando sistema ==="
apt-get update -q && apt-get upgrade -y -q

echo "=== [2/7] Instalando dependencias del sistema ==="
apt-get install -y -q python3 python3-pip python3-venv nginx certbot python3-certbot-nginx git

echo "=== [3/7] Creando usuario de sistema '$APP_USER' ==="
id -u "$APP_USER" &>/dev/null || useradd --system --create-home --shell /bin/bash "$APP_USER"

echo "=== [4/7] Clonando / copiando código ==="
mkdir -p "$APP_DIR"
# Si hay repositorio git:
# git clone <repo_url> "$APP_DIR"
# Si se copia manualmente (scp), este paso ya está hecho.
chown -R "$APP_USER":"$APP_USER" "$APP_DIR"

echo "=== [5/7] Instalando dependencias Python ==="
python3 -m venv "$APP_DIR/venv"
"$APP_DIR/venv/bin/pip" install --quiet --upgrade pip
"$APP_DIR/venv/bin/pip" install --quiet -r "$APP_DIR/requirements.txt"

echo "=== [6/7] Configurando nginx y TLS ==="
mkdir -p /var/log/sri_conciliacion
chown "$APP_USER":"$APP_USER" /var/log/sri_conciliacion
sed "s/TU_DOMINIO/$DOMAIN/g" "$APP_DIR/deploy/nginx.conf" > /etc/nginx/sites-available/sri_conciliacion
ln -sf /etc/nginx/sites-available/sri_conciliacion /etc/nginx/sites-enabled/
nginx -t && systemctl reload nginx
certbot --nginx -d "$DOMAIN" --non-interactive --agree-tos -m "admin@$DOMAIN"

echo "=== [7/7] Configurando servicio systemd ==="
cp "$APP_DIR/deploy/sri-conciliacion.service" /etc/systemd/system/
systemctl daemon-reload
systemctl enable sri-conciliacion
systemctl start sri-conciliacion

echo ""
echo "✅ Instalación completa."
echo "   App corriendo en: https://$DOMAIN"
echo ""
echo "Próximo paso — añadir el primer usuario:"
echo "   sudo -u $APP_USER $APP_DIR/venv/bin/python $APP_DIR/setup_usuario.py add <nombre>"
```

- [ ] **Step 4: Dar permisos de ejecución**

```bash
chmod +x deploy/install.sh
```

- [ ] **Step 5: Commit final**

```bash
git add deploy/ requirements.txt
git commit -m "feat: add VPS deployment files (systemd, nginx, install script)"
```

---

## Task 11: Test suite completa y verificación final

- [ ] **Step 1: Correr todos los tests**

```bash
cd C:/Users/User/Documents/2026/sri_conciliacion
python -m pytest tests/ -v --tb=short
```

Expected: todos los tests PASSED (sin errores ni warnings relevantes).

- [ ] **Step 2: Smoke test local del servidor**

```bash
# Añadir un usuario de prueba si no existe
python setup_usuario.py add demo

# Arrancar servidor
uvicorn app:app --port 8000
```

En otra terminal o en el navegador:
- Abrir `http://localhost:8000` → debe pedir credenciales
- Ingresar usuario `demo` con la contraseña creada → debe mostrar el formulario
- Subir `2390000831001_Recibidos.txt` + `reten_odoo.xlsx`
- Verificar que el progreso se actualiza cada ~5 segundos
- Al terminar, descargar el Excel y abrirlo

- [ ] **Step 3: Commit de cierre**

```bash
git add .
git commit -m "chore: complete web app implementation - all tests passing"
```

---

## Resumen de archivos nuevos/modificados

| Archivo | Estado |
|---|---|
| `requirements.txt` | Nuevo |
| `build_workbook_final.py` | Modificado (función `build_workbook()`) |
| `conciliar_odoo.py` | Modificado (función `conciliar()`) |
| `auth.py` | Nuevo |
| `job_manager.py` | Nuevo |
| `pipeline.py` | Nuevo |
| `app.py` | Nuevo |
| `templates/index.html` | Nuevo |
| `setup_usuario.py` | Nuevo |
| `.gitignore` | Nuevo |
| `deploy/sri-conciliacion.service` | Nuevo |
| `deploy/nginx.conf` | Nuevo |
| `deploy/install.sh` | Nuevo |
| `tests/test_auth.py` | Nuevo |
| `tests/test_job_manager.py` | Nuevo |
| `tests/test_app.py` | Nuevo |
| `tests/test_build_workbook.py` | Nuevo |
| `tests/test_conciliar.py` | Nuevo |
