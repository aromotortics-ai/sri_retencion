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
