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
