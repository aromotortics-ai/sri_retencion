import pandas as pd

from job_manager import create_job, get_job_dir, read_status
from odoo_client import OdooError


def _job(tmp_path, monkeypatch):
    monkeypatch.setattr("job_manager.JOBS_DIR", tmp_path / "jobs")
    import pipeline_compras
    job_id = create_job()
    (get_job_dir(job_id) / "sri").mkdir()
    monkeypatch.setattr(pipeline_compras, "schedule_cleanup", lambda *a, **k: None)
    return pipeline_compras, job_id


class _Lector:
    llamado = None

    def fetch_compras(self, desde, hasta, margen_dias=5):
        _Lector.llamado = (desde, hasta)
        return pd.DataFrame()


def test_pipeline_compras_ok(tmp_path, monkeypatch):
    pc, job_id = _job(tmp_path, monkeypatch)
    sri = pd.DataFrame({"fecha": pd.to_datetime(["2026-09-05", "2026-09-20"])})
    monkeypatch.setattr(pc, "leer_sri_txts", lambda rutas: sri)
    monkeypatch.setattr(pc.OdooReader, "from_env", classmethod(lambda cls: _Lector()))
    visto = {}

    def falso(rutas, odoo, desde, hasta, salida):
        visto.update(rutas=rutas, desde=desde, hasta=hasta, salida=salida)
        return {"por_tipo": {"FAC": {"coinciden": 1}}, "advertencias": []}
    monkeypatch.setattr(pc, "conciliar_compras", falso)
    (get_job_dir(job_id) / "sri" / "a.txt").write_text("x")
    pc.run_pipeline_compras(job_id)
    s = read_status(job_id)
    assert s["estado"] == "done"
    assert s["resumen"]["por_tipo"]["FAC"]["coinciden"] == 1
    from datetime import date
    assert _Lector.llamado == (date(2026, 9, 1), date(2026, 9, 30))
    assert (visto["desde"], visto["hasta"]) == (date(2026, 9, 1), date(2026, 9, 30))
    assert visto["salida"].endswith("resultado.xlsx")


def test_periodo_abarca_meses_de_las_fechas(tmp_path, monkeypatch):
    pc, _ = _job(tmp_path, monkeypatch)
    from datetime import date
    sri = pd.DataFrame({"fecha": pd.to_datetime(["2026-02-10", "2026-03-31"])})
    assert pc.periodo_sri(sri) == (date(2026, 2, 1), date(2026, 3, 31))


def test_pipeline_compras_sin_credenciales_error_claro(tmp_path, monkeypatch):
    pc, job_id = _job(tmp_path, monkeypatch)
    sri = pd.DataFrame({"fecha": pd.to_datetime(["2026-09-05"])})
    monkeypatch.setattr(pc, "leer_sri_txts", lambda rutas: sri)

    def sin_env(cls):
        raise OdooError("Faltan variables de entorno de Odoo: ODOO_URL")
    monkeypatch.setattr(pc.OdooReader, "from_env", classmethod(sin_env))
    pc.run_pipeline_compras(job_id)
    s = read_status(job_id)
    assert s["estado"] == "error"
    assert "ODOO_URL" in s["mensaje"]


def test_pipeline_compras_sin_filas_sri_es_error(tmp_path, monkeypatch):
    pc, job_id = _job(tmp_path, monkeypatch)
    monkeypatch.setattr(pc, "leer_sri_txts", lambda rutas: pd.DataFrame({"fecha": []}))
    pc.run_pipeline_compras(job_id)
    s = read_status(job_id)
    assert s["estado"] == "error"
    assert "factura" in s["mensaje"].lower()
