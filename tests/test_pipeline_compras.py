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


def _sri(*fechas):
    return pd.DataFrame({"fecha": pd.to_datetime(list(fechas))})


def test_periodo_mes_dominante_con_atipico_y_ventana_amplia(tmp_path, monkeypatch):
    pc, _ = _job(tmp_path, monkeypatch)
    from datetime import date
    sri = _sri("2026-07-20", "2026-09-05", "2026-09-06", "2026-09-30")
    assert pc.periodo_sri(sri) == (date(2026, 9, 1), date(2026, 9, 30),
                                   date(2026, 7, 20), date(2026, 9, 30))


def test_periodo_empate_gana_el_mes_mas_reciente(tmp_path, monkeypatch):
    pc, _ = _job(tmp_path, monkeypatch)
    from datetime import date
    r = pc.periodo_sri(_sri("2026-08-10", "2026-09-05"))
    assert r[:2] == (date(2026, 9, 1), date(2026, 9, 30))
    assert r[2:] == (date(2026, 8, 10), date(2026, 9, 30))


def test_periodo_varios_meses_domina_uno(tmp_path, monkeypatch):
    pc, _ = _job(tmp_path, monkeypatch)
    from datetime import date
    sri = _sri("2026-06-01", "2026-08-02", "2026-08-03", "2026-08-04", "2026-08-31", "2026-09-01")
    assert pc.periodo_sri(sri)[:2] == (date(2026, 8, 1), date(2026, 8, 31))
    assert pc.periodo_sri(sri)[2:] == (date(2026, 6, 1), date(2026, 9, 1))


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


def test_pipeline_ventana_amplia_para_odoo_y_mes_dominante_para_conciliar(tmp_path, monkeypatch):
    from datetime import date
    pc, job_id = _job(tmp_path, monkeypatch)
    monkeypatch.setattr(pc, "leer_sri_txts", lambda rutas: _sri("2026-07-20", "2026-09-05", "2026-09-06"))
    monkeypatch.setattr(pc.OdooReader, "from_env", classmethod(lambda cls: _Lector()))
    visto = {}
    monkeypatch.setattr(pc, "conciliar_compras",
                        lambda r, o, d, h, s: visto.update(d=d, h=h) or {"por_tipo": {}, "advertencias": []})
    pc.run_pipeline_compras(job_id)
    assert read_status(job_id)["estado"] == "done"
    assert _Lector.llamado == (date(2026, 7, 20), date(2026, 9, 30))
    assert (visto["d"], visto["h"]) == (date(2026, 9, 1), date(2026, 9, 30))


def test_periodo_diciembre_con_atipico_de_enero(tmp_path, monkeypatch):
    from datetime import date
    pc, _ = _job(tmp_path, monkeypatch)
    sri = _sri("2026-12-05", "2026-12-06", "2026-12-20", "2027-01-03")
    assert pc.periodo_sri(sri) == (date(2026, 12, 1), date(2026, 12, 31),
                                   date(2026, 12, 1), date(2027, 1, 3))


def test_pipeline_con_error_agenda_limpieza(tmp_path, monkeypatch):
    pc, job_id = _job(tmp_path, monkeypatch)
    llamadas = []
    monkeypatch.setattr(pc, "schedule_cleanup", lambda *a, **k: llamadas.append((a, k)))
    monkeypatch.setattr(pc, "leer_sri_txts", lambda rutas: pd.DataFrame({"fecha": []}))
    pc.run_pipeline_compras(job_id)
    assert read_status(job_id)["estado"] == "error"
    assert llamadas == [((job_id,), {"delay_seconds": 7200})]
