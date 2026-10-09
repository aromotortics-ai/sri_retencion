import json

from job_manager import create_job, get_job_dir, read_status


def test_pipeline_ventas_ok(tmp_path, monkeypatch):
    monkeypatch.setattr("job_manager.JOBS_DIR", tmp_path / "jobs")
    import pipeline_ventas
    job_id = create_job()
    (get_job_dir(job_id) / "emitidos").mkdir()
    monkeypatch.setattr(pipeline_ventas, "conciliar_ventas",
                        lambda odoo, emitidos, salida: {"coinciden": 3})
    monkeypatch.setattr(pipeline_ventas, "schedule_cleanup", lambda *a, **k: None)
    pipeline_ventas.run_pipeline_ventas(job_id)
    s = read_status(job_id)
    assert s["estado"] == "done"
    assert s["resumen"] == {"coinciden": 3}


def test_pipeline_ventas_error_se_reporta(tmp_path, monkeypatch):
    monkeypatch.setattr("job_manager.JOBS_DIR", tmp_path / "jobs")
    import pipeline_ventas
    job_id = create_job()

    def boom(*a):
        raise ValueError("formato incorrecto")
    monkeypatch.setattr(pipeline_ventas, "conciliar_ventas", boom)
    pipeline_ventas.run_pipeline_ventas(job_id)
    s = read_status(job_id)
    assert s["estado"] == "error"
    assert "formato incorrecto" in s["mensaje"]
