import pytest
import json
import time
from pathlib import Path
from unittest.mock import patch


def test_create_job_returns_id(tmp_path):
    from app.job_manager import create_job
    with patch("app.job_manager.JOBS_DIR", tmp_path):
        job_id = create_job()
    assert isinstance(job_id, str)
    assert len(job_id) == 8
    assert (tmp_path / job_id).is_dir()


def test_write_and_read_status(tmp_path):
    from app.job_manager import create_job, write_status, read_status
    with patch("app.job_manager.JOBS_DIR", tmp_path):
        job_id = create_job()
        write_status(job_id, "running", "Procesando...", 0.5)
        status = read_status(job_id)

    assert status["estado"] == "running"
    assert status["mensaje"] == "Procesando..."
    assert status["progreso"] == pytest.approx(0.5)


def test_read_status_missing_job(tmp_path):
    from app.job_manager import read_status
    with patch("app.job_manager.JOBS_DIR", tmp_path):
        result = read_status("noexiste")
    assert result is None


def test_write_status_done_with_resumen(tmp_path):
    from app.job_manager import create_job, write_status, read_status
    resumen = {"coinciden": 620, "solo_sri": 317}
    with patch("app.job_manager.JOBS_DIR", tmp_path):
        job_id = create_job()
        write_status(job_id, "done", "Completado", 1.0, resumen=resumen)
        status = read_status(job_id)

    assert status["estado"] == "done"
    assert status["resumen"]["coinciden"] == 620


def test_get_job_dir(tmp_path):
    from app.job_manager import create_job, get_job_dir
    with patch("app.job_manager.JOBS_DIR", tmp_path):
        job_id = create_job()
        d = get_job_dir(job_id)
    assert d == tmp_path / job_id


def _esperar(cond, segundos=5.0):
    limite = time.monotonic() + segundos
    while time.monotonic() < limite:
        if cond():
            return True
        time.sleep(0.02)
    return cond()


def test_cleanup_respeta_ultima_actividad_reciente(tmp_path):
    from app.job_manager import create_job, get_job_dir, schedule_cleanup
    with patch("app.job_manager.JOBS_DIR", tmp_path):
        job_id = create_job()
        (get_job_dir(job_id) / "sesion.json").write_text(
            json.dumps({"ultima_actividad": time.time() + 60}), encoding="utf-8")
        schedule_cleanup(job_id, delay_seconds=0.05)
        time.sleep(0.3)
        assert get_job_dir(job_id).exists()  # hubo actividad después: un timer posterior se encarga


def test_cleanup_borra_si_la_actividad_es_vieja(tmp_path):
    from app.job_manager import create_job, get_job_dir, schedule_cleanup
    with patch("app.job_manager.JOBS_DIR", tmp_path):
        job_id = create_job()
        (get_job_dir(job_id) / "sesion.json").write_text(
            json.dumps({"ultima_actividad": time.time() - 10}), encoding="utf-8")
        schedule_cleanup(job_id, delay_seconds=0.05)
        assert _esperar(lambda: not get_job_dir(job_id).exists())
