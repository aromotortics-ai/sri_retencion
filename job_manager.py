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
