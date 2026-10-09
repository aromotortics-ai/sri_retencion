"""
Runner del pipeline de conciliación de facturas de venta (Odoo vs SRI emitidos).
Se ejecuta en un hilo background lanzado por app.py.
"""
import traceback

from conciliar_ventas_sri import conciliar_ventas
from job_manager import write_status, schedule_cleanup, get_job_dir


def run_pipeline_ventas(job_id: str) -> None:
    job_dir = get_job_dir(job_id)
    try:
        write_status(job_id, "running", "Conciliando ventas Odoo vs SRI...", 0.3)
        resumen = conciliar_ventas(
            str(job_dir / "odoo_venta.xlsx"),
            str(job_dir / "emitidos"),
            str(job_dir / "resultado.xlsx"),
        )
        write_status(job_id, "done", "Completado", 1.0, resumen=resumen)
        schedule_cleanup(job_id, delay_seconds=7200)
    except Exception as exc:
        write_status(job_id, "error", f"{type(exc).__name__}: {exc}",
                     error=traceback.format_exc())
