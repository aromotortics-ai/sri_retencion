"""
Runner del pipeline de conciliación de compras (comprobantes recibidos SRI vs Odoo por API).
Se ejecuta en un hilo background lanzado por app.py.
"""
import calendar
import traceback
from datetime import date

from conciliar_compras import conciliar_compras, leer_sri_txts
from job_manager import write_status, schedule_cleanup, get_job_dir
from odoo_client import OdooReader


def periodo_sri(sri) -> tuple[date, date]:
    """Primer día del mes de la fecha más antigua y último día del mes de la más reciente."""
    menor, mayor = sri["fecha"].min(), sri["fecha"].max()
    return (date(menor.year, menor.month, 1),
            date(mayor.year, mayor.month, calendar.monthrange(mayor.year, mayor.month)[1]))


def run_pipeline_compras(job_id: str) -> None:
    job_dir = get_job_dir(job_id)
    try:
        write_status(job_id, "running", "Leyendo TXT del SRI...", 0.1)
        rutas = sorted(str(p) for p in (job_dir / "sri").glob("*.txt"))
        sri = leer_sri_txts(rutas)
        if sri.empty:
            raise ValueError("Los TXT no contienen facturas, notas de crédito ni notas de débito")
        desde, hasta = periodo_sri(sri)

        write_status(job_id, "running", "Leyendo compras de Odoo...", 0.4)
        odoo = OdooReader.from_env().fetch_compras(desde, hasta)

        write_status(job_id, "running", "Conciliando SRI vs Odoo...", 0.8)
        resumen = conciliar_compras(rutas, odoo, desde, hasta, str(job_dir / "resultado.xlsx"))
        write_status(job_id, "done", "Completado", 1.0, resumen=resumen)
        schedule_cleanup(job_id, delay_seconds=7200)
    except Exception as exc:
        write_status(job_id, "error", f"{type(exc).__name__}: {exc}",
                     error=traceback.format_exc())
