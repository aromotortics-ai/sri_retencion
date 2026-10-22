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


def periodo_sri(sri) -> tuple[date, date, date, date]:
    """(desde, hasta, fetch_desde, fetch_hasta).

    desde/hasta: mes con más filas del SRI (empate: el más reciente); es el período a reportar.
    fetch_*: rango min-max de fechas del SRI, nunca más estrecho que el período, para que los
    documentos atípicos encuentren su par en Odoo.
    """
    f = sri["fecha"]
    cuenta = f.dt.to_period("M").value_counts()
    mes = max(cuenta[cuenta == cuenta.max()].index)
    desde = date(mes.year, mes.month, 1)
    hasta = date(mes.year, mes.month, calendar.monthrange(mes.year, mes.month)[1])
    return (desde, hasta,
            min(desde, f.min().date()), max(hasta, f.max().date()))


def run_pipeline_compras(job_id: str) -> None:
    job_dir = get_job_dir(job_id)
    try:
        write_status(job_id, "running", "Leyendo TXT del SRI...", 0.1)
        rutas = sorted(str(p) for p in (job_dir / "sri").glob("*.txt"))
        sri = leer_sri_txts(rutas)
        if sri.empty:
            raise ValueError("Los TXT no contienen facturas, notas de crédito ni notas de débito")
        desde, hasta, fetch_desde, fetch_hasta = periodo_sri(sri)

        write_status(job_id, "running", "Leyendo compras de Odoo...", 0.4)
        odoo = OdooReader.from_env().fetch_compras(fetch_desde, fetch_hasta)

        write_status(job_id, "running", "Conciliando SRI vs Odoo...", 0.8)
        resumen = conciliar_compras(rutas, odoo, desde, hasta, str(job_dir / "resultado.xlsx"))
        write_status(job_id, "done", "Completado", 1.0, resumen=resumen)
        schedule_cleanup(job_id, delay_seconds=7200)
    except Exception as exc:
        write_status(job_id, "error", f"{type(exc).__name__}: {exc}",
                     error=traceback.format_exc())
        schedule_cleanup(job_id, delay_seconds=7200)
