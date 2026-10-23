"""
Runner del pipeline de conciliación de compras (comprobantes recibidos SRI vs Odoo por API).
Incluye las retenciones recibidas cuando se carga su TXT. Se ejecuta en un hilo background
lanzado por app.py.
"""
import calendar
import traceback
from datetime import date
from pathlib import Path

import pandas as pd

from conciliar_compras import conciliar_compras, leer_sri_txts
from conciliar_retenciones import conciliar_retenciones, leer_sri_retenciones_txt
from job_manager import write_status, schedule_cleanup, get_job_dir
from odoo_client import COLUMNAS as COLUMNAS_ODOO_COMPRAS, OdooReader
from xml_cache import descargar_xmls, lineas_sri

# Caché de XML compartido entre jobs (misma ruta que pipeline.CACHE_XML_DIR).
CACHE_XML_DIR = Path(__file__).parent / "cache" / "xml"

_DESCARGA_INI, _DESCARGA_FIN = 0.15, 0.65


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


def descargar_con_progreso(claves, descargar, cache_dir, avisar):
    """Descarga los XML de `claves` con `descargar` (firma de xml_cache.descargar_xmls) informando el
    avance entre 0,15 y 0,65 con avisar(mensaje, progreso). Devuelve los fallidos.

    La función de descarga, el caché y el aviso se reciben como argumentos para que cada pipeline
    use (y los tests sustituyan) los suyos."""
    total = len(claves)
    avance = {"hecho": False}

    def mensaje(n, k):
        return f"Descargando XML de retenciones ({n}/{total}, {k} fallidos)"

    def on_progress(n, tot, k):
        avance["hecho"] = True
        fr = n / tot if tot else 1.0
        avisar(mensaje(n, k), _DESCARGA_INI + (_DESCARGA_FIN - _DESCARGA_INI) * fr)

    avisar(mensaje(0, 0), _DESCARGA_INI)
    fallidos = descargar(claves, cache_dir, on_progress=on_progress)
    if not avance["hecho"]:  # todo estaba en caché: on_progress nunca se llamó
        avisar(mensaje(total, len(fallidos)), _DESCARGA_FIN)
    return fallidos


def _retenciones(job_id, sri_ret, lector):
    """Lee Odoo, descarga XML y concilia retenciones. Devuelve (ret, desde, hasta, n_sin_xml)."""
    con_fecha = sri_ret[sri_ret["fecha"].notna()]
    if con_fecha.empty:
        raise ValueError("El TXT de retenciones no tiene fechas de emisión válidas")
    desde, hasta, fetch_desde, fetch_hasta = periodo_sri(con_fecha)

    # Odoo primero: credenciales erróneas u Odoo caído fallan de inmediato, no tras la descarga.
    write_status(job_id, "running", "Leyendo retenciones de Odoo...", 0.12)
    odoo_ret = lector.fetch_retenciones(fetch_desde, fetch_hasta)

    claves = list(sri_ret["clave"])
    descargar_con_progreso(claves, descargar_xmls, CACHE_XML_DIR,
                           lambda m, p: write_status(job_id, "running", m, p))

    lineas = lineas_sri(claves, CACHE_XML_DIR)
    # Sin líneas = no descargado o XML ilegible/vacío: todos terminan como "XML no disponible".
    n_sin_xml = len(set(claves) - set(lineas["clave"]))

    write_status(job_id, "running", "Conciliando retenciones...", 0.7)
    ret = conciliar_retenciones(lineas, odoo_ret, sri_ret, desde, hasta)
    return ret, desde, hasta, n_sin_xml


def run_pipeline_compras(job_id: str) -> None:
    job_dir = get_job_dir(job_id)
    try:
        write_status(job_id, "running", "Leyendo TXT del SRI...", 0.1)
        rutas = sorted(str(p) for p in (job_dir / "sri").glob("*.txt"))
        sri = leer_sri_txts(rutas)
        sri_ret = leer_sri_retenciones_txt(rutas)
        if sri.empty and sri_ret.empty:
            raise ValueError("Los TXT no contienen facturas, notas de crédito, notas de débito "
                             "ni retenciones")
        lector = OdooReader.from_env()  # una sola conexión para compras y retenciones

        ret, n_sin_xml, periodo_ret = None, 0, None
        if not sri_ret.empty:
            ret, ret_desde, ret_hasta, n_sin_xml = _retenciones(job_id, sri_ret, lector)
            periodo_ret = (ret_desde, ret_hasta)
            desde, hasta = periodo_ret  # solo retenciones: hojas de compras con este período

        if not sri.empty:
            desde, hasta, fetch_desde, fetch_hasta = periodo_sri(sri)
            write_status(job_id, "running", "Leyendo compras de Odoo...", 0.75)
            odoo = lector.fetch_compras(fetch_desde, fetch_hasta)
        else:
            # Solo retenciones: hojas de compras vacías con el período de las retenciones.
            odoo = pd.DataFrame(columns=COLUMNAS_ODOO_COMPRAS)

        write_status(job_id, "running", "Conciliando SRI vs Odoo...", 0.85)
        resumen = conciliar_compras(rutas, odoo, desde, hasta, str(job_dir / "resultado.xlsx"),
                                    ret=ret, periodo_ret=periodo_ret)
        if ret is not None:
            resumen["retenciones_fallidos"] = n_sin_xml
            if n_sin_xml:
                resumen["advertencias"].append(
                    f"{n_sin_xml} XML de retenciones no se pudieron descargar o leer; sus retenciones "
                    "se listan como 'XML no disponible'")
        write_status(job_id, "done", "Completado", 1.0, resumen=resumen)
        schedule_cleanup(job_id, delay_seconds=7200)
    except Exception as exc:
        write_status(job_id, "error", f"{type(exc).__name__}: {exc}",
                     error=traceback.format_exc())
        schedule_cleanup(job_id, delay_seconds=7200)
