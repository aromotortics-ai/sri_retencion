"""
Runner del pipeline de conciliación de compras (comprobantes recibidos SRI vs Odoo por API).
Incluye las retenciones recibidas cuando se carga su TXT. `procesar_compras` es el cuerpo común
(lo usa la sesión de trabajo, que lo corre en un hilo background).
"""
import calendar
from datetime import date

import pandas as pd

from app.compras.clasificar import clasificar_compras
from app.compras.conciliar_compras import conciliar_y_excel, leer_sri_txts
from app.retenciones.conciliar_retenciones import (conciliar_retenciones, leer_sri_retenciones_txt,
                                                   resumir_por_documento)
from app.odoo.client import COLUMNAS as COLUMNAS_ODOO_COMPRAS, OdooReader
from app.reportes.tablero import construir_tablero
from app.sri.xml_cache import descargar_xmls, lineas_sri, resumir_motivos

from app.rutas import CACHE_XML_DIR  # caché de XML compartido entre jobs

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


def _sin_cancelar():
    """Punto de control nulo (sin sesión: nada puede superar la corrida)."""


def _retenciones(sri_ret, lector, avisar, cancelado):
    """Lee Odoo, descarga XML y concilia retenciones. Devuelve (ret, desde, hasta, n_sin_xml, fallidos)."""
    con_fecha = sri_ret[sri_ret["fecha"].notna()]
    if con_fecha.empty:
        raise ValueError("El TXT de retenciones no tiene fechas de emisión válidas")
    desde, hasta, fetch_desde, fetch_hasta = periodo_sri(con_fecha)

    # Odoo primero: credenciales erróneas u Odoo caído fallan de inmediato, no tras la descarga.
    avisar("Leyendo retenciones de Odoo...", 0.12)
    odoo_ret = lector.fetch_retenciones(fetch_desde, fetch_hasta)
    cancelado()

    claves = list(sri_ret["clave"])
    fallidos = descargar_con_progreso(claves, descargar_xmls, CACHE_XML_DIR, avisar)
    cancelado()

    lineas = lineas_sri(claves, CACHE_XML_DIR)
    # Sin líneas = no descargado o XML ilegible/vacío: todos terminan como "XML no disponible".
    n_sin_xml = len(set(claves) - set(lineas["clave"]))

    avisar("Conciliando retenciones...", 0.7)
    ret = conciliar_retenciones(lineas, odoo_ret, sri_ret, desde, hasta)
    return ret, desde, hasta, n_sin_xml, fallidos


def procesar_compras(rutas, salida, avisar, cancelado=None, advertencias=None):
    """Concilia los TXT `rutas` contra Odoo, escribe el Excel completo en `salida` y devuelve
    (resumen, tablero).

    avisar(mensaje, progreso): informa el avance. cancelado(): punto de control que lanza una
    excepción si la corrida fue superada (se llama tras leer los TXT, tras cada lectura de Odoo,
    tras la descarga de XML y antes de escribir). advertencias: se anteponen a las del resultado
    (p. ej. TXT excluidos por ilegibles)."""
    cancelado = cancelado or _sin_cancelar
    avisar("Leyendo TXT del SRI...", 0.1)
    sri = leer_sri_txts(rutas)
    sri_ret = leer_sri_retenciones_txt(rutas)
    if sri.empty and sri_ret.empty:
        raise ValueError("Los TXT no contienen facturas, notas de crédito, notas de débito "
                         "ni retenciones")
    cancelado()
    lector = OdooReader.from_env()  # una sola conexión para compras y retenciones

    ret, n_sin_xml, periodo_ret, fallidos = None, 0, None, []
    if not sri_ret.empty:
        ret, ret_desde, ret_hasta, n_sin_xml, fallidos = _retenciones(sri_ret, lector, avisar, cancelado)
        periodo_ret = (ret_desde, ret_hasta)
        desde, hasta = periodo_ret  # solo retenciones: hojas de compras con este período

    if not sri.empty:
        desde, hasta, fetch_desde, fetch_hasta = periodo_sri(sri)
        avisar("Leyendo compras de Odoo...", 0.75)
        odoo = lector.fetch_compras(fetch_desde, fetch_hasta)
        cancelado()
    else:
        # Solo retenciones: hojas de compras vacías con el período de las retenciones.
        odoo = pd.DataFrame(columns=COLUMNAS_ODOO_COMPRAS)

    avisar("Conciliando SRI vs Odoo...", 0.85)
    cancelado()
    resumen, pares = conciliar_y_excel(sri, odoo, desde, hasta, salida, ret=ret, periodo_ret=periodo_ret)
    resumen["advertencias"][:0] = list(advertencias or [])
    if ret is not None:
        resumen["retenciones_fallidos"] = n_sin_xml
        if n_sin_xml:
            resumen["advertencias"].append(
                f"{n_sin_xml} XML de retenciones no se pudieron descargar o leer; sus retenciones "
                "se listan como 'XML no disponible'")
            if fallidos:
                resumen["advertencias"].append(
                    f"Motivo de las descargas fallidas: {resumir_motivos(fallidos)}")
    # Mismos DataFrames (sin reindexar) que el Excel: pantalla y Excel salen del mismo cruce.
    docs = clasificar_compras(sri, odoo, pares)
    retdocs = resumir_por_documento(ret) if ret is not None else None
    tablero = construir_tablero("compras", docs, sri, odoo, retdocs, (desde, hasta),
                                resumen["advertencias"])
    return resumen, tablero
