# -*- coding: utf-8 -*-
"""
Runner del pipeline de conciliación de ventas: comprobantes EMITIDOS (facturas, NC, ND, liquidaciones
de compra y retenciones) del SRI vs Odoo por API. `procesar_ventas` lo usa la sesión de trabajo (app/sesion.py) en un hilo background.

Orden (falla rápido): TXT -> período -> Odoo (documentos y retenciones) -> XML de retenciones (solo si
hay) -> clasificación -> Excel.
"""

import pandas as pd

from app.ventas.conciliar_emitidos import clasificar_documentos, conciliar_retenciones_emitidas
from app.sri.emitidos_sri import TIPOS_DOC, leer_emitidos_txts, periodo_emitidos
from app.ventas.excel_emitidos import construir_excel_emitidos
from app.odoo.client import COLUMNAS_EMITIDOS, OdooReader
from app.compras.pipeline_compras import descargar_con_progreso
from app.reportes.tablero import construir_tablero
from app.retenciones.conciliar_retenciones import resumir_por_documento
from app.sri.xml_cache import descargar_xmls, lineas_sri, resumir_motivos

from app.rutas import CACHE_XML_DIR  # caché de XML compartido entre jobs

ADV_LIBREOFFICE = "No se pudo recalcular con LibreOffice; Excel recalculará las fórmulas al abrir el archivo"


_NOMBRE_TIPO = {"FAC": "facturas", "NC": "notas de crédito", "ND": "notas de débito",
                "LIQ": "liquidaciones"}


def _tipos_sin_txt(docs_sri, odoo, desde, hasta):
    """Advertencias por cada tipo sin filas en los TXT pero con registros de Odoo (publicados o
    anulados) dentro de [desde, hasta]: probablemente falta subir ese TXT."""
    fechas = pd.to_datetime(odoo["fecha"], errors="coerce")
    en_periodo = odoo[odoo["estado"].isin(["posted", "cancel"])
                      & (fechas >= pd.Timestamp(desde)) & (fechas <= pd.Timestamp(hasta))]
    advertencias = []
    for tipo in TIPOS_DOC:
        n = int((en_periodo["tipo"] == tipo).sum())
        if n and not (docs_sri["tipo"] == tipo).any():
            advertencias.append(f"Falta el TXT de {_NOMBRE_TIPO[tipo]}: los {n} registros de Odoo de "
                                "ese tipo aparecerán como 'Solo en Odoo'")
    return advertencias


def _sin_cancelar():
    """Punto de control nulo (sin sesión: nada puede superar la corrida)."""


def _retenciones(ret_txt, odoo_ret, desde, hasta, avisar, cancelado):
    """Descarga los XML (con caché) y concilia las retenciones emitidas. Devuelve (ret, n_sin_xml, fallidos)."""
    claves = list(ret_txt["clave"])
    fallidos = descargar_con_progreso(claves, descargar_xmls, CACHE_XML_DIR, avisar)
    cancelado()
    lineas = lineas_sri(claves, CACHE_XML_DIR, contraparte=True)
    # Sin líneas = no descargado o XML ilegible/vacío: todos terminan como "XML no disponible".
    n_sin_xml = len(set(claves) - set(lineas["clave"]))
    avisar("Conciliando retenciones...", 0.7)
    ret = conciliar_retenciones_emitidas(lineas, odoo_ret, ret_txt, desde, hasta)
    return ret, n_sin_xml, fallidos


def procesar_ventas(rutas, salida, avisar, cancelado=None, advertencias=None):
    """Concilia los TXT de emitidos `rutas` contra Odoo, escribe el Excel completo en `salida` y
    devuelve (resumen, tablero).

    avisar(mensaje, progreso): informa el avance. cancelado(): punto de control que lanza una
    excepción si la corrida fue superada (tras leer los TXT, tras cada lectura de Odoo, tras la
    descarga de XML y antes de escribir). advertencias: se anteponen a las del resultado."""
    cancelado = cancelado or _sin_cancelar
    avisar("Leyendo TXT del SRI...", 0.1)
    docs_sri, ret_txt, adv_txt = leer_emitidos_txts(rutas)
    advertencias = list(advertencias or []) + adv_txt
    if docs_sri.empty and ret_txt.empty:
        raise ValueError("Los TXT no contienen facturas, notas de crédito, notas de débito, "
                         "liquidaciones de compra ni retenciones")
    if not ret_txt.empty and ret_txt["fecha"].notna().sum() == 0:
        raise ValueError("El TXT de retenciones no tiene fechas de emisión válidas")
    desde, hasta, adv_periodo = periodo_emitidos(pd.concat([docs_sri["fecha"], ret_txt["fecha"]]))
    advertencias += adv_periodo
    cancelado()

    # Odoo primero: credenciales erróneas u Odoo caído fallan antes de descargar XML.
    lector = OdooReader.from_env()
    if not docs_sri.empty:
        avisar("Leyendo comprobantes emitidos de Odoo...", 0.12)
        odoo = lector.fetch_emitidos(desde, hasta)
        cancelado()
        advertencias += _tipos_sin_txt(docs_sri, odoo, desde, hasta)
    else:
        # Solo retenciones: sin documentos del SRI no se listan documentos de Odoo.
        odoo = pd.DataFrame(columns=COLUMNAS_EMITIDOS)
    odoo_ret = None
    if not ret_txt.empty:
        avisar("Leyendo retenciones emitidas de Odoo...", 0.13)
        odoo_ret = lector.fetch_retenciones_emitidas(desde, hasta)
        cancelado()

    ret, n_sin_xml = None, 0
    if odoo_ret is not None:
        ret, n_sin_xml, fallidos = _retenciones(ret_txt, odoo_ret, desde, hasta, avisar, cancelado)
        if n_sin_xml:
            advertencias.append(
                f"{n_sin_xml} XML de retenciones no se pudieron descargar o leer; sus retenciones "
                "se listan como 'XML no disponible'")
            if fallidos:
                advertencias.append(f"Motivo de las descargas fallidas: {resumir_motivos(fallidos)}")

    avisar("Conciliando SRI vs Odoo...", 0.85)
    docs = clasificar_documentos(docs_sri, odoo, desde, hasta)
    cancelado()
    res = construir_excel_emitidos(docs_sri, odoo, docs, ret, desde, hasta, salida,
                                   advertencias=list(advertencias))
    recalculado = bool(res["recalculado"])
    if not recalculado:
        advertencias.append(ADV_LIBREOFFICE)

    conteo_sri = docs_sri["tipo"].value_counts()
    comprobantes_sri = {t: int(conteo_sri.get(t, 0)) for t in TIPOS_DOC}
    comprobantes_sri["RET"] = int(len(ret_txt))
    resumen = {
        "por_tipo": res["documentos"],
        "periodo": [desde.isoformat(), hasta.isoformat()],
        "comprobantes_sri": comprobantes_sri,
        "advertencias": [str(a) for a in advertencias],
        "retenciones_fallidos": int(n_sin_xml),
        "recalculado": recalculado,
    }
    retdocs = resumir_por_documento(ret) if ret is not None else None
    tablero = construir_tablero("ventas", docs, docs_sri, odoo, retdocs, (desde, hasta),
                                resumen["advertencias"])
    return resumen, tablero
