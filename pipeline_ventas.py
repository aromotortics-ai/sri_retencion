# -*- coding: utf-8 -*-
"""
Runner del pipeline de conciliación de ventas: comprobantes EMITIDOS (facturas, NC, ND, liquidaciones
de compra y retenciones) del SRI vs Odoo por API. Se ejecuta en un hilo background lanzado por app.py.

Orden (falla rápido): TXT -> período -> Odoo (documentos y retenciones) -> XML de retenciones (solo si
hay) -> clasificación -> Excel.
"""
import traceback
from pathlib import Path

import pandas as pd

from conciliar_emitidos import clasificar_documentos, conciliar_retenciones_emitidas
from emitidos_sri import TIPOS_DOC, leer_emitidos_txts, periodo_emitidos
from excel_emitidos import construir_excel_emitidos
from job_manager import write_status, schedule_cleanup, get_job_dir
from odoo_client import COLUMNAS_EMITIDOS, OdooReader
from pipeline_compras import descargar_con_progreso
from xml_cache import descargar_xmls, lineas_sri, resumir_motivos

# Caché de XML compartido entre jobs (misma ruta que pipeline_compras.CACHE_XML_DIR).
CACHE_XML_DIR = Path(__file__).parent / "cache" / "xml"

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


def _retenciones(job_id, ret_txt, odoo_ret, desde, hasta):
    """Descarga los XML (con caché) y concilia las retenciones emitidas. Devuelve (ret, n_sin_xml, fallidos)."""
    claves = list(ret_txt["clave"])
    fallidos = descargar_con_progreso(claves, descargar_xmls, CACHE_XML_DIR,
                                      lambda m, p: write_status(job_id, "running", m, p))
    lineas = lineas_sri(claves, CACHE_XML_DIR, contraparte=True)
    # Sin líneas = no descargado o XML ilegible/vacío: todos terminan como "XML no disponible".
    n_sin_xml = len(set(claves) - set(lineas["clave"]))
    write_status(job_id, "running", "Conciliando retenciones...", 0.7)
    ret = conciliar_retenciones_emitidas(lineas, odoo_ret, ret_txt, desde, hasta)
    return ret, n_sin_xml, fallidos


def run_pipeline_ventas(job_id: str) -> None:
    job_dir = get_job_dir(job_id)
    try:
        write_status(job_id, "running", "Leyendo TXT del SRI...", 0.1)
        rutas = sorted(str(p) for p in (job_dir / "emitidos").glob("*.txt"))
        docs_sri, ret_txt, advertencias = leer_emitidos_txts(rutas)
        if docs_sri.empty and ret_txt.empty:
            raise ValueError("Los TXT no contienen facturas, notas de crédito, notas de débito, "
                             "liquidaciones de compra ni retenciones")
        if not ret_txt.empty and ret_txt["fecha"].notna().sum() == 0:
            raise ValueError("El TXT de retenciones no tiene fechas de emisión válidas")
        desde, hasta, adv_periodo = periodo_emitidos(pd.concat([docs_sri["fecha"], ret_txt["fecha"]]))
        advertencias += adv_periodo

        # Odoo primero: credenciales erróneas u Odoo caído fallan antes de descargar XML.
        lector = OdooReader.from_env()
        if not docs_sri.empty:
            write_status(job_id, "running", "Leyendo comprobantes emitidos de Odoo...", 0.12)
            odoo = lector.fetch_emitidos(desde, hasta)
            advertencias += _tipos_sin_txt(docs_sri, odoo, desde, hasta)
        else:
            # Solo retenciones: sin documentos del SRI no se listan documentos de Odoo.
            odoo = pd.DataFrame(columns=COLUMNAS_EMITIDOS)
        odoo_ret = None
        if not ret_txt.empty:
            write_status(job_id, "running", "Leyendo retenciones emitidas de Odoo...", 0.13)
            odoo_ret = lector.fetch_retenciones_emitidas(desde, hasta)

        ret, n_sin_xml = None, 0
        if odoo_ret is not None:
            ret, n_sin_xml, fallidos = _retenciones(job_id, ret_txt, odoo_ret, desde, hasta)
            if n_sin_xml:
                advertencias.append(
                    f"{n_sin_xml} XML de retenciones no se pudieron descargar o leer; sus retenciones "
                    "se listan como 'XML no disponible'")
                if fallidos:
                    advertencias.append(f"Motivo de las descargas fallidas: {resumir_motivos(fallidos)}")

        write_status(job_id, "running", "Conciliando SRI vs Odoo...", 0.85)
        docs = clasificar_documentos(docs_sri, odoo, desde, hasta)
        res = construir_excel_emitidos(docs_sri, odoo, docs, ret, desde, hasta,
                                       str(job_dir / "resultado.xlsx"), advertencias=list(advertencias))
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
        write_status(job_id, "done", "Completado", 1.0, resumen=resumen)
        schedule_cleanup(job_id, delay_seconds=7200)
    except Exception as exc:
        write_status(job_id, "error", f"{type(exc).__name__}: {exc}",
                     error=traceback.format_exc())
        schedule_cleanup(job_id, delay_seconds=7200)
