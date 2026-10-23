# -*- coding: utf-8 -*-
"""
Descarga de XML de retenciones por clave de acceso (web service público del SRI)
hacia el caché compartido y lectura de sus líneas como DataFrame.
"""
import os
import tempfile
import time
from pathlib import Path

import pandas as pd

from descargar_comprobantes_sri import evaluar_respuesta
from parse_comprobante_xml import parse_comprobante_xml
from sri_session import sri_session

ENDPOINT = "https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline"
SOAP_TEMPLATE = """<?xml version="1.0" encoding="UTF-8"?>
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ec="http://ec.gob.sri.ws.autorizacion">
  <soapenv:Header/>
  <soapenv:Body>
    <ec:autorizacionComprobante>
      <claveAccesoComprobante>{clave}</claveAccesoComprobante>
    </ec:autorizacionComprobante>
  </soapenv:Body>
</soapenv:Envelope>"""
HEADERS = {"Content-Type": "text/xml; charset=utf-8", "Connection": "close"}

COLUMNAS = ["clave", "n_ret", "fecha", "ruc", "proveedor",
            "factura", "cod", "pct", "base", "valor"]


def formatear_doc(v):
    """Num. de documento 3-3-9 con guiones (misma regla que build_workbook_final)."""
    if isinstance(v, str) and "-" not in v and len(v) >= 6:
        return f"{v[:3]}-{v[3:6]}-{v[6:]}"
    return v


def _escribir_atomico(destino, texto):
    """Escribe en un temporal del mismo directorio y lo renombra: otro job que lea el caché ve el
    archivo completo o no lo ve. El temporal (`<clave>.xml.<aleatorio>.tmp`) no coincide con
    `<clave>.xml`, así que ni la verificación de caché ni lineas_sri lo consideran."""
    destino = Path(destino)
    fd, tmp = tempfile.mkstemp(dir=destino.parent, prefix=destino.name + ".", suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(texto)
        os.replace(tmp, destino)
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


def descargar_xmls(claves, cache_dir, on_progress=None, sesion=None, pausa_seg=0.5):
    """Descarga al caché solo las claves sin XML. Devuelve [{"clave", "motivo"}]."""
    cache_dir = Path(cache_dir)
    cache_dir.mkdir(parents=True, exist_ok=True)
    total = len(claves)
    pendientes = [c for c in claves if not (cache_dir / f"{c}.xml").exists()]
    hechas = total - len(pendientes)
    fallidos = []
    if not pendientes:
        return fallidos
    sesion = sesion or sri_session()

    for i, clave in enumerate(pendientes):
        body = SOAP_TEMPLATE.format(clave=clave).encode("utf-8")
        try:
            r = sesion.post(ENDPOINT, data=body, headers=HEADERS, timeout=20)
            r.raise_for_status()
            autorizacion, motivo = evaluar_respuesta(r.text)
            if autorizacion:
                _escribir_atomico(cache_dir / f"{clave}.xml",
                                  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + autorizacion)
            else:
                fallidos.append({"clave": clave, "motivo": motivo})
        except Exception as exc:
            fallidos.append({"clave": clave, "motivo": f"{type(exc).__name__}: {exc}"})
        hechas += 1
        if pausa_seg:
            time.sleep(pausa_seg)
        if on_progress and ((i + 1) % 10 == 0 or i == len(pendientes) - 1):
            on_progress(hechas, total, len(fallidos))
    return fallidos


def lineas_sri(claves, cache_dir, contraparte: bool = False):
    """Una fila por línea de retención, solo de las claves dadas presentes en el caché.

    Con contraparte=True `proveedor` es la razón social del sujeto retenido (retenciones
    emitidas: el emisor es la empresa); `ruc` sigue siendo el del emisor."""
    cache_dir = Path(cache_dir)
    filas = []
    for clave in dict.fromkeys(claves):
        path = cache_dir / f"{clave}.xml"
        if not path.exists():
            continue
        try:
            cab, detalle = parse_comprobante_xml(path)
        except Exception:
            continue
        for d in detalle:
            filas.append({
                "clave": clave,
                "n_ret": f"{d['Establecimiento']}-{d['Punto_Emision']}-{d['Secuencial']}",
                "fecha": pd.to_datetime(d["Fecha_Emision"], format="%d/%m/%Y", errors="coerce"),
                "ruc": d["RUC_Emisor"],
                "proveedor": (d.get("Razon_Social_Sujeto_Retenido", cab.get("Razon_Social_Sujeto_Retenido", ""))
                              if contraparte else d["Razon_Social_Emisor"]),
                "factura": formatear_doc(d["Num_Doc_Sustento"]),
                "cod": d["Codigo_Retencion"],
                "pct": d["Porcentaje_Retener"],
                "base": d["Base_Imponible"],
                "valor": d["Valor_Retenido"],
            })
    return pd.DataFrame(filas, columns=COLUMNAS)
