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


ESPERAS_REINTENTO = (1, 3)  # segundos antes del 2.º y 3.º intento (3 intentos en total)
SIN_COMPROBANTE = "<numeroComprobantes>0</numeroComprobantes>"


def _es_transitorio(exc):
    """Error de red/timeout o HTTP 5xx: vale la pena reintentar. Un HTTP 4xx no."""
    resp = getattr(exc, "response", None)
    codigo = getattr(resp, "status_code", None)
    return codigo is None or codigo >= 500


def _consultar(sesion, clave, esperas):
    """Pide la autorización de `clave`. Devuelve (autorizacion, motivo); reintenta solo los fallos
    transitorios (excepciones de red, HTTP 5xx, respuesta sin <autorizacion>). Una respuesta definitiva
    del SRI (estado distinto de AUTORIZADO, o numeroComprobantes=0) no se reintenta."""
    body = SOAP_TEMPLATE.format(clave=clave).encode("utf-8")
    intentos = len(esperas) + 1
    motivo = None
    for n in range(intentos):
        if n:
            time.sleep(esperas[n - 1])
        transitorio = True
        try:
            r = sesion.post(ENDPOINT, data=body, headers=HEADERS, timeout=20)
            r.raise_for_status()
            autorizacion, motivo = evaluar_respuesta(r.text)
            if autorizacion:
                return autorizacion, None
            if SIN_COMPROBANTE in r.text:
                return None, ("el SRI no tiene el comprobante (numeroComprobantes=0: clave inexistente "
                              "o XML vencido)")
            transitorio = "<autorizacion>" not in r.text
        except Exception as exc:
            motivo = f"{type(exc).__name__}: {exc}"
            transitorio = _es_transitorio(exc)
        if not transitorio:
            break
    if transitorio and intentos > 1:
        motivo = f"{motivo} (tras {intentos} intentos)"
    return None, motivo


def descargar_xmls(claves, cache_dir, on_progress=None, sesion=None, pausa_seg=0.5,
                   esperas=ESPERAS_REINTENTO):
    """Descarga al caché solo las claves sin XML. Devuelve [{"clave", "motivo"}].

    Cada clave se intenta hasta len(esperas)+1 veces ante fallos transitorios del SRI."""
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
        autorizacion, motivo = _consultar(sesion, clave, esperas)
        if autorizacion:
            try:
                _escribir_atomico(cache_dir / f"{clave}.xml",
                                  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + autorizacion)
            except Exception as exc:
                fallidos.append({"clave": clave, "motivo": f"{type(exc).__name__}: {exc}"})
        else:
            fallidos.append({"clave": clave, "motivo": motivo})
        hechas += 1
        if pausa_seg:
            time.sleep(pausa_seg)
        if on_progress and ((i + 1) % 10 == 0 or i == len(pendientes) - 1):
            on_progress(hechas, total, len(fallidos))
    return fallidos


def resumir_motivos(fallidos, maximo=3):
    """Texto corto con los motivos de descarga más frecuentes ("motivo (×N); ...")."""
    conteo = {}
    for f in fallidos:
        conteo[f["motivo"]] = conteo.get(f["motivo"], 0) + 1
    top = sorted(conteo.items(), key=lambda kv: -kv[1])[:maximo]
    texto = "; ".join(f"{m} (×{n})" for m, n in top)
    return texto + ("; ..." if len(conteo) > maximo else "")


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
