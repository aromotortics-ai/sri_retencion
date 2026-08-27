"""
Descarga automática de comprobantes electrónicos del SRI por clave de acceso,
usando el web service PÚBLICO de autorización (no requiere usuario/clave del
portal, no requiere certificado — es el mismo servicio que usan terceros para
validar comprobantes ajenos).

WSDL (producción): https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantes?wsdl
Operación: autorizacionComprobante(claveAccesoComprobante)

IMPORTANTE: este script no se pudo probar en vivo desde este entorno porque
el sandbox donde se generó solo tiene salida de red a repositorios de
paquetes (pypi, npm, github, etc.), no a dominios .gob.ec. Antes de
automatizarlo en tu servidor, córrelo primero con una sola clave de acceso
y confirma que el XML de salida coincide con el que ya descargaste manualmente.

Uso:
    python descargar_comprobantes_sri.py claves.txt carpeta_salida/

    claves.txt: un número de clave de acceso (49 dígitos) por línea —
    exactamente la columna "Clave de acceso / Nro. autorización" que ya
    trae el reporte "Exportar Txt" del portal.
"""
import os
import re
import sys
import time
from sri_session import sri_session

# Endpoint probado en producción agosto 2026.
# AutorizacionComprobantes (sin Offline) cierra la conexión sin respuesta;
# AutorizacionComprobantesOffline funciona correctamente con Connection: close.
WSDL_ENDPOINT = "https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline"

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


def descargar_uno(clave_acceso, timeout=20):
    body = SOAP_TEMPLATE.format(clave=clave_acceso.strip())
    resp = sri_session().post(WSDL_ENDPOINT, data=body.encode("utf-8"),
                              headers=HEADERS, timeout=timeout)
    resp.raise_for_status()
    return resp.text  # respuesta SOAP completa; el XML del comprobante viene dentro


def extraer_autorizacion(soap_response_text):
    """Recorta el bloque <autorizacion>...</autorizacion> de la respuesta SOAP,
    que es el mismo formato del XML que se descarga manualmente desde el portal."""
    match = re.search(r"<autorizacion>.*?</autorizacion>", soap_response_text, re.DOTALL)
    if not match:
        return None
    return match.group(0)


def descargar_lote(claves, carpeta_salida, pausa_seg=0.5):
    os.makedirs(carpeta_salida, exist_ok=True)
    ok, fallidos = 0, []
    for clave in claves:
        clave = clave.strip()
        if not clave:
            continue
        try:
            soap_xml = descargar_uno(clave)
            autorizacion_xml = extraer_autorizacion(soap_xml)
            if autorizacion_xml is None:
                fallidos.append((clave, "no se encontró <autorizacion> en la respuesta"))
                continue
            destino = os.path.join(carpeta_salida, f"{clave}.xml")
            with open(destino, "w", encoding="utf-8") as f:
                f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + autorizacion_xml)
            ok += 1
        except Exception as e:
            fallidos.append((clave, str(e)))
        time.sleep(pausa_seg)  # trato cortés al servicio público, evita bloqueos por rate-limit

    print(f"Descargados correctamente: {ok}")
    if fallidos:
        print(f"Fallaron: {len(fallidos)}")
        for clave, err in fallidos[:20]:
            print(f"  {clave}: {err}")
    return ok, fallidos


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Uso: python descargar_comprobantes_sri.py claves.txt carpeta_salida/")
        sys.exit(1)
    with open(sys.argv[1], encoding="utf-8") as f:
        claves = [l.strip() for l in f if l.strip()]
    descargar_lote(claves, sys.argv[2])
