"""
Parser de comprobantes de retención electrónicos del SRI (Ecuador), en el
formato que devuelve tanto la descarga individual del portal como el web
service público AutorizacionComprobantes (mismo esquema <autorizacion>...).

Cada archivo XML produce:
  - un registro de CABECERA (un comprobante)
  - N registros de DETALLE (una fila por cada <impuesto>, que es el nivel
    en el que aparece numDocSustento = el documento de compra al que
    corresponde esa retención — la llave real para conciliar con el ERP)
"""
import glob
import xml.etree.ElementTree as ET
import pandas as pd


def _text(el, tag, default=""):
    node = el.find(tag)
    return node.text.strip() if node is not None and node.text else default


def parse_comprobante_xml(path):
    tree = ET.parse(path)
    root = tree.getroot()

    estado = _text(root, "estado")
    num_autorizacion = _text(root, "numeroAutorizacion")
    fecha_autorizacion = _text(root, "fechaAutorizacion")

    comprobante_cdata = root.find("comprobante").text
    comp_root = ET.fromstring(comprobante_cdata)

    info_trib = comp_root.find("infoTributaria")
    info_ret = comp_root.find("infoCompRetencion")

    cabecera = {
        "Estado": estado,
        "Clave_Acceso": _text(info_trib, "claveAcceso"),
        "Fecha_Autorizacion": fecha_autorizacion,
        "RUC_Emisor": _text(info_trib, "ruc"),
        "Razon_Social_Emisor": _text(info_trib, "razonSocial"),
        "Establecimiento": _text(info_trib, "estab"),
        "Punto_Emision": _text(info_trib, "ptoEmi"),
        "Secuencial": _text(info_trib, "secuencial"),
        "Fecha_Emision": _text(info_ret, "fechaEmision"),
        "Identificacion_Sujeto_Retenido": _text(info_ret, "identificacionSujetoRetenido"),
        "Razon_Social_Sujeto_Retenido": _text(info_ret, "razonSocialSujetoRetenido"),
        "Periodo_Fiscal": _text(info_ret, "periodoFiscal"),
    }

    detalle = []
    version = comp_root.get("version", "1")
    if version.startswith("2"):
        # v2.0.0: docsSustento/docSustento/retenciones/retencion
        for doc in comp_root.find("docsSustento").findall("docSustento"):
            num_doc = _text(doc, "numDocSustento")
            cod_doc = _text(doc, "codDocSustento")
            fecha_doc = _text(doc, "fechaEmisionDocSustento")
            retenciones = doc.find("retenciones")
            if retenciones is None:
                continue
            for ret in retenciones.findall("retencion"):
                fila = dict(cabecera)
                fila.update({
                    "Codigo_Retencion": _text(ret, "codigoRetencion"),
                    "Base_Imponible": float(_text(ret, "baseImponible", "0") or 0),
                    "Porcentaje_Retener": float(_text(ret, "porcentajeRetener", "0") or 0),
                    "Valor_Retenido": float(_text(ret, "valorRetenido", "0") or 0),
                    "Cod_Doc_Sustento": cod_doc,
                    "Num_Doc_Sustento": num_doc,
                    "Fecha_Emision_Doc_Sustento": fecha_doc,
                })
                detalle.append(fila)
    else:
        # v1.0.0: impuestos/impuesto
        for imp in comp_root.find("impuestos").findall("impuesto"):
            fila = dict(cabecera)
            fila.update({
                "Codigo_Retencion": _text(imp, "codigoRetencion"),
                "Base_Imponible": float(_text(imp, "baseImponible", "0") or 0),
                "Porcentaje_Retener": float(_text(imp, "porcentajeRetener", "0") or 0),
                "Valor_Retenido": float(_text(imp, "valorRetenido", "0") or 0),
                "Cod_Doc_Sustento": _text(imp, "codDocSustento"),
                "Num_Doc_Sustento": _text(imp, "numDocSustento"),
                "Fecha_Emision_Doc_Sustento": _text(imp, "fechaEmisionDocSustento"),
            })
            detalle.append(fila)

    cabecera["Valor_Total_Retenido"] = sum(f["Valor_Retenido"] for f in detalle)
    cabecera["Num_Doc_Sustento"] = detalle[0]["Num_Doc_Sustento"] if detalle else ""

    return cabecera, detalle


def parse_folder(folder_glob):
    """folder_glob: p.ej. 'xml_samples/*.xml'"""
    cabeceras, detalles = [], []
    for path in sorted(glob.glob(folder_glob)):
        try:
            cab, det = parse_comprobante_xml(path)
            cabeceras.append(cab)
            detalles.extend(det)
        except Exception as e:
            print(f"ERROR procesando {path}: {e}")

    df_cab = pd.DataFrame(cabeceras)
    df_det = pd.DataFrame(detalles)
    return df_cab, df_det


if __name__ == "__main__":
    import sys
    folder_glob = sys.argv[1] if len(sys.argv) > 1 else "xml_samples/*.xml"
    df_cab, df_det = parse_folder(folder_glob)
    print(f"{len(df_cab)} comprobantes, {len(df_det)} líneas de retención")
    df_cab.to_excel("comprobantes_cabecera.xlsx", index=False)
    df_det.to_excel("comprobantes_detalle.xlsx", index=False)
