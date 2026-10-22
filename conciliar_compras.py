# -*- coding: utf-8 -*-
"""Conciliacion de comprobantes electronicos recibidos (facturas, NC y ND) SRI vs Odoo."""
import pandas as pd

TIPOS = ("FAC", "NC", "ND")

_COLUMNAS_SRI = ["RUC_EMISOR", "RAZON_SOCIAL_EMISOR", "TIPO_COMPROBANTE", "SERIE_COMPROBANTE",
                 "CLAVE_ACCESO", "FECHA_EMISION", "VALOR_SIN_IMPUESTOS", "IVA", "IMPORTE_TOTAL",
                 "NUMERO_DOCUMENTO_MODIFICADO"]

_PREFIJOS_TIPO = (("Factura", "FAC"), ("Nota de Cr", "NC"), ("Nota de D", "ND"))


def _tipo(texto):
    texto = texto if isinstance(texto, str) else ""
    for prefijo, codigo in _PREFIJOS_TIPO:
        if texto.startswith(prefijo):
            return codigo
    return None


def leer_sri_txts(rutas):
    """Lee TXT del portal SRI (ISO-8859-1, tabulador) y devuelve FAC/NC/ND sin duplicados."""
    partes = []
    for ruta in rutas:
        df = pd.read_csv(ruta, sep="\t", encoding="iso-8859-1", dtype=str)
        faltan = [c for c in _COLUMNAS_SRI if c not in df.columns]
        if faltan:
            raise ValueError(f"{ruta}: faltan columnas del portal SRI: {', '.join(faltan)}")
        partes.append(df)
    s = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame(columns=_COLUMNAS_SRI)
    s["tipo"] = s["TIPO_COMPROBANTE"].map(_tipo)
    s = s[s["tipo"].notna()].drop_duplicates("CLAVE_ACCESO").copy()
    for col, n in (("VALOR_SIN_IMPUESTOS", "base"), ("IVA", "iva"), ("IMPORTE_TOTAL", "total")):
        s[n] = pd.to_numeric(s[col], errors="coerce")
    s["base"] = s["base"].fillna(0.0)
    s["iva"] = s["iva"].fillna(0.0)
    s["total"] = s["total"].fillna(s["base"] + s["iva"])
    out = pd.DataFrame({
        "clave": s["CLAVE_ACCESO"],
        "serie": s["SERIE_COMPROBANTE"],
        "ruc": s["RUC_EMISOR"],
        "proveedor": s["RAZON_SOCIAL_EMISOR"],
        "fecha": pd.to_datetime(s["FECHA_EMISION"], format="%d/%m/%Y"),
        "tipo": s["tipo"],
        "base": s["base"],
        "iva": s["iva"],
        "total": s["total"],
        "afectada": s["NUMERO_DOCUMENTO_MODIFICADO"].replace("", pd.NA),
    })
    return out.sort_values(["fecha", "serie"]).reset_index(drop=True)


def tipos_faltantes(sri):
    """Tipos de TIPOS sin ninguna fila en el DataFrame del SRI."""
    presentes = set(sri["tipo"])
    return [t for t in TIPOS if t not in presentes]
