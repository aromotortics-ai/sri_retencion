# -*- coding: utf-8 -*-
"""Lector de los TXT de comprobantes EMITIDOS exportados del portal SRI.

Los TXT (ISO-8859-1, tabulador) pueden mezclar tipos de comprobante; el tipo se detecta por la
columna COMPROBANTE, nunca por el nombre del archivo.
"""
import unicodedata

import pandas as pd

RUC_EMPRESA = "2390000831001"
TIPOS_DOC = ("FAC", "NC", "ND", "LIQ")

_COLUMNAS_EMITIDOS = ["COMPROBANTE", "SERIE_COMPROBANTE", "CLAVE_ACCESO", "FECHA_EMISION",
                      "VALOR_SIN_IMPUESTOS", "IVA", "IMPORTE_TOTAL"]
_COLUMNAS_DOCS = ["clave", "serie", "fecha", "tipo", "base", "iva", "total", "ruc", "proveedor",
                  "afectada"]
_COLUMNAS_RET = ["clave", "serie", "ruc", "proveedor", "fecha"]

# (prefijo normalizado, tipo); se compara solo hasta antes del caracter acentuado
_PREFIJOS = [("factura", "FAC"), ("nota de cr", "NC"), ("nota de d", "ND"),
             ("liquidaci", "LIQ"), ("comprobante de reten", "RET"), ("gu", "GUIA")]

_DIAS_MAX_PERIODO = 35


def _tipo(texto):
    t = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    t = t.strip().lower()
    for prefijo, tipo in _PREFIJOS:
        if t.startswith(prefijo):
            return tipo
    return None


def _fechas(serie):
    """dd/mm/aaaa o dd/mm/aaaa hh:mm:ss -> Timestamp normalizado (sin hora)."""
    s = serie.fillna("").astype(str).str.strip().str.slice(0, 10)
    return pd.to_datetime(s, format="%d/%m/%Y", errors="coerce")


def _num(serie):
    return pd.to_numeric(serie.fillna("").astype(str).str.strip().str.replace(",", "."),
                         errors="coerce")


def leer_emitidos_txts(rutas):
    """Devuelve (docs, ret, advertencias) a partir de los TXT de emitidos."""
    partes = []
    for ruta in rutas:
        df = pd.read_csv(ruta, sep="\t", encoding="iso-8859-1", dtype=str)
        faltan = [c for c in _COLUMNAS_EMITIDOS if c not in df.columns]
        if faltan:
            raise ValueError(f"{ruta}: faltan columnas del portal SRI: {', '.join(faltan)}")
        partes.append(df[_COLUMNAS_EMITIDOS])
    advertencias = []
    if not partes:
        return (pd.DataFrame(columns=_COLUMNAS_DOCS), pd.DataFrame(columns=_COLUMNAS_RET),
                advertencias)

    s = pd.concat(partes, ignore_index=True)
    s["CLAVE_ACCESO"] = s["CLAVE_ACCESO"].fillna("").str.strip()
    s["_tipo"] = s["COMPROBANTE"].map(_tipo)

    guias = int((s["_tipo"] == "GUIA").sum())
    if guias:
        advertencias.append(f"Se ignoraron {guias} guías de remisión (no se concilian).")
    otros = s[s["_tipo"].isna() & (s["COMPROBANTE"].fillna("").str.strip() != "")]
    if len(otros):
        advertencias.append(f"Se ignoraron {len(otros)} filas con tipo de comprobante desconocido: "
                            f"{', '.join(sorted(set(otros['COMPROBANTE'])))}.")
    s = s[s["_tipo"].isin(TIPOS_DOC + ("RET",))]

    # Duplicados por clave (conservando la primera fila); avisar si los datos difieren
    con_clave = s[s["CLAVE_ACCESO"] != ""]
    campos = ["_tipo", "SERIE_COMPROBANTE", "FECHA_EMISION", "VALOR_SIN_IMPUESTOS", "IVA",
              "IMPORTE_TOTAL"]
    for clave, grupo in con_clave[con_clave.duplicated("CLAVE_ACCESO", keep=False)].groupby(
            "CLAVE_ACCESO"):
        if len(grupo[campos].fillna("").drop_duplicates()) > 1:
            advertencias.append(f"La clave {clave} aparece repetida con datos distintos; "
                                f"se conservó la primera fila.")
    s = pd.concat([s[s["CLAVE_ACCESO"] == ""], con_clave.drop_duplicates("CLAVE_ACCESO")])

    # --- documentos (FAC/NC/ND/LIQ)
    d = s[s["_tipo"].isin(TIPOS_DOC)]
    base = _num(d["VALOR_SIN_IMPUESTOS"])
    iva = _num(d["IVA"])
    total = _num(d["IMPORTE_TOTAL"])
    total = total.fillna(base.fillna(0) + iva.fillna(0))
    docs = pd.DataFrame({
        "clave": d["CLAVE_ACCESO"],
        "serie": d["SERIE_COMPROBANTE"].str.strip(),
        "fecha": _fechas(d["FECHA_EMISION"]),
        "tipo": d["_tipo"],
        "base": base,
        "iva": iva,
        "total": total,
        "ruc": pd.NA,
        "proveedor": pd.NA,
        "afectada": pd.NA,
    }, columns=_COLUMNAS_DOCS)
    docs = docs.sort_values(["fecha", "serie", "clave"], kind="mergesort").reset_index(drop=True)

    # --- retenciones (clave valida de 45-49 digitos, completada a 49)
    r = s[s["_tipo"] == "RET"]
    valida = r["CLAVE_ACCESO"].str.fullmatch(r"\d{45,49}")
    if (~valida).sum():
        advertencias.append(f"{int((~valida).sum())} retenciones del TXT se ignoraron por no tener "
                            "clave de acceso válida")
    r = r[valida]
    ret = pd.DataFrame({
        "clave": r["CLAVE_ACCESO"].str.zfill(49),
        "serie": r["SERIE_COMPROBANTE"].str.strip(),
        "ruc": RUC_EMPRESA,
        "proveedor": pd.NA,
        "fecha": _fechas(r["FECHA_EMISION"]),
    }, columns=_COLUMNAS_RET)
    ret = ret.sort_values(["fecha", "serie", "clave"], kind="mergesort").reset_index(drop=True)
    return docs, ret, advertencias


def periodo_emitidos(fechas):
    """(desde, hasta, advertencias): min/max de las fechas; avisa si el rango supera 35 dias."""
    f = pd.to_datetime(fechas, errors="coerce").dropna()
    if f.empty:
        raise ValueError("No hay fechas de emisión válidas en los TXT.")
    desde, hasta = f.min().date(), f.max().date()
    advertencias = []
    dias = (hasta - desde).days
    if dias > _DIAS_MAX_PERIODO:
        advertencias.append(f"Las fechas de emisión abarcan {dias} días ({desde:%d/%m/%Y} a "
                            f"{hasta:%d/%m/%Y}); revise que los TXT sean del mismo período.")
    return desde, hasta, advertencias
