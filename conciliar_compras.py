# -*- coding: utf-8 -*-
"""Conciliacion de comprobantes electronicos recibidos (facturas, NC y ND) SRI vs Odoo."""
import os
import shutil
import subprocess
import tempfile

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


CAT_INFO = "No aplica (informativo)"
CAT_DUP = "Duplicado en Odoo"
CAT_SOLO_ODOO = "Solo en Odoo"
TOL_CENTAVOS = 2  # tolerancia de $0.02


def _centavos(v):
    return int(round(abs(float(v)) * 100)) if pd.notna(v) else None


def _mismo_total(a, b):
    ca, cb = _centavos(a), _centavos(b)
    return ca is not None and cb is not None and abs(ca - cb) <= TOL_CENTAVOS


def _ruc_compatible(a, b):
    """True si los RUC coinciden o si alguno de los lados no lo tiene."""
    return pd.isna(a) or pd.isna(b) or str(a) == str(b)


def conciliar(sri, odoo, desde, hasta):
    """Cruza SRI vs Odoo por tipo. Devuelve idx_sri, id_odoo, tipo, metodo, cat (una fila por pareja
    y una por cada lado sin pareja)."""
    por_clave, por_serie, por_afectada = {}, {}, {}
    for i, r in sri.iterrows():
        por_clave[(r["tipo"], r["clave"])] = i
        por_serie.setdefault((r["tipo"], r["serie"]), []).append(i)
        if r["tipo"] == "NC" and pd.notna(r["afectada"]):
            por_afectada.setdefault(r["afectada"], []).append(i)

    usados, filas = set(), []

    def por_clave_exacta(o):
        if pd.notna(o["clave"]) and (o["tipo"], o["clave"]) in por_clave:
            return por_clave[(o["tipo"], o["clave"])], "Clave de acceso"
        return None, ""

    def buscar_flexible(o):
        t = o["tipo"]
        cand = [i for i in por_serie.get((t, o["nfac"]), []) if _ruc_compatible(sri.at[i, "ruc"], o["ruc"])]
        iguales = [i for i in cand if _mismo_total(sri.at[i, "total"], o["total"])]
        if iguales:
            libres = [i for i in iguales if i not in usados]
            return (libres or iguales)[0], "N° documento + importe"
        if len(cand) == 1:
            return cand[0], "N° documento (importe distinto)"
        if t == "NC" and pd.notna(o["afectada"]):
            iguales = [i for i in por_afectada.get(o["afectada"], [])
                       if _ruc_compatible(sri.at[i, "ruc"], o["ruc"])
                       and _mismo_total(sri.at[i, "total"], o["total"])]
            if iguales:
                libres = [i for i in iguales if i not in usados]
                return (libres or iguales)[0], "NC: factura afectada + importe"
        return None, ""

    def registrar(o, idx, metodo):
        cat = CAT_DUP if idx in usados else ""
        usados.add(idx)
        return (idx, o["id"], o["tipo"], metodo, cat)

    # Pasada 1: clave de acceso para todas las filas elegibles (no depende del orden de Odoo).
    resultado = {}
    for pos, (_, o) in enumerate(odoo.iterrows()):
        if o["elegible"]:
            idx, metodo = por_clave_exacta(o)
            if idx is not None:
                resultado[pos] = registrar(o, idx, metodo)
    # Pasada 2: demas niveles solo para las filas sin pareja.
    for pos, (_, o) in enumerate(odoo.iterrows()):
        if pos in resultado:
            filas.append(resultado[pos])
            continue
        idx, metodo = (None, "") if not o["elegible"] else buscar_flexible(o)
        if idx is not None:
            filas.append(registrar(o, idx, metodo))
            continue
        if not (desde <= pd.Timestamp(o["fecha"]).date() <= hasta):
            continue
        filas.append((None, o["id"], o["tipo"], "", CAT_INFO if not o["elegible"] else CAT_SOLO_ODOO))

    for i in sri.index:
        if i not in usados:
            filas.append((i, None, sri.at[i, "tipo"], "", ""))
    out = pd.DataFrame(filas, columns=["idx_sri", "id_odoo", "tipo", "metodo", "cat"])
    out["idx_sri"] = out["idx_sri"].astype("Int64")
    out["id_odoo"] = out["id_odoo"].astype("Int64")
    return out


_NOMBRE_TIPO = {"FAC": "facturas", "NC": "notas de crédito", "ND": "notas de débito"}


def _recalcular(ruta):
    """Recalcula las fórmulas con LibreOffice. Devuelve True si se pudo."""
    soffice = shutil.which("soffice")
    if not soffice:
        return False
    try:
        with tempfile.TemporaryDirectory() as tmp:
            subprocess.run([soffice, "--headless", "--convert-to", "xlsx", "--outdir", tmp, ruta],
                           check=True, capture_output=True, timeout=120)
            nuevo = os.path.join(tmp, os.path.basename(ruta))
            if not os.path.exists(nuevo):
                return False
            shutil.move(nuevo, ruta)
        return True
    except Exception:
        return False


def conciliar_compras(rutas_txt, odoo, desde, hasta, salida, ret=None, periodo_ret=None):
    """Concilia TXT del SRI contra Odoo, escribe el Excel y devuelve
    {"por_tipo": {...}, "advertencias": [...]}.

    ret: conciliación de retenciones ya calculada (DataFrame de conciliar_retenciones) o None si
    no se cargó TXT de retenciones; con ret, por_tipo incluye "retenciones".
    periodo_ret: (desde, hasta) de las retenciones, rotulado en el Resumen del Excel."""
    from excel_compras import construir_excel  # import tardío: excel_compras importa este módulo
    sri = leer_sri_txts(rutas_txt)
    pares = conciliar(sri, odoo, desde, hasta)
    por_tipo = construir_excel(sri, odoo, pares, desde, hasta, salida, ret=ret, periodo_ret=periodo_ret)
    advertencias = [f"Falta el TXT de {_NOMBRE_TIPO[t]}" for t in tipos_faltantes(sri)]
    if ret is None:
        advertencias.append("Falta el TXT de retenciones")
    if not _recalcular(salida):
        advertencias.append("No se pudo recalcular con LibreOffice; Excel recalculará las fórmulas al abrir el archivo")
    return {"por_tipo": por_tipo, "advertencias": advertencias}
