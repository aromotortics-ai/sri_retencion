# -*- coding: utf-8 -*-
"""Conciliacion de retenciones RECIBIDAS: lineas del SRI (XML) vs lineas de Odoo (API).

Logica pura sobre DataFrames. Clave de cruce: (N° retencion, factura sustento, porcentaje),
tolerancia de $0.02 en el valor retenido (misma regla que conciliar_odoo.py), mas dos pasadas
sobre lo que queda sin pareja: factura incorrecta y porcentaje distinto.
"""
import pandas as pd

from conciliar_compras import _COLUMNAS_SRI

EST_OK = "Coincide"
EST_SOLO_SRI = "Solo en SRI"
EST_SOLO_ODOO = "Solo en Odoo"
EST_DIF = "Diferencia de valor"
EST_FAC = "Factura incorrecta en Odoo"
EST_XML = "XML no disponible"
EST_INFO = "Informativo (Odoo cancelada/borrador)"

TOL_CENTAVOS = 2  # tolerancia de $0.02

COLUMNAS_TXT = ["clave", "serie", "ruc", "proveedor", "fecha"]
COLUMNAS = ["estado", "n_ret", "fecha", "ruc", "proveedor", "factura_sri", "factura_odoo",
            "pct_sri", "pct_odoo", "base_sri", "base_odoo", "valor_sri", "valor_odoo", "dif_valor",
            "id_ret", "clave"]


def leer_sri_retenciones_txt(rutas):
    """Lee TXT del portal SRI (ISO-8859-1, tabulador) y devuelve solo los comprobantes de
    retencion con clave de acceso valida (45-49 digitos, completada a 49), sin duplicados."""
    partes = []
    for ruta in rutas:
        df = pd.read_csv(ruta, sep="\t", encoding="iso-8859-1", dtype=str)
        faltan = [c for c in _COLUMNAS_SRI if c not in df.columns]
        if faltan:
            raise ValueError(f"{ruta}: faltan columnas del portal SRI: {', '.join(faltan)}")
        partes.append(df)
    if not partes:
        return pd.DataFrame(columns=COLUMNAS_TXT)
    s = pd.concat(partes, ignore_index=True)
    s = s[s["TIPO_COMPROBANTE"].fillna("").str.startswith("Comprobante de Reten")].copy()
    clave = s["CLAVE_ACCESO"].fillna("").str.strip()
    s = s[clave.str.fullmatch(r"\d{45,49}")].copy()
    s["CLAVE_ACCESO"] = clave[s.index].str.zfill(49)
    s = s.drop_duplicates("CLAVE_ACCESO")
    out = pd.DataFrame({
        "clave": s["CLAVE_ACCESO"],
        "serie": s["SERIE_COMPROBANTE"].str.strip(),
        "ruc": s["RUC_EMISOR"],
        "proveedor": s["RAZON_SOCIAL_EMISOR"],
        "fecha": pd.to_datetime(s["FECHA_EMISION"], format="%d/%m/%Y", errors="coerce"),
    }, columns=COLUMNAS_TXT)
    return out.sort_values(["fecha", "serie", "clave"], kind="mergesort").reset_index(drop=True)


# ------------------------------------------------------------------------------ helpers

def _centavos(v):
    """Centavos enteros con signo (None si no hay valor)."""
    return int(round(float(v) * 100)) if pd.notna(v) else None


def _txt(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    v = str(v).strip()
    return v or None


def _pct(v):
    v = pd.to_numeric(v, errors="coerce")
    return None if pd.isna(v) else round(float(v), 4)


def _primero(serie):
    validos = serie.dropna()
    return validos.iloc[0] if len(validos) else None


def _agregar(df, campos_extra):
    """Agrupa por (n_ret, factura, pct) sumando base y valor. Las claves con NA se conservan
    (cada grupo con NA queda como una fila propia y nunca cruza por clave)."""
    if df.empty:
        return []
    d = df.copy()
    d["_n_ret"] = d["n_ret"].map(_txt)
    d["_fac"] = d["factura"].map(_txt)
    d["_pct"] = d["pct"].map(_pct)
    d["_cent_base"] = pd.to_numeric(d["base"], errors="coerce").map(_centavos)
    d["_cent_valor"] = pd.to_numeric(d["valor"], errors="coerce").map(_centavos)
    grupos = []
    for (n_ret, fac, pct), g in d.groupby(["_n_ret", "_fac", "_pct"], dropna=False, sort=False):
        n_ret, fac, pct = (None if pd.isna(x) else x for x in (n_ret, fac, pct))
        reg = {"n_ret": n_ret, "factura": fac, "pct": pct,
               "base": sum(c for c in g["_cent_base"] if c is not None),
               "valor": sum(c for c in g["_cent_valor"] if c is not None),
               "fecha": g["fecha"].min() if g["fecha"].notna().any() else pd.NaT}
        for campo in campos_extra:
            reg[campo] = _primero(g[campo]) if campo in g else None
        grupos.append(reg)
    return grupos


def _completa(r):
    return r["n_ret"] is not None and r["factura"] is not None and r["pct"] is not None


def _fila(estado, s=None, o=None, clave=None):
    s, o = s or {}, o or {}
    ambos = bool(s) and bool(o)
    fecha = s.get("fecha") if s and pd.notna(s.get("fecha")) else o.get("fecha", pd.NaT)
    return {
        "estado": estado,
        "n_ret": s.get("n_ret") or o.get("n_ret"),
        "fecha": fecha,
        "ruc": _txt(s.get("ruc")) or _txt(o.get("ruc")),
        "proveedor": _txt(s.get("proveedor")) or _txt(o.get("socio")),
        "factura_sri": s.get("factura"),
        "factura_odoo": o.get("factura"),
        "pct_sri": s.get("pct"),
        "pct_odoo": o.get("pct"),
        "base_sri": s["base"] / 100 if s else None,
        "base_odoo": o["base"] / 100 if o else None,
        "valor_sri": s["valor"] / 100 if s else None,
        "valor_odoo": o["valor"] / 100 if o else None,
        "dif_valor": (s["valor"] - o["valor"]) / 100 if ambos else None,
        "id_ret": o.get("id_ret"),
        "clave": s.get("clave") or clave,
    }


def _en_periodo(fecha, desde, hasta):
    if fecha is None or pd.isna(fecha):
        return True  # sin fecha no se puede afirmar que este fuera del periodo
    return desde <= pd.Timestamp(fecha).normalize() <= hasta


def _orden(r):
    return (r["n_ret"] or "", r["factura"] or "", r["pct"] if r["pct"] is not None else -1.0)


# --------------------------------------------------------------------------- conciliacion

def conciliar_retenciones(sri_lineas, odoo_lineas, sri_txt, desde, hasta):
    """Concilia lineas de retencion SRI vs Odoo. Una fila por linea conciliada (ver COLUMNAS)."""
    desde, hasta = pd.Timestamp(desde).normalize(), pd.Timestamp(hasta).normalize()
    sri_lineas = sri_lineas if sri_lineas is not None else pd.DataFrame()
    odoo_lineas = odoo_lineas if odoo_lineas is not None else pd.DataFrame()

    sri = sorted(_agregar(sri_lineas, ["clave", "ruc", "proveedor"]), key=_orden)
    if odoo_lineas.empty:
        posted, otros = odoo_lineas, odoo_lineas
    else:
        es_posted = odoo_lineas["estado"].fillna("") == "posted"
        posted, otros = odoo_lineas[es_posted], odoo_lineas[~es_posted]
    odoo = sorted(_agregar(posted, ["id_ret", "ruc", "socio"]), key=_orden)

    filas = []
    sri_libre = set(range(len(sri)))
    odoo_libre = set(range(len(odoo)))

    def emparejar(i, j, estado):
        filas.append(_fila(estado, sri[i], odoo[j]))
        sri_libre.discard(i)
        odoo_libre.discard(j)

    # (2) Cruce exacto por (n_ret, factura, pct). Tras agregar, cada clave es unica por lado.
    idx_odoo = {(o["n_ret"], o["factura"], o["pct"]): j for j, o in enumerate(odoo) if _completa(o)}
    for i, s in enumerate(sri):
        if not _completa(s):
            continue
        j = idx_odoo.get((s["n_ret"], s["factura"], s["pct"]))
        if j is not None:
            igual = abs(s["valor"] - odoo[j]["valor"]) <= TOL_CENTAVOS
            emparejar(i, j, EST_OK if igual else EST_DIF)

    # (3) Factura incorrecta: mismo n_ret, ultimos 9 digitos de factura y pct; distinta factura.
    for i in sorted(sri_libre):
        s = sri[i]
        if not _completa(s):
            continue
        cand = [j for j in sorted(odoo_libre)
                if _completa(odoo[j]) and odoo[j]["n_ret"] == s["n_ret"]
                and odoo[j]["pct"] == s["pct"] and odoo[j]["factura"] != s["factura"]
                and odoo[j]["factura"][-9:] == s["factura"][-9:]]
        if cand:
            emparejar(i, cand[0], EST_FAC)

    # (4) Porcentaje distinto: mismo n_ret y factura, distinto pct; uno a uno por valor cercano.
    pares = []
    for i in sri_libre:
        s = sri[i]
        if s["n_ret"] is None or s["factura"] is None:
            continue
        for j in odoo_libre:
            o = odoo[j]
            if o["n_ret"] == s["n_ret"] and o["factura"] == s["factura"] and o["pct"] != s["pct"]:
                pares.append((abs(s["valor"] - o["valor"]), i, j))
    for _, i, j in sorted(pares):
        if i in sri_libre and j in odoo_libre:
            emparejar(i, j, EST_DIF)

    # (5) Restos. SRI siempre genera fila; Odoo publicado solo si su fecha esta en el periodo.
    for i in sorted(sri_libre):
        filas.append(_fila(EST_SOLO_SRI, s=sri[i]))
    # XML faltante a nivel de comprobante, por (ruc emisor, n_ret): los numeros de retencion son
    # secuencias de cada emisor y pueden repetirse entre clientes distintos.
    con_lineas = set()
    if not sri_lineas.empty:
        con_lineas = {(_txt(r), _txt(n)) for r, n in zip(sri_lineas["ruc"], sri_lineas["n_ret"])}
    sin_xml = {}  # (ruc, serie) -> clave, solo comprobantes del TXT sin ninguna linea SRI
    if sri_txt is not None and not sri_txt.empty:
        for ruc, serie, clave in zip(sri_txt["ruc"], sri_txt["serie"], sri_txt["clave"]):
            k = (_txt(ruc), _txt(serie))
            if k[1] is not None and k not in con_lineas:
                sin_xml.setdefault(k, clave)
    for j in sorted(odoo_libre):
        o = odoo[j]
        if not _en_periodo(o["fecha"], desde, hasta):
            continue
        ruc = _txt(o["ruc"])
        if ruc is not None:
            k = (ruc, o["n_ret"])
            falta, clave = k in sin_xml, sin_xml.get(k)
        else:
            # Sin RUC en Odoo: basta que algun comprobante del TXT con ese n_ret no tenga XML.
            cand = [c for (_, serie), c in sin_xml.items() if serie == o["n_ret"]]
            falta, clave = bool(cand), (cand[0] if len(cand) == 1 else None)
        if falta:
            filas.append(_fila(EST_XML, o=o, clave=clave))
        else:
            filas.append(_fila(EST_SOLO_ODOO, o=o))

    # (6) Odoo cancelada/borrador: informativa, nunca participa del cruce.
    if not otros.empty:
        for _, g in otros.groupby("id_ret", dropna=False, sort=True):
            for o in sorted(_agregar(g, ["id_ret", "ruc", "socio"]), key=_orden):
                if _en_periodo(o["fecha"], desde, hasta):
                    filas.append(_fila(EST_INFO, o=o))

    return _a_dataframe(filas)


def _a_dataframe(filas):
    out = pd.DataFrame(filas, columns=COLUMNAS)
    for col in ("pct_sri", "pct_odoo", "base_sri", "base_odoo", "valor_sri", "valor_odoo",
                "dif_valor"):
        out[col] = pd.to_numeric(out[col], errors="coerce").astype(float)
    out["id_ret"] = pd.to_numeric(out["id_ret"], errors="coerce").astype("Int64")
    out["fecha"] = pd.to_datetime(out["fecha"], errors="coerce")
    for col in ("n_ret", "ruc", "proveedor", "factura_sri", "factura_odoo", "clave"):
        out[col] = out[col].astype(object).where(out[col].notna(), pd.NA)
    if out.empty:
        return out.reset_index(drop=True)
    orden = out.assign(
        _n=out["n_ret"].fillna(""),
        _f=out["factura_sri"].fillna(out["factura_odoo"]).fillna(""),
        _p=out["pct_sri"].fillna(out["pct_odoo"]).fillna(-1.0),
    ).sort_values(["_n", "_f", "_p"], kind="mergesort")
    return out.loc[orden.index].reset_index(drop=True)
