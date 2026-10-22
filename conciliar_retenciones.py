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

    sri = sorted(_agregar(sri_lineas, ["clave", "ruc", "proveedor", "cod"]), key=_orden)
    if odoo_lineas.empty:
        posted, otros = odoo_lineas, odoo_lineas
    else:
        es_posted = odoo_lineas["estado"].fillna("") == "posted"
        posted, otros = odoo_lineas[es_posted], odoo_lineas[~es_posted]
    odoo = sorted(_agregar(posted, ["id_ret", "ruc", "socio", "cod"]), key=_orden)

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

    # (4) Porcentaje distinto: mismo n_ret y factura, distinto pct; uno a uno, prefiriendo el mismo
    # codigo de retencion y luego el valor mas cercano. El codigo es solo preferencia (SRI y Odoo
    # pueden usar sistemas de codigos distintos), nunca filtro.
    pares = []
    for i in sri_libre:
        s = sri[i]
        if s["n_ret"] is None or s["factura"] is None:
            continue
        for j in odoo_libre:
            o = odoo[j]
            if o["n_ret"] == s["n_ret"] and o["factura"] == s["factura"] and o["pct"] != s["pct"]:
                cs, co = _txt(s.get("cod")), _txt(o.get("cod"))
                otro_cod = 0 if cs is not None and cs == co else 1
                pares.append((otro_cod, abs(s["valor"] - o["valor"]), i, j))
    for _, _, i, j in sorted(pares):
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
    sin_xml = {}  # (ruc, serie) -> fila del TXT, solo comprobantes sin ninguna linea SRI
    if sri_txt is not None and not sri_txt.empty:
        for t in sri_txt.to_dict("records"):
            k = (_txt(t["ruc"]), _txt(t["serie"]))
            if k[1] is not None and k not in con_lineas:
                sin_xml.setdefault(k, t)
    usados = set()  # comprobantes sin XML ya representados por una fila de Odoo
    for j in sorted(odoo_libre):
        o = odoo[j]
        if not _en_periodo(o["fecha"], desde, hasta):
            continue
        ruc = _txt(o["ruc"])
        if ruc is not None:
            k = (ruc, o["n_ret"])
            cand = [k] if k in sin_xml else []
        else:
            # Sin RUC en Odoo: basta que algun comprobante del TXT con ese n_ret no tenga XML.
            cand = [k for k in sin_xml if k[1] == o["n_ret"]]
        if cand:
            clave = sin_xml[cand[0]]["clave"] if len(cand) == 1 else None
            # Ambiguo (varios emisores con ese n_ret): consume uno solo, el primero libre, para que
            # cada comprobante sin XML aparezca una vez.
            usados.add(next((k for k in sorted(cand, key=str) if k not in usados), cand[0]))
            filas.append(_fila(EST_XML, o=o, clave=clave))
        else:
            filas.append(_fila(EST_SOLO_ODOO, o=o))
    # Comprobantes del TXT sin XML ni fila de Odoo: posibles "Solo en SRI"; se listan siempre (el
    # lado SRI no se filtra por periodo) con los datos del TXT.
    for k in sorted((k for k in sin_xml if k not in usados), key=str):
        t = sin_xml[k]
        fila = _fila(EST_XML, clave=_txt(t["clave"]))
        fila.update(n_ret=k[1], ruc=k[0], proveedor=_txt(t.get("proveedor")), fecha=t.get("fecha"))
        filas.append(fila)

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


# ------------------------------------------------------------------ resumen por documento

COLUMNAS_DOC = ["doc", "estado", "n_ret", "fecha", "ruc", "proveedor", "factura_sri", "factura_odoo",
                "n_lineas_sri", "n_lineas_odoo", "valor_sri", "valor_odoo", "detalle", "clave", "id_ret"]

SIN_NUMERO = "(sin N°)"
_COMODINES = str.maketrans({"*": "_", "?": "_", "~": "_"})  # comodines de MATCH/COUNTIFS en Excel


def documento_por_linea(ret):
    """Clave de documento (texto estable) de cada linea de `ret`, alineada con su indice.

    - Documento = N° de retencion.
    - Si un mismo N° aparece con dos o mas RUC distintos (clientes que comparten numeracion), cada
      RUC es un documento ("N° | RUC"); las lineas sin RUC de ese N° forman "N° | (sin RUC)". Con un
      solo RUC, las lineas sin RUC se unen a ese documento.
    - Lineas sin N°: un documento por registro de Odoo ("(sin N°) id X"), o por clave de acceso, o
      por fila si no tienen ninguno de los dos; nunca se mezclan registros distintos.
    Los comodines de Excel (* ? ~) se reemplazan por "_" para que la clave sirva en MATCH/SUMIFS.
    """
    if ret.empty:
        return pd.Series([], index=ret.index, dtype=object)
    n_rets = [_txt(x) for x in ret["n_ret"]]
    rucs = [_txt(x) for x in ret["ruc"]]
    rucs_por_n = {}
    for n, r in zip(n_rets, rucs):
        if n is not None and r is not None:
            rucs_por_n.setdefault(n, set()).add(r)
    claves = []
    for pos, (n, r, id_ret, clave) in enumerate(zip(n_rets, rucs, ret["id_ret"], ret["clave"])):
        if n is not None:
            k = n if len(rucs_por_n.get(n, ())) <= 1 else f"{n} | {r or '(sin RUC)'}"
        elif id_ret is not None and not pd.isna(id_ret):
            k = f"{SIN_NUMERO} id {int(id_ret)}"
        elif _txt(clave) is not None:
            k = f"{SIN_NUMERO} clave {_txt(clave)}"
        else:
            k = f"{SIN_NUMERO} fila {pos + 1}"
        claves.append(k.translate(_COMODINES))
    return pd.Series(claves, index=ret.index, dtype=object)


def _plural(n, singular, plural):
    return f"{n} {singular if n == 1 else plural}"


def _estado_documento(estados):
    """Estado del documento a partir de los estados de sus lineas (ver resumir_por_documento)."""
    utiles = [e for e in estados if e != EST_INFO] or list(estados)
    unicos = set(utiles)
    if unicos == {EST_INFO}:
        return EST_INFO
    if unicos == {EST_XML}:
        return EST_XML
    if EST_FAC in unicos:
        return EST_FAC
    for est in (EST_OK, EST_SOLO_SRI, EST_SOLO_ODOO):
        if unicos == {est}:
            return est
    return EST_DIF


_TXT_DETALLE = {EST_DIF: "valor", EST_SOLO_SRI: "solo en SRI", EST_SOLO_ODOO: "solo en Odoo",
                EST_XML: "XML no disponible", EST_FAC: "factura incorrecta"}


def _detalle(estado, utiles):
    n = len(utiles)
    lineas = _plural(n, "línea", "líneas")
    if estado == EST_OK:
        return f"{lineas}: coinciden"
    if estado == EST_SOLO_SRI:
        return f"{lineas} solo en SRI"
    if estado == EST_SOLO_ODOO:
        return f"{lineas} solo en Odoo"
    if estado == EST_XML:
        return f"XML no disponible ({lineas})"
    if estado == EST_INFO:
        return f"Odoo cancelada/borrador ({lineas})"
    malas = [e for e in utiles if e != EST_OK]
    partes = [f"{_TXT_DETALLE.get(e, e)}: {malas.count(e)}" for e in
              (EST_FAC, EST_DIF, EST_SOLO_SRI, EST_SOLO_ODOO, EST_XML) if e in malas]
    texto = f"{len(malas)} de {lineas} con diferencia ({', '.join(partes)})"
    return f"Factura sustento distinta en Odoo; {texto}" if estado == EST_FAC else texto


def _unicos(valores):
    vistos = []
    for v in map(_txt, valores):
        if v is not None and v not in vistos:
            vistos.append(v)
    return vistos


def resumir_por_documento(ret: pd.DataFrame) -> pd.DataFrame:
    """Agrupa las lineas de `conciliar_retenciones` en documentos: UNA fila por comprobante de
    retencion (ver `documento_por_linea` para la clave). Columnas: COLUMNAS_DOC.

    Estado del documento (en este orden):
      1. Las lineas "Informativo" (Odoo cancelada/borrador) se ignoran si hay otras lineas.
      2. Todas Informativo -> Informativo. Todas "XML no disponible" -> XML no disponible.
      3. Alguna "Factura incorrecta en Odoo" -> Factura incorrecta en Odoo.
      4. Todas Coincide -> Coincide; todas Solo en SRI -> Solo en SRI; todas Solo en Odoo -> Solo en Odoo.
      5. Cualquier otra cosa (alguna Diferencia de valor, o mezcla como una linea que coincide y otra
         que esta en un solo lado) -> Diferencia de valor; `detalle` dice cuantas lineas difieren y por que.
    n_lineas_* y valor_* cuentan/suman las lineas consideradas (sin las informativas, salvo que el
    documento sea solo informativo). Orden: fecha, N° de retencion, clave de documento.
    """
    if ret is None or ret.empty:
        return pd.DataFrame(columns=COLUMNAS_DOC)
    ret = ret.reset_index(drop=True)
    claves = documento_por_linea(ret)
    filas = []
    for doc, g in ret.groupby(claves, sort=False):
        estados = list(g["estado"])
        estado = _estado_documento(estados)
        cons = g if estado == EST_INFO else g[g["estado"] != EST_INFO]
        vs = pd.to_numeric(cons["valor_sri"], errors="coerce")
        vo = pd.to_numeric(cons["valor_odoo"], errors="coerce")
        fechas = pd.to_datetime(g["fecha"], errors="coerce").dropna()
        ids = [int(x) for x in cons["id_ret"] if x is not None and not pd.isna(x)]
        filas.append({
            "doc": doc,
            "estado": estado,
            "n_ret": _primero(cons["n_ret"]),
            "fecha": fechas.min() if len(fechas) else pd.NaT,
            "ruc": _primero(cons["ruc"]),
            "proveedor": _primero(cons["proveedor"]),
            "factura_sri": ", ".join(_unicos(cons["factura_sri"])) or None,
            "factura_odoo": ", ".join(_unicos(cons["factura_odoo"])) or None,
            "n_lineas_sri": int(vs.notna().sum()),
            "n_lineas_odoo": int(vo.notna().sum()),
            "valor_sri": round(float(vs.sum()), 2),
            "valor_odoo": round(float(vo.sum()), 2),
            "detalle": _detalle(estado, [e for e in estados if e != EST_INFO] or estados),
            "clave": ", ".join(_unicos(cons["clave"])) or None,
            "id_ret": ", ".join(str(i) for i in dict.fromkeys(ids)) or None,
        })
    out = pd.DataFrame(filas, columns=COLUMNAS_DOC)
    out["fecha"] = pd.to_datetime(out["fecha"], errors="coerce")
    for col in ("n_ret", "ruc", "proveedor", "factura_sri", "factura_odoo", "clave", "id_ret"):
        out[col] = out[col].astype(object).where(out[col].notna(), pd.NA)
    orden = out.assign(_n=out["n_ret"].fillna("")).sort_values(
        ["fecha", "_n", "doc"], kind="mergesort", na_position="last")
    return out.loc[orden.index].reset_index(drop=True)
