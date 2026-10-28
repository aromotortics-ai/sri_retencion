# -*- coding: utf-8 -*-
"""Cifras del tablero de conciliación (tarjetas por tipo y lista de excepciones), JSON-serializables."""
import pandas as pd

from app.reportes.observaciones import ACCIONABLES, QUE_CORREGIR
from app.sri.emitidos_sri import TIPOS_DOC

ETIQUETAS = {"FAC": "Facturas", "NC": "Notas de crédito", "ND": "Notas de débito",
             "LIQ": "Liquidaciones", "RET": "Retenciones"}
EST_OK = "Coincide"
_REPETIDOS = ("Duplicado en Odoo",)


def _v(x):
    """Valor Python serializable: NaN/NA/NaT -> None, fechas -> ISO, numpy -> nativo."""
    if x is None:
        return None
    try:
        if pd.isna(x):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(x, (pd.Timestamp,)) or hasattr(x, "isoformat"):
        return x.date().isoformat() if isinstance(x, pd.Timestamp) else x.isoformat()
    if hasattr(x, "item"):
        return x.item()
    return x


def _primero(*valores):
    for x in valores:
        x = _v(x)
        if x is not None and x != "":
            return x
    return None


def _num(x):
    x = _v(x)
    return None if x is None else round(float(x), 2)


def _fila(tipo, estado, documento, fecha, contraparte, ruc, vs, vo, detalle):
    vs, vo = _num(vs), _num(vo)
    accionable = estado in ACCIONABLES
    return {"tipo": tipo, "estado": estado, "documento": documento, "fecha": _v(fecha),
            "contraparte": contraparte, "ruc": ruc, "valor_sri": vs, "valor_odoo": vo,
            "diferencia": None if vs is None or vo is None else round(vs - vo, 2),
            "detalle": _v(detalle) or "", "que_corregir": QUE_CORREGIR.get(estado, ""),
            "accionable": accionable}


def _filas_docs(docs, sri, odoo):
    od = odoo.set_index("id") if odoo is not None and len(odoo) else None
    out = []
    for r in docs.itertuples(index=False):
        s = sri.loc[int(r.idx_sri)] if pd.notna(r.idx_sri) and sri is not None else None
        o = od.loc[int(r.id_odoo)] if pd.notna(r.id_odoo) and od is not None else None
        g = lambda f, col: None if f is None else f.get(col)  # noqa: E731
        out.append(_fila(
            r.tipo, r.estado,
            _primero(g(s, "serie"), g(o, "num")),
            _primero(g(s, "fecha"), g(o, "fecha")),
            _primero(g(s, "proveedor"), g(o, "socio")),
            _primero(g(s, "ruc"), g(o, "ruc")),
            g(s, "total"), g(o, "total"), r.detalle))
    return out


def _filas_ret(retdocs):
    out = []
    for r in retdocs.itertuples(index=False):
        out.append(_fila("RET", r.estado, _primero(r.n_ret, r.doc), r.fecha, _primero(r.proveedor),
                         _primero(r.ruc), r.valor_sri, r.valor_odoo, r.detalle))
    return out


def _tarjeta(tipo, estados, con_sri):
    """`total` = "Comprobantes SRI" del Resumen del Excel (documentos con lado SRI, sin repetidos);
    los que existen solo en Odoo no cuentan en `total` y se ven en `por_estado`."""
    por_estado = {}
    for e in estados:
        por_estado[e] = por_estado.get(e, 0) + 1
    return {"tipo": tipo, "etiqueta": ETIQUETAS[tipo], "total": int(con_sri),
            "coinciden": por_estado.get(EST_OK, 0), "por_estado": por_estado}


def construir_tablero(modo, docs, sri, odoo, retdocs, periodo, advertencias) -> dict:
    """`docs` = clasificación por documento (estado, tipo, idx_sri, id_odoo, metodo, detalle); `sri`/`odoo`
    son los DataFrames a los que apuntan idx_sri / id_odoo; `retdocs` = resumir_por_documento(...) o None."""
    filas, tarjetas = [], []
    if docs is not None and len(docs):
        filas += _filas_docs(docs, sri, odoo)
        for t in TIPOS_DOC:
            d = docs[docs["tipo"] == t]
            est = [str(e) for e in d["estado"]]
            if est:
                # mismo criterio que la columna "Comprobantes SRI" del Resumen: lado SRI y no repetido
                con_sri = (d["idx_sri"].notna() & ~d["estado"].astype(str).isin(_REPETIDOS)
                           & ~d["estado"].astype(str).str.startswith("Informativo")).sum()
                tarjetas.append(_tarjeta(t, est, con_sri))
    if retdocs is not None and len(retdocs):
        filas += _filas_ret(retdocs)
        # Resumen: COUNTIFS(Lineas_SRI > 0) sobre Ret_Documentos
        con_sri = (pd.to_numeric(retdocs["n_lineas_sri"], errors="coerce").fillna(0) > 0).sum()
        tarjetas.append(_tarjeta("RET", [str(e) for e in retdocs["estado"]], con_sri))
    return {"modo": modo,
            "periodo": None if periodo is None else [_v(p) for p in periodo],
            "tarjetas": tarjetas,
            "excepciones": [f for f in filas if f["estado"] != EST_OK],
            "advertencias": list(advertencias or [])}
