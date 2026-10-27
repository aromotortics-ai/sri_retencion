# -*- coding: utf-8 -*-
"""Estados de conciliación de comprobantes EMITIDOS (facturas, NC, ND, liquidaciones y retenciones)
SRI vs Odoo.

Lógica pura sobre DataFrames. Reutiliza el motor de cruce de Compras (`conciliar_compras.conciliar`)
y el de retenciones (`conciliar_retenciones.conciliar_retenciones`); aquí solo se asigna el estado.
"""
import pandas as pd

from app.compras.conciliar_compras import CAT_DUP, _mismo_total, conciliar
from app.retenciones.conciliar_retenciones import EST_ANULADO, EST_DIF, EST_OK, EST_SOLO_ODOO, EST_SOLO_SRI
from app.retenciones.conciliar_retenciones import EST_INFO as EST_INFO_RET
from app.retenciones.conciliar_retenciones import conciliar_retenciones, _txt

EST_DIF_VALOR = EST_DIF  # "Diferencia de valor"
EST_DIF_DATOS = "Diferencia de datos"
EST_DUP = "Duplicado en Odoo"
EST_INFO = "Informativo"

__all__ = ["EST_OK", "EST_SOLO_SRI", "EST_SOLO_ODOO", "EST_DIF_VALOR", "EST_DIF_DATOS", "EST_ANULADO",
           "EST_DUP", "EST_INFO", "COLUMNAS", "clasificar_documentos", "conciliar_retenciones_emitidas"]

COLUMNAS = ["estado", "tipo", "idx_sri", "id_odoo", "metodo", "detalle"]


def _dinero(v):
    return f"{abs(float(v)):.2f}"


def _fecha(v):
    return pd.Timestamp(v).strftime("%d/%m") if pd.notna(v) else "sin fecha"


def _misma_base(a, b):
    """Base con tolerancia; si algún lado no la tiene no se compara."""
    return pd.isna(a) or pd.isna(b) or _mismo_total(a, b)


def _clasificar_pareja(s, o, cat):
    """(estado, detalle) de una pareja SRI-Odoo."""
    cancelado = o["estado"] == "cancel"
    if cat == CAT_DUP:
        if cancelado:  # el documento ya cruza con otro registro vigente; este quedó anulado
            return EST_INFO, "duplicado anulado en Odoo"
        return EST_DUP, "otro registro de Odoo ya cruza con este documento"
    if cancelado:
        return EST_ANULADO, "cancelado en Odoo, autorizado en el SRI"
    difs = []
    if not _mismo_total(s["total"], o["total"]):
        difs.append(f"total SRI {_dinero(s['total'])} ≠ Odoo {_dinero(o['total'])}")
    if not _misma_base(s["base"], o["base"]):
        difs.append(f"base SRI {_dinero(s['base'])} ≠ Odoo {_dinero(o['base'])}")
    if difs:
        return EST_DIF_VALOR, "; ".join(difs)
    datos = []
    clave_o = _txt(o["clave"])
    if clave_o is None:
        datos.append("sin clave de acceso en Odoo")
    elif clave_o != _txt(s["clave"]):
        datos.append("clave de acceso distinta")
    if pd.isna(s["fecha"]) or pd.isna(o["fecha"]) or \
            pd.Timestamp(s["fecha"]).date() != pd.Timestamp(o["fecha"]).date():
        datos.append(f"fecha SRI {_fecha(s['fecha'])} ≠ Odoo {_fecha(o['fecha'])}")
    if not bool(o["electronico"]):
        datos.append("no electrónico en Odoo")
    if datos:
        return EST_DIF_DATOS, "; ".join(datos)
    return EST_OK, ""


def _clasificar_solo_odoo(o):
    """(estado, detalle) de un registro de Odoo sin pareja dentro del período."""
    if o["estado"] == "cancel":
        return EST_INFO, "cancelado en Odoo, sin autorización en el SRI"
    if o["estado"] != "posted":
        return EST_INFO, "borrador en Odoo"
    if not bool(o["electronico"]):
        return EST_INFO, "no electrónico en Odoo"
    return EST_SOLO_ODOO, "publicado en Odoo, sin autorización en el SRI"


def _preferir_publicado(pares, por_id):
    """Si la pareja principal de un documento del SRI es un registro CANCELADO y otro registro
    PUBLICADO quedó como duplicado (p. ej. el cancelado tiene la clave y el publicado cruzó solo por
    número), el publicado pasa a ser la pareja principal y el cancelado el duplicado."""
    pares = pares.copy()
    ambos = pares["idx_sri"].notna() & pares["id_odoo"].notna()
    for _, g in pares[ambos].groupby("idx_sri"):
        principal = g[g["cat"] != CAT_DUP]
        if len(principal) != 1 or por_id[principal["id_odoo"].iloc[0]]["estado"] != "cancel":
            continue
        publicados = [i for i in g.index
                      if g.at[i, "cat"] == CAT_DUP and por_id[g.at[i, "id_odoo"]]["estado"] == "posted"]
        if publicados:
            pares.at[principal.index[0], "cat"] = CAT_DUP
            pares.at[publicados[0], "cat"] = ""
    return pares


def clasificar_documentos(sri: pd.DataFrame, odoo: pd.DataFrame, desde, hasta) -> pd.DataFrame:
    """Estado de cada documento emitido (FAC/NC/ND/LIQ): una fila por pareja o lado sin pareja.

    Columnas: estado, tipo, idx_sri (índice de `sri`), id_odoo, metodo, detalle.
    `sri` lleva `ruc` NA (el TXT no trae el RUC del cliente). Los registros publicados de Odoo se
    cruzan antes que los cancelados, para que un documento anulado y vuelto a emitir no tape al vigente.
    Los registros de Odoo sin pareja fuera de [desde, hasta] no se listan."""
    if odoo.empty:
        odoo_ord = odoo
    else:
        odoo_ord = odoo.iloc[(odoo["estado"] != "posted").argsort(kind="stable")]
    pares = conciliar(sri, odoo_ord, desde, hasta)
    por_id = {o["id"]: o for _, o in odoo.iterrows()}
    pares = _preferir_publicado(pares, por_id)

    filas = []
    for p in pares.itertuples(index=False):
        hay_sri, hay_odoo = pd.notna(p.idx_sri), pd.notna(p.id_odoo)
        if hay_sri and hay_odoo:
            estado, detalle = _clasificar_pareja(sri.loc[p.idx_sri], por_id[p.id_odoo], p.cat)
        elif hay_sri:
            estado, detalle = EST_SOLO_SRI, "sin registro en Odoo"
        else:
            estado, detalle = _clasificar_solo_odoo(por_id[p.id_odoo])
        filas.append((estado, p.tipo, p.idx_sri, p.id_odoo, p.metodo, detalle))
    out = pd.DataFrame(filas, columns=COLUMNAS)
    out["idx_sri"] = out["idx_sri"].astype("Int64")
    out["id_odoo"] = out["id_odoo"].astype("Int64")
    return out


def conciliar_retenciones_emitidas(sri_lineas, odoo_ret, sri_txt, desde, hasta) -> pd.DataFrame:
    """Igual a `conciliar_retenciones`, pero una retención cancelada en Odoo y autorizada en el SRI
    queda como EST_ANULADO (en lugar de "Solo en SRI" más una fila informativa de Odoo)."""
    ret = conciliar_retenciones(sri_lineas, odoo_ret, sri_txt, desde, hasta)
    if ret.empty or odoo_ret is None or odoo_ret.empty:
        return ret
    cancel = odoo_ret[odoo_ret["estado"].fillna("") == "cancel"]
    ids_cancel = {}
    for n, id_ret in zip(cancel["n_ret"], cancel["id_ret"]):
        if _txt(n) is not None:
            ids_cancel.setdefault(_txt(n), id_ret)
    if not ids_cancel:
        return ret
    n_ret = ret["n_ret"].map(_txt)
    anular = (ret["estado"] == EST_SOLO_SRI) & n_ret.isin(ids_cancel)
    if not anular.any():
        return ret
    ret = ret.copy()
    ret.loc[anular, "estado"] = EST_ANULADO
    ret.loc[anular, "id_ret"] = [ids_cancel[n] for n in n_ret[anular]]
    anulados = set(n_ret[anular])
    duplicadas = (ret["estado"] == EST_INFO_RET) & n_ret.isin(anulados)
    return ret[~duplicadas].reset_index(drop=True)
