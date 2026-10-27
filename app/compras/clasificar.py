# -*- coding: utf-8 -*-
"""Estados por documento de Compras calculados en Python (misma regla que la columna Resultado del Excel)."""
import pandas as pd

from app.compras.conciliar_compras import CAT_DUP, CAT_INFO, CAT_SOLO_ODOO, _mismo_total

COLUMNAS = ["estado", "tipo", "idx_sri", "id_odoo", "metodo", "detalle"]


def _dinero(v):
    return f"{float(v):.2f}"


def clasificar_compras(sri: pd.DataFrame, odoo: pd.DataFrame, pares: pd.DataFrame) -> pd.DataFrame:
    """Una fila por fila de `pares` (mismo orden) con estado y detalle."""
    od = odoo.set_index("id")
    filas = []
    for r in pares.itertuples():
        hay_sri, hay_odoo = pd.notna(r.idx_sri), pd.notna(r.id_odoo)
        detalle = ""
        if r.cat == CAT_DUP:
            estado = "Duplicado en Odoo"
        elif hay_sri and hay_odoo:
            ts, to = sri.at[int(r.idx_sri), "total"], od.at[int(r.id_odoo), "total"]
            if _mismo_total(ts, to):
                estado = "Coincide"
            else:
                estado = "Diferencia de valor"
                detalle = f"total SRI {_dinero(ts)} ≠ Odoo {_dinero(to)}"
        elif hay_sri:
            estado = "Solo en SRI"
        elif r.cat == CAT_SOLO_ODOO:
            estado = "Solo en Odoo"
        elif r.cat == CAT_INFO:
            estado = "Informativo"
        else:
            estado = r.cat
        filas.append((estado, r.tipo, r.idx_sri, r.id_odoo, r.metodo, detalle))
    out = pd.DataFrame(filas, columns=COLUMNAS)
    out["idx_sri"] = out["idx_sri"].astype("Int64")
    out["id_odoo"] = out["id_odoo"].astype("Int64")
    return out
