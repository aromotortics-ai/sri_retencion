# -*- coding: utf-8 -*-
from datetime import date

import pandas as pd

from app.compras.clasificar import clasificar_compras
from app.compras.conciliar_compras import CAT_DUP, CAT_INFO, CAT_SOLO_ODOO
from app.compras.excel_compras import construir_excel
from app.ventas.conciliar_emitidos import COLUMNAS

F = pd.Timestamp("2026-09-05")
SRI_COLS = ["clave", "serie", "ruc", "proveedor", "fecha", "tipo", "base", "iva", "total", "afectada"]
ODOO_COLS = ["id", "num", "nfac", "socio", "ruc", "fecha", "tipo", "cod_comp", "clave", "base", "iva",
             "total", "afectada", "elegible"]


def _datos():
    def s(n, total):
        return [str(n) * 49, f"001-001-00000000{n}", "1790000000001", "PROV", F, "FAC", total, 0.0, total, pd.NA]

    def o(i, total, elegible=True):
        return [i, f"BILL/{i}", "x", "PROV", "1790000000001", F, "FAC", "01", None, total, 0.0, total,
                pd.NA, elegible]

    sri = pd.DataFrame([s(1, 100.0), s(2, 100.0), s(3, 100.0), s(4, 50.0), s(5, 70.0)], columns=SRI_COLS)
    odoo = pd.DataFrame([o(10, 100.0), o(11, 100.5), o(12, 99.98), o(13, 30.0), o(14, 5.0, False),
                         o(15, 70.0), o(16, 70.0)], columns=ODOO_COLS)
    filas = [(0, 10, "FAC", "Clave de acceso", ""), (1, 11, "FAC", "Clave de acceso", ""),
             (2, 12, "FAC", "Clave de acceso", ""), (3, None, "FAC", "", ""),
             (None, 13, "FAC", "", CAT_SOLO_ODOO), (None, 14, "FAC", "", CAT_INFO),
             (4, 15, "FAC", "Clave de acceso", ""), (4, 16, "FAC", "N° documento + importe", CAT_DUP)]
    pares = pd.DataFrame(filas, columns=["idx_sri", "id_odoo", "tipo", "metodo", "cat"])
    pares["idx_sri"] = pares["idx_sri"].astype("Int64")
    pares["id_odoo"] = pares["id_odoo"].astype("Int64")
    return sri, odoo, pares


def test_clasificar_compras_replica_el_excel():
    sri, odoo, pares = _datos()
    r = clasificar_compras(sri, odoo, pares)
    assert list(r.columns) == COLUMNAS
    assert list(r["estado"]) == ["Coincide", "Diferencia de valor", "Coincide", "Solo en SRI",
                                 "Solo en Odoo", "Informativo", "Coincide", "Duplicado en Odoo"]
    assert r.loc[1, "detalle"] == "total SRI 100.00 ≠ Odoo 100.50"
    assert r.loc[0, "detalle"] == ""


def test_conteos_igualan_construir_excel(tmp_path):
    sri, odoo, pares = _datos()
    r = clasificar_compras(sri, odoo, pares)
    res = construir_excel(sri, odoo, pares, date(2026, 9, 1), date(2026, 9, 30), str(tmp_path / "o.xlsx"))
    q = r[r["tipo"] == "FAC"]
    cuenta = q["estado"].value_counts()
    esperado = res["FAC"]
    assert cuenta.get("Coincide", 0) == esperado["coinciden"] == 3
    assert cuenta.get("Diferencia de valor", 0) == esperado["diferencias"] == 1
    assert cuenta.get("Solo en SRI", 0) == esperado["solo_sri"] == 1
    assert cuenta.get("Solo en Odoo", 0) == esperado["solo_odoo"] == 1
    assert cuenta.get("Informativo", 0) == esperado["informativos"] == 1
