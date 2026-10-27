# -*- coding: utf-8 -*-
import json
from datetime import date

import pandas as pd
import pytest

import app.compras.excel_compras as excel_compras
import app.retenciones.conciliar_retenciones as cr
import app.ventas.excel_emitidos as excel_emitidos
from app.compras.clasificar import clasificar_compras
from app.compras.excel_compras import construir_excel
from app.reportes.tablero import construir_tablero
from app.ventas.conciliar_emitidos import clasificar_documentos
from app.ventas.excel_emitidos import construir_excel_emitidos

DESDE, HASTA = date(2026, 9, 1), date(2026, 9, 30)
F = pd.Timestamp("2026-09-05")
SRI_V = ["clave", "serie", "fecha", "tipo", "base", "iva", "total", "ruc", "proveedor", "afectada"]
ODOO_V = ["id", "num", "nfac", "socio", "ruc", "fecha", "tipo", "clave", "base", "iva", "total",
          "afectada", "estado", "electronico", "elegible"]


@pytest.fixture(autouse=True)
def _sin_libreoffice(monkeypatch):
    monkeypatch.setattr(excel_emitidos, "_recalcular", lambda ruta: False)
    monkeypatch.setattr(excel_compras, "_recalcular", lambda ruta: False, raising=False)


# ---------- Ventas ----------
def _sri_v(*filas):
    base = {"clave": "1" * 49, "serie": "001-001-000000100", "fecha": F, "tipo": "FAC", "base": 100.0,
            "iva": 15.0, "total": 115.0, "ruc": pd.NA, "proveedor": pd.NA, "afectada": pd.NA}
    return pd.DataFrame([{**base, **f} for f in filas], columns=SRI_V)


def _odoo_v(*filas):
    base = {"id": 1, "num": "FAC/2026/0100", "nfac": "001-001-000000100", "socio": "CLIENTE SA",
            "ruc": "1790000000001", "fecha": F, "tipo": "FAC", "clave": "1" * 49, "base": 100.0,
            "iva": 15.0, "total": 115.0, "afectada": pd.NA, "estado": "posted", "electronico": True,
            "elegible": True}
    return pd.DataFrame([{**base, **f} for f in filas], columns=ODOO_V)


def _datos_v():
    sri = _sri_v(
        {},
        {"clave": "2" * 49, "serie": "001-001-000000101", "total": 120.0},
        {"clave": "3" * 49, "serie": "001-001-000000102"},
        {"clave": "4" * 49, "serie": "001-001-000000103"},
        {"clave": "5" * 49, "serie": "001-002-000000001", "tipo": "NC", "base": 10.0, "iva": 1.5,
         "total": 11.5},
    )
    odoo = _odoo_v(
        {},
        {"id": 2, "num": "FAC/2026/0101", "nfac": "001-001-000000101", "clave": "2" * 49},
        {"id": 4, "num": "FAC/2026/0103", "nfac": "001-001-000000103", "clave": "4" * 49, "estado": "cancel"},
        {"id": 5, "num": "FAC/2026/0104", "nfac": "001-001-000000104", "clave": "6" * 49},
        {"id": 6, "num": "NC/2026/0001", "nfac": "001-002-000000001", "clave": "5" * 49, "tipo": "NC",
         "base": 10.0, "iva": 1.5, "total": 11.5},
    )
    return sri, odoo


def _fila_ret(estado, n_ret, vs=10.0, vo=10.0, factura="001-001-000000010"):
    return {"estado": estado, "n_ret": n_ret, "fecha": F, "ruc": "1790000000001", "proveedor": "CLIENTE SA",
            "factura_sri": factura, "factura_odoo": factura, "pct_sri": 1.0,
            "pct_odoo": 1.0, "base_sri": 1000.0, "base_odoo": 1000.0, "valor_sri": vs, "valor_odoo": vo,
            "dif_valor": None, "id_ret": 7, "clave": "9" * 49}


def _ret():
    df = pd.DataFrame([_fila_ret(cr.EST_OK, "001-001-000000001"),
                       _fila_ret(cr.EST_OK, "001-001-000000001", factura="001-001-000000011"),
                       _fila_ret(cr.EST_SOLO_SRI, "001-001-000000002", vo=None),
                       _fila_ret(cr.EST_INFO, "001-001-000000003", vs=None)], columns=cr.COLUMNAS)
    df["id_ret"] = df["id_ret"].astype("Int64")
    return df


def test_tarjetas_cuentan_documentos_de_retenciones():
    ret = _ret()
    retdocs = cr.resumir_por_documento(ret)
    assert len(ret) == 4 and len(retdocs) == 3
    t = construir_tablero("ventas", pd.DataFrame(columns=["estado", "tipo", "idx_sri", "id_odoo", "metodo",
                                                          "detalle"]),
                          _sri_v(), _odoo_v().iloc[0:0], retdocs, [DESDE.isoformat(), HASTA.isoformat()], [])
    assert [x["tipo"] for x in t["tarjetas"]] == ["RET"]
    r = t["tarjetas"][0]
    assert r["etiqueta"] == "Retenciones" and r["total"] == 3 and r["coinciden"] == 1
    assert r["por_estado"] == {cr.EST_OK: 1, cr.EST_SOLO_SRI: 1, cr.EST_INFO: 1}


def test_excepciones_excluyen_coincide_e_incluyen_informativo():
    sri, odoo = _datos_v()
    odoo.loc[odoo["id"] == 5, "estado"] = "draft"  # borrador -> Informativo
    docs = clasificar_documentos(sri, odoo, DESDE, HASTA)
    t = construir_tablero("ventas", docs, sri, odoo, cr.resumir_por_documento(_ret()),
                          [DESDE.isoformat(), HASTA.isoformat()], ["aviso"])
    ex = t["excepciones"]
    assert all(e["estado"] != "Coincide" for e in ex)
    info = [e for e in ex if e["estado"].startswith("Informativo")]
    assert info and all(e["accionable"] is False for e in info)
    assert any(e["accionable"] for e in ex)
    assert {e["tipo"] for e in ex} >= {"FAC", "RET"}
    assert t["modo"] == "ventas" and t["advertencias"] == ["aviso"]
    # campos de presentacion
    dif = [e for e in ex if e["estado"] == "Diferencia de valor" and e["tipo"] == "FAC"][0]
    assert dif["documento"] == "001-001-000000101" and dif["contraparte"] == "CLIENTE SA"
    assert dif["valor_sri"] == 120.0 and dif["valor_odoo"] == 115.0 and dif["diferencia"] == 5.0
    assert dif["fecha"] == "2026-09-05" and dif["que_corregir"]
    json.dumps(t)  # serializable


def test_tablero_vacio_es_serializable():
    t = construir_tablero("compras", None, None, None, None, None, [])
    assert t["tarjetas"] == [] and t["excepciones"] == [] and t["periodo"] is None
    json.dumps(t)


# ---------- Compras ----------
def _datos_c():
    sri = pd.DataFrame([
        ["1" * 49, "001-001-000000001", "1790000000001", "PROV SA", F, "FAC", 100.0, 15.0, 115.0, pd.NA],
        ["2" * 49, "001-001-000000002", "1790000000001", "OTRO SA", F, "FAC", 200.0, 30.0, 240.0, pd.NA],
        ["3" * 49, "001-001-000000003", "1790000000001", "OTRO SA", F, "NC", 10.0, 1.5, 11.5, pd.NA],
    ], columns=["clave", "serie", "ruc", "proveedor", "fecha", "tipo", "base", "iva", "total", "afectada"])
    odoo = pd.DataFrame([
        [10, "BILL/1", "001-001-000000001", "PROV SA", "1790000000001", F, "FAC", "01", "1" * 49,
         100.0, 15.0, 115.0, pd.NA, True],
        [11, "BILL/2", "001-001-000000002", "OTRO SA", "1790000000001", F, "FAC", "01", "2" * 49,
         200.0, 30.0, 230.0, pd.NA, True],
    ], columns=["id", "num", "nfac", "socio", "ruc", "fecha", "tipo", "cod_comp", "clave", "base", "iva",
                "total", "afectada", "elegible"])
    pares = pd.DataFrame({
        "idx_sri": pd.array([0, 1, 2], dtype="Int64"),
        "id_odoo": pd.array([10, 11, None], dtype="Int64"),
        "tipo": ["FAC", "FAC", "NC"], "metodo": ["Clave de acceso", "Clave de acceso", ""],
        "cat": ["", "", ""]})
    return sri, odoo, pares


def test_cifras_tablero_igualan_excel_completo(tmp_path):
    # Ventas
    sri, odoo = _datos_v()
    docs = clasificar_documentos(sri, odoo, DESDE, HASTA)
    ret = _ret()
    res = construir_excel_emitidos(sri, odoo, docs, ret, DESDE, HASTA, str(tmp_path / "v.xlsx"))
    t = construir_tablero("ventas", docs, sri, odoo, cr.resumir_por_documento(ret), None, [])
    tarj = {x["tipo"]: x for x in t["tarjetas"]}
    assert set(tarj) == {"FAC", "NC", "RET"}
    for tipo, c in tarj.items():
        esperado = res["documentos"][tipo]
        assert c["total"] == esperado["total"], tipo
        for estado, n in c["por_estado"].items():
            assert esperado[estado] == n, (tipo, estado)
        assert c["coinciden"] == esperado.get("Coincide", 0)
        assert sum(c["por_estado"].values()) == c["total"]
    # Compras
    sri, odoo, pares = _datos_c()
    por_tipo = construir_excel(sri, odoo, pares, DESDE, HASTA, str(tmp_path / "c.xlsx"))
    docs = clasificar_compras(sri, odoo, pares)
    t = construir_tablero("compras", docs, sri, odoo, None, None, [])
    tarj = {x["tipo"]: x for x in t["tarjetas"]}
    assert set(tarj) == {"FAC", "NC"}
    for tipo, c in tarj.items():
        e = por_tipo[tipo]
        assert c["coinciden"] == e["coinciden"]
        assert c["por_estado"].get("Solo en SRI", 0) == e["solo_sri"]
        assert c["por_estado"].get("Solo en Odoo", 0) == e["solo_odoo"]
        assert c["por_estado"].get("Diferencia de valor", 0) == e["diferencias"]
        assert c["por_estado"].get("Informativo", 0) == e["informativos"]
    # contraparte de Compras: proveedor del SRI
    ex = [x for x in t["excepciones"] if x["estado"] == "Diferencia de valor"][0]
    assert ex["contraparte"] == "OTRO SA" and ex["diferencia"] == 10.0
    json.dumps(t)
