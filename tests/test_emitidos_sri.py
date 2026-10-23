# -*- coding: utf-8 -*-
import datetime as dt

import pandas as pd
import pytest

from emitidos_sri import RUC_EMPRESA, leer_emitidos_txts, periodo_emitidos

COLS = ["COMPROBANTE", "SERIE_COMPROBANTE", "CLAVE_ACCESO", "FECHA_AUTORIZACION", "FECHA_EMISION",
        "VALOR_SIN_IMPUESTOS", "IVA", "IMPORTE_TOTAL"]
COLS_DOCS = ["clave", "serie", "fecha", "tipo", "base", "iva", "total", "ruc", "proveedor", "afectada"]


def _clave(n):
    return str(n).zfill(49)


def _txt(tmp_path, nombre, filas, cols=COLS):
    p = tmp_path / nombre
    lineas = ["\t".join(cols)] + ["\t".join(f) for f in filas]
    p.write_text("\n".join(lineas) + "\n", encoding="iso-8859-1")
    return str(p)


def _fila(comp, serie, clave, fecha="01/09/2026", base="10.00", iva="1.50", total="11.50"):
    return [comp, serie, clave, fecha, fecha, base, iva, total]


def test_separa_tipos_por_columna_comprobante(tmp_path):
    ruta = _txt(tmp_path, "2390000831001_EmitidosAutorizados (1).txt", [
        _fila("Factura", "001-001-000000001", _clave(1)),
        _fila("Nota de Crédito", "001-001-000000002", _clave(2), total=""),
        _fila("Nota de Débito", "001-001-000000003", _clave(3)),
        _fila("Liquidación de Compra", "001-001-000000004", _clave(4)),
        _fila("Comprobante de Retención", "001-001-000000005", _clave(5), base="", iva="", total=""),
    ])
    docs, ret, adv = leer_emitidos_txts([ruta])
    assert set(docs["tipo"]) == {"FAC", "NC", "ND", "LIQ"}
    assert list(docs.columns) == COLS_DOCS
    assert len(ret) == 1
    assert ret.iloc[0]["ruc"] == RUC_EMPRESA
    assert list(ret.columns) == ["clave", "serie", "ruc", "proveedor", "fecha"]


def test_nc_total_vacio_es_base_mas_iva(tmp_path):
    ruta = _txt(tmp_path, "a.txt", [
        _fila("Nota de Crédito", "001-001-000000002", _clave(2), base="7.72", iva="1.16", total="")])
    docs, _, _ = leer_emitidos_txts([ruta])
    assert docs.iloc[0]["total"] == pytest.approx(8.88)


def test_fecha_con_hora(tmp_path):
    ruta = _txt(tmp_path, "a.txt", [
        _fila("Factura", "001-001-000000001", _clave(1), fecha="01/09/2026 00:00:00")])
    docs, _, _ = leer_emitidos_txts([ruta])
    assert docs.iloc[0]["fecha"] == pd.Timestamp(2026, 9, 1)


def test_duplicados_entre_txt(tmp_path):
    a = _txt(tmp_path, "a.txt", [_fila("Factura", "001-001-000000001", _clave(1))])
    b = _txt(tmp_path, "b.txt", [_fila("Factura", "001-001-000000001", _clave(1))])
    docs, _, adv = leer_emitidos_txts([a, b])
    assert len(docs) == 1
    assert adv == []
    c = _txt(tmp_path, "c.txt", [_fila("Factura", "001-001-000000001", _clave(1), total="99.00")])
    docs, _, adv = leer_emitidos_txts([a, c])
    assert len(docs) == 1
    assert any(_clave(1) in m for m in adv)


def test_guia_remision_se_ignora_con_advertencia(tmp_path):
    ruta = _txt(tmp_path, "a.txt", [
        _fila("Factura", "001-001-000000001", _clave(1)),
        _fila("Guía de Remisión", "001-001-000000009", _clave(9)),
        _fila("Guía de Remisión", "001-001-000000010", _clave(10)),
    ])
    docs, ret, adv = leer_emitidos_txts([ruta])
    assert list(docs["tipo"]) == ["FAC"]
    assert any("2" in m and "gu" in m.lower() for m in adv)


def test_faltan_columnas(tmp_path):
    ruta = _txt(tmp_path, "sinclave.txt", [["Factura", "x", "01/09/2026"]],
                cols=["COMPROBANTE", "SERIE_COMPROBANTE", "FECHA_EMISION"])
    with pytest.raises(ValueError) as e:
        leer_emitidos_txts([ruta])
    assert "sinclave.txt" in str(e.value)
    assert "CLAVE_ACCESO" in str(e.value)


def test_solo_retenciones(tmp_path):
    ruta = _txt(tmp_path, "a.txt", [
        _fila("Comprobante de Retención", "001-001-000000005", _clave(5), base="", iva="", total="")])
    docs, ret, adv = leer_emitidos_txts([ruta])
    assert docs.empty
    assert list(docs.columns) == COLS_DOCS
    assert len(ret) == 1


def test_retenciones_sin_clave_valida_se_ignoran_con_advertencia(tmp_path):
    ret_fila = lambda serie, clave: _fila("Comprobante de Retención", serie, clave, base="", iva="",
                                          total="")
    ruta = _txt(tmp_path, "a.txt", [
        ret_fila("001-001-000000005", _clave(5)),
        ret_fila("001-001-000000006", ""),
        ret_fila("001-001-000000007", "12345"),
    ])
    _, ret, adv = leer_emitidos_txts([ruta])
    assert len(ret) == 1
    assert ("2 retenciones del TXT se ignoraron por no tener clave de acceso válida") in adv
    _, _, adv = leer_emitidos_txts([_txt(tmp_path, "b.txt", [ret_fila("001-001-000000005", _clave(5))])])
    assert not any("clave de acceso válida" in a for a in adv)


def test_periodo_rango_largo():
    fechas = pd.Series([pd.Timestamp(2026, 9, d) for d in range(1, 16)] + [pd.Timestamp(2026, 7, 20)])
    desde, hasta, adv = periodo_emitidos(fechas)
    assert desde == dt.date(2026, 7, 20)
    assert hasta == dt.date(2026, 9, 15)
    assert adv
    _, _, adv2 = periodo_emitidos(fechas[:15])
    assert adv2 == []
