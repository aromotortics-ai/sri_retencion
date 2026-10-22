from datetime import date

import pandas as pd
import pytest

from conciliar_retenciones import (EST_DIF, EST_FAC, EST_INFO, EST_OK, EST_SOLO_ODOO, EST_SOLO_SRI,
                                   EST_XML, conciliar_retenciones, leer_sri_retenciones_txt)

SRI_COLS = ["clave", "n_ret", "fecha", "ruc", "proveedor", "factura", "cod", "pct", "base", "valor"]
ODOO_COLS = ["id_ret", "n_ret", "fecha", "estado", "socio", "ruc", "factura", "pct", "base", "valor",
             "cod"]
TXT_COLS = ["clave", "serie", "ruc", "proveedor", "fecha"]
SALIDA = ["estado", "n_ret", "fecha", "ruc", "proveedor", "factura_sri", "factura_odoo", "pct_sri",
          "pct_odoo", "base_sri", "base_odoo", "valor_sri", "valor_odoo", "dif_valor", "id_ret",
          "clave"]

DESDE, HASTA = date(2026, 9, 1), date(2026, 9, 30)
CLAVE = "1" * 49


def _sri(*filas):
    base = {"clave": CLAVE, "n_ret": "001-200-000004340", "fecha": pd.Timestamp("2026-09-10"),
            "ruc": "1790000000001", "proveedor": "CLIENTE SA", "factura": "001-003-000088766",
            "cod": "312", "pct": 2.0, "base": 580.0, "valor": 11.60}
    return pd.DataFrame([{**base, **f} for f in filas], columns=SRI_COLS)


def _odoo(*filas):
    base = {"id_ret": 10, "n_ret": "001-200-000004340", "fecha": pd.Timestamp("2026-09-10"),
            "estado": "posted", "socio": "CLIENTE SA", "ruc": "1790000000001",
            "factura": "001-003-000088766", "pct": 2.0, "base": 580.0, "valor": 11.60, "cod": "312"}
    return pd.DataFrame([{**base, **f} for f in filas], columns=ODOO_COLS)


def _txt_df(*series):
    return pd.DataFrame([{"clave": str(i) * 49, "serie": s, "ruc": "1790000000001",
                          "proveedor": "CLIENTE SA", "fecha": pd.Timestamp("2026-09-10")}
                         for i, s in enumerate(series, 1)], columns=TXT_COLS)


def _vacio_sri():
    return pd.DataFrame(columns=SRI_COLS)


def _vacio_odoo():
    return pd.DataFrame(columns=ODOO_COLS)


def _run(sri, odoo, txt=None):
    txt = _txt_df("001-200-000004340") if txt is None else txt
    out = conciliar_retenciones(sri, odoo, txt, DESDE, HASTA)
    assert list(out.columns) == SALIDA
    return out


def test_coincide_dentro_de_tolerancia():
    out = _run(_sri({"valor": 11.60}), _odoo({"valor": 11.58}))
    assert len(out) == 1
    r = out.iloc[0]
    assert r["estado"] == EST_OK
    assert r["valor_sri"] == pytest.approx(11.60) and r["valor_odoo"] == pytest.approx(11.58)
    assert r["id_ret"] == 10 and r["clave"] == CLAVE


def test_diferencia_de_valor_fuera_de_tolerancia():
    out = _run(_sri({"valor": 11.60}), _odoo({"valor": 11.57}))
    assert len(out) == 1
    r = out.iloc[0]
    assert r["estado"] == EST_DIF
    assert r["dif_valor"] == 0.03


def test_solo_en_sri():
    out = _run(_sri({}), _vacio_odoo())
    assert list(out["estado"]) == [EST_SOLO_SRI]
    assert pd.isna(out.iloc[0]["factura_odoo"]) and pd.isna(out.iloc[0]["id_ret"])
    assert out.iloc[0]["factura_sri"] == "001-003-000088766"


def test_solo_en_odoo():
    # La retención de Odoo no figura en el TXT del SRI: es "Solo en Odoo".
    out = _run(_vacio_sri(), _odoo({"n_ret": "001-200-000009999"}))
    assert list(out["estado"]) == [EST_SOLO_ODOO]
    r = out.iloc[0]
    assert r["factura_odoo"] == "001-003-000088766" and pd.isna(r["clave"])
    assert r["proveedor"] == "CLIENTE SA"


def test_factura_incorrecta_una_sola_fila():
    out = _run(_sri({"factura": "001-003-000088766"}), _odoo({"factura": "002-003-000088766"}))
    assert len(out) == 1
    r = out.iloc[0]
    assert r["estado"] == EST_FAC
    assert r["factura_sri"] == "001-003-000088766"
    assert r["factura_odoo"] == "002-003-000088766"


def test_porcentaje_distinto_es_diferencia_de_valor():
    out = _run(_sri({"pct": 2.0, "valor": 11.60}), _odoo({"pct": 1.0, "valor": 5.80}))
    assert len(out) == 1
    r = out.iloc[0]
    assert r["estado"] == EST_DIF
    assert r["pct_sri"] == 2.0 and r["pct_odoo"] == 1.0
    assert r["dif_valor"] == pytest.approx(5.80)


def test_porcentaje_distinto_empareja_por_valor_cercano():
    sri = _sri({"pct": 2.0, "valor": 11.60}, {"pct": 30.0, "cod": "725", "base": 87.0,
                                              "valor": 26.10})
    odoo = _odoo({"pct": 1.0, "valor": 5.80}, {"pct": 70.0, "base": 87.0, "valor": 60.90})
    out = _run(sri, odoo)
    assert list(out["estado"]) == [EST_DIF, EST_DIF]
    pares = sorted(zip(out["pct_sri"], out["pct_odoo"]))
    assert pares == [(2.0, 1.0), (30.0, 70.0)]


def test_xml_no_disponible_no_es_solo_odoo():
    out = _run(_vacio_sri(), _odoo({}), _txt_df("001-200-000004340"))
    assert list(out["estado"]) == [EST_XML]
    assert out.iloc[0]["clave"] == "1" * 49


def test_xml_no_disponible_mismo_n_ret_otro_cliente():
    # A y B usan el mismo n_ret (secuencias propias). El XML de A se descargó, el de B falló.
    ruc_a, ruc_b = "1790000000001", "0990000000001"
    clave_a, clave_b = "1" * 49, "2" * 49
    txt = pd.DataFrame([
        {"clave": clave_a, "serie": "001-200-000004340", "ruc": ruc_a, "proveedor": "A",
         "fecha": pd.Timestamp("2026-09-10")},
        {"clave": clave_b, "serie": "001-200-000004340", "ruc": ruc_b, "proveedor": "B",
         "fecha": pd.Timestamp("2026-09-10")},
    ], columns=TXT_COLS)
    sri = _sri({"clave": clave_a, "ruc": ruc_a})
    odoo = _odoo({"id_ret": 10, "ruc": ruc_a},
                 {"id_ret": 20, "ruc": ruc_b, "socio": "B", "factura": "005-001-000000777"})
    out = _run(sri, odoo, txt)
    por_id = {int(r["id_ret"]): r for _, r in out.iterrows()}
    assert len(out) == 2
    assert por_id[10]["estado"] == EST_OK
    assert por_id[20]["estado"] == EST_XML
    assert por_id[20]["clave"] == clave_b


def test_xml_no_disponible_odoo_sin_ruc_usa_solo_n_ret():
    out = _run(_vacio_sri(), _odoo({"ruc": pd.NA}), _txt_df("001-200-000004340"))
    assert list(out["estado"]) == [EST_XML]
    assert out.iloc[0]["clave"] == "1" * 49


def test_odoo_sin_ruc_con_xml_descargado_es_solo_odoo():
    # El único comprobante del TXT con ese n_ret sí tiene XML: no hay XML faltante.
    out = _run(_sri({"factura": "009-009-000000009"}), _odoo({"ruc": pd.NA}),
               _txt_df("001-200-000004340"))
    assert sorted(out["estado"]) == sorted([EST_SOLO_SRI, EST_SOLO_ODOO])


def test_odoo_cancelada_es_informativa():
    odoo = _odoo({"estado": "cancel", "n_ret": "001-200-000009999"},
                 {"estado": "draft", "id_ret": 11, "n_ret": "001-200-000004340"})
    out = _run(_sri({}), odoo)
    estados = sorted(out["estado"])
    assert estados == sorted([EST_INFO, EST_INFO, EST_SOLO_SRI])
    assert EST_SOLO_ODOO not in set(out["estado"])
    assert EST_OK not in set(out["estado"])


def test_lineas_repetidas_se_suman():
    odoo = _odoo({"base": 290.0, "valor": 5.80}, {"base": 290.0, "valor": 5.80})
    out = _run(_sri({"base": 580.0, "valor": 11.60}), odoo)
    assert len(out) == 1
    r = out.iloc[0]
    assert r["estado"] == EST_OK
    assert r["valor_odoo"] == pytest.approx(11.60) and r["base_odoo"] == pytest.approx(580.0)


def test_lineas_repetidas_en_sri_se_suman():
    sri = _sri({"base": 290.0, "valor": 5.80}, {"base": 290.0, "valor": 5.80})
    out = _run(sri, _odoo({}))
    assert list(out["estado"]) == [EST_OK]
    assert out.iloc[0]["valor_sri"] == pytest.approx(11.60)


def test_odoo_fuera_de_periodo_sin_par_no_genera_fila():
    odoo = _odoo({"n_ret": "001-200-000009999", "fecha": pd.Timestamp("2026-08-28")})
    out = _run(_vacio_sri(), odoo)
    assert out.empty


def test_odoo_fuera_de_periodo_con_par_si_genera_fila():
    out = _run(_sri({}), _odoo({"fecha": pd.Timestamp("2026-08-28")}))
    assert list(out["estado"]) == [EST_OK]


def test_sri_fuera_de_periodo_si_genera_fila():
    out = _run(_sri({"fecha": pd.Timestamp("2026-10-02")}), _vacio_odoo())
    assert list(out["estado"]) == [EST_SOLO_SRI]


def test_odoo_sin_factura_no_se_pierde():
    odoo = _odoo({"n_ret": "001-200-000009999", "factura": pd.NA, "ruc": pd.NA},
                 {"n_ret": "001-200-000009998", "factura": pd.NA, "ruc": pd.NA, "id_ret": 11})
    out = _run(_vacio_sri(), odoo)
    assert sorted(out["estado"]) == [EST_SOLO_ODOO, EST_SOLO_ODOO]
    assert sorted(out["id_ret"]) == [10, 11]
    assert out["factura_odoo"].isna().all()


def test_odoo_sin_factura_no_cruza_con_sri_sin_factura():
    out = _run(_sri({"factura": pd.NA}), _odoo({"factura": pd.NA}))
    assert sorted(out["estado"]) == sorted([EST_SOLO_SRI, EST_SOLO_ODOO])


@pytest.mark.parametrize("sri_vacio,odoo_vacio", [(True, True), (True, False), (False, True)])
def test_entradas_vacias(sri_vacio, odoo_vacio):
    sri = _vacio_sri() if sri_vacio else _sri({})
    odoo = _vacio_odoo() if odoo_vacio else _odoo({})
    out = conciliar_retenciones(sri, odoo, pd.DataFrame(columns=TXT_COLS), DESDE, HASTA)
    assert list(out.columns) == SALIDA
    assert len(out) == (0 if sri_vacio and odoo_vacio else 1)


# ---------------------------------------------------------------- leer_sri_retenciones_txt

TXT_COLUMNAS = ["RUC_EMISOR", "RAZON_SOCIAL_EMISOR", "TIPO_COMPROBANTE", "SERIE_COMPROBANTE",
                "CLAVE_ACCESO", "FECHA_AUTORIZACION", "FECHA_EMISION", "IDENTIFICACION_RECEPTOR",
                "VALOR_SIN_IMPUESTOS", "IVA", "IMPORTE_TOTAL", "NUMERO_DOCUMENTO_MODIFICADO"]


def _txt(path, filas):
    lineas = ["\t".join(TXT_COLUMNAS)]
    for f in filas:
        d = dict(zip(TXT_COLUMNAS, [""] * len(TXT_COLUMNAS)))
        d.update(f)
        lineas.append("\t".join(d[c] for c in TXT_COLUMNAS))
    path.write_bytes("\n".join(lineas).encode("iso-8859-1"))
    return str(path)


def _fila(clave, tipo="Comprobante de Retención", serie="001-158-003714396", fecha="05/09/2026"):
    return {"RUC_EMISOR": "1790000000001", "RAZON_SOCIAL_EMISOR": "CLIENTE ÑANDÚ SA",
            "TIPO_COMPROBANTE": tipo, "SERIE_COMPROBANTE": serie, "CLAVE_ACCESO": clave,
            "FECHA_EMISION": fecha}


def test_leer_sri_retenciones_txt(tmp_path):
    p = _txt(tmp_path / "r.txt", [_fila("1" * 49),
                                  _fila("2" * 49, serie="001-158-003714397"),
                                  _fila("3" * 49, tipo="Factura", serie="001-001-000000001")])
    df = leer_sri_retenciones_txt([p])
    assert list(df.columns) == ["clave", "serie", "ruc", "proveedor", "fecha"]
    assert len(df) == 2
    assert list(df["serie"]) == ["001-158-003714396", "001-158-003714397"]
    assert df["proveedor"].iloc[0] == "CLIENTE ÑANDÚ SA"
    assert df["fecha"].iloc[0] == pd.Timestamp("2026-09-05")


def test_leer_sri_retenciones_txt_deduplica_y_filtra_claves(tmp_path):
    p1 = _txt(tmp_path / "a.txt", [_fila("1" * 49), _fila("12345"), _fila("ABC" + "1" * 46)])
    p2 = _txt(tmp_path / "b.txt", [_fila("1" * 49), _fila("2" * 48, serie="001-158-000000002")])
    df = leer_sri_retenciones_txt([p1, p2])
    assert sorted(df["clave"]) == sorted(["1" * 49, "0" + "2" * 48])


def test_leer_sri_retenciones_txt_columnas_faltantes(tmp_path):
    p = tmp_path / "x.txt"
    p.write_bytes("A\tB\n1\t2".encode("iso-8859-1"))
    with pytest.raises(ValueError, match="faltan columnas"):
        leer_sri_retenciones_txt([str(p)])


def test_leer_sri_retenciones_txt_sin_rutas():
    df = leer_sri_retenciones_txt([])
    assert list(df.columns) == ["clave", "serie", "ruc", "proveedor", "fecha"]
    assert df.empty
