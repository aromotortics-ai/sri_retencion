from datetime import date

import pandas as pd

import conciliar_retenciones
from conciliar_emitidos import (EST_ANULADO, EST_DIF_DATOS, EST_DIF_VALOR, EST_DUP, EST_INFO, EST_OK,
                                EST_SOLO_ODOO, EST_SOLO_SRI, clasificar_documentos,
                                conciliar_retenciones_emitidas)
from conciliar_retenciones import EST_XML, resumir_por_documento

DESDE, HASTA = date(2026, 9, 1), date(2026, 9, 30)
CLAVE = "1" * 49

SRI_COLS = ["clave", "serie", "fecha", "tipo", "base", "iva", "total", "ruc", "proveedor", "afectada"]
ODOO_COLS = ["id", "num", "nfac", "socio", "ruc", "fecha", "tipo", "clave", "base", "iva", "total",
             "afectada", "estado", "electronico", "elegible"]
SALIDA = ["estado", "tipo", "idx_sri", "id_odoo", "metodo", "detalle"]


def _sri(*filas):
    base = {"clave": CLAVE, "serie": "001-001-000000100", "fecha": pd.Timestamp("2026-09-01"),
            "tipo": "FAC", "base": 100.0, "iva": 15.0, "total": 115.0, "ruc": pd.NA,
            "proveedor": pd.NA, "afectada": pd.NA}
    return pd.DataFrame([{**base, **f} for f in filas], columns=SRI_COLS)


def _odoo(*filas):
    base = {"id": 1, "num": "FAC/2026/0100", "nfac": "001-001-000000100", "socio": "CLIENTE SA",
            "ruc": "1790000000001", "fecha": pd.Timestamp("2026-09-01"), "tipo": "FAC",
            "clave": CLAVE, "base": 100.0, "iva": 15.0, "total": 115.0, "afectada": pd.NA,
            "estado": "posted", "electronico": True, "elegible": True}
    filas = [{**base, **f} for f in filas]
    for f in filas:
        f["elegible"] = f["estado"] in ("posted", "cancel")
    return pd.DataFrame(filas, columns=ODOO_COLS)


def _vacio_sri():
    return pd.DataFrame(columns=SRI_COLS)


def _vacio_odoo():
    return pd.DataFrame(columns=ODOO_COLS)


def _run(sri, odoo):
    out = clasificar_documentos(sri, odoo, DESDE, HASTA)
    assert list(out.columns) == SALIDA
    return out


def test_constantes():
    assert EST_OK == "Coincide"
    assert EST_SOLO_SRI == "Solo en SRI"
    assert EST_SOLO_ODOO == "Solo en Odoo"
    assert EST_DIF_VALOR == "Diferencia de valor"
    assert EST_DIF_DATOS == "Diferencia de datos"
    assert EST_ANULADO == "Odoo anulado, SRI autorizado"
    assert EST_DUP == "Duplicado en Odoo"
    assert EST_INFO == "Informativo"
    # una sola constante compartida con el módulo de retenciones
    assert EST_ANULADO is conciliar_retenciones.EST_ANULADO


def test_coincide_por_clave():
    out = _run(_sri({}), _odoo({}))
    assert len(out) == 1
    r = out.iloc[0]
    assert r["estado"] == EST_OK
    assert r["idx_sri"] == 0 and r["id_odoo"] == 1
    assert r["metodo"] == "Clave de acceso"
    assert r["tipo"] == "FAC"


def test_diferencia_de_valor_mayor_a_2_centavos():
    out = _run(_sri({}), _odoo({"total": 115.03}))
    assert list(out["estado"]) == [EST_DIF_VALOR]
    assert "115.00" in out.iloc[0]["detalle"] and "115.03" in out.iloc[0]["detalle"]


def test_diferencia_de_base_mayor_a_2_centavos_es_diferencia_de_valor():
    out = _run(_sri({}), _odoo({"base": 99.0, "iva": 16.0}))
    assert list(out["estado"]) == [EST_DIF_VALOR]


def test_diferencia_dentro_de_tolerancia_coincide():
    out = _run(_sri({}), _odoo({"total": 115.02, "base": 100.02}))
    assert list(out["estado"]) == [EST_OK]


def test_diferencia_de_fecha_es_diferencia_de_datos():
    out = _run(_sri({}), _odoo({"fecha": pd.Timestamp("2026-09-02")}))
    assert list(out["estado"]) == [EST_DIF_DATOS]
    assert "01/09" in out.iloc[0]["detalle"] and "02/09" in out.iloc[0]["detalle"]


def test_clave_ausente_en_odoo_cruza_por_numero_y_es_diferencia_de_datos():
    out = _run(_sri({}), _odoo({"clave": pd.NA}))
    assert list(out["estado"]) == [EST_DIF_DATOS]
    assert out.iloc[0]["idx_sri"] == 0 and out.iloc[0]["id_odoo"] == 1
    assert out.iloc[0]["metodo"] == "N° documento + importe"


def test_clave_distinta_en_odoo_es_diferencia_de_datos():
    out = _run(_sri({}), _odoo({"clave": "2" * 49}))
    assert list(out["estado"]) == [EST_DIF_DATOS]
    assert "clave" in out.iloc[0]["detalle"].lower()


def test_odoo_no_electronico_con_mismo_numero_es_diferencia_de_datos():
    out = _run(_sri({}), _odoo({"clave": pd.NA, "electronico": False}))
    assert list(out["estado"]) == [EST_DIF_DATOS]
    assert out.iloc[0]["id_odoo"] == 1
    assert "electrónico" in out.iloc[0]["detalle"]


def test_odoo_cancelado_con_sri_es_anulado():
    # sin importar importes
    out = _run(_sri({}), _odoo({"estado": "cancel", "total": 999.0}))
    assert list(out["estado"]) == [EST_ANULADO]
    assert out.iloc[0]["idx_sri"] == 0 and out.iloc[0]["id_odoo"] == 1


def test_odoo_cancelado_sin_sri_es_informativo():
    out = _run(_vacio_sri(), _odoo({"estado": "cancel"}))
    assert list(out["estado"]) == [EST_INFO]
    assert pd.isna(out.iloc[0]["idx_sri"])


def test_odoo_borrador_y_no_electronico_sin_sri_son_informativos():
    out = _run(_vacio_sri(), _odoo({"id": 1, "estado": "draft"},
                                   {"id": 2, "nfac": "001-001-000000200", "clave": pd.NA,
                                    "electronico": False}))
    assert list(out["estado"]) == [EST_INFO, EST_INFO]


def test_odoo_borrador_nunca_cruza():
    out = _run(_sri({}), _odoo({"estado": "draft"}))
    assert sorted(out["estado"]) == sorted([EST_SOLO_SRI, EST_INFO])


def test_solo_en_sri():
    out = _run(_sri({}), _vacio_odoo())
    assert list(out["estado"]) == [EST_SOLO_SRI]
    assert out.iloc[0]["idx_sri"] == 0 and pd.isna(out.iloc[0]["id_odoo"])


def test_solo_en_odoo_dentro_y_fuera_del_periodo():
    odoo = _odoo({"id": 1},
                 {"id": 2, "nfac": "001-001-000000200", "clave": "2" * 49,
                  "fecha": pd.Timestamp("2026-10-03")})
    out = _run(_vacio_sri(), odoo)
    assert list(out["estado"]) == [EST_SOLO_ODOO]
    assert list(out["id_odoo"]) == [1]


def test_nc_compara_valor_absoluto():
    out = _run(_sri({"tipo": "NC", "base": -100.0, "iva": -15.0, "total": -115.0}),
               _odoo({"tipo": "NC"}))
    assert list(out["estado"]) == [EST_OK]


def test_nota_debito_sintetica():
    out = _run(_sri({"tipo": "ND"}), _odoo({"tipo": "ND"}))
    assert list(out["estado"]) == [EST_OK]
    assert out.iloc[0]["tipo"] == "ND"


def test_duplicado_en_odoo():
    out = _run(_sri({}), _odoo({"id": 1}, {"id": 2, "clave": pd.NA}))
    por_id = dict(zip(out["id_odoo"], out["estado"]))
    assert por_id == {1: EST_OK, 2: EST_DUP}


def test_publicado_gana_al_cancelado_con_misma_clave():
    # Odoo trae primero el cancelado; el publicado debe quedarse con el par del SRI.
    out = _run(_sri({}), _odoo({"id": 1, "estado": "cancel"}, {"id": 2}))
    por_id = dict(zip(out["id_odoo"], out["estado"]))
    assert por_id[2] == EST_OK
    assert por_id[1] == EST_INFO


def test_publicado_sin_clave_gana_al_cancelado_con_clave():
    # El cancelado cruza por clave (pasada 1) y el publicado solo por número (pasada 2).
    out = _run(_sri({}), _odoo({"id": 1, "estado": "cancel"}, {"id": 2, "clave": pd.NA}))
    por_id = dict(zip(out["id_odoo"], out["estado"]))
    assert por_id == {1: EST_INFO, 2: EST_DIF_DATOS}
    assert EST_DUP not in set(out["estado"]) and EST_ANULADO not in set(out["estado"])


# ------------------------------------------------------------------------------ retenciones

RSRI_COLS = ["clave", "n_ret", "fecha", "ruc", "proveedor", "factura", "cod", "pct", "base", "valor"]
RODOO_COLS = ["id_ret", "n_ret", "fecha", "estado", "socio", "ruc", "factura", "pct", "base", "valor",
              "cod", "clave"]
TXT_COLS = ["clave", "serie", "ruc", "proveedor", "fecha"]
RUC = "2390000831001"
N_RET = "001-002-000000001"


def _rsri(*filas):
    base = {"clave": CLAVE, "n_ret": N_RET, "fecha": pd.Timestamp("2026-09-10"), "ruc": RUC,
            "proveedor": "PROVEEDOR SA", "factura": "001-001-000000555", "cod": "312", "pct": 2.0,
            "base": 500.0, "valor": 10.0}
    return pd.DataFrame([{**base, **f} for f in filas], columns=RSRI_COLS)


def _rodoo(*filas):
    base = {"id_ret": 7, "n_ret": N_RET, "fecha": pd.Timestamp("2026-09-10"), "estado": "posted",
            "socio": "PROVEEDOR SA", "ruc": RUC, "factura": "001-001-000000555", "pct": 2.0,
            "base": 500.0, "valor": 10.0, "cod": "312", "clave": CLAVE}
    return pd.DataFrame([{**base, **f} for f in filas], columns=RODOO_COLS)


def _txt(*claves_series):
    return pd.DataFrame([{"clave": c, "serie": s, "ruc": RUC, "proveedor": "PROVEEDOR SA",
                          "fecha": pd.Timestamp("2026-09-10")} for c, s in claves_series],
                        columns=TXT_COLS)


def test_retencion_coincide():
    out = conciliar_retenciones_emitidas(_rsri({}), _rodoo({}), _txt((CLAVE, N_RET)), DESDE, HASTA)
    assert list(out["estado"]) == [EST_OK]


def test_retencion_cancelada_en_odoo_con_sri_es_anulada():
    sri = _rsri({"pct": 2.0, "valor": 10.0}, {"cod": "725", "pct": 30.0, "base": 75.0, "valor": 22.5})
    odoo = _rodoo({"estado": "cancel"}, {"estado": "cancel", "cod": "725", "pct": 30.0, "base": 75.0,
                                         "valor": 22.5})
    out = conciliar_retenciones_emitidas(sri, odoo, _txt((CLAVE, N_RET)), DESDE, HASTA)
    assert set(out["estado"]) == {EST_ANULADO}
    assert len(out) == 2
    assert conciliar_retenciones.EST_INFO not in set(out["estado"])
    docs = resumir_por_documento(out)
    assert len(docs) == 1
    assert docs.iloc[0]["estado"] == EST_ANULADO
    assert docs.iloc[0]["n_ret"] == N_RET


def test_retencion_cancelada_sin_sri_sigue_informativa():
    out = conciliar_retenciones_emitidas(_rsri().iloc[0:0], _rodoo({"estado": "cancel"}),
                                         _txt(), DESDE, HASTA)
    assert list(out["estado"]) == [conciliar_retenciones.EST_INFO]


def test_retencion_sin_xml():
    out = conciliar_retenciones_emitidas(_rsri().iloc[0:0], _rodoo().iloc[0:0],
                                         _txt((CLAVE, N_RET)), DESDE, HASTA)
    assert list(out["estado"]) == [EST_XML]
    assert EST_SOLO_SRI not in set(out["estado"])
