import pandas as pd
import pytest

from conciliar_compras import TIPOS, leer_sri_txts, tipos_faltantes

COLS = ["RUC_EMISOR", "RAZON_SOCIAL_EMISOR", "TIPO_COMPROBANTE", "SERIE_COMPROBANTE",
        "CLAVE_ACCESO", "FECHA_AUTORIZACION", "FECHA_EMISION", "IDENTIFICACION_RECEPTOR",
        "VALOR_SIN_IMPUESTOS", "IVA", "IMPORTE_TOTAL", "NUMERO_DOCUMENTO_MODIFICADO"]


def _txt(path, filas):
    lineas = ["\t".join(COLS)]
    for f in filas:
        d = dict(zip(COLS, [""] * len(COLS)))
        d.update(f)
        lineas.append("\t".join(d[c] for c in COLS))
    path.write_bytes("\n".join(lineas).encode("iso-8859-1"))
    return str(path)


def _fila(clave, tipo="Factura", serie="001-001-000000001", fecha="05/09/2026", **kw):
    f = {"RUC_EMISOR": "1790000000001", "RAZON_SOCIAL_EMISOR": "PROVEEDOR SA",
         "TIPO_COMPROBANTE": tipo, "SERIE_COMPROBANTE": serie, "CLAVE_ACCESO": clave,
         "FECHA_EMISION": fecha, "VALOR_SIN_IMPUESTOS": "100.00", "IVA": "15.00",
         "IMPORTE_TOTAL": "115.00"}
    f.update(kw)
    return f


def test_total_vacio_en_nc_es_base_mas_iva(tmp_path):
    p = _txt(tmp_path / "nc.txt", [_fila("1" * 49, "Nota de Crédito", VALOR_SIN_IMPUESTOS="300.26",
                                         IVA="45.04", IMPORTE_TOTAL="")])
    df = leer_sri_txts([p])
    assert df.loc[0, "tipo"] == "NC"
    assert df.loc[0, "total"] == pytest.approx(345.30)


def test_tipo_detectado_por_texto_con_acentos_iso8859(tmp_path):
    p = _txt(tmp_path / "a.txt", [_fila("1" * 49, "Factura"),
                                  _fila("2" * 49, "Nota de Crédito", serie="001-001-000000002"),
                                  _fila("3" * 49, "Nota de Débito", serie="001-001-000000003")])
    df = leer_sri_txts([p])
    assert sorted(df["tipo"]) == ["FAC", "NC", "ND"]
    assert df["proveedor"].iloc[0] == "PROVEEDOR SA"


def test_ignora_retenciones_y_liquidaciones(tmp_path):
    p = _txt(tmp_path / "a.txt", [_fila("1" * 49, "Factura"),
                                  _fila("2" * 49, "Comprobante de Retención"),
                                  _fila("3" * 49, "Liquidación de compra de bienes y prestación de servicios")])
    df = leer_sri_txts([p])
    assert list(df["clave"]) == ["1" * 49]


def test_misma_clave_en_dos_archivos_se_deduplica(tmp_path):
    f = _fila("1" * 49)
    a = _txt(tmp_path / "a.txt", [f])
    b = _txt(tmp_path / "b.txt", [f])
    assert len(leer_sri_txts([a, b])) == 1


def test_afectada_viene_de_numero_documento_modificado(tmp_path):
    p = _txt(tmp_path / "a.txt", [
        _fila("1" * 49, "Nota de Crédito", NUMERO_DOCUMENTO_MODIFICADO="001-002-000000099", IMPORTE_TOTAL=""),
        _fila("2" * 49, "Factura", serie="001-001-000000002")])
    df = leer_sri_txts([p]).set_index("clave")
    assert df.loc["1" * 49, "afectada"] == "001-002-000000099"
    assert pd.isna(df.loc["2" * 49, "afectada"])
    assert df.loc["1" * 49, "fecha"] == pd.Timestamp(2026, 9, 5)


def test_tipos_faltantes_reporta_nd_si_no_hay_archivo(tmp_path):
    p = _txt(tmp_path / "a.txt", [_fila("1" * 49, "Factura"),
                                  _fila("2" * 49, "Nota de Crédito", serie="001-001-000000002")])
    assert tipos_faltantes(leer_sri_txts([p])) == ["ND"]
    assert TIPOS == ("FAC", "NC", "ND")


def test_columnas_faltantes_error_claro(tmp_path):
    p = tmp_path / "x.txt"
    p.write_bytes(b"A\tB\n1\t2")
    with pytest.raises(ValueError, match="columnas"):
        leer_sri_txts([str(p)])


# ------------------------------------------------------------------ conciliar ---
from datetime import date

from conciliar_compras import CAT_DUP, CAT_INFO, CAT_SOLO_ODOO, conciliar

D, H = date(2026, 9, 1), date(2026, 9, 30)


def _sri(*filas):
    base = {"clave": None, "serie": "001-001-000000001", "ruc": "1790000000001", "proveedor": "P",
            "fecha": pd.Timestamp(2026, 9, 5), "tipo": "FAC", "base": 100.0, "iva": 15.0,
            "total": 115.0, "afectada": pd.NA}
    return pd.DataFrame([{**base, **f} for f in filas])


def _odoo(*filas):
    base = {"id": 1, "num": "FAC/1", "nfac": "001-001-000000001", "socio": "P", "ruc": "1790000000001",
            "fecha": pd.Timestamp(2026, 9, 5), "tipo": "FAC", "cod_comp": "01", "clave": pd.NA,
            "base": 100.0, "iva": 15.0, "total": 115.0, "afectada": pd.NA, "elegible": True}
    return pd.DataFrame([{**base, **f} for f in filas])


def test_cruza_por_clave_con_48_digitos_rellenada():
    k = "0" + "1" * 48
    r = conciliar(_sri({"clave": k, "serie": "X"}), _odoo({"clave": k, "nfac": "Y"}), D, H)
    assert len(r) == 1
    assert r.loc[0, "metodo"] == "Clave de acceso"
    assert r.loc[0, "idx_sri"] == 0 and r.loc[0, "id_odoo"] == 1 and r.loc[0, "cat"] == ""


def test_cruza_por_numero_y_total_si_la_clave_de_odoo_esta_mal():
    r = conciliar(_sri({"clave": "1" * 49}), _odoo({"clave": "9" * 49, "total": 115.01}), D, H)
    assert len(r) == 1
    assert r.loc[0, "metodo"] == "N° documento + importe"


def test_numero_con_total_distinto_se_empareja_pero_marca_metodo():
    r = conciliar(_sri({"clave": "1" * 49}), _odoo({"total": 99.0}), D, H)
    assert len(r) == 1
    assert r.loc[0, "metodo"] == "N° documento (importe distinto)"
    assert r.loc[0, "id_odoo"] == 1


def test_mismo_numero_distinto_ruc_no_se_empareja():
    r = conciliar(_sri({"clave": "1" * 49}), _odoo({"ruc": "1790000000002"}), D, H)
    assert len(r) == 2
    assert set(r["cat"]) == {"", CAT_SOLO_ODOO}


def test_ruc_faltante_en_un_lado_no_impide_cruce():
    r = conciliar(_sri({"clave": "1" * 49}), _odoo({"ruc": pd.NA}), D, H)
    assert len(r) == 1 and r.loc[0, "metodo"] == "N° documento + importe"


def test_nc_cruza_por_factura_afectada():
    s = _sri({"clave": "1" * 49, "tipo": "NC", "serie": "001-001-000000050", "afectada": "001-002-000000099",
              "total": 50.0})
    o = _odoo({"tipo": "NC", "nfac": "OTRO", "afectada": "001-002-000000099", "total": -50.0})
    r = conciliar(s, o, D, H)
    assert len(r) == 1
    assert r.loc[0, "metodo"] == "NC: factura afectada + importe"


def test_no_elegible_es_informativo_no_error():
    r = conciliar(_sri(), _odoo({"id": 7, "nfac": "Z", "elegible": False}), D, H)
    fila = r[r["id_odoo"] == 7].iloc[0]
    assert fila["cat"] == CAT_INFO and pd.isna(fila["idx_sri"])


def test_odoo_fuera_de_periodo_sin_pareja_se_omite():
    r = conciliar(_sri({"clave": "1" * 49, "serie": "S"}), _odoo({"id": 7, "nfac": "Z", "fecha": pd.Timestamp(2026, 10, 2)}), D, H)
    assert 7 not in set(r["id_odoo"].dropna())
    assert len(r) == 1  # solo el SRI


def test_odoo_fuera_de_periodo_con_pareja_si_cruza():
    r = conciliar(_sri({"clave": "1" * 49}), _odoo({"fecha": pd.Timestamp(2026, 10, 2)}), D, H)
    assert len(r) == 1 and r.loc[0, "id_odoo"] == 1 and r.loc[0, "idx_sri"] == 0


def test_duplicado_en_odoo():
    o = _odoo({"id": 1}, {"id": 2, "num": "FAC/2"})
    r = conciliar(_sri({"clave": "1" * 49}), o, D, H)
    assert len(r) == 2
    assert r.set_index("id_odoo").loc[1, "cat"] == ""
    assert r.set_index("id_odoo").loc[2, "cat"] == CAT_DUP


def test_no_cruza_entre_tipos():
    s = _sri({"clave": "1" * 49, "tipo": "NC"})
    r = conciliar(s, _odoo({"tipo": "FAC"}), D, H)
    assert len(r) == 2
    assert r["idx_sri"].notna().sum() == 1 and r["id_odoo"].notna().sum() == 1
    assert not ((r["idx_sri"].notna()) & (r["id_odoo"].notna())).any()


def test_solo_sri_genera_fila_sin_id_odoo():
    r = conciliar(_sri({"clave": "1" * 49}), _odoo({"nfac": "Z", "total": 1.0}).iloc[0:0], D, H)
    assert len(r) == 1
    assert r.loc[0, "idx_sri"] == 0 and pd.isna(r.loc[0, "id_odoo"]) and r.loc[0, "tipo"] == "FAC"
