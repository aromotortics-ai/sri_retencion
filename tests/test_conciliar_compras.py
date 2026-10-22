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
