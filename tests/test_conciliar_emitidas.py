import pandas as pd
import pytest

from conciliar_emitidas import conciliar_emitidas, leer_odoo, leer_sri

SRI_HDR = ("COMPROBANTE\tSERIE_COMPROBANTE\tCLAVE_ACCESO\tFECHA_AUTORIZACION\t"
           "FECHA_EMISION\tVALOR_SIN_IMPUESTOS\tIVA\tIMPORTE_TOTAL\n")
ODOO_COLS = ["Factura", "Contacto", "Sucursal", "Posición fiscal", "No. de retención",
             "Fecha de retención", "Asiento contable", "Asiento contable/Fecha", "Base",
             "Porcentaje", "Valor", "Código", "Categoría"]


def _sri(n, clave, dia):
    return f"Comprobante de Retención\t{n}\t{clave}\t{dia} 10:00:00\t{dia} 00:00:00\t\t\t\n"


def _odoo(n, fecha, tipo, valor, factura="CE/1"):
    cat = "IR - Impuesto Renta" if tipo == "IR" else "IVA - Impuesto Valor Agregado"
    return [factura, "PROV", "COBRA", "PJ", n, pd.Timestamp(fecha), "AS/1", pd.Timestamp(fecha),
            100.0, 2, valor, "312", cat]


@pytest.fixture
def datos(tmp_path):
    sri = tmp_path / "sri"
    sri.mkdir()
    (sri / "emitidos_2026-03-31.txt").write_text(
        SRI_HDR + _sri("001-002-000000001", "K1", "31/03/2026")
        + _sri("001-002-000000002", "K2", "31/03/2026"), encoding="iso-8859-1")
    (sri / "emitidos_2026-04-01.txt").write_text(
        SRI_HDR + _sri("001-002-000000003", "K3", "01/04/2026"), encoding="iso-8859-1")
    xlsx = tmp_path / "odoo.xlsx"
    pd.DataFrame([
        _odoo("001-002-000000001", "2026-03-31", "IR", 2.0),
        _odoo("001-002-000000001", "2026-03-31", "IVA", 3.0),
        _odoo("001-002-000000003", "2026-03-30", "IR", 5.0),   # fecha distinta
        _odoo("001-002-000000009", "2026-04-02", "IR", 7.0),   # solo Odoo
    ], columns=ODOO_COLS).to_excel(xlsx, index=False)
    return str(xlsx), str(sri)


def test_estados_y_valores(datos):
    xlsx, sri = datos
    c = conciliar_emitidas(leer_odoo(xlsx), leer_sri(sri)).set_index("n_ret")
    assert c.loc["001-002-000000001", "resultado"] == "Coincide"
    assert c.loc["001-002-000000001", "valor_ir"] == 2.0
    assert c.loc["001-002-000000001", "valor_iva"] == 3.0
    assert c.loc["001-002-000000002", "resultado"] == "Solo en SRI"
    assert c.loc["001-002-000000009", "resultado"] == "Solo en Odoo"
    assert c.loc["001-002-000000003", "resultado"] == "Coincide"
    assert c.loc["001-002-000000003", "fecha_emision"] != c.loc["001-002-000000003", "fecha_odoo"]


def test_mes_usa_fecha_sri_y_filtra_rango(datos):
    xlsx, sri = datos
    c = conciliar_emitidas(leer_odoo(xlsx), leer_sri(sri)).set_index("n_ret")
    assert c.loc["001-002-000000003", "mes"] == "2026-04"   # mes SRI, no el de Odoo (marzo)
    solo_marzo = conciliar_emitidas(leer_odoo(xlsx, hasta="2026-03-31"), leer_sri(sri, hasta="2026-03-31"))
    assert set(solo_marzo["mes"]) == {"2026-03"}


def test_odoo_con_n_retencion_invalido(tmp_path):
    xlsx = tmp_path / "odoo.xlsx"
    pd.DataFrame([_odoo("1-2-3", "2026-03-31", "IR", 1.0)], columns=ODOO_COLS).to_excel(xlsx, index=False)
    with pytest.raises(ValueError, match="inválido"):
        leer_odoo(str(xlsx))
