import pandas as pd
import pytest

from conciliar_ventas_sri import conciliar_ventas

SRI_HDR = ("COMPROBANTE\tSERIE_COMPROBANTE\tCLAVE_ACCESO\tFECHA_AUTORIZACION\t"
           "FECHA_EMISION\tVALOR_SIN_IMPUESTOS\tIVA\tIMPORTE_TOTAL\n")
ODOO_COLS = ["Número", "Número de autorización", "Nombre del socio", "Sucursal",
             "Fecha de Factura/Recibo", "Fecha de autorización",
             "Importe sin impuestos con signo", "Total con signo", "Estado",
             "Estado electrónico"]


def _sri_line(num, clave, dia, base, iva, total):
    return (f"Factura\t{num}\t{clave}\t{dia} 10:00:00\t{dia} 00:00:00\t"
            f"{base}\t{iva}\t{total}\n")


def _odoo_row(num, clave, fecha, base, total, estado="Publicado", edi="authorized"):
    return [num, clave, "CLIENTE", "MATRIZ", f"{fecha} 00:00:00",
            f"{fecha} 10:00:00", str(base), str(total), estado, edi]


@pytest.fixture
def datos(tmp_path):
    emitidos = tmp_path / "emitidos"
    emitidos.mkdir()
    (emitidos / "emitidos_2026-03-01.txt").write_text(
        SRI_HDR
        + _sri_line("001-001-000000001", "K1", "01/03/2026", "10.00", "1.50", "11.50")
        + _sri_line("001-001-000000002", "K2", "01/03/2026", "20.00", "3.00", "23.00")
        + _sri_line("001-001-000000003", "K3", "01/03/2026", "30.00", "4.50", "34.50"),
        encoding="iso-8859-1",
    )
    (emitidos / "emitidos_2026-03-02.txt").write_text(
        SRI_HDR
        # duplicada de la 001 en otro día
        + _sri_line("001-001-000000001", "K1", "01/03/2026", "10.00", "1.50", "11.50"),
        encoding="iso-8859-1",
    )
    odoo = pd.DataFrame([
        _odoo_row("001-001-000000001", "K1", "2026-03-01", 10.0, 11.5),     # coincide
        _odoo_row("001-001-000000002", "K2", "2026-03-01", 20.0, 25.0),     # dif valor
        _odoo_row("001-001-000000004", "K4", "2026-03-02", 40.0, 46.0),     # solo odoo
        _odoo_row("001-001-000000005", "K5", "2026-03-02", 5.0, 5.75,
                  estado="Cancelado", edi="cancelled"),                     # no autorizado
    ], columns=ODOO_COLS)
    odoo_path = tmp_path / "odoo.xlsx"
    odoo.to_excel(odoo_path, index=False)
    return odoo_path, emitidos, tmp_path / "salida.xlsx"


def test_resumen_cuenta_cada_estado(datos):
    odoo, emitidos, salida = datos
    r = conciliar_ventas(str(odoo), str(emitidos), str(salida))
    assert r["coinciden"] == 1
    assert r["diferencias"] == 1
    assert r["solo_sri"] == 1      # 003
    assert r["solo_odoo"] == 1     # 004
    assert r["odoo_no_autorizado"] == 1
    assert r["duplicados_sri"] == 2  # ambas filas de la 001
    assert r["n_sri"] == 3
    assert r["n_odoo"] == 4
    assert salida.exists()


def test_excel_tiene_las_hojas(datos):
    odoo, emitidos, salida = datos
    conciliar_ventas(str(odoo), str(emitidos), str(salida))
    hojas = pd.ExcelFile(salida).sheet_names
    for h in ("Resumen", "Detalle", "Solo_en_SRI", "Solo_en_Odoo",
              "Diferencias", "Odoo_no_autorizado"):
        assert h in hojas


def test_odoo_con_columnas_incorrectas(datos, tmp_path):
    _, emitidos, salida = datos
    malo = tmp_path / "malo.xlsx"
    pd.DataFrame({"a": [1], "b": [2]}).to_excel(malo, index=False)
    with pytest.raises(ValueError, match="Odoo"):
        conciliar_ventas(str(malo), str(emitidos), str(salida))


def test_sri_sin_archivos(datos, tmp_path):
    odoo, _, salida = datos
    vacia = tmp_path / "vacia"
    vacia.mkdir()
    with pytest.raises(ValueError, match="emitidos"):
        conciliar_ventas(str(odoo), str(vacia), str(salida))


def test_sri_con_columnas_incorrectas(datos, tmp_path):
    odoo, _, salida = datos
    carpeta = tmp_path / "raro"
    carpeta.mkdir()
    (carpeta / "emitidos_2026-03-01.txt").write_text("X\tY\n1\t2\n", encoding="iso-8859-1")
    with pytest.raises(ValueError, match="SRI"):
        conciliar_ventas(str(odoo), str(carpeta), str(salida))
