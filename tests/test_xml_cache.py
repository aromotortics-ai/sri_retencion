from pathlib import Path

import pytest

import xml_cache
from _helpers import _autorizacion
from xml_cache import descargar_xmls, lineas_sri

CLAVE_V1 = "1" * 49
CLAVE_V2 = "2" * 49
CLAVE_OTRA = "3" * 49


def _escribir(cache, clave, version="1.0.0", num_doc="001003000088766"):
    (cache / f"{clave}.xml").write_text(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + _autorizacion(clave, version, num_doc),
        encoding="utf-8",
    )


class FakeResp:
    def __init__(self, text):
        self.text = text

    def raise_for_status(self):
        pass


class FakeSession:
    def __init__(self, respuestas):
        self.respuestas = respuestas  # clave -> texto | Exception
        self.llamadas = []

    def post(self, url, data=None, headers=None, timeout=None):
        clave = data.decode("utf-8").split("<claveAccesoComprobante>")[1].split("<")[0]
        self.llamadas.append((url, headers, timeout, clave))
        r = self.respuestas[clave]
        if isinstance(r, Exception):
            raise r
        return FakeResp(r)


def test_lineas_sri_v1_y_v2(tmp_path):
    _escribir(tmp_path, CLAVE_V1, "1.0.0")
    _escribir(tmp_path, CLAVE_V2, "2.0.0", num_doc="001003000088766")
    df = lineas_sri([CLAVE_V1, CLAVE_V2], tmp_path)
    assert list(df.columns) == ["clave", "n_ret", "fecha", "ruc", "proveedor",
                                "factura", "cod", "pct", "base", "valor"]
    assert len(df) == 2
    for _, f in df.iterrows():
        assert f["n_ret"] == "001-200-000004340"
        assert f["factura"] == "001-003-000088766"
        assert f["pct"] == 1.75 and f["base"] == 100.0 and f["valor"] == 1.75
        assert f["cod"] == "312"
        assert f["ruc"] == "1790000000001" and f["proveedor"] == "PROVEEDOR SA"
    assert df["fecha"].iloc[0].year == 2026 and df["fecha"].iloc[0].month == 9


def test_lineas_sri_factura_ya_con_guiones(tmp_path):
    _escribir(tmp_path, CLAVE_V1, "1.0.0", num_doc="001-003-000088766")
    df = lineas_sri([CLAVE_V1], tmp_path)
    assert df["factura"].iloc[0] == "001-003-000088766"


def test_lineas_sri_filtra_por_claves(tmp_path):
    _escribir(tmp_path, CLAVE_V1)
    _escribir(tmp_path, CLAVE_OTRA)
    df = lineas_sri([CLAVE_V1], tmp_path)
    assert set(df["clave"]) == {CLAVE_V1}


def test_lineas_sri_un_xml_corrupto_no_aborta(tmp_path):
    _escribir(tmp_path, CLAVE_V1)
    (tmp_path / f"{CLAVE_V2}.xml").write_text("no es xml <<<", encoding="utf-8")
    df = lineas_sri([CLAVE_V1, CLAVE_V2], tmp_path)
    assert set(df["clave"]) == {CLAVE_V1}


def test_lineas_sri_vacio_conserva_columnas(tmp_path):
    df = lineas_sri([CLAVE_V1], tmp_path)
    assert df.empty
    assert "n_ret" in df.columns


def test_descargar_xmls_usa_cache(tmp_path):
    _escribir(tmp_path, CLAVE_V1)
    ses = FakeSession({})
    fallidos = descargar_xmls([CLAVE_V1], tmp_path, sesion=ses, pausa_seg=0)
    assert fallidos == []
    assert ses.llamadas == []


def test_descargar_xmls_guarda_autorizadas_y_reporta_fallidas(tmp_path):
    rechazada = "<autorizacion><estado>NO AUTORIZADO</estado></autorizacion>"
    ses = FakeSession({
        CLAVE_V1: _autorizacion(CLAVE_V1, "1.0.0"),
        CLAVE_V2: rechazada,
        CLAVE_OTRA: ConnectionError("boom"),
    })
    fallidos = descargar_xmls([CLAVE_V1, CLAVE_V2, CLAVE_OTRA], tmp_path,
                              sesion=ses, pausa_seg=0)
    assert (tmp_path / f"{CLAVE_V1}.xml").read_text(encoding="utf-8").startswith("<?xml")
    assert not (tmp_path / f"{CLAVE_V2}.xml").exists()
    por_clave = {f["clave"]: f["motivo"] for f in fallidos}
    assert set(por_clave) == {CLAVE_V2, CLAVE_OTRA}
    assert "NO AUTORIZADO" in por_clave[CLAVE_V2]
    assert "ConnectionError" in por_clave[CLAVE_OTRA]
    assert ses.llamadas[0][1]["Connection"] == "close"


def test_descargar_xmls_progreso(tmp_path):
    claves = [str(i).zfill(49) for i in range(12)]
    ses = FakeSession({c: _autorizacion(c, "1.0.0") for c in claves})
    eventos = []
    descargar_xmls(claves, tmp_path, on_progress=lambda *a: eventos.append(a),
                   sesion=ses, pausa_seg=0)
    assert eventos[0] == (10, 12, 0)
    assert eventos[-1] == (12, 12, 0)
