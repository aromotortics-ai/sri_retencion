from pathlib import Path

import pytest

import xml_cache
from _helpers import _autorizacion
from xml_cache import descargar_xmls, lineas_sri, resumir_motivos

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
                              sesion=ses, pausa_seg=0, esperas=(0, 0))
    assert (tmp_path / f"{CLAVE_V1}.xml").read_text(encoding="utf-8").startswith("<?xml")
    assert not (tmp_path / f"{CLAVE_V2}.xml").exists()
    por_clave = {f["clave"]: f["motivo"] for f in fallidos}
    assert set(por_clave) == {CLAVE_V2, CLAVE_OTRA}
    assert "NO AUTORIZADO" in por_clave[CLAVE_V2]
    assert "ConnectionError" in por_clave[CLAVE_OTRA]
    assert "tras 3 intentos" in por_clave[CLAVE_OTRA]
    assert "tras" not in por_clave[CLAVE_V2]  # respuesta definitiva: sin reintentos
    assert ses.llamadas[0][1]["Connection"] == "close"


def test_descargar_xmls_escritura_atomica_sin_restos(tmp_path):
    ses = FakeSession({CLAVE_V1: _autorizacion(CLAVE_V1, "1.0.0")})
    assert descargar_xmls([CLAVE_V1], tmp_path, sesion=ses, pausa_seg=0) == []
    assert sorted(p.name for p in tmp_path.iterdir()) == [f"{CLAVE_V1}.xml"]


def test_descargar_xmls_escritura_fallida_no_deja_parcial(tmp_path, monkeypatch):
    ses = FakeSession({CLAVE_V1: _autorizacion(CLAVE_V1, "1.0.0")})

    def falla(*a, **k):
        raise OSError("disco lleno")
    monkeypatch.setattr(xml_cache.os, "replace", falla)
    fallidos = descargar_xmls([CLAVE_V1], tmp_path, sesion=ses, pausa_seg=0, esperas=(0, 0))
    assert [f["clave"] for f in fallidos] == [CLAVE_V1]
    assert list(tmp_path.iterdir()) == []


def test_tmp_huerfano_no_cuenta_como_cache(tmp_path):
    (tmp_path / f"{CLAVE_V1}.xml.abc.tmp").write_text("<parcial", encoding="utf-8")
    ses = FakeSession({CLAVE_V1: _autorizacion(CLAVE_V1, "1.0.0")})
    assert descargar_xmls([CLAVE_V1], tmp_path, sesion=ses, pausa_seg=0) == []
    assert len(ses.llamadas) == 1
    assert set(lineas_sri([CLAVE_V1], tmp_path)["clave"]) == {CLAVE_V1}


def test_descargar_xmls_progreso(tmp_path):
    claves = [str(i).zfill(49) for i in range(12)]
    ses = FakeSession({c: _autorizacion(c, "1.0.0") for c in claves})
    eventos = []
    descargar_xmls(claves, tmp_path, on_progress=lambda *a: eventos.append(a),
                   sesion=ses, pausa_seg=0)
    assert eventos[0] == (10, 12, 0)
    assert eventos[-1] == (12, 12, 0)


def test_lineas_sri_contraparte(tmp_path):
    _escribir(tmp_path, CLAVE_V2, "2.0.0")
    normal = lineas_sri([CLAVE_V2], tmp_path)
    assert normal["proveedor"].iloc[0] == "PROVEEDOR SA"
    cp = lineas_sri([CLAVE_V2], tmp_path, contraparte=True)
    assert cp["proveedor"].iloc[0] == "EMPRESA"
    assert cp["ruc"].iloc[0] == "1790000000001"
    assert list(cp.columns) == list(normal.columns)


class SesionIntermitente:
    """Falla `fallos` veces con `error` y luego responde `texto`."""
    def __init__(self, fallos, error, texto):
        self.fallos, self.error, self.texto, self.n = fallos, error, texto, 0

    def post(self, url, data=None, headers=None, timeout=None):
        self.n += 1
        if self.n <= self.fallos:
            raise self.error
        return FakeResp(self.texto)


def test_descargar_xmls_reintenta_fallo_transitorio(tmp_path):
    ses = SesionIntermitente(2, ConnectionError("RemoteDisconnected"), _autorizacion(CLAVE_V1, "1.0.0"))
    assert descargar_xmls([CLAVE_V1], tmp_path, sesion=ses, pausa_seg=0, esperas=(0, 0)) == []
    assert ses.n == 3
    assert (tmp_path / f"{CLAVE_V1}.xml").exists()


def test_descargar_xmls_no_reintenta_rechazo_ni_sin_comprobante(tmp_path):
    rechazada = "<autorizacion><estado>NO AUTORIZADO</estado></autorizacion>"
    vacia = ("<RespuestaAutorizacionComprobante><claveAccesoConsultada>x</claveAccesoConsultada>"
             "<numeroComprobantes>0</numeroComprobantes><autorizaciones/></RespuestaAutorizacionComprobante>")
    ses = FakeSession({CLAVE_V1: rechazada, CLAVE_V2: vacia})
    fallidos = descargar_xmls([CLAVE_V1, CLAVE_V2], tmp_path, sesion=ses, pausa_seg=0, esperas=(0, 0))
    assert len(ses.llamadas) == 2  # un intento por clave
    assert "numeroComprobantes=0" in {f["clave"]: f["motivo"] for f in fallidos}[CLAVE_V2]


def test_descargar_xmls_http_4xx_no_se_reintenta(tmp_path):
    class Resp4xx(FakeResp):
        status_code = 404

    class Err(Exception):
        response = Resp4xx("")
    ses = SesionIntermitente(9, Err("404"), "")
    fallidos = descargar_xmls([CLAVE_V1], tmp_path, sesion=ses, pausa_seg=0, esperas=(0, 0))
    assert ses.n == 1 and len(fallidos) == 1


def test_descargar_xmls_espera_entre_intentos(tmp_path, monkeypatch):
    pausas = []
    monkeypatch.setattr(xml_cache.time, "sleep", pausas.append)
    ses = SesionIntermitente(9, TimeoutError("t"), "")
    descargar_xmls([CLAVE_V1], tmp_path, sesion=ses, pausa_seg=0)
    assert pausas == [1, 3]


def test_resumir_motivos_agrupa_y_ordena():
    f = [{"clave": "a", "motivo": "X"}, {"clave": "b", "motivo": "Y"}, {"clave": "c", "motivo": "Y"}]
    assert resumir_motivos(f) == "Y (×2); X (×1)"
    assert resumir_motivos(f, maximo=1).endswith("; ...")
