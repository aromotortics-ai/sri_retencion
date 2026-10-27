# -*- coding: utf-8 -*-
"""Sesión de trabajo: TXT por sesión y recálculo cancelable por generación.

TXT sintéticos, conciliación y Excel reales; Odoo, descarga de XML, caché y LibreOffice falsos
(nunca red, nunca Odoo, nunca el caché real)."""
import json
import threading

import pandas as pd
import pytest

import app.odoo.client as odoo_client
from app.job_manager import get_job_dir, read_status

COLS_TXT = ["RUC_EMISOR", "RAZON_SOCIAL_EMISOR", "TIPO_COMPROBANTE", "SERIE_COMPROBANTE",
            "CLAVE_ACCESO", "FECHA_AUTORIZACION", "FECHA_EMISION", "IDENTIFICACION_RECEPTOR",
            "VALOR_SIN_IMPUESTOS", "IVA", "IMPORTE_TOTAL", "NUMERO_DOCUMENTO_MODIFICADO"]


def _txt(*facturas):
    """TXT de recibidos del portal (ISO-8859-1, tabulador) con una factura por número."""
    lineas = ["\t".join(COLS_TXT)]
    for n in facturas:
        d = dict.fromkeys(COLS_TXT, "")
        d.update({"RUC_EMISOR": "1790000000001", "RAZON_SOCIAL_EMISOR": "PROVEEDOR SA",
                  "TIPO_COMPROBANTE": "Factura", "SERIE_COMPROBANTE": f"001-001-{n:09d}",
                  "CLAVE_ACCESO": str(n).zfill(49), "FECHA_EMISION": "05/09/2026",
                  "VALOR_SIN_IMPUESTOS": "100.00", "IVA": "15.00", "IMPORTE_TOTAL": "115.00"})
        lineas.append("\t".join(d[c] for c in COLS_TXT))
    return "\n".join(lineas).encode("iso-8859-1")


class _Odoo:
    """Lector falso de compras sin registros; cuenta las lecturas."""

    def __init__(self):
        self.lecturas = 0

    def fetch_compras(self, desde, hasta, margen_dias=5):
        self.lecturas += 1
        return pd.DataFrame(columns=odoo_client.COLUMNAS)

    def fetch_retenciones(self, desde, hasta, margen_dias=5):
        self.lecturas += 1
        return pd.DataFrame(columns=odoo_client.COLUMNAS_RET)


@pytest.fixture
def entorno(tmp_path, monkeypatch):
    """Devuelve (sesion, pc, instancias, lanzados): módulos, llamadas a from_env y recálculos pedidos."""
    monkeypatch.setattr("app.job_manager.JOBS_DIR", tmp_path / "jobs")
    import app.compras.pipeline_compras as pc
    import app.sesion as sesion
    monkeypatch.setattr(sesion, "schedule_cleanup", lambda *a, **k: None)
    monkeypatch.setattr("shutil.which", lambda *_: None)  # sin LibreOffice
    monkeypatch.setattr(pc, "CACHE_XML_DIR", tmp_path / "cache" / "xml")
    monkeypatch.setattr(pc, "descargar_xmls", lambda *a, **k: [])
    lector = _Odoo()
    instancias = []

    def from_env(cls):
        instancias.append(1)
        return lector
    monkeypatch.setattr(pc.OdooReader, "from_env", classmethod(from_env))
    lanzados = []
    monkeypatch.setattr(sesion, "_lanzar", lambda sid, gen: lanzados.append((sid, gen)))
    return sesion, pc, lector, instancias, lanzados


def _tablero(sid):
    return json.loads((get_job_dir(sid) / "tablero.json").read_text(encoding="utf-8"))


def test_mismo_nombre_reemplaza(entorno):
    sesion, *_ = entorno
    sid = sesion.crear_sesion("compras")
    sesion.agregar_txt(sid, "facturas.txt", _txt(1))
    r = sesion.agregar_txt(sid, "facturas.txt", _txt(1, 2))
    lista = sesion.listar_txt(sid)
    assert [a["nombre"] for a in lista] == ["facturas.txt"]
    assert r == lista[0]
    assert r["estado"] == "ok" and r["filas"] == 2 and r["tipos"]["FAC"] == 2 and r["error"] is None
    assert sorted(p.name for p in (get_job_dir(sid) / "txt").iterdir()) == ["facturas.txt"]


def test_extension_y_ruta_normalizadas(entorno):
    sesion, *_ = entorno
    sid = sesion.crear_sesion("compras")
    r = sesion.agregar_txt(sid, "..\\..\\otra/carpeta/SEMANA.TXT", _txt(1))
    assert r["nombre"] == "SEMANA.txt"
    assert (get_job_dir(sid) / "txt" / "SEMANA.txt").exists()


def test_txt_ilegible_se_excluye_y_el_resto_concilia(entorno):
    sesion, _, lector, _, lanzados = entorno
    sid = sesion.crear_sesion("compras")
    roto = sesion.agregar_txt(sid, "roto.txt", b"esto no es\tun TXT del portal\n1\t2\n")
    assert roto["estado"] == "error" and roto["error"] and roto["filas"] == 0
    assert "roto.txt" in roto["error"] and str(get_job_dir(sid)) not in roto["error"]
    assert sesion.agregar_txt(sid, "fac.txt", _txt(1, 2))["estado"] == "ok"
    assert [a["estado"] for a in sesion.listar_txt(sid)] == ["error", "ok"]

    gen = sesion.solicitar_recalculo(sid)
    assert lanzados == [(sid, gen)]
    sesion.recalcular(sid, gen)
    s = read_status(sid)
    assert s["estado"] == "done", s
    t = _tablero(sid)
    assert t["generacion"] == gen and t["modo"] == "compras"
    fac = [c for c in t["tarjetas"] if c["tipo"] == "FAC"]
    assert fac and fac[0]["total"] == 2
    assert any("roto.txt" in a for a in t["advertencias"])
    assert any("roto.txt" in a for a in s["resumen"]["advertencias"])
    assert s["resumen"]["por_tipo"]["FAC"]["solo_sri"] == 2
    assert (get_job_dir(sid) / "resultado.xlsx").exists()
    assert (get_job_dir(sid) / "observaciones.xlsx").exists()
    assert lector.lecturas == 1


def test_quitar_ultimo_txt_deja_tablero_vacio(entorno):
    sesion, _, _, instancias, _ = entorno
    sid = sesion.crear_sesion("compras")
    sesion.agregar_txt(sid, "fac.txt", _txt(1))
    sesion.recalcular(sid, sesion.solicitar_recalculo(sid))
    assert (get_job_dir(sid) / "resultado.xlsx").exists()
    instancias.clear()

    sesion.quitar_txt(sid, "fac.txt")
    assert sesion.listar_txt(sid) == []
    assert not (get_job_dir(sid) / "txt" / "fac.txt").exists()
    gen = sesion.solicitar_recalculo(sid)
    sesion.recalcular(sid, gen)
    s = read_status(sid)
    assert s["estado"] == "done", s
    t = _tablero(sid)
    assert t["tarjetas"] == [] and t["excepciones"] == [] and t["periodo"] is None
    assert t["generacion"] == gen
    assert instancias == []  # Odoo no se invoca
    assert (get_job_dir(sid) / "observaciones.xlsx").exists()
    assert not (get_job_dir(sid) / "resultado.xlsx").exists()  # sin Excel completo desfasado


def test_recalculo_superado_no_pisa_el_nuevo(entorno):
    sesion, _, lector, _, _ = entorno
    entro, soltar = threading.Event(), threading.Event()
    original = lector.fetch_compras

    def bloqueante(desde, hasta, margen_dias=5):
        if not entro.is_set():  # solo la primera lectura (generación 1) se bloquea
            entro.set()
            assert soltar.wait(10)
        return original(desde, hasta, margen_dias)
    lector.fetch_compras = bloqueante

    sid = sesion.crear_sesion("compras")
    sesion.agregar_txt(sid, "a.txt", _txt(1))
    g1 = sesion.solicitar_recalculo(sid)
    hilo = threading.Thread(target=sesion.recalcular, args=(sid, g1), daemon=True)
    hilo.start()
    assert entro.wait(10)

    sesion.agregar_txt(sid, "b.txt", _txt(2, 3))
    g2 = sesion.solicitar_recalculo(sid)
    assert g2 == g1 + 1
    sesion.recalcular(sid, g2)
    nuevo = _tablero(sid)
    estado_nuevo = read_status(sid)
    assert nuevo["generacion"] == g2 and estado_nuevo["estado"] == "done"
    assert [c["total"] for c in nuevo["tarjetas"] if c["tipo"] == "FAC"] == [3]

    soltar.set()
    hilo.join(10)
    assert not hilo.is_alive()
    assert lector.lecturas == 2  # la generación 1 sí terminó de leer Odoo antes de cancelarse
    assert _tablero(sid) == nuevo
    assert read_status(sid)["resumen"] == estado_nuevo["resumen"]
    assert read_status(sid)["estado"] == "done"
    restos = [p.name for p in get_job_dir(sid).iterdir() if "tmp" in p.name or p.name.startswith("_g")]
    assert restos == []


def test_limite_de_txt_por_modo(entorno):
    sesion, *_ = entorno
    sid = sesion.crear_sesion("compras")
    for i in range(5):
        sesion.agregar_txt(sid, f"t{i}.txt", _txt(i + 1))
    with pytest.raises(ValueError):
        sesion.agregar_txt(sid, "t5.txt", _txt(9))
    assert len(sesion.listar_txt(sid)) == 5
    assert not (get_job_dir(sid) / "txt" / "t5.txt").exists()
    sesion.agregar_txt(sid, "t0.txt", _txt(7))  # reemplazar no cuenta para el límite
    assert len(sesion.listar_txt(sid)) == 5
    assert sesion.LIMITES == {"compras": 5, "ventas": 400}


def test_modo_invalido(entorno):
    sesion, *_ = entorno
    with pytest.raises(ValueError):
        sesion.crear_sesion("retenciones")


def test_ventas_cuenta_tipos_del_lector(entorno):
    sesion, *_ = entorno
    sid = sesion.crear_sesion("ventas")
    cols = ["COMPROBANTE", "SERIE_COMPROBANTE", "CLAVE_ACCESO", "FECHA_AUTORIZACION", "FECHA_EMISION",
            "VALOR_SIN_IMPUESTOS", "IVA", "IMPORTE_TOTAL"]
    filas = [["Factura", "001-001-000000001", "1" * 49, "05/09/2026", "05/09/2026", "100", "15", "115"],
             ["Comprobante de Retención", "001-200-000000001", "2" * 49, "05/09/2026", "05/09/2026",
              "", "", ""]]
    contenido = "\n".join(["\t".join(cols)] + ["\t".join(f) for f in filas]).encode("iso-8859-1")
    r = sesion.agregar_txt(sid, "semana.txt", contenido)
    assert r["estado"] == "ok"
    assert r["tipos"] == {"FAC": 1, "NC": 0, "ND": 0, "LIQ": 0, "RET": 1} and r["filas"] == 2
