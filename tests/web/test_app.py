import pytest
import bcrypt
import yaml
from pathlib import Path
from fastapi.testclient import TestClient

FIXTURES = Path(__file__).parent.parent


@pytest.fixture
def users_yaml(tmp_path):
    hashed = bcrypt.hashpw(b"testpass", bcrypt.gensalt()).decode()
    data = {"users": [{"username": "testuser", "password_hash": hashed}]}
    f = tmp_path / "users.yaml"
    f.write_text(yaml.dump(data))
    return f


@pytest.fixture
def client(users_yaml, tmp_path, monkeypatch):
    monkeypatch.setenv("USERS_FILE", str(users_yaml))
    monkeypatch.setattr("app.job_manager.JOBS_DIR", tmp_path / "jobs")
    import app.web.main as app_module
    monkeypatch.setattr(app_module, "USERS_FILE", str(users_yaml))
    return TestClient(app_module.app, raise_server_exceptions=False, follow_redirects=False)


def _login(client) -> dict:
    """Hace POST /login y devuelve las cookies de sesión."""
    r = client.post("/login", data={"username": "testuser", "password": "testpass"})
    assert r.status_code == 302
    return dict(r.cookies)


def test_root_redirects_when_no_session(client):
    r = client.get("/")
    assert r.status_code == 302
    assert "/login" in r.headers["location"]


def test_login_page_returns_html(client):
    r = client.get("/login")
    assert r.status_code == 200
    assert "text/html" in r.headers["content-type"]


def test_login_bad_credentials_redirects_with_error(client):
    r = client.post("/login", data={"username": "testuser", "password": "wrong"})
    assert r.status_code == 302
    assert "error=1" in r.headers["location"]


def test_root_returns_html_with_valid_session(client):
    cookies = _login(client)
    r = client.get("/", cookies=cookies)
    assert r.status_code == 200
    assert "text/html" in r.headers["content-type"]


def test_status_unknown_job(client):
    cookies = _login(client)
    r = client.get("/status/xxxxxxxx", cookies=cookies)
    assert r.status_code == 404


def test_download_unknown_job(client):
    cookies = _login(client)
    r = client.get("/download/xxxxxxxx", cookies=cookies)
    assert r.status_code == 404


def test_logout_clears_session(client):
    cookies = _login(client)
    r = client.get("/logout", cookies=cookies)
    assert r.status_code == 302
    # Después del logout, / debe redirigir a /login
    r2 = client.get("/")
    assert r2.status_code == 302
    assert "/login" in r2.headers["location"]


def test_download_ventas_name(client):
    from app.job_manager import create_job, get_job_dir
    cookies = _login(client)
    job_id = create_job()
    (get_job_dir(job_id) / "resultado.xlsx").write_bytes(b"x")
    r = client.get(f"/download/{job_id}", cookies=cookies)
    assert r.status_code == 200

# ------------------------------------------------------------------ API de la sesión

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@pytest.fixture
def entorno_sesion(client, tmp_path, monkeypatch):
    """Cliente con Odoo, descarga de XML, LibreOffice y limpieza falsos (sin red ni Odoo real)."""
    import pandas as pd
    import app.odoo.client as odoo_client
    import app.compras.pipeline_compras as pc
    import app.sesion as sesion

    class Lector:
        def fetch_compras(self, desde, hasta, margen_dias=5):
            return pd.DataFrame(columns=odoo_client.COLUMNAS)

        def fetch_retenciones(self, desde, hasta, margen_dias=5):
            return pd.DataFrame(columns=odoo_client.COLUMNAS_RET)

    programadas = []
    monkeypatch.setattr(sesion, "schedule_cleanup", lambda *a, **k: programadas.append(a))
    monkeypatch.setattr("shutil.which", lambda *_: None)
    monkeypatch.setattr(pc, "CACHE_XML_DIR", tmp_path / "cache" / "xml")
    monkeypatch.setattr(pc, "descargar_xmls", lambda *a, **k: [])
    monkeypatch.setattr(pc.OdooReader, "from_env", classmethod(lambda cls: Lector()))
    return client, _login(client), programadas


def _crear(client, cookies, modo="compras"):
    r = client.post("/sesion", cookies=cookies, json={"modo": modo})
    assert r.status_code == 200, r.text
    return r.json()["sesion_id"]


def _subir(client, cookies, sid, *archivos):
    return client.post(f"/sesion/{sid}/txt", cookies=cookies,
                       files=[("archivos", (n, c, "text/plain")) for n, c in archivos])


def _esperar_done(client, cookies, sid, timeout=30):
    import time
    limite = time.monotonic() + timeout
    while time.monotonic() < limite:
        s = client.get(f"/status/{sid}", cookies=cookies).json()
        if s["estado"] in ("done", "error"):
            return s
        time.sleep(0.05)
    raise AssertionError("el recálculo no terminó")


def _conciliar(client, cookies, sid):
    r = client.post(f"/sesion/{sid}/conciliar", cookies=cookies)
    assert r.status_code == 200
    gen = r.json()["generacion"]
    s = _esperar_done(client, cookies, sid)
    assert s["estado"] == "done", s
    return gen


def test_flujo_crear_subir_conciliar_tablero(entorno_sesion):
    from test_sesion import _txt
    client, cookies, _ = entorno_sesion
    sid = _crear(client, cookies)
    r = _subir(client, cookies, sid, ("facturas.txt", _txt(1, 2)))
    assert r.status_code == 200
    cuerpo = r.json()
    assert [a["nombre"] for a in cuerpo["archivos"]] == ["facturas.txt"]
    assert cuerpo["archivos"][0]["estado"] == "ok"
    assert cuerpo["generacion"] == 0  # subir no recalcula
    gen = _conciliar(client, cookies, sid)
    assert gen == 1
    t = client.get(f"/tablero/{sid}", cookies=cookies)
    assert t.status_code == 200
    assert t.json()["tarjetas"] and t.json()["generacion"] == 1


def test_tablero_404_antes_de_conciliar(entorno_sesion):
    client, cookies, _ = entorno_sesion
    sid = _crear(client, cookies)
    assert client.get(f"/tablero/{sid}", cookies=cookies).status_code == 404


def test_subir_sin_sesion_da_401(client):
    assert client.post("/sesion", json={"modo": "compras"}).status_code == 401
    r = client.post("/sesion/abcdef01/txt", files=[("archivos", ("a.txt", b"x", "text/plain"))])
    assert r.status_code == 401
    assert client.delete("/sesion/abcdef01/txt/a.txt").status_code == 401
    assert client.post("/sesion/abcdef01/conciliar").status_code == 401
    assert client.get("/tablero/abcdef01").status_code == 401
    assert client.get("/download/abcdef01").status_code == 401


def test_modo_invalido_se_rechaza(entorno_sesion):
    client, cookies, _ = entorno_sesion
    assert client.post("/sesion", cookies=cookies, json={"modo": "otro"}).status_code in (400, 422)
    assert client.post("/sesion", cookies=cookies, json={}).status_code in (400, 422)


def test_nombre_con_ruta_se_sanea(entorno_sesion):
    from test_sesion import _txt
    from app.job_manager import get_job_dir
    client, cookies, _ = entorno_sesion
    sid = _crear(client, cookies)
    r = _subir(client, cookies, sid, ("../../x.txt", _txt(1)), ("..\\..\\y.TXT", _txt(2)))
    assert r.status_code == 200
    assert [a["nombre"] for a in r.json()["archivos"]] == ["x.txt", "y.txt"]
    assert sorted(p.name for p in (get_job_dir(sid) / "txt").iterdir()) == ["x.txt", "y.txt"]
    assert not (get_job_dir(sid).parent.parent / "x.txt").exists()


def test_rechaza_archivo_no_txt(entorno_sesion):
    from test_sesion import _txt
    client, cookies, _ = entorno_sesion
    sid = _crear(client, cookies)
    r = _subir(client, cookies, sid, ("ok.txt", _txt(1)), ("odoo.xlsx", b"x"))
    assert r.status_code == 400
    # nada se guardó: se valida todo el envío antes de tocar la sesión
    assert client.delete(f"/sesion/{sid}/txt/ok.txt", cookies=cookies).json()["archivos"] == []


def test_limite_de_txt_da_400(entorno_sesion):
    from test_sesion import _txt
    client, cookies, _ = entorno_sesion
    sid = _crear(client, cookies)  # compras: máximo 5
    r = _subir(client, cookies, sid, *[(f"f{i}.txt", _txt(i + 1)) for i in range(5)])
    assert r.status_code == 200 and len(r.json()["archivos"]) == 5
    antes = client.delete(f"/sesion/{sid}/txt/nada.txt", cookies=cookies).json()["archivos"]
    r = _subir(client, cookies, sid, ("f6.txt", _txt(9)))
    assert r.status_code == 400
    assert "5" in r.json()["detail"]
    assert client.delete(f"/sesion/{sid}/txt/nada.txt", cookies=cookies).json()["archivos"] == antes
    # reemplazar uno existente sigue permitido
    assert _subir(client, cookies, sid, ("f0.txt", _txt(1, 2))).status_code == 200


def test_lote_con_mas_archivos_que_el_limite_se_rechaza_antes_de_leer(entorno_sesion):
    from test_sesion import _txt
    client, cookies, _ = entorno_sesion
    sid = _crear(client, cookies)
    _subir(client, cookies, sid, ("a.txt", _txt(1)))
    r = _subir(client, cookies, sid, *[(f"n{i}.txt", _txt(i + 2)) for i in range(6)])
    assert r.status_code == 400
    assert [a["nombre"] for a in client.delete(f"/sesion/{sid}/txt/nada.txt",
                                               cookies=cookies).json()["archivos"]] == ["a.txt"]


def test_lote_que_no_cabe_no_guarda_nada(entorno_sesion):
    from test_sesion import _txt
    client, cookies, _ = entorno_sesion
    sid = _crear(client, cookies)
    _subir(client, cookies, sid, *[(f"f{i}.txt", _txt(i + 1)) for i in range(4)])
    r = _subir(client, cookies, sid, ("x.txt", _txt(7)), ("y.txt", _txt(8)))  # 4 + 2 > 5
    assert r.status_code == 400
    nombres = [a["nombre"] for a in client.delete(f"/sesion/{sid}/txt/nada.txt",
                                                  cookies=cookies).json()["archivos"]]
    assert nombres == [f"f{i}.txt" for i in range(4)]


def test_reemplazo_dentro_de_un_lote_no_cuenta_contra_el_limite(entorno_sesion):
    from test_sesion import _txt
    client, cookies, _ = entorno_sesion
    sid = _crear(client, cookies)
    _subir(client, cookies, sid, *[(f"f{i}.txt", _txt(i + 1)) for i in range(4)])
    # f0 ya existe (reemplazo) y z se repite dentro del lote (cuenta una vez): 4 + 1 = 5
    r = _subir(client, cookies, sid, ("f0.txt", _txt(1, 2)), ("z.txt", _txt(5)),
               ("z.TXT", _txt(7)))
    assert r.status_code == 200, r.text


def test_archivo_sobre_el_maximo_da_413(entorno_sesion, monkeypatch):
    import app.web.main as app_module
    client, cookies, _ = entorno_sesion
    monkeypatch.setattr(app_module, "MAX_BYTES_TXT", 100)
    sid = _crear(client, cookies)
    assert _subir(client, cookies, sid, ("g.txt", b"x" * 101)).status_code == 413
    assert _subir(client, cookies, sid, ("p.txt", b"x" * 50)).status_code == 200
    monkeypatch.setattr(app_module, "MAX_BYTES_LOTE", 120)
    assert _subir(client, cookies, sid, ("a.txt", b"x" * 80), ("b.txt", b"x" * 80)).status_code == 413


def test_subida_masiva_programa_una_sola_limpieza(entorno_sesion):
    client, cookies, programadas = entorno_sesion
    sid = _crear(client, cookies, "ventas")
    antes = len(programadas)
    r = _subir(client, cookies, sid, *[(f"e{i}.txt", b"x") for i in range(40)])
    assert r.status_code == 200 and len(r.json()["archivos"]) == 40
    assert len(programadas) - antes == 1


def test_quitar_txt(entorno_sesion):
    from test_sesion import _txt
    client, cookies, _ = entorno_sesion
    sid = _crear(client, cookies)
    _subir(client, cookies, sid, ("a.txt", _txt(1)), ("b.txt", _txt(2)))
    r = client.delete(f"/sesion/{sid}/txt/a.txt", cookies=cookies)
    assert r.status_code == 200
    assert [a["nombre"] for a in r.json()["archivos"]] == ["b.txt"]


def test_descargar_observaciones_y_completo(entorno_sesion):
    from test_sesion import _txt
    client, cookies, _ = entorno_sesion
    sid = _crear(client, cookies)
    _subir(client, cookies, sid, ("facturas.txt", _txt(1, 2)))
    _conciliar(client, cookies, sid)
    for url, nombre in ((f"/download/{sid}?tipo=observaciones", "Observaciones_SRI.xlsx"),
                        (f"/download/{sid}?tipo=completo", "Conciliacion_SRI.xlsx"),
                        (f"/download/{sid}", "Conciliacion_SRI.xlsx")):
        r = client.get(url, cookies=cookies)
        assert r.status_code == 200, url
        assert r.headers["content-type"].startswith(XLSX)
        assert nombre in r.headers["content-disposition"]
        assert r.content[:2] == b"PK"  # un .xlsx es un zip
    assert client.get(f"/download/{sid}?tipo=otro", cookies=cookies).status_code in (400, 422)


def test_descarga_completo_404_sin_txt_validos(entorno_sesion):
    client, cookies, _ = entorno_sesion
    sid = _crear(client, cookies)
    _conciliar(client, cookies, sid)  # sin TXT: tablero vacío, sin Excel completo
    assert client.get(f"/tablero/{sid}", cookies=cookies).status_code == 200
    assert client.get(f"/download/{sid}", cookies=cookies).status_code == 404
    assert client.get(f"/download/{sid}?tipo=observaciones", cookies=cookies).status_code == 200


def test_sesion_inexistente_404(entorno_sesion):
    from test_sesion import _txt
    client, cookies, _ = entorno_sesion
    for sid in ("abcdef01", "xxxxxxxx", "ABCDEF01", "abcdef012"):
        assert _subir(client, cookies, sid, ("a.txt", _txt(1))).status_code == 404, sid
        assert client.delete(f"/sesion/{sid}/txt/a.txt", cookies=cookies).status_code == 404, sid
        assert client.post(f"/sesion/{sid}/conciliar", cookies=cookies).status_code == 404, sid
        assert client.get(f"/tablero/{sid}", cookies=cookies).status_code == 404, sid
        assert client.get(f"/download/{sid}", cookies=cookies).status_code == 404, sid
    assert client.get("/tablero/%2e%2e", cookies=cookies).status_code == 404


def test_index_menciona_retenciones_en_compras(client):
    cookies = _login(client)
    html = client.get("/", cookies=cookies).text
    assert "facturas, notas de crédito, notas de débito y retenciones" in html
    assert "cualquier subconjunto" in html


def test_index_ventas_informativo_de_retenciones_es_neutro(client):
    import re
    from app.retenciones.conciliar_retenciones import EST_INFO
    cookies = _login(client)
    html = client.get("/", cookies=cookies).text
    neutros = re.search(r"const neutros = new Set\(\[(.*?)\]\);", html).group(1)
    assert f'"{EST_INFO}"' in neutros
