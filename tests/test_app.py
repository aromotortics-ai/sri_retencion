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
    monkeypatch.setattr("job_manager.JOBS_DIR", tmp_path / "jobs")
    import app as app_module
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


def test_upload_ventas_returns_401_without_session(client):
    assert client.post("/upload-ventas").status_code == 401


def test_upload_ventas_missing_files(client):
    cookies = _login(client)
    assert client.post("/upload-ventas", cookies=cookies).status_code == 422


def test_upload_ventas_acepta_txt_con_nombre_del_portal(client, monkeypatch):
    import app as app_module
    from job_manager import get_job_dir
    monkeypatch.setattr(app_module, "run_pipeline_ventas", lambda job_id: None)

    cookies = _login(client)
    r = client.post(
        "/upload-ventas",
        cookies=cookies,
        files=[
            ("emitidos", ("2390000831001_EmitidosAutorizados (1).txt", b"a", "text/plain")),
            ("emitidos", ("2390000831001_EmitidosAutorizados (1).txt", b"b", "text/plain")),
        ],
    )
    assert r.status_code == 200
    d = get_job_dir(r.json()["job_id"]) / "emitidos"
    nombres = sorted(p.name for p in d.iterdir())
    assert "2390000831001_EmitidosAutorizados (1).txt" in nombres
    assert len(nombres) == 2  # el repetido se conserva con prefijo numérico


def test_upload_ventas_nombre_con_ruta_no_escapa_del_job(client, monkeypatch):
    import app as app_module
    from job_manager import get_job_dir
    monkeypatch.setattr(app_module, "run_pipeline_ventas", lambda job_id: None)
    cookies = _login(client)
    r = client.post("/upload-ventas", cookies=cookies,
                    files=[("emitidos", ("../../evil.txt", b"x", "text/plain"))])
    assert r.status_code == 200
    d = get_job_dir(r.json()["job_id"])
    assert [p.name for p in (d / "emitidos").iterdir()] == ["evil.txt"]


def _nombres_guardados(client, monkeypatch, ruta, campo, pipeline, archivos):
    import app as app_module
    from job_manager import get_job_dir
    monkeypatch.setattr(app_module, pipeline, lambda job_id: None)
    cookies = _login(client)
    r = client.post(ruta, cookies=cookies,
                    files=[(campo, (n, b"x", "text/plain")) for n in archivos])
    assert r.status_code == 200
    d = get_job_dir(r.json()["job_id"]) / ("emitidos" if campo == "emitidos" else "sri")
    nombres = sorted(p.name for p in d.iterdir())
    # el pipeline usa glob("*.txt"), sensible a mayúsculas en Linux
    assert all(n.endswith(".txt") for n in nombres), nombres
    assert sorted(p.name for p in d.glob("*.txt")) == nombres
    return nombres


def test_upload_ventas_normaliza_extension_mayuscula(client, monkeypatch):
    nombres = _nombres_guardados(client, monkeypatch, "/upload-ventas", "emitidos",
                                 "run_pipeline_ventas", ["SEMANA.TXT"])
    assert nombres == ["SEMANA.txt"]


def test_upload_ventas_extensiones_distintas_no_se_pisan(client, monkeypatch):
    nombres = _nombres_guardados(client, monkeypatch, "/upload-ventas", "emitidos",
                                 "run_pipeline_ventas", ["a.TXT", "a.txt"])
    assert len(nombres) == 2


def test_upload_compras_normaliza_extension_mayuscula(client, monkeypatch):
    nombres = _nombres_guardados(client, monkeypatch, "/upload-compras", "sri",
                                 "run_pipeline_compras", ["SEMANA.TXT", "b.Txt"])
    assert nombres == ["SEMANA.txt", "b.txt"]


def test_upload_ventas_rechaza_no_txt(client, monkeypatch):
    import app as app_module
    monkeypatch.setattr(app_module, "run_pipeline_ventas", lambda job_id: None)
    cookies = _login(client)
    r = client.post("/upload-ventas", cookies=cookies,
                    files=[("emitidos", ("odoo.xlsx", b"x", "application/octet-stream"))])
    assert r.status_code == 400


def test_upload_ventas_acepta_mas_de_60_archivos_y_rechaza_sobre_el_maximo(client, monkeypatch):
    import app as app_module
    monkeypatch.setattr(app_module, "run_pipeline_ventas", lambda job_id: None)
    cookies = _login(client)
    n_max = app_module.MAX_TXT_VENTAS
    mas_de_60 = [("emitidos", (f"e{i}.txt", b"x", "text/plain")) for i in range(61)]
    assert client.post("/upload-ventas", cookies=cookies, files=mas_de_60).status_code == 200
    sobre_max = [("emitidos", (f"e{i}.txt", b"x", "text/plain")) for i in range(n_max + 1)]
    assert client.post("/upload-ventas", cookies=cookies, files=sobre_max).status_code == 400


def test_upload_ventas_ya_no_pide_excel_de_odoo(client, monkeypatch):
    import app as app_module
    monkeypatch.setattr(app_module, "run_pipeline_ventas", lambda job_id: None)
    cookies = _login(client)
    r = client.post("/upload-ventas", cookies=cookies,
                    files=[("emitidos", ("a.txt", b"x", "text/plain"))])
    assert r.status_code == 200
    assert "job_id" in r.json()


def test_download_ventas_name(client):
    from job_manager import create_job, get_job_dir
    cookies = _login(client)
    job_id = create_job()
    (get_job_dir(job_id) / "resultado.xlsx").write_bytes(b"x")
    r = client.get(f"/download/{job_id}", cookies=cookies)
    assert r.status_code == 200


def test_upload_compras_requiere_sesion(client):
    r = client.post("/upload-compras", files=[("sri", ("a.txt", b"x", "text/plain"))])
    assert r.status_code == 401


def test_upload_compras_devuelve_job_id(client, monkeypatch):
    import app as app_module
    monkeypatch.setattr(app_module, "run_pipeline_compras", lambda job_id: None)
    cookies = _login(client)
    r = client.post("/upload-compras", cookies=cookies,
                    files=[("sri", ("a.txt", b"x", "text/plain")), ("sri", ("b.txt", b"y", "text/plain"))])
    assert r.status_code == 200
    from job_manager import get_job_dir
    d = get_job_dir(r.json()["job_id"]) / "sri"
    assert sorted(p.name for p in d.iterdir()) == ["a.txt", "b.txt"]


def test_upload_compras_nombre_con_ruta_no_escapa_del_job(client, monkeypatch):
    import app as app_module
    monkeypatch.setattr(app_module, "run_pipeline_compras", lambda job_id: None)
    cookies = _login(client)
    r = client.post("/upload-compras", cookies=cookies,
                    files=[("sri", ("../../evil.txt", b"x", "text/plain")),
                           ("sri", (r"..\..\evil.txt", b"y", "text/plain"))])
    assert r.status_code == 200
    from job_manager import get_job_dir
    d = get_job_dir(r.json()["job_id"])
    assert not (d.parent.parent / "evil.txt").exists()
    assert len(list((d / "sri").iterdir())) == 2


def test_upload_compras_rechaza_no_txt_mas_de_5_y_vacio(client, monkeypatch):
    import app as app_module
    monkeypatch.setattr(app_module, "run_pipeline_compras", lambda job_id: None)
    cookies = _login(client)
    r = client.post("/upload-compras", cookies=cookies, files=[("sri", ("a.xlsx", b"x", "text/plain"))])
    assert r.status_code == 400
    seis = [("sri", (f"{i}.txt", b"x", "text/plain")) for i in range(6)]
    assert client.post("/upload-compras", cookies=cookies, files=seis).status_code == 400
    assert client.post("/upload-compras", cookies=cookies).status_code in (400, 422)


def test_upload_compras_acepta_txt_de_retenciones(client, monkeypatch):
    import app as app_module
    monkeypatch.setattr(app_module, "run_pipeline_compras", lambda job_id: None)
    cookies = _login(client)
    r = client.post("/upload-compras", cookies=cookies,
                    files=[("sri", ("retenciones.txt", b"x", "text/plain"))])
    assert r.status_code == 200
    assert "job_id" in r.json()


def test_index_menciona_retenciones_en_compras(client):
    cookies = _login(client)
    html = client.get("/", cookies=cookies).text
    assert "facturas, notas de crédito, notas de débito y retenciones" in html
    assert "cualquier subconjunto" in html


def test_index_ventas_informativo_de_retenciones_es_neutro(client):
    import re
    from conciliar_retenciones import EST_INFO
    cookies = _login(client)
    html = client.get("/", cookies=cookies).text
    neutros = re.search(r"const neutros = new Set\(\[(.*?)\]\);", html).group(1)
    assert f'"{EST_INFO}"' in neutros
