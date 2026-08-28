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


def test_upload_returns_401_without_session(client):
    r = client.post("/upload")
    assert r.status_code == 401


def test_upload_missing_files_with_session(client):
    cookies = _login(client)
    r = client.post("/upload", cookies=cookies)
    assert r.status_code == 422


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


def test_upload_creates_job(client, tmp_path, monkeypatch):
    """Upload con archivos válidos debe retornar job_id y estado pending."""
    import app as app_module
    monkeypatch.setattr(app_module, "run_pipeline", lambda job_id: None)

    cookies = _login(client)
    txt_content  = (FIXTURES / "2390000831001_Recibidos.txt").read_bytes()
    xlsx_content = (FIXTURES / "reten_odoo.xlsx").read_bytes()

    r = client.post(
        "/upload",
        cookies=cookies,
        files={
            "recibidos":   ("Recibidos.txt",  txt_content,  "text/plain"),
            "reten_odoo":  ("reten_odoo.xlsx", xlsx_content,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        },
    )
    assert r.status_code == 200
    data = r.json()
    assert "job_id" in data
    assert len(data["job_id"]) == 8

    job_id = data["job_id"]
    r2 = client.get(f"/status/{job_id}", cookies=cookies)
    assert r2.status_code == 200
    assert r2.json()["estado"] in ("pending", "running", "done")
