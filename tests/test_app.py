import pytest
import bcrypt
import yaml
from pathlib import Path
from fastapi.testclient import TestClient
import base64

FIXTURES = Path(__file__).parent.parent


def _auth(username="testuser", password="testpass"):
    token = base64.b64encode(f"{username}:{password}".encode()).decode()
    return {"Authorization": f"Basic {token}"}


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
    return TestClient(app_module.app, raise_server_exceptions=False)


def test_root_requires_auth(client):
    r = client.get("/", auth=None)
    assert r.status_code == 401


def test_root_returns_html(client):
    r = client.get("/", auth=("testuser", "testpass"))
    assert r.status_code == 200
    assert "text/html" in r.headers["content-type"]


def test_upload_requires_auth(client):
    r = client.post("/upload")
    assert r.status_code == 401


def test_upload_missing_files(client):
    r = client.post("/upload", auth=("testuser", "testpass"))
    assert r.status_code == 422   # FastAPI validation error


def test_status_unknown_job(client):
    r = client.get("/status/xxxxxxxx", auth=("testuser", "testpass"))
    assert r.status_code == 404


def test_download_unknown_job(client):
    r = client.get("/download/xxxxxxxx", auth=("testuser", "testpass"))
    assert r.status_code == 404


def test_upload_creates_job(client, tmp_path, monkeypatch):
    """Upload con archivos válidos debe retornar job_id y estado pending."""
    import app as app_module
    # Mock run_pipeline para no ejecutar el pipeline real
    monkeypatch.setattr(app_module, "run_pipeline", lambda job_id: None)

    txt_content = (FIXTURES / "2390000831001_Recibidos.txt").read_bytes()
    xlsx_content = (FIXTURES / "reten_odoo.xlsx").read_bytes()

    r = client.post(
        "/upload",
        auth=("testuser", "testpass"),
        files={
            "recibidos": ("Recibidos.txt", txt_content, "text/plain"),
            "reten_odoo": ("reten_odoo.xlsx", xlsx_content,
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        },
    )
    assert r.status_code == 200
    data = r.json()
    assert "job_id" in data
    assert len(data["job_id"]) == 8

    # Status debe existir
    job_id = data["job_id"]
    r2 = client.get(f"/status/{job_id}", auth=("testuser", "testpass"))
    assert r2.status_code == 200
    assert r2.json()["estado"] in ("pending", "running", "done")
