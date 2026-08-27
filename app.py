"""
Aplicación FastAPI para el pipeline de conciliación SRI.
Endpoints: GET /, POST /upload, GET /status/{job_id}, GET /download/{job_id}
Auth: HTTP Basic Auth contra users.yaml (bcrypt)
"""
import threading
from pathlib import Path

from fastapi import FastAPI, Depends, HTTPException, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.security import HTTPBasic, HTTPBasicCredentials

from auth import verify_user
from job_manager import create_job, read_status, get_job_dir, mark_stale_jobs_as_error
from pipeline import run_pipeline

BASE_DIR    = Path(__file__).parent
USERS_FILE  = str(BASE_DIR / "users.yaml")
TEMPLATE    = BASE_DIR / "templates" / "index.html"

app = FastAPI(title="Conciliación SRI", docs_url=None, redoc_url=None)
security = HTTPBasic()


def get_user(credentials: HTTPBasicCredentials = Depends(security)) -> str:
    if not verify_user(credentials.username, credentials.password, users_file=USERS_FILE):
        raise HTTPException(
            status_code=401,
            detail="Credenciales inválidas",
            headers={"WWW-Authenticate": "Basic"},
        )
    return credentials.username


@app.on_event("startup")
def startup():
    mark_stale_jobs_as_error()


@app.get("/", response_class=HTMLResponse)
def index(_user: str = Depends(get_user)):
    return TEMPLATE.read_text(encoding="utf-8")


@app.post("/upload")
def upload(
    recibidos: UploadFile = File(...),
    reten_odoo: UploadFile = File(...),
    _user: str = Depends(get_user),
):
    job_id  = create_job()
    job_dir = get_job_dir(job_id)

    (job_dir / "Recibidos.txt").write_bytes(recibidos.file.read())
    (job_dir / "reten_odoo.xlsx").write_bytes(reten_odoo.file.read())

    t = threading.Thread(target=run_pipeline, args=(job_id,), daemon=True)
    t.start()

    return {"job_id": job_id}


@app.get("/status/{job_id}")
def status(job_id: str, _user: str = Depends(get_user)):
    s = read_status(job_id)
    if s is None:
        raise HTTPException(status_code=404, detail="Job no encontrado")
    return s


@app.get("/download/{job_id}")
def download(job_id: str, _user: str = Depends(get_user)):
    excel = get_job_dir(job_id) / "resultado.xlsx"
    if not excel.exists():
        raise HTTPException(status_code=404, detail="Resultado no disponible aún")
    return FileResponse(
        path=str(excel),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="Conciliacion_SRI.xlsx",
    )
