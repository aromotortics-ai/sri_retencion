"""
Aplicación FastAPI para el pipeline de conciliación SRI.
Endpoints: GET/POST /login, GET /logout, GET /, POST /upload-ventas (emitidos), POST /upload-compras (recibidos),
GET /status/{job_id}, GET /download/{job_id}
Auth: cookie de sesión firmada con itsdangerous (SESSION_HOURS, default 8 h)
"""
import os
import secrets
import threading
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, Depends, HTTPException, Request, UploadFile, File
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse
from itsdangerous import TimestampSigner, BadSignature, SignatureExpired

from auth import verify_user
from job_manager import create_job, read_status, get_job_dir, mark_stale_jobs_as_error
from pipeline_ventas import run_pipeline_ventas
from pipeline_compras import run_pipeline_compras

BASE_DIR       = Path(__file__).parent
USERS_FILE     = str(BASE_DIR / "data" / "users.yaml")
TEMPLATE       = BASE_DIR / "templates" / "index.html"
LOGIN_TEMPLATE = BASE_DIR / "templates" / "login.html"

# Clave secreta para firmar cookies — leer de env var en producción.
# Si no se configura, se genera una nueva en cada arranque (invalida sesiones activas).
SECRET_KEY    = os.environ.get("SECRET_KEY") or secrets.token_hex(32)
SESSION_HOURS = int(os.environ.get("SESSION_HOURS", "8"))
MAX_TXT_VENTAS = 400  # TXT por envío en Ventas (un mes de TXT diarios de varios tipos cabe de sobra)
COOKIE_NAME   = "sri_session"


def _make_token(username: str) -> str:
    return TimestampSigner(SECRET_KEY).sign(username).decode()


def _verify_token(token: str) -> str | None:
    try:
        return TimestampSigner(SECRET_KEY).unsign(
            token, max_age=SESSION_HOURS * 3600
        ).decode()
    except (BadSignature, SignatureExpired):
        return None


def _current_user(request: Request) -> str | None:
    token = request.cookies.get(COOKIE_NAME)
    return _verify_token(token) if token else None


def _nombres_txt(archivos: list[UploadFile]) -> list[str]:
    """Basename de cada archivo (evita path traversal); 400 si alguno no es .txt."""
    nombres = [Path((f.filename or "").replace("\\", "/")).name for f in archivos]
    if any(not n.lower().endswith(".txt") for n in nombres):
        raise HTTPException(status_code=400, detail="Todos los archivos deben ser .txt")
    return nombres


async def _guardar_txts(archivos: list[UploadFile], nombres: list[str], carpeta: Path) -> None:
    """Guarda los TXT con la extensión normalizada a `.txt` minúscula (los pipelines usan
    glob("*.txt"), sensible a mayúsculas en Linux); un nombre repetido lleva prefijo numérico."""
    carpeta.mkdir()
    for i, (f, nombre) in enumerate(zip(archivos, nombres)):
        nombre = Path(nombre).stem + ".txt"
        dest = carpeta / nombre
        if dest.exists():
            dest = carpeta / f"{i}_{nombre}"
        dest.write_bytes(await f.read())


def require_session(request: Request) -> str:
    """Dependencia para endpoints de API: devuelve 401 JSON si no hay sesión."""
    user = _current_user(request)
    if not user:
        raise HTTPException(status_code=401, detail="Sesión expirada — vuelve a iniciar sesión")
    return user


@asynccontextmanager
async def lifespan(app: FastAPI):
    mark_stale_jobs_as_error()
    yield


app = FastAPI(title="Conciliación SRI", docs_url=None, redoc_url=None, lifespan=lifespan)


# ── Autenticación ──────────────────────────────────────────────────────────────

@app.get("/login", response_class=HTMLResponse)
def login_page(error: str = ""):
    html = LOGIN_TEMPLATE.read_text(encoding="utf-8")
    if error:
        html = html.replace(
            "<!-- LOGIN_ERROR -->",
            '<p class="login-error">Usuario o contraseña incorrectos.</p>',
        )
    return html


@app.post("/login")
async def login(request: Request):
    form = await request.form()
    username = str(form.get("username", ""))
    password = str(form.get("password", ""))
    if not verify_user(username, password, users_file=USERS_FILE):
        return RedirectResponse("/login?error=1", status_code=302)
    resp = RedirectResponse("/", status_code=302)
    resp.set_cookie(
        COOKIE_NAME, _make_token(username),
        httponly=True, samesite="lax",
        max_age=SESSION_HOURS * 3600,
    )
    return resp


@app.get("/logout")
def logout():
    resp = RedirectResponse("/login", status_code=302)
    resp.delete_cookie(COOKIE_NAME)
    return resp


# ── Páginas y API ──────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    if not _current_user(request):
        return RedirectResponse("/login", status_code=302)
    return TEMPLATE.read_text(encoding="utf-8")


@app.post("/upload-ventas")
async def upload_ventas(
    emitidos: list[UploadFile] = File(...),
    _user: str = Depends(require_session),
):
    if not 1 <= len(emitidos) <= MAX_TXT_VENTAS:
        raise HTTPException(status_code=400,
                            detail=f"Sube entre 1 y {MAX_TXT_VENTAS} archivos TXT del SRI")
    nombres = _nombres_txt(emitidos)

    job_id = create_job()
    await _guardar_txts(emitidos, nombres, get_job_dir(job_id) / "emitidos")

    threading.Thread(target=run_pipeline_ventas, args=(job_id,), daemon=True).start()
    return {"job_id": job_id}


@app.post("/upload-compras")
async def upload_compras(
    sri: list[UploadFile] = File(...),
    _user: str = Depends(require_session),
):
    if not 1 <= len(sri) <= 5:
        raise HTTPException(status_code=400, detail="Sube entre 1 y 5 archivos TXT del SRI")
    nombres = _nombres_txt(sri)

    job_id = create_job()
    await _guardar_txts(sri, nombres, get_job_dir(job_id) / "sri")

    threading.Thread(target=run_pipeline_compras, args=(job_id,), daemon=True).start()
    return {"job_id": job_id}


@app.get("/status/{job_id}")
def status(job_id: str, _user: str = Depends(require_session)):
    s = read_status(job_id)
    if s is None:
        raise HTTPException(status_code=404, detail="Job no encontrado")
    return s


@app.get("/download/{job_id}")
def download(job_id: str, _user: str = Depends(require_session)):
    excel = get_job_dir(job_id) / "resultado.xlsx"
    if not excel.exists():
        raise HTTPException(status_code=404, detail="Resultado no disponible aún")
    return FileResponse(
        path=str(excel),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="Conciliacion_SRI.xlsx",
    )
