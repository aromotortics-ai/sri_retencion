"""
Aplicación FastAPI de la conciliación SRI vs Odoo (tablero en vivo).
Endpoints: GET/POST /login, GET /logout, GET /,
sesión de trabajo: POST /sesion {modo}, POST /sesion/{id}/txt (campo `archivos`),
DELETE /sesion/{id}/txt/{nombre}, POST /sesion/{id}/conciliar,
resultado: GET /status/{id}, GET /tablero/{id}, GET /download/{id}?tipo=completo|observaciones
Auth: cookie de sesión firmada con itsdangerous (SESSION_HOURS, default 8 h); las rutas de API
responden 401 JSON sin sesión. Los ids de sesión son de 8 hex (create_job) y se validan antes de
tocar el disco.
"""
import os
import re
import secrets
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Literal

from fastapi import FastAPI, Depends, HTTPException, Request, UploadFile, File
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from pydantic import BaseModel
from itsdangerous import TimestampSigner, BadSignature, SignatureExpired

from app.rutas import DATA_DIR
from app.web.auth import verify_user
from app import sesion
from app.job_manager import read_status, get_job_dir, mark_stale_jobs_as_error

TEMPLATES_DIR  = Path(__file__).parent / "templates"
USERS_FILE     = str(DATA_DIR / "users.yaml")
TEMPLATE       = TEMPLATES_DIR / "index.html"
LOGIN_TEMPLATE = TEMPLATES_DIR / "login.html"

# Clave secreta para firmar cookies — leer de env var en producción.
# Si no se configura, se genera una nueva en cada arranque (invalida sesiones activas).
SECRET_KEY    = os.environ.get("SECRET_KEY") or secrets.token_hex(32)
SESSION_HOURS = int(os.environ.get("SESSION_HOURS", "8"))
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


_ID_RE = re.compile(r"[0-9a-f]{8}")  # formato de create_job: evita path traversal
_PUBLICADOS = {"tablero": "tablero.json", "completo": "resultado.xlsx",
               "observaciones": "observaciones.xlsx"}
_DESCARGAS = {"completo": "Conciliacion_SRI.xlsx", "observaciones": "Observaciones_SRI.xlsx"}
MAX_BYTES_TXT = 20 * 1024 * 1024    # por archivo (un TXT diario del portal pesa mucho menos)
MAX_BYTES_LOTE = 200 * 1024 * 1024  # por solicitud
_CHUNK = 1024 * 1024
SLACK_MULTIPART = 1024 * 1024       # sobrecarga del multipart sobre MAX_BYTES_LOTE
_RUTA_SUBIDA = re.compile(r"/sesion/[^/]+/txt")
_MSG_413 = b'{"detail":"Archivo o env\\u00edo demasiado grande"}'


class _CuerpoExcesivo(Exception):
    pass


class LimiteSubida:
    """Guardia ASGI: corta la subida de TXT antes de que se analice (y se guarde) el cuerpo.

    Usa Content-Length si viene; si no (chunked), cuenta los bytes mientras llegan.
    Los topes finos por archivo siguen en el handler."""

    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if (scope["type"] != "http" or scope["method"] != "POST"
                or not _RUTA_SUBIDA.fullmatch(scope["path"])):
            return await self.app(scope, receive, send)
        limite = MAX_BYTES_LOTE + SLACK_MULTIPART
        cl = dict(scope["headers"]).get(b"content-length")
        try:
            if cl is not None and int(cl) > limite:
                return await self._rechazar(send)
        except ValueError:
            return await self._rechazar(send, 400)
        recibidos = 0
        iniciada = False

        async def receive_limitado():
            nonlocal recibidos
            msg = await receive()
            if msg["type"] == "http.request":
                recibidos += len(msg.get("body", b""))
                if recibidos > limite:
                    raise _CuerpoExcesivo
            return msg

        async def send_visto(msg):
            nonlocal iniciada
            if msg["type"] == "http.response.start":
                iniciada = True
            await send(msg)

        try:
            await self.app(scope, receive_limitado, send_visto)
        except _CuerpoExcesivo:
            if not iniciada:
                await self._rechazar(send)

    @staticmethod
    async def _rechazar(send, status=413):
        cuerpo = _MSG_413 if status == 413 else b'{"detail":"Content-Length inv\\u00e1lido"}'
        await send({"type": "http.response.start", "status": status,
                    "headers": [(b"content-type", b"application/json"),
                                (b"content-length", str(len(cuerpo)).encode()),
                                (b"connection", b"close")]})
        await send({"type": "http.response.body", "body": cuerpo})


app.add_middleware(LimiteSubida)
XLSX_MEDIA = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


class NuevaSesion(BaseModel):
    modo: Literal["compras", "ventas"]


def _sesion_o_404(sesion_id: str) -> str:
    """Valida el formato del id antes de tocar el disco y que la sesión exista."""
    if not _ID_RE.fullmatch(sesion_id) or not sesion.existe(sesion_id):
        raise HTTPException(status_code=404, detail="Sesión no encontrada o expirada")
    return sesion_id


def _ocupada(exc: TimeoutError) -> HTTPException:
    return HTTPException(status_code=503, detail=str(exc))


@app.post("/sesion")
def crear_sesion(datos: NuevaSesion, _user: str = Depends(require_session)):
    return {"sesion_id": sesion.crear_sesion(datos.modo)}


@app.post("/sesion/{sesion_id}/txt")
async def subir_txt(
    sesion_id: str,
    archivos: list[UploadFile] = File(...),
    _user: str = Depends(require_session),
):
    _sesion_o_404(sesion_id)
    nombres = [Path((f.filename or "").replace("\\", "/")).name for f in archivos]
    if any(not n.lower().endswith(".txt") for n in nombres):
        raise HTTPException(status_code=400, detail="Todos los archivos deben ser .txt")
    modo = sesion.modo_de(sesion_id)
    if len(archivos) > sesion.LIMITES[modo]:  # primera barrera, antes de leer bytes
        raise HTTPException(status_code=400,
                            detail=f"Se admiten como máximo {sesion.LIMITES[modo]} TXT en {modo}")
    lote, total = [], 0
    for n, f in zip(nombres, archivos):
        partes, tam = [], 0
        while chunk := await f.read(_CHUNK):
            tam += len(chunk)
            total += len(chunk)
            if tam > MAX_BYTES_TXT or total > MAX_BYTES_LOTE:
                raise HTTPException(status_code=413, detail="Archivo o envío demasiado grande")
            partes.append(chunk)
        lote.append((n, b"".join(partes)))
    try:
        resultados = await run_in_threadpool(sesion.agregar_txts, sesion_id, lote)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except KeyError:
        raise HTTPException(status_code=404, detail="Sesión no encontrada o expirada")
    except TimeoutError as exc:
        raise _ocupada(exc)
    # No recalcula: el cliente pide POST /sesion/{id}/conciliar cuando quiere.
    return {"archivos": resultados, "generacion": sesion.generacion_actual(sesion_id)}


@app.delete("/sesion/{sesion_id}/txt/{nombre}")
def quitar_txt(sesion_id: str, nombre: str, _user: str = Depends(require_session)):
    _sesion_o_404(sesion_id)
    try:
        sesion.quitar_txt(sesion_id, nombre)
        return {"archivos": sesion.listar_txt(sesion_id)}
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except KeyError:
        raise HTTPException(status_code=404, detail="Sesión no encontrada o expirada")
    except TimeoutError as exc:
        raise _ocupada(exc)


@app.post("/sesion/{sesion_id}/conciliar")
def conciliar(sesion_id: str, _user: str = Depends(require_session)):
    _sesion_o_404(sesion_id)
    try:
        return {"generacion": sesion.solicitar_recalculo(sesion_id)}
    except KeyError:
        raise HTTPException(status_code=404, detail="Sesión no encontrada o expirada")
    except TimeoutError as exc:
        raise _ocupada(exc)


@app.get("/tablero/{sesion_id}")
def tablero(sesion_id: str, _user: str = Depends(require_session)):
    _sesion_o_404(sesion_id)
    try:
        datos = sesion.leer_publicado(sesion_id, _PUBLICADOS["tablero"])
    except TimeoutError as exc:
        raise _ocupada(exc)
    if datos is None:
        raise HTTPException(status_code=404, detail="Tablero aún no disponible")
    return Response(content=datos, media_type="application/json")


@app.get("/status/{job_id}")
def status(job_id: str, _user: str = Depends(require_session)):
    s = read_status(job_id)
    if s is None:
        raise HTTPException(status_code=404, detail="Job no encontrado")
    return s


@app.get("/download/{job_id}")
def download(job_id: str, tipo: Literal["completo", "observaciones"] = "completo",
             _user: str = Depends(require_session)):
    if not _ID_RE.fullmatch(job_id) or not get_job_dir(job_id).is_dir():
        raise HTTPException(status_code=404, detail="Resultado no disponible aún")
    try:
        datos = sesion.leer_publicado(job_id, _PUBLICADOS[tipo])
    except TimeoutError as exc:
        raise _ocupada(exc)
    if datos is None:
        raise HTTPException(status_code=404, detail="Resultado no disponible aún")
    return Response(
        content=datos, media_type=XLSX_MEDIA,
        headers={"Content-Disposition": f'attachment; filename="{_DESCARGAS[tipo]}"'},
    )
