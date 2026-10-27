# -*- coding: utf-8 -*-
"""Sesión de trabajo: un job (`jobs/<id>/`) que guarda los TXT del usuario y recalcula la
conciliación completa (Compras o Ventas) cada vez que se pide.

Todo el estado vive en disco porque gunicorn corre varios workers:

    jobs/<id>/txt/<nombre>.txt   TXT subidos (mismo nombre = reemplazo)
    jobs/<id>/sesion.json        {"modo", "generacion", "archivos": [...], "ultima_actividad"}
    jobs/<id>/tablero.json       resultado para la pantalla (incluye "generacion")
    jobs/<id>/resultado.xlsx     Excel completo (no existe si no hay ningún TXT válido)
    jobs/<id>/observaciones.xlsx Excel de observaciones
    jobs/<id>/status.json        estado del último recálculo (job_manager)

Cancelación: cada `solicitar_recalculo` incrementa `generacion` y lanza un hilo con ese número.
Un recálculo compara su generación con la de disco en los puntos de control del pipeline y,
justo antes de publicar, bajo el bloqueo de la sesión; si fue superado termina sin escribir nada
(sus salidas se escriben con nombres temporales y solo se renombran a los definitivos si sigue
vigente). El bloqueo es un archivo `jobs/<id>/.lock` creado con O_EXCL, así que serializa a los
hilos de un mismo proceso y a los distintos workers.
"""
import json
import os
import threading
import time
import traceback
from contextlib import contextmanager
from pathlib import Path

from app.compras.conciliar_compras import TIPOS as TIPOS_COMPRAS
from app.compras.conciliar_compras import leer_sri_txts
from app.compras.pipeline_compras import procesar_compras
from app.job_manager import create_job, get_job_dir, schedule_cleanup, write_status
from app.reportes.observaciones import escribir_excel_observaciones
from app.retenciones.conciliar_retenciones import leer_sri_retenciones_txt
from app.sri.emitidos_sri import TIPOS_DOC, leer_emitidos_txts
from app.ventas.pipeline_ventas import procesar_ventas

LIMITES = {"compras": 5, "ventas": 400}
TIPOS = {"compras": TIPOS_COMPRAS + ("RET",), "ventas": TIPOS_DOC + ("RET",)}
LIMPIEZA_SEG = 7200  # 2 h desde la última actividad

_SIN_TIPOS = {
    "compras": "El TXT no contiene facturas, notas de crédito, notas de débito ni retenciones",
    "ventas": "El TXT no contiene facturas, notas de crédito, notas de débito, liquidaciones de "
              "compra ni retenciones",
}
_SALIDAS = ("resultado.xlsx", "observaciones.xlsx", "tablero.json")
_LOCK_ESPERA_SEG = 10.0   # tiempo máximo esperando el bloqueo
_LOCK_HUERFANO_SEG = 30.0  # un .lock más viejo quedó de un proceso muerto: se descarta


class _Cancelado(Exception):
    """El recálculo fue superado por una generación más nueva."""


# ----------------------------------------------------------------------------- disco

def _ruta_sesion(sesion_id):
    return get_job_dir(sesion_id) / "sesion.json"


def _carpeta_txt(sesion_id):
    return get_job_dir(sesion_id) / "txt"


def _reemplazar(origen: Path, destino: Path):
    """os.replace con reintentos: en Windows falla mientras otro hilo tiene abierto el destino."""
    for intento in range(50):
        try:
            os.replace(origen, destino)
            return
        except PermissionError:
            if intento == 49:
                raise
            time.sleep(0.02)


def _escribir_atomico(destino: Path, datos: bytes):
    tmp = destino.with_name(f"{destino.name}.{os.getpid()}.{threading.get_ident()}.tmp")
    tmp.write_bytes(datos)
    _reemplazar(tmp, destino)


def _leer_sesion(sesion_id) -> dict:
    ruta = _ruta_sesion(sesion_id)
    if not ruta.exists():
        raise KeyError(f"La sesión {sesion_id} no existe o expiró")
    for intento in range(50):
        try:
            return json.loads(ruta.read_text(encoding="utf-8"))
        except PermissionError:  # Windows: reemplazo en curso
            if intento == 49:
                raise
            time.sleep(0.02)


def _guardar_sesion(sesion_id, ses: dict):
    ses["ultima_actividad"] = time.time()
    _escribir_atomico(_ruta_sesion(sesion_id), json.dumps(ses, ensure_ascii=False).encode("utf-8"))


@contextmanager
def _bloqueo(sesion_id):
    """Exclusión mutua por sesión entre hilos y procesos (archivo .lock con O_EXCL)."""
    ruta = get_job_dir(sesion_id) / ".lock"
    limite = time.monotonic() + _LOCK_ESPERA_SEG
    while True:
        try:
            os.close(os.open(ruta, os.O_CREAT | os.O_EXCL | os.O_WRONLY))
            break
        except (FileExistsError, PermissionError):
            try:
                if time.time() - ruta.stat().st_mtime > _LOCK_HUERFANO_SEG:
                    ruta.unlink(missing_ok=True)
                    continue
            except OSError:
                pass
            if time.monotonic() > limite:
                raise TimeoutError(f"La sesión {sesion_id} está ocupada; intente de nuevo")
            time.sleep(0.005)
    try:
        yield
    finally:
        try:
            ruta.unlink(missing_ok=True)
        except OSError:
            pass


def _actividad(sesion_id):
    """Reprograma la limpieza: el job se borra 2 h después de la última actividad."""
    schedule_cleanup(sesion_id, delay_seconds=LIMPIEZA_SEG)


# ----------------------------------------------------------------------------- TXT

def _nombre_txt(nombre: str) -> str:
    """Basename (sin rutas, con / o \\) y extensión normalizada a `.txt` minúscula."""
    base = Path((nombre or "").replace("\\", "/")).name
    stem = Path(base).stem if base.lower().endswith(".txt") else base
    if not stem.strip(" ."):
        raise ValueError("Nombre de archivo vacío o inválido")
    return stem + ".txt"


def _validar(modo, ruta: Path, nombre: str) -> dict:
    """Lee el TXT solo (con los lectores reales del pipeline) y devuelve su entrada de la lista."""
    tipos = dict.fromkeys(TIPOS[modo], 0)
    error = None
    try:
        if modo == "compras":
            sri = leer_sri_txts([str(ruta)])
            for t, n in sri["tipo"].value_counts().items():
                tipos[t] = int(n)
            tipos["RET"] = int(len(leer_sri_retenciones_txt([str(ruta)])))
            avisos = []
        else:
            docs, ret, avisos = leer_emitidos_txts([str(ruta)])
            for t, n in docs["tipo"].value_counts().items():
                tipos[t] = int(n)
            tipos["RET"] = int(len(ret))
        if not sum(tipos.values()):
            error = " ".join([_SIN_TIPOS[modo] + "."] + [str(a) for a in avisos])
    except Exception as exc:  # TXT ilegible: el mensaje no debe exponer la ruta del servidor
        texto = str(exc).replace(str(ruta), nombre) or type(exc).__name__
        error = f"No se pudo leer {nombre}: {texto}" if nombre not in texto else texto
    if error:
        return {"nombre": nombre, "estado": "error", "tipos": dict.fromkeys(TIPOS[modo], 0),
                "filas": 0, "error": error}
    return {"nombre": nombre, "estado": "ok", "tipos": tipos, "filas": int(sum(tipos.values())),
            "error": None}


def crear_sesion(modo: str) -> str:
    if modo not in LIMITES:
        raise ValueError(f"Modo inválido: {modo!r} (use 'compras' o 'ventas')")
    sesion_id = create_job()
    _carpeta_txt(sesion_id).mkdir(parents=True, exist_ok=True)
    _guardar_sesion(sesion_id, {"modo": modo, "generacion": 0, "archivos": []})
    write_status(sesion_id, "done", "Sin TXT cargados", 1.0)
    _actividad(sesion_id)
    return sesion_id


def agregar_txt(sesion_id: str, nombre: str, contenido: bytes) -> dict:
    """Guarda el TXT (mismo nombre = reemplazo) y lo valida; ValueError si se excede el límite."""
    nombre = _nombre_txt(nombre)
    ses = _leer_sesion(sesion_id)
    modo = ses["modo"]

    def excede(s):
        return (nombre not in {a["nombre"] for a in s["archivos"]}
                and len(s["archivos"]) >= LIMITES[modo])
    if excede(ses):
        raise ValueError(f"Se admiten como máximo {LIMITES[modo]} TXT en {modo}")
    carpeta = _carpeta_txt(sesion_id)
    carpeta.mkdir(exist_ok=True)
    # Se valida una copia temporal: el TXT anterior con el mismo nombre sigue intacto hasta publicar.
    tmp = carpeta / f".{nombre}.{os.getpid()}.{threading.get_ident()}.tmp"
    try:
        tmp.write_bytes(contenido)
        entrada = _validar(modo, tmp, nombre)
        with _bloqueo(sesion_id):
            ses = _leer_sesion(sesion_id)
            if excede(ses):
                raise ValueError(f"Se admiten como máximo {LIMITES[modo]} TXT en {modo}")
            _reemplazar(tmp, carpeta / nombre)
            archivos = ses["archivos"]
            pos = next((i for i, a in enumerate(archivos) if a["nombre"] == nombre), None)
            if pos is None:
                archivos.append(entrada)
            else:
                archivos[pos] = entrada
            _guardar_sesion(sesion_id, ses)
    finally:
        tmp.unlink(missing_ok=True)
    _actividad(sesion_id)
    return entrada


def quitar_txt(sesion_id: str, nombre: str) -> None:
    """Quita el TXT de la sesión (sin efecto si no estaba)."""
    nombre = _nombre_txt(nombre)
    with _bloqueo(sesion_id):
        ses = _leer_sesion(sesion_id)
        ses["archivos"] = [a for a in ses["archivos"] if a["nombre"] != nombre]
        (_carpeta_txt(sesion_id) / nombre).unlink(missing_ok=True)
        _guardar_sesion(sesion_id, ses)
    _actividad(sesion_id)


def listar_txt(sesion_id: str) -> list[dict]:
    return list(_leer_sesion(sesion_id)["archivos"])


# ----------------------------------------------------------------------------- recálculo

def _lanzar(sesion_id, generacion):
    threading.Thread(target=recalcular, args=(sesion_id, generacion), daemon=True).start()


def solicitar_recalculo(sesion_id: str) -> int:
    """Incrementa la generación (bajo el bloqueo de la sesión) y lanza el recálculo en un hilo."""
    with _bloqueo(sesion_id):
        ses = _leer_sesion(sesion_id)
        ses["generacion"] = int(ses.get("generacion", 0)) + 1
        _guardar_sesion(sesion_id, ses)
        generacion = ses["generacion"]
        write_status(sesion_id, "running", "Actualizando...", 0.0)
    _actividad(sesion_id)
    _lanzar(sesion_id, generacion)
    return generacion


def _vigente(sesion_id, generacion) -> bool:
    try:
        return int(_leer_sesion(sesion_id).get("generacion", 0)) == generacion
    except KeyError:  # la sesión se borró: nada que publicar
        return False


def _resumen_vacio(modo, advertencias):
    if modo == "compras":
        return {"por_tipo": {}, "advertencias": list(advertencias)}
    return {"por_tipo": {}, "periodo": None,
            "comprobantes_sri": dict.fromkeys(TIPOS["ventas"], 0),
            "advertencias": list(advertencias), "retenciones_fallidos": 0, "recalculado": False}


def recalcular(sesion_id: str, generacion: int) -> None:
    """Recalcula la sesión con sus TXT en estado ok. Si otra generación lo supera, termina sin
    escribir nada (ni status.json, ni tablero, ni Excel).

    Sin ningún TXT ok: tablero vacío y estado done sin consultar Odoo; resultado.xlsx se elimina
    (no hay Excel completo que ofrecer) y observaciones.xlsx dice que no hay observaciones."""
    job_dir = get_job_dir(sesion_id)
    tmp = {n: job_dir / f"_g{generacion}_{n}" for n in _SALIDAS}

    def cancelado():
        if not _vigente(sesion_id, generacion):
            raise _Cancelado

    def avisar(mensaje, progreso):
        with _bloqueo(sesion_id):
            cancelado()
            write_status(sesion_id, "running", mensaje, progreso)

    try:
        cancelado()
        ses = _leer_sesion(sesion_id)
        modo = ses["modo"]
        validos = [a["nombre"] for a in ses["archivos"] if a["estado"] == "ok"]
        excluidos = [f"El archivo {a['nombre']} se excluyó por no poder leerse: {a['error']}"
                     for a in ses["archivos"] if a["estado"] != "ok"]
        rutas = sorted(str(_carpeta_txt(sesion_id) / n) for n in validos)
        if rutas:
            procesar = procesar_compras if modo == "compras" else procesar_ventas
            resumen, tablero = procesar(rutas, str(tmp["resultado.xlsx"]), avisar,
                                        cancelado=cancelado, advertencias=excluidos)
        else:
            resumen = _resumen_vacio(modo, excluidos)
            tablero = {"modo": modo, "periodo": None, "tarjetas": [], "excepciones": [],
                       "advertencias": list(excluidos)}
        tablero["generacion"] = generacion
        cancelado()
        tmp["tablero.json"].write_text(json.dumps(tablero, ensure_ascii=False), encoding="utf-8")
        escribir_excel_observaciones(tablero, str(tmp["observaciones.xlsx"]))
        with _bloqueo(sesion_id):
            cancelado()
            for nombre, origen in tmp.items():
                if origen.exists():
                    _reemplazar(origen, job_dir / nombre)
                else:  # sin TXT válidos no hay Excel completo: no dejar el de otra generación
                    (job_dir / nombre).unlink(missing_ok=True)
            write_status(sesion_id, "done", "Completado", 1.0, resumen=resumen)
    except _Cancelado:
        pass
    except Exception as exc:
        try:
            with _bloqueo(sesion_id):
                if _vigente(sesion_id, generacion):
                    write_status(sesion_id, "error", f"{type(exc).__name__}: {exc}",
                                 error=traceback.format_exc())
        except Exception:
            pass  # la sesión se borró o está bloqueada: no hay dónde informar
    finally:
        for origen in tmp.values():
            try:
                origen.unlink(missing_ok=True)
            except OSError:
                pass
