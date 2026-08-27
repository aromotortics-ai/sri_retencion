"""
Runner del pipeline de conciliación SRI.
Se ejecuta en un hilo background lanzado por app.py.
"""
import re
import time
import threading
from pathlib import Path

from job_manager import write_status, schedule_cleanup, get_job_dir
from sri_session import sri_session

CACHE_XML_DIR = Path(__file__).parent / "cache" / "xml"

ENDPOINT = "https://cel.sri.gob.ec/comprobantes-electronicos-ws/AutorizacionComprobantesOffline"
SOAP_TEMPLATE = """<?xml version="1.0" encoding="UTF-8"?>
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ec="http://ec.gob.sri.ws.autorizacion">
  <soapenv:Header/>
  <soapenv:Body>
    <ec:autorizacionComprobante>
      <claveAccesoComprobante>{clave}</claveAccesoComprobante>
    </ec:autorizacionComprobante>
  </soapenv:Body>
</soapenv:Envelope>"""
HEADERS = {"Content-Type": "text/xml; charset=utf-8", "Connection": "close"}


def run_pipeline(job_id: str) -> None:
    """
    Orquesta los 4 pasos del pipeline para un job dado.
    Escribe progreso en status.json cada vez que avanza.
    Llamar en un hilo daemon.
    """
    job_dir = get_job_dir(job_id)
    try:
        # ── Paso 1: Leer TXT ─────────────────────────────────────────────────
        write_status(job_id, "running", "Leyendo TXT del SRI...", 0.02)
        from parse_recibidos_txt import parse_recibidos_txt
        df = parse_recibidos_txt(str(job_dir / "Recibidos.txt"))
        claves = df["CLAVE_ACCESO"].dropna().tolist()
        write_status(job_id, "running", f"TXT leído: {len(claves)} claves de acceso", 0.05)

        # ── Paso 2: Descargar XMLs (con caché) ───────────────────────────────
        CACHE_XML_DIR.mkdir(parents=True, exist_ok=True)
        pendientes = [c for c in claves if not (CACHE_XML_DIR / f"{c}.xml").exists()]
        ya_cache = len(claves) - len(pendientes)

        session = sri_session()
        ok = ya_cache
        fallidos = []

        for i, clave in enumerate(pendientes):
            body = SOAP_TEMPLATE.format(clave=clave).encode("utf-8")
            try:
                r = session.post(ENDPOINT, data=body, headers=HEADERS, timeout=20)
                r.raise_for_status()
                match = re.search(r"<autorizacion>.*?</autorizacion>", r.text, re.DOTALL)
                if match:
                    dest = CACHE_XML_DIR / f"{clave}.xml"
                    dest.write_text(
                        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + match.group(0),
                        encoding="utf-8",
                    )
                    ok += 1
                else:
                    fallidos.append(clave)
            except Exception:
                fallidos.append(clave)
            time.sleep(0.5)

            # Actualizar progreso cada 10 descargas
            if (i + 1) % 10 == 0 or i == len(pendientes) - 1:
                progreso = 0.05 + (ok / max(len(claves), 1)) * 0.70
                msg = f"Descargando XMLs del SRI ({ok}/{len(claves)}"
                if fallidos:
                    msg += f", {len(fallidos)} fallidos"
                if ya_cache:
                    msg += f", {ya_cache} desde caché"
                msg += ")"
                write_status(job_id, "running", msg, progreso)

        # ── Paso 3: Armar Excel ───────────────────────────────────────────────
        write_status(job_id, "running", "Armando Excel de conciliación...", 0.80)
        xml_glob = str(CACHE_XML_DIR / "[0-9]*.xml")
        output_path = str(job_dir / "resultado.xlsx")

        from build_workbook_final import build_workbook
        stats = build_workbook(
            txt_path=str(job_dir / "Recibidos.txt"),
            xml_glob=xml_glob,
            output_path=output_path,
        )

        # ── Paso 4: Conciliar con Odoo ────────────────────────────────────────
        write_status(job_id, "running", "Conciliando con Odoo...", 0.92)
        from conciliar_odoo import conciliar
        resumen = conciliar(
            excel_path=output_path,
            reten_odoo_path=str(job_dir / "reten_odoo.xlsx"),
        )
        resumen["n_listado"] = stats["n_listado"]
        resumen["n_xml"]     = stats["n_xml"]
        resumen["fallidos"]  = len(fallidos)

        # ── Completado ────────────────────────────────────────────────────────
        write_status(job_id, "done", "Completado", 1.0, resumen=resumen)
        schedule_cleanup(job_id, delay_seconds=7200)

    except Exception as exc:
        import traceback
        write_status(
            job_id,
            "error",
            f"{type(exc).__name__}: {exc}",
            error=traceback.format_exc(),
        )
