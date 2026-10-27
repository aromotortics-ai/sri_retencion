"""Helpers compartidos entre tests (importar como `from _helpers import ...`)."""
import traceback

from app.job_manager import get_job_dir, write_status


def correr_pipeline(modulo, job_id, modo):
    """Ejecuta `procesar_compras`/`procesar_ventas` de `modulo` sobre los TXT de la carpeta del job
    (`sri/` o `emitidos/`) y deja status.json como lo hace la sesión (done/error). Reemplaza al
    antiguo `run_pipeline_*`; los espías de estado parchean `_helpers.write_status`."""
    job_dir = get_job_dir(job_id)
    carpeta = "sri" if modo == "compras" else "emitidos"
    procesar = getattr(modulo, f"procesar_{modo}")
    try:
        rutas = sorted(str(p) for p in (job_dir / carpeta).glob("*.txt"))
        resumen, _ = procesar(rutas, str(job_dir / "resultado.xlsx"),
                              lambda m, p: write_status(job_id, "running", m, p))
        write_status(job_id, "done", "Completado", 1.0, resumen=resumen)
    except Exception as exc:
        write_status(job_id, "error", f"{type(exc).__name__}: {exc}", error=traceback.format_exc())


def _autorizacion(clave, version, num_doc="001003000088766"):
    if version == "1.0.0":
        cuerpo = (
            "<impuestos><impuesto><codigo>1</codigo><codigoRetencion>312</codigoRetencion>"
            "<baseImponible>100.00</baseImponible><porcentajeRetener>1.75</porcentajeRetener>"
            f"<valorRetenido>1.75</valorRetenido><codDocSustento>01</codDocSustento>"
            f"<numDocSustento>{num_doc}</numDocSustento>"
            "<fechaEmisionDocSustento>01/09/2026</fechaEmisionDocSustento></impuesto></impuestos>"
        )
    else:
        cuerpo = (
            "<docsSustento><docSustento><codDocSustento>01</codDocSustento>"
            f"<numDocSustento>{num_doc}</numDocSustento>"
            "<fechaEmisionDocSustento>01/09/2026</fechaEmisionDocSustento>"
            "<retenciones><retencion><codigo>1</codigo><codigoRetencion>312</codigoRetencion>"
            "<baseImponible>100.00</baseImponible><porcentajeRetener>1.75</porcentajeRetener>"
            "<valorRetenido>1.75</valorRetenido></retencion></retenciones>"
            "</docSustento></docsSustento>"
        )
    comp = (
        f'<comprobanteRetencion version="{version}"><infoTributaria>'
        f"<razonSocial>PROVEEDOR SA</razonSocial><ruc>1790000000001</ruc>"
        f"<claveAcceso>{clave}</claveAcceso><estab>001</estab><ptoEmi>200</ptoEmi>"
        "<secuencial>000004340</secuencial></infoTributaria>"
        "<infoCompRetencion><fechaEmision>05/09/2026</fechaEmision>"
        "<identificacionSujetoRetenido>2390000831001</identificacionSujetoRetenido>"
        "<razonSocialSujetoRetenido>EMPRESA</razonSocialSujetoRetenido>"
        f"<periodoFiscal>09/2026</periodoFiscal></infoCompRetencion>{cuerpo}"
        "</comprobanteRetencion>"
    )
    return (
        "<autorizacion><estado>AUTORIZADO</estado><numeroAutorizacion>"
        f"{clave}</numeroAutorizacion><fechaAutorizacion>2026-09-05</fechaAutorizacion>"
        f"<comprobante><![CDATA[{comp}]]></comprobante></autorizacion>"
    )
