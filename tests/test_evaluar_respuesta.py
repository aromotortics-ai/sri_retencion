from descargar_comprobantes_sri import evaluar_respuesta

RECHAZADA = ("<autorizaciones><autorizacion><estado>RECHAZADA</estado><mensajes><mensaje>"
             "<mensaje>ERROR EN LA ESTRUCTURA DE LA CLAVE DE ACCESO</mensaje>"
             "<informacionAdicional>fecha fuera del rango</informacionAdicional>"
             "</mensaje></mensajes></autorizacion></autorizaciones>")
AUTORIZADO = ("<autorizaciones><autorizacion><estado>AUTORIZADO</estado>"
              "<comprobante>&lt;comprobanteRetencion/&gt;</comprobante></autorizacion></autorizaciones>")


def test_rechazada_no_se_guarda():
    xml, motivo = evaluar_respuesta(RECHAZADA)
    assert xml is None
    assert "RECHAZADA" in motivo and "fuera del rango" in motivo


def test_autorizado_se_guarda():
    xml, motivo = evaluar_respuesta(AUTORIZADO)
    assert motivo is None and xml.startswith("<autorizacion>")


def test_autorizado_sin_comprobante():
    xml, motivo = evaluar_respuesta("<autorizacion><estado>AUTORIZADO</estado></autorizacion>")
    assert xml is None and "sin <comprobante>" in motivo


def test_sin_autorizacion():
    assert evaluar_respuesta("<x/>")[0] is None
