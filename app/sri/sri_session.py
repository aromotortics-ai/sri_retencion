"""
Sesión requests con SSL completo para los web services del SRI.

Problema: el SRI redirige requests de cel.sri.gob.ec a la IP 181.113.227.222,
que sirve el certificado de 'srienlinea.sri.gob.ec'. El cert es válido (firmado
por DigiCert Global G2, incluido en certifi) pero la verificación estándar falla
porque la IP no aparece en las SANs — es una mala configuración de infraestructura
del SRI.

Solución: usar urllib3 con assert_hostname='srienlinea.sri.gob.ec' al conectar
a esa IP. Esto verifica COMPLETAMENTE el TLS:
  - Cadena de CA: certifi (DigiCert incluido) ✓
  - Hostname: verifica contra 'srienlinea.sri.gob.ec', que sí está en el cert ✓
  - No se deshabilita ninguna verificación ✓
"""
import certifi
import urllib3
from requests import Session
from requests.adapters import HTTPAdapter
from requests.models import Response
from requests.utils import get_encoding_from_headers

_SRI_IP        = "181.113.227.222"
_SRI_ALT_HOST  = "srienlinea.sri.gob.ec"   # hostname real del cert servido por esa IP


class SRIAdapter(HTTPAdapter):
    """
    Adapter que detecta la IP de redirección del SRI y hace la verificación TLS
    correcta usando assert_hostname en lugar de deshabilitar check_hostname.
    """

    def send(self, request, stream=False, timeout=None, verify=True, cert=None, proxies=None):
        from urllib.parse import urlparse
        if urlparse(request.url).hostname != _SRI_IP:
            # Conexión normal al dominio — comportamiento estándar con verify completo
            return super().send(request, stream=stream, timeout=timeout,
                                verify=verify, cert=cert, proxies=proxies)

        # Conexión a la IP del SRI: verificar TLS completamente pero contra el
        # hostname correcto del certificado (assert_hostname).
        pool = urllib3.HTTPSConnectionPool(
            host=_SRI_IP,
            port=443,
            ca_certs=certifi.where(),
            assert_hostname=_SRI_ALT_HOST,   # verifica que el cert es de este host
            server_hostname=_SRI_ALT_HOST,   # SNI: pide el cert correcto al servidor
            cert_reqs="CERT_REQUIRED",
        )
        urllib3_timeout = urllib3.Timeout(connect=timeout, read=timeout) if timeout else None
        raw = pool.urlopen(
            method=request.method,
            url=request.path_url or "/",
            body=request.body,
            headers=dict(request.headers),
            redirect=False,
            assert_same_host=False,
            preload_content=True,
            timeout=urllib3_timeout,
        )

        resp = Response()
        resp.status_code = raw.status
        resp.headers = dict(raw.headers)
        resp.encoding = get_encoding_from_headers(resp.headers)
        resp._content = raw.data
        resp.url = request.url
        resp.request = request
        return resp


def sri_session() -> Session:
    """Sesión requests con TLS completo para el SRI (dominio + IP de redirección)."""
    s = Session()
    s.mount("https://", SRIAdapter())
    s.headers.update({"Connection": "close"})
    return s
