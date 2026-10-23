"""Cliente de SOLO LECTURA de Odoo 14 (XML-RPC) para las compras recibidas.

Solo usa authenticate, search_read y read. Las credenciales llegan por variables
de entorno (ODOO_URL, ODOO_DB, ODOO_USER, ODOO_PASSWORD) o por el `.env` local, y
nunca se incluyen en mensajes de error.
"""
import os
import re
import xmlrpc.client
from datetime import date, timedelta

import pandas as pd

from emitidos_sri import RUC_EMPRESA

RUTA_ENV = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
VARIABLES = ("ODOO_URL", "ODOO_DB", "ODOO_USER", "ODOO_PASSWORD")
PAGINA = 500
CAMPOS = ["name", "x_number", "partner_id", "invoice_date", "move_type", "state",
          "x_receipt_type_id", "x_authorization_number", "x_is_edocument",
          "amount_untaxed", "amount_tax", "amount_total", "reversed_entry_id"]
TIPOS = {"01": "FAC", "04": "NC", "05": "ND"}
COLUMNAS = ["id", "num", "nfac", "socio", "ruc", "fecha", "tipo", "cod_comp", "clave",
            "base", "iva", "total", "afectada", "elegible"]
COLUMNAS_RET = ["id_ret", "n_ret", "fecha", "estado", "socio", "ruc", "factura", "pct",
                "base", "valor", "cod"]
COLUMNAS_EMITIDOS = ["id", "num", "nfac", "socio", "ruc", "fecha", "tipo", "clave", "base",
                     "iva", "total", "afectada", "estado", "electronico", "elegible"]
# Verificado en solo lectura: los ids 3 (Liquidacion de compra) y 15 (Comprobante por
# reembolso) de account.receipt.type tienen code "03". El mapa codigo->tipo es TIPOS.
CODIGOS_LIQUIDACION = ("03",)
_NUM_DOC = re.compile(r"\d{3}-\d{3}-\d{9}")


def _num_doc(*valores):
    """Primer patron NNN-NNN-NNNNNNNNN hallado en los textos dados, o pd.NA."""
    for v in valores:
        t = _txt(v)
        if t is not pd.NA and (m := _NUM_DOC.search(t)):
            return m.group(0)
    return pd.NA


class OdooError(Exception):
    pass


TIMEOUT_SEGUNDOS = 60


class _TransportTimeout(xmlrpc.client.Transport):
    def make_connection(self, host):
        conn = super().make_connection(host)
        conn.timeout = TIMEOUT_SEGUNDOS
        return conn


class _SafeTransportTimeout(xmlrpc.client.SafeTransport):
    def make_connection(self, host):
        conn = super().make_connection(host)
        conn.timeout = TIMEOUT_SEGUNDOS
        return conn


def _proxy_con_timeout(url, **kw):
    """ServerProxy con timeout por conexión (sin tocar el timeout global de sockets)."""
    transporte = _SafeTransportTimeout() if url.lower().startswith("https") else _TransportTimeout()
    return xmlrpc.client.ServerProxy(url, transport=transporte, **kw)


def cargar_env(ruta_env: str | None = None) -> dict[str, str]:
    """os.environ, completado con las claves faltantes del archivo .env."""
    env = dict(os.environ)
    ruta = ruta_env or RUTA_ENV
    if os.path.isfile(ruta):
        with open(ruta, encoding="utf-8") as f:
            for linea in f:
                linea = linea.strip()
                if not linea or linea.startswith("#") or "=" not in linea:
                    continue
                k, v = linea.split("=", 1)
                k, v = k.strip(), v.strip().strip('"').strip("'")
                env.setdefault(k, v)
    return env


def _txt(v):
    """Texto de Odoo o pd.NA (Odoo devuelve False para vacío)."""
    return v.strip() if isinstance(v, str) and v.strip() else pd.NA


def _clave(v):
    t = _txt(v)
    if t is pd.NA or not t.isdigit() or not 45 <= len(t) <= 49:
        return pd.NA
    return t.zfill(49)


def _fila_base(m, vat):
    """Campos comunes de una fila (compras y emitidos); importes en valor absoluto."""
    pid = m["partner_id"][0] if m["partner_id"] else None
    return {
        "id": m["id"],
        "num": m["name"],
        "nfac": _txt(m["x_number"]),
        "socio": m["partner_id"][1] if m["partner_id"] else "",
        "ruc": _txt(vat.get(pid)),
        "fecha": pd.Timestamp(m["invoice_date"]) if m["invoice_date"] else pd.NaT,
        "base": abs(float(m["amount_untaxed"])),
        "iva": abs(float(m["amount_tax"])),
        "total": abs(float(m["amount_total"])),
    }


def _afectada(m, tipo, rev):
    """x_number del documento afectado (solo notas de credito)."""
    rev_id = m["reversed_entry_id"][0] if m["reversed_entry_id"] else None
    return _txt(rev.get(rev_id)) if tipo == "NC" and rev_id else pd.NA


class OdooReader:
    def __init__(self, url, db, user, password, proxy_factory=_proxy_con_timeout):
        self._url = url.rstrip("/")
        self._db = db
        self._user = user
        self._password = password
        self._factory = proxy_factory
        self._uid = None
        self._obj = None

    @classmethod
    def from_env(cls) -> "OdooReader":
        env = cargar_env()
        faltan = [k for k in VARIABLES if not env.get(k)]
        if faltan:
            raise OdooError("Falta configurar " + "/".join(faltan))
        return cls(env["ODOO_URL"], env["ODOO_DB"], env["ODOO_USER"], env["ODOO_PASSWORD"])

    def _conectar(self):
        if self._uid is None:
            common = self._factory(self._url + "/xmlrpc/2/common", allow_none=True)
            uid = common.authenticate(self._db, self._user, self._password, {})
            if not uid:
                raise OdooError("Odoo rechazó las credenciales configuradas")
            self._uid = uid
            self._obj = self._factory(self._url + "/xmlrpc/2/object", allow_none=True)

    def _call(self, model, method, args, **kw):
        return self._obj.execute_kw(self._db, self._uid, self._password, model, method,
                                    args, kw)

    def _search_read_todo(self, model, dominio, campos):
        filas, off = [], 0
        while True:
            lote = self._call(model, "search_read", [dominio], fields=campos,
                              limit=PAGINA, offset=off, order="id")
            filas += lote
            if len(lote) < PAGINA:
                return filas
            off += PAGINA

    def _mapa_codigos(self):
        """id de account.receipt.type -> code."""
        return {t["id"]: t["code"] for t in
                self._call("account.receipt.type", "search_read", [[]], fields=["code"])}

    def _mapa_vat(self, moves):
        """id de partner -> vat, para los socios de los asientos dados."""
        pids = sorted({m["partner_id"][0] for m in moves if m["partner_id"]})
        return ({p["id"]: p["vat"] for p in self._call("res.partner", "read", [pids], fields=["vat"])}
                if pids else {})

    def _mapa_afectadas(self, moves):
        """id de asiento revertido -> x_number, para los asientos dados."""
        rids = sorted({m["reversed_entry_id"][0] for m in moves if m["reversed_entry_id"]})
        return ({r["id"]: r["x_number"] for r in
                 self._call("account.move", "search_read", [[("id", "in", rids)]],
                            fields=["x_number"])} if rids else {})

    def fetch_compras(self, desde: date, hasta: date, margen_dias: int = 5) -> pd.DataFrame:
        try:
            return self._fetch_compras(desde, hasta, margen_dias)
        except OdooError:
            raise
        except Exception:
            # Mensaje genérico a propósito: el original podría contener URL con credenciales.
            raise OdooError("No se pudo consultar Odoo (error de conexión o de la consulta)") from None

    def _fetch_compras(self, desde, hasta, margen_dias):
        self._conectar()
        dominio = [("state", "=", "posted"),
                   ("move_type", "in", ["in_invoice", "in_refund"]),
                   ("invoice_date", ">=", (desde - timedelta(days=margen_dias)).isoformat()),
                   ("invoice_date", "<=", (hasta + timedelta(days=margen_dias)).isoformat())]
        moves = self._search_read_todo("account.move", dominio, CAMPOS)
        if not moves:
            return pd.DataFrame(columns=COLUMNAS)

        codigos = self._mapa_codigos()
        vat = self._mapa_vat(moves)
        rev = self._mapa_afectadas(moves)

        filas = []
        for m in moves:
            cod = codigos.get(m["x_receipt_type_id"][0] if m["x_receipt_type_id"] else None, "")
            tipo = TIPOS.get(cod, "OTRO")
            clave = _clave(m["x_authorization_number"])
            fila = _fila_base(m, vat)
            fila.update({
                "tipo": tipo,
                "cod_comp": cod,
                "clave": clave,
                "afectada": _afectada(m, tipo, rev),
                "elegible": cod in TIPOS and clave is not pd.NA and not bool(m["x_is_edocument"]),
            })
            filas.append(fila)
        return pd.DataFrame(filas, columns=COLUMNAS)

    def fetch_emitidos(self, desde: date, hasta: date, margen_dias: int = 5) -> pd.DataFrame:
        """Documentos EMITIDOS por la empresa: FAC/NC/ND (ventas) y LIQ (liquidaciones de compra)."""
        try:
            return self._fetch_emitidos(desde, hasta, margen_dias)
        except OdooError:
            raise
        except Exception:
            raise OdooError("No se pudo consultar Odoo (error de conexión o de la consulta)") from None

    def _fetch_emitidos(self, desde, hasta, margen_dias):
        self._conectar()
        dominio = [("state", "in", ["posted", "cancel", "draft"]),
                   ("move_type", "in", ["out_invoice", "out_refund", "in_invoice"]),
                   ("invoice_date", ">=", (desde - timedelta(days=margen_dias)).isoformat()),
                   ("invoice_date", "<=", (hasta + timedelta(days=margen_dias)).isoformat())]
        moves = self._search_read_todo("account.move", dominio, CAMPOS)
        if not moves:
            return pd.DataFrame(columns=COLUMNAS_EMITIDOS)

        codigos = self._mapa_codigos()

        def tipo_de(m):
            cod = codigos.get(m["x_receipt_type_id"][0] if m["x_receipt_type_id"] else None, "")
            mt = m["move_type"]
            if mt == "in_invoice":  # compras: solo las liquidaciones se emiten
                return "LIQ" if cod in CODIGOS_LIQUIDACION else None
            tipo = TIPOS.get(cod)
            # el codigo debe ser coherente con el tipo de asiento
            if (tipo == "NC") == (mt == "out_refund") and mt in ("out_invoice", "out_refund"):
                return tipo
            return None

        moves = [(m, t) for m in moves if (t := tipo_de(m))]
        if not moves:
            return pd.DataFrame(columns=COLUMNAS_EMITIDOS)

        vat = self._mapa_vat([m for m, _ in moves])
        rev = self._mapa_afectadas([m for m, t in moves if t == "NC"])

        filas = []
        for m, tipo in moves:
            fila = _fila_base(m, vat)
            fila.update({
                "tipo": tipo,
                "clave": _clave(m["x_authorization_number"]),
                "afectada": _afectada(m, tipo, rev),
                "estado": m["state"],
                "electronico": bool(m["x_is_edocument"]),
                "elegible": m["state"] in ("posted", "cancel"),
            })
            filas.append(fila)
        return pd.DataFrame(filas, columns=COLUMNAS_EMITIDOS)

    def fetch_retenciones(self, desde: date, hasta: date, margen_dias: int = 5) -> pd.DataFrame:
        """Retenciones RECIBIDAS (out_retention), todos los estados; una fila por linea."""
        try:
            return self._fetch_retenciones(desde, hasta, margen_dias)
        except OdooError:
            raise
        except Exception:
            raise OdooError("No se pudo consultar Odoo (error de conexión o de la consulta)") from None

    def fetch_retenciones_emitidas(self, desde: date, hasta: date, margen_dias: int = 5) -> pd.DataFrame:
        """Retenciones EMITIDAS (in_retention), todos los estados; una fila por linea.

        `ruc` es siempre RUC_EMPRESA (la empresa es la emisora) y `socio` el proveedor retenido."""
        try:
            return self._fetch_retenciones(desde, hasta, margen_dias, "in_retention", emitidas=True)
        except OdooError:
            raise
        except Exception:
            raise OdooError("No se pudo consultar Odoo (error de conexión o de la consulta)") from None

    def _fetch_retenciones(self, desde, hasta, margen_dias, tipo="out_retention", emitidas=False):
        columnas = COLUMNAS_RET + ["clave"] if emitidas else COLUMNAS_RET
        self._conectar()
        dominio = [("type", "=", tipo),
                   ("retention_date", ">=", (desde - timedelta(days=margen_dias)).isoformat()),
                   ("retention_date", "<=", (hasta + timedelta(days=margen_dias)).isoformat())]
        campos = ["retention_number", "retention_date", "state", "partner_id", "invoice_id", "type"]
        if emitidas:
            campos.append("auth_retention_number")
        rets = self._search_read_todo("account.move.retention", dominio, campos)
        if not rets:
            return pd.DataFrame(columns=columnas)

        ids = [r["id"] for r in rets]
        lineas = []
        for i in range(0, len(ids), PAGINA):
            lineas += self._search_read_todo(
                "account.move.retention.line", [("retention_id", "in", ids[i:i + PAGINA])],
                ["retention_id", "amount_untaxed", "amount_retention", "tax_percent", "tax_id"])
        pids = sorted({r["partner_id"][0] for r in rets if r["partner_id"]})
        vat = ({p["id"]: p["vat"] for p in self._call("res.partner", "read", [pids], fields=["vat"])}
               if pids else {})
        fids = sorted({r["invoice_id"][0] for r in rets if r["invoice_id"]})
        facturas = {}
        for i in range(0, len(fids), PAGINA):
            for f in self._call("account.move", "search_read", [[("id", "in", fids[i:i + PAGINA])]],
                                fields=["name", "x_number"]):
                facturas[f["id"]] = _num_doc(f.get("x_number"), f.get("name"))

        por_id = {r["id"]: r for r in rets}
        filas = []
        for ln in lineas:
            r = por_id.get(ln["retention_id"][0])
            if r is None:
                continue
            pid = r["partner_id"][0] if r["partner_id"] else None
            fid = r["invoice_id"][0] if r["invoice_id"] else None
            nombre_tax = ln["tax_id"][1] if ln["tax_id"] else ""
            cod = _txt(nombre_tax.split(" - ", 1)[0]) if " - " in nombre_tax else pd.NA
            fila = {
                "id_ret": r["id"],
                "n_ret": _txt(r["retention_number"]),
                "fecha": pd.Timestamp(r["retention_date"]) if r["retention_date"] else pd.NaT,
                "estado": r["state"],
                "socio": r["partner_id"][1] if r["partner_id"] else "",
                "ruc": RUC_EMPRESA if emitidas else _txt(vat.get(pid)),
                "factura": facturas.get(fid, pd.NA),
                "pct": float(ln["tax_percent"] or 0),
                "base": float(ln["amount_untaxed"] or 0),
                "valor": float(ln["amount_retention"] or 0),
                "cod": cod,
            }
            if emitidas:
                fila["clave"] = _clave(r.get("auth_retention_number"))
            filas.append(fila)
        return pd.DataFrame(filas, columns=columnas)
