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

RUTA_ENV = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
VARIABLES = ("ODOO_URL", "ODOO_DB", "ODOO_USER", "ODOO_PASSWORD")
PAGINA = 500
CAMPOS = ["name", "x_number", "partner_id", "invoice_date", "move_type", "state",
          "x_receipt_type_id", "x_authorization_number", "x_is_edocument",
          "amount_untaxed", "amount_tax", "amount_total", "reversed_entry_id"]
TIPOS = {"01": "FAC", "04": "NC", "05": "ND"}
COLUMNAS = ["id", "num", "nfac", "socio", "ruc", "fecha", "tipo", "cod_comp", "clave",
            "base", "iva", "total", "afectada", "elegible"]


class OdooError(Exception):
    pass


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


class OdooReader:
    def __init__(self, url, db, user, password, proxy_factory=xmlrpc.client.ServerProxy):
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

        tipos = {t["id"]: t["code"] for t in
                 self._call("account.receipt.type", "search_read", [[]], fields=["code"])}
        pids = sorted({m["partner_id"][0] for m in moves if m["partner_id"]})
        vat = {p["id"]: p["vat"] for p in self._call("res.partner", "read", [pids], fields=["vat"])} if pids else {}
        rids = sorted({m["reversed_entry_id"][0] for m in moves if m["reversed_entry_id"]})
        rev = ({r["id"]: r["x_number"] for r in
                self._call("account.move", "search_read", [[("id", "in", rids)]],
                           fields=["x_number"])} if rids else {})

        filas = []
        for m in moves:
            cod = tipos.get(m["x_receipt_type_id"][0] if m["x_receipt_type_id"] else None, "")
            tipo = TIPOS.get(cod, "OTRO")
            clave = _clave(m["x_authorization_number"])
            propio = bool(m["x_is_edocument"])
            nfac = _txt(m["x_number"])
            pid = m["partner_id"][0] if m["partner_id"] else None
            rev_id = m["reversed_entry_id"][0] if m["reversed_entry_id"] else None
            filas.append({
                "id": m["id"],
                "num": m["name"],
                "nfac": nfac,
                "socio": m["partner_id"][1] if m["partner_id"] else "",
                "ruc": _txt(vat.get(pid)),
                "fecha": pd.Timestamp(m["invoice_date"]) if m["invoice_date"] else pd.NaT,
                "tipo": tipo,
                "cod_comp": cod,
                "clave": clave,
                "base": abs(float(m["amount_untaxed"])),
                "iva": abs(float(m["amount_tax"])),
                "total": abs(float(m["amount_total"])),
                "afectada": _txt(rev.get(rev_id)) if tipo == "NC" and rev_id else pd.NA,
                "elegible": cod in TIPOS and clave is not pd.NA and not propio,
            })
        return pd.DataFrame(filas, columns=COLUMNAS)
