from datetime import date

import pandas as pd
import pytest

from odoo_client import COLUMNAS_EMITIDOS, OdooError, OdooReader

CLAVE49 = "3009202601239000083100120010050001507450015074515"
TIPOS_REC = [{"id": 1, "code": "01"}, {"id": 4, "code": "04"}, {"id": 5, "code": "05"},
             {"id": 3, "code": "03"}, {"id": 15, "code": "03"}, {"id": 2, "code": "02"}]


class FakeProxy:
    """Simula /common y /object de Odoo sin red (solo lectura)."""

    def __init__(self, moves, partners=None, reversed_=None, falla=False):
        self.moves = moves
        self.partners = partners or [{"id": 7, "vat": "1790000000001"}]
        self.reversed = reversed_ or []
        self.falla = falla
        self.llamadas = []

    def authenticate(self, db, user, pwd, ctx):
        if self.falla:
            raise OSError("conexion rechazada a http://odoo.local con secreto123")
        return 9

    def execute_kw(self, db, uid, pwd, model, method, args, kw=None):
        self.llamadas.append((model, method, args, kw))
        assert method in {"search_read", "read", "search_count", "fields_get"}
        if model == "account.move" and method == "search_read":
            dom = args[0]
            if any(t[0] == "id" for t in dom):
                return self.reversed
            off = kw.get("offset", 0)
            return self.moves[off:off + kw.get("limit", 500)]
        if model == "account.receipt.type":
            return TIPOS_REC
        if model == "res.partner":
            return self.partners
        return []


def mov(**kw):
    base = {"id": 1, "name": "FAC/2026/09/0001", "x_number": "001-005-000150745",
            "partner_id": [7, "Cliente SA"], "invoice_date": "2026-09-10",
            "move_type": "out_invoice", "state": "posted", "x_receipt_type_id": [1, "Factura"],
            "x_authorization_number": CLAVE49, "x_is_edocument": True,
            "amount_untaxed": 100.0, "amount_tax": 15.0, "amount_total": 115.0,
            "reversed_entry_id": False}
    base.update(kw)
    return base


def fetch(moves, **kw):
    proxy = FakeProxy(moves, **kw)
    lector = OdooReader("http://odoo.local", "db", "u", "secreto123",
                        proxy_factory=lambda url, **k: proxy)
    return lector.fetch_emitidos(date(2026, 9, 1), date(2026, 9, 30)), proxy


def test_fetch_emitidos_tipos():
    moves = [
        mov(id=1),
        mov(id=2, move_type="out_refund", x_receipt_type_id=[4, "NC"],
            amount_total=-115.0, amount_untaxed=-100.0, amount_tax=-15.0,
            reversed_entry_id=[50, "FAC/2026/09/0009"]),
        mov(id=3, x_receipt_type_id=[5, "ND"]),
        mov(id=4, move_type="in_invoice", x_receipt_type_id=[3, "Liquidación"]),
        mov(id=5, move_type="in_invoice", x_receipt_type_id=[15, "Reembolso"]),
        mov(id=6, move_type="in_invoice", x_receipt_type_id=[1, "Factura"]),  # compra normal
    ]
    df, _ = fetch(moves, reversed_=[{"id": 50, "x_number": "001-002-000000123"}])
    assert list(df["id"]) == [1, 2, 3, 4, 5]
    assert list(df["tipo"]) == ["FAC", "NC", "ND", "LIQ", "LIQ"]
    assert list(df.columns) == COLUMNAS_EMITIDOS
    assert df.loc[1, "total"] == 115.0 and df.loc[1, "base"] == 100.0 and df.loc[1, "iva"] == 15.0
    assert df.loc[1, "afectada"] == "001-002-000000123"
    assert pd.isna(df.loc[0, "afectada"])
    assert df.loc[0, "clave"] == CLAVE49
    assert df.loc[0, "ruc"] == "1790000000001"


def test_fetch_emitidos_estados():
    moves = [mov(id=1, state="posted"), mov(id=2, state="cancel"), mov(id=3, state="draft")]
    df, proxy = fetch(moves)
    assert list(df["estado"]) == ["posted", "cancel", "draft"]
    assert list(df["elegible"]) == [True, True, False]
    dom = next(a[0] for m, meth, a, kw in proxy.llamadas
               if m == "account.move" and meth == "search_read")
    assert ("state", "in", ["posted", "cancel", "draft"]) in dom
    assert ("invoice_date", ">=", "2026-08-27") in dom
    assert ("invoice_date", "<=", "2026-10-05") in dom
    assert {meth for _, meth, _, _ in proxy.llamadas} <= {"search_read", "read", "search_count"}


def test_fetch_emitidos_incluye_no_electronicos():
    df, _ = fetch([mov(id=1, x_is_edocument=False, x_authorization_number=False)])
    assert len(df) == 1
    assert bool(df.loc[0, "electronico"]) is False
    assert pd.isna(df.loc[0, "clave"])


def test_fetch_emitidos_error_sin_credenciales():
    proxy = FakeProxy([], falla=True)
    lector = OdooReader("http://odoo.local", "db", "u", "secreto123",
                        proxy_factory=lambda url, **k: proxy)
    with pytest.raises(OdooError) as e:
        lector.fetch_emitidos(date(2026, 9, 1), date(2026, 9, 30))
    assert "secreto123" not in str(e.value) and "odoo.local" not in str(e.value)


def test_fetch_emitidos_vacio():
    df, _ = fetch([])
    assert df.empty
    assert list(df.columns) == COLUMNAS_EMITIDOS
