from datetime import date

import pandas as pd
import pytest

import odoo_client
from odoo_client import OdooError, OdooReader

CLAVE49 = "3009202601239000083100120010050001507450015074515"


class FakeProxy:
    """Simula /common y /object de Odoo sin red."""

    def __init__(self, moves, tipos=None, partners=None, reversed_=None, falla=False):
        self.moves = moves
        self.tipos = tipos or [{"id": 1, "code": "01"}, {"id": 4, "code": "04"},
                                {"id": 5, "code": "05"}, {"id": 2, "code": "02"}]
        self.partners = partners or [{"id": 7, "vat": "1790000000001"}]
        self.reversed = reversed_ or []
        self.falla = falla
        self.llamadas = []

    def authenticate(self, db, user, pwd, ctx):
        if self.falla:
            raise OSError("conexion rechazada a http://u:secreto123@host")
        return 9

    def execute_kw(self, db, uid, pwd, model, method, args, kw=None):
        self.llamadas.append((model, method, args, kw))
        assert method in {"search_read", "read", "search_count", "fields_get"}
        if model == "account.move" and method == "search_read":
            dom = args[0]
            if any(t[0] == "id" for t in dom):  # lectura en lote de asientos revertidos
                return self.reversed
            off = kw.get("offset", 0)
            return self.moves[off:off + kw.get("limit", 500)]
        if model == "account.receipt.type":
            return self.tipos
        if model == "res.partner":
            return self.partners
        return []


def mov(**kw):
    base = {"id": 1, "name": "CE/2026/09/0001", "x_number": "001-005-000150745",
            "partner_id": [7, "Proveedor SA"], "invoice_date": "2026-09-10",
            "move_type": "in_invoice", "state": "posted", "x_receipt_type_id": [1, "Factura"],
            "x_authorization_number": CLAVE49, "x_is_edocument": False,
            "amount_untaxed": 100.0, "amount_tax": 15.0, "amount_total": 115.0,
            "reversed_entry_id": False}
    base.update(kw)
    return base


def lector(proxy):
    return OdooReader("http://odoo", "db", "u", "secreto123", proxy_factory=lambda url, **k: proxy)


def fetch(moves, **kw):
    proxy = FakeProxy(moves, **kw)
    return lector(proxy).fetch_compras(date(2026, 9, 1), date(2026, 9, 30)), proxy


def test_clave_48_digitos_se_rellena_a_49():
    df, _ = fetch([mov(x_authorization_number=CLAVE49[1:])])
    assert df.loc[0, "clave"] == "0" + CLAVE49[1:]
    assert len(df.loc[0, "clave"]) == 49
    assert bool(df.loc[0, "elegible"]) is True


def test_tipo_por_codigo_comprobante():
    moves = [mov(id=i, x_receipt_type_id=[t, "x"]) for i, t in [(1, 1), (2, 4), (3, 5), (4, 2)]]
    df, _ = fetch(moves)
    assert list(df["tipo"]) == ["FAC", "NC", "ND", "OTRO"]
    assert list(df["cod_comp"]) == ["01", "04", "05", "02"]


def test_elegible_exige_clave_y_no_ser_documento_propio():
    moves = [mov(id=1), mov(id=2, x_is_edocument=True),
             mov(id=3, x_authorization_number=False), mov(id=4, x_authorization_number="123"),
             mov(id=5, x_receipt_type_id=[2, "x"])]
    df, _ = fetch(moves)
    assert list(df["elegible"]) == [True, False, False, False, False]
    assert pd.isna(df.loc[2, "clave"])


def test_total_es_valor_absoluto():
    df, _ = fetch([mov(amount_total=-115.0, amount_untaxed=-100.0, amount_tax=-15.0)])
    assert df.loc[0, "total"] == 115.0
    assert df.loc[0, "base"] == 100.0 and df.loc[0, "iva"] == 15.0


def test_ruc_viene_del_partner():
    df, _ = fetch([mov(), mov(id=2, partner_id=[8, "Sin RUC"])],
                  partners=[{"id": 7, "vat": "1790000000001"}, {"id": 8, "vat": False}])
    assert df.loc[0, "ruc"] == "1790000000001"
    assert pd.isna(df.loc[1, "ruc"])
    assert df.loc[0, "socio"] == "Proveedor SA"


def test_afectada_de_reversed_entry_para_nc():
    nc = mov(id=2, move_type="in_refund", x_receipt_type_id=[4, "NC"],
             reversed_entry_id=[50, "CE/2026/09/0002"])
    df, _ = fetch([mov(), nc], reversed_=[{"id": 50, "x_number": "001-002-000000123"}])
    assert pd.isna(df.loc[0, "afectada"])
    assert df.loc[1, "afectada"] == "001-002-000000123"
    assert df.loc[1, "nfac"] == "001-005-000150745"
    assert df.loc[1, "fecha"] == pd.Timestamp("2026-09-10")


def test_consulta_usa_solo_posted_y_margen_de_fechas():
    _, proxy = fetch([mov()])
    dom = next(a[0] for m, meth, a, kw in proxy.llamadas
               if m == "account.move" and meth == "search_read")
    assert ("state", "=", "posted") in dom
    assert ("invoice_date", ">=", "2026-08-27") in dom
    assert ("invoice_date", "<=", "2026-10-05") in dom
    assert ("move_type", "in", ["in_invoice", "in_refund"]) in dom
    assert {meth for _, meth, _, _ in proxy.llamadas} <= {"search_read", "read", "search_count"}


def test_pagina_en_bloques_de_500():
    moves = [mov(id=i) for i in range(1, 1201)]
    df, proxy = fetch(moves)
    assert len(df) == 1200
    pags = [c for c in proxy.llamadas if c[0] == "account.move" and c[1] == "search_read"]
    assert len(pags) == 3


def test_from_env_lista_nombres_faltantes_sin_valores(monkeypatch, tmp_path):
    for k in ("ODOO_URL", "ODOO_DB", "ODOO_USER", "ODOO_PASSWORD"):
        monkeypatch.delenv(k, raising=False)
    monkeypatch.setattr(odoo_client, "RUTA_ENV", str(tmp_path / ".env"))
    (tmp_path / ".env").write_text('# c\nODOO_URL="http://x"\nODOO_PASSWORD=clave-secreta\n',
                                   encoding="utf-8")
    with pytest.raises(OdooError) as e:
        OdooReader.from_env()
    msg = str(e.value)
    assert "ODOO_DB" in msg and "ODOO_USER" in msg
    assert "ODOO_URL" not in msg and "ODOO_PASSWORD" not in msg
    assert "clave-secreta" not in msg and "http://x" not in msg


def test_cargar_env_prioriza_os_environ(monkeypatch, tmp_path):
    f = tmp_path / ".env"
    f.write_text("A=1\nB='2'\n", encoding="utf-8")
    monkeypatch.setenv("A", "9")
    env = odoo_client.cargar_env(str(f))
    assert env["A"] == "9" and env["B"] == "2"


def test_fallo_de_conexion_lanza_odooerror_sin_credenciales():
    proxy = FakeProxy([], falla=True)
    with pytest.raises(OdooError) as e:
        lector(proxy).fetch_compras(date(2026, 9, 1), date(2026, 9, 30))
    assert "secreto123" not in str(e.value)
    assert "secreto123" not in repr(e.value.__cause__ or "")
