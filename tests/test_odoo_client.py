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
    assert "secreto123" not in repr(e.value)
    assert e.value.__suppress_context__ is True


class _ProxyTimeout(FakeProxy):
    def authenticate(self, *a):
        raise TimeoutError("timed out contra http://u:secreto123@host")


def test_timeout_lanza_odooerror_generico():
    with pytest.raises(OdooError) as e:
        lector(_ProxyTimeout([])).fetch_compras(date(2026, 9, 1), date(2026, 9, 30))
    assert str(e.value) == "No se pudo consultar Odoo (error de conexión o de la consulta)"
    assert "secreto123" not in repr(e.value) and e.value.__suppress_context__ is True


def test_factory_por_defecto_usa_transport_con_timeout_60():
    import xmlrpc.client
    for url, cls in (("http://odoo/xmlrpc/2/common", xmlrpc.client.Transport),
                     ("https://odoo/xmlrpc/2/common", xmlrpc.client.SafeTransport)):
        proxy = odoo_client._proxy_con_timeout(url, allow_none=True)
        tr = proxy._ServerProxy__transport
        assert isinstance(tr, cls)
        assert tr.make_connection("odoo").timeout == odoo_client.TIMEOUT_SEGUNDOS == 60


# ---------------------------------------------------------------- retenciones recibidas

class FakeProxyRet:
    def __init__(self, rets, lineas, facturas=None, partners=None, falla_en=None):
        self.rets, self.lineas = rets, lineas
        self.facturas = facturas or []
        self.partners = partners or [{"id": 7, "vat": "1790000000001"}]
        self.falla_en = falla_en
        self.llamadas = []

    def authenticate(self, db, user, pwd, ctx):
        return 9

    def execute_kw(self, db, uid, pwd, model, method, args, kw=None):
        self.llamadas.append((model, method, args, kw))
        assert method in {"search_read", "read", "search_count", "fields_get"}
        if self.falla_en == model:
            raise OSError("fallo http://u:secreto123@host")
        kw = kw or {}
        data = {"account.move.retention": self.rets,
                "account.move.retention.line": self.lineas,
                "account.move": self.facturas}.get(model)
        if data is not None and method == "search_read":
            dom = args[0]
            for cond in dom:
                if cond[0] in ("id", "retention_id") and cond[1] == "in":
                    campo = "id" if cond[0] == "id" else "retention_id"
                    ids = set(cond[2])
                    data = [d for d in data if (d[campo][0] if isinstance(d[campo], list) else d[campo]) in ids]
            off = kw.get("offset", 0)
            return data[off:off + kw.get("limit", 500)]
        if model == "res.partner":
            return self.partners
        return []


def ret(**kw):
    base = {"id": 10, "retention_number": "001-200-000004340", "retention_date": "2026-09-12",
            "state": "posted", "partner_id": [7, "Cliente SA"], "invoice_id": [100, "FAC"],
            "type": "out_retention"}
    base.update(kw)
    return base


def lin(i, rid, pct=2.0, base=100.0, val=2.0, tax="312 - Transferencia de bienes (2%)"):
    return {"id": i, "retention_id": [rid, "x"], "amount_untaxed": base,
            "amount_retention": val, "tax_percent": pct, "tax_id": [50 + i, tax]}


def fac(i=100, name="FAC 006-001-000003385", x_number=False):
    return {"id": i, "name": name, "x_number": x_number}


def proxy_ret(**kw):
    rets = kw.pop("rets", [ret(), ret(id=11, retention_number="001-200-000004341",
                                      state="cancel", invoice_id=[101, "x"])])
    lineas = kw.pop("lineas", [lin(1, 10), lin(2, 10, pct=30.0, val=30.0, tax="725 - IVA 30% (30%)"),
                               lin(3, 11)])
    facturas = kw.pop("facturas", [fac(100, "006-001-000003385", False),
                                   fac(101, "FAC", "001-002-000000009")])
    return FakeProxyRet(rets, lineas, facturas, **kw)


def fetch_ret(proxy):
    return lector(proxy).fetch_retenciones(date(2026, 9, 1), date(2026, 9, 30))


def test_fetch_retenciones_una_fila_por_linea():
    df = fetch_ret(proxy_ret())
    assert len(df) == 3
    assert list(df.columns) == odoo_client.COLUMNAS_RET
    assert df["ruc"].iloc[0] == "1790000000001"
    assert df["fecha"].iloc[0] == pd.Timestamp("2026-09-12")


def test_fetch_retenciones_solo_recibidas():
    p = proxy_ret()
    fetch_ret(p)
    dom = next(a[0] for m, me, a, k in p.llamadas
               if m == "account.move.retention" and me == "search_read")
    assert ("type", "=", "out_retention") in dom
    assert ("retention_date", ">=", "2026-08-27") in dom
    assert ("retention_date", "<=", "2026-10-05") in dom
    assert "in_retention" not in repr(dom)


def test_fetch_retenciones_codigo_y_pct():
    df = fetch_ret(proxy_ret())
    assert df["cod"].iloc[0] == "312" and df["pct"].iloc[0] == 2.0
    assert df["cod"].iloc[1] == "725" and df["pct"].iloc[1] == 30.0


@pytest.mark.parametrize("name,xn", [("006-001-000003385", False),
                                     ("FAC 006-001-000003385", False),
                                     ("FAC", "006-001-000003385")])
def test_fetch_retenciones_factura_formato_sri(name, xn):
    p = proxy_ret(rets=[ret()], lineas=[lin(1, 10)], facturas=[fac(100, name, xn)])
    assert fetch_ret(p)["factura"].iloc[0] == "006-001-000003385"


def test_fetch_retenciones_estado_cancelada_se_conserva():
    df = fetch_ret(proxy_ret())
    assert df.loc[df["id_ret"] == 11, "estado"].tolist() == ["cancel"]
    assert df.loc[df["id_ret"] == 10, "estado"].tolist() == ["posted", "posted"]
    assert df.loc[df["id_ret"] == 11, "factura"].iloc[0] == "001-002-000000009"


def test_fetch_retenciones_error_generico():
    with pytest.raises(OdooError) as e:
        fetch_ret(proxy_ret(falla_en="account.move.retention.line"))
    assert "secreto123" not in str(e.value) and "secreto123" not in repr(e.value)
    assert e.value.__suppress_context__ is True


def test_fetch_retenciones_vacio():
    df = fetch_ret(FakeProxyRet([], []))
    assert df.empty and list(df.columns) == odoo_client.COLUMNAS_RET
