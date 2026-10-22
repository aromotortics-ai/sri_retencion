# -*- coding: utf-8 -*-
"""
Conciliación de FACTURAS EMITIDAS (ventas) por la empresa: SRI vs Odoo.

Entradas (carpeta facturas_emitidas/):
  - emitidos_AAAA-MM-DD.txt: TXT diarios del portal SRI "Comprobantes electrónicos
    emitidos" -> Autorizados -> Factura (ISO-8859-1, tab). Trae base, IVA e importe.
  - fact_emi_2025_2026.xlsx: facturas de cliente exportadas de Odoo (13 columnas:
    Número, Cliente, Sucursal, Fecha, Fecha vencimiento, Estado, Estado electrónico,
    Subtotal IVA, Subtotal IVA0, Importe sin impuestos, Impuesto, Total, N° autorización).

Uso:
    python conciliar_facturas_emitidas.py facturas_emitidas salida.xlsx

Clave de cruce: clave de acceso (N° de autorización de Odoo = CLAVE_ACCESO del SRI).
Las facturas de Odoo sin autorización (estado electrónico 'returned') se reportan como
Solo en Odoo. Requiere xlsxwriter (pip install xlsxwriter): escribe en modo de memoria
constante, necesario para ~100 mil filas, y guarda el valor calculado de cada fórmula.
"""
import glob
import os
import sys

import pandas as pd
import xlsxwriter

FONT = "Arial"
TOL = 0.02
RES_RET = "Autorizada en SRI / returned en Odoo"
COLORES = {"Resumen_Mensual": "#1F4E79", "Resumen_Sucursal": "#1F4E79", "Conciliacion": "#1F4E79",
           "Solo_en_SRI": "#ED7D31", "Solo_en_Odoo": "#C00000", "Diferencia_Valor": "#FFC000",
           "Returned_en_Odoo": "#FFC000", "Fecha_o_Serie_distinta": "#7030A0"}
COLS_ODOO = ["num", "socio", "suc", "fecha", "venc", "estado", "eest", "iva", "iva0", "sinimp",
             "imp", "total", "aut"]


def leer_sri(carpeta):
    archivos = sorted(glob.glob(os.path.join(carpeta, "emitidos_*.txt")))
    if not archivos:
        raise ValueError("No se encontraron archivos emitidos_*.txt del SRI.")
    s = pd.concat([pd.read_csv(f, sep="\t", encoding="iso-8859-1", dtype=str) for f in archivos])
    s = s.drop_duplicates("CLAVE_ACCESO").copy()
    s["fe"] = pd.to_datetime(s["FECHA_EMISION"].str[:10], format="%d/%m/%Y")
    for col, n in (("VALOR_SIN_IMPUESTOS", "base"), ("IVA", "iva"), ("IMPORTE_TOTAL", "total")):
        s[n] = pd.to_numeric(s[col], errors="coerce")
    return s.rename(columns={"CLAVE_ACCESO": "clave", "SERIE_COMPROBANTE": "serie"}).reset_index(drop=True)


def leer_odoo(path):
    o = pd.read_excel(path, dtype={12: str})
    if len(o.columns) < len(COLS_ODOO):
        raise ValueError(f"El Excel de Odoo tiene {len(o.columns)} columnas; se esperan {len(COLS_ODOO)}.")
    o = o.iloc[:, :len(COLS_ODOO)].copy()
    o.columns = COLS_ODOO
    o["fecha"] = pd.to_datetime(o["fecha"])
    o["clave"] = o["aut"].astype("string").str.strip()
    return o


def conciliar(s, o):
    c = s[["clave", "serie", "fe", "base", "total"]].merge(
        o[["clave", "num", "socio", "suc", "fecha", "eest", "sinimp", "total"]]
        .rename(columns={"total": "total_odoo", "sinimp": "base_odoo"}),
        on="clave", how="outer", indicator=True)
    # facturas de Odoo sin autorización (clave nula) no se unen entre sí: se reponen
    c = c[~((c["_merge"] == "right_only") & c["clave"].isna())]
    sin_aut = o[o["clave"].isna()].assign(_merge="right_only")
    sin_aut = sin_aut.rename(columns={"total": "total_odoo", "sinimp": "base_odoo"})
    c = pd.concat([c, sin_aut[["clave", "num", "socio", "suc", "fecha", "eest", "base_odoo", "total_odoo", "_merge"]]],
                  ignore_index=True)
    # Autorizadas en el SRI que Odoo dejó sin clave ('returned'): se cruzan por N° de serie
    solo_sri = c[(c["_merge"] == "left_only")]
    solo_odoo = c[(c["_merge"] == "right_only") & c["clave"].isna()]
    pares = solo_sri.reset_index().merge(solo_odoo.reset_index(), left_on="serie", right_on="num",
                                         suffixes=("", "_o"))
    for r in pares.itertuples():
        for col in ("num", "socio", "suc", "fecha", "eest", "base_odoo", "total_odoo"):
            c.at[r.index, col] = getattr(r, col + "_o")
        c.at[r.index, "_merge"] = "both"
    c = c.drop(index=pares["index_o"]).reset_index(drop=True)
    c["en_sri"] = c["_merge"].isin(["both", "left_only"])
    c["en_odoo"] = c["_merge"].isin(["both", "right_only"])
    c["fecha_ref"] = c["fe"].fillna(c["fecha"])
    c["mes"] = c["fecha_ref"].dt.strftime("%Y-%m")
    c["serie_ref"] = c["serie"].fillna(c["num"])
    c = c.sort_values(["fecha_ref", "serie_ref"], kind="stable").reset_index(drop=True)
    return c


def resultado(r):
    if r.en_sri and r.en_odoo:
        if r.eest == "returned":
            return RES_RET
        if pd.isna(r.total_odoo):
            return "Coincide" if r.total == 0 else "Odoo sin importe"
        return "Diferencia de valor" if abs(r.total - r.total_odoo) > TOL else "Coincide"
    return "Solo en SRI" if r.en_sri else "Solo en Odoo"


def construir_excel(c, salida, desde, hasta):
    wb = xlsxwriter.Workbook(salida, {"constant_memory": True, "default_date_format": "dd/mm/yyyy"})
    f = lambda **k: wb.add_format({"font_name": FONT, "font_size": 10, **k})
    fn, fd, fm, fp = f(), f(num_format="dd/mm/yyyy"), f(num_format="#,##0.00"), f(num_format="0.0%")
    ftxt = f(italic=True, font_size=9)
    ftit = f(bold=True, font_size=14, font_color="#1F4E79")
    ftot = f(bold=True, bg_color="#DDEBF7")
    ftotm = f(bold=True, bg_color="#DDEBF7", num_format="#,##0.00")
    ftotp = f(bold=True, bg_color="#DDEBF7", num_format="0.0%")
    fb = f(bold=True)

    def enc(color):
        return f(bold=True, font_color="#FFFFFF", bg_color=color, align="center", valign="vcenter", text_wrap=True)

    n = len(c)
    c = c.copy()
    c["res"] = [resultado(r) for r in c.itertuples()]
    c["dif"] = (c["total"] - c["total_odoo"]).round(2)
    c["dif_dias"] = (c["fe"] - c["fecha"]).dt.days
    c["serie_dif"] = c["en_sri"] & c["en_odoo"] & (c["serie"] != c["num"])

    # El orden de creación de hojas define el orden de pestañas; en modo de memoria
    # constante cada hoja se escribe completa antes de pasar a la siguiente.
    ws_r = wb.add_worksheet("Resumen_Mensual")
    ws_s = wb.add_worksheet("Resumen_Sucursal")
    ws_c = wb.add_worksheet("Conciliacion")
    # ---- Conciliacion (datos + fórmulas con valor en caché) ----
    ws_c.set_tab_color(COLORES["Conciliacion"])
    cab = ["Clave_Acceso", "Serie_SRI", "Numero_Odoo", "Mes", "Fecha_SRI", "Fecha_Odoo", "Cliente", "Sucursal",
           "Estado_Electronico_Odoo", "En_SRI", "En_Odoo", "Base_SRI", "Total_SRI", "Total_Odoo", "Diferencia",
           "Resultado"]
    h = enc(COLORES["Conciliacion"])
    for j, t in enumerate(cab):
        ws_c.write(0, j, t, h)
    for i, r in enumerate(c.itertuples(), start=1):
        x = i + 1
        ws_c.write_string(i, 0, "" if pd.isna(r.clave) else r.clave, fn)
        ws_c.write_string(i, 1, "" if pd.isna(r.serie) else r.serie, fn)
        ws_c.write_string(i, 2, "" if pd.isna(r.num) else r.num, fn)
        ws_c.write_string(i, 3, r.mes, fn)
        ws_c.write_datetime(i, 4, r.fe.to_pydatetime(), fd) if pd.notna(r.fe) else ws_c.write_blank(i, 4, None, fn)
        ws_c.write_datetime(i, 5, r.fecha.to_pydatetime(), fd) if pd.notna(r.fecha) else ws_c.write_blank(i, 5, None, fn)
        ws_c.write_string(i, 6, "" if pd.isna(r.socio) else r.socio, fn)
        ws_c.write_string(i, 7, "" if pd.isna(r.suc) else r.suc, fn)
        ws_c.write_string(i, 8, "" if pd.isna(r.eest) else r.eest, fn)
        ws_c.write_string(i, 9, "Sí" if r.en_sri else "No", fn)
        ws_c.write_string(i, 10, "Sí" if r.en_odoo else "No", fn)
        ws_c.write_number(i, 11, r.base, fm) if pd.notna(r.base) else ws_c.write_blank(i, 11, None, fm)
        ws_c.write_number(i, 12, r.total, fm) if pd.notna(r.total) else ws_c.write_blank(i, 12, None, fm)
        ws_c.write_number(i, 13, r.total_odoo, fm) if pd.notna(r.total_odoo) else ws_c.write_blank(i, 13, None, fm)
        ws_c.write_formula(i, 14, f'=IF(AND(M{x}<>"",N{x}<>""),ROUND(M{x}-N{x},2),"")', fm,
                           "" if pd.isna(r.dif) else float(r.dif))
        ws_c.write_formula(
            i, 15,
            f'=IF(AND(J{x}="Sí",K{x}="Sí"),IF(I{x}="returned","{RES_RET}",IF(N{x}="",IF(M{x}=0,"Coincide","Odoo sin importe"),IF(ABS(O{x})>{TOL},"Diferencia de valor","Coincide"))),'
            f'IF(J{x}="Sí","Solo en SRI","Solo en Odoo"))', fn, r.res)
    ws_c.freeze_panes(1, 1)
    ws_c.autofilter(0, 0, n, len(cab) - 1)
    for j, w in enumerate([52, 19, 19, 9, 12, 12, 42, 13, 14, 8, 8, 12, 12, 12, 11, 20]):
        ws_c.set_column(j, j, w)

    # ---- Resumen mensual (fórmulas sobre Conciliacion) ----
    last = n + 1
    rng = lambda col: f"Conciliacion!${col}$2:${col}${last}"
    ws_r.set_tab_color(COLORES["Resumen_Mensual"])
    ws_r.write(0, 0, "Conciliación de facturas EMITIDAS (ventas): SRI vs Odoo — resumen mensual", ftit)
    ws_r.write(1, 0, f"RUC 2390000831001  |  Período {desde} a {hasta}  |  Clave de cruce: clave de acceso  |  "
                     "Mes = mes de emisión según SRI (si no está en SRI, mes de Odoo)", ftxt)
    e = ["Mes", "Facturas SRI", "Facturas Odoo", "Coinciden", "Diferencia de valor", RES_RET,
         "Solo en SRI", "Solo en Odoo", "% Coincidencia", "Total SRI", "Total Odoo", "Dif. total"]
    for j, t in enumerate(e):
        ws_r.write(3, j, t, enc(COLORES["Resumen_Mensual"]))
    ws_r.set_row(3, 30)
    g = c.groupby("mes")
    meses = sorted(c["mes"].unique())
    for k, m in enumerate(meses):
        r, x = 4 + k, 5 + k
        d = c[c["mes"] == m]
        cnt = lambda v: int((d["res"] == v).sum())
        vals = [int(d["en_sri"].sum()), int(d["en_odoo"].sum()), cnt("Coincide"), cnt("Diferencia de valor"),
                cnt(RES_RET), cnt("Solo en SRI"), cnt("Solo en Odoo")]
        ws_r.write(r, 0, m, fn)
        ws_r.write_formula(r, 1, f'=COUNTIFS({rng("D")},$A{x},{rng("J")},"Sí")', fn, vals[0])
        ws_r.write_formula(r, 2, f'=COUNTIFS({rng("D")},$A{x},{rng("K")},"Sí")', fn, vals[1])
        for j, (txt, v) in enumerate(zip(["Coincide", "Diferencia de valor", RES_RET, "Solo en SRI",
                                          "Solo en Odoo"], vals[2:])):
            ws_r.write_formula(r, 3 + j, f'=COUNTIFS({rng("D")},$A{x},{rng("P")},"{txt}")', fn, v)
        ws_r.write_formula(r, 8, f"=IF(B{x}=0,0,D{x}/B{x})", fp, vals[2] / vals[0] if vals[0] else 0)
        ts, to = float(d["total"].sum()), float(d["total_odoo"].sum())
        ws_r.write_formula(r, 9, f'=SUMIFS({rng("M")},{rng("D")},$A{x})', fm, ts)
        ws_r.write_formula(r, 10, f'=SUMIFS({rng("N")},{rng("D")},$A{x})', fm, to)
        ws_r.write_formula(r, 11, f"=J{x}-K{x}", fm, ts - to)
    ult, tr = 4 + len(meses), 5 + len(meses)
    ws_r.write(ult, 0, "Total", ftot)
    for j in range(1, 12):
        col = chr(65 + j)
        if j == 8:
            ws_r.write_formula(ult, j, f"=IF(B{tr}=0,0,D{tr}/B{tr})", ftotp,
                               int((c["res"] == "Coincide").sum()) / max(int(c["en_sri"].sum()), 1))
            continue
        v = {1: int(c["en_sri"].sum()), 2: int(c["en_odoo"].sum()), 3: int((c["res"] == "Coincide").sum()),
             4: int((c["res"] == "Diferencia de valor").sum()), 5: int((c["res"] == RES_RET).sum()),
             6: int((c["res"] == "Solo en SRI").sum()), 7: int((c["res"] == "Solo en Odoo").sum()),
             9: float(c["total"].sum()), 10: float(c["total_odoo"].sum()),
             11: float(c["total"].sum() - c["total_odoo"].sum())}[j]
        ws_r.write_formula(ult, j, f"=SUM({col}5:{col}{ult})", ftotm if j >= 9 else ftot, v)
    o0 = ult + 3
    ws_r.write(o0, 0, "Observaciones", fb)
    nf = int((c["en_sri"] & c["en_odoo"] & (c["dif_dias"].abs() > 0)).sum())
    ns = int(c["serie_dif"].sum())
    notas = [
        "Coincide = misma clave de acceso en SRI y Odoo con importe total igual (tolerancia $0.02).",
        "Solo en SRI = factura autorizada en el SRI sin registro en Odoo. Solo en Odoo = factura en Odoo sin "
        "autorización ni equivalente en el SRI, o con clave que no figura en el SRI.",
        "Autorizada en SRI / returned en Odoo = el SRI la autorizó pero Odoo la dejó en estado 'returned' sin "
        "clave de acceso (se cruza por N° de serie): hay que reprocesarla en Odoo. Facturas de $0 se cuentan como Coincide.",
        f"Coinciden por clave pero con fecha de emisión distinta entre SRI y Odoo: {nf}; con N° de serie distinto: {ns} "
        "(ver hoja Fecha_o_Serie_distinta).",
        "El Mes de la hoja Conciliacion y de este resumen es el mes de emisión SRI; si no está en SRI, la fecha de Odoo.",
    ]
    for k, t in enumerate(notas, 1):
        ws_r.write(o0 + k, 0, t, ftxt)
    ws_r.freeze_panes(4, 1)
    for j, w in enumerate([10, 13, 13, 12, 14, 12, 12, 12, 13, 16, 16, 14]):
        ws_r.set_column(j, j, w)

    # ---- Resumen por sucursal (Odoo) ----
    ws_s.set_tab_color(COLORES["Resumen_Sucursal"])
    ws_s.write(0, 0, "Resumen por sucursal (según Odoo; las 'Solo en SRI' no tienen sucursal)", ftit)
    es = ["Sucursal", "Facturas Odoo", "Coinciden", "Diferencia de valor", RES_RET, "Solo en Odoo",
          "Total Odoo"]
    for j, t in enumerate(es):
        ws_s.write(2, j, t, enc(COLORES["Resumen_Sucursal"]))
    ws_s.set_row(2, 30)
    sucs = sorted(c["suc"].dropna().unique())
    for k, sname in enumerate(sucs):
        r, x = 3 + k, 4 + k
        d = c[c["suc"] == sname]
        ws_s.write(r, 0, sname, fn)
        ws_s.write_formula(r, 1, f'=COUNTIFS({rng("H")},$A{x},{rng("K")},"Sí")', fn, len(d))
        for j, txt in enumerate(["Coincide", "Diferencia de valor", RES_RET, "Solo en Odoo"]):
            ws_s.write_formula(r, 2 + j, f'=COUNTIFS({rng("H")},$A{x},{rng("P")},"{txt}")', fn, int((d["res"] == txt).sum()))
        ws_s.write_formula(r, 6, f'=SUMIFS({rng("N")},{rng("H")},$A{x})', fm, float(d["total_odoo"].sum()))
    ult2, t2 = 3 + len(sucs), 4 + len(sucs)
    ws_s.write(ult2, 0, "Total", ftot)
    con_suc = c[c["suc"].notna()]
    tot_v = [len(con_suc)] + [int((con_suc["res"] == t).sum()) for t in
                              ("Coincide", "Diferencia de valor", RES_RET, "Solo en Odoo")] + \
            [float(con_suc["total_odoo"].sum())]
    for j in range(1, 7):
        ws_s.write_formula(ult2, j, f"=SUM({chr(65 + j)}4:{chr(65 + j)}{ult2})", ftotm if j == 6 else ftot,
                           tot_v[j - 1])
    for j, w in enumerate([16, 14, 12, 14, 14, 13, 16]):
        ws_s.set_column(j, j, w)

    # ---- Listas de excepciones ----
    def lista(nombre, cab, filas, anchos, fmts=None):
        ws = wb.add_worksheet(nombre)
        ws.set_tab_color(COLORES[nombre])
        for j, t in enumerate(cab):
            ws.write(0, j, t, enc(COLORES[nombre]))
        for i, fila in enumerate(filas, start=1):
            for j, v in enumerate(fila):
                if isinstance(v, pd.Timestamp):
                    ws.write_datetime(i, j, v.to_pydatetime(), fd)
                elif isinstance(v, (int, float)) and not pd.isna(v):
                    ws.write_number(i, j, v, fm if isinstance(v, float) else fn)
                else:
                    ws.write(i, j, "" if pd.isna(v) else v, fn)
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, max(len(filas), 1), len(cab) - 1)
        for j, w in enumerate(anchos):
            ws.set_column(j, j, w)

    d = c[c["res"] == "Solo en SRI"]
    lista("Solo_en_SRI", ["Fecha_Emision", "Serie", "Base", "Total", "Clave_Acceso"],
          [[r.fe, r.serie, r.base, r.total, r.clave] for r in d.itertuples()], [13, 19, 12, 12, 52])
    d = c[c["res"] == "Solo en Odoo"]
    lista("Solo_en_Odoo", ["Numero_Odoo", "Fecha", "Cliente", "Sucursal", "Estado_Electronico", "Total", "Clave_Odoo"],
          [[r.num, r.fecha, r.socio, r.suc, r.eest, r.total_odoo, r.clave] for r in d.itertuples()],
          [19, 12, 42, 13, 16, 12, 52])
    d = c[c["res"] == "Diferencia de valor"]
    lista("Diferencia_Valor", ["Fecha_SRI", "Serie", "Cliente", "Sucursal", "Total_SRI", "Total_Odoo", "Diferencia"],
          [[r.fe, r.serie, r.socio, r.suc, r.total, r.total_odoo, r.dif] for r in d.itertuples()],
          [12, 19, 42, 13, 12, 12, 11])
    d = c[c["res"] == RES_RET]
    lista("Returned_en_Odoo", ["Numero_Odoo", "Fecha", "Cliente", "Sucursal", "Total_SRI", "Total_Odoo", "Clave_Acceso_SRI"],
          [[r.num, r.fecha, r.socio, r.suc, r.total, r.total_odoo, r.clave] for r in d.itertuples()],
          [19, 12, 42, 13, 12, 12, 52])
    d = c[c["en_sri"] & c["en_odoo"] & ((c["dif_dias"].abs() > 0) | c["serie_dif"])]
    lista("Fecha_o_Serie_distinta", ["Serie_SRI", "Numero_Odoo", "Fecha_SRI", "Fecha_Odoo", "Dif_Dias", "Cliente",
                                      "Clave_Acceso"],
          [[r.serie, r.num, r.fe, r.fecha, int(r.dif_dias), r.socio, r.clave] for r in d.itertuples()],
          [19, 19, 12, 12, 9, 42, 52])
    wb.close()
    return c


def main(carpeta, salida, desde="2025-04-01", hasta="2026-09-30"):
    s = leer_sri(carpeta)
    o = leer_odoo(os.path.join(carpeta, "fact_emi_2025_2026.xlsx"))
    c = conciliar(s, o)
    c = construir_excel(c, salida, desde, hasta)
    print(c["res"].value_counts().to_string())
    print("SRI", len(s), "Odoo", len(o))
    return c


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    main(*sys.argv[1:4])
