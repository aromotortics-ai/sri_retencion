"""
Concilia SRI_Detalle_Retenciones vs reporte de retenciones de Odoo.

Clave de cruce: N° Retencion  +  Factura/Num Doc Sustento  +  Porcentaje

Salida: nuevas hojas en Comprobantes_Retencion_Agosto2026_Final.xlsx
  - Conciliacion_Odoo  : resumen ejecutivo por retención
  - Solo_en_SRI        : retenciones/líneas del SRI sin registro en Odoo
  - Solo_en_Odoo       : retenciones/líneas en Odoo sin autorización SRI
  - Diferencias_Valor  : líneas presentes en ambos con montos distintos
"""
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

EXCEL_FINAL = "Comprobantes_Retencion_Agosto2026_Final.xlsx"
EXCEL_ODOO  = "reten_odoo.xlsx"

FONT      = "Arial"
HDR_FILL  = PatternFill("solid", fgColor="1F4E78")
HDR_FONT  = Font(name=FONT, bold=True, color="FFFFFF", size=10)
OK_FILL   = PatternFill("solid", fgColor="E2EFDA")   # verde claro
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")   # naranja claro
ERR_FILL  = PatternFill("solid", fgColor="FFDCE1")   # rojo claro
BODY_FONT = Font(name=FONT, size=10)
THIN      = Side(style="thin", color="BFBFBF")
BORDER    = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NUM_FMT   = '#,##0.00'


def _hdr(ws, cols, widths=None):
    for j, h in enumerate(cols, 1):
        c = ws.cell(1, j, h)
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    if widths:
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A2"


def _write_df(ws, df, money_cols=(), fill_col=None):
    cols = list(df.columns)
    _hdr(ws, cols, widths=[max(14, len(c)+2) for c in cols])
    for i, row in enumerate(df.itertuples(index=False), 2):
        row_fill = None
        if fill_col and fill_col in cols:
            estado = getattr(row, fill_col.replace(" ", "_").replace("°","").strip(), None)
            if isinstance(estado, str):
                if "Solo SRI"   in estado: row_fill = WARN_FILL
                elif "Solo Odoo" in estado: row_fill = ERR_FILL
                elif "Diferencia" in estado: row_fill = PatternFill("solid", fgColor="FFF2CC")
                elif "Coincide"   in estado: row_fill = OK_FILL
        for j, val in enumerate(row, 1):
            c = ws.cell(i, j, val)
            c.font = BODY_FONT
            c.border = BORDER
            if cols[j-1] in money_cols:
                c.number_format = NUM_FMT
            if row_fill:
                c.fill = row_fill


# ── Leer SRI ────────────────────────────────────────────────────────────────
wb_main = openpyxl.load_workbook(EXCEL_FINAL)
ws_sri  = wb_main["SRI_Detalle_Retenciones"]
hdr_sri = [ws_sri.cell(1, j).value for j in range(1, ws_sri.max_column + 1)]
df_sri  = pd.DataFrame(
    [{hdr_sri[j-1]: ws_sri.cell(i, j).value for j in range(1, ws_sri.max_column + 1)}
     for i in range(2, ws_sri.max_row + 1)]
)
df_sri = df_sri.rename(columns={
    "N Retencion":       "N_Ret",
    "Num Doc Sustento":  "Factura",
    "Razon Social Emisor": "Proveedor_SRI",
    "Base Imponible":    "Base_SRI",
    "Porcentaje Retener":"Pct",
    "Valor Retenido":    "Valor_SRI",
    "Codigo Retencion":  "Cod_SRI",
})

# ── Leer Odoo ────────────────────────────────────────────────────────────────
wb_odoo  = openpyxl.load_workbook(EXCEL_ODOO)
ws_odoo  = wb_odoo.active
hdr_odoo = [ws_odoo.cell(1, j).value for j in range(1, ws_odoo.max_column + 1)]
df_odoo  = pd.DataFrame(
    [{hdr_odoo[j-1]: ws_odoo.cell(i, j).value for j in range(1, ws_odoo.max_column + 1)}
     for i in range(2, ws_odoo.max_row + 1)]
)
df_odoo = df_odoo.rename(columns={
    "No. de retención":                          "N_Ret",
    "Factura":                                   "Factura",
    "Contacto":                                  "Proveedor_Odoo",
    "Líneas de retención/Base imponible":        "Base_Odoo",
    "Líneas de retención/Porcentaje retención":  "Pct",
    "Líneas de retención/Valor retenido":        "Valor_Odoo",
    "Líneas de retención/% de retención":        "Cod_Odoo",
    "Fecha de retención":                        "Fecha_Ret",
})

# Normalizar porcentaje a float en ambos
df_sri["Pct"]  = pd.to_numeric(df_sri["Pct"],  errors="coerce")
df_odoo["Pct"] = pd.to_numeric(df_odoo["Pct"], errors="coerce")

KEY = ["N_Ret", "Factura", "Pct"]

# ── Merge exacto ──────────────────────────────────────────────────────────────
merged = df_sri[KEY + ["Proveedor_SRI", "Base_SRI", "Valor_SRI", "Cod_SRI"]].merge(
    df_odoo[KEY + ["Proveedor_Odoo", "Base_Odoo", "Valor_Odoo", "Fecha_Ret"]],
    on=KEY, how="outer", indicator=True
)

merged["Estado"] = merged["_merge"].map({
    "left_only":  "Solo SRI (falta en Odoo)",
    "right_only": "Solo Odoo (no está en SRI)",
    "both":       "Coincide",
}).astype(str)

# Detectar diferencias de valor en los que coinciden
tol = 0.02   # tolerancia ±2 centavos
mask_diff = (
    (merged["_merge"] == "both") &
    (abs(merged["Valor_SRI"].fillna(0) - merged["Valor_Odoo"].fillna(0)) > tol)
)
merged.loc[mask_diff, "Estado"] = "Diferencia de valor"
merged["Diferencia"] = (merged["Valor_SRI"].fillna(0) - merged["Valor_Odoo"].fillna(0)).round(2)
merged.loc[merged["_merge"] != "both", "Diferencia"] = None
merged = merged.drop(columns=["_merge"])

# ── Detectar falsos negativos: N°+secuencial+% coinciden pero establecimiento de factura difiere ──
solo_sri_idx  = merged["Estado"] == "Solo SRI (falta en Odoo)"
solo_odoo_idx = merged["Estado"] == "Solo Odoo (no está en SRI)"

tmp_sri  = merged[solo_sri_idx].copy()
tmp_odoo = merged[solo_odoo_idx].copy()
tmp_sri["Fac_seq"]  = tmp_sri["Factura"].str[-9:]
tmp_odoo["Fac_seq"] = tmp_odoo["Factura"].str[-9:]

partial = tmp_sri[["N_Ret","Fac_seq","Pct","Factura"]].merge(
    tmp_odoo[["N_Ret","Fac_seq","Pct","Factura"]],
    on=["N_Ret","Fac_seq","Pct"], suffixes=("_SRI","_Odoo")
)

if len(partial):
    # Marcar estas filas con estado especial
    for _, p in partial.iterrows():
        mask_s = (merged["N_Ret"] == p["N_Ret"]) & (merged["Factura"] == p["Factura_SRI"]) & (merged["Pct"] == p["Pct"])
        mask_o = (merged["N_Ret"] == p["N_Ret"]) & (merged["Factura"] == p["Factura_Odoo"]) & (merged["Pct"] == p["Pct"])
        merged.loc[mask_s, "Estado"] = f"Factura incorrecta en Odoo (SRI:{p['Factura_SRI']} | Odoo:{p['Factura_Odoo']})"
        merged.loc[mask_o, "Estado"] = f"Factura incorrecta en Odoo (SRI:{p['Factura_SRI']} | Odoo:{p['Factura_Odoo']})"

# ── Subconjuntos ─────────────────────────────────────────────────────────────
df_solo_sri   = merged[merged["Estado"] == "Solo SRI (falta en Odoo)"].copy()
df_solo_odoo  = merged[merged["Estado"] == "Solo Odoo (no está en SRI)"].copy()
df_diff_val   = merged[merged["Estado"] == "Diferencia de valor"].copy()
df_fac_error  = merged[merged["Estado"].str.startswith("Factura incorrecta", na=False)].copy()
df_ok         = merged[merged["Estado"] == "Coincide"].copy()

# ── Resumen ejecutivo ────────────────────────────────────────────────────────
resumen = pd.DataFrame([
    ["Total líneas SRI",                              len(df_sri)],
    ["Total líneas Odoo",                             len(df_odoo)],
    ["Líneas coinciden (valor OK)",                   len(df_ok)],
    ["Solo en SRI — faltan en Odoo",                  len(df_solo_sri)],
    ["Solo en Odoo — no están en SRI",                len(df_solo_odoo)],
    ["Factura incorrecta en Odoo (N° correcto)",      len(df_fac_error) // 2],
    ["Coinciden pero con diferencia de valor",        len(df_diff_val)],
    ["Retenciones únicas SRI",                        df_sri["N_Ret"].nunique()],
    ["Retenciones únicas Odoo",                       df_odoo["N_Ret"].nunique()],
], columns=["Concepto", "Cantidad"])

# ── Escribir hojas ───────────────────────────────────────────────────────────
for sheet_name in ["Conciliacion_Odoo", "Solo_en_SRI", "Solo_en_Odoo", "Diferencias_Valor", "Factura_Incorrecta_Odoo"]:
    if sheet_name in wb_main.sheetnames:
        del wb_main[sheet_name]

MONEY = ("Base_SRI", "Valor_SRI", "Base_Odoo", "Valor_Odoo", "Diferencia")

# Hoja resumen
ws_res = wb_main.create_sheet("Conciliacion_Odoo")
ws_res.sheet_properties.tabColor = "1F4E78"
_hdr(ws_res, list(resumen.columns), widths=[45, 12])
for i, row in enumerate(resumen.itertuples(index=False), 2):
    c1 = ws_res.cell(i, 1, row.Concepto)
    c2 = ws_res.cell(i, 2, row.Cantidad)
    c1.font = BODY_FONT; c1.border = BORDER
    c2.font = Font(name=FONT, bold=True, size=10); c2.border = BORDER
    if "Solo en SRI"   in row.Concepto: fill = WARN_FILL
    elif "Solo en Odoo" in row.Concepto: fill = ERR_FILL
    elif "diferencia"   in row.Concepto: fill = PatternFill("solid", fgColor="FFF2CC")
    elif "coinciden"    in row.Concepto.lower() and "diferencia" not in row.Concepto.lower(): fill = OK_FILL
    else: fill = PatternFill()
    c1.fill = fill; c2.fill = fill

# Hoja Solo_en_SRI
ws_ssri = wb_main.create_sheet("Solo_en_SRI")
ws_ssri.sheet_properties.tabColor = "FF9900"
cols_sri = ["N_Ret", "Factura", "Pct", "Proveedor_SRI", "Base_SRI", "Valor_SRI", "Cod_SRI"]
_write_df(ws_ssri, df_solo_sri[cols_sri], money_cols=MONEY)

# Hoja Solo_en_Odoo
ws_sodo = wb_main.create_sheet("Solo_en_Odoo")
ws_sodo.sheet_properties.tabColor = "C00000"
cols_odo = ["N_Ret", "Factura", "Pct", "Proveedor_Odoo", "Base_Odoo", "Valor_Odoo", "Fecha_Ret"]
_write_df(ws_sodo, df_solo_odoo[cols_odo], money_cols=MONEY)

# Hoja Diferencias_Valor
ws_dif = wb_main.create_sheet("Diferencias_Valor")
ws_dif.sheet_properties.tabColor = "FFCC00"
cols_dif = ["N_Ret", "Factura", "Pct",
            "Proveedor_SRI", "Base_SRI", "Valor_SRI",
            "Base_Odoo", "Valor_Odoo", "Diferencia"]
_write_df(ws_dif, df_diff_val[cols_dif], money_cols=MONEY)

# Hoja Factura_Incorrecta_Odoo
ws_fac = wb_main.create_sheet("Factura_Incorrecta_Odoo")
ws_fac.sheet_properties.tabColor = "9933FF"
# Mostrar pares SRI/Odoo con la diferencia clara
fac_err_sri  = df_fac_error[df_fac_error["Proveedor_SRI"].notna()][["N_Ret","Factura","Pct","Proveedor_SRI","Base_SRI","Valor_SRI"]].copy()
fac_err_odoo = df_fac_error[df_fac_error["Proveedor_Odoo"].notna()][["N_Ret","Factura","Pct","Proveedor_Odoo","Base_Odoo","Valor_Odoo"]].copy()
fac_err_sri.columns  = ["N_Ret","Factura_SRI","Pct","Proveedor","Base","Valor"]
fac_err_odoo.columns = ["N_Ret","Factura_Odoo","Pct","Proveedor","Base","Valor"]
fac_combined = fac_err_sri.merge(fac_err_odoo[["N_Ret","Factura_Odoo","Pct"]], on=["N_Ret","Pct"])
fac_combined["Accion"] = "Corregir factura en Odoo: cambiar Factura_Odoo → Factura_SRI"
_write_df(ws_fac, fac_combined, money_cols=("Base","Valor"))

wb_main.save(EXCEL_FINAL)

print("=" * 60)
print(f"  Total líneas SRI                    : {len(df_sri):>6}")
print(f"  Total líneas Odoo                   : {len(df_odoo):>6}")
print(f"  Coinciden (valor OK)                : {len(df_ok):>6}")
print(f"  Solo en SRI — faltan en Odoo        : {len(df_solo_sri):>6}")
print(f"  Solo en Odoo — no están en SRI      : {len(df_solo_odoo):>6}")
print(f"  Factura incorrecta en Odoo           : {len(df_fac_error)//2:>6}")
print(f"  Diferencias de valor                : {len(df_diff_val):>6}")
print("=" * 60)
print(f"Guardado: {EXCEL_FINAL}")
