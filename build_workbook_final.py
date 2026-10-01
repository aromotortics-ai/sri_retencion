from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from parse_recibidos_txt import parse_recibidos_txt
from parse_comprobante_xml import parse_folder

FONT = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
WARN_FILL = PatternFill("solid", fgColor="C00000")
HEADER_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
BODY_FONT = Font(name=FONT, size=10)
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def escribir_encabezado(ws, headers):
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=j, value=h.replace("_", " "))
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER


def build_workbook(txt_path: str, xml_glob: str, output_path: str) -> dict:
    """
    Construye el workbook Excel de conciliación SRI.

    Args:
        txt_path: Ruta al archivo TXT de comprobantes recibidos
        xml_glob: Glob pattern para archivos XML (ej. "xml_samples/[0-9]*.xml")
        output_path: Ruta de salida del archivo Excel

    Returns:
        dict con claves n_listado, n_xml, n_pendientes
    """
    df_listado = parse_recibidos_txt(txt_path)          # 578 filas: TODO el mes, sin montos
    df_cab, df_det = parse_folder(xml_glob)              # solo los XML ya descargados

    # El caché de XML es compartido entre períodos: quedarse solo con los
    # comprobantes del TXT cargado, o se mezclan retenciones de otros meses.
    claves_txt = set(df_listado["CLAVE_ACCESO"].dropna())
    if len(df_cab):
        df_cab = df_cab[df_cab["Clave_Acceso"].isin(claves_txt)].reset_index(drop=True)
    if len(df_det):
        df_det = df_det[df_det["Clave_Acceso"].isin(claves_txt)].reset_index(drop=True)

    claves_con_xml = set(df_cab["Clave_Acceso"]) if len(df_cab) else set()
    df_listado["XML_Descargado"] = df_listado["CLAVE_ACCESO"].isin(claves_con_xml)

    if len(df_det):
        # N° Retencion: Establecimiento-Punto_Emision-Secuencial  (ej. 001-200-000004340)
        df_det.insert(
            df_det.columns.get_loc("Secuencial") + 1,
            "N_Retencion",
            df_det["Establecimiento"] + "-" + df_det["Punto_Emision"] + "-" + df_det["Secuencial"],
        )
        # Num_Doc_Sustento con guiones: primeros 3 + "-" + sig 3 + "-" + resto  (ej. 001-003-000088766)
        df_det["Num_Doc_Sustento"] = df_det["Num_Doc_Sustento"].apply(
            lambda v: f"{v[:3]}-{v[3:6]}-{v[6:]}" if isinstance(v, str) and len(v) >= 6 else v
        )
        df_agg = (df_det.groupby(["Num_Doc_Sustento", "RUC_Emisor", "Razon_Social_Emisor",
                                   "Fecha_Emision_Doc_Sustento"], as_index=False)
                  .agg(Valor_Total_Retenido=("Valor_Retenido", "sum")))
    else:
        df_agg = df_det  # vacío

    wb = Workbook()

    # ---------- Hoja 1: SRI_Listado (las 578 filas del TXT, sin montos) ----------
    ws1 = wb.active
    ws1.title = "SRI_Listado"
    headers1 = list(df_listado.columns)
    escribir_encabezado(ws1, headers1)
    col_xml_desc = headers1.index("XML_Descargado") + 1
    for i, row in enumerate(df_listado.itertuples(index=False), start=2):
        for j, val in enumerate(row, start=1):
            c = ws1.cell(row=i, column=j, value=val)
            c.font = BODY_FONT
            c.border = BORDER
        c_flag = ws1.cell(row=i, column=col_xml_desc)
        if not c_flag.value:
            c_flag.fill = WARN_FILL
            c_flag.font = Font(name=FONT, size=10, color="FFFFFF", bold=True)
    for j in range(1, len(headers1) + 1):
        ws1.column_dimensions[get_column_letter(j)].width = 20
    ws1.freeze_panes = "A2"
    n_listado = len(df_listado)

    resumen_row = n_listado + 3
    ws1.cell(row=resumen_row, column=1, value="Total comprobantes del período:").font = Font(name=FONT, bold=True, size=10)
    ws1.cell(row=resumen_row, column=2, value=f"=COUNTA(A2:A{1+n_listado})").font = Font(name=FONT, bold=True, size=10)
    ws1.cell(row=resumen_row + 1, column=1, value="Con XML descargado (con detalle de valores):").font = Font(name=FONT, bold=True, size=10)
    ws1.cell(row=resumen_row + 1, column=2, value=f"=COUNTIF({get_column_letter(col_xml_desc)}2:{get_column_letter(col_xml_desc)}{1+n_listado},TRUE)").font = Font(name=FONT, bold=True, size=10)
    ws1.cell(row=resumen_row + 2, column=1, value="Pendientes de descargar:").font = Font(name=FONT, bold=True, size=10)
    ws1.cell(row=resumen_row + 2, column=2, value=f"=COUNTIF({get_column_letter(col_xml_desc)}2:{get_column_letter(col_xml_desc)}{1+n_listado},FALSE)").font = Font(name=FONT, bold=True, size=10)

    # ---------- Hoja 2: SRI_Detalle_Retenciones (de los XML ya descargados) ----------
    ws2 = wb.create_sheet("SRI_Detalle_Retenciones")
    if len(df_det):
        headers2 = list(df_det.columns)
        escribir_encabezado(ws2, headers2)
        for i, row in enumerate(df_det.itertuples(index=False), start=2):
            for j, val in enumerate(row, start=1):
                c = ws2.cell(row=i, column=j, value=val)
                c.font = BODY_FONT
                c.border = BORDER
                if headers2[j - 1] in ("Base_Imponible", "Valor_Retenido"):
                    c.number_format = '#,##0.00'
        for j in range(1, len(headers2) + 1):
            ws2.column_dimensions[get_column_letter(j)].width = 20
        ws2.freeze_panes = "A2"
    else:
        ws2["A1"] = "Aún no hay XML descargados en xml_samples/. Corre descargar_comprobantes_sri.py primero."

    # ---------- Hoja 3: SRI_Por_Doc_Sustento ----------
    ws3 = wb.create_sheet("SRI_Por_Doc_Sustento")
    n_agg = len(df_agg)
    if n_agg:
        headers3 = list(df_agg.columns)
        escribir_encabezado(ws3, headers3)
        for i, row in enumerate(df_agg.itertuples(index=False), start=2):
            for j, val in enumerate(row, start=1):
                c = ws3.cell(row=i, column=j, value=val)
                c.font = BODY_FONT
                c.border = BORDER
                if headers3[j - 1] == "Valor_Total_Retenido":
                    c.number_format = '#,##0.00'
        for j in range(1, len(headers3) + 1):
            ws3.column_dimensions[get_column_letter(j)].width = 22
        ws3.freeze_panes = "A2"
    else:
        ws3["A1"] = "Sin datos todavía (depende de los XML descargados)."

    # ---------- Hoja 4: ERP_Compras (plantilla) ----------
    ws4 = wb.create_sheet("ERP_Compras")
    headers4 = ["Num_Documento_Compra", "RUC_Proveedor", "Razon_Social",
                "Fecha_Comprobante", "Base_Imponible", "Total_Retenido_ERP"]
    escribir_encabezado(ws4, headers4)
    ejemplo = ["001003000088744", "1001789088001", "FUENTES VASQUEZ OSCAR VINICIO",
               "2026-07-31", 109.36, 2.24]
    for j, val in enumerate(ejemplo, start=1):
        c = ws4.cell(row=2, column=j, value=val)
        c.font = Font(name=FONT, size=10, italic=True, color="808080")
        c.border = BORDER
    N_ERP = 600
    for r in range(3, 3 + N_ERP):
        for j in range(1, len(headers4) + 1):
            c = ws4.cell(row=r, column=j)
            c.fill = INPUT_FILL
            c.border = BORDER
            c.font = BODY_FONT
    for j, w in enumerate([22, 16, 32, 16, 14, 16], start=1):
        ws4.column_dimensions[get_column_letter(j)].width = w
    ws4.freeze_panes = "A2"
    ws4.cell(row=3 + N_ERP + 1, column=1,
             value=('Leyenda: fila 2 (cursiva) es ejemplo — bórrela. "Num Documento Compra" debe '
                    'ser el mismo número que numDocSustento del XML del SRI, solo dígitos.')
             ).font = Font(name=FONT, size=9, italic=True, color="808080")
    n_erp_start, n_erp_end = 3, 3 + N_ERP - 1

    # ---------- Hoja 5: Conciliacion ----------
    ws5 = wb.create_sheet("Conciliacion")
    headers5 = ["Num_Doc_Sustento", "RUC_Emisor", "Razon_Social_Emisor",
                "Valor_Retenido_SRI", "Total_Retenido_ERP", "Diferencia", "Estado"]
    escribir_encabezado(ws5, headers5)
    if n_agg:
        for i in range(2, 2 + n_agg):
            ws5.cell(row=i, column=1, value=f"=SRI_Por_Doc_Sustento!A{i}")
            ws5.cell(row=i, column=2, value=f"=SRI_Por_Doc_Sustento!B{i}")
            ws5.cell(row=i, column=3, value=f"=SRI_Por_Doc_Sustento!C{i}")
            ws5.cell(row=i, column=4, value=f"=SRI_Por_Doc_Sustento!E{i}")
            ws5.cell(
                row=i, column=5,
                value=(f'=IFERROR(INDEX(ERP_Compras!$F${n_erp_start}:$F${n_erp_end},'
                        f'MATCH(A{i},ERP_Compras!$A${n_erp_start}:$A${n_erp_end},0)),"")')
            )
            ws5.cell(row=i, column=6, value=f'=IF(E{i}="","",D{i}-E{i})')
            ws5.cell(
                row=i, column=7,
                value=(f'=IF(E{i}="","No encontrado en ERP",'
                        f'IF(ABS(D{i}-E{i})<0.01,"Coincide","Diferencia de valor"))')
            )
            for col in range(1, 8):
                c = ws5.cell(row=i, column=col)
                c.font = BODY_FONT
                c.border = BORDER
                if headers5[col - 1] in ("Valor_Retenido_SRI", "Total_Retenido_ERP", "Diferencia"):
                    c.number_format = '#,##0.00'
    for j, w in enumerate([22, 16, 32, 18, 18, 12, 20], start=1):
        ws5.column_dimensions[get_column_letter(j)].width = w
    ws5.freeze_panes = "A2"

    wb.save(output_path)

    return {
        "n_listado": n_listado,
        "n_xml": len(df_cab),
        "n_pendientes": n_listado - len(df_cab),
    }


if __name__ == "__main__":
    stats = build_workbook(
        txt_path="2390000831001_Recibidos.txt",
        xml_glob="xml_samples/[0-9]*.xml",
        output_path="Comprobantes_Retencion_Agosto2026_Final.xlsx",
    )
    print(f"OK. Listado: {stats['n_listado']} | XML: {stats['n_xml']} | Pendientes: {stats['n_pendientes']}")
