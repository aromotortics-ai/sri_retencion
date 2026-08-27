"""
Parser del archivo TXT que exporta el portal SRI en línea desde
"Comprobantes electrónicos recibidos" -> Exportar Txt.

Formato real observado:
  - Codificación: ISO-8859-1 (Latin-1), no UTF-8.
  - Delimitador: tabulador.
  - 12 columnas, con encabezado en la primera fila:
    RUC_EMISOR, RAZON_SOCIAL_EMISOR, TIPO_COMPROBANTE, SERIE_COMPROBANTE,
    CLAVE_ACCESO, FECHA_AUTORIZACION, FECHA_EMISION, IDENTIFICACION_RECEPTOR,
    VALOR_SIN_IMPUESTOS, IVA, IMPORTE_TOTAL, NUMERO_DOCUMENTO_MODIFICADO

IMPORTANTE (confirmado con tu archivo real): para "Comprobante de Retención",
el SRI deja VACÍAS las columnas VALOR_SIN_IMPUESTOS, IVA e IMPORTE_TOTAL en
este export — el portal no calcula esos totales para retenciones (a
diferencia de facturas). Los valores retenidos reales solo existen dentro
del XML individual de cada comprobante (ver parse_comprobante_xml.py), en el
detalle por cada <impuesto>. Por eso este TXT sirve para obtener la LISTA de
comprobantes y sus claves de acceso, pero no para los montos.
"""
import sys
import pandas as pd

COLUMNAS_ESPERADAS = [
    "RUC_EMISOR", "RAZON_SOCIAL_EMISOR", "TIPO_COMPROBANTE", "SERIE_COMPROBANTE",
    "CLAVE_ACCESO", "FECHA_AUTORIZACION", "FECHA_EMISION", "IDENTIFICACION_RECEPTOR",
    "VALOR_SIN_IMPUESTOS", "IVA", "IMPORTE_TOTAL", "NUMERO_DOCUMENTO_MODIFICADO",
]


def parse_recibidos_txt(path, encoding="ISO-8859-1"):
    df = pd.read_csv(path, sep="\t", dtype=str, encoding=encoding)
    df.columns = [c.strip() for c in df.columns]

    faltantes = [c for c in COLUMNAS_ESPERADAS if c not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas esperadas en el TXT: {faltantes}")

    for col in ("VALOR_SIN_IMPUESTOS", "IVA", "IMPORTE_TOTAL"):
        df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def resumen_vacios(df):
    total = len(df)
    resumen = {}
    for col in ("VALOR_SIN_IMPUESTOS", "IVA", "IMPORTE_TOTAL"):
        vacios = df[col].isna().sum()
        resumen[col] = (vacios, total)
    return resumen


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python parse_recibidos_txt.py archivo.txt [claves_salida.txt]")
        sys.exit(1)

    df = parse_recibidos_txt(sys.argv[1])
    print(f"{len(df)} comprobantes leídos.")

    print("\nCampos de valor vacíos (columna: vacíos / total):")
    for col, (vacios, total) in resumen_vacios(df).items():
        pct = 100 * vacios / total if total else 0
        print(f"  {col}: {vacios}/{total}  ({pct:.0f}% vacío)")

    print("\nComprobantes por tipo:")
    print(df["TIPO_COMPROBANTE"].value_counts().to_string())

    if len(sys.argv) >= 3:
        with open(sys.argv[2], "w", encoding="utf-8") as f:
            for clave in df["CLAVE_ACCESO"].dropna():
                f.write(clave.strip() + "\n")
        print(f"\n{df['CLAVE_ACCESO'].notna().sum()} claves de acceso escritas en {sys.argv[2]}")
