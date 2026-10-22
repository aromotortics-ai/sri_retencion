# -*- coding: utf-8 -*-
"""
Mueve los TXT descargados del SRI (2390000831001_EmitidosAutorizados*.txt) de
Descargas a facturas_emitidas/emitidos_AAAA-MM-DD.txt, usando la fecha de
emisión que trae el propio archivo.

Uso: python mover_emitidos.py <epoch_desde> [destino]
  Solo toma archivos con fecha de modificación >= epoch_desde.
"""
import glob
import os
import shutil
import sys
from collections import Counter
from datetime import datetime

DESCARGAS = os.path.expanduser("~/Downloads")


def main(desde, destino="facturas_emitidas"):
    os.makedirs(destino, exist_ok=True)
    archivos = [f for f in glob.glob(os.path.join(DESCARGAS, "2390000831001_EmitidosAutorizados*.txt"))
                if os.path.getmtime(f) >= desde]
    for f in sorted(archivos, key=os.path.getmtime):
        with open(f, encoding="iso-8859-1") as fh:
            lineas = fh.read().splitlines()
        cab = lineas[0].split("\t")
        if "FECHA_EMISION" not in cab or len(lineas) < 2:
            print("OMITIDO (sin datos):", os.path.basename(f))
            continue
        i = cab.index("FECHA_EMISION")
        fechas = Counter(l.split("\t")[i][:10] for l in lineas[1:] if l.strip())
        dia = fechas.most_common(1)[0][0]
        nombre = datetime.strptime(dia, "%d/%m/%Y").strftime("emitidos_%Y-%m-%d.txt")
        if len(fechas) > 1:
            print("AVISO: varias fechas en", os.path.basename(f), dict(fechas))
        shutil.move(f, os.path.join(destino, nombre))
        print(nombre, len(lineas) - 1)


if __name__ == "__main__":
    main(float(sys.argv[1]), *sys.argv[2:3])
