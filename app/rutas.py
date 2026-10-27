# -*- coding: utf-8 -*-
"""Rutas de ejecución. Todo cuelga de la raíz del repositorio (volúmenes de Easypanel)."""
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
JOBS_DIR = RAIZ / "jobs"
CACHE_XML_DIR = RAIZ / "cache" / "xml"
DATA_DIR = RAIZ / "data"
