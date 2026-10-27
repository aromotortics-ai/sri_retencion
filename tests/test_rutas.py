from app.rutas import CACHE_XML_DIR, DATA_DIR, JOBS_DIR, RAIZ


def test_rutas_apuntan_a_la_raiz():
    assert (RAIZ / "Dockerfile").exists()
    assert JOBS_DIR == RAIZ / "jobs"
    assert CACHE_XML_DIR == RAIZ / "cache" / "xml"
    assert DATA_DIR == RAIZ / "data"
