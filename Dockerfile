FROM python:3.12-slim

WORKDIR /app

# Instalar dependencias primero (capa cacheable)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar código fuente
COPY *.py .
COPY templates/ templates/

# Crear directorios para volúmenes persistentes
RUN mkdir -p cache/xml jobs

EXPOSE 8000

CMD ["gunicorn", "app:app", \
     "-w", "2", \
     "-k", "uvicorn.workers.UvicornWorker", \
     "--bind", "0.0.0.0:8000", \
     "--timeout", "30"]
