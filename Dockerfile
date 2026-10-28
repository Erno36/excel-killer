# Dockerfile — cruzador-excel-web
# Build context debe ser desarrollo/ (directorio padre):
#   docker build -f cruzador_excel_web/Dockerfile -t cruzador-excel-web ..
# O usar docker-compose.yml que configura el contexto automáticamente.
#
# La app importa módulos de ../cruzador_excel/ via sys.path en app.py.
# Por eso el contexto de build es el directorio padre (desarrollo/).

FROM python:3.12-slim

# Dependencias de sistema mínimas para pandas/openpyxl
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Estructura en el contenedor:
#   /app/cruzador_excel/      <- módulos compartidos del motor
#   /app/cruzador_excel_web/  <- aplicación web
WORKDIR /app

COPY cruzador_excel/ ./cruzador_excel/
COPY cruzador_excel_web/ ./cruzador_excel_web/

WORKDIR /app/cruzador_excel_web

RUN pip install --no-cache-dir -r requirements.txt

# Directorios de datos persistentes (montados como volúmenes en producción)
RUN mkdir -p /var/lib/cruzador-excel-web/recipes \
             /var/log/cruzador-excel-web

# Usuario no-root para reducir superficie de ataque
RUN useradd --system --no-create-home --shell /bin/false cruzador
RUN chown -R cruzador:cruzador /var/lib/cruzador-excel-web /var/log/cruzador-excel-web

USER cruzador

EXPOSE 8000

# 1 worker, 4 threads — CRÍTICO: múltiples workers rompen _workflows (estado en RAM)
# Timeout 120s para operaciones pesadas con archivos Excel grandes
CMD ["gunicorn", \
     "--workers", "1", \
     "--threads", "4", \
     "--bind", "0.0.0.0:8000", \
     "--timeout", "120", \
     "--access-logfile", "-", \
     "--error-logfile", "-", \
     "app:app"]
