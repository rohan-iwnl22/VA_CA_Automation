FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY src/ src/
COPY static/ static/
COPY templates/ templates/

EXPOSE 8000

CMD ["uvicorn", "src.va_ca_automation.api.main:app", "--host", "0.0.0.0", "--port", "8000"]