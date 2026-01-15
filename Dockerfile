FROM python:3.9-slim

# LibreOffice'i (PDF çeviri için) yükle
RUN apt-get update && apt-get install -y \
    libreoffice \
    default-jre \
    libreoffice-java-common \
    && apt-get clean

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Uygulamayı başlat
CMD ["gunicorn", "-b", "0.0.0.0:10000", "app:app"]