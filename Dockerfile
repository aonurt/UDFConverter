FROM python:3.9-slim

# LibreOffice ve Java Kurulumu (PDF çevirisi için)
RUN apt-get update && apt-get install -y \
    libreoffice \
    default-jre \
    libreoffice-java-common \
    && apt-get clean

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# DEĞİŞİKLİK BURADA: --timeout 120 ekledik (İşlem süresini 30 saniyeden 2 dakikaya çıkardık)
CMD ["gunicorn", "--timeout", "120", "-b", "0.0.0.0:10000", "app:app"]
