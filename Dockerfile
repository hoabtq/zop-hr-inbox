FROM python:3.11-slim

# Cài LibreOffice Writer để convert docx → PDF
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice-writer \
        libreoffice-java-common \
        default-jre-headless \
        fonts-liberation \
        fonts-dejavu && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p uploads output_grouped archive_data

EXPOSE 8080

ENV PORT=8080
# Tắt sandbox của LibreOffice khi chạy trong container
ENV LIBREOFFICE_ARGS="--headless --norestore --nofirststartwizard"

CMD ["python", "-u", "app.py"]