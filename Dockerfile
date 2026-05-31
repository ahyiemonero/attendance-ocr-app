FROM python:3.11-slim

WORKDIR /app

# Install Tesseract OCR and required libraries
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    libgl1 \
    libglib2.0-0 \
    libjpeg-dev \
    zlib1g-dev \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .

RUN pip install --no-cache-dir --upgrade pip
RUN pip install --no-cache-dir -r requirements.txt
RUN pip install --no-cache-dir gunicorn

COPY . .

RUN mkdir -p uploads/time_in uploads/time_out generated static

EXPOSE 7869

CMD ["gunicorn", "--workers", "1", "--timeout", "300", "--bind", "0.0.0.0:7869", "app:app"]