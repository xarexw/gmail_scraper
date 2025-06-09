# Базовий образ
FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    build-essential \
    libssl-dev \
    libffi-dev \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Створюємо робочу директорію всередині контейнера
WORKDIR /app

# Копіюємо залежності й встановлюємо
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Копіюємо весь код, крім того, що проігноровано в .dockerignore
COPY . .

# Гарантуємо, що папка для завантажень існує
RUN mkdir -p price_files

# Запускаємо скрипт
CMD ["python", "download_prices.py"]
