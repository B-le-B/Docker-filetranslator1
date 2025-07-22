FROM python:3.9-slim

WORKDIR /app

# 安装pandoc和基本依赖
RUN apt-get update && apt-get install -y \
    pandoc \
    curl \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 7860

CMD ["python", "app.py"]
