# Python 3.14ベースイメージ
FROM python:3.14-slim

# 作業ディレクトリ
WORKDIR /app

# システムパッケージのインストール（Tesseract & Poppler & OpenCV依存関係）
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-jpn \
    poppler-utils \
    libgl1 \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

# 環境変数を設定（重要！）
ENV RUNNING_IN_DOCKER=1

# Pythonライブラリのインストール
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# アプリケーションファイルをコピー
COPY . .

# ポート公開
EXPOSE 8000

# アプリ起動（Gunicornを使用）
CMD ["gunicorn", "--bind", "0.0.0.0:8000", "--timeout", "300", "app:app"]