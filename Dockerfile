FROM python:3.11-slim

WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

# Deps systeme pour docTR (OpenCV + decodage PDF)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libgl1 \
    libglib2.0-0 \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

# Install Python deps en premier (cache Docker)
COPY requirements.txt ./
RUN pip install --upgrade pip && pip install -r requirements.txt

# Copie du code
COPY . .

# Pas de serveur : CMD neutre, l'utilisateur lance via `docker compose run`
CMD ["python", "--version"]
