# Dockerfile
FROM python:3.9-slim

WORKDIR /app

# Instalar dependências do sistema
RUN apt-get update && apt-get install -y \
    gcc \
    python3-dev \
    && rm -rf /var/lib/apt/lists/*

# Copiar requirements primeiro para cache
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copiar o restante dos arquivos
COPY . .

# Configurar porta e comando de execução
EXPOSE 8080
ENTRYPOINT ["streamlit", "run", "main.py", "--server.port=8080", "--server.address=0.0.0.0"]