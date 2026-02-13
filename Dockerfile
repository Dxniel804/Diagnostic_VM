# Usa uma imagem leve do Python
FROM python:3.11-slim

# Instala dependências de sistema para o pandas e lxml
RUN apt-get update && apt-get install -y \
    build-essential \
    libxml2-dev \
    libxslt-dev \
    && rm -rf /var/lib/apt/lists/*

# Define o diretório de trabalho
WORKDIR /app

# Copia os arquivos do seu projeto
COPY . .

# Instala as dependências do Python
RUN pip install --no-cache-dir -r requirements.txt

# Porta que o Render geralmente usa
EXPOSE 8080

# Comando para rodar a aplicação (ajuste 'app:app' se seu arquivo tiver outro nome)
CMD ["gunicorn", "--bind", "0.0.0.0:8080", "--timeout", "300", "--workers", "1", "--threads", "4", "--worker-class", "gthread", "app:app"]