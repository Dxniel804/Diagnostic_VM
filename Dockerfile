# Usa uma imagem leve do Python
FROM python:3.11-slim

# Define o diretório de trabalho
WORKDIR /app

# Copia os arquivos do seu projeto
COPY requirements.txt .

# Instala as dependências do Python
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Porta que o Render geralmente usa
EXPOSE 5000

# Comando para rodar a aplicação (ajuste 'app:app' se seu arquivo tiver outro nome)
CMD ["gunicorn", "--bind", "0.0.0.0:8080", "--timeout", "300", "--workers", "1", "--threads", "4", "--worker-class", "gthread", "app:app"]