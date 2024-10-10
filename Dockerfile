# Utilizando imagem Python para amd64 com versão completa
FROM --platform=linux/amd64 python:3.9

# Definindo o diretório de trabalho dentro do container
WORKDIR /app

# Copiando os arquivos de dependências
COPY requirements.txt .

# Instalando as dependências do Python
RUN pip install --no-cache-dir -r requirements.txt

# Instalando sshpass para suportar a execução de comandos remotos via SSH
RUN apt-get update && apt-get install -y sshpass

# Copiando o restante do projeto para dentro do container
COPY . .

# Comando para executar o script Python ao iniciar o container
CMD [ "python", "script.py" ]
