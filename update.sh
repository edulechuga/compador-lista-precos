#!/bin/bash
# Script de Atualização para o Container LXC
# Coloque este script na raiz do projeto dentro do container e dê permissão de execução: chmod +x update.sh

echo "Iniciando atualização do serviço..."

# 1. Puxar as últimas alterações do GitHub
echo "Puxando alterações do GitHub..."
git pull origin main

# 2. Atualizar dependências (se necessário)
echo "Atualizando dependências do Python..."
source .venv/bin/activate
pip install -r requirements.txt

# 3. Reiniciar o serviço (substitua 'nome_do_servico' pelo nome correto do seu serviço systemd)
# Exemplo: sudo systemctl restart meu_app.service
echo "Reiniciando o serviço..."
# sudo systemctl restart NOME_DO_SEU_SERVICO.service

echo "Atualização concluída com sucesso!"
