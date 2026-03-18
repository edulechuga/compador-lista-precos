# Agente — Instância Pessoal

Sistema de agente com arquitetura de 3 camadas conforme `AGENTE.md`.

## Estrutura

```
.
├── AGENTE.md                  # Instruções para o agente (Layer 2)
├── directives/                # SOPs em Markdown (Layer 1 — O QUÊ fazer)
│   └── _template.md           # Template para novas diretivas
├── execution/                 # Scripts Python determinísticos (Layer 3 — COMO fazer)
│   ├── utils.py               # Utilitários compartilhados
│   └── exemplo_tarefa.py      # Script de exemplo / template
├── .tmp/                      # Arquivos intermediários (nunca commitar)
├── .env                       # Variáveis de ambiente (nunca commitar)
├── .env.example               # Template do .env
├── .gitignore
└── pyproject.toml
```

## Setup inicial

```bash
# 1. Criar ambiente virtual e instalar dependências
python -m venv .venv
source .venv/bin/activate       # mac/linux
# ou .venv\Scripts\activate     # windows

pip install -e .

# 2. Configurar variáveis de ambiente
cp .env.example .env
# edite .env com suas chaves de API

# 3. (Opcional) Google OAuth
# Coloque credentials.json na raiz e rode o script que precisar de auth Google
```

## Como usar

1. **Nova tarefa?** → Crie um arquivo em `directives/` usando `_template.md` como base
2. **Novo script?** → Copie `execution/exemplo_tarefa.py` e adapte
3. **Algo quebrou?** → O agente lê o erro, corrige o script, atualiza a diretiva
4. **Entregáveis** → Sempre em serviços cloud (Google Sheets, Slides, etc.)
5. **Intermediários** → Sempre em `.tmp/` (descartáveis)

## Fluxo do agente

```
Usuário pede algo
    ↓
Agente lê diretiva em directives/
    ↓
Agente chama script em execution/
    ↓
Script processa → salva em .tmp/
    ↓
Agente entrega resultado (cloud) ao usuário
    ↓
Se erro → corrige script → atualiza diretiva → testa
```
