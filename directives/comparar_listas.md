# Directive: Comparar Listas de Preços

## Objetivo
Comparar duas versões de uma lista de preços em Excel e identificar:
1. **Produtos novos** — presentes apenas na lista nova
2. **Alterações de preço** — nas colunas `1-5`, `5-10` e `10+`
3. **Percentual de variação** — positivo = aumento, negativo = redução

Gerar planilha de saída apenas quando houver diferenças.

## Inputs
- `arquivo_antigo`: arquivo Excel com a lista de preços antiga (.xlsx / .xls)
- `arquivo_novo`: arquivo Excel com a lista de preços nova (.xlsx / .xls)

## Colunas esperadas
| Coluna           | Candidatos detectados automaticamente                            |
|------------------|------------------------------------------------------------------|
| PN (chave)       | `PN`, `PART NUMBER`, `PART_NUMBER`, `CODIGO`, `CÓDIGO`, `COD`  |
| Descrição        | `DESC`, `DESCRIÇÃO`, `DESCRICAO`, `NOME`, `PRODUTO`              |
| Preço faixa 1–5  | qualquer coluna contendo `1-5`                                   |
| Preço faixa 5–10 | qualquer coluna contendo `5-10`                                  |
| Preço faixa 10+  | qualquer coluna contendo `10+`                                   |

## Script de execução
```
execution/comparar_listas.py
```
Chamado indiretamente via `app.py` (servidor Flask) na rota `POST /comparar`.

## Outputs
- **Sem diferença**: mensagem na tela, nenhum arquivo gerado
- **Com diferença**: arquivo `comparacao_precos.xlsx` para download
  - Todas as colunas da lista nova
  - Colunas `% 1-5`, `% 5-10`, `% 10+` inseridas após a coluna `10+`
  - Coluna `STATUS` ao final (`NOVO`, `ALTERADO`, `IGUAL`)
  - Formatação por cor (verde = novo, vermelho = aumento, amarelo = redução)

## Estratégia de match entre listas
1. **Match exato por PN** (normalizado: trim + upper)
2. **Match fuzzy por PN** (rapidfuzz ≥ 85%) — cobre typos e variações
3. **Match fuzzy por descrição** (rapidfuzz partial_ratio ≥ 80%) — cobre mudança de PN com descrição similar
4. **Produto novo** — nenhum match encontrado

## Como rodar o servidor web
```bash
# 1. Ambiente virtual
python -m venv .venv && source .venv/bin/activate

# 2. Instalar dependências
pip install -e .

# 3. Iniciar servidor
python app.py

# 4. Acessar no navegador
# http://localhost:5000
```

## Edge Cases & Aprendizados
- Preços com vírgula decimal (`R$ 1.234,56`) → `_to_float()` normaliza automaticamente
- Colunas com prefixo `R$` → removido antes da conversão
- Ordem dos produtos não importa — matching é feito por lookup dict
- Planilhas com até 50 MB são aceitas pelo servidor
- Se PN não for encontrado em nenhum arquivo → erro claro exibido na UI

## Limites conhecidos
- Fuzzy match pode gerar falsos positivos em PNs muito curtos (< 4 caracteres)
- O threshold de 85% para PN pode ser ajustado em `comparar_listas.py` → `_fuzzy_match_pn`
