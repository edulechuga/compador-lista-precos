"""
Script de exemplo — estrutura base para todos os scripts de execução.

Uso:
    python execution/exemplo_tarefa.py

Descrição:
    Demonstra o padrão de script da Layer 3:
    - Carrega .env via utils
    - Aceita argumentos (argparse)
    - Faz logging estruturado
    - Salva resultados intermediários em .tmp/
    - Falha com mensagens claras
"""
import argparse
import sys
from utils import get_logger, save_json

logger = get_logger("exemplo_tarefa")


def main(input1: str, input2: str) -> None:
    logger.info("Iniciando tarefa com input1=%s, input2=%s", input1, input2)

    # ── substitua esta lógica pela tarefa real ──
    resultado = {
        "status": "ok",
        "input1": input1,
        "input2": input2,
        "output": f"{input1} + {input2} processados com sucesso",
    }
    # ────────────────────────────────────────────

    output_path = save_json(resultado, "resultado.json")
    logger.info("Resultado salvo em %s", output_path)
    print(f"✅ Concluído. Resultado em: {output_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Script de exemplo")
    parser.add_argument("--input1", required=True, help="Primeiro input")
    parser.add_argument("--input2", required=True, help="Segundo input")
    args = parser.parse_args()
    main(args.input1, args.input2)
