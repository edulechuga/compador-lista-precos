"""
Flask backend — Comparador de Listas de Preços
"""
import base64
import io
import logging
import sys
import traceback
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import Flask, jsonify, render_template, request

sys.path.insert(0, str(Path(__file__).parent / "execution"))
from comparar_listas import compare_excels  # noqa: E402

# ─────────────────────── Logging ───────────────────────────────
LOG_DIR = Path(__file__).parent / "logs"
LOG_DIR.mkdir(exist_ok=True)
LOG_FILE = LOG_DIR / "app.log"

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s  %(levelname)-8s  %(name)s — %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("comparador")
log.info("=" * 70)
log.info("Servidor iniciado — log em %s", LOG_FILE)
log.info("=" * 70)

# ───────────────────────── App ─────────────────────────────────
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB


def _get_sheet_names(file_bytes: bytes) -> list[str]:
    """Retorna as abas do Excel."""
    try:
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        return xl.sheet_names
    except Exception as exc:
        log.warning("Não foi possível listar abas: %s", exc)
        return []


def _resolve_sheet(raw_value, file_bytes: bytes):
    """
    Recebe o valor da aba selecionada pelo usuário (sempre string) e retorna
    o nome correto para passar ao pandas.

    Regra:
      - Se não veio nada → 0 (primeira aba)
      - Se é um nome válido de aba → usa o nome (string)
      - Se é um número puro E não existe como nome de aba → usa como índice (int)
    """
    if not raw_value:
        log.debug("Sheet não informada → usando índice 0")
        return 0

    sheet_names = _get_sheet_names(file_bytes)
    log.debug("Abas disponíveis: %s | Solicitado: %r", sheet_names, raw_value)

    # Prefere match exato por nome
    if raw_value in sheet_names:
        log.debug("Sheet resolvida por nome: %r", raw_value)
        return raw_value

    # Tenta match case-insensitive
    raw_lower = raw_value.strip().lower()
    for name in sheet_names:
        if str(name).strip().lower() == raw_lower:
            log.debug("Sheet resolvida por nome (case-insensitive): %r → %r", raw_value, name)
            return name

    # Só converte para int se o valor for numérico E estiver dentro do range
    if raw_value.isdigit():
        idx = int(raw_value)
        if idx < len(sheet_names):
            log.debug("Sheet resolvida por índice: %d", idx)
            return idx
        else:
            log.warning(
                "Índice %d fora do range (%d abas) — usando índice 0", idx, len(sheet_names)
            )
            return 0

    # Fallback
    log.warning("Sheet %r não encontrada, usando índice 0", raw_value)
    return 0


# ─────────────────────── Rotas ─────────────────────────────────
@app.route("/")
def index():
    log.info("GET / — página principal")
    return render_template("index.html")


@app.route("/abas", methods=["POST"])
def abas():
    """Retorna as abas disponíveis de um arquivo para o usuário selecionar."""
    try:
        f = request.files.get("arquivo")
        if not f:
            return jsonify({"error": "Arquivo não enviado."}), 400
        file_bytes = f.read()
        sheets = _get_sheet_names(file_bytes)
        log.info("POST /abas — arquivo=%r  abas=%s", f.filename, sheets)
        return jsonify({"sheets": sheets})
    except Exception as exc:
        log.error("POST /abas — erro: %s\n%s", exc, traceback.format_exc())
        return jsonify({"error": str(exc)}), 500


@app.route("/comparar", methods=["POST"])
def comparar():
    req_id = datetime.now().strftime("%H%M%S_%f")
    log.info("─" * 60)
    log.info("POST /comparar  [req=%s]", req_id)

    try:
        arquivo_antigo = request.files.get("arquivo_antigo")
        arquivo_novo   = request.files.get("arquivo_novo")

        if not arquivo_antigo or not arquivo_novo:
            log.warning("[%s] Arquivos não enviados", req_id)
            return jsonify({"error": "Ambos os arquivos são obrigatórios."}), 400

        log.info("[%s] ANTIGO=%r  NOVO=%r", req_id, arquivo_antigo.filename, arquivo_novo.filename)

        for f, nome in [(arquivo_antigo, "ANTIGO"), (arquivo_novo, "NOVO")]:
            if not f.filename.lower().endswith((".xlsx", ".xls")):
                msg = f"O arquivo {nome} deve ser um Excel (.xlsx ou .xls)."
                log.warning("[%s] %s", req_id, msg)
                return jsonify({"error": msg}), 400

        old_bytes = arquivo_antigo.read()
        new_bytes = arquivo_novo.read()
        log.info("[%s] Bytes lidos — ANTIGO=%d B  NOVO=%d B", req_id, len(old_bytes), len(new_bytes))

        # ── Resolução de abas (BUG FIX: não converte nome p/ int cegamente) ──
        raw_old = request.form.get("old_sheet", "")
        raw_new = request.form.get("new_sheet", "")
        log.info("[%s] Abas recebidas do form — old_sheet=%r  new_sheet=%r", req_id, raw_old, raw_new)

        old_sheet = _resolve_sheet(raw_old, old_bytes)
        new_sheet = _resolve_sheet(raw_new, new_bytes)
        log.info("[%s] Abas resolvidas — old=%r  new=%r", req_id, old_sheet, new_sheet)

        # ── Comparação ────────────────────────────────────────────────────────
        result = compare_excels(old_bytes, new_bytes,
                                old_sheet=old_sheet, new_sheet=new_sheet)

        log.info(
            "[%s] Resultado — idêntico=%s  novos=%d  removidos=%d  alterados=%d  iguais=%d  total_new=%d  total_old=%d",
            req_id,
            result["identical"],
            result["new_count"],
            result.get("removed_count", 0),
            result["modified_count"],
            result["unchanged_count"],
            result["total_new_file"],
            result.get("total_old_file", 0),
        )
        log.debug("[%s] Colunas detectadas — %s", req_id, result.get("col_info", {}))

        response = {
            "identical":       result["identical"],
            "new_count":       result["new_count"],
            "removed_count":   result.get("removed_count", 0),
            "modified_count":  result["modified_count"],
            "unchanged_count": result["unchanged_count"],
            "total_new_file":  result["total_new_file"],
            "total_old_file":  result.get("total_old_file", 0),
            "col_info": {
                k: {ck: cv for ck, cv in v.items()}
                for k, v in result.get("col_info", {}).items()
            },
        }

        if not result["identical"] and "excel_bytes" in result:
            response["excel_b64"]  = base64.b64encode(result["excel_bytes"]).decode("utf-8")
            response["excel_name"] = "comparacao_precos.xlsx"
            log.info("[%s] Excel gerado (%d B)", req_id, len(result["excel_bytes"]))

        return jsonify(response)

    except ValueError as exc:
        tb = traceback.format_exc()
        log.error("[%s] ValueError: %s\n%s", req_id, exc, tb)
        return jsonify({"error": str(exc), "detail": tb}), 400

    except Exception as exc:
        tb = traceback.format_exc()
        log.error("[%s] Erro inesperado: %s\n%s", req_id, exc, tb)
        return jsonify({"error": f"Erro inesperado: {exc}", "detail": tb}), 500


if __name__ == "__main__":
    log.info("Iniciando servidor Flask em modo debug na porta 5000 (0.0.0.0)")
    app.run(host="0.0.0.0", port=5000, debug=True)

