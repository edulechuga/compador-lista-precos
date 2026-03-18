"""
Utility functions shared across execution scripts.
- Load environment variables from .env
- Common helpers: logging, file I/O, etc.
"""
import os
import json
import logging
from pathlib import Path
from dotenv import load_dotenv

# ─────────────────────────────────────────────
# Paths
# ─────────────────────────────────────────────
ROOT = Path(__file__).resolve().parent.parent
TMP_DIR = ROOT / ".tmp"
TMP_DIR.mkdir(exist_ok=True)

# ─────────────────────────────────────────────
# Environment
# ─────────────────────────────────────────────
load_dotenv(ROOT / ".env")


def get_env(key: str, required: bool = True) -> str:
    """Fetch an env var; raise clearly if missing and required."""
    value = os.getenv(key)
    if required and not value:
        raise EnvironmentError(
            f"Missing required environment variable: {key}\n"
            f"Add it to your .env file (see .env.example)."
        )
    return value or ""


# ─────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────
def get_logger(name: str) -> logging.Logger:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    return logging.getLogger(name)


# ─────────────────────────────────────────────
# File helpers
# ─────────────────────────────────────────────
def save_json(data: dict | list, filename: str) -> Path:
    """Save data as JSON in .tmp/."""
    path = TMP_DIR / filename
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def load_json(filename: str) -> dict | list:
    """Load JSON from .tmp/."""
    path = TMP_DIR / filename
    if not path.exists():
        raise FileNotFoundError(f".tmp/{filename} not found.")
    return json.loads(path.read_text(encoding="utf-8"))
