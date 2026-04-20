"""Carrega configuração e diretório base da aplicação."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import tomllib

APP_ROOT = Path(__file__).resolve().parent

DEFAULT_CONFIG: dict[str, Any] = {
    "outlook": {"folder_name": "Ouvidoria_Teste"},
    "sla": {"hours": 24},
    "sync": {"interval_minutes": 10},
    "paths": {
        "spreadsheet": "Controle_Ouvidoria.xlsx",
        "state_file": "state.json",
        "pending_file": "pending_chamados.json",
    },
    "notifications": {
        "enabled": True,
        "blocked_debounce_minutes": 10,
    },
}


def _deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    out = dict(base)
    for k, v in override.items():
        if k in out and isinstance(out[k], dict) and isinstance(v, dict):
            out[k] = _deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def load_config(config_path: Path | None = None) -> dict[str, Any]:
    path = config_path or (APP_ROOT / "config.toml")
    merged = dict(DEFAULT_CONFIG)
    if path.is_file():
        with open(path, "rb") as f:
            loaded = tomllib.load(f)
        merged = _deep_merge(merged, loaded)
    return merged


def abs_path_under_root(relative: str) -> Path:
    p = Path(relative)
    if p.is_absolute():
        return p
    return APP_ROOT / p
