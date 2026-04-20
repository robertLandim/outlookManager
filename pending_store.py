"""Persistência de chamados pendentes quando o Excel está bloqueado."""

from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any, List

from logger_config import logger


def _serialize_row(row: dict[str, Any]) -> dict[str, Any]:
    out = {}
    for k, v in row.items():
        if isinstance(v, datetime):
            out[k] = v.isoformat()
        else:
            out[k] = v
    return out


def _deserialize_row(row: dict[str, Any]) -> dict[str, Any]:
    out = dict(row)
    dc = out.get("Data Chegada")
    if isinstance(dc, str):
        try:
            out["Data Chegada"] = datetime.fromisoformat(dc.replace("Z", "+00:00"))
        except ValueError:
            pass
    return out


def load_pending(path: Path) -> List[dict[str, Any]]:
    if not path.is_file():
        return []
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        rows = data.get("rows", [])
        return [_deserialize_row(r) for r in rows]
    except (json.JSONDecodeError, OSError) as e:
        logger.warning("Não foi possível carregar pendências em %s: %s", path, e)
        return []


def save_pending(path: Path, rows: List[dict[str, Any]]) -> None:
    payload = {"rows": [_serialize_row(r) for r in rows]}
    tmp = path.with_name(path.name + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def clear_pending(path: Path) -> None:
    if path.is_file():
        try:
            path.unlink()
        except OSError as e:
            logger.warning("Não foi possível remover pendências %s: %s", path, e)


def merge_by_message_id(
    pending: List[dict[str, Any]],
    novos: List[dict[str, Any]],
) -> List[dict[str, Any]]:
    """Une pendentes e novos; para o mesmo ID Mensagem, prevalece o dado mais recente (lista novos)."""
    seen: set[str] = set()
    merged: List[dict[str, Any]] = []
    for row in novos + pending:
        mid = str(row.get("ID Mensagem", "") or "").strip()
        if mid:
            if mid in seen:
                continue
            seen.add(mid)
        merged.append(row)
    return merged


def max_data_chegada(rows: List[dict[str, Any]]) -> datetime | None:
    best: datetime | None = None
    for row in rows:
        dc = row.get("Data Chegada")
        if isinstance(dc, datetime):
            if best is None or dc > best:
                best = dc
    return best
