"""Notificações toast no Windows com debounce."""

from __future__ import annotations

import threading
import time
from pathlib import Path

from logger_config import logger

# winotify exige app_id sem caracteres especiais (ex.: travessão unicode) — quebram o PowerShell/XML.
_APP_ID = "OutlookManager.Ouvidoria"

_last_blocked_ts: float = 0.0
_lock = threading.Lock()


def notify_excel_blocked(
    caminho_planilha: str,
    debounce_minutes: float,
    enabled: bool,
) -> None:
    if not enabled or debounce_minutes <= 0:
        return
    global _last_blocked_ts
    now = time.monotonic()
    debounce_sec = debounce_minutes * 60.0
    with _lock:
        if now - _last_blocked_ts < debounce_sec:
            return
        _last_blocked_ts = now

    nome = Path(caminho_planilha).name
    try:
        from winotify import Notification

        toast = Notification(
            app_id=_APP_ID,
            title="Ouvidoria - planilha em uso",
            msg=(
                "Novos chamados aguardando gravação. Feche o Excel em "
                f"'{nome}' para o robô salvar as atualizações."
            ),
        )
        toast.show()
    except Exception as e:
        logger.warning("Falha ao exibir notificacao Windows (planilha bloqueada): %s", e)
