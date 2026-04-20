"""
Robô de bandeja: lê e-mails no Outlook (pasta configurável), calcula SLA e grava em Excel.

Operação: manter o Outlook Desktop aberto e logado; fechar a planilha no Excel quando o robô
precisar gravar (logs em logs/outlook_manager.log ao lado deste projeto).
"""

from __future__ import annotations

import threading

import pystray
from pystray import MenuItem as item
from PIL import Image, ImageDraw

from config import abs_path_under_root, load_config
from excel_manager import ArquivoBloqueadoError, ExcelManager, ExcelReadError
from logger_config import logger
from notifications_win import notify_excel_blocked
from outlook_api import OutlookReader, write_watermark
from pending_store import (
    clear_pending,
    load_pending,
    max_data_chegada,
    merge_by_message_id,
    save_pending,
)
from sla_engine import calcular_prazo_sla

# Evento para “Sincronizar agora” (mesma thread do worker COM)
_manual_sync: threading.Event | None = None


def _run_sync_cycle(cfg: dict) -> None:
    paths = cfg["paths"]
    state_path = abs_path_under_root(paths["state_file"])
    spreadsheet_path = abs_path_under_root(paths["spreadsheet"])
    pending_path = abs_path_under_root(paths["pending_file"])
    nome_pasta = cfg["outlook"]["folder_name"]
    horas_sla = int(cfg["sla"]["hours"])
    notif = cfg["notifications"]

    logger.info("Iniciando ciclo de sincronização da Ouvidoria...")
    pending_rows = load_pending(pending_path)

    emails: list = []
    try:
        _outlook = OutlookReader(str(state_path))
        emails = _outlook.buscar_novos_emails(nome_pasta=nome_pasta)
    except RuntimeError as e:
        logger.error("Falha ao conectar ao Outlook: %s", e)
        if not pending_rows:
            return

    dados_novos = []
    for mail in emails:
        try:
            prazo_sla = calcular_prazo_sla(mail["data_recebimento"], horas_sla)
        except Exception as ex:
            logger.error("Erro ao calcular SLA para o e-mail %s: %s", mail, ex)
            prazo_sla = ""
        dados_novos.append(
            {
                "Assunto": mail.get("assunto", ""),
                "Remetente": mail.get("remetente", ""),
                "Data Chegada": mail.get("data_recebimento", ""),
                "Prazo SLA": prazo_sla,
                "ID Mensagem": mail.get("entry_id", "") or "",
            }
        )

    if not emails and not pending_rows:
        logger.info("Nenhum e-mail novo e nada pendente. Aguardando próximo ciclo.")
        return

    to_save = merge_by_message_id(pending_rows, dados_novos)
    if not to_save:
        logger.info("Nada a gravar após mesclar pendências.")
        return

    excel = ExcelManager(spreadsheet_path)
    try:
        excel.salvar_chamados(to_save)
    except ArquivoBloqueadoError as e:
        logger.warning("Não foi possível salvar — arquivo em uso: %s", e)
        notify_excel_blocked(
            str(spreadsheet_path),
            float(notif.get("blocked_debounce_minutes", 10)),
            bool(notif.get("enabled", True)),
        )
        save_pending(pending_path, to_save)
        return
    except ExcelReadError as e:
        logger.error("%s", e)
        return
    except Exception as e:
        logger.error("Erro ao salvar chamados: %s", e)
        return

    clear_pending(pending_path)
    dt_cursor = max_data_chegada(to_save)
    if dt_cursor is None:
        logger.warning("Sem datas em 'Data Chegada'; watermark não atualizado.")
        return

    write_watermark(str(state_path), dt_cursor)
    logger.info(
        "Sincronização finalizada: %s linha(s) gravada(s); watermark atualizado.",
        len(to_save),
    )


def _sync_worker(cfg: dict) -> None:
    """Um único thread: COM inicializado aqui (Outlook via pywin32)."""
    import pythoncom

    pythoncom.CoInitialize()
    try:
        interval_min = float(cfg["sync"]["interval_minutes"])
        interval_sec = max(interval_min * 60.0, 1.0)
        while True:
            try:
                _run_sync_cycle(cfg)
            except Exception as e:
                logger.exception("Erro na rotina de sincronização: %s", e)
            ev = _manual_sync
            if ev is None:
                break
            triggered = ev.wait(timeout=interval_sec)
            if triggered:
                ev.clear()
    finally:
        pythoncom.CoUninitialize()


def criar_icone_ok() -> Image.Image:
    img = Image.new("RGBA", (32, 32), "blue")
    draw = ImageDraw.Draw(img)
    draw.ellipse((8, 8, 24, 24), fill="white")
    return img


def acao_sincronizar_manual(icon: pystray.Icon, _item: pystray.MenuItem) -> None:
    logger.info("Sincronização manual acionada pelo menu.")
    if _manual_sync is not None:
        _manual_sync.set()


def acao_sair(icon: pystray.Icon, _item: pystray.MenuItem) -> None:
    logger.info("Encerrando pelo menu da bandeja.")
    icon.stop()


def main() -> None:
    global _manual_sync

    cfg = load_config()
    _manual_sync = threading.Event()

    worker = threading.Thread(target=_sync_worker, args=(cfg,), daemon=True)
    worker.start()

    menu = (
        item("Sincronizar agora", acao_sincronizar_manual),
        item("Sair", acao_sair),
    )
    icone = pystray.Icon(
        "OuvidoriaBot",
        criar_icone_ok(),
        "Outlook Manager — Ouvidoria (ativo)",
        menu=menu,
    )
    logger.info(
        "Sistema carregado (intervalo %s min). Ícone na bandeja.",
        cfg["sync"]["interval_minutes"],
    )
    icone.run()


if __name__ == "__main__":
    main()
