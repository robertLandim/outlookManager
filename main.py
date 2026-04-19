import threading
import time
import pystray
from pystray import MenuItem as item
from PIL import Image, ImageDraw

from outlook_api import OutlookReader
from sla_engine import calcular_prazo_sla
from excel_manager import ExcelManager, ArquivoBloqueadoError
from logger_config import logger

# Configurações globais
NOME_PASTA_OUTLOOK = "Ouvidoria_Teste"
HORAS_SLA = 24  # SLA de 24 horas úteis por chamado
INTERVALO_MINUTOS = 10  # Tempo entre ciclos, em minutos

def rotina_sincronizacao():
    logger.info("Iniciando rotina de sincronização do sistema Ouvidoria...")
    outlook = OutlookReader()
    excel = ExcelManager()

    try:
        emails = outlook.buscar_novos_emails(nome_pasta=NOME_PASTA_OUTLOOK)
    except Exception as e:
        logger.error(f"Falha ao buscar emails: {e}")
        return

    if not emails:
        logger.info("Nenhum novo e-mail encontrado na rotina. Aguardando próximo ciclo.")
        return

    dados_para_excel = []
    for mail in emails:
        try:
            prazo_sla = calcular_prazo_sla(mail['data_recebimento'], HORAS_SLA)
        except Exception as e:
            logger.error(f"Erro ao calcular SLA para o e-mail '{mail}': {e}")
            prazo_sla = ""  # Armazena string vazia, ou pode 'continue' (option)
        dados_para_excel.append({
            "Assunto": mail.get("assunto", ""),
            "Remetente": mail.get("remetente", ""),
            "Data Chegada": mail.get("data_recebimento", ""),
            "Prazo SLA": prazo_sla,
            "ID Mensagem": mail.get("entry_id", "") or "",
        })
    try:
        excel.salvar_chamados(dados_para_excel)
        # Atualiza o cursor (data última processada)
        data_ultimo = emails[-1]['data_recebimento']
        outlook.atualizar_cursor(data_ultimo)
        logger.info(f"Sincronização finalizada. {len(dados_para_excel)} chamados processados e cursor atualizado.")
        print(f"Sincronização finalizada. {len(dados_para_excel)} chamados processados e cursor atualizado.")
    except ArquivoBloqueadoError as e:
        logger.warning(f"Não foi possível salvar porque o arquivo está aberto: {e}")
        # Não avança o cursor, para tentar novamente no próximo ciclo
        return
    except Exception as e:
        logger.error(f"Erro inesperado ao salvar chamados ou atualizar cursor: {e}")
        print(f"Erro inesperado ao salvar chamados ou atualizar cursor: {e}")
        return

def criar_icone():
    # Ícone: quadrado azul com círculo branco no centro (32x32)
    img = Image.new('RGBA', (32, 32), "blue")
    draw = ImageDraw.Draw(img)
    draw.ellipse((8, 8, 24, 24), fill="white")
    return img

def thread_da_sincronizacao():
    logger.info("Thread de sincronização automática iniciada (background).")
    while True:
        try:
            rotina_sincronizacao()
        except Exception as e:
            logger.error(f"Erro na rotina sincronizada: {e}")
        time.sleep(INTERVALO_MINUTOS * 1)

def acao_sincronizar_manual(icon, item):
    logger.info("Sincronização manual acionada pelo usuário.")
    thread = threading.Thread(target=rotina_sincronizacao)
    thread.daemon = True
    thread.start()

def acao_sair(icon, item):
    logger.info("Sistema de Ouvidoria finalizando pelo menu do usuário.")
    icon.stop()

def main():
    # Inicia thread de background
    t = threading.Thread(target=thread_da_sincronizacao, daemon=True)
    t.start()

    # Cria e configura pystray Icon
    menu = (
        item("Sincronizar Agora", acao_sincronizar_manual),
        item("Sair", acao_sair)
    )
    icone = pystray.Icon("OuvidoriaBot", criar_icone(), "Outlook Manager Robot - Ativo", menu=menu)
    logger.info("Sistema da Ouvidoria carregado. O ícone foi iniciado na bandeja do sistema.")
    icone.run()

if __name__ == "__main__":
    main()