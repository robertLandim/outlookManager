from logger_config import logger
from datetime import datetime, timedelta
import os
import json
from typing import List, Dict, Any, Union

class OutlookReader:
    """
    Classe robusta para automação do Outlook Desktop via COM Automation.
    Garante leitura incremental de e-mails utilizando controle de Watermark baseado em data/hora persistida em arquivo local.
    """

    def __init__(self):
        """
        Inicializa a conexão com a API MAPI do Outlook. 
        Caso o Outlook não esteja aberto, uma exceção amigável é lançada.
        """
        try:
            import win32com.client
            self.win32com = win32com
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
        except Exception as e:
            logger.error(f"Falha ao conectar na API do Outlook: {e}")
            raise RuntimeError(
                "Não foi possível conectar ao Outlook. Certifique-se de que o Outlook Desktop está aberto e logado."
            ) from e
        self.state_file = "state.json"

    def _ler_watermark(self) -> datetime:
        """
        Lê a última data/hora salva em state.json.
        Caso o arquivo não exista ou esteja corrompido, retorna a data de 24 horas atrás 
        como fallback de segurança para não processar a caixa inteira.
        """
        if os.path.exists(self.state_file):
            with open(self.state_file, "r", encoding="utf-8") as f:
                try:
                    data = json.load(f)
                    ultima_leitura = data.get("ultima_leitura")
                    if ultima_leitura:
                        return datetime.strptime(ultima_leitura, "%Y-%m-%d %H:%M:%S")
                except (json.JSONDecodeError, ValueError, KeyError):
                    pass # Se o arquivo estiver corrompido, ignora e vai para o fallback abaixo
        
        # Fallback de segurança: Exatamente 24 horas atrás
        logger.warning("state.json não encontrado. Assumindo leitura das últimas 24 horas.")
        return datetime.now() - timedelta(days=1)

    def _salvar_watermark(self, nova_data: datetime) -> None:
        """
        Salva a data/hora mais recente lida no arquivo state.json.
        """
        with open(self.state_file, "w", encoding="utf-8") as f:
            json.dump({"ultima_leitura": nova_data.strftime("%Y-%m-%d %H:%M:%S")}, f, ensure_ascii=False, indent=2)

    def buscar_novos_emails(self, nome_pasta: str = "Inbox") -> List[Dict[str, Any]]:
        """
        Busca e-mails recebidos após a última leitura registrada no state.json.
        """
        pasta_ref = None
        for folder in self.namespace.GetDefaultFolder(6).Parent.Folders:
            if folder.Name.lower() == nome_pasta.lower():
                pasta_ref = folder
                break
                
        if pasta_ref is None:
            pasta_ref = self.namespace.GetDefaultFolder(6)  # Fallback para Inbox padrão

        ultima_lida = self._ler_watermark()
        filtro = ""
        if ultima_lida:
            # O Outlook prefere o formato de data local ou dd/mm/yyyy HH:MM para filtros COM
            dt_str = ultima_lida.strftime("%d/%m/%Y %H:%M:%S")
            filtro = f"[ReceivedTime] > '{dt_str}'"
            
        emails = pasta_ref.Items
        if filtro:
            emails = emails.Restrict(filtro)
            
        # False = Ascending (Do mais antigo para o mais novo)
        emails.Sort("[ReceivedTime]", False) 

        resultados: List[Dict[str, Any]] = []
        
        for item in emails:
            # Class == 43 garante que o item é um E-mail real (ignora convites de calendário, alertas, etc)
            if item.Class == 43: 
                # Conversão segura do pywintypes.datetime para datetime do Python
                raw_date = item.ReceivedTime
                dt_recebimento = datetime(
                    raw_date.year, raw_date.month, raw_date.day,
                    raw_date.hour, raw_date.minute, raw_date.second
                )
                
                resultados.append({
                    "assunto": item.Subject,
                    "remetente": item.SenderName,  # SenderName é mais seguro que Sender.Name
                    "data_recebimento": dt_recebimento,
                    "entry_id": getattr(item, "EntryID", None) or "",
                    "objeto_email": item,
                })
                
        return resultados

    def atualizar_cursor(self, nova_data: Union[datetime, str]) -> None:
        """
        Atualiza o arquivo state.json com a data do e-mail mais recente lido.

        Parâmetro:
            nova_data: datetime ou string "YYYY-MM-DD HH:MM:SS".
        """
        if isinstance(nova_data, datetime):
            dt = nova_data
        else:
            dt = datetime.strptime(nova_data, "%Y-%m-%d %H:%M:%S")
        self._salvar_watermark(dt)
