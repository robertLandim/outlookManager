from typing import List, Dict
import os
import pandas as pd
from logger_config import logger

class ArquivoBloqueadoError(Exception):
    """Exceção customizada para indicar que o arquivo Excel está bloqueado para escrita."""
    pass

class ExcelManager:
    def __init__(self, caminho_arquivo: str = "Controle_Ouvidoria.xlsx") -> None:
        self.caminho_arquivo = caminho_arquivo

    def salvar_chamados(self, lista_dados: List[Dict]) -> None:
        """
        Salva novos chamados em um arquivo Excel, criando ou atualizando conforme o caso.

        Parâmetros:
            lista_dados (List[Dict]): Lista de dicionários com as chaves:
                'Assunto', 'Remetente', 'Data Chegada', 'Prazo SLA'
        """
        colunas = ['Assunto', 'Remetente', 'Data Chegada', 'Prazo SLA', 'ID Mensagem']

        if not lista_dados:
            logger.warning("Nenhum dado recebido para salvar no Excel.")
            return

        df_novos = pd.DataFrame(lista_dados, columns=colunas)

        if not os.path.exists(self.caminho_arquivo):
            df_pronto = df_novos
        else:
            try:
                df_existente = pd.read_excel(self.caminho_arquivo)
            except Exception as e:
                logger.error(f"Erro ao ler o arquivo Excel existente: {e}")
                df_existente = pd.DataFrame(columns=colunas)
            for c in colunas:
                if c not in df_existente.columns:
                    df_existente[c] = pd.NA
            df_existente = df_existente[colunas]
            df_pronto = pd.concat([df_existente, df_novos], ignore_index=True)
            id_col = 'ID Mensagem'
            mask = df_pronto[id_col].notna() & (df_pronto[id_col].astype(str).str.strip() != '')
            if mask.any():
                df_pronto = df_pronto.drop_duplicates(subset=[id_col], keep='first')

        try:
            df_pronto.to_excel(self.caminho_arquivo, index=False)
            logger.info(f"{len(lista_dados)} chamados salvos em '{self.caminho_arquivo}'.")
        except PermissionError:
            logger.warning(
                f"Arquivo '{self.caminho_arquivo}' está bloqueado para escrita. "
                "Feche o arquivo no Excel para continuar."
            )
            raise ArquivoBloqueadoError(
                f"O arquivo '{self.caminho_arquivo}' está aberto em outro programa. "
                "Não foi possível salvar os dados."
            )
        except Exception as e:
            logger.error(f"Erro inesperado ao salvar o Excel: {e}")
            raise