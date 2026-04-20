from __future__ import annotations

import os
import shutil
from datetime import datetime
from pathlib import Path
from typing import List, Dict

import pandas as pd

from logger_config import logger


class ArquivoBloqueadoError(Exception):
    """Exceção customizada para indicar que o arquivo Excel está bloqueado para escrita."""
    pass


class ExcelReadError(Exception):
    """Falha ao ler o Excel existente; gravação abortada para evitar perda de dados."""
    pass


class ExcelManager:
    def __init__(self, caminho_arquivo: str | Path) -> None:
        self.caminho_arquivo = str(Path(caminho_arquivo))

    def _backup_corrupt_file(self, exc: Exception) -> None:
        src = Path(self.caminho_arquivo)
        if not src.is_file():
            return
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        bak = src.with_name(f"{src.stem}_readerror_{stamp}{src.suffix}.bak")
        try:
            shutil.copy2(src, bak)
            logger.error(
                "Cópia de segurança do Excel criada em '%s' após erro de leitura: %s",
                bak,
                exc,
            )
        except OSError as copy_err:
            logger.error("Não foi possível copiar o Excel para backup: %s", copy_err)

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
                self._backup_corrupt_file(e)
                logger.error("Erro ao ler o arquivo Excel existente: %s", e)
                raise ExcelReadError(
                    f"Não foi possível ler '{self.caminho_arquivo}'. "
                    "Foi criada uma cópia .bak se possível; a gravação foi abortada."
                ) from e
            for c in colunas:
                if c not in df_existente.columns:
                    df_existente[c] = pd.NA
            df_existente = df_existente[colunas]
            df_pronto = pd.concat([df_existente, df_novos], ignore_index=True)
            id_col = 'ID Mensagem'
            mask = df_pronto[id_col].notna() & (df_pronto[id_col].astype(str).str.strip() != '')
            if mask.any():
                df_pronto = df_pronto.drop_duplicates(subset=[id_col], keep='last')

        try:
            df_pronto.to_excel(self.caminho_arquivo, index=False)
            logger.info("%s chamados salvos em '%s'.", len(lista_dados), self.caminho_arquivo)
        except PermissionError:
            logger.warning(
                "Arquivo '%s' está bloqueado para escrita. "
                "Feche o arquivo no Excel para continuar.",
                self.caminho_arquivo,
            )
            raise ArquivoBloqueadoError(
                f"O arquivo '{self.caminho_arquivo}' está aberto em outro programa. "
                "Não foi possível salvar os dados."
            )
        except Exception as e:
            logger.error("Erro inesperado ao salvar o Excel: %s", e)
            raise
