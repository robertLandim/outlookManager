from unittest.mock import patch

import pandas as pd
import pytest

from excel_manager import (
    ArquivoBloqueadoError,
    ExcelManager,
    ExcelReadError,
)


def test_permission_error_raises_blocked(tmp_path):
    xlsx = tmp_path / "t.xlsx"
    df = pd.DataFrame(
        [{"Assunto": "a", "Remetente": "b", "Data Chegada": 1, "Prazo SLA": 2, "ID Mensagem": "x"}]
    )
    df.to_excel(xlsx, index=False)

    em = ExcelManager(xlsx)
    dados = [
        {
            "Assunto": "n",
            "Remetente": "r",
            "Data Chegada": pd.Timestamp("2026-01-01"),
            "Prazo SLA": "",
            "ID Mensagem": "id1",
        }
    ]

    with patch.object(pd.DataFrame, "to_excel", side_effect=PermissionError("locked")):
        with pytest.raises(ArquivoBloqueadoError):
            em.salvar_chamados(dados)


def test_read_error_raises_without_clobber(tmp_path):
    xlsx = tmp_path / "bad.xlsx"
    xlsx.write_bytes(b"not an xlsx")

    em = ExcelManager(xlsx)
    dados = [
        {
            "Assunto": "n",
            "Remetente": "r",
            "Data Chegada": pd.Timestamp("2026-01-01"),
            "Prazo SLA": "",
            "ID Mensagem": "id1",
        }
    ]

    with pytest.raises(ExcelReadError):
        em.salvar_chamados(dados)


def test_novo_arquivo_grava(tmp_path):
    xlsx = tmp_path / "new.xlsx"
    em = ExcelManager(xlsx)
    dados = [
        {
            "Assunto": "s",
            "Remetente": "r",
            "Data Chegada": pd.Timestamp("2026-01-01"),
            "Prazo SLA": "",
            "ID Mensagem": "id1",
        }
    ]
    em.salvar_chamados(dados)
    assert xlsx.is_file()
    df = pd.read_excel(xlsx)
    assert len(df) == 1
    assert df.iloc[0]["Assunto"] == "s"
