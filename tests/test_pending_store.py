from datetime import datetime

from pending_store import merge_by_message_id, max_data_chegada


def test_merge_prefere_novos_quando_mesmo_id():
    pend = [{"ID Mensagem": "A", "x": 1}]
    nov = [{"ID Mensagem": "A", "x": 2}]
    m = merge_by_message_id(pend, nov)
    assert len(m) == 1
    assert m[0]["x"] == 2


def test_max_data_chegada():
    rows = [
        {"Data Chegada": datetime(2026, 1, 1, 10, 0)},
        {"Data Chegada": datetime(2026, 2, 1, 10, 0)},
    ]
    assert max_data_chegada(rows) == datetime(2026, 2, 1, 10, 0)


def test_sem_id_nao_dedupa():
    a = [{"ID Mensagem": "", "k": 1}]
    b = [{"ID Mensagem": "", "k": 2}]
    m = merge_by_message_id(a, b)
    assert len(m) == 2
