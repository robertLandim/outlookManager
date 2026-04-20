"""Testes automatizados do motor de SLA (horário comercial)."""

from datetime import datetime

import pytest

from sla_engine import calcular_prazo_sla


def test_segunda_09h_24h_sla():
    chegada = datetime(2026, 4, 13, 9, 0)
    fim = calcular_prazo_sla(chegada, 24)
    # 24h úteis com 9h/dia ≈ 2 dias + 6h a partir da segunda 09:00 → quarta 15:00
    assert fim.weekday() == 2
    assert fim.hour == 15
    assert fim.minute == 0


def test_sexta_17h_4h_sla_pula_fds():
    chegada = datetime(2026, 4, 17, 17, 0)
    fim = calcular_prazo_sla(chegada, 4)
    assert fim.weekday() == 0  # segunda
    assert fim.hour == 12


def test_sabado_14h_8h_sla_comeca_segunda():
    chegada = datetime(2026, 4, 18, 14, 0)
    fim = calcular_prazo_sla(chegada, 8)
    assert fim.weekday() == 0
    assert fim.hour == 17


def test_exatamente_inicio_expediente():
    chegada = datetime(2026, 4, 13, 9, 0, 0)
    fim = calcular_prazo_sla(chegada, 1)
    assert fim.hour == 10


@pytest.mark.parametrize(
    "hora,minuto",
    [(18, 0), (17, 59)],
)
def test_borda_fim_expediente(hora, minuto):
    chegada = datetime(2026, 4, 13, hora, minuto)
    fim = calcular_prazo_sla(chegada, 1)
    assert fim.date() >= chegada.date()
