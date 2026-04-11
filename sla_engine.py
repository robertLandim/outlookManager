from datetime import datetime, timedelta
from logger_config import logger

def calcular_prazo_sla(data_chegada: datetime, horas_sla: int) -> datetime:
    """
    Calcula o prazo de SLA com base em uma janela de horário comercial (09:00 às 18:00),
    ignorando completamente finais de semana, e ajustando o início da contagem caso fora do expediente.

    Parâmetros:
        data_chegada (datetime): Momento em que o item chegou ou foi recebido.
        horas_sla (int): Número total de horas de SLA a serem somadas dentro do expediente.

    Retorna:
        datetime: Data e hora exatas do vencimento do SLA, respeitando janelas de trabalho.

    Regras de Negócio:
        - Contagem apenas no expediente (09:00 - 18:00).
        - Finais de semana (sábado/domingo) são totalmente desconsiderados.
        - Antes das 09:00 ou depois das 18:00, o prazo só começa a contar às 09:00 do próximo dia útil.
        - Progride o cálculo dia a dia, deduzindo apenas horas viáveis de cada janela.
    """
    # Constantes de expediente
    HORA_INICIO = 9
    HORA_FIM = 18
    HORAS_POR_DIA = HORA_FIM - HORA_INICIO

    logger.debug(f"Início do cálculo do SLA: chegada={data_chegada}, horas_sla={horas_sla}")

    # Função utilitária para verificar se é final de semana
    def eh_fim_de_semana(data: datetime) -> bool:
        return data.weekday() >= 5  # 5: sábado, 6: domingo

    # Ajusta data_chegada para o próximo horário útil, se necessário
    atual = data_chegada

    # Fora do expediente: ou antes de abrir ou depois de fechar (ou fim de semana)
    if eh_fim_de_semana(atual):
        logger.info("Data de chegada caiu em fim de semana. Avançando para próximo dia útil às 09:00.")
        # Avançar para a próxima segunda-feira às 09:00
        while eh_fim_de_semana(atual):
            atual = atual + timedelta(days=1)
        atual = atual.replace(hour=HORA_INICIO, minute=0, second=0, microsecond=0)
    elif atual.hour < HORA_INICIO or (atual.hour == HORA_INICIO and atual.minute == 0 and atual.second == 0):
        logger.info("Data de chegada antes do expediente. Ajustando para 09:00 do mesmo dia.")
        atual = atual.replace(hour=HORA_INICIO, minute=0, second=0, microsecond=0)
    elif atual.hour >= HORA_FIM:
        logger.info("Data de chegada após o expediente. Avançando para próximo dia útil às 09:00.")
        # Avançar para o próximo dia
        atual = atual + timedelta(days=1)
        # Caso caia em final de semana, avançar mais
        while eh_fim_de_semana(atual):
            atual = atual + timedelta(days=1)
        atual = atual.replace(hour=HORA_INICIO, minute=0, second=0, microsecond=0)
    else:
        # Caso esteja no expediente, checa se deve ajustar para o começo do horário útil (caso horário igual a 18:00)
        if atual.hour == HORA_FIM and (atual.minute > 0 or atual.second > 0):
            atual = atual + timedelta(days=1)
            while eh_fim_de_semana(atual):
                atual = atual + timedelta(days=1)
            atual = atual.replace(hour=HORA_INICIO, minute=0, second=0, microsecond=0)

    horas_restantes = horas_sla

    # Loop incremental: adiciona horas respeitando o expediente e pulando finais de semana
    while horas_restantes > 0:
        # Determina fim do expediente do dia atual
        fim_do_expediente = atual.replace(hour=HORA_FIM, minute=0, second=0, microsecond=0)
        # Horas disponíveis hoje
        delta_hoje = (fim_do_expediente - atual).total_seconds() / 3600
        # Se não está em expediente válido (ou horário incossistente), pula para o próximo útil
        if delta_hoje <= 0:
            # Avança para o próximo dia útil às 09:00
            atual = atual + timedelta(days=1)
            while eh_fim_de_semana(atual):
                atual = atual + timedelta(days=1)
            atual = atual.replace(hour=HORA_INICIO, minute=0, second=0, microsecond=0)
            continue

        horas_a_somar = min(horas_restantes, delta_hoje)
        logger.debug(
            f"Somando {horas_a_somar}h em {atual.date()} (restantes: {horas_restantes}h, disponíveis hoje: {delta_hoje}h)"
        )

        atual = atual + timedelta(hours=horas_a_somar)
        horas_restantes -= horas_a_somar

        # Se ainda faltar horas, cai no próximo ciclo (próximo dia útil às 09:00)
        if horas_restantes > 0:
            atual = atual.replace(hour=HORA_INICIO, minute=0, second=0, microsecond=0)
            atual = atual + timedelta(days=1)
            while eh_fim_de_semana(atual):
                atual = atual + timedelta(days=1)
            atual = atual.replace(hour=HORA_INICIO, minute=0, second=0, microsecond=0)

    logger.info(f"SLA calculado: vencimento em {atual.strftime('%Y-%m-%d %H:%M:%S')}")
    return atual