from datetime import datetime
from sla_engine import calcular_prazo_sla

def rodar_testes_de_sla():
    print("Iniciando bateria de testes do Motor de SLA...\n")
    
    # Cenário 1: E-mail chega na Segunda de manhã (SLA de 24 horas úteis)
    # Como são 9h úteis por dia, 24 horas = 2 dias completos + 6 horas.
    chegada_1 = datetime(2026, 4, 13, 9, 0) # Segunda-feira, 09:00
    resultado_1 = calcular_prazo_sla(chegada_1, 24)
    print(f"Cenário 1 (Chegada Segunda 09:00 + 24h úteis):")
    print(f"-> Vencimento Calculado: {resultado_1}")
    print("-" * 40)
    
    # Cenário 2: E-mail chega na Sexta-feira no fim do expediente
    chegada_2 = datetime(2026, 4, 17, 17, 0) # Sexta-feira, 17:00
    resultado_2 = calcular_prazo_sla(chegada_2, 4) # SLA curto de 4h
    print(f"Cenário 2 (Chegada Sexta 17:00 + 4h úteis):")
    print(f"-> Vencimento Calculado: {resultado_2} (Deve pular para Segunda-feira!)")
    print("-" * 40)

    # Cenário 3: E-mail chega no Sábado (Fora do expediente)
    chegada_3 = datetime(2026, 4, 18, 14, 0) # Sábado, 14:00
    resultado_3 = calcular_prazo_sla(chegada_3, 8) # SLA de 8h
    print(f"Cenário 3 (Chegada Sábado 14:00 + 8h úteis):")
    print(f"-> Vencimento Calculado: {resultado_3} (Deve ignorar o fds e contar a partir de Segunda!)")

if __name__ == "__main__":
    rodar_testes_de_sla()