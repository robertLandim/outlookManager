import os
from excel_manager import ExcelManager, ArquivoBloqueadoError
from logger_config import logger

def rodar_teste_excel():
    # Usamos um arquivo com nome diferente para não sujar o de produção depois
    arquivo_teste = "Controle_Ouvidoria_Teste.xlsx"
    
    # Limpa o ambiente apagando o teste anterior, se existir
    if os.path.exists(arquivo_teste):
        try:
            os.remove(arquivo_teste)
        except PermissionError:
            print(f"⚠️ Feche o arquivo '{arquivo_teste}' antes de começar o teste!")
            return

    gerenciador = ExcelManager(arquivo_teste)
    
    # Dados simulados (Mock Data)
    dados_lote_1 = [{
        "Assunto": "[Ouvidoria] Reclamação de Atraso",
        "Remetente": "cliente_a@teste.com",
        "Data Chegada": "2026-04-10 10:00:00",
        "Prazo SLA": "2026-04-13 10:00:00"
    }]
    
    dados_lote_2 = [{
        "Assunto": "[Ouvidoria] Dúvida sobre Faturamento",
        "Remetente": "financeiro@teste.com",
        "Data Chegada": "2026-04-10 14:00:00",
        "Prazo SLA": "2026-04-13 14:00:00"
    }]

    print("Iniciando bateria de testes do Excel...\n")
    
    try:
        # Cenário 1: Criação do Zero
        print("Passo 1: Criando arquivo e inserindo o Lote 1...")
        gerenciador.salvar_chamados(dados_lote_1)
        print("✅ Sucesso! Arquivo criado na raiz do projeto.")
        
        # Cenário 2: Adicionando Linhas (Append)
        print("\nPasso 2: Abrindo arquivo existente e adicionando o Lote 2...")
        gerenciador.salvar_chamados(dados_lote_2)
        print("✅ Sucesso! Lote 2 adicionado sem apagar os dados anteriores.")
        
        # Cenário 3: Simulação de Usuário Interrompendo (O Teste de Fogo)
        print("\n" + "="*55)
        print("🚨 HORA DA SIMULAÇÃO DE ESTRESSE 🚨")
        print(f"1. Abra a pasta do projeto e dê um duplo clique em '{arquivo_teste}'.")
        print("2. Deixe o arquivo aberto no seu Excel na tela.")
        input("3. APÓS abrir o Excel, clique neste terminal e aperte [ENTER]...")
        print("="*55 + "\n")
        
        print("Passo 3: Tentando salvar com a gestora mexendo na planilha...")
        
        # Tentamos injetar mais dados com o arquivo aberto
        gerenciador.salvar_chamados(dados_lote_1) 
        
        # Se o código chegar na linha abaixo, significa que a nossa exceção não funcionou
        print("❌ Falha: O código conseguiu salvar (ou o arquivo não estava aberto no aplicativo Excel).")

    except ArquivoBloqueadoError as e:
        # Este é o resultado esperado para o Passo 3!
        print("✅ SUCESSO ABSOLUTO! O sistema detectou o bloqueio perfeitamente.")
        print(f"Mensagem da Exceção Capturada: {e}")
        print("O seu robô está blindado contra a interferência do usuário!")
        
    except Exception as e:
        print(f"❌ Erro inesperado e não tratado: {e}")

if __name__ == "__main__":
    rodar_teste_excel()