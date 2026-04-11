from outlook_api import OutlookReader
from logger_config import logger

def rodar_cenario_de_teste():
    logger.info("Iniciando cenário de teste isolado do Outlook...")
    print("Iniciando cenário de teste isolado...\n")
    
    try:
        # 1. Instancia o nosso leitor
        leitor = OutlookReader()
        
        # 2. Define a pasta que criamos no Outlook para o teste
        nome_da_pasta = "Ouvidoria_Teste"
        print(f"Buscando e-mails na pasta: '{nome_da_pasta}'...")
        
        # 3. Puxa os e-mails
        emails_encontrados = leitor.buscar_novos_emails(nome_pasta=nome_da_pasta)
        
        # 4. Exibe os resultados
        if not emails_encontrados:
            print("\nNenhum e-mail novo encontrado desde a última leitura.")
            logger.info("Teste concluído: 0 e-mails encontrados.")
            return

        print(f"\n✅ SUCESSO! Encontrados {len(emails_encontrados)} e-mails.")
        print("-" * 50)
        
        for i, email in enumerate(emails_encontrados, 1):
            print(f"E-mail {i}:")
            print(f"  Assunto: {email['assunto']}")
            print(f"  De:      {email['remetente']}")
            print(f"  Data:    {email['data_recebimento'].strftime('%d/%m/%Y %H:%M:%S')}")
            print("-" * 50)
            
        # 5. Simulando o avanço do Watermark (state.json)
        ultimo_email = emails_encontrados[-1] 
        print(f"\nSimulando gravação do state.json com a data: {ultimo_email['data_recebimento']}")
        
        leitor.atualizar_cursor(ultimo_email['data_recebimento'])
        
        print("Arquivo state.json atualizado com sucesso!")
        logger.info(f"Teste concluído com sucesso. Watermark atualizado para {ultimo_email['data_recebimento']}")
            
    except Exception as e:
        logger.error(f"Erro crítico no teste: {e}")
        print(f"❌ Erro crítico no teste: {e}")

if __name__ == "__main__":
    # Lembre-se de deletar o state.json antes da primeira execução 
    # para testar o fallback de 24 horas!
    rodar_cenario_de_teste()