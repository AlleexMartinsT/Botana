import time
import re
import gspread
from google.oauth2.service_account import Credentials

import config

print("Inicializando script de correção do Google Sheets...")

# EN: Setup authentication client for Google Sheets
# BR: Eu configuro a autenticação para o Google Sheets através de conta de serviço local.
def getSheetsClient():
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        credentialsGoogle = Credentials.from_service_account_file(
            config.GOOGLE_CREDENTIALS_SHEETS, scopes=SCOPES
        )
        clienteGoogle = gspread.authorize(credentialsGoogle)
        print("Sucesso na autenticação com as credenciais!")
        return clienteGoogle
    except Exception as erro:
        print(f"Erro ao autenticar: {erro}")
        return None

# EN: Process a single spreadsheet based on its ID
# BR: Eu processo uma única planilha procurada a partir do seu ID.
def processarPlanilha(clienteSheets, idPlanilha, nomeMatriz):
    if not idPlanilha:
        return
    try:
        planilhaAtual = clienteSheets.open_by_key(idPlanilha)
        print(f"\nVerificando planilha: {planilhaAtual.title} ({nomeMatriz})")
    except Exception as erroOpen:
        print(f"Erro ao abrir {nomeMatriz} com id {idPlanilha}: {erroOpen}")
        return
    
    for abaAtual in planilhaAtual.worksheets():
        print(f" -> Lendo aba: {abaAtual.title}")
        try:
            # EN: Sleep to avoid API limit bottlenecks
            # BR: Eu coloco o robô para dormir um instante para evitar esgotar a cota da API (Too Many Requests).
            time.sleep(1)
            registrosAba = abaAtual.get_all_values()
            
            # EN: Expected structure is A (index 0) and B (index 1) - where B is the "Description"
            # BR: A estrutura esperada coloca a "Descrição" na coluna B (índice 1). get_all_values() retorna listas completas.
            quantidadeAlteracoes = 0
            
            for indiceLinha, linhaCompleta in enumerate(registrosAba):
                numeroLinha = indiceLinha + 1 # 1-based no Sheets
                
                if len(linhaCompleta) > 1:
                    descricaoAtual = linhaCompleta[1]
                    
                    # EN: Find specific "BLT" long numbering formats inside descriptions
                    # BR: Eu busco encontrar formatos "BLT" com longa numeração dentro das descrições para refaze-las.
                    def atualizarRegex(matchEncontrado):
                        textoPrefixo = matchEncontrado.group(1)
                        numeroLongo = matchEncontrado.group(2)
                        
                        # EN: Finding trailing zeros chunk to cleanly separate ticket sequence 
                        # BR: Eu busco um longo agrupamento de 0s seguidos (ao menos 4) e puxo o texto posterior sem zeros
                        limpezaMatch = re.search(r'0{4,}([1-9][0-9]*(-[0-9A-Za-z]+)?)$', numeroLongo)
                        
                        if limpezaMatch:
                            return textoPrefixo + limpezaMatch.group(1)
                        return matchEncontrado.group(0)
                        
                    novaDescricao = re.sub(r'(BLT\s+)([0-9-]+)', atualizarRegex, descricaoAtual)
                    
                    if novaDescricao != descricaoAtual:
                        print(f"      Corrigindo Linha {numeroLinha}:")
                        print(f"         De: {descricaoAtual}")
                        print(f"         Aa: {novaDescricao}")
                        
                        # EN: Col 2 corresponds to 'B' which holds Description 
                        # BR: A coluna 2 é a letra B, responsável por abrigar a descrição.
                        abaAtual.update_cell(numeroLinha, 2, novaDescricao)
                        quantidadeAlteracoes += 1
                        time.sleep(1.5) # EN: Extra precaution delay. BR: Atraso extra de precaução proativa.
                        
            print(f"      Terminou aba {abaAtual.title}. Células corrigidas na rodada: {quantidadeAlteracoes}")
            
        except Exception as erroAba:
            print(f"Erro ao processar aba {abaAtual.title}: {erroAba}")

def iniciar_assistente_em_background():
    import threading
    def tarefa_assistente():
        print("[Assistente Botana] Iniciando varredura retroativa...")
        clienteAtivo = getSheetsClient()
        if clienteAtivo:
            for nomeCompanhia, dictAnos in config.PLANILHAS.items():
                for anoPlanilha, idGoogle in dictAnos.items():
                    if idGoogle:
                        processarPlanilha(clienteAtivo, idGoogle, f"{nomeCompanhia}-{anoPlanilha}")
        print("[Assistente Botana] Varredura finalizada.")
        
    threadLimpeza = threading.Thread(target=tarefa_assistente, daemon=True)
    threadLimpeza.start()
    return True

# EN: Main standalone launcher inside Botana hub
# BR: Eu declaro a função de inicialização se operada fora de outro código.
if __name__ == "__main__":
    iniciar_assistente_em_background()
