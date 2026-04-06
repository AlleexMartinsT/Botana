import time
import re
import threading
import gspread
from google.oauth2.service_account import Credentials

import config

print("Inicializando script de correção do Google Sheets...")

# EN: In-memory log buffer shared across threads for real-time progress tracking.
# BR: Eu mantenho um buffer de log em memória compartilhado entre threads para acompanhamento em tempo real.
_logCorrecao = []
_logLock = threading.Lock()
_correcaoAtiva = False
_correcaoLock = threading.Lock()

def obterLog(desdeIndice: int = 0) -> dict:
    """EN: I return the current log entries starting from a given index.
    BR: Eu retorno as entradas de log a partir de um indice especificado."""
    with _logLock:
        entradas = list(_logCorrecao[desdeIndice:])
    with _correcaoLock:
        ativo = _correcaoAtiva
    return {"entries": entradas, "ativo": ativo, "total": len(_logCorrecao)}

def _adicionarLog(tipo: str, mensagem: str):
    """EN: I append a log entry with timestamp and type.
    BR: Eu adiciono uma entrada de log com timestamp e tipo."""
    entrada = {
        "ts": time.strftime("%H:%M:%S"),
        "tipo": tipo,
        "msg": mensagem,
    }
    with _logLock:
        _logCorrecao.append(entrada)

def _limparLog():
    """EN: I clear all log entries for a fresh run.
    BR: Eu limpo todas as entradas de log para uma nova execução."""
    with _logLock:
        _logCorrecao.clear()


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
        _adicionarLog("info", "Autenticação com Google Sheets bem-sucedida.")
        return clienteGoogle
    except Exception as erro:
        _adicionarLog("erro", f"Erro ao autenticar: {erro}")
        return None

# EN: Process a single spreadsheet based on its ID, optionally filtering by tab name.
# BR: Eu processo uma única planilha procurada a partir do seu ID, opcionalmente filtrando por aba.
def processarPlanilha(clienteSheets, idPlanilha, nomeMatriz, filtroAba: str = ""):
    if not idPlanilha:
        return
    try:
        planilhaAtual = clienteSheets.open_by_key(idPlanilha)
        _adicionarLog("info", f"Verificando planilha: {planilhaAtual.title} ({nomeMatriz})")
    except Exception as erroOpen:
        _adicionarLog("erro", f"Erro ao abrir {nomeMatriz} com id {idPlanilha}: {erroOpen}")
        return
    
    for abaAtual in planilhaAtual.worksheets():
        # EN: If a tab filter is specified, I only process matching tabs.
        # BR: Se um filtro de aba foi especificado, eu so processo as abas correspondentes.
        if filtroAba and filtroAba.lower() not in abaAtual.title.lower():
            continue

        _adicionarLog("info", f"  Lendo aba: {abaAtual.title}")
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
                    
                    # EN: I fix BLT numbering issues in descriptions.
                    # BR: Eu corrijo problemas de numeração de BLT nas descrições.
                    # Rule: All BLT numbers should be 5 digits starting with 1 (10001-19999 range).
                    def corrigirBlt(matchEncontrado):
                        textoPrefixo = matchEncontrado.group(1)
                        numeroBruto = matchEncontrado.group(2)
                        
                        # EN: I separate the numeric part from any dash-suffix like "-8" or "-53".
                        # BR: Eu separo a parte numérica de qualquer sufixo com traço como "-8" ou "-53".
                        partesMatch = re.match(r'^(\d+)(-.+)?$', numeroBruto)
                        if not partesMatch:
                            return matchEncontrado.group(0)
                        
                        numPuro = partesMatch.group(1)
                        sufixo = partesMatch.group(2) or ""
                        
                        # EN: Already correct: 5 digits starting with 1 → don't touch.
                        # BR: Já está correto: 5 dígitos começando com 1 → não mexo.
                        if len(numPuro) == 5 and numPuro[0] == '1':
                            return matchEncontrado.group(0)
                        
                        # EN: 4 digits starting with 0 → prepend 1 (the leading 1 was truncated).
                        # BR: 4 dígitos começando com 0 → adiciono 1 na frente (o 1 inicial foi perdido).
                        # Ex: 0005→10005, 0097→10097, 0240→10240
                        if len(numPuro) == 4 and numPuro[0] == '0':
                            return textoPrefixo + "1" + numPuro + sufixo
                        
                        # EN: 1 digit → was corrupted, pad to 5 digits (1000X).
                        # BR: 1 dígito → foi corrompido, restauro para 5 dígitos (1000X).
                        # Ex: 5→10005, 1→10001, 8→10008
                        if len(numPuro) == 1:
                            return textoPrefixo + "1000" + numPuro + sufixo
                        
                        # EN: 2 digits → was corrupted, pad to 5 digits (100XX).
                        # BR: 2 dígitos → foi corrompido, restauro para 5 dígitos (100XX).
                        if len(numPuro) == 2:
                            return textoPrefixo + "100" + numPuro + sufixo
                        
                        # EN: 3 digits → was corrupted, pad to 5 digits (10XXX).
                        # BR: 3 dígitos → foi corrompido, restauro para 5 dígitos (10XXX).
                        if len(numPuro) == 3:
                            return textoPrefixo + "10" + numPuro + sufixo
                        
                        # EN: Anything else (6+ digits) → don't touch for safety.
                        # BR: Qualquer outra coisa (6+ dígitos) → não mexo por segurança.
                        return matchEncontrado.group(0)
                        
                    novaDescricao = re.sub(r'(BLT[\s-]+)(\d+(?:-[\w]+)?)', corrigirBlt, descricaoAtual)
                    
                    if novaDescricao != descricaoAtual:
                        _adicionarLog("correcao", f"  Linha {numeroLinha} | De: {descricaoAtual} → Para: {novaDescricao}")
                        
                        # EN: Col 2 corresponds to 'B' which holds Description 
                        # BR: A coluna 2 é a letra B, responsável por abrigar a descrição.
                        abaAtual.update_cell(numeroLinha, 2, novaDescricao)
                        quantidadeAlteracoes += 1
                        time.sleep(1.5) # EN: Extra precaution delay. BR: Atraso extra de precaução proativa.
                        
            _adicionarLog("info", f"  Aba {abaAtual.title} finalizada. Correções: {quantidadeAlteracoes}")
            
        except Exception as erroAba:
            _adicionarLog("erro", f"Erro na aba {abaAtual.title}: {erroAba}")

def atualizarHistoricosRetroativos():
    import os, re
    from config import RELATORIO_DIR
    if os.path.exists(RELATORIO_DIR):
        _adicionarLog("info", "Atualizando histórico (TXT)...")
        qt = 0
        for f in os.listdir(RELATORIO_DIR):
            if f.endswith('.txt'):
                p = os.path.join(RELATORIO_DIR, f)
                with open(p, 'r', encoding='utf-8') as file:
                    lines = file.readlines()
                changed = False
                for i, line in enumerate(lines):
                    if 'HIST_JSON' in line:
                        def corrigirBltTexto(m):
                            pref = m.group(1)
                            raw = m.group(2)
                            pm = re.match(r'^(\d+)(-.+)?$', raw)
                            if not pm: return m.group(0)
                            n = pm.group(1)
                            s = pm.group(2) or ""
                            if len(n) == 5 and n[0] == '1': return m.group(0)
                            if len(n) == 4 and n[0] == '0': return pref + "1" + n + s
                            if len(n) == 1: return pref + "1000" + n + s
                            if len(n) == 2: return pref + "100" + n + s
                            if len(n) == 3: return pref + "10" + n + s
                            return m.group(0)
                        
                        nl = re.sub(r'(BLT[\s-]+)(\d+(?:-[\w]+)?)', corrigirBltTexto, line)
                        if nl != line:
                            lines[i] = nl
                            changed = True
                            qt += 1
                if changed:
                    with open(p, 'w', encoding='utf-8') as file:
                        file.writelines(lines)
        _adicionarLog("info", f"Histórico atualizado. Linhas corrigidas: {qt}")

def iniciar_assistente_em_background(empresa: str = "todos", filtroAba: str = ""):
    """EN: I start the correction assistant in a background thread with optional filters.
    BR: Eu inicio o assistente de correção em thread separada com filtros opcionais."""
    global _correcaoAtiva
    
    with _correcaoLock:
        if _correcaoAtiva:
            return False
        _correcaoAtiva = True
    
    _limparLog()
    _adicionarLog("info", f"Iniciando varredura retroativa... (Empresa: {empresa}, Aba: {filtroAba or 'todas'})")
    
    def tarefa_assistente():
        global _correcaoAtiva
        try:
            clienteAtivo = getSheetsClient()
            if clienteAtivo:
                for nomeCompanhia, dictAnos in config.PLANILHAS.items():
                    # EN: I filter by empresa if specified.
                    # BR: Eu filtro por empresa se especificada.
                    if empresa and empresa != "todos" and nomeCompanhia != empresa:
                        continue
                    for anoPlanilha, idGoogle in dictAnos.items():
                        if idGoogle:
                            processarPlanilha(clienteAtivo, idGoogle, f"{nomeCompanhia}-{anoPlanilha}", filtroAba)
            atualizarHistoricosRetroativos()
            _adicionarLog("info", "Varredura finalizada com sucesso!")
        except Exception as erroGeral:
            _adicionarLog("erro", f"Erro geral na varredura: {erroGeral}")
        finally:
            with _correcaoLock:
                _correcaoAtiva = False
        
    threadLimpeza = threading.Thread(target=tarefa_assistente, daemon=True)
    threadLimpeza.start()
    return True

# EN: Main standalone launcher inside Botana hub
# BR: Eu declaro a função de inicialização se operada fora de outro código.
if __name__ == "__main__":
    iniciar_assistente_em_background()
