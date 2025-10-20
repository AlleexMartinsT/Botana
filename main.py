import os, re
import time
import gspread
import threading
from tray_icon import *
from google.oauth2.service_account import Credentials
from config import PLANILHAS, CNPJ_MVA, CNPJ_EH, INTERVALO, DOWNLOAD_DIR, GOOGLE_CREDENTIALS_SHEETS
from gmail_service import getGmailService, buscarMessagesEnviados, baixar_anexos_de_mensagem
from xml_parser import extrairDadosXML
from sheets_writer import atualizarPlanilha
import colorlog, logging

stop_event = threading.Event()  # usado para parar o loop com segurança
running = False # indica se o loop principal está ativo

handler = colorlog.StreamHandler()
handler.setFormatter(colorlog.ColoredFormatter(
    "%(log_color)s%(asctime)s [%(levelname)s] %(message)s",
    log_colors={
        "INFO": "green",
        "WARNING": "yellow",
        "ERROR": "red",
        "DEBUG": "blue"
    }
))

logging.basicConfig(level=logging.INFO, handlers=[handler])
logger = logging.getLogger("bot.main")

def escolher_planilha_por_cnpj_e_ano(cnpj: str, ano: str):
    if cnpj == CNPJ_MVA:
        return PLANILHAS["MVA"].get(ano)
    if cnpj == CNPJ_EH:
        return PLANILHAS["EH"].get(ano)
    return None

def processar_emails_enviados():
    service = getGmailService()
    msgs = buscarMessagesEnviados(service, max_results=100)
    if not msgs:
        logger.info("Nenhuma mensagem enviada com XML encontrada.")
        return

    total_processados = 0
    for m in msgs:
        msg_id = m.get("id")
        logger.info("📧 Abrindo mensagem ID: %s", msg_id)
        arquivos = baixar_anexos_de_mensagem(service, msg_id)
        if not arquivos:
            logger.info("Nenhum anexo salvo para mensagem %s", msg_id)
            continue

        dados_xml = None
        num_boleto = None
        sucesso_na_mensagem = False

        for arquivo in arquivos:
            # 1️⃣ XML
            if arquivo.lower().endswith(".xml"):
                try:
                    dados = extrairDadosXML(arquivo)
                    if not dados:
                        logger.info("Ignorado (destinatário é o nosso).")
                        os.remove(arquivo)
                        continue
                    dados_xml = dados
                    os.remove(arquivo)
                except Exception as e:
                    logger.exception("Erro extraindo XML %s: %s", arquivo, e)
                    os.remove(arquivo)
                    continue

            # 2️⃣ PDF → só extrai número BLT ou BOLETO do nome do arquivo
            elif arquivo.lower().endswith(".pdf"):
                nome_arquivo = os.path.basename(arquivo)
                
                # Aceita separadores comuns antes de BLT ou BOLETO
                if re.search(r"[_\s-]?(BLT|BOLETO)", nome_arquivo.upper()):
                    match = re.findall(r"([0-9]{2,}-?[0-9]+)", nome_arquivo)
                    if match:
                        num_boleto = match[-1]  # pega o último número, que geralmente é o boleto
                        logger.info("🔢 Boleto identificado no nome: %s (BLT %s)", nome_arquivo, num_boleto)

                    else:
                        logger.info("Nenhum número de boleto encontrado no nome: %s", nome_arquivo)
                else:
                    logger.info("Arquivo não identificado como boleto: %s", nome_arquivo)

                try:
                    os.remove(arquivo)
                except OSError:
                    pass


        # 3️⃣ Se não houver XML, pula
        if not dados_xml:
            logger.info("Nenhum XML válido encontrado neste e-mail.")
            continue

        # Define planilha
        cnpj_emit = dados_xml.get("cnpjEmitente")
        ano = dados_xml.get("anoVencimento")
        planilha_id = escolher_planilha_por_cnpj_e_ano(cnpj_emit, ano)
        if not planilha_id:
            logger.warning("CNPJ %s ou ano %s sem planilha configurada.", cnpj_emit, ano)
            continue

        # Descrição final: nome do destinatário + BLT (se tiver)
        if num_boleto:
            dados_xml["descricao"] = f"{dados_xml['destinatario']} BLT {num_boleto} (Bot)"
        else:
            if "18471209000107" in cnpj_emit.upper():
                dados_xml["descricao"] = f"{dados_xml['destinatario']} DEP BR (Bot)"
            else:
                dados_xml["descricao"] = f"{dados_xml['destinatario']} DEP CX (Bot)"

                # Atualiza planilha com tratamento de limite da API
        for tentativa in range(5):
            try:
                creds = Credentials.from_service_account_file(
                    GOOGLE_CREDENTIALS_SHEETS,
                    scopes=["https://www.googleapis.com/auth/spreadsheets"]
                )
                gc = gspread.authorize(creds)

                # Cache simples de planilhas já abertas
                if not hasattr(processar_emails_enviados, "_cache"):
                    processar_emails_enviados._cache = {}

                cache = processar_emails_enviados._cache
                if planilha_id not in cache:
                    cache[planilha_id] = gc.open_by_key(planilha_id)

                planilha = cache[planilha_id]

                atualizarPlanilha(planilha, dados_xml)
                total_processados += 1
                break  # saiu com sucesso ✅

            except gspread.exceptions.APIError as e:
                if "429" in str(e):
                    logger.warning("⚠️ Limite da API atingido (tentativa %d/5). Aguardando 30 segundos...", tentativa + 1)
                    from sheets_writer import apiCooldown
                    apiCooldown()
                    continue  # tenta novamente
                else:
                    logger.exception("Erro ao atualizar planilha: %s", e)
                    break

            except Exception as e:
                logger.exception("Falha inesperada ao atualizar planilha: %s", e)
                break

    logger.info("Ciclo finalizado. Total processado: %d", total_processados)

def main():
    logger.info("🚀 Inicializando bot de envios (monitorando 'Sent')...")
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    while True:
        try:
            processar_emails_enviados()
        except Exception as e:
            logger.exception("Erro no ciclo principal: %s", e)
        logger.info("⏳ Aguardando %d segundos para próxima verificação...", INTERVALO)
        time.sleep(INTERVALO)

def iniciar_verificacao():
    """Inicia o loop principal em thread separada (chamado pelo tray)."""
    global running
    if not running:
        stop_event.clear()
        t = threading.Thread(target=main, daemon=True)
        t.start()
        print("[Main] Loop principal iniciado.")
    else:
        print("[Main] Loop já está em execução.")


def parar_verificacao():
    """Interrompe o loop principal."""
    global running
    if running:
        print("[Main] Parando loop principal...")
        stop_event.set()
        running = False
    else:
        print("[Main] Nenhum loop ativo para encerrar.")


def on_quit():
    """Chamado quando o usuário clica em 'Sair' no tray."""
    parar_verificacao()
    print("[Main] Encerrando Finance Bot...")
    time.sleep(1)
    sys.exit(0)

# =========================
# EXECUÇÃO PRINCIPAL
# =========================
if __name__ == "__main__":
    # Passa callbacks para o tray (para permitir controle)
    run_tray(on_quit_callback=on_quit, start_callback=iniciar_verificacao)
