import os, re, time, gspread, threading
from tray_icon import *
from datetime import datetime
from google.oauth2.service_account import Credentials
from config import PLANILHAS, CNPJ_MVA, CNPJ_EH, INTERVALO, DOWNLOAD_DIR, GOOGLE_CREDENTIALS_SHEETS
from gmail_service import getGmailService, buscarMessagesEnviados, baixar_anexos_de_mensagem
from reporter import escreverRelatorio, registrarEvento, consolidarRelatorioTMP
from xml_parser import extrairDadosXML
from sheets_writer import atualizarPlanilha
from gmail_service import marcar_mensagem_com_label
from logger_config import logger, cor_ciano, reset


stop_event = threading.Event()  # usado para parar o loop com segurança
running = False # indica se o loop principal está ativo

def escolher_planilha_por_cnpj_e_ano(cnpj: str, ano: str):
    if cnpj == CNPJ_MVA:
        return PLANILHAS["MVA"].get(ano)
    if cnpj == CNPJ_EH:
        return PLANILHAS["EH"].get(ano)
    return None

def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

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

        dados_xmls = []
        boletos = []

        # 🔁 Processa todos os anexos baixados
        for arquivo in arquivos:
            nome_arquivo = os.path.basename(arquivo)

            try:
                # =============================
                # 📄 XML → extrai dados
                # =============================
                if arquivo.lower().endswith(".xml"):
                    try:
                        dados = extrairDadosXML(arquivo)
                        # 🔍 Ignora vendas à vista
                        nat_op = dados.get("naturezaOperacao", "").strip().upper()
                        dest = dados.get("destinatario", "")
                        if ( "VISTA" in nat_op or "VENDA A VISTA" in nat_op):
                            # Checa se a mensagem ja foi processada no relatorio atual:
                            if dados.get('nf') not in consolidarRelatorioTMP(): 
                                escreverRelatorio(f"{_now()} - 💰 NF {dados.get('nf')} ignorada (venda à vista).")
                                continue
                            else: logger.info(f"{cor_ciano}NF {dados['nf']} já registrada no relatório, não duplicando a mensagem de ignorada.{reset}") 
                            continue
                        if ( CNPJ_MVA.replace(".", "").replace("/", "").replace("-", "") in dest or
                             CNPJ_EH.replace(".", "").replace("/", "").replace("-", "") in dest ):
                            logger.info(f"[DEBUG IGNORE RESULT] NF {dados['nf']} ignorada (destinatário é o nosso: {dest})")
                            escreverRelatorio(f"{_now()} - 💰 NF {dados.get('nf')} ignorada (destinatário é o nosso).")
                            continue
                        if not dados:
                            motivo = dados.get("motivo_ignoracao", "Desconhecido") if isinstance(dados, dict) else "Desconhecido"
                            logger.info(f"Ignorado XML (motivo: {motivo}).")
                            escreverRelatorio(f"{_now()} - ⚠️ XML {nome_arquivo} ignorado (motivo: {motivo})")
                            continue

                        dados_xmls.append(dados)

                    except Exception as e:
                        escreverRelatorio(f"{_now()} - ❌ Erro extraindo XML {nome_arquivo}: {e}")
                        logger.exception("Erro extraindo XML %s: %s", arquivo, e)

                # =============================
                # 📑 PDF → tenta identificar boleto
                # =============================
                elif arquivo.lower().endswith(".pdf"): # mudar pra elif se o bloco de cima for realmente necessário
                    nome_upper = nome_arquivo.upper()

                    # 🔍 Trata nomes parecidos com BOLETO (erros comuns tipo BOLTO, BOLETA, BOLETT, etc)
                    padrao_boleto = r"[_\s-]?(BLT|BOLET[OA]?|BOLTO|BOLETOO|BOLETT?)"

                    if re.search(padrao_boleto, nome_upper):
                        match = re.findall(r"([0-9]{2,}-?[0-9]+)", nome_upper)
                        if match:
                            num_boleto = match[-1]
                            boletos.append(num_boleto)
                            logger.info("🔢 Boleto identificado no nome: %s (BLT %s)", nome_arquivo, num_boleto)
                        else:
                            logger.info("Nenhum número de boleto encontrado no nome: %s", nome_arquivo)
                    elif arquivo.lower().endswith(".pdf"):
                        nome_upper = nome_arquivo.upper()

                        # 🔍 Palavras que indicam boleto (considera erros comuns)
                        padrao_boleto = r"\b(BOLET[OA]?|BOLTO|BOLETOO|BOLETT?|BLT)\b"

                        # Só tenta identificar número se o nome realmente tiver algo próximo de "boleto"
                        if re.search(padrao_boleto, nome_upper):
                            match = re.findall(r"([0-9]{2,}-?[0-9]+)", nome_upper)
                            if match:
                                num_boleto = match[-1]
                                boletos.append(num_boleto)
                                logger.info("🔢 Boleto identificado no nome: %s (BLT %s)", nome_arquivo, num_boleto)
                            else:
                                logger.info("📎 Possível boleto sem número identificado: %s", nome_arquivo)
                        else:
                            logger.info("📄 PDF ignorado (não parece boleto): %s", nome_arquivo)

                else:
                    logger.info("Arquivo não identificado como boleto: %s", nome_arquivo)

            finally:
                # 🧹 Remove sempre o anexo local (independente do tipo)
                try:
                    os.remove(arquivo)
                    logger.debug(f"🧹 Anexo removido: {arquivo}")
                except FileNotFoundError:
                    pass
                except Exception as e:
                    logger.warning(f"⚠️ Falha ao remover {arquivo}: {e}")

        # =============================
        # 🏷️ Marca o e-mail como processado
        # =============================
        try:
            marcar_mensagem_com_label(service, msg_id)
            logger.info("🏷️ E-mail %s marcado com 'XML Processado Botana'", msg_id)
        except Exception as e:
            logger.exception("Falha ao aplicar rótulo: %s", e)
            
        # ⚠️ Nenhum XML → pula este e-mail
        if not dados_xmls:
            logger.info("Nenhum XML válido encontrado neste e-mail.")
            continue

        # =============================
        # 🧾 Atualiza planilhas
        # =============================
        for dados_xml in dados_xmls:
            cnpj_emit = dados_xml.get("cnpjEmitente")
            ano = dados_xml.get("anoVencimento")
            planilha_id = escolher_planilha_por_cnpj_e_ano(cnpj_emit, ano)

            if not planilha_id:
                logger.warning("CNPJ %s ou ano %s sem planilha configurada.", cnpj_emit, ano)
                continue

            # Itera sobre todas as parcelas — MAPEAMENTO correto de boletos → parcelas
            parcelas = dados_xml.get("parcelas", [])
            n_parcelas = len(parcelas)
            n_boletos = len(boletos)

            # monta lista de boletos por parcela (mesmo tamanho de parcelas)
            if n_parcelas == 0:
                continue  # nada a fazer

            if n_boletos == 0:
                boletos_map = [None] * n_parcelas
            else:
                # Se tiver igual, mapeia 1:1; se menor, preenche em ordem; se maior, usa só os primeiros N
                boletos_map = [boletos[i] if i < n_boletos else None for i in range(n_parcelas)]
                if n_boletos > n_parcelas:
                    logger.info("⚠️ Mais boletos (%d) que parcelas (%d). Sobraram: %s", n_boletos, n_parcelas, boletos[n_parcelas:])

            # Agora processa 1 vez por parcela, usando o boleto mapeado (ou None)
            for idx, parcela in enumerate(parcelas):
                num_boleto = boletos_map[idx]
                dados_parcela = dados_xml.copy()
                dados_parcela.update({
                    "vencimento": parcela["vencimento"],
                    "numParcela": parcela["numParcela"],
                    "valorParcela": parcela["valor"],
                    "boleto": num_boleto  # adiciona campo explícito (opcional)
                })

                # Ajusta descrição com o boleto mapeado (se houver)
                if num_boleto:
                    dados_parcela["descricao"] = f"{dados_parcela['destinatario']} BLT {num_boleto} (Bot)"
                else:
                    if "18471209000107" in cnpj_emit.upper():
                        dados_parcela["descricao"] = f"{dados_parcela['destinatario']} DEP BR (Bot)"
                    else:
                        dados_parcela["descricao"] = f"{dados_parcela['destinatario']} DEP CX (Bot)"

                # Tenta atualizar planilha com retry
                for tentativa in range(5):
                    try:
                        creds = Credentials.from_service_account_file(
                            GOOGLE_CREDENTIALS_SHEETS,
                            scopes=["https://www.googleapis.com/auth/spreadsheets"]
                        )
                        gc = gspread.authorize(creds)

                        if not hasattr(processar_emails_enviados, "_cache"):
                            processar_emails_enviados._cache = {}

                        cache = processar_emails_enviados._cache

                        if planilha_id not in cache:
                            cache[planilha_id] = gc.open_by_key(planilha_id)

                        planilha = cache[planilha_id]
                        atualizarPlanilha(planilha, dados_parcela, gc)
                        total_processados += 1
                        break         
                    except gspread.exceptions.APIError as e:
                        if "429" in str(e):
                            from sheets_writer import apiCooldown
                            apiCooldown()
                            continue
                        else:
                            logger.exception("Erro ao atualizar planilha: %s", e)
                            break
                    except Exception as e:
                        logger.exception("Falha inesperada ao atualizar planilha: %s", e)
                        break
    logger.info("Ciclo finalizado. Total processado: %d", total_processados)

def main():
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
    time.sleep(1)
    sys.exit(0)

# =========================
# EXECUÇÃO PRINCIPAL
# =========================
if __name__ == "__main__":
    # Passa callbacks para o tray (para permitir controle)
    run_tray(on_quit_callback=on_quit, start_callback=iniciar_verificacao)
