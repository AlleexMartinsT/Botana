import argparse
import json
import os, re, time, gspread, threading, sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse
from datetime import datetime
from google.oauth2.service_account import Credentials
from config import PLANILHAS, CNPJ_MVA, CNPJ_EH, INTERVALO, DOWNLOAD_DIR, GOOGLE_CREDENTIALS_SHEETS
from gmail_service import getGmailService, buscarMessagesEnviados, baixar_anexos_de_mensagem
from reporter import escreverRelatorio, consolidarRelatorioTMP
from xml_parser import extrairDadosXML
from sheets_writer import atualizarPlanilha
from gmail_service import marcar_mensagem_com_label
from logger_config import logger, cor_ciano, reset
try:
    from tray_icon import run_tray
except Exception:
    run_tray = None

# -----------------------
# FILTROS PARA DEBUG / ANÁLISE ISOLADA
# -----------------------
# Defina manualmente aqui (string) ou via variável de ambiente:
# Ex.: set SKIP_UNTIL_NF=12345       (Windows CMD)

# Se quiser que o script ignore tudo até achar a NF X, defina SKIP_UNTIL_NF
SKIP_UNTIL_NF = os.environ.get("SKIP_UNTIL_NF") or None  # ex: "12345"

# Se quiser processar somente uma NF específica (ignorar todas as outras), defina NF_ALVO
NF_ALVO = os.environ.get("NF_ALVO") or None  # ex: "12345"

# Se NF_ALVO for usado e quiser que o script pare após processar essa NF, coloque True
STOP_AFTER_NF = os.environ.get("STOP_AFTER_NF", "False").lower() in ("1", "true", "yes")
# -----------------------

stop_event = threading.Event()  # usado para parar o loop com segurança
running = False # indica se o loop principal está ativo
last_status = {"ok": True, "message": "Aguardando", "at": None}

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
            # --- FILTRAGEM POR NF (para debug/análise isolada) ---
            nf_num = str(dados_xml.get("nf", "")).strip()

            # NF_ALVO: processa somente essa NF (ignora as outras)
            if NF_ALVO:
                if nf_num != str(NF_ALVO):
                    logger.info(f"🔎 Pulando NF {nf_num} (NF_ALVO ativo: {NF_ALVO})")
                    continue
                else:
                    logger.info(f"✅ NF_ALVO encontrada: {nf_num}")

            # SKIP_UNTIL_NF: ignora tudo até encontrar essa NF; quando encontrada, passa a processar normalmente
            if SKIP_UNTIL_NF:
                # usa atributo da função para manter estado entre ciclos enquanto o processo está vivo
                if not hasattr(processar_emails_enviados, "_skip_reached"):
                    processar_emails_enviados._skip_reached = False

                if not processar_emails_enviados._skip_reached:
                    if nf_num == str(SKIP_UNTIL_NF):
                        processar_emails_enviados._skip_reached = True
                        logger.info(f"🎯 SKIP_UNTIL_NF: NF {nf_num} encontrada — a partir daqui será processada.")
                    else:
                        logger.info(f"⏭ SKIP_UNTIL_NF ativo, pulando NF {nf_num}")
                        continue

            # Se chegou até aqui, a NF será processada normalmente.
            # Se NF_ALVO + STOP_AFTER_NF: após processar, se encerra o loop/principal para análise isolada.

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
                        # Se NF_ALVO + STOP_AFTER_NF -> encerra o processo principal para análise isolada.
                        if NF_ALVO and STOP_AFTER_NF:
                            logger.info(f"🏁 NF_ALVO {NF_ALVO} processada. STOP_AFTER_NF=True -> encerrando execução.")
                            # força saída limpa do loop principal retornando da função
                            return
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

def main_loop():
    global running, last_status
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    running = True
    logger.info("[Botana] Loop iniciado")
    while not stop_event.is_set():
        try:
            processar_emails_enviados()
            last_status = {"ok": True, "message": "Ciclo executado com sucesso", "at": datetime.now().isoformat()}
        except Exception as e:
            logger.exception("Erro no ciclo principal: %s", e)
            last_status = {"ok": False, "message": f"Erro no ciclo: {e}", "at": datetime.now().isoformat()}
        if stop_event.wait(INTERVALO):
            break
    running = False
    logger.info("[Botana] Loop finalizado")


def executar_um_ciclo():
    global last_status
    try:
        processar_emails_enviados()
        last_status = {"ok": True, "message": "Execução manual concluída", "at": datetime.now().isoformat()}
        return True, "Execução manual concluída"
    except Exception as exc:
        logger.exception("Erro na execução manual: %s", exc)
        last_status = {"ok": False, "message": f"Erro na execução manual: {exc}", "at": datetime.now().isoformat()}
        return False, str(exc)


def iniciar_verificacao():
    """Inicia o loop principal em thread separada."""
    global running
    if running:
        return False
    stop_event.clear()
    t = threading.Thread(target=main_loop, daemon=True, name="botana-loop")
    t.start()
    return True


def parar_verificacao():
    """Interrompe o loop principal."""
    global running
    if not running:
        return False
    stop_event.set()
    running = False
    return True


def on_quit():
    """Chamado quando o usuário clica em 'Sair' no tray."""
    parar_verificacao()
    time.sleep(1)
    sys.exit(0)


def _json_response(handler: BaseHTTPRequestHandler, status: int, payload: dict):
    raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(raw)))
    handler.end_headers()
    handler.wfile.write(raw)


def _html_response(handler: BaseHTTPRequestHandler, status: int, html: str):
    raw = html.encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "text/html; charset=utf-8")
    handler.send_header("Content-Length", str(len(raw)))
    handler.end_headers()
    handler.wfile.write(raw)


def _render_server_html() -> str:
    return """<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>AnaBot - Painel</title>
<style>
body{font-family:Arial,sans-serif;background:#f2f2f2;margin:0;padding:16px}
.card{max-width:760px;margin:0 auto;background:#fff;padding:16px;border-radius:10px;border:1px solid #ddd}
.row{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}
button{padding:8px 12px;border:0;border-radius:8px;background:#176fe5;color:#fff;font-weight:700;cursor:pointer}
button.sec{background:#6b4128}
.muted{color:#666}
pre{background:#fafafa;border:1px solid #eee;border-radius:8px;padding:10px;white-space:pre-wrap}
</style></head><body>
<section class="card"><h2>AnaBot - Painel de execução</h2>
<div id="status" class="muted">Carregando status...</div>
<div class="row">
<button onclick="startLoop()">Iniciar loop</button>
<button class="sec" onclick="stopLoop()">Parar loop</button>
<button onclick="runNow()">Executar agora</button>
</div>
<h3>Último status</h3>
<pre id="details">-</pre>
</section>
<script>
async function api(path,opts){const r=await fetch(path,opts);return r.json();}
async function refresh(){const j=await api('/api/state');document.getElementById('status').textContent='Loop: '+(j.running?'ativo':'parado')+' | Intervalo: '+j.interval_seconds+'s';document.getElementById('details').textContent=JSON.stringify(j.last_status||{},null,2);}
async function startLoop(){await api('/api/start',{method:'POST'});refresh();}
async function stopLoop(){await api('/api/stop',{method:'POST'});refresh();}
async function runNow(){await api('/api/run-now',{method:'POST'});refresh();}
refresh();setInterval(refresh,3000);
</script></body></html>"""


def start_server(host: str, port: int, no_loop: bool = False):
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, fmt, *args):
            return

        def do_GET(self):
            parsed = urlparse(self.path)
            if parsed.path == "/":
                return _html_response(self, 200, _render_server_html())
            if parsed.path == "/api/state":
                return _json_response(
                    self,
                    200,
                    {
                        "ok": True,
                        "running": bool(running),
                        "interval_seconds": int(INTERVALO),
                        "last_status": dict(last_status),
                    },
                )
            return _json_response(self, 404, {"ok": False, "message": "Não encontrado"})

        def do_POST(self):
            parsed = urlparse(self.path)
            if parsed.path == "/api/start":
                started = iniciar_verificacao()
                return _json_response(self, 200, {"ok": True, "started": bool(started)})
            if parsed.path == "/api/stop":
                stopped = parar_verificacao()
                return _json_response(self, 200, {"ok": True, "stopped": bool(stopped)})
            if parsed.path == "/api/run-now":
                ok, msg = executar_um_ciclo()
                return _json_response(self, 200 if ok else 500, {"ok": bool(ok), "message": msg})
            return _json_response(self, 404, {"ok": False, "message": "Não encontrado"})

    if not no_loop:
        iniciar_verificacao()
    server = ThreadingHTTPServer((host, port), Handler)
    print(f"[AnaBot] Painel online em http://{host}:{port}")
    print("Ctrl+C para encerrar")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
        parar_verificacao()


def parse_args():
    p = argparse.ArgumentParser(description="AnaBot")
    p.add_argument("--server", action="store_true", help="Executa em modo servidor HTTP")
    p.add_argument("--host", default="127.0.0.1")
    p.add_argument("--port", type=int, default=8865)
    p.add_argument("--no-loop", action="store_true", help="Não inicia o loop automaticamente no modo servidor")
    return p.parse_args()


# =========================
# EXECUÇÃO PRINCIPAL
# =========================
if __name__ == "__main__":
    args = parse_args()
    if args.server:
        start_server(args.host, args.port, no_loop=args.no_loop)
    else:
        if run_tray is None:
            # fallback para ambientes sem tray (ex.: servidor sem interface)
            start_server("127.0.0.1", 8865, no_loop=False)
        else:
            run_tray(on_quit_callback=on_quit, start_callback=iniciar_verificacao)
