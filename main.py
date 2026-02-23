import argparse
import hashlib
import hmac
import json
import secrets
import os, re, time, gspread, threading, sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse, parse_qs
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from config import PLANILHAS, CNPJ_MVA, CNPJ_EH, INTERVALO, DOWNLOAD_DIR, GOOGLE_CREDENTIALS_SHEETS, GOOGLE_CREDENTIALS_GMAIL
from gmail_service import getGmailService, buscarMessagesEnviados, baixar_anexos_de_mensagem, ensure_label, LABEL_NAME
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
# FILTROS PARA DEBUG / ANÃLISE ISOLADA
# -----------------------
# Defina manualmente aqui (string) ou via variÃ¡vel de ambiente:
# Ex.: set SKIP_UNTIL_NF=12345       (Windows CMD)

# Se quiser que o script ignore tudo atÃ© achar a NF X, defina SKIP_UNTIL_NF
SKIP_UNTIL_NF = os.environ.get("SKIP_UNTIL_NF") or None  # ex: "12345"

# Se quiser processar somente uma NF especÃ­fica (ignorar todas as outras), defina NF_ALVO
NF_ALVO = os.environ.get("NF_ALVO") or None  # ex: "12345"

# Se NF_ALVO for usado e quiser que o script pare apÃ³s processar essa NF, coloque True
STOP_AFTER_NF = os.environ.get("STOP_AFTER_NF", "False").lower() in ("1", "true", "yes")
# -----------------------

stop_event = threading.Event()  # usado para parar o loop com seguranÃ§a
running = False # indica se o loop principal estÃ¡ ativo
last_status = {"ok": True, "message": "Aguardando", "at": None}
APPDATA_BASE = Path(os.getenv("APPDATA", str(Path.home() / "AppData" / "Roaming"))) / "Botana"
APPDATA_BASE.mkdir(parents=True, exist_ok=True)
_SETTINGS_FILE = APPDATA_BASE / "panel_settings.json"
_AUTH_FILE = APPDATA_BASE / "panel_auth.json"
_SETTINGS_LOCK = threading.RLock()
_AUTH_LOCK = threading.Lock()
_SESSIONS = {}
_SESSIONS_LOCK = threading.Lock()
_COOKIE_SESSION = "botana_session"
_SESSION_TTL_SECONDS = 8 * 60 * 60
_RUNTIME_SETTINGS = {
    "gmail_filter_mode": "last_30_days",
    "gmail_max_pages": 3,
    "gmail_page_size": 50,
    "loop_interval_minutes": max(1, int(INTERVALO // 60) if int(INTERVALO) > 0 else 30),
    "interval_seconds": int(INTERVALO),
    "max_messages": 100,
}
_EMAIL_CACHE = {"email": "", "error": "", "at": 0.0}
_NEXT_RUN_AT = 0.0
_GMAIL_SERVICE_LOCK = threading.Lock()


def _get_gmail_service_locked():
    with _GMAIL_SERVICE_LOCK:
        return getGmailService()


def _load_settings():
    with _SETTINGS_LOCK:
        if not _SETTINGS_FILE.exists():
            loop_min = max(1, min(720, int(_RUNTIME_SETTINGS.get("loop_interval_minutes", 30))))
            out = {
                "gmail_filter_mode": str(_RUNTIME_SETTINGS.get("gmail_filter_mode", "last_30_days")),
                "gmail_max_pages": max(1, min(20, int(_RUNTIME_SETTINGS.get("gmail_max_pages", 3)))),
                "gmail_page_size": max(1, min(500, int(_RUNTIME_SETTINGS.get("gmail_page_size", 50)))),
                "loop_interval_minutes": loop_min,
                "interval_seconds": max(30, min(86400, loop_min * 60)),
                "max_messages": max(1, min(1000, int(_RUNTIME_SETTINGS.get("max_messages", 100)))),
            }
            _RUNTIME_SETTINGS.update(out)
            _SETTINGS_FILE.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
            return dict(out)
        try:
            raw = json.loads(_SETTINGS_FILE.read_text(encoding="utf-8"))
        except Exception:
            raw = {}
        mode = str(raw.get("gmail_filter_mode", "last_30_days")).strip()
        if mode not in {
            "last_15_days",
            "last_30_days",
            "last_45_days",
            "last_60_days",
            "current_week",
            "previous_month",
            "current_and_previous_month",
        }:
            mode = "last_30_days"
        loop_min = max(1, min(720, int(raw.get("loop_interval_minutes", _RUNTIME_SETTINGS.get("loop_interval_minutes", 30)))))
        max_pages = max(1, min(20, int(raw.get("gmail_max_pages", _RUNTIME_SETTINGS.get("gmail_max_pages", 3)))))
        page_size = max(1, min(500, int(raw.get("gmail_page_size", _RUNTIME_SETTINGS.get("gmail_page_size", 50)))))
        out = {
            "gmail_filter_mode": mode,
            "gmail_max_pages": max_pages,
            "gmail_page_size": page_size,
            "loop_interval_minutes": loop_min,
            "interval_seconds": max(30, min(86400, loop_min * 60)),
            "max_messages": max(1, min(1000, int(raw.get("max_messages", max_pages * page_size)))),
        }
        _RUNTIME_SETTINGS.update(out)
        global _NEXT_RUN_AT
        if _NEXT_RUN_AT <= 0:
            _NEXT_RUN_AT = time.time() + int(out.get("interval_seconds", INTERVALO))
        return out


def _save_settings(data: dict):
    with _SETTINGS_LOCK:
        mode = str(data.get("gmail_filter_mode", _RUNTIME_SETTINGS.get("gmail_filter_mode", "last_30_days"))).strip()
        if mode not in {
            "last_15_days",
            "last_30_days",
            "last_45_days",
            "last_60_days",
            "current_week",
            "previous_month",
            "current_and_previous_month",
        }:
            mode = "last_30_days"
        max_pages = max(1, min(20, int(data.get("gmail_max_pages", _RUNTIME_SETTINGS.get("gmail_max_pages", 3)))))
        page_size = max(1, min(500, int(data.get("gmail_page_size", _RUNTIME_SETTINGS.get("gmail_page_size", 50)))))
        loop_min = max(1, min(720, int(data.get("loop_interval_minutes", _RUNTIME_SETTINGS.get("loop_interval_minutes", 30)))))
        out = {
            "gmail_filter_mode": mode,
            "gmail_max_pages": max_pages,
            "gmail_page_size": page_size,
            "loop_interval_minutes": loop_min,
            "interval_seconds": max(30, min(86400, loop_min * 60)),
            "max_messages": max(1, min(1000, int(data.get("max_messages", max_pages * page_size)))),
        }
        _RUNTIME_SETTINGS.update(out)
        _SETTINGS_FILE.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
        global _NEXT_RUN_AT
        _NEXT_RUN_AT = time.time() + int(out.get("interval_seconds", INTERVALO))


def _password_hash(password: str, salt_hex: str) -> str:
    try:
        salt = bytes.fromhex(str(salt_hex or ""))
        digest = hashlib.pbkdf2_hmac("sha256", (password or "").encode("utf-8"), salt, 120000)
        return digest.hex()
    except Exception:
        return ""


def _normalize_username(username: str) -> str:
    return str(username or "").strip().lower()


def _load_auth():
    with _AUTH_LOCK:
        if _AUTH_FILE.exists():
            try:
                data = json.loads(_AUTH_FILE.read_text(encoding="utf-8"))
                if isinstance(data, dict) and isinstance(data.get("users"), list) and data["users"]:
                    return data
            except Exception:
                pass
        salt_hex = secrets.token_hex(16)
        data = {
            "users": [
                {
                    "username": "dev",
                    "role": "dev",
                    "salt": salt_hex,
                    "password_hash": _password_hash("dev", salt_hex),
                    "created_at": datetime.now().isoformat(),
                }
            ]
        }
        _AUTH_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        print("[Botana] Login inicial: usuario=dev senha=dev")
        return data


def _verify_login(username: str, password: str) -> bool:
    user = _normalize_username(username)
    try:
        data = _load_auth()
        for item in data.get("users", []):
            if _normalize_username(item.get("username", "")) != user:
                continue
            calc = _password_hash(password or "", str(item.get("salt", "")))
            saved = str(item.get("password_hash", "") or "")
            if not calc or not saved:
                return False
            return hmac.compare_digest(calc, saved)
        return False
    except Exception as exc:
        logger.exception("Falha ao validar login: %s", exc)
        return False


def _role_of(username: str) -> str:
    user = _normalize_username(username)
    data = _load_auth()
    for item in data.get("users", []):
        if _normalize_username(item.get("username", "")) == user:
            role = str(item.get("role", "user")).lower()
            return role if role in {"dev", "admin", "user"} else "user"
    return "user"


def _can_operate(username: str) -> bool:
    return _role_of(username) in {"dev", "admin"}


def _create_session(username: str) -> str:
    token = secrets.token_urlsafe(32)
    with _SESSIONS_LOCK:
        _SESSIONS[token] = {"user": _normalize_username(username), "exp": time.time() + _SESSION_TTL_SECONDS}
    return token


def _read_cookie_session(handler: BaseHTTPRequestHandler) -> str:
    raw = str(handler.headers.get("Cookie", "") or "")
    for part in raw.split(";"):
        p = part.strip()
        if p.startswith(f"{_COOKIE_SESSION}="):
            return p.split("=", 1)[1].strip()
    return ""


def _current_session_user(handler: BaseHTTPRequestHandler) -> str | None:
    token = _read_cookie_session(handler)
    if not token:
        return None
    with _SESSIONS_LOCK:
        item = _SESSIONS.get(token)
        if not item:
            return None
        if float(item.get("exp", 0)) < time.time():
            _SESSIONS.pop(token, None)
            return None
        return str(item.get("user", "")).strip() or None

def escolher_planilha_por_cnpj_e_ano(cnpj: str, ano: str):
    if cnpj == CNPJ_MVA:
        return PLANILHAS["MVA"].get(ano)
    if cnpj == CNPJ_EH:
        return PLANILHAS["EH"].get(ano)
    return None

def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def processar_emails_enviados():
    service = _get_gmail_service_locked()
    msgs = buscarMessagesEnviados(service, max_results=int(_RUNTIME_SETTINGS.get("max_messages", 100)))
    if not msgs:
        logger.info("Nenhuma mensagem enviada com XML encontrada.")
        return

    total_processados = 0

    for m in msgs:   
        msg_id = m.get("id")
        logger.info("ðŸ“§ Abrindo mensagem ID: %s", msg_id)

        arquivos = baixar_anexos_de_mensagem(service, msg_id)
        if not arquivos:
            logger.info("Nenhum anexo salvo para mensagem %s", msg_id)
            continue

        dados_xmls = []
        boletos = []

        # ðŸ” Processa todos os anexos baixados
        for arquivo in arquivos:
            nome_arquivo = os.path.basename(arquivo)

            try:
                # =============================
                # ðŸ“„ XML â†’ extrai dados
                # =============================
                if arquivo.lower().endswith(".xml"):
                    try:
                        dados = extrairDadosXML(arquivo)
                        # ðŸ” Ignora vendas Ã  vista
                        nat_op = dados.get("naturezaOperacao", "").strip().upper()
                        dest = dados.get("destinatario", "")
                        if ( "VISTA" in nat_op or "VENDA A VISTA" in nat_op):
                            # Checa se a mensagem ja foi processada no relatorio atual:
                            if dados.get('nf') not in consolidarRelatorioTMP(): 
                                escreverRelatorio(f"{_now()} - ðŸ’° NF {dados.get('nf')} ignorada (venda Ã  vista).")
                                continue
                            else: logger.info(f"{cor_ciano}NF {dados['nf']} jÃ¡ registrada no relatÃ³rio, nÃ£o duplicando a mensagem de ignorada.{reset}") 
                            continue
                        if ( CNPJ_MVA.replace(".", "").replace("/", "").replace("-", "") in dest or
                             CNPJ_EH.replace(".", "").replace("/", "").replace("-", "") in dest ):
                            logger.info(f"[DEBUG IGNORE RESULT] NF {dados['nf']} ignorada (destinatÃ¡rio Ã© o nosso: {dest})")
                            escreverRelatorio(f"{_now()} - ðŸ’° NF {dados.get('nf')} ignorada (destinatÃ¡rio Ã© o nosso).")
                            continue
                        if not dados:
                            motivo = dados.get("motivo_ignoracao", "Desconhecido") if isinstance(dados, dict) else "Desconhecido"
                            logger.info(f"Ignorado XML (motivo: {motivo}).")
                            escreverRelatorio(f"{_now()} - âš ï¸ XML {nome_arquivo} ignorado (motivo: {motivo})")
                            continue

                        dados_xmls.append(dados)

                    except Exception as e:
                        escreverRelatorio(f"{_now()} - âŒ Erro extraindo XML {nome_arquivo}: {e}")
                        logger.exception("Erro extraindo XML %s: %s", arquivo, e)

                # =============================
                # ðŸ“‘ PDF â†’ tenta identificar boleto
                # =============================
                elif arquivo.lower().endswith(".pdf"): # mudar pra elif se o bloco de cima for realmente necessÃ¡rio
                    nome_upper = nome_arquivo.upper()

                    # ðŸ” Trata nomes parecidos com BOLETO (erros comuns tipo BOLTO, BOLETA, BOLETT, etc)
                    padrao_boleto = r"[_\s-]?(BLT|BOLET[OA]?|BOLTO|BOLETOO|BOLETT?)"

                    if re.search(padrao_boleto, nome_upper):
                        match = re.findall(r"([0-9]{2,}-?[0-9]+)", nome_upper)
                        if match:
                            num_boleto = match[-1]
                            boletos.append(num_boleto)
                            logger.info("ðŸ”¢ Boleto identificado no nome: %s (BLT %s)", nome_arquivo, num_boleto)
                        else:
                            logger.info("Nenhum nÃºmero de boleto encontrado no nome: %s", nome_arquivo)
                    elif arquivo.lower().endswith(".pdf"):
                        nome_upper = nome_arquivo.upper()

                        # ðŸ” Palavras que indicam boleto (considera erros comuns)
                        padrao_boleto = r"\b(BOLET[OA]?|BOLTO|BOLETOO|BOLETT?|BLT)\b"

                        # SÃ³ tenta identificar nÃºmero se o nome realmente tiver algo prÃ³ximo de "boleto"
                        if re.search(padrao_boleto, nome_upper):
                            match = re.findall(r"([0-9]{2,}-?[0-9]+)", nome_upper)
                            if match:
                                num_boleto = match[-1]
                                boletos.append(num_boleto)
                                logger.info("ðŸ”¢ Boleto identificado no nome: %s (BLT %s)", nome_arquivo, num_boleto)
                            else:
                                logger.info("ðŸ“Ž PossÃ­vel boleto sem nÃºmero identificado: %s", nome_arquivo)
                        else:
                            logger.info("ðŸ“„ PDF ignorado (nÃ£o parece boleto): %s", nome_arquivo)

                else:
                    logger.info("Arquivo nÃ£o identificado como boleto: %s", nome_arquivo)

            finally:
                # ðŸ§¹ Remove sempre o anexo local (independente do tipo)
                try:
                    os.remove(arquivo)
                    logger.debug(f"ðŸ§¹ Anexo removido: {arquivo}")
                except FileNotFoundError:
                    pass
                except Exception as e:
                    logger.warning(f"âš ï¸ Falha ao remover {arquivo}: {e}")

        # =============================
        # ðŸ·ï¸ Marca o e-mail como processado
        # =============================
        try:
            marcar_mensagem_com_label(service, msg_id)
            logger.info("ðŸ·ï¸ E-mail %s marcado com 'XML Processado Botana'", msg_id)
        except Exception as e:
            logger.exception("Falha ao aplicar rÃ³tulo: %s", e)
            
        # âš ï¸ Nenhum XML â†’ pula este e-mail
        if not dados_xmls:
            logger.info("Nenhum XML vÃ¡lido encontrado neste e-mail.")
            continue

        # =============================
        # ðŸ§¾ Atualiza planilhas
        # =============================
        for dados_xml in dados_xmls:
            # --- FILTRAGEM POR NF (para debug/anÃ¡lise isolada) ---
            nf_num = str(dados_xml.get("nf", "")).strip()

            # NF_ALVO: processa somente essa NF (ignora as outras)
            if NF_ALVO:
                if nf_num != str(NF_ALVO):
                    logger.info(f"ðŸ”Ž Pulando NF {nf_num} (NF_ALVO ativo: {NF_ALVO})")
                    continue
                else:
                    logger.info(f"âœ… NF_ALVO encontrada: {nf_num}")

            # SKIP_UNTIL_NF: ignora tudo atÃ© encontrar essa NF; quando encontrada, passa a processar normalmente
            if SKIP_UNTIL_NF:
                # usa atributo da funÃ§Ã£o para manter estado entre ciclos enquanto o processo estÃ¡ vivo
                if not hasattr(processar_emails_enviados, "_skip_reached"):
                    processar_emails_enviados._skip_reached = False

                if not processar_emails_enviados._skip_reached:
                    if nf_num == str(SKIP_UNTIL_NF):
                        processar_emails_enviados._skip_reached = True
                        logger.info(f"ðŸŽ¯ SKIP_UNTIL_NF: NF {nf_num} encontrada â€” a partir daqui serÃ¡ processada.")
                    else:
                        logger.info(f"â­ SKIP_UNTIL_NF ativo, pulando NF {nf_num}")
                        continue

            # Se chegou atÃ© aqui, a NF serÃ¡ processada normalmente.
            # Se NF_ALVO + STOP_AFTER_NF: apÃ³s processar, se encerra o loop/principal para anÃ¡lise isolada.

            cnpj_emit = dados_xml.get("cnpjEmitente")
            ano = dados_xml.get("anoVencimento")
            planilha_id = escolher_planilha_por_cnpj_e_ano(cnpj_emit, ano)

            if not planilha_id:
                logger.warning("CNPJ %s ou ano %s sem planilha configurada.", cnpj_emit, ano)
                continue

            # Itera sobre todas as parcelas â€” MAPEAMENTO correto de boletos â†’ parcelas
            parcelas = dados_xml.get("parcelas", [])
            n_parcelas = len(parcelas)
            n_boletos = len(boletos)

            # monta lista de boletos por parcela (mesmo tamanho de parcelas)
            if n_parcelas == 0:
                continue  # nada a fazer

            if n_boletos == 0:
                boletos_map = [None] * n_parcelas
            else:
                # Se tiver igual, mapeia 1:1; se menor, preenche em ordem; se maior, usa sÃ³ os primeiros N
                boletos_map = [boletos[i] if i < n_boletos else None for i in range(n_parcelas)]
                if n_boletos > n_parcelas:
                    logger.info("âš ï¸ Mais boletos (%d) que parcelas (%d). Sobraram: %s", n_boletos, n_parcelas, boletos[n_parcelas:])

            # Agora processa 1 vez por parcela, usando o boleto mapeado (ou None)
            for idx, parcela in enumerate(parcelas):
                num_boleto = boletos_map[idx]
                dados_parcela = dados_xml.copy()
                dados_parcela.update({
                    "vencimento": parcela["vencimento"],
                    "numParcela": parcela["numParcela"],
                    "valorParcela": parcela["valor"],
                    "boleto": num_boleto  # adiciona campo explÃ­cito (opcional)
                })

                # Ajusta descriÃ§Ã£o com o boleto mapeado (se houver)
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
                        # Se NF_ALVO + STOP_AFTER_NF -> encerra o processo principal para anÃ¡lise isolada.
                        if NF_ALVO and STOP_AFTER_NF:
                            logger.info(f"ðŸ NF_ALVO {NF_ALVO} processada. STOP_AFTER_NF=True -> encerrando execuÃ§Ã£o.")
                            # forÃ§a saÃ­da limpa do loop principal retornando da funÃ§Ã£o
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
    global running, last_status, _NEXT_RUN_AT
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    running = True
    logger.info("[Botana] Loop iniciado")
    while not stop_event.is_set():
        _NEXT_RUN_AT = time.time() + int(_RUNTIME_SETTINGS.get("interval_seconds", INTERVALO))
        try:
            processar_emails_enviados()
            last_status = {"ok": True, "message": "Ciclo executado com sucesso", "at": datetime.now().isoformat()}
        except Exception as e:
            logger.exception("Erro no ciclo principal: %s", e)
            last_status = {"ok": False, "message": f"Erro no ciclo: {e}", "at": datetime.now().isoformat()}
        if stop_event.wait(int(_RUNTIME_SETTINGS.get("interval_seconds", INTERVALO))):
            break
    running = False
    logger.info("[Botana] Loop finalizado")


def executar_um_ciclo():
    global last_status, _NEXT_RUN_AT
    try:
        processar_emails_enviados()
        last_status = {"ok": True, "message": "ExecuÃ§Ã£o manual concluÃ­da", "at": datetime.now().isoformat()}
        _NEXT_RUN_AT = time.time() + int(_RUNTIME_SETTINGS.get("interval_seconds", INTERVALO))
        return True, "ExecuÃ§Ã£o manual concluÃ­da"
    except Exception as exc:
        logger.exception("Erro na execuÃ§Ã£o manual: %s", exc)
        last_status = {"ok": False, "message": f"Erro na execuÃ§Ã£o manual: {exc}", "at": datetime.now().isoformat()}
        _NEXT_RUN_AT = time.time() + int(_RUNTIME_SETTINGS.get("interval_seconds", INTERVALO))
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
    """Chamado quando o usuÃ¡rio clica em 'Sair' no tray."""
    parar_verificacao()
    time.sleep(1)
    sys.exit(0)


def _read_json(handler: BaseHTTPRequestHandler) -> dict:
    try:
        raw_len = handler.headers.get("Content-Length", "0")
        size = int(raw_len) if str(raw_len).strip().isdigit() else 0
        if size <= 0:
            return {}
        raw = handler.rfile.read(size)
        if not raw:
            return {}
        data = json.loads(raw.decode("utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _json_response(handler: BaseHTTPRequestHandler, status: int, payload: dict, extra_headers: dict | None = None):
    raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(raw)))
    if extra_headers:
        for k, v in extra_headers.items():
            handler.send_header(k, v)
    handler.end_headers()
    handler.wfile.write(raw)


def _html_response(handler: BaseHTTPRequestHandler, status: int, html: str):
    raw = html.encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "text/html; charset=utf-8")
    handler.send_header("Content-Length", str(len(raw)))
    handler.end_headers()
    handler.wfile.write(raw)


def _find_store_image_path() -> Path | None:
    names = [
        "Arte MVA logo Metalico (1).png",
        "arte mva logo metalico (1).png",
    ]
    local_rel = Path("assets") / "branding"
    candidates = []

    for n in names:
        candidates.append(Path.cwd() / local_rel / n)
        candidates.append(Path.cwd() / n)
        candidates.append(Path(__file__).resolve().parent / local_rel / n)
        candidates.append(Path(__file__).resolve().parent / n)
        candidates.append(Path(sys.executable).resolve().parent / local_rel / n)
        candidates.append(Path(sys.executable).resolve().parent / n)
        candidates.append(Path.home() / "Desktop" / n)

    # Fallback para a imagem do FinanceBot, se não houver imagem local do Botana.
    for n in names:
        candidates.append(Path("C:/FinanceBot/assets/branding") / n)
        candidates.append(Path("C:/FinanceBot") / n)

    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        for n in names:
            candidates.append(Path(meipass) / local_rel / n)
            candidates.append(Path(meipass) / n)

    for p in candidates:
        if p.exists() and p.is_file():
            return p
    return None


def _send_store_image(handler: BaseHTTPRequestHandler) -> bool:
    img = _find_store_image_path()
    if not img:
        return False
    try:
        raw = img.read_bytes()
        handler.send_response(200)
        handler.send_header("Content-Type", "image/png")
        handler.send_header("Cache-Control", "no-cache")
        handler.send_header("Content-Length", str(len(raw)))
        handler.end_headers()
        handler.wfile.write(raw)
        return True
    except Exception:
        return False


def _history_from_reports(limit: int = 300, query: str = "") -> list[dict]:
    out = []
    q = str(query or "").strip().lower()
    rel_dir = Path(RELATORIO_DIR)
    if not rel_dir.exists():
        return out
    files = sorted(rel_dir.glob("*.txt"), key=lambda p: p.stat().st_mtime, reverse=True)
    for fp in files:
        try:
            for line in fp.read_text(encoding="utf-8", errors="ignore").splitlines():
                line = line.strip()
                if not line:
                    continue
                if q and q not in line.lower():
                    continue
                # formato típico: "YYYY-MM-DD HH:MM:SS - mensagem"
                if " - " in line:
                    dt, msg = line.split(" - ", 1)
                else:
                    dt, msg = "", line
                out.append({"at": dt.strip(), "message": msg.strip(), "raw": line})
                if len(out) >= limit:
                    return out
        except Exception:
            continue
    return out


def _latest_report_file() -> Path | None:
    rel_dir = Path(RELATORIO_DIR)
    if not rel_dir.exists():
        return None
    files = sorted(rel_dir.glob("relatorio_*.txt"), key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0] if files else None


def _daily_report_data() -> dict:
    report_file = _latest_report_file()
    if not report_file:
        return {
            "exists": False,
            "path": "",
            "updated_at": "",
            "totals": {"processados": 0, "ignorados": 0, "avisos_ciclo": 0, "avisos_dia": 0},
            "processados": [],
            "ignorados": [],
            "avisos": [],
        }

    try:
        lines = report_file.read_text(encoding="utf-8", errors="ignore").splitlines()
    except Exception:
        lines = []

    processados: list[str] = []
    ignorados: list[str] = []
    avisos: list[str] = []

    for raw in lines:
        line = str(raw or "").strip()
        if not line:
            continue
        text = line.split(" - ", 1)[1].strip() if " - " in line else line
        low = text.lower()
        if "erro" in low or "falha" in low:
            avisos.append(text)
            continue
        if "ignorada" in low or "ignorado" in low:
            ignorados.append(text)
            continue
        processados.append(text)

    totals = {
        "processados": len(processados),
        "ignorados": len(ignorados),
        "avisos_ciclo": len(avisos),
        "avisos_dia": len(avisos),
    }
    updated_at = datetime.fromtimestamp(report_file.stat().st_mtime).strftime("%d/%m/%Y, %H:%M:%S")
    return {
        "exists": True,
        "path": str(report_file),
        "updated_at": updated_at,
        "totals": totals,
        "processados": processados[-8:],
        "ignorados": ignorados[-8:],
        "avisos": avisos[-8:],
    }


def _connected_email(force: bool = False) -> dict:
    now = time.time()
    if not force and _EMAIL_CACHE.get("at", 0.0) and (now - float(_EMAIL_CACHE.get("at", 0.0)) < 120):
        return dict(_EMAIL_CACHE)
    try:
        service = _get_gmail_service_locked()
        profile = service.users().getProfile(userId="me").execute()
        _EMAIL_CACHE.update(
            {
                "email": str(profile.get("emailAddress", "")).strip(),
                "error": "",
                "at": now,
            }
        )
    except Exception as exc:
        _EMAIL_CACHE.update({"error": str(exc), "at": now})
    return dict(_EMAIL_CACHE)


def _reprocess_recent(days: int, max_messages: int, mark_unread: bool) -> dict:
    service = _get_gmail_service_locked()
    label_id = ensure_label(service, LABEL_NAME)
    after = (datetime.now() - timedelta(days=max(1, int(days)))).strftime("%Y/%m/%d")
    query = f'after:{after} label:"{LABEL_NAME}"'
    resp = service.users().messages().list(userId="me", q=query, maxResults=max(1, min(1000, int(max_messages)))).execute()
    messages = resp.get("messages", []) or []
    changed = 0
    failed = 0
    for item in messages:
        msg_id = str(item.get("id", "")).strip()
        if not msg_id:
            continue
        body = {"removeLabelIds": [label_id], "addLabelIds": ["UNREAD"] if mark_unread else []}
        try:
            service.users().messages().modify(userId="me", id=msg_id, body=body).execute()
            changed += 1
        except Exception:
            failed += 1
    return {
        "ok": True,
        "matched": len(messages),
        "changed": changed,
        "failed": failed,
        "mark_unread": bool(mark_unread),
    }


def _reauthenticate_gmail() -> dict:
    token_path = str(GOOGLE_CREDENTIALS_GMAIL).replace(".json", "_token.json")
    try:
        if os.path.exists(token_path):
            os.remove(token_path)
    except Exception:
        pass
    _get_gmail_service_locked()
    _connected_email(force=True)
    return {"ok": True, "message": "Reautenticação concluída"}



def _render_login_html() -> str:
    return """<!doctype html>
<html lang="pt-br"><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Botana - Login</title>
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Lexend:wght@300;400;600;700;800&display=swap" rel="stylesheet">
<style>
:root{--o:#da7a1c;--o2:#ee9b2f;--b:#4a2b18;--b2:#6b4128}
*{box-sizing:border-box}
body{margin:0;min-height:100vh;font-family:'Lexend',Arial,sans-serif;background:linear-gradient(160deg,rgba(41,22,11,.78),rgba(95,56,28,.72)),url('/assets/store-bg') center/cover fixed;display:flex;justify-content:center;align-items:center;padding:12px;color:#2a1b12}
.card{width:min(420px,96vw);border-radius:16px;border:1px solid rgba(231,200,168,.9);background:linear-gradient(180deg,rgba(255,250,246,.96),rgba(255,245,235,.92));box-shadow:0 24px 60px rgba(21,11,6,.35);padding:16px}
h1{margin:0 0 6px;color:var(--b);font-size:1.2rem}
p{margin:0 0 12px;color:#6b4128}
label{display:block;margin-top:8px;font-weight:600;color:#5c341c}
input{width:100%;padding:10px;margin-top:4px;border:1px solid #d6b18f;border-radius:8px;background:#fffdfb;font-family:inherit}
button{margin-top:12px;width:100%;padding:10px 12px;border:0;border-radius:9px;background:linear-gradient(90deg,var(--o),var(--o2));color:#2b1408;font-weight:700;cursor:pointer}
.msg{margin-top:10px;font-size:.9rem;color:#9c2c1d;min-height:20px}
</style></head><body>
<section class="card">
<h1>Acesso ao Botana</h1>
<p>Entre com usuário e senha para continuar</p>
<label>Usuário</label><input id="u" type="text" autocomplete="username"/>
<label>Senha</label><input id="p" type="password" autocomplete="current-password"/>
<button id="b" onclick="login()">Entrar</button>
<div id="m" class="msg"></div>
</section>
<script>
async function login(){
  const u=document.getElementById('u').value||'';
  const p=document.getElementById('p').value||'';
  const m=document.getElementById('m');
  const b=document.getElementById('b');
  b.disabled=true;
  m.textContent='Validando acesso';
  try{
    const r=await fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:u,password:p})});
    const j=await r.json();
    if(r.ok&&j.ok){window.location.href='/';return;}
    m.textContent=j.message||'Usuário ou senha inválidos';
  }catch(_){
    m.textContent='Falha ao conectar com o servidor';
  }finally{b.disabled=false;}
}
['u','p'].forEach(id=>{document.getElementById(id).addEventListener('keydown',(e)=>{if(e.key==='Enter')login();});});
</script></body></html>"""
def _render_server_html() -> str:
    return """<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Botana - Painel</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Lexend:wght@300;400;600;700;800&display=swap');
:root{--o:#da7a1c;--o2:#ee9b2f;--b:#4a2b18;--bg:#f8efe6;--br:#e4c6a7}
*{box-sizing:border-box}
body{margin:0;min-height:100vh;font-family:'Lexend',Arial,sans-serif;background:linear-gradient(160deg,rgba(41,22,11,.78),rgba(95,56,28,.72)),url('/assets/store-bg') center/cover fixed;display:flex;justify-content:center;align-items:center;padding:12px;color:#2a1b12}
.app{width:min(1150px,100%);border-radius:18px;overflow:hidden;border:1px solid rgba(231,200,168,.9);background:linear-gradient(180deg,rgba(255,250,246,.96),rgba(255,245,235,.92));box-shadow:0 24px 60px rgba(21,11,6,.35)}
.top{padding:14px 20px;background:linear-gradient(90deg,var(--b),var(--o));color:#fff9f3;font-weight:700;display:flex;justify-content:space-between;align-items:center;gap:8px}
.top-right{display:flex;align-items:center;gap:8px}
.whoami{font-size:.82rem;opacity:.95}
.logout-btn{padding:7px 10px;border:1px solid rgba(255,244,234,.5);border-radius:8px;background:rgba(255,244,234,.12);color:#fff9f3;font-weight:700;cursor:pointer}
.logout-btn:hover{background:rgba(255,244,234,.2)}
.status-pill{display:inline-flex;align-items:center;gap:6px;padding:3px 8px;border-radius:999px;font-size:.76rem;font-weight:700;border:1px solid}
.ok{background:#e8f6ea;color:#2e7d32;border-color:#b6dfbf}
.off{background:#fff3e0;color:#8b4f19;border-color:#f2c8a3}
.err{background:#fdecec;color:#b42b2b;border-color:#f1bbbb}
.tabs{display:flex;gap:8px;padding:10px;margin-bottom:6px}
.tab-btn{background:#fff5ea;color:#5a311b;border:1px solid #d7b393;border-radius:9px;padding:8px 12px;font-weight:700;cursor:pointer}
.tab-btn.active{background:linear-gradient(90deg,var(--o),var(--o2));border-color:transparent;color:#2b1408}
.hidden{display:none}
#tabMain{padding:0 10px 10px;display:grid;gap:9px}
#tabHist,#tabDiag{padding:0 10px 10px}
.card{background:rgba(255,248,240,.92);border:1px solid #e7c8a8;border-radius:13px;padding:10px;box-shadow:0 8px 20px rgba(21,11,6,.06)}
h3{margin:0 0 8px;color:var(--b);font-size:.98rem}
label{display:block;font-weight:600;color:#5c341c;font-size:.9rem}
.muted{color:#6c4a35;font-size:.84rem}
.btns{display:flex;gap:8px;flex-wrap:wrap}
button{padding:9px 12px;border:0;border-radius:9px;background:linear-gradient(90deg,var(--o),var(--o2));color:#2b1408;font-weight:700;cursor:pointer;font-size:.9rem}
button.sec{background:linear-gradient(90deg,#7a4d30,#5b341f);color:#fff9f3}
button.warn{background:linear-gradient(90deg,#bc2d2d,#8f2020);color:#fff}
.inp{padding:8px;margin-top:4px;border:1px solid #d6b18f;border-radius:8px;background:#fffdfb;font-family:inherit;width:145px}
pre{margin:0;background:#fff7ef;border:1px dashed #cf9f78;padding:8px;border-radius:10px;overflow:auto;max-height:340px;white-space:pre-wrap}
.kpi{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px}
.k{border:1px solid #deb999;border-radius:10px;padding:9px;background:#fff9f3}
.k .n{font-size:1.2rem;font-weight:800;color:#6b4128}
.k .t{font-size:.78rem;color:#6b4128}
.kpi.daily{grid-template-columns:repeat(4,minmax(0,1fr))}
.rmeta{margin-top:6px}
.lists{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px;margin-top:8px}
.box{border:1px solid #e4c6a7;border-radius:10px;background:#fffdfb;padding:8px}
.box h4{margin:0 0 6px;font-size:.85rem;color:#58311b}
.box ul{margin:0;padding-left:16px;max-height:160px;overflow:auto}
.box li{margin:3px 0;font-size:.8rem}
.status-grid{display:grid;grid-template-columns:1fr;gap:8px}
.s{border:1px solid #d5b08f;background:#fffaf6;border-radius:11px;padding:10px}
.h{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;font-weight:700;color:#5b321c}
.problem{color:#862818;font-size:.82rem;margin-top:4px}
input,select{padding:8px;margin-top:4px;border:1px solid #d6b18f;border-radius:8px;background:#fffdfb;font-family:inherit}
.cfg-grid{display:grid;grid-template-columns:minmax(0,1fr) minmax(170px,210px) minmax(260px,320px);gap:10px;align-items:start}
.cfg-main{display:grid;gap:8px}
.cfg-fields{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px;align-items:end}
.cfg-fields > div{display:flex;flex-direction:column}
.cfg-fields > div input,.cfg-fields > div select{width:100%}
.cfg-actions{display:flex;justify-content:flex-start;align-items:center;gap:8px;flex-wrap:wrap}
.auth-card .btns{flex-direction:column}
.auth-card .btns button{width:100%}
.reproc-card .btns{margin-top:8px}
.reproc-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px;align-items:end}
.reproc-grid > div{display:flex;flex-direction:column}
.cb{margin-top:8px;display:inline-flex;align-items:center;gap:8px}
.ov{position:fixed;inset:0;z-index:99999;display:none;align-items:center;justify-content:center;background:rgba(22,10,5,.78);backdrop-filter:blur(3px)}
.ov.show{display:flex}
.ovb{width:min(440px,92vw);border-radius:14px;border:1px solid #f0c89d;background:linear-gradient(180deg,#fff6ec,#ffe8d4);text-align:center;padding:18px}
.cnt{margin-top:12px;font-size:2.4rem;font-weight:800;color:#b05714}
@media(max-width:900px){.lists{grid-template-columns:1fr}.cfg-grid{grid-template-columns:1fr}.cfg-fields{grid-template-columns:1fr 1fr}.reproc-grid{grid-template-columns:1fr}}
@media(max-width:640px){.top-right{flex-direction:column;align-items:flex-end}}
</style></head><body>
<div id="ov" class="ov"><div class="ovb"><h4>Reautenticação em andamento</h4><p>Troque para a conta correta no navegador<br/>A autenticação começará em:</p><div id="cnt" class="cnt">5</div></div></div>
<main class="app">
  <section class="top">
    <span>Botana - Painel de Controle MVA</span>
    <div class="top-right">
      <span id="who" class="whoami">Usuário: -</span>
      <button class="logout-btn" onclick="logout()">Sair</button>
      <span id="pill" class="status-pill off"><span>●</span><span>Aguardando</span></span>
    </div>
  </section>

  <div class="tabs">
    <button id="tabBtnMain" class="tab-btn active" onclick="switchTab('main')">Painel</button>
    <button id="tabBtnHist" class="tab-btn" onclick="switchTab('hist')">Histórico</button>
    <button id="tabBtnDiag" class="tab-btn" onclick="switchTab('diag')">Diagnóstico</button>
  </div>

  <section id="tabMain">
    <section class="card">
      <h3>Status da conta de e-mail</h3>
      <div class="status-grid">
        <div class="s">
          <div class="h">
            <span>Conta Botana</span>
            <span id="accBadge" class="status-pill off"><span>●</span><span>Aguardando</span></span>
          </div>
          <div id="accEmail" class="muted">E-mail conectado: -</div>
          <div id="accDetail" class="muted">Aguardando leitura</div>
          <div id="accProblem" class="problem hidden">-</div>
        </div>
      </div>
      <div id="cool" class="muted" style="margin-top:8px">Próxima verificação automática: -</div>
    </section>

    <section class="card">
      <h3>Relatórios diários</h3>
      <div class="kpi daily">
        <div class="k"><div id="kp1" class="n">0</div><div class="t">Processados</div></div>
        <div class="k"><div id="kp2" class="n">0</div><div class="t">Ignorados</div></div>
        <div class="k"><div id="kp3" class="n">0</div><div class="t">Avisos no ciclo</div></div>
        <div class="k"><div id="kp4" class="n">0</div><div class="t">Avisos no dia</div></div>
      </div>
      <div id="rmeta" class="muted rmeta">Sem relatório encontrado ainda</div>
      <div class="lists">
        <div class="box"><h4>Últimos processados</h4><ul id="lp"><li>Sem itens</li></ul></div>
        <div class="box"><h4>Últimos ignorados</h4><ul id="li"><li>Sem itens</li></ul></div>
        <div class="box"><h4>Avisos recentes</h4><ul id="la"><li>Sem itens</li></ul></div>
      </div>
    </section>

    <section class="cfg-grid">
      <article class="card">
        <h3>Configuração do Gmail</h3>
        <div class="cfg-main">
          <div class="cfg-fields">
            <div>
              <label>Período</label>
              <select id="mode">
                <option value="last_15_days">Últimos 15 dias</option>
                <option value="last_30_days">Últimos 30 dias</option>
                <option value="last_45_days">Últimos 45 dias</option>
                <option value="last_60_days">Últimos 60 dias</option>
                <option value="current_week">Semana atual</option>
                <option value="previous_month">Mês anterior</option>
                <option value="current_and_previous_month">Mês atual + mês anterior</option>
              </select>
            </div>
            <div>
              <label>Máx páginas</label>
              <input id="maxPages" type="number" min="1" max="20"/>
            </div>
            <div>
              <label>Tamanho da página</label>
              <input id="pageSize" type="number" min="1" max="500"/>
            </div>
            <div>
              <label>Intervalo de leitura</label>
              <input id="intervalMin" type="number" min="1" max="720"/>
            </div>
          </div>
          <div class="cfg-actions">
            <button onclick="saveSettings()">Salvar configuração</button>
            <input id="last" type="text" readonly value="-" style="min-width:260px"/>
          </div>
        </div>
      </article>
      <article class="card auth-card">
        <h3>Autenticação</h3>
        <div class="btns">
          <button id="reauthBtn" class="sec" onclick="reauth('principal')">Principal</button>
        </div>
      </article>
      <article class="card reproc-card">
        <h3>Reprocessar e-mails</h3>
        <div class="reproc-grid">
          <div>
            <label>Conta</label>
            <select id="account">
              <option value="all">Todos</option>
              <option value="principal">E-mail Principal</option>
            </select>
          </div>
          <div>
            <label>Dias para trás</label>
            <input id="days" type="number" value="30" min="1" max="365"/>
          </div>
          <div>
            <label>Limite de mensagens</label>
            <input id="limit" type="number" value="100" min="1" max="1000"/>
          </div>
        </div>
        <label class="cb"><input id="unread" type="checkbox" checked/>Marcar como não lido</label>
        <div class="btns">
          <button onclick="reprocess()">Remover labels para reprocessar</button>
          <button class="sec" onclick="runNow()">Executar agora</button>
        </div>
      </article>
    </section>
  </section>

  <section id="tabHist" class="hidden">
    <section class="card" style="margin-top:10px">
      <h3>Histórico</h3>
      <div class="btns">
        <input id="hQuery" class="inp" type="text" placeholder="Buscar no histórico">
        <input id="hLimit" class="inp" type="number" min="10" max="2000" value="300" placeholder="Limite">
        <button onclick="loadHistory()">Aplicar filtros</button>
      </div>
      <pre id="historyList" style="margin-top:8px">Carregando...</pre>
    </section>
  </section>

  <section id="tabDiag" class="hidden">
    <section class="card" style="margin-top:10px">
      <h3>Diagnóstico</h3>
      <pre id="details">-</pre>
    </section>
  </section>
</main>
<script>
async function api(path,opts){const r=await fetch(path,opts);const j=await r.json().catch(()=>({}));if(r.status===401){window.location.href='/login';throw new Error('nao autenticado');}return j;}
let _nextRemain=0;
function _fmtSec(total){
  const t=Math.max(0, Number(total||0));
  const h=Math.floor(t/3600);
  const m=Math.floor((t%3600)/60);
  const s=Math.floor(t%60);
  if(h>0) return `${h}h ${m}min ${s}s`;
  if(m>0) return `${m}min ${s}s`;
  return `${s}s`;
}
function _tickNext(){
  const el=document.getElementById('cool');
  if(!el) return;
  if(_nextRemain>0){
    el.textContent='Próxima verificação automática em '+_fmtSec(_nextRemain);
    _nextRemain=Math.max(0,_nextRemain-1);
  }else{
    el.textContent='Próxima verificação automática: sem contagem no momento';
  }
}
function switchTab(tab){
  const m=tab==='main';
  const h=tab==='hist';
  const d=tab==='diag';
  document.getElementById('tabMain').classList.toggle('hidden',!m);
  document.getElementById('tabHist').classList.toggle('hidden',!h);
  document.getElementById('tabDiag').classList.toggle('hidden',!d);
  document.getElementById('tabBtnMain').classList.toggle('active',m);
  document.getElementById('tabBtnHist').classList.toggle('active',h);
  document.getElementById('tabBtnDiag').classList.toggle('active',d);
}
function setPill(ok,running){const p=document.getElementById('pill');if(running){p.className='status-pill ok';p.innerHTML='<span>●</span><span>Em execução</span>';return;}if(ok){p.className='status-pill off';p.innerHTML='<span>●</span><span>Aguardando</span>';return;}p.className='status-pill err';p.innerHTML='<span>●</span><span>Com erro</span>';}
function setAccBadge(kind,label){
  const b=document.getElementById('accBadge');
  b.className='status-pill '+kind;
  b.innerHTML='<span>●</span><span>'+label+'</span>';
}
function updAccount(state){
  const e=document.getElementById('accEmail');
  const d=document.getElementById('accDetail');
  const p=document.getElementById('accProblem');
  e.textContent='E-mail conectado: '+String(state.email||'-');
  d.textContent=String(state.friendly||'Aguardando leitura');
  const raw=String(state.error||'').trim();
  if(raw){
    p.textContent=raw;
    p.classList.remove('hidden');
  }else{
    p.textContent='-';
    p.classList.add('hidden');
  }
  const st=String(state.status||'waiting');
  if(st==='running'){setAccBadge('ok','Funcionando');return;}
  if(st==='error'){setAccBadge('err','Com problema');return;}
  setAccBadge('off','Aguardando');
}
function setList(id,items){
  const el=document.getElementById(id);
  if(!el) return;
  const arr=Array.isArray(items)?items:[];
  if(!arr.length){el.innerHTML='<li>Sem itens</li>';return;}
  el.innerHTML=arr.map(v=>`<li>${String(v||'').replace(/</g,'&lt;').replace(/>/g,'&gt;')}</li>`).join('');
}
function updDaily(rep){
  const t=(rep&&rep.totals)||{};
  document.getElementById('kp1').textContent=String(t.processados||0);
  document.getElementById('kp2').textContent=String(t.ignorados||0);
  document.getElementById('kp3').textContent=String(t.avisos_ciclo||0);
  document.getElementById('kp4').textContent=String(t.avisos_dia||0);
  const meta=document.getElementById('rmeta');
  if(rep&&rep.exists){
    meta.textContent='Atualizado em: '+String(rep.updated_at||'-')+' | Arquivo: '+String(rep.path||'-');
  }else{
    meta.textContent='Sem relatório encontrado ainda';
  }
  setList('lp',(rep&&rep.processados)||[]);
  setList('li',(rep&&rep.ignorados)||[]);
  setList('la',(rep&&rep.avisos)||[]);
}
async function refresh(){const j=await api('/api/state');const running=!!j.running;const ok=!!(j.last_status&&j.last_status.ok);const s=(j.settings||{});document.getElementById('who').textContent='Usuário: '+String((j.auth&&j.auth.user)||'-');document.getElementById('mode').value=String(s.gmail_filter_mode||'last_30_days');document.getElementById('maxPages').value=String(s.gmail_max_pages||3);document.getElementById('pageSize').value=String(s.gmail_page_size||50);document.getElementById('intervalMin').value=String(s.loop_interval_minutes||30);document.getElementById('last').value=String((j.last_status&&j.last_status.message)||'-');document.getElementById('details').textContent=JSON.stringify(j.last_status||{},null,2);_nextRemain=Number((j.scheduler&&j.scheduler.next_in_seconds)||0);_tickNext();updAccount(j.account||{});setPill(ok,running);updDaily(j.daily_report||{});}
async function startLoop(){await api('/api/start',{method:'POST'});refresh();}
async function stopLoop(){await api('/api/stop',{method:'POST'});refresh();}
async function runNow(){const account=(document.getElementById('account').value||'principal');await api('/api/run-now',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({account})});refresh();}
async function saveSettings(){const payload={gmail_filter_mode:document.getElementById('mode').value,gmail_max_pages:Number(document.getElementById('maxPages').value||3),gmail_page_size:Number(document.getElementById('pageSize').value||50),loop_interval_minutes:Number(document.getElementById('intervalMin').value||30)};await api('/api/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});refresh();}
async function countdown(sec){const ov=document.getElementById('ov');const c=document.getElementById('cnt');let n=Number(sec||5);if(c)c.textContent=String(n);if(ov)ov.classList.add('show');await new Promise((res)=>{const t=setInterval(()=>{n-=1;if(c)c.textContent=String(Math.max(n,0));if(n<=0){clearInterval(t);res();}},1000);});if(ov)ov.classList.remove('show');}
async function reauth(account){
  const btn=document.getElementById('reauthBtn');
  if(btn){btn.disabled=true;btn.textContent='Autenticando...';}
  try{
    await countdown(5);
    await api('/api/reauth',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({account:account||'principal'})});
    await refresh();
  }finally{
    if(btn){btn.disabled=false;btn.textContent='Principal';}
  }
}
async function reprocess(){const payload={account:document.getElementById('account').value||'principal',days:Number(document.getElementById('days').value||30),max_messages:Number(document.getElementById('limit').value||100),mark_unread:document.getElementById('unread').checked};await api('/api/reprocess',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});refresh();}
async function loadHistory(){
  const q=(document.getElementById('hQuery').value||'').trim();
  const l=Number(document.getElementById('hLimit').value||300);
  const p=new URLSearchParams();
  if(q)p.set('q',q);
  p.set('limit',String(Math.max(10,Math.min(2000,l||300))));
  const j=await api('/api/history?'+p.toString());
  const items=j.items||[];
  const lines=items.map(i=>`${i.at||''} - ${i.message||''}`.trim());
  document.getElementById('historyList').textContent=lines.length?lines.join('\\n'):'Sem itens.';
}
async function logout(){await fetch('/api/logout',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'}).catch(()=>{});window.location.href='/login';}
['mode','maxPages','pageSize','intervalMin'].forEach(id=>{const el=document.getElementById(id);if(!el)return;el.addEventListener('keydown',(e)=>{if(e.key==='Enter'){e.preventDefault();saveSettings();}});});
['account','days','limit'].forEach(id=>{const el=document.getElementById(id);if(!el)return;el.addEventListener('keydown',(e)=>{if(e.key==='Enter'){e.preventDefault();reprocess();}});});
refresh();loadHistory();setInterval(refresh,3000);setInterval(_tickNext,1000);
</script></body></html>"""

def start_server(host: str, port: int, no_loop: bool = False):
    _load_auth()
    _load_settings()

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, fmt, *args):
            return

        def _require_auth(self, parsed_path: str) -> str | None:
            user = _current_session_user(self)
            if user:
                return user
            if parsed_path.startswith("/api/"):
                _json_response(self, 401, {"ok": False, "message": "Não autenticado"})
            else:
                _html_response(self, 200, _render_login_html())
            return None

        def do_GET(self):
            parsed = urlparse(self.path)
            if parsed.path == "/assets/store-bg":
                if _send_store_image(self):
                    return
                self.send_response(404)
                self.end_headers()
                return
            if parsed.path == "/login":
                if _current_session_user(self):
                    return _html_response(self, 200, _render_server_html())
                return _html_response(self, 200, _render_login_html())

            user = self._require_auth(parsed.path)
            if not user:
                return

            if parsed.path == "/":
                return _html_response(self, 200, _render_server_html())
            if parsed.path == "/api/state":
                email_info = _connected_email()
                last_msg = str((last_status or {}).get("message", "") or "")
                if running:
                    acc_status = "running"
                    friendly = "Lendo os e-mails agora"
                elif (last_status or {}).get("ok", True):
                    acc_status = "waiting"
                    friendly = "Aguardando a próxima verificação automática"
                else:
                    acc_status = "error"
                    friendly = "Falha na comunicação com a API. Veja os detalhes técnicos para identificar a causa."
                return _json_response(
                    self,
                    200,
                    {
                        "ok": True,
                        "running": bool(running),
                        "interval_seconds": int(_RUNTIME_SETTINGS.get("interval_seconds", INTERVALO)),
                        "max_messages": int(_RUNTIME_SETTINGS.get("max_messages", 100)),
                        "settings": {
                            "gmail_filter_mode": str(_RUNTIME_SETTINGS.get("gmail_filter_mode", "last_30_days")),
                            "gmail_max_pages": int(_RUNTIME_SETTINGS.get("gmail_max_pages", 3)),
                            "gmail_page_size": int(_RUNTIME_SETTINGS.get("gmail_page_size", 50)),
                            "loop_interval_minutes": int(_RUNTIME_SETTINGS.get("loop_interval_minutes", 30)),
                        },
                        "last_status": dict(last_status),
                        "account": {
                            "email": str(email_info.get("email", "")),
                            "status": acc_status,
                            "friendly": friendly,
                            "error": "" if not email_info.get("error") else str(email_info.get("error")),
                            "detail": last_msg,
                        },
                        "scheduler": {
                            "next_in_seconds": max(0, int(_NEXT_RUN_AT - time.time())) if _NEXT_RUN_AT > 0 else int(_RUNTIME_SETTINGS.get("interval_seconds", INTERVALO)),
                        },
                        "daily_report": _daily_report_data(),
                        "auth": {"user": user, "role": _role_of(user)},
                    },
                )
            if parsed.path == "/api/history":
                qs = parse_qs(parsed.query or "")
                query = (qs.get("q", [""])[0] or "").strip()
                try:
                    limit = int((qs.get("limit", ["300"])[0] or "300").strip())
                except Exception:
                    limit = 300
                items = _history_from_reports(limit=max(10, min(limit, 2000)), query=query)
                return _json_response(self, 200, {"items": items})
            return _json_response(self, 404, {"ok": False, "message": "Não encontrado"})

        def do_POST(self):
            try:
                parsed = urlparse(self.path)
                data = _read_json(self)

                if parsed.path == "/api/login":
                    username = str(data.get("username", "")).strip()
                    password = str(data.get("password", ""))
                    if not _verify_login(username, password):
                        return _json_response(self, 401, {"ok": False, "message": "Usuário ou senha inválidos"})
                    token = _create_session(username)
                    cookie = f"{_COOKIE_SESSION}={token}; Path=/; HttpOnly; Max-Age={_SESSION_TTL_SECONDS}; SameSite=Lax"
                    return _json_response(self, 200, {"ok": True, "message": "Login efetuado"}, {"Set-Cookie": cookie})

                if parsed.path == "/api/logout":
                    token = _read_cookie_session(self)
                    if token:
                        with _SESSIONS_LOCK:
                            _SESSIONS.pop(token, None)
                    return _json_response(
                        self,
                        200,
                        {"ok": True},
                        {"Set-Cookie": f"{_COOKIE_SESSION}=; Path=/; HttpOnly; Max-Age=0; SameSite=Lax"},
                    )

                user = self._require_auth(parsed.path)
                if not user:
                    return

                if parsed.path == "/api/start":
                    started = iniciar_verificacao()
                    return _json_response(self, 200, {"ok": True, "started": bool(started)})
                if parsed.path == "/api/stop":
                    stopped = parar_verificacao()
                    return _json_response(self, 200, {"ok": True, "stopped": bool(stopped)})
                if parsed.path == "/api/run-now":
                    _ = str(data.get("account", "principal"))
                    ok, msg = executar_um_ciclo()
                    return _json_response(self, 200 if ok else 500, {"ok": bool(ok), "message": msg})
                if parsed.path == "/api/settings":
                    if not _can_operate(user):
                        return _json_response(self, 403, {"ok": False, "message": "Sem permissão"})
                    _save_settings(
                        {
                            "gmail_filter_mode": data.get("gmail_filter_mode", _RUNTIME_SETTINGS.get("gmail_filter_mode", "last_30_days")),
                            "gmail_max_pages": data.get("gmail_max_pages", _RUNTIME_SETTINGS.get("gmail_max_pages", 3)),
                            "gmail_page_size": data.get("gmail_page_size", _RUNTIME_SETTINGS.get("gmail_page_size", 50)),
                            "loop_interval_minutes": data.get("loop_interval_minutes", _RUNTIME_SETTINGS.get("loop_interval_minutes", 30)),
                            "max_messages": data.get("max_messages", _RUNTIME_SETTINGS.get("max_messages", 100)),
                        }
                    )
                    return _json_response(self, 200, {"ok": True, "message": "Configuração salva"})
                if parsed.path == "/api/reprocess":
                    if not _can_operate(user):
                        return _json_response(self, 403, {"ok": False, "message": "Sem permissão"})
                    try:
                        days = max(1, min(365, int(data.get("days", 30))))
                    except Exception:
                        days = 30
                    try:
                        max_messages = max(1, min(1000, int(data.get("max_messages", 100))))
                    except Exception:
                        max_messages = 100
                    mark_unread = bool(data.get("mark_unread", True))
                    result = _reprocess_recent(days=days, max_messages=max_messages, mark_unread=mark_unread)
                    friendly = f"Reprocessamento concluído: {result.get('changed', 0)} de {result.get('matched', 0)} mensagens atualizadas"
                    return _json_response(self, 200, {"ok": True, "result": result, "friendly": friendly})
                if parsed.path == "/api/reauth":
                    if not _can_operate(user):
                        return _json_response(self, 403, {"ok": False, "message": "Sem permissão"})
                    account = str(data.get("account", "principal")).strip().lower()
                    if account != "principal":
                        account = "principal"
                    info = _reauthenticate_gmail()
                    return _json_response(self, 200, {"ok": True, "account": account, "message": info.get("message", "Reautenticação concluída"), "friendly": "Reautenticação concluída"})
                return _json_response(self, 404, {"ok": False, "message": "Não encontrado"})
            except Exception as exc:
                logger.exception("Erro no endpoint POST %s: %s", self.path, exc)
                return _json_response(self, 500, {"ok": False, "message": "Erro interno no servidor"})

    if not no_loop:
        iniciar_verificacao()
    server = ThreadingHTTPServer((host, port), Handler)
    print(f"[Botana] Painel online em http://{host}:{port}")
    print("Ctrl+C para encerrar")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
        parar_verificacao()


def parse_args():
    p = argparse.ArgumentParser(description="Botana")
    p.add_argument("--server", action="store_true", help="Executa em modo servidor HTTP")
    p.add_argument("--host", default="127.0.0.1")
    p.add_argument("--port", type=int, default=8865)
    p.add_argument("--no-loop", action="store_true", help="NÃ£o inicia o loop automaticamente no modo servidor")
    return p.parse_args()


# =========================
# EXECUÃ‡ÃƒO PRINCIPAL
# =========================
if __name__ == "__main__":
    try:
        os.system("cls" if os.name == "nt" else "clear")
    except Exception:
        pass
    args = parse_args()
    if args.server:
        start_server(args.host, args.port, no_loop=args.no_loop)
    else:
        if run_tray is None:
            # fallback para ambientes sem tray (ex.: servidor sem interface)
            start_server("127.0.0.1", 8865, no_loop=False)
        else:
            run_tray(on_quit_callback=on_quit, start_callback=iniciar_verificacao)












