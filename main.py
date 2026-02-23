import argparse
import hashlib
import hmac
import json
import secrets
import os, re, time, gspread, threading, sys
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse, parse_qs
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
_RUNTIME_SETTINGS = {"interval_seconds": int(INTERVALO), "max_messages": 100}
_EMAIL_CACHE = {"email": "", "error": "", "at": 0.0}


def _load_settings():
    with _SETTINGS_LOCK:
        if not _SETTINGS_FILE.exists():
            out = {
                "interval_seconds": max(30, min(86400, int(_RUNTIME_SETTINGS.get("interval_seconds", INTERVALO)))),
                "max_messages": max(1, min(1000, int(_RUNTIME_SETTINGS.get("max_messages", 100)))),
            }
            _RUNTIME_SETTINGS.update(out)
            _SETTINGS_FILE.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
            return dict(out)
        try:
            raw = json.loads(_SETTINGS_FILE.read_text(encoding="utf-8"))
        except Exception:
            raw = {}
        out = {
            "interval_seconds": max(30, min(86400, int(raw.get("interval_seconds", INTERVALO)))),
            "max_messages": max(1, min(1000, int(raw.get("max_messages", 100)))),
        }
        _RUNTIME_SETTINGS.update(out)
        return out


def _save_settings(data: dict):
    with _SETTINGS_LOCK:
        out = {
            "interval_seconds": max(30, min(86400, int(data.get("interval_seconds", INTERVALO)))),
            "max_messages": max(1, min(1000, int(data.get("max_messages", 100)))),
        }
        _RUNTIME_SETTINGS.update(out)
        _SETTINGS_FILE.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")


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
    service = getGmailService()
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
        if stop_event.wait(int(_RUNTIME_SETTINGS.get("interval_seconds", INTERVALO))):
            break
    running = False
    logger.info("[Botana] Loop finalizado")


def executar_um_ciclo():
    global last_status
    try:
        processar_emails_enviados()
        last_status = {"ok": True, "message": "ExecuÃ§Ã£o manual concluÃ­da", "at": datetime.now().isoformat()}
        return True, "ExecuÃ§Ã£o manual concluÃ­da"
    except Exception as exc:
        logger.exception("Erro na execuÃ§Ã£o manual: %s", exc)
        last_status = {"ok": False, "message": f"Erro na execuÃ§Ã£o manual: {exc}", "at": datetime.now().isoformat()}
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
    candidates = [
        Path(__file__).resolve().parent / "assets" / "branding" / "Arte MVA logo Metalico (1).png",
        Path(__file__).resolve().parent / "assets" / "branding" / "arte mva logo metalico (1).png",
        Path(__file__).resolve().parent / "Arte MVA logo Metalico (1).png",
        Path(__file__).resolve().parent / "arte mva logo metalico (1).png",
        Path.home() / "Desktop" / "Arte MVA logo Metalico (1).png",
    ]
    for p in candidates:
        if p.exists():
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


def _connected_email(force: bool = False) -> dict:
    now = time.time()
    if not force and _EMAIL_CACHE.get("at", 0.0) and (now - float(_EMAIL_CACHE.get("at", 0.0)) < 120):
        return dict(_EMAIL_CACHE)
    try:
        service = getGmailService()
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
body{margin:0;font-family:'Lexend',Arial,sans-serif;background:linear-gradient(160deg,rgba(41,22,11,.78),rgba(95,56,28,.72)),url('/assets/store-bg') center/cover fixed;color:#2c1b12;padding:12px}
.app{max-width:1120px;margin:0 auto}
.top{display:flex;align-items:center;justify-content:space-between;background:linear-gradient(100deg,var(--b),#7a4d30);color:#fff9f3;border-radius:14px;padding:12px 14px;border:1px solid rgba(255,235,214,.3)}
.brand{font-weight:800;letter-spacing:.2px}
.status-pill{display:inline-flex;align-items:center;gap:6px;padding:4px 10px;border-radius:999px;font-size:.82rem;font-weight:700;border:1px solid}
.ok{background:#e8f6ea;color:#2e7d32;border-color:#b6dfbf}
.off{background:#fff3e0;color:#8b4f19;border-color:#f2c8a3}
.err{background:#fdecec;color:#b42b2b;border-color:#f1bbbb}
.tabs{display:flex;gap:8px;padding:10px 2px 0}
.tab-btn{background:#f2e5d6;color:#55321c;border:1px solid #d9b690;border-radius:10px;padding:6px 12px;font-weight:700;cursor:pointer}
.tab-btn.active{background:linear-gradient(90deg,var(--o),var(--o2));border-color:transparent;color:#2b1408}
.hidden{display:none}
.grid{display:grid;grid-template-columns:1.1fr .9fr;gap:10px;margin-top:10px}
.card{background:rgba(255,250,245,.9);border:1px solid var(--br);border-radius:14px;padding:12px;box-shadow:0 8px 20px rgba(21,11,6,.06)}
h3{margin:0 0 10px;color:#5a311b}
.muted{color:#6d4a35}
.btns{display:flex;gap:8px;flex-wrap:wrap}
button{padding:9px 12px;border:0;border-radius:9px;background:linear-gradient(90deg,var(--o),var(--o2));color:#2b1408;font-weight:700;cursor:pointer}
button.sec{background:linear-gradient(90deg,#7a4d30,#5b341f);color:#fff9f3}
button.warn{background:linear-gradient(90deg,#bc2d2d,#8f2020);color:#fff}
.inp{padding:8px;border:1px solid #d6b18f;border-radius:8px;background:#fffdfb;width:145px}
pre{margin:0;background:#fff7ef;border:1px dashed #cf9f78;padding:10px;border-radius:10px;overflow:auto;max-height:340px;white-space:pre-wrap}
.kpi{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px}
.k{border:1px solid #e8c9aa;border-radius:10px;padding:10px;background:#fffdfb}
.k .n{font-size:1.1rem;font-weight:800;color:#6b4128}
.k .t{font-size:.84rem;color:#7c5034}
.status-grid{display:grid;grid-template-columns:1fr;gap:8px}
.s{border:1px solid #d5b08f;background:#fffaf6;border-radius:11px;padding:10px}
.h{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;font-weight:700;color:#5b321c}
.problem{color:#862818;font-size:.82rem;margin-top:4px}
@media(max-width:900px){.grid{grid-template-columns:1fr}}
</style></head><body>
<main class="app">
  <section class="top">
    <div class="brand">Botana - Painel de Controle</div>
    <div style="display:flex;align-items:center;gap:10px">
      <span id="who">Usuário: -</span>
      <button class="sec" style="padding:6px 10px" onclick="logout()">Sair</button>
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
      <h3>Status das contas de e-mail</h3>
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

    <div class="grid">
      <article class="card">
        <h3>Status da execução</h3>
        <div id="status" class="muted">Carregando status...</div>
        <div class="btns" style="margin-top:10px">
          <button onclick="startLoop()">Iniciar loop</button>
          <button class="sec" onclick="runNow()">Executar agora</button>
          <button class="warn" onclick="stopLoop()">Parar loop</button>
        </div>
      </article>
      <article class="card">
        <h3>Resumo</h3>
        <div class="kpi">
          <div class="k"><div id="maxMsgs" class="n">-</div><div class="t">Máx. mensagens</div></div>
          <div class="k"><div id="interval" class="n">-</div><div class="t">Intervalo (s)</div></div>
          <div class="k"><div id="stateTxt" class="n">-</div><div class="t">Estado</div></div>
        </div>
      </article>
    </div>

    <section class="card" style="margin-top:10px">
      <h3>Configuração do Botana</h3>
      <div class="btns">
        <input id="cfgInterval" class="inp" type="number" min="30" max="86400" placeholder="Intervalo (s)">
        <input id="cfgMax" class="inp" type="number" min="1" max="1000" placeholder="Máx. mensagens">
        <button onclick="saveSettings()">Salvar configuração</button>
      </div>
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
async function refresh(){const j=await api('/api/state');const running=!!j.running;const ok=!!(j.last_status&&j.last_status.ok);document.getElementById('who').textContent='Usuário: '+String((j.auth&&j.auth.user)||'-');document.getElementById('status').textContent='Loop: '+(running?'ativo':'parado')+' | Intervalo: '+j.interval_seconds+' segundos';document.getElementById('interval').textContent=String(j.interval_seconds||'-');document.getElementById('maxMsgs').textContent=String(j.max_messages||'-');document.getElementById('stateTxt').textContent=running?'Ativo':'Parado';document.getElementById('cfgInterval').value=String(j.interval_seconds||'');document.getElementById('cfgMax').value=String(j.max_messages||'');document.getElementById('details').textContent=JSON.stringify(j.last_status||{},null,2);const left=Number((j.scheduler&&j.scheduler.next_in_seconds)||0);document.getElementById('cool').textContent=left>0?('Próxima verificação automática em '+left+'s'):'Próxima verificação automática: sem contagem no momento';updAccount(j.account||{});setPill(ok,running);}
async function startLoop(){await api('/api/start',{method:'POST'});refresh();}
async function stopLoop(){await api('/api/stop',{method:'POST'});refresh();}
async function runNow(){await api('/api/run-now',{method:'POST'});refresh();}
async function saveSettings(){const interval_seconds=Number(document.getElementById('cfgInterval').value||0);const max_messages=Number(document.getElementById('cfgMax').value||0);await api('/api/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({interval_seconds,max_messages})});refresh();}
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
refresh();loadHistory();setInterval(refresh,3000);
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
                        "last_status": dict(last_status),
                        "account": {
                            "email": str(email_info.get("email", "")),
                            "status": acc_status,
                            "friendly": friendly,
                            "error": "" if not email_info.get("error") else str(email_info.get("error")),
                            "detail": last_msg,
                        },
                        "scheduler": {
                            "next_in_seconds": int(_RUNTIME_SETTINGS.get("interval_seconds", INTERVALO)),
                        },
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
                    ok, msg = executar_um_ciclo()
                    return _json_response(self, 200 if ok else 500, {"ok": bool(ok), "message": msg})
                if parsed.path == "/api/settings":
                    if not _can_operate(user):
                        return _json_response(self, 403, {"ok": False, "message": "Sem permissão"})
                    _save_settings(
                        {
                            "interval_seconds": data.get("interval_seconds", _RUNTIME_SETTINGS.get("interval_seconds", INTERVALO)),
                            "max_messages": data.get("max_messages", _RUNTIME_SETTINGS.get("max_messages", 100)),
                        }
                    )
                    return _json_response(self, 200, {"ok": True, "message": "Configuração salva"})
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












