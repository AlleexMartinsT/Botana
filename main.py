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
from config import PLANILHAS, CNPJ_MVA, CNPJ_EH, INTERVALO, DOWNLOAD_DIR, RELATORIO_DIR, GOOGLE_CREDENTIALS_SHEETS, GOOGLE_CREDENTIALS_GMAIL
from gmail_service import getGmailService, buscarMessagesEnviadosPagina, baixar_anexos_de_mensagem, ensure_label, LABEL_NAME
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
# SKIP_UNTIL_NF = "19843"

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
_EMAIL_CACHE = {"email": "", "error": "", "pending": False, "at": 0.0}
_NEXT_RUN_AT = 0.0
_GMAIL_SERVICE_LOCK = threading.Lock()
_IS_READING = False
_PROCESS_STATS_LOCK = threading.Lock()
_PROCESS_STATS = {
    "current": {
        "active": False,
        "started_at": "",
        "messages": 0,
        "attachments": 0,
        "xmls": 0,
        "launched": 0,
    },
    "last": {
        "ok": None,
        "started_at": "",
        "finished_at": "",
        "messages": 0,
        "attachments": 0,
        "xmls": 0,
        "launched": 0,
        "error": "",
    },
}


def _process_start():
    now = datetime.now().isoformat()
    with _PROCESS_STATS_LOCK:
        _PROCESS_STATS["current"] = {
            "active": True,
            "started_at": now,
            "messages": 0,
            "attachments": 0,
            "xmls": 0,
            "launched": 0,
        }


def _process_update(messages: int | None = None, attachments: int | None = None, xmls: int | None = None, launched: int | None = None):
    with _PROCESS_STATS_LOCK:
        cur = _PROCESS_STATS.get("current", {})
        if messages is not None:
            cur["messages"] = max(0, int(messages))
        if attachments is not None:
            cur["attachments"] = max(0, int(attachments))
        if xmls is not None:
            cur["xmls"] = max(0, int(xmls))
        if launched is not None:
            cur["launched"] = max(0, int(launched))


def _process_finish(ok: bool, error: str = ""):
    now = datetime.now().isoformat()
    with _PROCESS_STATS_LOCK:
        cur = dict(_PROCESS_STATS.get("current", {}))
        _PROCESS_STATS["last"] = {
            "ok": bool(ok),
            "started_at": str(cur.get("started_at", "") or ""),
            "finished_at": now,
            "messages": int(cur.get("messages", 0) or 0),
            "attachments": int(cur.get("attachments", 0) or 0),
            "xmls": int(cur.get("xmls", 0) or 0),
            "launched": int(cur.get("launched", 0) or 0),
            "error": str(error or ""),
        }
        cur["active"] = False
        _PROCESS_STATS["current"] = cur


def _reading_state() -> dict:
    """
    Estado consistente da leitura, com auto-correcao de divergencia entre flag e ciclo ativo.
    """
    global _IS_READING
    with _PROCESS_STATS_LOCK:
        cur = dict(_PROCESS_STATS.get("current", {}))
    current_active = bool(cur.get("active"))
    flag_raw = bool(_IS_READING)
    corrected = False
    source = "idle"

    if flag_raw and not current_active:
        _IS_READING = False
        corrected = True
        source = "stale_flag_corrected"
    elif current_active and not flag_raw:
        _IS_READING = True
        corrected = True
        source = "synced_from_current"
    elif current_active and flag_raw:
        source = "flag_and_current"

    reading_now = bool(_IS_READING) and current_active
    if not reading_now and source == "idle" and current_active:
        source = "current_only"

    return {
        "reading": bool(reading_now),
        "source": source,
        "flag_raw": flag_raw,
        "current_active": current_active,
        "corrected": corrected,
    }


def _reading_active() -> bool:
    return bool(_reading_state().get("reading"))


def _process_snapshot() -> dict:
    r = _reading_state()
    with _PROCESS_STATS_LOCK:
        cur = dict(_PROCESS_STATS.get("current", {}))
        last = dict(_PROCESS_STATS.get("last", {}))
    return {
        "running": bool(running),
        "reading": bool(r.get("reading", False)),
        "reading_diag": r,
        "current": cur,
        "last": last,
    }


def _get_gmail_service_locked(timeout: float | None = None):
    if timeout is None:
        with _GMAIL_SERVICE_LOCK:
            return getGmailService()
    acquired = _GMAIL_SERVICE_LOCK.acquire(timeout=max(0.01, float(timeout)))
    if not acquired:
        raise TimeoutError("AUTH_IN_PROGRESS")
    try:
        return getGmailService()
    finally:
        _GMAIL_SERVICE_LOCK.release()


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
    if username == "hub_internal":
        return True
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


def _normalize_report_text(text: str) -> str:
    out = str(text or "")
    markers = ("Ãƒ", "Ã‚", "Ã¢", "Ã°Å¸", "ï¿½", "Ñ€ÑŸ")
    if any(m in out for m in markers):
        for enc in ("latin-1", "cp1252", "cp1251"):
            try:
                fixed = out.encode(enc).decode("utf-8")
                if fixed:
                    out = fixed
                    break
            except Exception:
                continue
    replacements = {
        "ÃƒÂ¡": "Ã¡",
        "ÃƒÂ¢": "Ã¢",
        "ÃƒÂ£": "Ã£",
        "Ãƒ ": "Ã ",
        "ÃƒÂ©": "Ã©",
        "ÃƒÂª": "Ãª",
        "ÃƒÂ­": "Ã­",
        "ÃƒÂ³": "Ã³",
        "ÃƒÂ´": "Ã´",
        "ÃƒÂµ": "Ãµ",
        "ÃƒÂº": "Ãº",
        "ÃƒÂ§": "Ã§",
        "Ã‚ ": "",
        "Ã¢â‚¬â€œ": "â€“",
        "Ã¢â‚¬â€": "â€”",
        "Ã¢â‚¬Å“": "â€œ",
        "Ã¢â‚¬Â": "â€",
        "Ã¢â‚¬Ëœ": "â€˜",
        "Ã¢â‚¬â„¢": "â€™",
    }
    for src, dst in replacements.items():
        out = out.replace(src, dst)
    # Remove emojis/symbols pictograficos para manter o painel textual limpo.
    out = re.sub(r"[\u200d\ufe0f]", "", out)
    out = re.sub(
        r"[\U0001F1E6-\U0001F1FF\U0001F300-\U0001FAFF\U00002700-\U000027BF\U00002600-\U000026FF]",
        "",
        out,
    )
    out = re.sub(r"\s{2,}", " ", out)
    return out.strip()


def _write_history_launch_event(dados_xml: dict, dados_parcela: dict, result: dict):
    """Registra no relatorio um evento estruturado apenas para lancamentos validos."""
    try:
        if not isinstance(result, dict) or not bool(result.get("inserted")):
            return
        nf = str(dados_parcela.get("nf") or dados_xml.get("nf") or "").strip()
        if not nf:
            return
        payload = {
            "type": "boleto_lancado",
            "at": _now(),
            "nf": nf,
            "cliente": str(dados_parcela.get("destinatario") or dados_xml.get("destinatario") or "").strip(),
            "cnpj_emit": re.sub(r"\D+", "", str(dados_xml.get("cnpjEmitente") or "")),
            "cnpj_dest": re.sub(r"\D+", "", str(dados_xml.get("cnpjDestinatario") or "")),
            "vencimento": str(result.get("vencimento") or dados_parcela.get("vencimento") or "").strip(),
            "descricao": str(result.get("descricao") or dados_parcela.get("descricao") or "").strip(),
            "valor_total": float(result.get("valor_total", dados_parcela.get("valorTotal", 0)) or 0),
            "qtd_parcelas": int(result.get("qtd_parcelas", dados_parcela.get("qtdParcelas", 1)) or 1),
            "parcela": str(result.get("parcela") or dados_parcela.get("numParcela") or "").strip(),
            "valor_parcela": float(result.get("valor_parcela", dados_parcela.get("valorParcela", 0)) or 0),
            "valor_pago": str(result.get("valor_pago") or "").strip(),
            "status": str(result.get("status") or "").strip(),
            "sheet_title": str(result.get("sheet_title") or "").strip(),
            "sheet_type": str(result.get("sheet_type") or "").strip(),
            "aba": str(result.get("aba") or "").strip(),
        }
        escreverRelatorio(f"{_now()} - HIST_JSON {json.dumps(payload, ensure_ascii=False)}")
    except Exception as exc:
        logger.warning("Falha ao registrar evento estruturado no historico: %s", exc)

def processar_emails_enviados():
    global _IS_READING
    _process_start()
    service = _get_gmail_service_locked()
    batch_size = int(_RUNTIME_SETTINGS.get("max_messages", 100))
    total_msgs = 0
    anexos_lidos = 0
    xmls_lidos = 0
    total_processados = 0
    page_token = None
    primeira_pagina = True
    interativo_cmd = bool(sys.stdin and sys.stdin.isatty())
    def _summary() -> dict:
        return {
            "messages": int(total_msgs),
            "attachments": int(anexos_lidos),
            "xmls": int(xmls_lidos),
            "launched": int(total_processados),
        }

    def _sync_progress():
        _process_update(
            messages=total_msgs,
            attachments=anexos_lidos,
            xmls=xmls_lidos,
            launched=total_processados,
        )

    while True:
        msgs, next_page_token = buscarMessagesEnviadosPagina(
            service,
            max_results=batch_size,
            page_token=page_token,
        )

        if primeira_pagina and not msgs:
            logger.info("Nenhuma mensagem enviada com XML encontrada.")
            escreverRelatorio(f"{_now()} - CICLO: 0 e-mails lidos, 0 anexos, 0 XML, 0 lanÃ§amentos.")
            _sync_progress()
            return _summary()

        primeira_pagina = False

        for m in msgs:
            total_msgs += 1
            _sync_progress()
            msg_id = m.get("id")
            logger.info("Abrindo mensagem ID: %s", msg_id)

            arquivos = baixar_anexos_de_mensagem(service, msg_id)
            if not arquivos:
                logger.info("Nenhum anexo salvo para mensagem %s", msg_id)
                continue
            anexos_lidos += len(arquivos)
            _sync_progress()

            dados_xmls = []
            boletos = []

            # Processa todos os anexos baixados
            for arquivo in arquivos:
                nome_arquivo = os.path.basename(arquivo)

                try:
                    # =============================
                    # XML -> extrai dados
                    # =============================
                    if arquivo.lower().endswith(".xml"):
                        xmls_lidos += 1
                        _sync_progress()
                        try:
                            dados = extrairDadosXML(arquivo)
                            # Ignora vendas Ã  vista
                            nat_op = dados.get("naturezaOperacao", "").strip().upper()
                            dest_nome = dados.get("destinatario", "")
                            dest_cnpj = re.sub(r"\D+", "", str(dados.get("cnpjDestinatario") or ""))
                            if ( "VISTA" in nat_op or "VENDA A VISTA" in nat_op):
                                # Checa se a mensagem jÃ¡ foi processada no relatÃ³rio atual:
                                if dados.get('nf') not in consolidarRelatorioTMP():
                                    escreverRelatorio(f"{_now()} - NF {dados.get('nf')} ignorada (venda Ã  vista).")
                                    continue
                                else: logger.info(f"{cor_ciano}NF {dados['nf']} jÃ¡ registrada no relatÃ³rio, nÃ£o duplicando a mensagem de ignorada.{reset}")
                                continue
                            cnpj_mva = re.sub(r"\D+", "", str(CNPJ_MVA or ""))
                            cnpj_eh = re.sub(r"\D+", "", str(CNPJ_EH or ""))
                            if dest_cnpj and (dest_cnpj == cnpj_mva or dest_cnpj == cnpj_eh):
                                logger.info(f"[DEBUG IGNORE RESULT] NF {dados['nf']} ignorada (destinatÃ¡rio Ã© o nosso: {dest_nome} / {dest_cnpj})")
                                escreverRelatorio(f"{_now()} - NF {dados.get('nf')} ignorada (destinatÃ¡rio Ã© o nosso).")
                                continue
                            if not dados:
                                motivo = dados.get("motivo_ignoracao", "Desconhecido") if isinstance(dados, dict) else "Desconhecido"
                                logger.info(f"Ignorado XML (motivo: {motivo}).")
                                escreverRelatorio(f"{_now()} - XML {nome_arquivo} ignorado (motivo: {motivo})")
                                continue

                            dados_xmls.append(dados)

                        except Exception as e:
                            escreverRelatorio(f"{_now()} - Erro extraindo XML {nome_arquivo}: {e}")
                            logger.exception("Erro extraindo XML %s: %s", arquivo, e)

                    # =============================
                    # PDF -> tenta identificar boleto
                    # =============================
                    elif arquivo.lower().endswith(".pdf"): # mudar pra elif se o bloco de cima for realmente necessÃ¡rio
                        nome_upper = nome_arquivo.upper()

                        # Trata nomes parecidos com BOLETO (erros comuns tipo BOLTO, BOLETA, BOLETT, etc)
                        padrao_boleto = r"[_\s-]?(BLT|BOLET[OA]?|BOLTO|BOLETOO|BOLETT?)"

                        if re.search(padrao_boleto, nome_upper):
                            match = re.findall(r"([0-9]{2,}-?[0-9]+)", nome_upper)
                            if match:
                                num_boleto = match[-1]
                                m_clean = re.search(r'0{4,}([1-9][0-9]*(-[0-9a-zA-Z]+)?)$', num_boleto)
                                if m_clean:
                                    num_boleto = m_clean.group(1)
                                if num_boleto == "0136" or num_boleto == "136": num_boleto = "10136"
                                elif num_boleto.startswith("0136-"): num_boleto = num_boleto.replace("0136-", "10136-", 1)
                                elif num_boleto.startswith("136-"): num_boleto = num_boleto.replace("136-", "10136-", 1)
                                boletos.append(num_boleto)
                                logger.info("Boleto identificado no nome: %s (BLT %s)", nome_arquivo, num_boleto)
                            else:
                                logger.info("Nenhum nÃºmero de boleto encontrado no nome: %s", nome_arquivo)
                        elif arquivo.lower().endswith(".pdf"):
                            nome_upper = nome_arquivo.upper()

                            # Palavras que indicam boleto (considera erros comuns)
                            padrao_boleto = r"\b(BOLET[OA]?|BOLTO|BOLETOO|BOLETT?|BLT)\b"

                            # SÃ³ tenta identificar nÃºmero se o nome realmente tiver algo prÃ³ximo de "boleto"
                            if re.search(padrao_boleto, nome_upper):
                                match = re.findall(r"([0-9]{2,}-?[0-9]+)", nome_upper)
                                if match:
                                    num_boleto = match[-1]
                                    m_clean = re.search(r'0{4,}([1-9][0-9]*(-[0-9a-zA-Z]+)?)$', num_boleto)
                                    if m_clean:
                                        num_boleto = m_clean.group(1)
                                    if num_boleto == "0136" or num_boleto == "136": num_boleto = "10136"
                                    elif num_boleto.startswith("0136-"): num_boleto = num_boleto.replace("0136-", "10136-", 1)
                                    elif num_boleto.startswith("136-"): num_boleto = num_boleto.replace("136-", "10136-", 1)
                                    boletos.append(num_boleto)
                                    logger.info("Boleto identificado no nome: %s (BLT %s)", nome_arquivo, num_boleto)
                                else:
                                    logger.info("PossÃ­vel boleto sem nÃºmero identificado: %s", nome_arquivo)
                            else:
                                logger.info("PDF ignorado (nÃ£o parece boleto): %s", nome_arquivo)

                    else:
                        logger.info("Arquivo nÃ£o identificado como boleto: %s", nome_arquivo)

                finally:
                    # Remove sempre o anexo local (independente do tipo)
                    try:
                        os.remove(arquivo)
                        logger.debug(f"Anexo removido: {arquivo}")
                    except FileNotFoundError:
                        pass
                    except Exception as e:
                        logger.warning(f"Falha ao remover {arquivo}: {e}")

            # =============================
            # Marca o e-mail como processado
            # =============================
            try:
                marcar_mensagem_com_label(service, msg_id)
                logger.info("E-mail %s marcado com 'XML Processado Botana'", msg_id)
            except Exception as e:
                logger.exception("Falha ao aplicar rÃ³tulo: %s", e)

            # Nenhum XML -> pula este e-mail
            if not dados_xmls:
                logger.info("Nenhum XML vÃ¡lido encontrado neste e-mail.")
                continue

            # =============================
            # Atualiza planilhas
            # =============================
            for dados_xml in dados_xmls:
                # --- FILTRAGEM POR NF (para debug/anÃ¡lise isolada) ---
                nf_num = str(dados_xml.get("nf", "")).strip()

                # NF_ALVO: processa somente essa NF (ignora as outras)
                if NF_ALVO:
                    if nf_num != str(NF_ALVO):
                        logger.info(f"Pulando NF {nf_num} (NF_ALVO ativo: {NF_ALVO})")
                        continue
                    else:
                        logger.info(f"NF_ALVO encontrada: {nf_num}")

                # SKIP_UNTIL_NF: ignora tudo atÃ© encontrar essa NF; quando encontrada, passa a processar normalmente
                if SKIP_UNTIL_NF:
                    # usa atributo da funÃ§Ã£o para manter estado entre ciclos enquanto o processo estÃ¡ vivo
                    if not hasattr(processar_emails_enviados, "_skip_reached"):
                        processar_emails_enviados._skip_reached = False

                    if not processar_emails_enviados._skip_reached:
                        if nf_num == str(SKIP_UNTIL_NF):
                            processar_emails_enviados._skip_reached = True
                            logger.info(f"SKIP_UNTIL_NF: NF {nf_num} encontrada - a partir daqui serÃ¡ processada.")
                        else:
                            logger.info(f"SKIP_UNTIL_NF ativo, pulando NF {nf_num}")
                            continue

                # Se chegou atÃ© aqui, a NF serÃ¡ processada normalmente.
                # Se NF_ALVO + STOP_AFTER_NF: apÃ³s processar, se encerra o loop/principal para anÃ¡lise isolada.

                cnpj_emit = re.sub(r"\D+", "", str(dados_xml.get("cnpjEmitente") or ""))
                ano = dados_xml.get("anoVencimento")
                planilha_id = escolher_planilha_por_cnpj_e_ano(cnpj_emit, ano)

                if not planilha_id:
                    logger.warning("CNPJ %s ou ano %s sem planilha configurada.", cnpj_emit, ano)
                    continue

                # Itera sobre todas as parcelas - mapeamento correto de boletos -> parcelas
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
                        logger.info("Mais boletos (%d) que parcelas (%d). Sobraram: %s", n_boletos, n_parcelas, boletos[n_parcelas:])

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
                        if cnpj_emit == "18471209000107":
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
                            resultado = atualizarPlanilha(planilha, dados_parcela, gc)
                            if isinstance(resultado, dict) and bool(resultado.get("inserted")):
                                total_processados += 1
                                _write_history_launch_event(dados_xml, dados_parcela, resultado)
                                _sync_progress()
                            # Se NF_ALVO + STOP_AFTER_NF -> encerra o processo principal para anÃ¡lise isolada.
                            if NF_ALVO and STOP_AFTER_NF and isinstance(resultado, dict) and bool(resultado.get("inserted")):
                                logger.info(f"NF_ALVO {NF_ALVO} processada. STOP_AFTER_NF=True -> encerrando execuÃ§Ã£o.")
                                # forÃ§a saÃ­da limpa do loop principal retornando da funÃ§Ã£o
                                return _summary()
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

        skip_reached = bool(getattr(processar_emails_enviados, "_skip_reached", False))
        if SKIP_UNTIL_NF and not skip_reached and next_page_token:
            if interativo_cmd:
                resposta = input(
                    f"\nNF {SKIP_UNTIL_NF} nÃ£o encontrada neste lote de {batch_size}. "
                    f"Deseja continuar com mais {batch_size}? [s/N]: "
                ).strip().lower()
                if resposta in ("s", "sim", "y", "yes"):
                    page_token = next_page_token
                    continue
                logger.info("Busca interrompida pelo usuÃ¡rio antes de encontrar a NF %s.", SKIP_UNTIL_NF)
            else:
                logger.info(
                    "NF %s nÃ£o encontrada no lote atual e execuÃ§Ã£o nÃ£o interativa. "
                    "Encerrando sem carregar prÃ³ximas pÃ¡ginas.",
                    SKIP_UNTIL_NF,
                )
        break
    logger.info("Ciclo finalizado. Total processado: %d", total_processados)
    escreverRelatorio(
        f"{_now()} - CICLO: {total_msgs} e-mails lidos, {anexos_lidos} anexos, {xmls_lidos} XML, {total_processados} lanÃ§amentos."
    )
    _sync_progress()
    return _summary()

def main_loop():
    global running, last_status, _NEXT_RUN_AT, _IS_READING
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    running = True
    logger.info("[Botana] Loop iniciado")
    while not stop_event.is_set():
        _NEXT_RUN_AT = time.time() + int(_RUNTIME_SETTINGS.get("interval_seconds", INTERVALO))
        try:
            _IS_READING = True
            summary = processar_emails_enviados()
            _process_finish(ok=True, error="")
            msg = (
                f"Ciclo concluÃ­do: {int((summary or {}).get('messages', 0))} e-mails, "
                f"{int((summary or {}).get('attachments', 0))} anexos, "
                f"{int((summary or {}).get('xmls', 0))} XML, "
                f"{int((summary or {}).get('launched', 0))} lanÃ§amentos."
            )
            last_status = {"ok": True, "message": msg, "at": datetime.now().isoformat()}
        except Exception as e:
            logger.exception("Erro no ciclo principal: %s", e)
            _process_finish(ok=False, error=str(e))
            last_status = {"ok": False, "message": f"Erro no ciclo: {e}", "at": datetime.now().isoformat()}
        finally:
            _IS_READING = False
        if stop_event.wait(int(_RUNTIME_SETTINGS.get("interval_seconds", INTERVALO))):
            break
    running = False
    logger.info("[Botana] Loop finalizado")


def executar_um_ciclo():
    global last_status, _NEXT_RUN_AT, _IS_READING
    try:
        _IS_READING = True
        summary = processar_emails_enviados()
        _process_finish(ok=True, error="")
        msg = (
            f"ExecuÃ§Ã£o manual concluÃ­da: {int((summary or {}).get('messages', 0))} e-mails, "
            f"{int((summary or {}).get('attachments', 0))} anexos, "
            f"{int((summary or {}).get('xmls', 0))} XML, "
            f"{int((summary or {}).get('launched', 0))} lanÃ§amentos."
        )
        last_status = {"ok": True, "message": msg, "at": datetime.now().isoformat()}
        _NEXT_RUN_AT = time.time() + int(_RUNTIME_SETTINGS.get("interval_seconds", INTERVALO))
        return True, msg
    except Exception as exc:
        logger.exception("Erro na execuÃ§Ã£o manual: %s", exc)
        _process_finish(ok=False, error=str(exc))
        last_status = {"ok": False, "message": f"Erro na execuÃ§Ã£o manual: {exc}", "at": datetime.now().isoformat()}
        _NEXT_RUN_AT = time.time() + int(_RUNTIME_SETTINGS.get("interval_seconds", INTERVALO))
        return False, str(exc)
    finally:
        _IS_READING = False


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

    # Fallback para a imagem do FinanceBot, se nÃ£o houver imagem local do Botana.
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



def _gerar_relatorio_nfs(filtro: str, mes: str, nf_inicio: str, nf_fim: str, empresa: str) -> dict:
    """Retorna as NFs faltantes dentro de um intervalo.
    # I collect all existing NF numbers then compare with the expected range.
    # Eu coleto todos os numeros de NF existentes e comparo com o intervalo esperado.
    """
    import gspread
    from google.oauth2.service_account import Credentials
    from config import GOOGLE_CREDENTIALS_SHEETS, PLANILHAS

    creds = Credentials.from_service_account_file(
        GOOGLE_CREDENTIALS_SHEETS,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    gc = gspread.authorize(creds)

    # I collect all numeric NFs found in the sheets, filtering by empresa and optionally by month.
    # Eu coleto todas as NFs numéricas encontradas nas planilhas, filtrando por empresa e opcionalmente por mês.
    nfsEncontradas = set()
    nfsDetalhes = {}  # nf_num -> {Data, Descricao, Planilha, Aba}

    def matchMes(vencimento: str) -> bool:
        """I check if the date string matches the selected month.
        # Eu verifico se a data corresponde ao mes selecionado."""
        if not mes:
            return True
        partes = mes.split("-")
        if len(partes) != 2:
            return True
        m = f"/{partes[1]}/"
        a = partes[0]
        if m in vencimento and (a in vencimento or str(a)[-2:] in vencimento):
            return True
        if mes in vencimento:
            return True
        return False

    for tipoEmpresa, anos in PLANILHAS.items():
        if empresa and empresa != "todos" and tipoEmpresa != empresa:
            continue
        for ano, idPlanilha in anos.items():
            if not idPlanilha:
                continue
            try:
                planilha = gc.open_by_key(idPlanilha)
                for aba in planilha.worksheets():
                    try:
                        linhas = aba.get_all_values()
                        for i, linha in enumerate(linhas):
                            if i == 0 or len(linha) < 3:
                                continue
                            nfTexto = str(linha[2]).strip()
                            if not nfTexto or not nfTexto[0].isdigit():
                                continue
                            try:
                                nfNumero = int(re.sub(r"\D+", "", nfTexto))
                            except ValueError:
                                continue

                            # I apply month filter when relevant.
                            # Eu aplico o filtro de mes quando relevante.
                            if filtro == "mes" and not matchMes(linha[0]):
                                continue

                            nfsEncontradas.add(nfNumero)
                            nfsDetalhes[nfNumero] = {
                                "Data": linha[0],
                                "Descricao": linha[1],
                                "Planilha": tipoEmpresa + " " + (ano or ""),
                                "Aba": aba.title,
                            }
                    except Exception:
                        pass
            except Exception:
                pass

    # I determine the range to check.
    # Eu determino o intervalo para verificar.
    if filtro == "nfs":
        try:
            rangeInicio = int(nf_inicio)
        except (ValueError, TypeError):
            rangeInicio = min(nfsEncontradas) if nfsEncontradas else 0
        try:
            rangeFim = int(nf_fim)
        except (ValueError, TypeError):
            rangeFim = max(nfsEncontradas) if nfsEncontradas else 0
    else:
        # For month or all filters I use min/max from found NFs as the range.
        # Para filtro de mes ou todos eu uso o min/max das NFs encontradas.
        if nfsEncontradas:
            rangeInicio = min(nfsEncontradas)
            rangeFim = max(nfsEncontradas)
        else:
            rangeInicio = 0
            rangeFim = 0

    # I build the list of missing NFs within the range.
    # Eu monto a lista das NFs faltantes dentro do intervalo.
    faltantes = []
    if rangeInicio > 0 and rangeFim >= rangeInicio:
        intervaloEsperado = set(range(rangeInicio, rangeFim + 1))
        nfsFaltantes = sorted(intervaloEsperado - nfsEncontradas)
        for nf in nfsFaltantes:
            faltantes.append({"NF": nf})

    return {
        "faltantes": faltantes,
        "rangeInicio": rangeInicio,
        "rangeFim": rangeFim,
        "totalEsperado": (rangeFim - rangeInicio + 1) if rangeFim >= rangeInicio else 0,
        "totalEncontrado": len(nfsEncontradas),
        "totalFaltante": len(faltantes),
    }

def _history_from_reports(
    limit: int = 300,
    query: str = "",
    at_filter: str = "",
    venc_filter: str = "",
    nf_filter: str = "",
    cliente_filter: str = "",
    aba_filter: str = "",
    dt_from: str = "",
    dt_to: str = "",
    cnpj_emit: str = "",
    cnpj_dest: str = "",
) -> list[dict]:
    out = []
    q = str(query or "").strip().lower()
    f_at = str(at_filter or "").strip().lower()
    f_venc = str(venc_filter or "").strip().lower()
    f_nf = str(nf_filter or "").strip().lower()
    f_cliente = str(cliente_filter or "").strip().lower()
    f_aba = str(aba_filter or "").strip().lower()
    f_emit = re.sub(r"\D+", "", str(cnpj_emit or ""))
    f_dest = re.sub(r"\D+", "", str(cnpj_dest or ""))
    try:
        dt_from_obj = datetime.strptime(str(dt_from or "").strip(), "%Y-%m-%d").date() if str(dt_from or "").strip() else None
    except Exception:
        dt_from_obj = None
    try:
        dt_to_obj = datetime.strptime(str(dt_to or "").strip(), "%Y-%m-%d").date() if str(dt_to or "").strip() else None
    except Exception:
        dt_to_obj = None

    def _safe_float(v):
        try:
            if isinstance(v, str):
                vv = v.strip()
                if not vv:
                    return 0.0
                if "," in vv and "." in vv:
                    vv = vv.replace(".", "").replace(",", ".")
                elif "," in vv:
                    vv = vv.replace(",", ".")
                return float(vv)
            return float(v or 0)
        except Exception:
            return 0.0

    def _safe_int(v, default=0):
        try:
            return int(v)
        except Exception:
            return int(default)

    def _parse_date_obj(dt_text: str):
        t = str(dt_text or "").strip()
        if not t:
            return None
        candidates = [
            (t[:19], "%Y-%m-%d %H:%M:%S"),
            (t[:19], "%Y-%m-%dT%H:%M:%S"),
            (t[:10], "%Y-%m-%d"),
        ]
        for cand, fmt in candidates:
            try:
                return datetime.strptime(cand, fmt).date()
            except Exception:
                continue
        try:
            return datetime.fromisoformat(t.replace("Z", "+00:00")).date()
        except Exception:
            return None

    rel_dir = Path(RELATORIO_DIR)
    if not rel_dir.exists():
        return out

    files = sorted(rel_dir.glob("*.txt"), key=lambda p: p.stat().st_mtime, reverse=True)
    for fp in files:
        try:
            for raw_line in fp.read_text(encoding="utf-8", errors="ignore").splitlines():
                line = str(raw_line or "").strip()
                if not line:
                    continue
                if " - " in line:
                    dt, msg = line.split(" - ", 1)
                else:
                    dt, msg = "", line

                msg = str(msg or "").strip()
                if not msg.startswith("HIST_JSON "):
                    continue

                raw_json = msg[len("HIST_JSON ") :].strip()
                try:
                    payload = json.loads(raw_json)
                except Exception:
                    continue
                if not isinstance(payload, dict):
                    continue
                if str(payload.get("type") or "boleto_lancado") != "boleto_lancado":
                    continue

                at = str(payload.get("at") or dt or "").strip()
                date_obj = _parse_date_obj(at)
                if dt_from_obj and date_obj and date_obj < dt_from_obj:
                    continue
                if dt_to_obj and date_obj and date_obj > dt_to_obj:
                    continue

                emit = re.sub(r"\D+", "", str(payload.get("cnpj_emit") or ""))
                dest = re.sub(r"\D+", "", str(payload.get("cnpj_dest") or ""))
                if f_emit and f_emit not in emit:
                    continue
                if f_dest and f_dest not in dest:
                    continue

                nf = _normalize_report_text(str(payload.get("nf") or "").strip())
                cliente = _normalize_report_text(str(payload.get("cliente") or "").strip())
                descricao = _normalize_report_text(str(payload.get("descricao") or "").strip())
                vencimento = _normalize_report_text(str(payload.get("vencimento") or "").strip())
                parcela = _normalize_report_text(str(payload.get("parcela") or "").strip())
                valor_parcela = _safe_float(payload.get("valor_parcela"))
                valor_total = _safe_float(payload.get("valor_total"))
                valor_pago = payload.get("valor_pago")
                valor_pago_text = _normalize_report_text(str(valor_pago or "").strip())
                status = _normalize_report_text(str(payload.get("status") or "").strip())
                sheet_title = _normalize_report_text(str(payload.get("sheet_title") or "").strip())
                aba = _normalize_report_text(str(payload.get("aba") or "").strip())
                local = "/".join([x for x in (sheet_title, aba) if x]) or "Botana/RelatÃ³rio"

                at_search_parts = [at]
                try:
                    at_dt = datetime.fromisoformat(at.replace("Z", "+00:00"))
                    at_search_parts.extend(
                        [
                            at_dt.strftime("%d/%m/%Y"),
                            at_dt.strftime("%H:%M"),
                            at_dt.strftime("%d/%m/%Y %H:%M"),
                            at_dt.strftime("%d/%m/%Y, %H:%M:%S"),
                            at_dt.strftime("%Y-%m-%d"),
                            at_dt.strftime("%Y-%m-%d %H:%M"),
                        ]
                    )
                except Exception:
                    pass
                if f_at:
                    hay_at = " ".join(x for x in at_search_parts if x).lower()
                    if f_at not in hay_at:
                        continue
                if f_venc and f_venc not in vencimento.lower():
                    continue
                if f_nf and f_nf not in nf.lower():
                    continue
                if f_cliente:
                    hay_cliente = " ".join([cliente, descricao]).lower()
                    if f_cliente not in hay_cliente:
                        continue
                if f_aba:
                    hay_aba = " ".join([sheet_title, aba, local]).lower()
                    if f_aba not in hay_aba:
                        continue

                item = {
                    "type": "boleto_lancado",
                    "at": at,
                    "nf": nf,
                    "numero": nf,
                    "doc_tipo": "NF" if nf else "-",
                    "cliente": cliente,
                    "fornecedor": cliente,
                    "descricao": descricao,
                    "vencimento": vencimento,
                    "parcela": parcela,
                    "valor_parcela": valor_parcela,
                    "valor_total": valor_total,
                    "valor_pago": valor_pago_text,
                    "status": status,
                    "sheet_title": sheet_title,
                    "aba": aba,
                    "local_lancamento": local,
                    "cnpj_emit": emit,
                    "cnpj_dest": dest,
                    "qtd_parcelas": _safe_int(payload.get("qtd_parcelas"), 1),
                    "raw": _normalize_report_text(line),
                    "message": _normalize_report_text(f"NF {nf} - {cliente} - {descricao}"),
                }

                if q:
                    hay = " ".join(
                        [
                            at,
                            nf,
                            cliente,
                            descricao,
                            vencimento,
                            parcela,
                            str(valor_parcela),
                            str(valor_total),
                            valor_pago_text,
                            status,
                            local,
                            emit,
                            dest,
                        ]
                    ).lower()
                    if q not in hay:
                        continue

                out.append(item)
                if len(out) >= limit:
                    break
        except Exception:
            continue

    # I detect duplicate entries (same NF + same parcela) and mark them.
    # Eu detecto entradas duplicadas (mesma NF + mesma parcela) e as marco.
    contadorChaves = {}
    for item in out:
        chave = (str(item.get("nf", "")).strip(), str(item.get("parcela", "")).strip())
        if chave[0]:  # Only count if NF is not empty / Só conto se a NF não é vazia
            contadorChaves[chave] = contadorChaves.get(chave, 0) + 1
    for item in out:
        chave = (str(item.get("nf", "")).strip(), str(item.get("parcela", "")).strip())
        item["duplicata"] = bool(chave[0] and contadorChaves.get(chave, 0) > 1)

    return out


def _delete_history_entry(nf: str, parcela: str, at: str) -> dict:
    """I remove a specific HIST_JSON entry from the report files.
    # Eu removo uma entrada HIST_JSON especifica dos arquivos de relatorio."""
    nfAlvo = str(nf or "").strip()
    parcelaAlvo = str(parcela or "").strip()
    atAlvo = str(at or "").strip()
    if not nfAlvo:
        return {"ok": False, "message": "NF não informada"}

    relDir = Path(RELATORIO_DIR)
    if not relDir.exists():
        return {"ok": False, "message": "Diretório de relatórios não encontrado"}

    arquivos = sorted(relDir.glob("relatorio_*.txt"), key=lambda p: p.stat().st_mtime, reverse=True)
    for arquivo in arquivos:
        try:
            linhas = arquivo.read_text(encoding="utf-8", errors="ignore").splitlines()
        except Exception:
            continue

        linhasNovas = []
        removido = False
        for linha in linhas:
            textoLinha = str(linha or "").strip()
            if not textoLinha:
                linhasNovas.append(linha)
                continue
            # I check if this line is the HIST_JSON entry I want to delete.
            # Eu verifico se esta linha e a entrada HIST_JSON que eu quero deletar.
            if "HIST_JSON " in textoLinha and not removido:
                posJson = textoLinha.find("HIST_JSON ")
                if posJson >= 0:
                    jsonBruto = textoLinha[posJson + len("HIST_JSON "):].strip()
                    try:
                        payload = json.loads(jsonBruto)
                        nfPayload = str(payload.get("nf") or "").strip()
                        parcelaPayload = str(payload.get("parcela") or "").strip()
                        atPayload = str(payload.get("at") or "").strip()
                        # I match by NF + parcela, and optionally by timestamp.
                        # Eu faço match por NF + parcela, e opcionalmente pelo timestamp.
                        if nfPayload == nfAlvo and parcelaPayload == parcelaAlvo:
                            if not atAlvo or atPayload == atAlvo:
                                removido = True
                                continue  # I skip this line (delete it) / Eu pulo esta linha (deleto ela)
                    except Exception:
                        pass
            linhasNovas.append(linha)

        if removido:
            try:
                arquivo.write_text("\n".join(linhasNovas) + "\n", encoding="utf-8")
                logger.info("Entrada NF %s parcela %s removida do relatório %s", nfAlvo, parcelaAlvo, arquivo.name)
                return {"ok": True, "message": f"NF {nfAlvo} ({parcelaAlvo}) removida com sucesso"}
            except Exception as exc:
                logger.exception("Falha ao reescrever relatório %s: %s", arquivo.name, exc)
                return {"ok": False, "message": f"Falha ao salvar: {exc}"}

    return {"ok": False, "message": f"Entrada NF {nfAlvo} ({parcelaAlvo}) não encontrada nos relatórios"}


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
        if text.startswith("HIST_JSON "):
            try:
                payload = json.loads(text[len("HIST_JSON ") :].strip())
                nf = str(payload.get("nf") or "-").strip()
                cliente = str(payload.get("cliente") or "-").strip()
                parcela = str(payload.get("parcela") or "-").strip()
                venc = str(payload.get("vencimento") or "-").strip()
                processados.append(f"NF {nf} | Cliente: {cliente} | {parcela} | Venc: {venc}")
                continue
            except Exception:
                pass
        text = _normalize_report_text(text)
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
        service = _get_gmail_service_locked(timeout=0.3 if not force else 1.0)
        profile = service.users().getProfile(userId="me").execute()
        _EMAIL_CACHE.update(
            {
                "email": str(profile.get("emailAddress", "")).strip(),
                "error": "",
                "pending": False,
                "at": now,
            }
        )
    except TimeoutError:
        _EMAIL_CACHE.update({"error": "", "pending": True, "at": now})
    except Exception as exc:
        msg = str(exc).strip() or getattr(exc, "__class__", type(exc)).__name__ or "Falha ao obter perfil do e-mail"
        _EMAIL_CACHE.update({"error": msg, "pending": False, "at": now})
    return dict(_EMAIL_CACHE)


def _scheduler_next_seconds() -> int:
    global _NEXT_RUN_AT
    interval = max(30, int(_RUNTIME_SETTINGS.get("interval_seconds", INTERVALO)))
    now = time.time()
    if _NEXT_RUN_AT <= 0 or _NEXT_RUN_AT <= now:
        _NEXT_RUN_AT = now + interval
    return max(0, int(_NEXT_RUN_AT - now))


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
.btn-sec{background:linear-gradient(90deg,#6b4128,#4a2b18);color:#fff9f3}
.hidden{display:none!important}
.msg{margin-top:10px;font-size:.9rem;color:#9c2c1d;min-height:20px}
</style></head><body>
<section class="card">
<h1>Acesso ao Botana</h1>
<p>Entre com usuário e senha para continuar</p>
<label>Usuário</label><input id="u" type="text" autocomplete="username"/>
<label>Senha</label><input id="p" type="password" autocomplete="current-password"/>
<button id="b" onclick="login()">Entrar</button>
<button id="hubBackLogin" class="btn-sec hidden" type="button" onclick="backToHub()">Voltar ao HUB</button>
<div id="m" class="msg"></div>
</section>
<script>
const _PATH_RESERVED=new Set(['','login','logout','api','assets','static','store-image','favicon.ico']);
function _basePrefix(){const p=String(window.location.pathname||'/');const segs=p.split('/').filter(Boolean);if(!segs.length)return '';const first=String(segs[0]||'').toLowerCase();if(_PATH_RESERVED.has(first))return '';return `/${segs[0]}`;}
const _BASE_PREFIX=_basePrefix();
function _url(path){const p=String(path||'');if(!p.startsWith('/'))return p;if(!_BASE_PREFIX)return p;return p.startsWith(`${_BASE_PREFIX}/`)||p===_BASE_PREFIX?p:`${_BASE_PREFIX}${p}`;}
function backToHub(){
  try{
    const ref=document.referrer?new URL(document.referrer):null;
    if(ref&&ref.origin){window.location.assign(ref.origin+'/');return;}
  }catch(_){}
  window.location.assign(new URL('/',window.location.origin).toString());
}
function initHubBackLogin(){const b=document.getElementById('hubBackLogin');if(!b)return;if(_BASE_PREFIX)b.classList.remove('hidden');}
async function login(){
  const u=document.getElementById('u').value||'';
  const p=document.getElementById('p').value||'';
  const m=document.getElementById('m');
  const b=document.getElementById('b');
  b.disabled=true;
  m.textContent='Validando acesso';
  try{
    const r=await fetch(_url('/api/login'),{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:u,password:p})});
    const j=await r.json();
    if(r.ok&&j.ok){window.location.href=_url('/');return;}
    m.textContent=j.message||'Usuário ou senha inválidos';
  }catch(_){
    m.textContent='Falha ao conectar com o servidor';
  }finally{b.disabled=false;}
}
['u','p'].forEach(id=>{document.getElementById(id).addEventListener('keydown',(e)=>{if(e.key==='Enter')login();});});
initHubBackLogin();
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
.hub-back-btn{margin-right:8px}
.status-pill{display:inline-flex;align-items:center;gap:6px;padding:3px 8px;border-radius:999px;font-size:.76rem;font-weight:700;border:1px solid}
.ok{background:#e8f6ea;color:#2e7d32;border-color:#b6dfbf}
.off{background:#fff3e0;color:#8b4f19;border-color:#f2c8a3}
.err{background:#fdecec;color:#b42b2b;border-color:#f1bbbb}
.tabs{display:flex;gap:8px;padding:10px;margin-bottom:6px}
.tab-btn{background:#fff5ea;color:#5a311b;border:1px solid #d7b393;border-radius:9px;padding:8px 12px;font-weight:700;cursor:pointer}
.tab-btn.active{background:linear-gradient(90deg,var(--o),var(--o2));border-color:transparent;color:#2b1408}
.hidden{display:none!important}
.tab-panel.hidden{display:none!important}
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
.proc-grid{margin-top:8px;display:grid;gap:4px}
.proc-line{font-size:.82rem;color:#5e3a24}
.proc-progress{margin-top:2px;display:grid;gap:4px}
.proc-track{width:100%;height:10px;border-radius:999px;background:#f1ddcb;border:1px solid #debb9a;overflow:hidden}
.proc-fill{height:100%;width:0%;background:linear-gradient(90deg,var(--o),var(--o2));transition:width .25s ease}
.proc-label{font-size:.78rem;color:#6b4128}
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
.cfg-grid{display:grid;grid-template-columns:minmax(560px,1fr) minmax(120px,145px) minmax(230px,280px);gap:10px;align-items:stretch}
.cfg-grid > .card{height:100%;display:flex;flex-direction:column}
.cfg-main{display:grid;gap:8px;flex:1}
.cfg-main-card h3{text-align:center}
.cfg-fields{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px;align-items:end;justify-items:center}
.cfg-fields > div{display:flex;flex-direction:column;align-items:center}
.cfg-fields > div label{text-align:center}
.cfg-fields > div input,.cfg-fields > div select{width:min(160px,100%);text-align:center}
.cfg-fields > div input.num-sm{width:min(78px,100%)}
.cfg-fields > div select.mode-wide{width:min(190px,100%)}
.cfg-save{display:flex;justify-content:center;align-items:center}
.cfg-save button{min-width:0;width:auto}
.cfg-status{display:flex;flex-direction:column;align-items:center}
.cfg-status label{text-align:center}
.cfg-status input{width:min(240px,100%);text-align:center}
.auth-card .btns{flex-direction:column;justify-content:center;flex:1}
.auth-card h3{text-align:center}
.auth-card .btns button{width:100%}
.reproc-card .btns{margin-top:8px;justify-content:center}
.reproc-card h3{text-align:center}
.reproc-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px;align-items:end;justify-items:center}
.reproc-grid > div{display:flex;flex-direction:column;align-items:center}
.reproc-grid > div label{text-align:center}
.reproc-grid > div input,.reproc-grid > div select{text-align:center}
.cb{margin-top:8px;display:inline-flex;align-items:center;gap:8px}
.hist-filters{display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:8px;align-items:end}
.hist-filters > div{display:flex;flex-direction:column;justify-content:center;align-items:center}
.hist-filters > div label{width:100%;text-align:center}
.hist-filters > div input,.hist-filters > div select{width:100%;text-align:center}
.table-wrap{width:100%;overflow:auto;border:1px solid #d9af86;border-radius:10px;background:#fffdfb;box-shadow:inset 0 0 0 1px rgba(217,175,134,.22)}
.hist-toolbar{margin-top:8px;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}
.hist-note{flex:1 1 420px}
.hist-reset-btn{padding:6px 10px;font-size:.78rem}
.hist-table{width:100%;min-width:1240px;border-collapse:collapse;font-size:.8rem;table-layout:fixed;border:1px solid #ddb38d}
.hist-table th,.hist-table td{border:1px solid #e7c4a5;padding:7px 8px;text-align:center;vertical-align:middle;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.hist-table th{position:sticky;top:0;background:#fff1e3;color:#5c341c;z-index:1;padding-right:18px;border-bottom:2px solid #cf9c73;overflow:visible}
.hist-table th.sortable{cursor:pointer;user-select:none}
.hist-table th.sortable:after{content:""}
.hist-table tbody tr:nth-child(even){background:rgba(255,244,232,.92)}
.hist-table tbody tr:hover{background:rgba(238,155,47,.08)}
.hist-table th.is-resizing{background:#ffe5cf}
.col-resizer{position:absolute;top:0;right:-6px;width:12px;height:100%;cursor:col-resize;user-select:none;touch-action:none;z-index:4}
.col-resizer::after{content:"";position:absolute;top:7px;bottom:7px;left:5px;width:2px;background:#c68551;border-radius:999px;opacity:.78}
.cell-menu{position:relative;display:flex;align-items:center;gap:6px;justify-content:center;width:100%}
.cell-btn{display:block;width:100%;padding:0;border:0;background:transparent;color:#5a311b;font-weight:700;cursor:pointer;text-align:center;max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.cell-btn:hover{text-decoration:underline}
.cell-pop{position:absolute;top:100%;left:0;background:#fffaf6;border:1px solid #e7c8a8;border-radius:8px;padding:6px;box-shadow:0 8px 20px rgba(21,11,6,.15);display:none;z-index:5;min-width:160px}
.cell-pop button{width:100%;border:0;background:#fff1e3;padding:6px;border-radius:6px;cursor:pointer;font-size:.78rem;color:#5a311b}
.cell-menu.open .cell-pop{display:block}
.dup-row{background:rgba(220,53,69,.16)!important}
.dup-row:hover{background:rgba(220,53,69,.24)!important}
.dup-badge{display:inline-block;padding:2px 7px;border-radius:999px;font-size:.7rem;font-weight:700;background:#fde7ea;color:#a61d2d;border:1px solid #dc3545;margin-left:4px}
.del-btn{padding:4px 8px;border:1px solid #e0a0a0;border-radius:6px;background:#fdecec;color:#b42b2b;font-size:.75rem;font-weight:600;cursor:pointer;white-space:nowrap}
.del-btn:hover{background:#f8d7d7;border-color:#c45050}
.ov{position:fixed;inset:0;z-index:99999;display:none;align-items:center;justify-content:center;background:rgba(22,10,5,.78);backdrop-filter:blur(3px)}
.ov.show{display:flex}
.ovb{width:min(440px,92vw);border-radius:14px;border:1px solid #f0c89d;background:linear-gradient(180deg,#fff6ec,#ffe8d4);text-align:center;padding:18px}
.cnt{margin-top:12px;font-size:2.4rem;font-weight:800;color:#b05714}
@media(max-width:900px){.lists{grid-template-columns:1fr}.cfg-grid{grid-template-columns:1fr}.cfg-fields{grid-template-columns:1fr 1fr}.reproc-grid{grid-template-columns:1fr}}
@media(max-width:1020px){.hist-filters{grid-template-columns:1fr 1fr 1fr}}
@media(max-width:640px){.top-right{flex-direction:column;align-items:flex-end}.hist-filters{grid-template-columns:1fr}}
</style></head><body>
<div id="ov" class="ov"><div class="ovb"><h4>Reautenticação em andamento</h4><p>Troque para a conta correta no navegador<br/>A autenticação começará em:</p><div id="cnt" class="cnt">5</div></div></div>
<main class="app">
  <section class="top">
    <span>Botana - Painel de Controle MVA</span>
    <div class="top-right">
      <span id="who" class="whoami">Usuário: -</span>
      <button id="backHubBtn" class="logout-btn hub-back-btn hidden" onclick="goHub()">Voltar ao HUB</button>
      <button class="logout-btn" onclick="logout()">Sair</button>
      <span id="pill" class="status-pill off"><span>•</span><span>Aguardando</span></span>
    </div>
  </section>

  <div class="tabs">
    <button id="tabBtnMain" type="button" class="tab-btn active" onclick="switchTab('main')">Painel</button>
    <button id="tabBtnHist" type="button" class="tab-btn" onclick="switchTab('hist')">Histórico</button>
    <button id="tabBtnDiag" type="button" class="tab-btn" onclick="switchTab('diag')">Diagnóstico</button>
  </div>

  <section id="tabMain" class="tab-panel">
    <section class="card">
      <h3>Status da conta de e-mail</h3>
      <div class="status-grid">
        <div class="s">
          <div class="h">
            <span>Conta Botana</span>
            <span id="accBadge" class="status-pill off"><span>•</span><span>Aguardando</span></span>
          </div>
          <div id="accEmail" class="muted">E-mail conectado: -</div>
          <div id="accDetail" class="muted">Aguardando leitura</div>
          <div id="accProblem" class="problem hidden">-</div>
        </div>
      </div>
      <div id="cool" class="muted" style="margin-top:8px">Próxima verificação automática: -</div>
      <div class="proc-grid">
        <div id="procRun" class="proc-line">Loop: -</div>
        <div class="proc-progress">
          <div class="proc-track"><div id="procBarFill" class="proc-fill"></div></div>
          <div id="procBarLabel" class="proc-label">Progresso: 0/0 (0%)</div>
        </div>
        <div id="procNow" class="proc-line">Ciclo atual: -</div>
        <div id="procLast" class="proc-line">Último ciclo: -</div>
      </div>
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
      <article class="card cfg-main-card">
        <h3>Configuração do Gmail</h3>
        <div class="cfg-main">
          <div class="cfg-fields">
            <div>
              <label>Período</label>
              <select id="mode" class="mode-wide">
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
              <input id="maxPages" class="num-sm" type="number" min="1" max="20"/>
            </div>
            <div>
              <label>Tamanho da página</label>
              <input id="pageSize" class="num-sm" type="number" min="1" max="500"/>
            </div>
            <div>
              <label>Intervalo de leitura</label>
              <input id="intervalMin" class="num-sm" type="number" min="1" max="720"/>
            </div>
          </div>
          <div class="cfg-save"><button onclick="saveSettings()">Salvar configuração</button></div>
          <div class="cfg-status"><label>Status da última execução</label><input id="last" type="text" readonly value="-"/></div>
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

  <section id="tabHist" class="tab-panel hidden">
    <section class="card" style="margin-top:10px">
      <h3>Histórico de processamento e lançamentos</h3>
      <div class="hist-filters">
        <div><label>Data / Horário</label><input id="hAt" type="text" placeholder="06/04/2026 12:30 ou 2026-04-06"/></div>
        <div><label>Vencimento</label><input id="hVenc" type="date"/></div>
        <div><label>NF</label><input id="hNf" type="text" placeholder="49001"/></div>
        <div><label>Cliente</label><input id="hCliente" type="text" placeholder="Nome curto ou completo"/></div>
        <div><label>Aba</label><input id="hAba" type="text" placeholder="Janeiro ou MVA/Janeiro"/></div>
        <div><label>Limite</label><input id="hLimit" type="number" min="10" max="2000" value="300"/></div>
        <div style="display:flex;align-items:end"><button onclick="loadHistory()">Aplicar filtros</button></div>
      </div>
      <div class="hist-toolbar">
        <div class="muted hist-note">O botão Excluir remove somente o registro do histórico/relatório. A linha da planilha não é apagada. Arraste a divisória do cabeçalho para reajustar as colunas.</div>
        <button type="button" class="sec hist-reset-btn" onclick="resetHistColumnWidths()">Resetar larguras</button>
      </div>
      <div class="table-wrap" style="margin-top:10px">
        <table class="hist-table">
          <colgroup>
            <col id="histCol-at" style="width:150px"/>
            <col id="histCol-venc" style="width:110px"/>
            <col id="histCol-doc" style="width:100px"/>
            <col id="histCol-cliente" style="width:240px"/>
            <col id="histCol-parcela" style="width:90px"/>
            <col id="histCol-vparcela" style="width:130px"/>
            <col id="histCol-vtotal" style="width:130px"/>
            <col id="histCol-local" style="width:150px"/>
            <col id="histCol-acao" style="width:110px"/>
          </colgroup>
          <thead>
            <tr>
              <th class="sortable" data-key="at" data-colid="at">Data</th>
              <th class="sortable" data-key="venc" data-colid="venc">Venc.</th>
              <th class="sortable" data-key="doc" data-colid="doc">NF</th>
              <th class="sortable" data-key="cliente" data-colid="cliente">Cliente</th>
              <th class="sortable" data-key="parcela" data-colid="parcela">Parc.</th>
              <th class="sortable" data-key="vparcela" data-colid="vparcela">Parcela</th>
              <th class="sortable" data-key="vtotal" data-colid="vtotal">Total</th>
              <th class="sortable" data-key="local" data-colid="local">Aba</th>
              <th data-colid="acao">Ação</th>
            </tr>
          </thead>
          <tbody id="hBody"><tr><td colspan="9">Sem dados</td></tr></tbody>
        </table>
      </div>
    </section>
  </section>

  <section id="tabDiag" class="tab-panel hidden">
    <section class="card" style="margin-top:10px">
      <h3>Diagnóstico</h3>
      <pre id="details">-</pre>
    </section>
  </section>
</main>
<script>
const _PATH_RESERVED=new Set(['','login','logout','api','assets','static','store-image','favicon.ico']);
function _basePrefix(){const p=String(window.location.pathname||'/');const segs=p.split('/').filter(Boolean);if(!segs.length)return '';const first=String(segs[0]||'').toLowerCase();if(_PATH_RESERVED.has(first))return '';return `/${segs[0]}`;}
const _BASE_PREFIX=_basePrefix();
function _url(path){const p=String(path||'');if(!p.startsWith('/'))return p;if(!_BASE_PREFIX)return p;return p.startsWith(`${_BASE_PREFIX}/`)||p===_BASE_PREFIX?p:`${_BASE_PREFIX}${p}`;}
function goHub(){
  try{
    const ref=document.referrer?new URL(document.referrer):null;
    if(ref&&ref.origin&&ref.origin!==window.location.origin){
      window.location.assign(ref.origin+'/');
      return;
    }
  }catch(_){}
  const target=new URL('/',window.location.origin).toString();
  window.location.assign(target);
}
function initHubBackButton(){const b=document.getElementById('backHubBtn');if(!b)return;if(_BASE_PREFIX)b.classList.remove('hidden');else b.classList.add('hidden');}
async function api(path,opts){const r=await fetch(_url(path),opts);const j=await r.json().catch(()=>({}));if(r.status===401){window.location.href=_url('/login');throw new Error('não autenticado');}if(!r.ok){throw new Error(String((j&&j.message)||`HTTP ${r.status}`));}return j;}
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
let _activeTab='main';
function _tabFromLocation(){
  const h=String(window.location.hash||'').replace('#','').trim().toLowerCase();
  if(h==='hist'||h==='diag'||h==='main') return h;
  const q=new URLSearchParams(window.location.search||'');
  const t=String(q.get('tab')||'').trim().toLowerCase();
  if(t==='hist'||t==='diag'||t==='main') return t;
  return 'main';
}
function switchTab(tab){
  const next=(tab==='hist'||tab==='diag')?tab:'main';
  const changed=_activeTab!==next;
  const m=next==='main';
  const h=next==='hist';
  const d=next==='diag';
  document.getElementById('tabMain').classList.toggle('hidden',!m);
  document.getElementById('tabHist').classList.toggle('hidden',!h);
  document.getElementById('tabDiag').classList.toggle('hidden',!d);
  document.getElementById('tabBtnMain').classList.toggle('active',m);
  document.getElementById('tabBtnHist').classList.toggle('active',h);
  document.getElementById('tabBtnDiag').classList.toggle('active',d);
  _activeTab=next;
  if(next==='hist'){
    loadHistory().catch(()=>{});
  }
  const nextHash='#'+next;
  if(window.location.hash!==nextHash){
    try{history.replaceState(null,'',nextHash);}catch(_){}
  }
  if(changed){
    try{window.scrollTo({top:0,behavior:'auto'});}catch(_){window.scrollTo(0,0);}
  }
}
function setPill(ok,running){const p=document.getElementById('pill');if(running){p.className='status-pill ok';p.innerHTML='<span>•</span><span>Em execução</span>';return;}if(ok){p.className='status-pill off';p.innerHTML='<span>•</span><span>Aguardando</span>';return;}p.className='status-pill err';p.innerHTML='<span>•</span><span>Com erro</span>';}
function setAccBadge(kind,label){
  const b=document.getElementById('accBadge');
  b.className='status-pill '+kind;
  b.innerHTML='<span>•</span><span>'+label+'</span>';
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
function _fmtCycleShort(c){
  const it=c||{};
  const m=Number(it.messages||0);
  const a=Number(it.attachments||0);
  const x=Number(it.xmls||0);
  const l=Number(it.launched||0);
  return `E-mails ${m} | Anexos ${a} | XML ${x} | Lançamentos ${l}`;
}
function updProcessing(proc,maxMessages){
  const runEl=document.getElementById('procRun');
  const nowEl=document.getElementById('procNow');
  const lastEl=document.getElementById('procLast');
  const barFill=document.getElementById('procBarFill');
  const barLabel=document.getElementById('procBarLabel');
  if(!runEl||!nowEl||!lastEl||!barFill||!barLabel) return;
  const p=proc||{};
  const reading=!!p.reading;
  const running=!!p.running;
  if(reading) runEl.textContent='Loop: executando ciclo agora';
  else if(running) runEl.textContent='Loop: ativo (aguardando proximo ciclo)';
  else runEl.textContent='Loop: pausado';

  const cur=p.current||{};
  const curStart=cur.started_at?_fmtDateTime(cur.started_at):'-';
  nowEl.textContent=`Ciclo atual: inicio ${curStart} | ${_fmtCycleShort(cur)}`;

  const last=p.last||{};
  const lastEnd=last.finished_at?_fmtDateTime(last.finished_at):'-';
  let statusTxt='-';
  if(last.ok===true) statusTxt='OK';
  else if(last.ok===false) statusTxt='Erro';
  let msg=`Ultimo ciclo: ${statusTxt} em ${lastEnd} | ${_fmtCycleShort(last)}`;
  const err=String(last.error||'').trim();
  if(statusTxt==='Erro'&&err) msg += ` | ${err}`;
  lastEl.textContent=msg;

  const maxV=Math.max(1, Number(maxMessages||100));
  let curV=Number(cur.messages||0);
  if(!Number.isFinite(curV)||curV<0) curV=0;
  if(curV>maxV) curV=maxV;
  const perc=Math.max(0,Math.min(100,Math.round((curV/maxV)*100)));
  barFill.style.width=String(perc)+'%';
  barLabel.textContent=`Progresso: ${curV}/${maxV} (${perc}%)`;
}
async function refresh(){
  try{
    const j=await api('/api/state');
    const reading=!!j.reading;
    const ok=!!(j.last_status&&j.last_status.ok);
    const s=(j.settings||{});
    document.getElementById('who').textContent='Usuário: '+String((j.auth&&j.auth.user)||'-');
    document.getElementById('mode').value=String(s.gmail_filter_mode||'last_30_days');
    document.getElementById('maxPages').value=String(s.gmail_max_pages||3);
    document.getElementById('pageSize').value=String(s.gmail_page_size||50);
    document.getElementById('intervalMin').value=String(s.loop_interval_minutes||30);
    document.getElementById('last').value=String((j.last_status&&j.last_status.message)||'-');
    const diag={
      last_status:(j.last_status||{}),
      account:(j.account||{}),
      scheduler:(j.scheduler||{}),
      processing:(j.processing||{}),
      diagnostic:(j.diagnostic||{})
    };
    document.getElementById('details').textContent=JSON.stringify(diag,null,2);
    _nextRemain=Number((j.scheduler&&j.scheduler.next_in_seconds)||0);
    _tickNext();
    updAccount(j.account||{});
    setPill(ok,reading);
    updDaily(j.daily_report||{});
    updProcessing(j.processing||{}, Number(j.max_messages||100));
  }catch(err){
    document.getElementById('details').textContent=JSON.stringify({erro:String((err&&err.message)||err||'Falha ao atualizar estado')},null,2);
  }
}
async function startLoop(){await api('/api/start',{method:'POST'});refresh();}
async function stopLoop(){await api('/api/stop',{method:'POST'});refresh();}
async function runNow(){await api('/api/run-now',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({account:'principal'})});refresh();}
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
async function reprocess(){const payload={account:'principal',days:Number(document.getElementById('days').value||30),max_messages:Number(document.getElementById('limit').value||100),mark_unread:document.getElementById('unread').checked};await api('/api/reprocess',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});refresh();}
function _fmtDateTime(v){if(!v)return '-';try{return new Date(v).toLocaleString('pt-BR');}catch(_){return String(v);}}
function _esc(s){return String(s??'').replace(/[&<>\"']/g,(c)=>({'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;',\"'\":'&#39;'}[c]));}
function _compactSpaces(s){return String(s||'').replace(/\\s+/g,' ').trim();}
function _compactClienteLabel(cliente,descricao){
  const fonte=_compactSpaces(descricao||cliente||'');
  if(!fonte)return '-';
  const blt=(fonte.match(/\\bBLT[\\s:-]*(\\d+(?:-\\d+)*)\\b/i)||[])[1]||'';
  let base=fonte.replace(/\\bBLT[\\s:-]*\\d+(?:-\\d+)*.*$/i,'').trim();
  if(!base)base=fonte;
  const palavras=base.split(/\\s+/).filter(Boolean);
  const uteis=palavras.filter(p=>!/^(DE|DA|DO|DAS|DOS|E)$/i.test(p));
  const escolhidas=(uteis.length?uteis:palavras).slice(0,blt?2:3).join(' ');
  const nomeCurto=_compactSpaces(escolhidas||base);
  return _compactSpaces(blt?`${nomeCurto} ${blt}`:nomeCurto);
}
function _toggleMenu(ev,btn){ev.stopPropagation();const wrap=btn.closest('.cell-menu');document.querySelectorAll('.cell-menu.open').forEach(x=>{if(x!==wrap)x.classList.remove('open');});wrap.classList.toggle('open');}
async function _showCnpj(ev,btn){ev.stopPropagation();const cnpj=btn.getAttribute('data-cnpj')||'-';try{await navigator.clipboard.writeText(cnpj);}catch(_){}const wrap=btn.closest('.cell-menu');if(wrap)wrap.classList.remove('open');}
document.addEventListener('click',()=>{document.querySelectorAll('.cell-menu.open').forEach(x=>x.classList.remove('open'));});
let _histItems=[];
let _histSort={key:'at',dir:'desc'};
const _histColStorageKey='botana.hist.colwidths.v1';
const _histColDefaults={at:150,venc:110,doc:100,cliente:240,parcela:90,vparcela:130,vtotal:130,local:150,acao:110};
let _histColWidths={..._histColDefaults};
function _saveHistColWidths(){try{localStorage.setItem(_histColStorageKey,JSON.stringify(_histColWidths));}catch(_){}}
function _loadHistColWidths(){
  try{
    const raw=localStorage.getItem(_histColStorageKey);
    if(!raw)return;
    const parsed=JSON.parse(raw);
    if(!parsed||typeof parsed!=='object')return;
    _histColWidths={..._histColDefaults};
    Object.keys(_histColDefaults).forEach((key)=>{
      const next=Number(parsed[key]);
      if(Number.isFinite(next)&&next>60)_histColWidths[key]=Math.round(next);
    });
  }catch(_){}
}
function _applyHistColWidths(){
  Object.entries(_histColDefaults).forEach(([key,baseWidth])=>{
    const col=document.getElementById(`histCol-${key}`);
    if(!col)return;
    const minWidth=key==='cliente'?180:80;
    const width=Math.max(minWidth,Number(_histColWidths[key]||baseWidth));
    col.style.width=`${width}px`;
  });
}
function resetHistColumnWidths(){
  _histColWidths={..._histColDefaults};
  _applyHistColWidths();
  _saveHistColWidths();
}
function _initHistoryColumnResize(){
  const table=document.querySelector('.hist-table');
  if(!table||table.dataset.resizeReady==='1')return;
  table.dataset.resizeReady='1';
  _loadHistColWidths();
  _applyHistColWidths();
  table.querySelectorAll('thead th[data-colid]').forEach((th)=>{
    if(th.querySelector('.col-resizer'))return;
    const key=th.dataset.colid;
    const handle=document.createElement('span');
    handle.className='col-resizer';
    handle.title='Arraste para reajustar';
    handle.addEventListener('click',(ev)=>{ev.preventDefault();ev.stopPropagation();});
    const beginResize=(startX)=>{
      const col=document.getElementById(`histCol-${key}`);
      if(!col)return;
      const startWidth=col.getBoundingClientRect().width;
      th.classList.add('is-resizing');
      const onMove=(moveX)=>{
        const minWidth=key==='cliente'?180:80;
        _histColWidths[key]=Math.max(minWidth,Math.round(startWidth+(moveX-startX)));
        _applyHistColWidths();
      };
      const onMouseMove=(moveEv)=>{onMove(moveEv.clientX);};
      const onMouseUp=()=>{
        th.classList.remove('is-resizing');
        window.removeEventListener('mousemove',onMouseMove);
        window.removeEventListener('mouseup',onMouseUp);
        _saveHistColWidths();
      };
      window.addEventListener('mousemove',onMouseMove);
      window.addEventListener('mouseup',onMouseUp);
    };
    handle.addEventListener('mousedown',(ev)=>{
      ev.preventDefault();
      ev.stopPropagation();
      beginResize(ev.clientX);
    });
    th.appendChild(handle);
  });
}
function _fmtLocal(local){const s=String(local||'');if(!s)return '-';const parts=s.split('/');if(parts.length>=2)return parts.slice(-2).join('/');return s;}
function _fmtMoney(v){
  if(v===null||v===undefined) return '-';
  if(typeof v==='number'&&Number.isFinite(v)){
    return v.toLocaleString('pt-BR',{style:'currency',currency:'BRL'});
  }
  const txt=String(v).trim();
  if(!txt) return '-';
  let norm=txt;
  if(norm.includes(',')&&norm.includes('.')) norm=norm.replace(/[.]/g,'').replace(',','.');
  else if(norm.includes(',')) norm=norm.replace(',','.');
  const n=Number(norm);
  if(!Number.isFinite(n)) return txt;
  return n.toLocaleString('pt-BR',{style:'currency',currency:'BRL'});
}
function _getSortValue(it,key){
  if(key==='at')return it.at||'';
  if(key==='venc')return it.vencimento||'';
  if(key==='doc')return it._doc||'';
  if(key==='cliente')return it._cliente_view||it.descricao||it.cliente||'';
  if(key==='parcela')return it.parcela||'';
  if(key==='vparcela')return Number(it.valor_parcela||0);
  if(key==='vtotal')return Number(it.valor_total||0);
  if(key==='local')return it._local||'';
  return '';
}
function _sortHist(items){const k=_histSort.key;const dir=_histSort.dir==='asc'?1:-1;return [...items].sort((a,b)=>{const va=_getSortValue(a,k);const vb=_getSortValue(b,k);if(va<vb)return -1*dir;if(va>vb)return 1*dir;return 0;});}
function _renderHistory(items){
  _histItems=Array.isArray(items)?items:[];
  const body=document.getElementById('hBody');
  if(!body)return;
  body.innerHTML='';
  let arr=_histItems.filter(it=>it&&it.type==='boleto_lancado');
  arr=arr.map(it=>{
    const nf=String(it.nf||it.numero||'').trim();
    const doc=nf?`NF ${nf}`:'-';
    const local=_fmtLocal(it.local_lancamento);
    const clienteView=_compactClienteLabel(it.cliente,it.descricao);
    return {...it,_doc:doc,_local:local,_cliente_view:clienteView};
  });
  if(!arr.length){body.innerHTML='<tr><td colspan="9">Sem dados para os filtros selecionados</td></tr>';return;}
  arr=_sortHist(arr);
  arr.forEach(it=>{
    const tr=document.createElement('tr');
    if(it.duplicata)tr.classList.add('dup-row');
    const emit=String(it.cnpj_emit||'-');
    const dupTag=it.duplicata?'<span class="dup-badge">DUPLICADA</span>':'';
    const tituloCliente=_compactSpaces(it.descricao||it.cliente||'-');
    const menu=`<div class=\"cell-menu\"><button class=\"cell-btn\" title=\"${_esc(tituloCliente)}\" onclick=\"_toggleMenu(event,this)\">${_esc(it._cliente_view||'-')}</button><div class=\"cell-pop\"><button data-cnpj=\"${_esc(emit)}\" onclick=\"_showCnpj(event,this)\">Copiar CNPJ emitente</button></div></div>`;
    const delBtn=`<button class=\"del-btn\" title=\"Remove apenas este registro do histórico\" onclick=\"deleteEntry('${_esc(it.nf||'')}','${_esc(it.parcela||'')}','${_esc(it.at||'')}')\">Excluir</button>`;
    tr.innerHTML=`<td title=\"${_esc(_fmtDateTime(it.at))}\">${_fmtDateTime(it.at)}</td><td title=\"${_esc(it.vencimento||'-')}\">${_esc(it.vencimento||'-')}</td><td title=\"${_esc(it._doc)}\">${_esc(it._doc)}${dupTag}</td><td>${menu}</td><td title=\"${_esc(it.parcela||'-')}\">${_esc(it.parcela||'-')}</td><td title=\"${_esc(_fmtMoney(it.valor_parcela))}\">${_fmtMoney(it.valor_parcela)}</td><td title=\"${_esc(_fmtMoney(it.valor_total))}\">${_fmtMoney(it.valor_total)}</td><td title=\"${_esc(it._local)}\">${_esc(it._local)}</td><td>${delBtn}</td>`;
    body.appendChild(tr);
  });
}
function _setSort(key){
  const ths=document.querySelectorAll('.hist-table th.sortable');
  ths.forEach(th=>{th.classList.remove('asc');th.classList.remove('desc');});
  if(_histSort.key===key){_histSort.dir=_histSort.dir==='asc'?'desc':'asc';}
  else{_histSort.key=key;_histSort.dir='asc';}
  const th=document.querySelector(`.hist-table th.sortable[data-key="${key}"]`);
  if(th)th.classList.add(_histSort.dir);
  _renderHistory(_histItems);
}
_initHistoryColumnResize();
document.querySelectorAll('.hist-table th.sortable').forEach(th=>{th.addEventListener('click',()=>_setSort(th.dataset.key));});
async function loadHistory(silent=false){
  try{
    const p=new URLSearchParams();
    const vAt=((document.getElementById('hAt')||{}).value||'').trim();
    const vVenc=((document.getElementById('hVenc')||{}).value||'').trim();
    const vNf=((document.getElementById('hNf')||{}).value||'').trim();
    const vCliente=((document.getElementById('hCliente')||{}).value||'').trim();
    const vAba=((document.getElementById('hAba')||{}).value||'').trim();
    const vLimit=Number((document.getElementById('hLimit')||{}).value||300);
    if(vAt)p.set('at',vAt);
    if(vVenc)p.set('venc',vVenc);
    if(vNf)p.set('nf',vNf);
    if(vCliente)p.set('cliente',vCliente);
    if(vAba)p.set('aba',vAba);
    p.set('limit',String(Math.max(10,Math.min(2000,vLimit||300))));
    const j=await api('/api/history?'+p.toString());
    const items=(j&&Array.isArray(j.items))?j.items:[];
    _renderHistory(items);
  }catch(err){
    if(!silent)console.warn('Erro ao carregar histórico:',err);
    const body=document.getElementById('hBody');
    if(body)body.innerHTML='<tr><td colspan="9">Erro de rede: '+_esc(String(err&&err.message||err))+'</td></tr>';
  }
}
async function deleteEntry(nf,parcela,at){
  if(!nf)return;
  const msg=`Tem certeza que deseja excluir a NF ${nf} (${parcela||'-'})?\nIsso remove apenas o registro do histórico/relatório, não a linha da planilha.`;
  if(!confirm(msg))return;
  try{
    const r=await api('/api/history/delete',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({nf:nf,parcela:parcela,at:at})});
    if(r&&r.ok){await loadHistory();}
    else{alert(r&&r.message||'Falha ao excluir');}
  }catch(err){alert('Erro ao excluir: '+(err&&err.message||err));}
}
async function logout(){await fetch(_url('/api/logout'),{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'}).catch(()=>{});window.location.href=_url('/login');}
['mode','maxPages','pageSize','intervalMin'].forEach(id=>{const el=document.getElementById(id);if(!el)return;el.addEventListener('keydown',(e)=>{if(e.key==='Enter'){e.preventDefault();saveSettings();}});});
['days','limit'].forEach(id=>{const el=document.getElementById(id);if(!el)return;el.addEventListener('keydown',(e)=>{if(e.key==='Enter'){e.preventDefault();reprocess();}});});
document.querySelectorAll('#hAt,#hVenc,#hNf,#hCliente,#hAba,#hLimit').forEach(el=>{el.addEventListener('keydown',(e)=>{if(e.key==='Enter'){e.preventDefault();loadHistory();}});});
window.addEventListener('hashchange',()=>{const t=_tabFromLocation();if(t!==_activeTab)switchTab(t);});
refresh();loadHistory();switchTab(_tabFromLocation());setInterval(refresh,3000);setInterval(_tickNext,1000);setInterval(loadHistory,10000);
initHubBackButton();
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
            
            # Hub acessa o botana pelo localhost, permitimos essas acoes pelo proxy sem token
            if self.client_address[0] in ["127.0.0.1", "::1", "localhost"] and parsed_path in ["/api/relatorio-nfs", "/api/clean-sheets", "/api/clean-sheets/log"]:
                return "hub_internal"

            if parsed_path.startswith("/api/"):
                _json_response(self, 401, {"ok": False, "message": "NÃ£o autenticado"})
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
                if not str(email_info.get("email", "")).strip() and not str(email_info.get("error", "")).strip():
                    email_info = _connected_email(force=True)
                last_msg = str((last_status or {}).get("message", "") or "")
                reading_info = _reading_state()
                reading_now = bool(reading_info.get("reading"))
                email_value = str(email_info.get("email", "")).strip()
                email_err = str(email_info.get("error", "")).strip()
                email_pending = bool(email_info.get("pending", False))
                if email_err:
                    acc_status = "error"
                    friendly = "Falha ao validar o e-mail conectado"
                elif email_pending:
                    acc_status = "waiting"
                    friendly = "Autenticação em andamento. Aguarde a confirmação no navegador"
                elif not email_value:
                    acc_status = "waiting"
                    friendly = "Autenticação pendente. Clique em Autenticação > Principal"
                elif reading_now:
                    acc_status = "running"
                    friendly = "Lendo os e-mails agora"
                elif not running:
                    acc_status = "waiting"
                    friendly = "Monitoramento pausado. Use Executar agora ou inicie o loop."
                elif (last_status or {}).get("ok", True):
                    acc_status = "waiting"
                    friendly = "Aguardando a próxima verificação automática"
                else:
                    acc_status = "error"
                    friendly = "Falha na comunicação com a API. Veja os detalhes técnicos para identificar a causa."
                status_view = {
                    "ok": acc_status != "error",
                    "message": (email_err or friendly or last_msg or "Aguardando"),
                    "at": (last_status or {}).get("at"),
                }
                if not status_view["at"] and (email_value or reading_now):
                    status_view["at"] = datetime.now().isoformat()
                return _json_response(
                    self,
                    200,
                    {
                        "ok": True,
                        "running": bool(running),
                        "reading": bool(reading_now),
                        "interval_seconds": int(_RUNTIME_SETTINGS.get("interval_seconds", INTERVALO)),
                        "max_messages": int(_RUNTIME_SETTINGS.get("max_messages", 100)),
                        "settings": {
                            "gmail_filter_mode": str(_RUNTIME_SETTINGS.get("gmail_filter_mode", "last_30_days")),
                            "gmail_max_pages": int(_RUNTIME_SETTINGS.get("gmail_max_pages", 3)),
                            "gmail_page_size": int(_RUNTIME_SETTINGS.get("gmail_page_size", 50)),
                            "loop_interval_minutes": int(_RUNTIME_SETTINGS.get("loop_interval_minutes", 30)),
                        },
                        "last_status": status_view,
                        "account": {
                            "email": email_value,
                            "status": acc_status,
                            "friendly": friendly,
                            "error": email_err,
                            "detail": last_msg,
                        },
                        "scheduler": {
                            "next_in_seconds": _scheduler_next_seconds(),
                        },
                        "processing": _process_snapshot(),
                        "diagnostic": {
                            "reading_source": str(reading_info.get("source", "")),
                            "reading_flag_raw": bool(reading_info.get("flag_raw", False)),
                            "reading_current_active": bool(reading_info.get("current_active", False)),
                            "reading_corrected": bool(reading_info.get("corrected", False)),
                        },
                        "daily_report": _daily_report_data(),
                        "auth": {"user": user, "role": _role_of(user)},
                    },
                )
            if parsed.path == "/api/relatorio-nfs":
                qs = parse_qs(parsed.query or "")
                filtro = (qs.get("filtro", [""])[0] or "todos").strip()
                mes = (qs.get("mes", [""])[0] or "").strip()
                nf_inicio = (qs.get("nf_inicio", [""])[0] or "").strip()
                nf_fim = (qs.get("nf_fim", [""])[0] or "").strip()
                empresa = (qs.get("empresa", [""])[0] or "todos").strip()
                try:
                    resultado = _gerar_relatorio_nfs(filtro, mes, nf_inicio, nf_fim, empresa)
                    return _json_response(self, 200, {"status": "success", **resultado})
                except Exception as e:
                    return _json_response(self, 500, {"status": "error", "message": str(e)})

            if parsed.path == "/api/clean-sheets/log":
                qs = parse_qs(parsed.query or "")
                try:
                    desdeIndice = int((qs.get("desde", ["0"])[0] or "0").strip())
                except (ValueError, TypeError):
                    desdeIndice = 0
                try:
                    import correcao_planilhas
                    logData = correcao_planilhas.obterLog(desdeIndice)
                    return _json_response(self, 200, {"ok": True, **logData})
                except Exception as e:
                    return _json_response(self, 500, {"ok": False, "message": str(e)})

            if parsed.path == "/api/history":
                qs = parse_qs(parsed.query or "")
                at_filter = (qs.get("at", [""])[0] or "").strip()
                venc_filter = (qs.get("venc", [""])[0] or "").strip()
                nf_filter = (qs.get("nf", [""])[0] or "").strip()
                cliente_filter = (qs.get("cliente", [""])[0] or "").strip()
                aba_filter = (qs.get("aba", [""])[0] or "").strip()
                dt_from = (qs.get("from", [""])[0] or "").strip()
                dt_to = (qs.get("to", [""])[0] or "").strip()
                cnpj_emit = (qs.get("cnpj_emit", [""])[0] or "").strip()
                cnpj_dest = (qs.get("cnpj_dest", [""])[0] or "").strip()
                query = (qs.get("q", [""])[0] or "").strip()
                try:
                    limit = int((qs.get("limit", ["300"])[0] or "300").strip())
                except Exception:
                    limit = 300
                items = _history_from_reports(
                    limit=max(10, min(limit, 2000)),
                    query=query,
                    at_filter=at_filter,
                    venc_filter=venc_filter,
                    nf_filter=nf_filter,
                    cliente_filter=cliente_filter,
                    aba_filter=aba_filter,
                    dt_from=dt_from,
                    dt_to=dt_to,
                    cnpj_emit=cnpj_emit,
                    cnpj_dest=cnpj_dest,
                )
                return _json_response(self, 200, {"items": items})
            return _json_response(self, 404, {"ok": False, "message": "NÃ£o encontrado"})

        def do_POST(self):
            try:
                parsed = urlparse(self.path)
                data = _read_json(self)

                if parsed.path == "/api/login":
                    username = str(data.get("username", "")).strip()
                    password = str(data.get("password", ""))
                    if not _verify_login(username, password):
                        return _json_response(self, 401, {"ok": False, "message": "UsuÃ¡rio ou senha invÃ¡lidos"})
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
                        return _json_response(self, 403, {"ok": False, "message": "Sem permissÃ£o"})
                    _save_settings(
                        {
                            "gmail_filter_mode": data.get("gmail_filter_mode", _RUNTIME_SETTINGS.get("gmail_filter_mode", "last_30_days")),
                            "gmail_max_pages": data.get("gmail_max_pages", _RUNTIME_SETTINGS.get("gmail_max_pages", 3)),
                            "gmail_page_size": data.get("gmail_page_size", _RUNTIME_SETTINGS.get("gmail_page_size", 50)),
                            "loop_interval_minutes": data.get("loop_interval_minutes", _RUNTIME_SETTINGS.get("loop_interval_minutes", 30)),
                            "max_messages": data.get("max_messages", _RUNTIME_SETTINGS.get("max_messages", 100)),
                        }
                    )
                    return _json_response(self, 200, {"ok": True, "message": "ConfiguraÃ§Ã£o salva"})
                if parsed.path == "/api/reprocess":
                    if not _can_operate(user):
                        return _json_response(self, 403, {"ok": False, "message": "Sem permissÃ£o"})
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
                    friendly = f"Reprocessamento concluÃ­do: {result.get('changed', 0)} de {result.get('matched', 0)} mensagens atualizadas"
                    return _json_response(self, 200, {"ok": True, "result": result, "friendly": friendly})
                if parsed.path == "/api/reauth":
                    if not _can_operate(user):
                        return _json_response(self, 403, {"ok": False, "message": "Sem permissÃ£o"})
                    account = str(data.get("account", "principal")).strip().lower()
                    if account != "principal":
                        account = "principal"
                    info = _reauthenticate_gmail()
                    return _json_response(self, 200, {"ok": True, "account": account, "message": info.get("message", "ReautenticaÃ§Ã£o concluÃ­da"), "friendly": "ReautenticaÃ§Ã£o concluÃ­da"})
                if parsed.path == "/api/clean-sheets":
                    if not _can_operate(user):
                        return _json_response(self, 403, {"ok": False, "message": "Sem permissão"})
                    try:
                        import correcao_planilhas
                        empresaFiltro = str(data.get("empresa", "todos")).strip() or "todos"
                        abaFiltro = str(data.get("aba", "")).strip()
                        iniciou = correcao_planilhas.iniciar_assistente_em_background(empresaFiltro, abaFiltro)
                        if iniciou:
                            return _json_response(self, 200, {"ok": True, "friendly": "Assistente de correção iniciado."})
                        else:
                            return _json_response(self, 409, {"ok": False, "friendly": "Já existe uma correção em andamento."})
                    except Exception as e:
                        return _json_response(self, 500, {"ok": False, "friendly": f"Falha ao iniciar o assistente: {e}"})

                if parsed.path == "/api/history/delete":
                    if not _can_operate(user):
                        return _json_response(self, 403, {"ok": False, "message": "Sem permissão"})
                    nfDel = str(data.get("nf", "")).strip()
                    parcelaDel = str(data.get("parcela", "")).strip()
                    atDel = str(data.get("at", "")).strip()
                    if not nfDel:
                        return _json_response(self, 400, {"ok": False, "message": "NF não informada"})
                    resultado = _delete_history_entry(nf=nfDel, parcela=parcelaDel, at=atDel)
                    statusCode = 200 if resultado.get("ok") else 404
                    return _json_response(self, statusCode, resultado)

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
