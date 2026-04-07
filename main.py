import argparse
import csv
import hashlib
import hmac
import io
import json
import secrets
import os, re, time, gspread, threading, sys
import unicodedata
from email.utils import parseaddr, parsedate_to_datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse, parse_qs
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from config import PLANILHAS, CNPJ_MVA, CNPJ_EH, INTERVALO, DOWNLOAD_DIR, RELATORIO_DIR, GOOGLE_CREDENTIALS_SHEETS, GOOGLE_CREDENTIALS_GMAIL
from gmail_service import (
    getGmailService,
    buscarMessagesEnviadosPagina,
    baixar_anexos_de_mensagem,
    build_sent_xml_query,
    list_botana_label_ids,
    marcar_mensagem_com_label,
    marcar_mensagem_para_reprocessar,
    listar_mensagens_com_labels_botana,
)
from reporter import escreverRelatorio, consolidarRelatorioTMP
from xml_parser import extrairDadosXML
from sheets_writer import atualizarPlanilha
from logger_config import logger, cor_ciano, reset
try:
    from tray_icon import run_tray
except Exception:
    run_tray = None
try:
    from PyPDF2 import PdfReader
except Exception:
    PdfReader = None

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
_LOOP_THREAD = None
last_status = {"ok": True, "message": "Aguardando", "at": None}
APPDATA_BASE = Path(os.getenv("APPDATA", str(Path.home() / "AppData" / "Roaming"))) / "Botana"
APPDATA_BASE.mkdir(parents=True, exist_ok=True)
_SETTINGS_FILE = APPDATA_BASE / "panel_settings.json"
_AUTH_FILE = APPDATA_BASE / "panel_auth.json"
_WATCH_SEARCH_NAMES_FILE = APPDATA_BASE / "watch_search_names.txt"
_SETTINGS_LOCK = threading.RLock()
_AUTH_LOCK = threading.Lock()
_WATCH_SEARCH_NAMES_LOCK = threading.Lock()
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
_PROCESS_EXEC_LOCK = threading.Lock()
_PROCESS_STATS = {
    "current": {
        "active": False,
        "started_at": "",
        "messages": 0,
        "attachments": 0,
        "xmls": 0,
        "launched": 0,
        "duplicates": 0,
    },
    "last": {
        "ok": None,
        "started_at": "",
        "finished_at": "",
        "messages": 0,
        "attachments": 0,
        "xmls": 0,
        "launched": 0,
        "duplicates": 0,
        "error": "",
    },
}
_MANUAL_ACTION_LOCK = threading.Lock()
_MANUAL_ACTION = {
    "active": False,
    "kind": "",
    "phase": "",
    "label": "",
    "status": "idle",
    "message": "Nenhuma acao manual em andamento.",
    "detail": "",
    "started_at": "",
    "finished_at": "",
    "progress_current": 0,
    "progress_total": 0,
    "changed": 0,
    "failed": 0,
    "matched": 0,
    "inspected": 0,
    "current_email": "",
    "current_subject": "",
    "current_date": "",
    "requested_limit": 0,
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
            "duplicates": 0,
        }


def _process_update(
    messages: int | None = None,
    attachments: int | None = None,
    xmls: int | None = None,
    launched: int | None = None,
    duplicates: int | None = None,
):
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
        if duplicates is not None:
            cur["duplicates"] = max(0, int(duplicates))


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
            "duplicates": int(cur.get("duplicates", 0) or 0),
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


def _manual_action_snapshot() -> dict:
    with _MANUAL_ACTION_LOCK:
        return dict(_MANUAL_ACTION)


def _manual_action_begin(kind: str, label: str, message: str, detail: str = "", progress_total: int = 0, **extra) -> tuple[bool, dict]:
    with _MANUAL_ACTION_LOCK:
        if bool(_MANUAL_ACTION.get("active")):
            return False, dict(_MANUAL_ACTION)
        _MANUAL_ACTION.update(
            {
                "active": True,
                "kind": str(kind or "").strip(),
                "phase": "",
                "label": str(label or "").strip(),
                "status": "running",
                "message": str(message or "").strip() or "Acao manual em andamento.",
                "detail": str(detail or "").strip(),
                "started_at": datetime.now().isoformat(),
                "finished_at": "",
                "progress_current": 0,
                "progress_total": max(0, int(progress_total or 0)),
                "changed": 0,
                "failed": 0,
                "matched": 0,
                "inspected": 0,
                "current_email": "",
                "current_subject": "",
                "current_date": "",
                "requested_limit": 0,
            }
        )
        for key, value in extra.items():
            _MANUAL_ACTION[key] = value
        return True, dict(_MANUAL_ACTION)


def _manual_action_update(
    *,
    message: str | None = None,
    detail: str | None = None,
    progress_current: int | None = None,
    progress_total: int | None = None,
    **extra,
) -> dict:
    with _MANUAL_ACTION_LOCK:
        if message is not None:
            _MANUAL_ACTION["message"] = str(message or "").strip()
        if detail is not None:
            _MANUAL_ACTION["detail"] = str(detail or "").strip()
        if progress_current is not None:
            _MANUAL_ACTION["progress_current"] = max(0, int(progress_current))
        if progress_total is not None:
            _MANUAL_ACTION["progress_total"] = max(0, int(progress_total))
        for key, value in extra.items():
            _MANUAL_ACTION[key] = value
        return dict(_MANUAL_ACTION)


def _manual_action_finish(ok: bool, message: str, detail: str = "", **extra) -> dict:
    with _MANUAL_ACTION_LOCK:
        _MANUAL_ACTION["active"] = False
        _MANUAL_ACTION["status"] = "success" if ok else "error"
        _MANUAL_ACTION["message"] = str(message or "").strip() or ("Acao concluida." if ok else "Acao concluida com erro.")
        _MANUAL_ACTION["detail"] = str(detail or "").strip()
        _MANUAL_ACTION["finished_at"] = datetime.now().isoformat()
        _MANUAL_ACTION["phase"] = ""
        _MANUAL_ACTION["current_email"] = ""
        _MANUAL_ACTION["current_subject"] = ""
        _MANUAL_ACTION["current_date"] = ""
        _MANUAL_ACTION["requested_limit"] = 0
        _MANUAL_ACTION["matched"] = 0
        _MANUAL_ACTION["inspected"] = 0
        for key, value in extra.items():
            _MANUAL_ACTION[key] = value
        return dict(_MANUAL_ACTION)


def _manual_action_busy_message() -> str:
    state = _manual_action_snapshot()
    if bool(state.get("active")):
        label = str(state.get("label") or "Acao manual").strip()
        return f"{label} ja esta em andamento."
    if _reading_active():
        return "Ja existe uma leitura/processamento em andamento."
    return ""


def _start_run_now_background(max_messages_override: int | None = None) -> tuple[bool, dict]:
    snap = _manual_action_snapshot()
    if bool(snap.get("active")):
        if not snap.get("message"):
            label = str(snap.get("label") or "Acao manual").strip() or "Acao manual"
            snap["message"] = f"{label} ja esta em andamento."
        return False, snap
    started, snap = _manual_action_begin(
        "run_now",
        "Execu\u00e7\u00e3o manual",
        "Execu\u00e7\u00e3o manual iniciada.",
        detail=(
            f"O ciclo manual vai ler at\u00e9 {int(max_messages_override)} mensagens."
            if max_messages_override
            else "Acompanhe os contadores de leitura e lan\u00e7amentos no painel."
        ),
        requested_limit=int(max_messages_override or 0),
    )
    if not started:
        return False, snap

    def _worker():
        resume_loop = bool(running)
        try:
            if resume_loop:
                _manual_action_update(
                    message="Interrompendo ciclo autom\u00e1tico para executar agora.",
                    detail="O loop autom\u00e1tico ser\u00e1 retomado depois da execu\u00e7\u00e3o manual.",
                )
                parar_verificacao(wait=True, timeout=120.0)
                if (_LOOP_THREAD and _LOOP_THREAD.is_alive()) or not _wait_for_processing_idle(timeout=5.0):
                    _manual_action_finish(
                        False,
                        "N\u00e3o foi poss\u00edvel executar agora.",
                        detail="O ciclo autom\u00e1tico n\u00e3o liberou a leitura a tempo.",
                    )
                    return
            stop_event.clear()
            _manual_action_update(
                message="Execu\u00e7\u00e3o manual em andamento.",
                detail=(
                    f"Leitura manual em andamento com limite de {int(max_messages_override)} mensagens."
                    if max_messages_override
                    else "Acompanhe os contadores de leitura e lan\u00e7amentos no painel."
                ),
            )
            ok, msg = executar_um_ciclo(max_messages_override=max_messages_override)
            proc = _process_snapshot().get("last", {})
            detail = (
                f"E-mails: {int(proc.get('messages', 0) or 0)} | "
                f"Anexos: {int(proc.get('attachments', 0) or 0)} | "
                f"XML: {int(proc.get('xmls', 0) or 0)} | "
                f"Lan\u00e7amentos: {int(proc.get('launched', 0) or 0)} | "
                f"Duplicadas: {int(proc.get('duplicates', 0) or 0)}"
            )
            if resume_loop:
                restarted = iniciar_verificacao()
                detail = f"{detail} | Loop autom\u00e1tico {'retomado' if restarted else 'n\u00e3o retomado'}."
            _manual_action_finish(
                ok,
                msg,
                detail=detail,
                progress_current=int(proc.get("messages", 0) or 0),
                progress_total=int(max_messages_override or 0),
                requested_limit=int(max_messages_override or 0),
            )
        except Exception as exc:
            logger.exception("Falha na execução manual em background: %s", exc)
            if resume_loop:
                try:
                    iniciar_verificacao()
                except Exception:
                    logger.exception("Falha ao retomar loop automatico apos erro na execucao manual.")
            _manual_action_finish(False, "Erro na execu\u00e7\u00e3o manual.", detail=str(exc))

    threading.Thread(target=_worker, daemon=True, name="botana-run-now").start()
    return True, snap


def _start_reprocess_background(max_messages: int, mark_unread: bool) -> tuple[bool, dict]:
    snap = _manual_action_snapshot()
    if bool(snap.get("active")):
        if not snap.get("message"):
            label = str(snap.get("label") or "Acao manual").strip() or "Acao manual"
            snap["message"] = f"{label} ja esta em andamento."
        return False, snap
    started, snap = _manual_action_begin(
        "reprocess",
        "Reprocessamento",
        "Reprocessamento iniciado.",
        detail=f"At\u00e9 {int(max_messages)} mensagens mais recentes com label do Botana ser\u00e3o remarcadas e relidas neste ciclo.",
        progress_total=int(max_messages),
        requested_limit=int(max_messages),
    )
    if not started:
        return False, snap

    def _progress(**kwargs):
        _manual_action_update(**kwargs)

    def _worker():
        resume_loop = bool(running)
        try:
            if resume_loop:
                _manual_action_update(
                    message="Interrompendo ciclo autom\u00e1tico para reprocessar.",
                    detail="O loop autom\u00e1tico ser\u00e1 retomado depois do reprocessamento.",
                )
                parar_verificacao(wait=True, timeout=120.0)
                if (_LOOP_THREAD and _LOOP_THREAD.is_alive()) or not _wait_for_processing_idle(timeout=5.0):
                    _manual_action_finish(
                        False,
                        "N\u00e3o foi poss\u00edvel iniciar o reprocessamento.",
                        detail="O ciclo autom\u00e1tico n\u00e3o liberou a leitura a tempo.",
                        requested_limit=int(max_messages),
                    )
                    return
            stop_event.clear()
            _manual_action_update(
                message="Reprocessamento em andamento.",
                detail="Atualizando a label do Botana nas mensagens mais recentes antes de reler os e-mails.",
                phase="marking",
            )
            result = _reprocess_recent(max_messages=max_messages, mark_unread=mark_unread, progress_cb=_progress)
            targets = list(result.get("targets") or [])
            changed = int(result.get("changed", 0) or 0)
            failed = int(result.get("failed", 0) or 0)
            matched = int(result.get("matched", 0) or 0)
            if targets:
                _manual_action_update(
                    phase="processing",
                    progress_current=0,
                    progress_total=len(targets),
                    changed=changed,
                    failed=failed,
                    current_email="",
                    current_subject="",
                    current_date="",
                    message="Labels atualizadas. Iniciando leitura dos e-mails reprocessados.",
                    detail=f"O Botana vai reler at\u00e9 {len(targets)} mensagens reprocessadas para tentar relan\u00e7ar na planilha.",
                )
                ok, msg = executar_um_ciclo(
                    max_messages_override=len(targets),
                    messages_override=targets,
                    preserve_reprocess_label=True,
                )
                proc = _process_snapshot().get("last", {})
                detail = (
                    f"Labels atualizadas: {changed}/{matched} | "
                    f"Falhas ao marcar: {failed} | "
                    f"E-mails: {int(proc.get('messages', 0) or 0)} | "
                    f"Anexos: {int(proc.get('attachments', 0) or 0)} | "
                    f"XML: {int(proc.get('xmls', 0) or 0)} | "
                    f"Lan\u00e7amentos: {int(proc.get('launched', 0) or 0)} | "
                    f"Duplicadas: {int(proc.get('duplicates', 0) or 0)}"
                )
                if resume_loop:
                    restarted = iniciar_verificacao()
                    detail = f"{detail} | Loop autom\u00e1tico {'retomado' if restarted else 'n\u00e3o retomado'}."
                _manual_action_finish(
                    ok,
                    msg,
                    detail=detail,
                    progress_current=int(proc.get("messages", 0) or 0),
                    progress_total=len(targets),
                    changed=changed,
                    failed=failed,
                    requested_limit=int(max_messages),
                )
                return
            if matched <= 0:
                friendly = "Nenhuma mensagem com label do Botana foi encontrada para reprocessar."
            elif changed <= 0:
                friendly = "Nenhuma mensagem foi marcada para reprocessar."
            else:
                friendly = f"Reprocessamento concluido: {changed} de {matched} mensagens atualizadas."
            detail = f"Falhas: {failed} | Marcar como nao lido: {'sim' if mark_unread else 'nao'}"
            if resume_loop:
                restarted = iniciar_verificacao()
                detail = f"{detail} | Loop autom\u00e1tico {'retomado' if restarted else 'n\u00e3o retomado'}."
            _manual_action_finish(
                changed > 0 or matched == 0,
                friendly,
                detail=detail,
                progress_current=matched,
                progress_total=matched,
                changed=changed,
                failed=failed,
                requested_limit=int(max_messages),
            )
        except Exception as exc:
            logger.exception("Falha no reprocessamento em background: %s", exc)
            if resume_loop:
                try:
                    iniciar_verificacao()
                except Exception:
                    logger.exception("Falha ao retomar loop automatico apos erro no reprocessamento.")
            _manual_action_finish(False, "Erro no reprocessamento.", detail=str(exc), requested_limit=int(max_messages))

    threading.Thread(target=_worker, daemon=True, name="botana-reprocess").start()
    return True, snap


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
                "max_messages": max(
                    1,
                    min(
                        1000,
                        max(
                            int(_RUNTIME_SETTINGS.get("max_messages", 100)),
                            max(1, min(20, int(_RUNTIME_SETTINGS.get("gmail_max_pages", 3))))
                            * max(1, min(500, int(_RUNTIME_SETTINGS.get("gmail_page_size", 50)))),
                        ),
                    ),
                ),
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
        effective_max_messages = max(
            max_pages * page_size,
            max(1, min(1000, int(raw.get("max_messages", max_pages * page_size)))),
        )
        out = {
            "gmail_filter_mode": mode,
            "gmail_max_pages": max_pages,
            "gmail_page_size": page_size,
            "loop_interval_minutes": loop_min,
            "interval_seconds": max(30, min(86400, loop_min * 60)),
            "max_messages": max(1, min(1000, effective_max_messages)),
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
        effective_max_messages = max(
            max_pages * page_size,
            max(1, min(1000, int(data.get("max_messages", max_pages * page_size)))),
        )
        out = {
            "gmail_filter_mode": mode,
            "gmail_max_pages": max_pages,
            "gmail_page_size": page_size,
            "loop_interval_minutes": loop_min,
            "interval_seconds": max(30, min(86400, loop_min * 60)),
            "max_messages": max(1, min(1000, effective_max_messages)),
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


def _safe_money_text(value: str) -> float:
    try:
        txt = str(value or "").strip().replace("\xa0", " ")
        if not txt:
            return 0.0
        txt = re.sub(r"[^\d,.\-]+", "", txt)
        if not txt or txt in {"-", ",", "."}:
            return 0.0
        if "," in txt and "." in txt:
            if txt.rfind(",") > txt.rfind("."):
                txt = txt.replace(".", "").replace(",", ".")
            else:
                txt = txt.replace(",", "")
        elif "," in txt:
            txt = txt.replace(",", ".")
        return float(txt)
    except Exception:
        return 0.0


def _normalize_ddmmyyyy(date_raw: str) -> str:
    if not date_raw:
        return ""
    candidates = [
        str(date_raw).strip(),
        str(date_raw).strip().replace(".", "/"),
        str(date_raw).strip().replace("-", "/"),
    ]
    formats = (
        "%d/%m/%Y",
        "%Y/%m/%d",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d.%m.%Y",
    )
    for cand in candidates:
        for fmt in formats:
            try:
                return datetime.strptime(cand, fmt).strftime("%d/%m/%Y")
            except Exception:
                continue
    return ""


def _is_boleto_pdf_name(name: str) -> bool:
    nome_upper = str(name or "").upper()
    return bool(re.search(r"[_\s-]?(BLT|BOLET[OA]?|BOLTO|BOLETOO|BOLETT?)", nome_upper))


def _extract_boleto_number_from_name(name: str) -> str:
    nome_upper = str(name or "").upper()
    if not _is_boleto_pdf_name(nome_upper):
        return ""
    match = re.findall(r"([0-9]{2,}-?[0-9]+)", nome_upper)
    if not match:
        return ""
    num_boleto = match[-1]
    m_clean = re.search(r"0{4,}([1-9][0-9]*(-[0-9A-Z]+)?)$", num_boleto)
    if m_clean:
        num_boleto = m_clean.group(1)
    if num_boleto in {"0136", "136"}:
        num_boleto = "10136"
    elif num_boleto.startswith("0136-"):
        num_boleto = num_boleto.replace("0136-", "10136-", 1)
    elif num_boleto.startswith("136-"):
        num_boleto = num_boleto.replace("136-", "10136-", 1)
    return num_boleto


def _extract_nf_number(value: str) -> str:
    txt = str(value or "").upper()
    match = re.search(r"\bNF\s*0*([0-9]{3,})\b", txt)
    return str(match.group(1) or "").strip() if match else ""


def _extract_pdf_text(file_path: str) -> str:
    if PdfReader is None:
        return ""
    chunks = []
    try:
        reader = PdfReader(file_path)
    except Exception as exc:
        logger.warning("Falha ao abrir PDF %s para leitura de fallback: %s", os.path.basename(file_path), exc)
        return ""
    for page in reader.pages[:3]:
        try:
            chunks.append(page.extract_text() or "")
        except Exception:
            continue
    return _normalize_report_text(" ".join(chunks))


def _extract_boleto_due_and_value(pdf_text: str) -> tuple[str, float]:
    text = _normalize_report_text(pdf_text)
    if not text:
        return "", 0.0
    upper = text.upper()
    regions = []
    for marker in ("VENCIMENTO", "VENCTO", "DATA DE VENCIMENTO"):
        pos = upper.find(marker)
        if pos >= 0:
            regions.append(text[pos : pos + 260])
    if not regions:
        regions.append(text[:320])
    for region in regions:
        date_match = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", region)
        if not date_match:
            continue
        vencimento = _normalize_ddmmyyyy(date_match.group(1))
        tail = region[date_match.end() : date_match.end() + 120]
        value_match = re.search(r"\b(\d{1,3}(?:\.\d{3})*,\d{2})\b", tail)
        if vencimento and value_match:
            valor = _safe_money_text(value_match.group(1))
            if valor > 0:
                return vencimento, valor
    return "", 0.0


def _extract_boleto_pdf_info(file_path: str) -> dict | None:
    nome_arquivo = os.path.basename(file_path)
    if not _is_boleto_pdf_name(nome_arquivo):
        return None
    numero = _extract_boleto_number_from_name(nome_arquivo)
    pdf_text = _extract_pdf_text(file_path)
    vencimento, valor = _extract_boleto_due_and_value(pdf_text)
    nf = _extract_nf_number(nome_arquivo) or _extract_nf_number(pdf_text)
    sort_date = "9999-12-31"
    try:
        if vencimento:
            sort_date = datetime.strptime(vencimento, "%d/%m/%Y").strftime("%Y-%m-%d")
    except Exception:
        pass
    info = {
        "numero": numero,
        "nf": nf,
        "vencimento": vencimento,
        "valor": float(valor or 0.0),
        "arquivo": nome_arquivo,
        "_sort_date": sort_date,
        "_sort_num": _safe_money_text(re.sub(r"[^\d]", "", numero or "")),
    }
    if not info["numero"] and not info["vencimento"] and info["valor"] <= 0:
        return None
    return info


def _boletos_for_xml(dados_xml: dict, boleto_infos: list[dict]) -> list[dict]:
    nf = str(dados_xml.get("nf") or "").strip()
    candidatos = [dict(item) for item in list(boleto_infos or []) if isinstance(item, dict)]
    if nf:
        matched = [item for item in candidatos if str(item.get("nf") or "").strip() == nf]
        if matched:
            candidatos = matched
    candidatos.sort(
        key=lambda item: (
            str(item.get("_sort_date") or "9999-12-31"),
            float(item.get("_sort_num") or 0.0),
            str(item.get("numero") or ""),
            str(item.get("arquivo") or ""),
        )
    )
    return candidatos


def _infer_parcelas_from_boleto_pdfs(dados_xml: dict, boleto_infos: list[dict]) -> dict:
    payload = dict(dados_xml or {})
    parcelas = list(payload.get("parcelas") or [])
    if str(payload.get("parcelas_source") or "").strip().lower() != "fat":
        return payload
    candidatos = _boletos_for_xml(payload, boleto_infos)
    if len(parcelas) != 1 or len(candidatos) <= 1:
        return payload
    if not all(str(item.get("vencimento") or "").strip() and float(item.get("valor") or 0.0) > 0 for item in candidatos):
        return payload
    valor_total = float(payload.get("valorTotal") or 0.0)
    soma_boletos = round(sum(float(item.get("valor") or 0.0) for item in candidatos), 2)
    if valor_total > 0 and abs(soma_boletos - valor_total) > 0.05:
        logger.warning(
            "NF %s com XML de fatura unica teve %d boletos PDF, mas a soma %.2f difere do total %.2f. Mantendo XML.",
            payload.get("nf"),
            len(candidatos),
            soma_boletos,
            valor_total,
        )
        return payload
    parcelas_inferidas = []
    for idx, boleto in enumerate(candidatos, start=1):
        parcelas_inferidas.append(
            {
                "numero": idx,
                "numParcela": f"{idx}ª Parcela",
                "vencimento": str(boleto.get("vencimento") or "").strip(),
                "valor": float(boleto.get("valor") or 0.0),
            }
        )
    payload["parcelas"] = parcelas_inferidas
    payload["qtdParcelas"] = len(parcelas_inferidas)
    payload["parcelas_source"] = "pdf_fallback"
    payload["vencimento"] = parcelas_inferidas[0]["vencimento"]
    payload["numParcela"] = parcelas_inferidas[0]["numParcela"]
    payload["valorParcela"] = parcelas_inferidas[0]["valor"]
    try:
        payload["anoVencimento"] = datetime.strptime(payload["vencimento"], "%d/%m/%Y").strftime("%Y")
    except Exception:
        pass
    logger.info(
        "NF %s inferiu %d parcelas a partir dos PDFs de boleto porque o XML veio apenas com fatura total.",
        payload.get("nf"),
        len(parcelas_inferidas),
    )
    return payload


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

def processar_emails_enviados(
    max_messages_override: int | None = None,
    messages_override: list[dict] | None = None,
    preserve_reprocess_label: bool = False,
):
    global _IS_READING
    _process_start()
    service = _get_gmail_service_locked()
    batch_limit = max(1, min(1000, int(max_messages_override or _RUNTIME_SETTINGS.get("max_messages", 100))))
    page_size = max(1, min(500, int(_RUNTIME_SETTINGS.get("gmail_page_size", 50) or 50)))
    max_pages = max(1, min(50, int(_RUNTIME_SETTINGS.get("gmail_max_pages", 3) or 3)))
    if max_messages_override:
        max_pages = max(max_pages, (batch_limit + page_size - 1) // page_size)
    search_query = _cycle_search_query()
    force_messages = messages_override is not None
    forced_messages = []
    skip_botana_ids = set()
    if force_messages:
        for item in list(messages_override or []):
            if not isinstance(item, dict):
                continue
            msg_id = str(item.get("id", "")).strip()
            if not msg_id:
                continue
            forced_messages.append(dict(item))
    else:
        try:
            skip_botana_ids = {str(label_id).strip() for label_id in list_botana_label_ids(service) if str(label_id).strip()}
        except Exception as exc:
            logger.warning("Falha ao carregar labels do Botana para o filtro de leitura: %s", exc)
    total_msgs = 0
    anexos_lidos = 0
    xmls_lidos = 0
    total_processados = 0
    total_duplicados = 0
    page_token = None
    primeira_pagina = True
    paginas_lidas = 0
    vistos = set()
    interativo_cmd = bool(sys.stdin and sys.stdin.isatty())
    abort_logged = False
    def _summary() -> dict:
        return {
            "messages": int(total_msgs),
            "attachments": int(anexos_lidos),
            "xmls": int(xmls_lidos),
            "launched": int(total_processados),
            "duplicates": int(total_duplicados),
        }

    def _sync_progress():
        _process_update(
            messages=total_msgs,
            attachments=anexos_lidos,
            xmls=xmls_lidos,
            launched=total_processados,
            duplicates=total_duplicados,
        )

    def _abort_if_requested() -> bool:
        nonlocal abort_logged
        if not _processing_abort_requested():
            return False
        if not abort_logged:
            abort_logged = True
            logger.info("Ciclo interrompido por solicitacao.")
            escreverRelatorio(f"{_now()} - CICLO interrompido por solicitacao.")
        _sync_progress()
        return True

    while True:
        if _abort_if_requested():
            return _summary()
        if force_messages:
            msgs = list(forced_messages)
            next_page_token = None
            force_messages = False
        else:
            restante = max(0, batch_limit - total_msgs)
            if restante <= 0:
                break
            msgs, next_page_token = buscarMessagesEnviadosPagina(
                service,
                max_results=min(page_size, max(1, restante)),
                page_token=page_token,
                query=search_query,
                skip_label_ids=skip_botana_ids,
            )
            paginas_lidas += 1

        if primeira_pagina and not msgs:
            logger.info("Nenhuma mensagem enviada com XML encontrada.")
            escreverRelatorio(f"{_now()} - CICLO: 0 e-mails lidos, 0 anexos, 0 XML, 0 lanÃ§amentos.")
            _sync_progress()
            return _summary()

        primeira_pagina = False

        for m in msgs:
            if _abort_if_requested():
                return _summary()
            if total_msgs >= batch_limit:
                break
            msg_id = str(m.get("id", "")).strip()
            if not msg_id or msg_id in vistos:
                continue
            vistos.add(msg_id)
            total_msgs += 1
            _sync_progress()
            logger.info("Abrindo mensagem ID: %s", msg_id)

            arquivos = baixar_anexos_de_mensagem(service, msg_id)
            if not arquivos:
                logger.info("Nenhum anexo salvo para mensagem %s", msg_id)
                continue
            anexos_lidos += len(arquivos)
            _sync_progress()

            dados_xmls = []
            boleto_infos = []

            # Processa todos os anexos baixados
            for arquivo in arquivos:
                if _abort_if_requested():
                    return _summary()
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
                        boleto_info = _extract_boleto_pdf_info(arquivo)
                        if boleto_info:
                            boleto_infos.append(boleto_info)
                            logger.info(
                                "Boleto identificado: %s (BLT %s, venc %s, valor %.2f)",
                                nome_arquivo,
                                str(boleto_info.get("numero") or "-"),
                                str(boleto_info.get("vencimento") or "-"),
                                float(boleto_info.get("valor") or 0.0),
                            )
                        else:
                            logger.info("PDF ignorado (nÃ£o parece boleto ou nÃ£o foi possÃ­vel extrair metadados): %s", nome_arquivo)

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
            if _abort_if_requested():
                return _summary()
            try:
                label_aplicada = marcar_mensagem_com_label(
                    service,
                    msg_id,
                    existing_label_ids=m.get("labelIds", []),
                    reprocessed=True if preserve_reprocess_label else None,
                )
                if label_aplicada:
                    logger.info("E-mail %s marcado com '%s'", msg_id, label_aplicada)
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
                if _abort_if_requested():
                    return _summary()
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

                xml_boletos = _boletos_for_xml(dados_xml, boleto_infos)
                dados_xml = _infer_parcelas_from_boleto_pdfs(dados_xml, xml_boletos)
                cnpj_emit = re.sub(r"\D+", "", str(dados_xml.get("cnpjEmitente") or ""))
                ano = dados_xml.get("anoVencimento")
                planilha_id = escolher_planilha_por_cnpj_e_ano(cnpj_emit, ano)

                if not planilha_id:
                    logger.warning("CNPJ %s ou ano %s sem planilha configurada.", cnpj_emit, ano)
                    continue

                # Itera sobre todas as parcelas - mapeamento correto de boletos -> parcelas
                parcelas = dados_xml.get("parcelas", [])
                n_parcelas = len(parcelas)
                n_boletos = len(xml_boletos)

                # monta lista de boletos por parcela (mesmo tamanho de parcelas)
                if n_parcelas == 0:
                    continue  # nada a fazer

                if n_boletos == 0:
                    boletos_map = [None] * n_parcelas
                else:
                    # Se tiver igual, mapeia 1:1; se menor, preenche em ordem; se maior, usa sÃ³ os primeiros N
                    boletos_map = [
                        str((xml_boletos[i] or {}).get("numero") or "").strip() if i < n_boletos else None
                        for i in range(n_parcelas)
                    ]
                    if n_boletos > n_parcelas:
                        logger.info(
                            "Mais boletos (%d) que parcelas (%d). Sobraram: %s",
                            n_boletos,
                            n_parcelas,
                            [str((item or {}).get("numero") or "").strip() for item in xml_boletos[n_parcelas:]],
                        )

                # Agora processa 1 vez por parcela, usando o boleto mapeado (ou None)
                for idx, parcela in enumerate(parcelas):
                    if _abort_if_requested():
                        return _summary()
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
                        if _abort_if_requested():
                            return _summary()
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
                            elif isinstance(resultado, dict) and bool(resultado.get("duplicate")):
                                total_duplicados += 1
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

        if messages_override is not None or total_msgs >= batch_limit:
            break
        skip_reached = bool(getattr(processar_emails_enviados, "_skip_reached", False))
        if SKIP_UNTIL_NF and not skip_reached and next_page_token:
            if interativo_cmd:
                resposta = input(
                    f"\nNF {SKIP_UNTIL_NF} nÃ£o encontrada neste lote de {page_size}. "
                    f"Deseja continuar com mais {page_size}? [s/N]: "
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
        if next_page_token and paginas_lidas < max_pages:
            page_token = next_page_token
            continue
        break
    logger.info("Ciclo finalizado. Total processado: %d", total_processados)
    escreverRelatorio(
        f"{_now()} - CICLO: {total_msgs} e-mails lidos, {anexos_lidos} anexos, {xmls_lidos} XML, {total_processados} lanÃ§amentos."
    )
    _sync_progress()
    return _summary()

def main_loop():
    global running, last_status, _NEXT_RUN_AT, _IS_READING, _LOOP_THREAD
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    running = True
    logger.info("[Botana] Loop iniciado")
    while not stop_event.is_set():
        _NEXT_RUN_AT = time.time() + int(_RUNTIME_SETTINGS.get("interval_seconds", INTERVALO))
        try:
            _IS_READING = True
            with _PROCESS_EXEC_LOCK:
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
    _LOOP_THREAD = None
    logger.info("[Botana] Loop finalizado")


def executar_um_ciclo(
    max_messages_override: int | None = None,
    messages_override: list[dict] | None = None,
    preserve_reprocess_label: bool = False,
):
    global last_status, _NEXT_RUN_AT, _IS_READING
    try:
        _IS_READING = True
        with _PROCESS_EXEC_LOCK:
            summary = processar_emails_enviados(
                max_messages_override=max_messages_override,
                messages_override=messages_override,
                preserve_reprocess_label=preserve_reprocess_label,
            )
        _process_finish(ok=True, error="")
        msg = _format_cycle_status("Execução manual concluída", summary)
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


def _format_cycle_status(prefix: str, summary: dict | None) -> str:
    item = summary or {}
    return (
        f"{prefix}: {int(item.get('messages', 0) or 0)} e-mails, "
        f"{int(item.get('attachments', 0) or 0)} anexos, "
        f"{int(item.get('xmls', 0) or 0)} XML, "
        f"{int(item.get('launched', 0) or 0)} lan\u00e7amentos, "
        f"{int(item.get('duplicates', 0) or 0)} duplicadas."
    )


def main_loop():
    global running, last_status, _NEXT_RUN_AT, _IS_READING, _LOOP_THREAD
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    running = True
    logger.info("[Botana] Loop iniciado")
    while not stop_event.is_set():
        _NEXT_RUN_AT = time.time() + int(_RUNTIME_SETTINGS.get("interval_seconds", INTERVALO))
        try:
            _IS_READING = True
            with _PROCESS_EXEC_LOCK:
                summary = processar_emails_enviados()
            _process_finish(ok=True, error="")
            msg = _format_cycle_status("Ciclo conclu\u00eddo", summary)
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
    _LOOP_THREAD = None
    logger.info("[Botana] Loop finalizado")


def executar_um_ciclo(
    max_messages_override: int | None = None,
    messages_override: list[dict] | None = None,
    preserve_reprocess_label: bool = False,
):
    global last_status, _NEXT_RUN_AT, _IS_READING
    try:
        _IS_READING = True
        with _PROCESS_EXEC_LOCK:
            summary = processar_emails_enviados(
                max_messages_override=max_messages_override,
                messages_override=messages_override,
                preserve_reprocess_label=preserve_reprocess_label,
            )
        _process_finish(ok=True, error="")
        msg = _format_cycle_status("Execu\u00e7\u00e3o manual conclu\u00edda", summary)
        last_status = {"ok": True, "message": msg, "at": datetime.now().isoformat()}
        _NEXT_RUN_AT = time.time() + int(_RUNTIME_SETTINGS.get("interval_seconds", INTERVALO))
        return True, msg
    except Exception as exc:
        logger.exception("Erro na execu\u00e7\u00e3o manual: %s", exc)
        _process_finish(ok=False, error=str(exc))
        last_status = {"ok": False, "message": f"Erro na execu\u00e7\u00e3o manual: {exc}", "at": datetime.now().isoformat()}
        _NEXT_RUN_AT = time.time() + int(_RUNTIME_SETTINGS.get("interval_seconds", INTERVALO))
        return False, str(exc)
    finally:
        _IS_READING = False


def iniciar_verificacao():
    """Inicia o loop principal em thread separada."""
    global running, _LOOP_THREAD
    if running or (_LOOP_THREAD and _LOOP_THREAD.is_alive()):
        return False
    stop_event.clear()
    t = threading.Thread(target=main_loop, daemon=True, name="botana-loop")
    _LOOP_THREAD = t
    t.start()
    return True


def parar_verificacao(wait: bool = False, timeout: float = 30.0):
    """Interrompe o loop principal."""
    global running, _LOOP_THREAD
    thread = _LOOP_THREAD
    if not running and not (thread and thread.is_alive()):
        return False
    stop_event.set()
    running = False
    if wait and thread and thread.is_alive() and thread is not threading.current_thread():
        thread.join(max(0.1, float(timeout or 0.1)))
        if not thread.is_alive():
            _LOOP_THREAD = None
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


def _csv_response(
    handler: BaseHTTPRequestHandler,
    status: int,
    filename: str,
    rows: list[list[str]],
    delimiter: str = ";",
):
    buffer = io.StringIO(newline="")
    writer = csv.writer(buffer, delimiter=delimiter, quoting=csv.QUOTE_MINIMAL)
    for row in rows:
        writer.writerow(row)
    raw = buffer.getvalue().encode("utf-8-sig")
    safe_name = re.sub(r'[^A-Za-z0-9._-]+', "_", str(filename or "export.csv")).strip("._") or "export.csv"
    if not safe_name.lower().endswith(".csv"):
        safe_name += ".csv"
    handler.send_response(status)
    handler.send_header("Content-Type", "text/csv; charset=utf-8")
    handler.send_header("Content-Disposition", f'attachment; filename="{safe_name}"')
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

    files = _report_files_for_reading(include_all_txt=True)
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


def _report_base_name(path: Path) -> str:
    name = str(path.name or "").strip()
    return name[:-4] if name.endswith(".tmp") else name


def _format_currency_brl(value) -> str:
    try:
        number = float(value or 0)
    except Exception:
        number = 0.0
    text = f"{number:,.2f}"
    return "R$ " + text.replace(",", "X").replace(".", ",").replace("X", ".")


def _history_csv_rows(items: list[dict]) -> list[list[str]]:
    rows = [[
        "Data/Horário",
        "Vencimento",
        "NF",
        "Cliente",
        "Descrição",
        "Parcela",
        "Valor da Parcela",
        "Valor Total",
        "Valor Pago",
        "Status",
        "Aba",
        "CNPJ Emitente",
        "CNPJ Destinatário",
        "Duplicada",
    ]]
    for item in items or []:
        rows.append(
            [
                str(item.get("at") or ""),
                str(item.get("vencimento") or ""),
                str(item.get("nf") or ""),
                str(item.get("cliente") or ""),
                str(item.get("descricao") or ""),
                str(item.get("parcela") or ""),
                _format_currency_brl(item.get("valor_parcela")),
                _format_currency_brl(item.get("valor_total")),
                str(item.get("valor_pago") or ""),
                str(item.get("status") or ""),
                str(item.get("local_lancamento") or ""),
                str(item.get("cnpj_emit") or ""),
                str(item.get("cnpj_dest") or ""),
                "Sim" if bool(item.get("duplicata")) else "Não",
            ]
        )
    return rows


def _report_files_for_reading(include_all_txt: bool = False) -> list[Path]:
    rel_dir = Path(RELATORIO_DIR)
    if not rel_dir.exists():
        return []
    patterns = ("*.txt", "*.txt.tmp") if include_all_txt else ("relatorio_*.txt", "relatorio_*.txt.tmp")
    files: list[Path] = []
    seen = set()
    for pattern in patterns:
        for fp in rel_dir.glob(pattern):
            key = str(fp.resolve())
            if key in seen:
                continue
            seen.add(key)
            files.append(fp)
    return sorted(files, key=lambda p: p.stat().st_mtime, reverse=True)


def _latest_report_files() -> list[Path]:
    files = _report_files_for_reading(include_all_txt=False)
    if not files:
        return []
    latest_base = _report_base_name(files[0])
    grouped = [fp for fp in files if _report_base_name(fp) == latest_base]
    return sorted(grouped, key=lambda p: p.stat().st_mtime)


def _parse_report_datetime(dt_text: str):
    t = str(dt_text or "").strip()
    if not t:
        return None
    candidates = [
        (t[:19], "%Y-%m-%d %H:%M:%S"),
        (t[:19], "%Y-%m-%dT%H:%M:%S"),
        (t[:10], "%Y-%m-%d"),
        (t[:10], "%d/%m/%Y"),
    ]
    for cand, fmt in candidates:
        try:
            return datetime.strptime(cand, fmt)
        except Exception:
            continue
    try:
        return datetime.fromisoformat(t.replace("Z", "+00:00"))
    except Exception:
        return None


def _history_parcela_key(item: dict) -> str:
    parcela = _normalize_report_text(str((item or {}).get("parcela") or "")).strip()
    parcela_key = unicodedata.normalize("NFKD", parcela).encode("ascii", "ignore").decode("ascii")
    parcela_key = re.sub(r"\s+", " ", parcela_key).strip().upper()
    if parcela_key:
        patterns = (
            r"\b(\d+)\s*(?:A|O)?\s*PARC(?:ELA)?\b",
            r"\bPARC(?:ELA)?\s*(\d+)\b",
            r"\b(\d+)\s*/\s*\d+\b",
            r"\b(\d+)\s+DE\s+\d+\b",
            r"\b(\d+)\b",
        )
        for pattern in patterns:
            match = re.search(pattern, parcela_key)
            if not match:
                continue
            try:
                return f"parcela:{int(match.group(1))}"
            except Exception:
                continue
        return f"parcela:{parcela_key.lower()}"
    vencimento = _normalize_report_text(str((item or {}).get("vencimento") or "")).strip()
    valor = 0.0
    try:
        valor = round(float((item or {}).get("valor_parcela") or 0), 2)
    except Exception:
        valor = 0.0
    if vencimento or valor:
        return f"fallback:{vencimento}|{valor:.2f}"
    return f"raw:{_normalize_report_text(str((item or {}).get('at') or '')).strip()}"


_AUDIT_MONTH_MAP = {
    "jan": 1,
    "fev": 2,
    "feb": 2,
    "mar": 3,
    "abr": 4,
    "apr": 4,
    "mai": 5,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "ago": 8,
    "aug": 8,
    "set": 9,
    "sep": 9,
    "out": 10,
    "oct": 10,
    "nov": 11,
    "dez": 12,
    "dec": 12,
}


def _audit_safe_float(v):
    try:
        if isinstance(v, str):
            vv = v.strip().replace("\xa0", " ")
            if not vv:
                return 0.0
            vv = re.sub(r"[^\d,.\-]+", "", vv)
            if not vv or vv in {"-", ",", "."}:
                return 0.0
            if "," in vv and "." in vv:
                if vv.rfind(",") > vv.rfind("."):
                    vv = vv.replace(".", "").replace(",", ".")
                else:
                    vv = vv.replace(",", "")
            elif "," in vv:
                vv = vv.replace(",", ".")
            return float(vv)
        return float(v or 0)
    except Exception:
        return 0.0


def _audit_safe_int(v, default=0):
    try:
        if isinstance(v, str):
            vv = re.sub(r"[^\d-]+", "", v)
            return int(vv) if vv else int(default)
        return int(v)
    except Exception:
        return int(default)


def _parse_audit_date(dt_text: str):
    t = str(dt_text or "").strip()
    if not t:
        return None
    candidates = [
        (t[:10], "%d/%m/%Y"),
        (t[:10], "%Y-%m-%d"),
        (t[:10], "%d-%m-%Y"),
        (t[:10], "%d.%m.%Y"),
        (t[:19], "%Y-%m-%d %H:%M:%S"),
        (t[:19], "%Y-%m-%dT%H:%M:%S"),
    ]
    for cand, fmt in candidates:
        try:
            return datetime.strptime(cand, fmt)
        except Exception:
            continue
    try:
        return datetime.fromisoformat(t.replace("Z", "+00:00"))
    except Exception:
        return None


def _audit_month_key_from_value(value: str) -> str:
    dt = _parse_audit_date(value)
    return dt.strftime("%Y-%m") if dt else ""


def _audit_month_key_from_sheet_title(title: str) -> str:
    txt = str(title or "").strip()
    if not txt:
        return ""
    m = re.match(r"^\s*([^/]+?)\s*/\s*(\d{4})\s*$", txt)
    if not m:
        return ""
    month_token = unicodedata.normalize("NFKD", m.group(1)).encode("ascii", "ignore").decode("ascii").strip().lower()
    month_num = _AUDIT_MONTH_MAP.get(month_token[:3]) or _AUDIT_MONTH_MAP.get(month_token)
    if not month_num:
        return ""
    return f"{int(m.group(2)):04d}-{month_num:02d}"


def _audit_sheet_values(worksheet) -> list[list[str]]:
    from sheets_writer import apiCooldown

    for _ in range(3):
        try:
            return worksheet.get_all_values()
        except gspread.exceptions.APIError as exc:
            if "429" in str(exc):
                apiCooldown()
                continue
            raise
    return []


def _load_audit_sheet_rows() -> tuple[list[dict], dict]:
    creds = Credentials.from_service_account_file(
        GOOGLE_CREDENTIALS_SHEETS,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    gc = gspread.authorize(creds)
    rows = []
    planilhas_lidas = 0
    abas_lidas = 0

    for tipo_empresa, anos in PLANILHAS.items():
        for ano, planilha_id in (anos or {}).items():
            if not planilha_id:
                continue
            try:
                planilha = gc.open_by_key(planilha_id)
                planilhas_lidas += 1
            except Exception as exc:
                logger.warning("Falha ao abrir planilha %s %s para conferencia: %s", tipo_empresa, ano, exc)
                continue
            for worksheet in planilha.worksheets():
                abas_lidas += 1
                try:
                    linhas = _audit_sheet_values(worksheet)
                except Exception as exc:
                    logger.warning("Falha ao ler aba %s/%s: %s", getattr(planilha, "title", tipo_empresa), worksheet.title, exc)
                    continue
                worksheet_month = _audit_month_key_from_sheet_title(worksheet.title)
                for idx, linha in enumerate(linhas):
                    if idx == 0:
                        continue
                    row = list(linha or [])
                    if len(row) < 3:
                        continue
                    row += [""] * max(0, 9 - len(row))
                    nf_raw = str(row[2] or "").strip()
                    if not nf_raw:
                        continue
                    nf_digits = re.sub(r"\D+", "", nf_raw)
                    if not nf_digits:
                        continue
                    vencimento = str(row[0] or "").strip()
                    descricao = _normalize_report_text(str(row[1] or "").strip())
                    parcela = _normalize_report_text(str(row[5] or "").strip())
                    raw_cells = [_normalize_report_text(str(cell or "").strip()) for cell in row]
                    rows.append(
                        {
                            "group_key": f"{tipo_empresa}:{ano}:{nf_digits}",
                            "planilha_id": str(planilha_id or "").strip(),
                            "sheet_type": str(tipo_empresa or "").strip(),
                            "sheet_year": str(ano or "").strip(),
                            "sheet_title": _normalize_report_text(str(getattr(planilha, "title", "") or "").strip()),
                            "aba": _normalize_report_text(str(worksheet.title or "").strip()),
                            "worksheet_title": _normalize_report_text(str(worksheet.title or "").strip()),
                            "row_number": int(idx + 1),
                            "scope_month": _audit_month_key_from_value(vencimento) or worksheet_month,
                            "vencimento": _normalize_report_text(vencimento),
                            "descricao": descricao,
                            "cliente": descricao,
                            "nf": nf_digits,
                            "nf_num": _audit_safe_int(nf_digits, 0),
                            "valor_total_raw": _normalize_report_text(str(row[3] or "").strip()),
                            "valor_total": _audit_safe_float(row[3]),
                            "qtd_parcelas": max(1, _audit_safe_int(row[4], 1)),
                            "parcela": parcela,
                            "valor_parcela_raw": _normalize_report_text(str(row[6] or "").strip()),
                            "valor_parcela": _audit_safe_float(row[6]),
                            "valor_pago": _normalize_report_text(str(row[7] or "").strip()),
                            "status_planilha": _normalize_report_text(str(row[8] or "").strip()),
                            "raw_cells": raw_cells,
                        }
                    )

    meta = {
        "loaded_at": datetime.now().isoformat(),
        "planilhas_lidas": planilhas_lidas,
        "abas_lidas": abas_lidas,
        "linhas_lidas": len(rows),
        "source": "planilhas",
    }
    return rows, meta


def _audit_row_ref(item: dict) -> tuple[str, str, int]:
    return (
        str((item or {}).get("planilha_id") or "").strip(),
        str((item or {}).get("worksheet_title") or (item or {}).get("aba") or "").strip(),
        max(0, _audit_safe_int((item or {}).get("row_number"), 0)),
    )


def _audit_row_is_safe_delete(item: dict) -> bool:
    row_number = max(0, _audit_safe_int((item or {}).get("row_number"), 0))
    if row_number <= 1:
        return False
    if _sheet_watch_is_baixado(item):
        return False
    return _sheet_status_is_pending(str((item or {}).get("status_planilha") or ""))


def _audit_row_keep_key(item: dict):
    raw_cells = list((item or {}).get("raw_cells") or [])
    filled_cells = sum(1 for cell in raw_cells if str(cell or "").strip())
    descricao_len = len(str((item or {}).get("descricao") or "").strip())
    row_number = max(0, _audit_safe_int((item or {}).get("row_number"), 0)) or 999999
    safe_delete = _audit_row_is_safe_delete(item)
    status_value = str((item or {}).get("status_planilha") or "").strip()
    valor_pago = str((item or {}).get("valor_pago") or "").strip()
    return (
        0 if not safe_delete else 1,
        -filled_cells,
        -len(status_value),
        -len(valor_pago),
        row_number,
        -descricao_len,
    )


def _audit_identity_order_key(item: dict):
    identity = str(_history_parcela_key(item) or "").strip()
    parcela_num = 999999
    if identity.startswith("parcela:"):
        try:
            parcela_num = int(identity.split(":", 1)[1])
        except Exception:
            parcela_num = 999999
    venc_dt = _parse_audit_date(str((item or {}).get("vencimento") or "").strip())
    venc_key = venc_dt.strftime("%Y-%m-%d") if venc_dt else "9999-12-31"
    row_number = max(0, _audit_safe_int((item or {}).get("row_number"), 0)) or 999999
    return (parcela_num, venc_key, row_number, identity)


def _audit_delete_candidates(nf_items: list[dict], qtd_esperada: int) -> list[dict]:
    identities = {}
    removable = {}
    for item in list(nf_items or []):
        identity = str(_history_parcela_key(item) or "").strip()
        if not identity:
            continue
        identities.setdefault(identity, []).append(item)

    keepers = {}
    for identity, rows in identities.items():
        ordered = sorted(rows, key=_audit_row_keep_key)
        if not ordered:
            continue
        keepers[identity] = ordered[0]
        for extra in ordered[1:]:
            if _audit_row_is_safe_delete(extra):
                removable[_audit_row_ref(extra)] = extra

    if qtd_esperada > 0 and len(keepers) > qtd_esperada:
        ordered_identities = sorted(keepers.items(), key=lambda pair: _audit_identity_order_key(pair[1]))
        keep_identity_keys = {identity for identity, _ in ordered_identities[:qtd_esperada]}
        for identity, keeper in ordered_identities[qtd_esperada:]:
            if identity in keep_identity_keys:
                continue
            if _audit_row_is_safe_delete(keeper):
                removable[_audit_row_ref(keeper)] = keeper

    rows = list(removable.values())
    rows.sort(
        key=lambda item: (
            str(item.get("planilha_id") or "").strip(),
            str(item.get("worksheet_title") or item.get("aba") or "").strip(),
            -max(0, _audit_safe_int(item.get("row_number"), 0)),
        )
    )
    return rows


def _delete_audit_rows(audit_key: str) -> dict:
    target_key = str(audit_key or "").strip()
    if not target_key:
        return {"ok": False, "message": "NF da conferência não informada."}

    linhas, _ = _load_audit_sheet_rows()
    nf_items = [item for item in linhas if str(item.get("group_key") or "").strip() == target_key]
    if not nf_items:
        return {"ok": False, "message": "NF não encontrada na conferência atual."}

    qtd_esperada = 0
    for item in nf_items:
        qtd_esperada = max(qtd_esperada, max(1, _audit_safe_int(item.get("qtd_parcelas"), 1)))

    candidatos = _audit_delete_candidates(nf_items, qtd_esperada)
    if not candidatos:
        nf = str((nf_items[0] or {}).get("nf") or "").strip()
        return {
            "ok": False,
            "message": f"A NF {nf or '-'} não possui linhas pendentes removíveis automaticamente.",
            "deleted": 0,
        }

    creds = Credentials.from_service_account_file(
        GOOGLE_CREDENTIALS_SHEETS,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    gc = gspread.authorize(creds)
    from sheets_writer import apiCooldown

    book_cache = {}
    sheet_cache = {}
    deleted = []

    for item in candidatos:
        planilha_id, worksheet_title, row_number = _audit_row_ref(item)
        if not planilha_id or not worksheet_title or row_number <= 1:
            continue
        if planilha_id not in book_cache:
            book_cache[planilha_id] = gc.open_by_key(planilha_id)
        cache_key = (planilha_id, worksheet_title)
        if cache_key not in sheet_cache:
            sheet_cache[cache_key] = book_cache[planilha_id].worksheet(worksheet_title)
        worksheet = sheet_cache[cache_key]
        for _ in range(3):
            try:
                worksheet.delete_rows(row_number)
                deleted.append(
                    {
                        "nf": str(item.get("nf") or "").strip(),
                        "aba": worksheet_title,
                        "row_number": row_number,
                        "parcela": str(item.get("parcela") or "").strip(),
                    }
                )
                break
            except gspread.exceptions.APIError as exc:
                if "429" in str(exc):
                    apiCooldown()
                    continue
                raise

    nf = str((nf_items[0] or {}).get("nf") or "").strip()
    return {
        "ok": bool(deleted),
        "message": (
            f"{len(deleted)} linha(s) removida(s) da planilha para a NF {nf}."
            if deleted
            else f"Nenhuma linha foi removida para a NF {nf}."
        ),
        "deleted": len(deleted),
        "items": deleted,
    }


def _gerar_conferencia_parcelas(filtro: str, mes: str, nf_inicio: str, nf_fim: str) -> dict:
    filtro_normalizado = str(filtro or "mes").strip().lower()
    if filtro_normalizado not in {"mes", "nfs", "todos"}:
        filtro_normalizado = "mes"
    linhas, meta = _load_audit_sheet_rows()
    grupos = {}
    for item in linhas:
        group_key = str(item.get("group_key") or "").strip()
        if not group_key:
            continue
        grupos.setdefault(group_key, []).append(item)

    try:
        nf_inicio_num = int(re.sub(r"\D+", "", str(nf_inicio or ""))) if str(nf_inicio or "").strip() else None
    except Exception:
        nf_inicio_num = None
    try:
        nf_fim_num = int(re.sub(r"\D+", "", str(nf_fim or ""))) if str(nf_fim or "").strip() else None
    except Exception:
        nf_fim_num = None
    if nf_inicio_num is not None and nf_fim_num is not None and nf_inicio_num > nf_fim_num:
        nf_inicio_num, nf_fim_num = nf_fim_num, nf_inicio_num

    def _matches_scope(nf_items: list[dict]) -> bool:
        if not nf_items:
            return False
        nf_num = _audit_safe_int((nf_items[0] or {}).get("nf_num"), 0)
        if filtro_normalizado == "todos":
            return True
        if filtro_normalizado == "nfs":
            if nf_inicio_num is not None and nf_num < nf_inicio_num:
                return False
            if nf_fim_num is not None and nf_num > nf_fim_num:
                return False
            return True
        if not mes:
            return True
        return any(str(entry.get("scope_month") or "").strip() == mes for entry in nf_items)

    itens_saida = []
    resumo = {
        "nfs_verificadas": 0,
        "nfs_ok": 0,
        "nfs_com_divergencia": 0,
        "parcelas_esperadas": 0,
        "parcelas_lancadas": 0,
        "parcelas_duplicadas": 0,
    }

    for _, nf_items in grupos.items():
        if not _matches_scope(nf_items):
            continue

        nf = str((nf_items[0] or {}).get("nf") or "").strip()
        parcelas_contagem = {}
        cliente = ""
        descricao = ""
        aba_principal = ""
        local = ""
        abas = []
        vencimentos = set()
        ultimo_venc_dt = None
        qtd_esperada = 0

        for item in nf_items:
            chave_parcela = _history_parcela_key(item)
            parcelas_contagem[chave_parcela] = parcelas_contagem.get(chave_parcela, 0) + 1
            qtd_item = max(1, _audit_safe_int(item.get("qtd_parcelas"), 1))
            qtd_esperada = max(qtd_esperada, qtd_item)
            if not cliente:
                cliente = str(item.get("cliente") or "").strip()
            if not descricao:
                descricao = str(item.get("descricao") or "").strip()
            aba_item = str(item.get("aba") or "").strip()
            if aba_item and aba_item not in abas:
                abas.append(aba_item)
            if not aba_principal:
                aba_principal = aba_item
            if not local:
                local = "/".join(
                    [
                        x
                        for x in (
                            str(item.get("sheet_type") or "").strip(),
                            aba_item,
                        )
                        if x
                    ]
                )
            venc = str(item.get("vencimento") or "").strip()
            if venc:
                vencimentos.add(venc)
                venc_dt = _parse_audit_date(venc)
                if venc_dt and (ultimo_venc_dt is None or venc_dt > ultimo_venc_dt):
                    ultimo_venc_dt = venc_dt

        qtd_lancada = len(parcelas_contagem)
        qtd_bruta = sum(parcelas_contagem.values())
        qtd_duplicada = max(0, qtd_bruta - qtd_lancada)
        qtd_faltando = max(0, qtd_esperada - qtd_lancada)
        qtd_excedente = max(0, qtd_lancada - qtd_esperada)
        delete_candidates = _audit_delete_candidates(nf_items, qtd_esperada)
        abas_view = sorted(abas, key=lambda value: (_audit_month_key_from_sheet_title(value), value))
        if len(abas_view) > 1:
            aba_view = f"{abas_view[0]} +{len(abas_view) - 1}"
        else:
            aba_view = abas_view[0] if abas_view else aba_principal
        local_view = " - ".join(
            [
                x
                for x in (
                    str((nf_items[0] or {}).get("sheet_type") or "").strip(),
                    aba_view,
                )
                if x
            ]
        ) or local
        ultimo_vencimento = ultimo_venc_dt.strftime("%Y-%m-%d") if ultimo_venc_dt else ""

        duplicadas = []
        for item in nf_items:
            chave_parcela = _history_parcela_key(item)
            label = _normalize_report_text(str(item.get("parcela") or item.get("vencimento") or "-")).strip() or "-"
            if parcelas_contagem.get(chave_parcela, 0) > 1 and label not in duplicadas:
                duplicadas.append(label)

        if qtd_faltando == 0 and qtd_excedente == 0 and qtd_duplicada == 0:
            status = "ok"
            status_label = "OK"
        elif qtd_faltando > 0 and qtd_duplicada > 0:
            status = "erro"
            status_label = "Faltando + duplicada"
        elif qtd_faltando > 0:
            status = "erro"
            status_label = "Faltando"
        elif qtd_excedente > 0:
            status = "erro"
            status_label = "A mais"
        else:
            status = "erro"
            status_label = "Duplicada"

        itens_saida.append(
            {
                "audit_key": str((nf_items[0] or {}).get("group_key") or "").strip(),
                "nf": nf,
                "cliente": _normalize_report_text(cliente),
                "descricao": _normalize_report_text(descricao),
                "qtd_esperada": qtd_esperada,
                "qtd_lancada": qtd_lancada,
                "qtd_bruta": qtd_bruta,
                "qtd_faltando": qtd_faltando,
                "qtd_excedente": qtd_excedente,
                "qtd_duplicada": qtd_duplicada,
                "parcelas_duplicadas": duplicadas,
                "vencimentos": sorted(vencimentos),
                "ultimo_vencimento": ultimo_vencimento,
                "aba": _normalize_report_text(aba_view),
                "local_lancamento": _normalize_report_text(local_view),
                "status": status,
                "status_label": status_label,
                "delete_candidates": len(delete_candidates),
                "can_delete_rows": bool(delete_candidates),
            }
        )

        resumo["nfs_verificadas"] += 1
        resumo["parcelas_esperadas"] += qtd_esperada
        resumo["parcelas_lancadas"] += qtd_lancada
        resumo["parcelas_duplicadas"] += qtd_duplicada
        if status == "ok":
            resumo["nfs_ok"] += 1
        else:
            resumo["nfs_com_divergencia"] += 1

    itens_saida.sort(
        key=lambda item: (
            0 if item.get("status") == "erro" else 1 if item.get("status") == "aviso" else 2,
            -int(re.sub(r"\D+", "", str(item.get("nf") or "0")) or 0),
        )
    )

    return {
        "filtro": filtro_normalizado,
        "mes": mes,
        "nf_inicio": nf_inicio_num,
        "nf_fim": nf_fim_num,
        "summary": resumo,
        "meta": meta,
        "items": itens_saida,
    }


def _normalize_ascii_key(value: str) -> str:
    txt = unicodedata.normalize("NFKD", str(value or "")).encode("ascii", "ignore").decode("ascii")
    txt = re.sub(r"\s+", " ", txt).strip().upper()
    return txt


def _normalize_watch_search_name(value: str) -> str:
    txt = _normalize_report_text(str(value or "").strip())
    txt = re.sub(r"\s+", " ", txt).strip()
    return txt[:160]


def _load_watch_search_names() -> list[str]:
    with _WATCH_SEARCH_NAMES_LOCK:
        if not _WATCH_SEARCH_NAMES_FILE.exists():
            return []
        try:
            lines = _WATCH_SEARCH_NAMES_FILE.read_text(encoding="utf-8", errors="ignore").splitlines()
        except Exception:
            return []
        out = []
        seen = set()
        for raw in lines:
            name = _normalize_watch_search_name(raw)
            if not name:
                continue
            key = _normalize_ascii_key(name)
            if not key or key in seen:
                continue
            seen.add(key)
            out.append(name)
        return out


def _remember_watch_search_names(*values: str):
    incoming = []
    for raw in values:
        name = _normalize_watch_search_name(raw)
        if name:
            incoming.append(name)
    if not incoming:
        return
    with _WATCH_SEARCH_NAMES_LOCK:
        current = []
        if _WATCH_SEARCH_NAMES_FILE.exists():
            try:
                current = _WATCH_SEARCH_NAMES_FILE.read_text(encoding="utf-8", errors="ignore").splitlines()
            except Exception:
                current = []
        merged = []
        seen = set()
        for raw in incoming + current:
            name = _normalize_watch_search_name(raw)
            if not name:
                continue
            key = _normalize_ascii_key(name)
            if not key or key in seen:
                continue
            seen.add(key)
            merged.append(name)
            if len(merged) >= 200:
                break
        try:
            _WATCH_SEARCH_NAMES_FILE.write_text("\n".join(merged) + ("\n" if merged else ""), encoding="utf-8")
        except Exception as exc:
            logger.warning("Falha ao salvar autocomplete de nomes da busca de prazos: %s", exc)


def _load_watch_search_suggestions() -> list[str]:
    out = []
    seen = set()

    def add_name(value: str):
        name = _normalize_watch_search_name(value)
        if not name:
            return
        key = _normalize_ascii_key(name)
        if not key or key in seen:
            return
        seen.add(key)
        out.append(name)

    for raw in _load_watch_search_names():
        add_name(raw)

    try:
        linhas, _ = _load_audit_sheet_rows()
    except Exception as exc:
        logger.warning("Falha ao montar catálogo de autocomplete da busca de prazos: %s", exc)
        return out[:500]

    for item in linhas:
        if _sheet_watch_is_baixado(item):
            continue
        if not _sheet_status_is_pending(item.get("status_planilha")):
            continue
        if _sheet_watch_kind(item.get("descricao")) != "boleto":
            continue
        add_name(str(item.get("cliente") or item.get("descricao") or "").strip())

    return out[:500]


def _business_days_distance(from_date, to_date) -> int:
    if not from_date or not to_date or from_date == to_date:
        return 0
    step = 1 if to_date > from_date else -1
    cur = from_date
    total = 0
    while cur != to_date:
        cur += timedelta(days=step)
        if cur.weekday() < 5:
            total += step
    return total


def _sheet_status_is_pending(status_value: str) -> bool:
    key = _normalize_ascii_key(status_value)
    return not key or key == "A RECEBER"


def _sheet_watch_is_baixado(item: dict) -> bool:
    ignore_markers = ("BAIXADO", "BAIXADA", "ESTORNADO", "ESTORNADA")
    for field in ("descricao", "cliente", "parcela", "valor_pago", "status_planilha", "valor_total_raw", "valor_parcela_raw"):
        key = _normalize_ascii_key((item or {}).get(field, ""))
        if any(marker in key for marker in ignore_markers):
            return True
    for cell in (item or {}).get("raw_cells") or []:
        key = _normalize_ascii_key(cell)
        if any(marker in key for marker in ignore_markers):
            return True
    return False


def _sheet_watch_kind(descricao: str) -> str:
    key = _normalize_ascii_key(descricao)
    if not key:
        return ""
    if "BLT" in key or "BOLETO" in key:
        return "boleto"
    if re.search(r"\bDEP\b", key):
        return "deposito"
    return ""


def _format_days_label(days: int, future: bool = False) -> str:
    qtd = abs(int(days or 0))
    if qtd == 0:
        return "Hoje"
    unidade = "dia" if qtd == 1 else "dias"
    if future:
        return f"Em {qtd} {unidade}"
    return f"{qtd} {unidade} atrasado" if qtd == 1 else f"{qtd} {unidade} atrasados"


def _watch_boleto_status_payload(dias_uteis: int) -> tuple[str, str, str]:
    if dias_uteis > 0:
        return "aviso", "A vencer", _format_days_label(dias_uteis, future=True)
    if dias_uteis == 0:
        return "erro", "Vence hoje", "Hoje"
    atraso = abs(dias_uteis)
    return "erro", "Vencido", _format_days_label(atraso, future=False)


def _gerar_relacao_pendencias(boletos_dias: int, depositos_dias: int) -> dict:
    boleto_limit = max(1, min(7, _audit_safe_int(boletos_dias, 7)))
    deposito_limit = max(1, min(7, _audit_safe_int(depositos_dias, 7)))
    linhas, meta = _load_audit_sheet_rows()
    hoje = datetime.now().date()
    itens = []
    resumo = {
        "total_itens": 0,
        "boletos_a_vencer": 0,
        "boletos_vencidos": 0,
        "depositos_atrasados": 0,
    }

    for item in linhas:
        if _sheet_watch_is_baixado(item):
            continue
        if not _sheet_status_is_pending(item.get("status_planilha")):
            continue
        tipo = _sheet_watch_kind(item.get("descricao"))
        if not tipo:
            continue
        venc_dt = _parse_audit_date(str(item.get("vencimento") or "").strip())
        if not venc_dt:
            continue
        venc_date = venc_dt.date()
        dias_uteis = _business_days_distance(hoje, venc_date)
        if tipo == "boleto":
            if dias_uteis > boleto_limit:
                continue
            status, status_label, dias_label = _watch_boleto_status_payload(dias_uteis)
            if dias_uteis > 0:
                resumo["boletos_a_vencer"] += 1
            else:
                resumo["boletos_vencidos"] += 1
            tipo_label = "Boleto"
        else:
            if dias_uteis > -deposito_limit:
                continue
            atraso = abs(dias_uteis)
            status = "erro"
            status_label = "Dep\u00f3sito atrasado"
            dias_label = _format_days_label(atraso, future=False)
            tipo_label = "Dep\u00f3sito"
            resumo["depositos_atrasados"] += 1

        valor_base = _audit_safe_float(item.get("valor_parcela"))
        if valor_base <= 0:
            valor_base = _audit_safe_float(item.get("valor_total"))
        nf = str(item.get("nf") or "").strip()
        local = " - ".join(
            [x for x in (str(item.get("sheet_type") or "").strip(), str(item.get("aba") or "").strip()) if x]
        )
        itens.append(
            {
                "tipo": tipo,
                "tipo_label": tipo_label,
                "status": status,
                "status_label": status_label,
                "dias_uteis": dias_uteis,
                "dias_label": dias_label,
                "vencimento": str(item.get("vencimento") or "").strip(),
                "descricao": _normalize_report_text(str(item.get("descricao") or "").strip()),
                "cliente": _normalize_report_text(str(item.get("cliente") or "").strip()),
                "nf": nf,
                "valor": valor_base,
                "aba": _normalize_report_text(str(item.get("aba") or "").strip()),
                "local": _normalize_report_text(local),
                "status_planilha": _normalize_report_text(str(item.get("status_planilha") or "").strip()),
                "_sort_date": venc_date.toordinal(),
                "_sort_nf": _audit_safe_int(nf, 0),
            }
        )
        resumo["total_itens"] += 1

    itens.sort(
        key=lambda row: (
            0 if row.get("status") == "erro" else 1,
            int(row.get("_sort_date") or 0),
            0 if row.get("tipo") == "boleto" else 1,
            -int(row.get("_sort_nf") or 0),
        )
    )
    for row in itens:
        row.pop("_sort_date", None)
        row.pop("_sort_nf", None)

    return {
        "summary": resumo,
        "meta": {**meta, "loaded_at": datetime.now().isoformat()},
        "limits": {"boletos_dias": boleto_limit, "depositos_dias": deposito_limit},
        "items": itens,
    }


def _buscar_boletos_em_aberto_por_nome(nome: str) -> dict:
    nome_busca = _normalize_report_text(str(nome or "").strip())
    termo = _normalize_ascii_key(nome_busca)
    if not termo:
        raise ValueError("Informe um nome para buscar.")
    linhas, meta = _load_audit_sheet_rows()
    hoje = datetime.now().date()
    tokens = [item for item in termo.split(" ") if item]
    itens = []
    suggestions = []
    suggestions_seen = set()

    def add_suggestion(value: str):
        name = _normalize_watch_search_name(value)
        if not name:
            return
        key = _normalize_ascii_key(name)
        if not key or key in suggestions_seen:
            return
        suggestions_seen.add(key)
        suggestions.append(name)

    for raw in _load_watch_search_names():
        add_suggestion(raw)

    for item in linhas:
        if _sheet_watch_is_baixado(item):
            continue
        if not _sheet_status_is_pending(item.get("status_planilha")):
            continue
        if _sheet_watch_kind(item.get("descricao")) != "boleto":
            continue
        add_suggestion(str(item.get("cliente") or item.get("descricao") or "").strip())
        haystack = _normalize_ascii_key(
            " ".join(
                [
                    str(item.get("cliente") or "").strip(),
                    str(item.get("descricao") or "").strip(),
                    str(item.get("nf") or "").strip(),
                ]
            )
        )
        if not haystack or not all(token in haystack for token in tokens):
            continue
        venc_dt = _parse_audit_date(str(item.get("vencimento") or "").strip())
        if not venc_dt:
            continue
        venc_date = venc_dt.date()
        dias_uteis = _business_days_distance(hoje, venc_date)
        status, status_label, dias_label = _watch_boleto_status_payload(dias_uteis)
        valor_base = _audit_safe_float(item.get("valor_parcela"))
        if valor_base <= 0:
            valor_base = _audit_safe_float(item.get("valor_total"))
        nf = str(item.get("nf") or "").strip()
        local = " - ".join(
            [x for x in (str(item.get("sheet_type") or "").strip(), str(item.get("aba") or "").strip()) if x]
        )
        itens.append(
            {
                "tipo": "boleto",
                "tipo_label": "Boleto",
                "status": status,
                "status_label": status_label,
                "dias_uteis": dias_uteis,
                "dias_label": dias_label,
                "vencimento": str(item.get("vencimento") or "").strip(),
                "descricao": _normalize_report_text(str(item.get("descricao") or "").strip()),
                "cliente": _normalize_report_text(str(item.get("cliente") or "").strip()),
                "nf": nf,
                "valor": valor_base,
                "aba": _normalize_report_text(str(item.get("aba") or "").strip()),
                "local": _normalize_report_text(local),
                "status_planilha": _normalize_report_text(str(item.get("status_planilha") or "").strip()),
                "_sort_date": venc_date.toordinal(),
                "_sort_nf": _audit_safe_int(nf, 0),
            }
        )
    itens.sort(
        key=lambda row: (
            0 if row.get("status") == "erro" else 1,
            int(row.get("_sort_date") or 0),
            -int(row.get("_sort_nf") or 0),
        )
    )
    for row in itens:
        row.pop("_sort_date", None)
        row.pop("_sort_nf", None)
    count = len(itens)
    related_names = [nome_busca]
    related_names.extend(str(item.get("cliente") or "").strip() for item in itens)
    _remember_watch_search_names(*related_names)
    for raw in related_names:
        add_suggestion(raw)
    if count <= 0:
        message = f"Não existem pendências para '{nome_busca}'."
    elif count == 1:
        message = f"Foi encontrado 1 boleto em aberto para '{nome_busca}'."
    else:
        message = f"Foram encontrados {count} boletos em aberto para '{nome_busca}'."
    return {
        "query": nome_busca,
        "count": count,
        "message": message,
        "meta": {**meta, "loaded_at": datetime.now().isoformat()},
        "items": itens,
        "suggestions": suggestions[:500],
    }


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

    arquivos = _report_files_for_reading(include_all_txt=False)
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


def _daily_report_data() -> dict:
    report_files = _latest_report_files()
    if not report_files:
        return {
            "exists": False,
            "path": "",
            "updated_at": "",
            "totals": {"processados": 0, "ignorados": 0, "avisos_ciclo": 0, "avisos_dia": 0},
            "processados": [],
            "ignorados": [],
            "avisos": [],
        }

    lines = []
    for report_file in report_files:
        try:
            lines.extend(report_file.read_text(encoding="utf-8", errors="ignore").splitlines())
        except Exception:
            continue

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
    latest_mtime = max(fp.stat().st_mtime for fp in report_files)
    updated_at = datetime.fromtimestamp(latest_mtime).strftime("%d/%m/%Y, %H:%M:%S")
    return {
        "exists": True,
        "path": " | ".join(str(fp) for fp in report_files),
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


def _processing_abort_requested() -> bool:
    return bool(stop_event.is_set())


def _wait_for_processing_idle(timeout: float = 30.0) -> bool:
    deadline = time.time() + max(0.5, float(timeout or 0))
    while time.time() < deadline:
        if not _reading_active():
            return True
        time.sleep(0.2)
    return not _reading_active()


def _parse_iso_date_input(value: str):
    txt = str(value or "").strip()
    if not txt:
        return None
    try:
        return datetime.strptime(txt, "%Y-%m-%d").date()
    except Exception:
        return None


def _join_gmail_query(*parts: str) -> str:
    return " ".join(str(part or "").strip() for part in parts if str(part or "").strip())


def _cycle_search_query() -> str:
    return build_sent_xml_query(filter_mode=str(_RUNTIME_SETTINGS.get("gmail_filter_mode", "last_30_days")))


def _recover_nf_query_terms(nf_start: str = "", nf_end: str = "") -> str:
    start_nf, end_nf = _parse_nf_filter_range(nf_start, nf_end)
    if start_nf is None or end_nf is None:
        return ""
    if start_nf == end_nf:
        return str(start_nf)
    if end_nf - start_nf <= 20:
        return "{" + " ".join(str(numero) for numero in range(start_nf, end_nf + 1)) + "}"
    return ""


def _recover_search_query(nf_start: str = "", nf_end: str = "", date_from: str = "", date_to: str = "") -> str:
    extra_parts = []
    start_date = _parse_iso_date_input(date_from)
    end_date = _parse_iso_date_input(date_to)
    if start_date:
        extra_parts.append(f"after:{start_date.strftime('%Y/%m/%d')}")
    if end_date:
        extra_parts.append(f"before:{(end_date + timedelta(days=1)).strftime('%Y/%m/%d')}")
    nf_query = _recover_nf_query_terms(nf_start=nf_start, nf_end=nf_end)
    if nf_query:
        extra_parts.append(nf_query)
    return _join_gmail_query(build_sent_xml_query(filter_mode="", extra_query=""), " ".join(extra_parts))


def _parse_nf_filter_range(nf_start: str = "", nf_end: str = "") -> tuple[int | None, int | None]:
    start_txt = re.sub(r"\D+", "", str(nf_start or "").strip())
    end_txt = re.sub(r"\D+", "", str(nf_end or "").strip())
    start_val = int(start_txt) if start_txt else None
    end_val = int(end_txt) if end_txt else None
    if start_val is not None and end_val is None:
        end_val = start_val
    elif end_val is not None and start_val is None:
        start_val = end_val
    if start_val is not None and end_val is not None and start_val > end_val:
        start_val, end_val = end_val, start_val
    return start_val, end_val


def _extract_nf_numbers_from_text(text: str) -> list[int]:
    values = []
    for raw in re.findall(r"\bNF\s*0*(\d{3,})\b", str(text or "").upper()):
        try:
            values.append(int(raw))
        except Exception:
            continue
    return sorted(set(values))


def _preview_matches_nf_range(subject: str, nf_start: int | None, nf_end: int | None) -> bool:
    if nf_start is None or nf_end is None:
        return True
    numeros = _extract_nf_numbers_from_text(subject)
    if not numeros:
        return False
    return any(nf_start <= numero <= nf_end for numero in numeros)


def _manual_scan_page_limit(wanted: int, page_size: int) -> int:
    runtime_pages = max(1, min(50, int(_RUNTIME_SETTINGS.get("gmail_max_pages", 3) or 3)))
    estimated = max(1, (max(1, int(wanted or 1)) + max(1, page_size) - 1) // max(1, page_size))
    return max(runtime_pages, min(50, max(10, estimated * 4)))


def _describe_recovery_filters(nf_start: str = "", nf_end: str = "", date_from: str = "", date_to: str = "") -> str:
    parts = []
    start_nf, end_nf = _parse_nf_filter_range(nf_start, nf_end)
    if start_nf is not None and end_nf is not None:
        parts.append(f"NFs {start_nf} a {end_nf}" if start_nf != end_nf else f"NF {start_nf}")
    start_date = _parse_iso_date_input(date_from)
    end_date = _parse_iso_date_input(date_to)
    if start_date and end_date:
        parts.append(f"período {start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')}")
    elif start_date:
        parts.append(f"a partir de {start_date.strftime('%d/%m/%Y')}")
    elif end_date:
        parts.append(f"até {end_date.strftime('%d/%m/%Y')}")
    return " | ".join(parts)


def _reprocess_message_preview(service, msg_id: str) -> dict:
    preview = {"email": "", "subject": "", "date": "", "timestamp": 0}
    if not msg_id:
        return preview
    try:
        meta = service.users().messages().get(
            userId="me",
            id=msg_id,
            format="metadata",
            metadataHeaders=["From", "Subject", "Date"],
        ).execute()
        headers = ((meta or {}).get("payload") or {}).get("headers", []) or []
        header_map = {}
        for item in headers:
            name = str((item or {}).get("name", "")).strip().lower()
            if not name or name in header_map:
                continue
            header_map[name] = str((item or {}).get("value", "")).strip()
        from_raw = str(header_map.get("from", "")).strip()
        subject = str(header_map.get("subject", "")).strip()
        date_raw = str(header_map.get("date", "")).strip()
        internal_ts = int(str(meta.get("internalDate", "0") or "0").strip() or "0")
        _, email_addr = parseaddr(from_raw)
        email_view = email_addr or from_raw
        date_view = date_raw
        stamp = int(internal_ts or 0)
        if date_raw:
            try:
                dt = parsedate_to_datetime(date_raw)
                if dt.tzinfo is not None:
                    dt = dt.astimezone()
                date_view = dt.strftime("%d/%m/%Y %H:%M")
                if not stamp:
                    stamp = int(dt.timestamp() * 1000)
            except Exception:
                pass
        preview.update({"email": email_view, "subject": subject, "date": date_view, "timestamp": stamp})
    except Exception as exc:
        logger.warning("Falha ao carregar cabecalhos da mensagem %s: %s", msg_id, exc)
    return preview


def _reprocess_recent(max_messages: int, mark_unread: bool, progress_cb=None) -> dict:
    service = _get_gmail_service_locked()
    wanted = max(1, min(1000, int(max_messages)))
    messages_raw = listar_mensagens_com_labels_botana(service, max_results=1000)
    mensagens_com_meta = []
    for item in messages_raw:
        msg_id = str(item.get("id", "")).strip()
        if not msg_id:
            continue
        preview = _reprocess_message_preview(service, msg_id)
        mensagens_com_meta.append(
            {
                "id": msg_id,
                "threadId": str(item.get("threadId", "")).strip(),
                "botana_label": str(item.get("botana_label", "")).strip(),
                "preview": preview,
            }
        )
    mensagens_com_meta.sort(key=lambda item: int((item.get("preview") or {}).get("timestamp", 0) or 0), reverse=True)
    messages = mensagens_com_meta[:wanted]
    changed = 0
    failed = 0
    targets = []
    if callable(progress_cb):
        progress_cb(
            progress_current=0,
            progress_total=len(messages),
            changed=changed,
            failed=failed,
            current_email="",
            current_subject="",
            current_date="",
            message=f"{len(messages)} mensagens mais recentes encontradas para reprocessar.",
            detail="Atualizando a label do Botana e exibindo o remetente/data da mensagem atual.",
        )
    for idx, item in enumerate(messages, start=1):
        msg_id = str(item.get("id", "")).strip()
        if not msg_id:
            continue
        preview = dict(item.get("preview") or {})
        current_email = str(preview.get("email", "")).strip()
        current_subject = str(preview.get("subject", "")).strip()
        current_date = str(preview.get("date", "")).strip()
        current_label = str(item.get("botana_label", "")).strip()
        if callable(progress_cb):
            progress_cb(
                progress_current=changed + failed,
                progress_total=len(messages),
                changed=changed,
                failed=failed,
                current_email=current_email,
                current_subject=current_subject,
                current_date=current_date,
                message=f"Analisando mensagens: {idx} de {len(messages)}.",
                detail=(f"Label atual: {current_label} | Assunto: {current_subject}" if current_subject else f"Label atual: {current_label}"),
            )
        novo_label = ""
        try:
            novo_label = marcar_mensagem_para_reprocessar(service, msg_id, mark_unread=mark_unread)
            if not novo_label:
                raise RuntimeError("Falha ao atualizar label de reprocessamento")
            changed += 1
            targets.append(
                {
                    "id": msg_id,
                    "threadId": str(item.get("threadId", "")).strip(),
                    "labelIds": [],
                    "snippet": "",
                }
            )
        except Exception:
            failed += 1
        if callable(progress_cb):
            progress_cb(
                progress_current=changed + failed,
                progress_total=len(messages),
                changed=changed,
                failed=failed,
                current_email=current_email,
                current_subject=current_subject,
                current_date=current_date,
                message=f"Reprocessando mensagens: {changed + failed} de {len(messages)}.",
                detail=(f"Nova label: {novo_label} | Assunto: {current_subject}" if novo_label and current_subject else (f"Nova label: {novo_label}" if novo_label else f"Atualizadas: {changed} | Falhas: {failed}")),
            )
    return {
        "ok": True,
        "matched": len(messages),
        "changed": changed,
        "failed": failed,
        "mark_unread": bool(mark_unread),
        "targets": targets,
    }


def _find_missing_messages(
    max_messages: int,
    nf_start: str = "",
    nf_end: str = "",
    date_from: str = "",
    date_to: str = "",
    progress_cb=None,
) -> dict:
    start_nf, end_nf = _parse_nf_filter_range(nf_start, nf_end)
    start_date = _parse_iso_date_input(date_from)
    end_date = _parse_iso_date_input(date_to)
    if start_nf is None and end_nf is None and not start_date and not end_date:
        raise ValueError("Informe uma faixa de NF e/ou um período para recuperar.")
    service = _get_gmail_service_locked()
    wanted = max(1, min(1000, int(max_messages)))
    page_size = max(1, min(500, int(_RUNTIME_SETTINGS.get("gmail_page_size", 50) or 50)))
    page_limit = _manual_scan_page_limit(wanted, page_size)
    query = _recover_search_query(nf_start=nf_start, nf_end=nf_end, date_from=date_from, date_to=date_to)
    skip_botana_ids = {str(label_id).strip() for label_id in list_botana_label_ids(service) if str(label_id).strip()}
    targets = []
    seen_ids = set()
    inspected = 0
    pages = 0
    page_token = None
    criteria_desc = _describe_recovery_filters(nf_start=nf_start, nf_end=nf_end, date_from=date_from, date_to=date_to) or "filtros informados"
    if callable(progress_cb):
        progress_cb(
            phase="searching",
            progress_current=0,
            progress_total=wanted,
            matched=0,
            inspected=0,
            current_email="",
            current_subject="",
            current_date="",
            message="Varrendo Gmail em busca de mensagens sem label do Botana.",
            detail=f"Critérios: {criteria_desc}.",
        )
    while pages < page_limit and len(targets) < wanted:
        batch, next_page_token = buscarMessagesEnviadosPagina(
            service,
            max_results=page_size,
            page_token=page_token,
            query=query,
            skip_label_ids=skip_botana_ids,
        )
        pages += 1
        if not batch and not next_page_token:
            break
        for item in batch:
            msg_id = str(item.get("id", "")).strip()
            if not msg_id or msg_id in seen_ids:
                continue
            seen_ids.add(msg_id)
            preview = _reprocess_message_preview(service, msg_id)
            inspected += 1
            current_email = str(preview.get("email", "")).strip()
            current_subject = str(preview.get("subject", "")).strip()
            current_date = str(preview.get("date", "")).strip()
            if callable(progress_cb):
                progress_cb(
                    phase="searching",
                    progress_current=min(len(targets), wanted),
                    progress_total=wanted,
                    matched=len(targets),
                    inspected=inspected,
                    current_email=current_email,
                    current_subject=current_subject,
                    current_date=current_date,
                    message=f"Analisando mensagens sem label: {inspected} verificadas.",
                    detail=f"Encontradas {len(targets)} dentro dos filtros. Página {pages} de até {page_limit}.",
                )
            if not _preview_matches_nf_range(current_subject, start_nf, end_nf):
                continue
            targets.append(
                {
                    "id": msg_id,
                    "threadId": str(item.get("threadId", "")).strip(),
                    "labelIds": list(item.get("labelIds", []) or []),
                    "snippet": str(item.get("snippet", "") or ""),
                }
            )
            if callable(progress_cb):
                progress_cb(
                    phase="searching",
                    progress_current=min(len(targets), wanted),
                    progress_total=wanted,
                    matched=len(targets),
                    inspected=inspected,
                    current_email=current_email,
                    current_subject=current_subject,
                    current_date=current_date,
                    message=f"Mensagens encontradas: {len(targets)} de {wanted}.",
                    detail=f"Critérios: {criteria_desc}. Página {pages} de até {page_limit}.",
                )
            if len(targets) >= wanted:
                break
        if len(targets) >= wanted or not next_page_token:
            break
        page_token = next_page_token
    return {
        "ok": True,
        "matched": len(targets),
        "inspected": inspected,
        "pages": pages,
        "query": query,
        "criteria": criteria_desc,
        "targets": targets,
    }


def _start_recover_missing_background(
    max_messages: int,
    nf_start: str = "",
    nf_end: str = "",
    date_from: str = "",
    date_to: str = "",
) -> tuple[bool, dict]:
    criteria_desc = _describe_recovery_filters(nf_start=nf_start, nf_end=nf_end, date_from=date_from, date_to=date_to)
    if not criteria_desc:
        return False, {"message": "Informe uma faixa de NF e/ou um período para recuperar."}
    snap = _manual_action_snapshot()
    if bool(snap.get("active")):
        if not snap.get("message"):
            label = str(snap.get("label") or "Acao manual").strip() or "Acao manual"
            snap["message"] = f"{label} ja esta em andamento."
        return False, snap
    started, snap = _manual_action_begin(
        "recover_missing",
        "Recuperação de faltantes",
        "Recuperação iniciada.",
        detail=f"Buscando até {int(max_messages)} mensagens sem label do Botana em {criteria_desc}.",
        progress_total=int(max_messages),
        requested_limit=int(max_messages),
        matched=0,
        inspected=0,
    )
    if not started:
        return False, snap

    def _progress(**kwargs):
        _manual_action_update(**kwargs)

    def _worker():
        resume_loop = bool(running)
        try:
            if resume_loop:
                _manual_action_update(
                    message="Interrompendo ciclo automático para recuperar faltantes.",
                    detail="O loop automático será retomado depois da recuperação.",
                )
                parar_verificacao(wait=True, timeout=120.0)
                if (_LOOP_THREAD and _LOOP_THREAD.is_alive()) or not _wait_for_processing_idle(timeout=5.0):
                    _manual_action_finish(
                        False,
                        "Não foi possível iniciar a recuperação.",
                        detail="O ciclo automático não liberou a leitura a tempo.",
                        requested_limit=int(max_messages),
                    )
                    return
            stop_event.clear()
            _manual_action_update(
                message="Recuperação em andamento.",
                detail=f"Varrendo Gmail em busca de mensagens sem label do Botana em {criteria_desc}.",
                phase="searching",
                matched=0,
                inspected=0,
            )
            result = _find_missing_messages(
                max_messages=max_messages,
                nf_start=nf_start,
                nf_end=nf_end,
                date_from=date_from,
                date_to=date_to,
                progress_cb=_progress,
            )
            targets = list(result.get("targets") or [])
            matched = int(result.get("matched", 0) or 0)
            inspected = int(result.get("inspected", 0) or 0)
            if targets:
                _manual_action_update(
                    phase="processing",
                    progress_current=0,
                    progress_total=len(targets),
                    matched=matched,
                    inspected=inspected,
                    current_email="",
                    current_subject="",
                    current_date="",
                    message="Mensagens encontradas. Iniciando leitura.",
                    detail=f"O Botana vai reler {len(targets)} mensagens sem label do Botana em {criteria_desc}.",
                )
                ok, msg = executar_um_ciclo(
                    max_messages_override=len(targets),
                    messages_override=targets,
                )
                proc = _process_snapshot().get("last", {})
                detail = (
                    f"Encontradas: {matched} | "
                    f"Analisadas: {inspected} | "
                    f"E-mails: {int(proc.get('messages', 0) or 0)} | "
                    f"Anexos: {int(proc.get('attachments', 0) or 0)} | "
                    f"XML: {int(proc.get('xmls', 0) or 0)} | "
                    f"Lançamentos: {int(proc.get('launched', 0) or 0)} | "
                    f"Duplicadas: {int(proc.get('duplicates', 0) or 0)}"
                )
                if resume_loop:
                    restarted = iniciar_verificacao()
                    detail = f"{detail} | Loop automático {'retomado' if restarted else 'não retomado'}."
                _manual_action_finish(
                    ok,
                    msg,
                    detail=detail,
                    progress_current=int(proc.get("messages", 0) or 0),
                    progress_total=len(targets),
                    matched=matched,
                    inspected=inspected,
                    requested_limit=int(max_messages),
                )
                return
            detail = f"Nenhuma mensagem sem label do Botana combinou com {criteria_desc}. Analisadas: {inspected}."
            if resume_loop:
                restarted = iniciar_verificacao()
                detail = f"{detail} | Loop automático {'retomado' if restarted else 'não retomado'}."
            _manual_action_finish(
                True,
                "Nenhuma mensagem pendente foi encontrada para os filtros informados.",
                detail=detail,
                progress_current=matched,
                progress_total=int(max_messages),
                matched=matched,
                inspected=inspected,
                requested_limit=int(max_messages),
            )
        except Exception as exc:
            logger.exception("Falha na recuperação de faltantes em background: %s", exc)
            if resume_loop:
                try:
                    iniciar_verificacao()
                except Exception:
                    logger.exception("Falha ao retomar loop automatico apos erro na recuperação.")
            _manual_action_finish(False, "Erro na recuperação de faltantes.", detail=str(exc), requested_limit=int(max_messages))

    threading.Thread(target=_worker, daemon=True, name="botana-recover-missing").start()
    return True, snap


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
#tabHist,#tabDiag,#tabWatch{padding:0 10px 10px}
.card{background:rgba(255,248,240,.92);border:1px solid #e7c8a8;border-radius:13px;padding:10px;box-shadow:0 8px 20px rgba(21,11,6,.06)}
h3{margin:0 0 8px;color:var(--b);font-size:.98rem}
label{display:block;font-weight:600;color:#5c341c;font-size:.9rem}
.muted{color:#6c4a35;font-size:.84rem}
.btns{display:flex;gap:8px;flex-wrap:wrap}
button{padding:9px 12px;border:0;border-radius:9px;background:linear-gradient(90deg,var(--o),var(--o2));color:#2b1408;font-weight:700;cursor:pointer;font-size:.9rem}
button.sec{background:linear-gradient(90deg,#7a4d30,#5b341f);color:#fff9f3}
button.warn{background:linear-gradient(90deg,#bc2d2d,#8f2020);color:#fff}
button:disabled{opacity:.62;cursor:not-allowed;filter:saturate(.8)}
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
.cfg-grid{display:grid;grid-template-columns:minmax(560px,1fr) minmax(120px,145px) minmax(230px,280px);gap:10px;align-items:start}
.cfg-grid > .card{height:auto;display:flex;flex-direction:column;align-self:start}
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
.reproc-grid{display:grid;grid-template-columns:minmax(180px,240px);gap:8px;align-items:end;justify-content:center;justify-items:center}
.reproc-grid > div{display:flex;flex-direction:column;align-items:center}
.reproc-grid > div label{text-align:center}
.reproc-grid > div input,.reproc-grid > div select{width:min(220px,100%);text-align:center}
.reproc-card .muted{text-align:center}
.recover-card h3{text-align:center}
.recover-grid{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:8px;align-items:end}
.recover-grid > div{display:flex;flex-direction:column;justify-content:center;align-items:center}
.recover-grid > div label{width:100%;text-align:center}
.recover-grid > div input{width:min(180px,100%);text-align:center}
.recover-note{margin-top:8px;text-align:center}
.cb{margin-top:8px;display:inline-flex;align-items:center;gap:8px}
.action-box{margin-top:10px;border:1px solid #d8b391;border-radius:10px;background:#fffaf5;padding:9px;display:grid;gap:6px}
.action-head{display:flex;justify-content:space-between;align-items:center;gap:8px}
.action-title{font-size:.84rem;font-weight:700;color:#5b321c}
.action-detail{font-size:.78rem;color:#6c4a35}
.action-progress{font-size:.78rem;color:#6b4128}
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
.audit-filters{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:8px;align-items:end}
.audit-filters > div{display:flex;flex-direction:column;justify-content:center;align-items:center}
.audit-filters > div label{width:100%;text-align:center}
.audit-filters > div input,.audit-filters > div select{width:100%;text-align:center}
.audit-title{text-align:center}
.audit-toolbar{margin-top:8px;display:flex;flex-direction:column;justify-content:center;align-items:center;gap:6px}
.audit-note{max-width:900px;text-align:center}
.audit-state{min-height:20px;text-align:center;font-size:.83rem;color:#6b4126}
.audit-state.loading{color:#a25b18;font-weight:700}
.audit-summary{margin-top:10px;display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:10px}
.audit-summary .k{border:1px solid #e2b58d;border-radius:10px;background:linear-gradient(180deg,#fff7ef,#fff1e3);padding:10px;text-align:center}
.audit-summary .n{font-size:1.3rem;font-weight:800;color:#7a3d11}
.audit-summary .t{font-size:.8rem;color:#6b4126}
.audit-table{width:100%;min-width:1080px;border-collapse:collapse;font-size:.8rem;table-layout:fixed;border:1px solid #ddb38d}
.audit-table th,.audit-table td{border:1px solid #e7c4a5;padding:7px 8px;text-align:center;vertical-align:middle;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.audit-table th{position:sticky;top:0;background:#fff1e3;color:#5c341c;z-index:1;border-bottom:2px solid #cf9c73}
.audit-col-status{width:92px}
.audit-col-nf{width:86px}
.audit-col-cliente{width:260px}
.audit-col-sm{width:78px}
.audit-col-date{width:126px}
.audit-col-aba{width:136px}
.audit-table tbody tr:nth-child(even){background:rgba(255,244,232,.92)}
.audit-table tbody tr:hover{background:rgba(238,155,47,.08)}
.audit-status{display:inline-flex;align-items:center;justify-content:center;padding:3px 8px;border-radius:999px;font-size:.72rem;font-weight:700;border:1px solid transparent}
.audit-status-btn{cursor:pointer;transition:transform .15s ease, box-shadow .15s ease}
.audit-status-btn:hover{transform:translateY(-1px);box-shadow:0 4px 10px rgba(92,52,28,.12)}
.audit-status.ok{background:#e9f8ec;color:#1c6a32;border-color:#87c69a}
.audit-status.aviso{background:#fff3dd;color:#8b5a00;border-color:#e7bf6e}
.audit-status.erro{background:#fde7ea;color:#a61d2d;border-color:#dc3545}
.audit-row-aviso{background:rgba(255,193,7,.12)!important}
.audit-row-aviso:hover{background:rgba(255,193,7,.2)!important}
.audit-row-erro{background:rgba(220,53,69,.14)!important}
.audit-row-erro:hover{background:rgba(220,53,69,.22)!important}
.watch-title{text-align:center}
.watch-filters{display:grid;grid-template-columns:1fr;gap:8px;align-items:center;justify-items:center}
.watch-filters > div{display:flex;flex-direction:column;justify-content:center;align-items:center}
.watch-filters > div label{width:100%;text-align:center}
.watch-filters > div input{width:min(88px,100%);text-align:center}
.watch-actions{display:flex;justify-content:center;align-items:center;gap:8px;flex-wrap:wrap;margin-top:14px}
.watch-toolbar{margin-top:8px;display:flex;flex-direction:column;justify-content:center;align-items:center;gap:6px}
.watch-note{max-width:900px;text-align:center}
.watch-state{min-height:20px;text-align:center;font-size:.83rem;color:#6b4126}
.watch-state.loading{color:#a25b18;font-weight:700}
.watch-summary{margin-top:10px;display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px}
.watch-summary .k{border:1px solid #e2b58d;border-radius:10px;background:linear-gradient(180deg,#fff7ef,#fff1e3);padding:10px;text-align:center}
.watch-summary .n{font-size:1.3rem;font-weight:800;color:#7a3d11}
.watch-summary .t{font-size:.8rem;color:#6b4126}
.watch-table{width:100%;min-width:980px;border-collapse:collapse;font-size:.8rem;table-layout:fixed;border:1px solid #ddb38d}
.watch-table th,.watch-table td{border:1px solid #e7c4a5;padding:7px 8px;text-align:center;vertical-align:middle;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.watch-table th{position:sticky;top:0;background:#fff1e3;color:#5c341c;z-index:1;border-bottom:2px solid #cf9c73}
.watch-col-type{width:88px}
.watch-col-status{width:126px}
.watch-col-date{width:110px}
.watch-col-days{width:126px}
.watch-col-client{width:260px}
.watch-col-nf{width:86px}
.watch-col-value{width:118px}
.watch-col-aba{width:130px}
.watch-table tbody tr:nth-child(even){background:rgba(255,244,232,.92)}
.watch-table tbody tr:hover{background:rgba(238,155,47,.08)}
.watch-row-aviso{background:rgba(255,193,7,.12)!important}
.watch-row-aviso:hover{background:rgba(255,193,7,.2)!important}
.watch-row-erro{background:rgba(220,53,69,.14)!important}
.watch-row-erro:hover{background:rgba(220,53,69,.22)!important}
.watch-badge{display:inline-flex;align-items:center;justify-content:center;padding:3px 8px;border-radius:999px;font-size:.72rem;font-weight:700;border:1px solid transparent}
.watch-badge.aviso{background:#fff3dd;color:#8b5a00;border-color:#e7bf6e}
.watch-badge.erro{background:#fde7ea;color:#a61d2d;border-color:#dc3545}
.watch-pop{position:fixed;inset:0;z-index:99998;display:none;align-items:center;justify-content:center;background:rgba(22,10,5,.68);backdrop-filter:blur(3px);padding:20px}
.watch-pop.show{display:flex}
.watch-pop-box{width:min(920px,94vw);max-height:min(82vh,760px);overflow:hidden;display:grid;grid-template-rows:auto auto auto 1fr;background:linear-gradient(180deg,#fff9f3,#fff2e5);border:1px solid #efc9a3;border-radius:16px;box-shadow:0 18px 42px rgba(20,10,4,.22)}
.watch-pop-head{position:relative;display:flex;justify-content:center;align-items:center;gap:12px;padding:14px 16px 10px;border-bottom:1px solid #efd6bf}
.watch-pop-head h4{margin:0;color:#5f341a;text-align:center}
.watch-pop-close{position:absolute;right:16px;top:10px;border:1px solid #d7ab82;background:#fff7ef;color:#6d3b1a;border-radius:10px;padding:6px 10px;cursor:pointer;font-weight:700}
.watch-pop-close:hover{background:#ffeddc}
.watch-pop-search{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:8px;padding:14px 16px 10px;align-items:end}
.watch-pop-field{position:relative;display:flex;flex-direction:column;align-items:center;width:100%}
.watch-pop-search label{text-align:center;width:100%}
.watch-pop-search input{width:min(520px,100%);margin-top:4px;text-align:center}
.watch-pop-suggest{display:none;width:min(520px,100%);margin-top:6px;max-height:210px;overflow:auto;border:1px solid #e7c4a5;border-radius:12px;background:#fffaf6;box-shadow:0 10px 24px rgba(20,10,4,.12)}
.watch-pop-suggest.show{display:block}
.watch-pop-suggest-item{display:block;width:100%;border:0;background:transparent;padding:9px 10px;text-align:center;color:#5f341a;cursor:pointer;font-size:.84rem}
.watch-pop-suggest-item + .watch-pop-suggest-item{border-top:1px solid #f0dac6}
.watch-pop-suggest-item:hover,.watch-pop-suggest-item.active{background:#fff0e1}
.watch-pop-state{padding:0 16px 10px;min-height:22px;text-align:center;color:#6b4126;font-size:.84rem}
.watch-pop-state.loading{color:#a25b18;font-weight:700}
.watch-pop-results{padding:0 16px 16px;overflow:auto}
.watch-pop-empty{border:1px dashed #e0b68c;border-radius:12px;background:#fffaf6;padding:16px;text-align:center;color:#6b4126}
.watch-pop-table{width:100%;min-width:760px;border-collapse:collapse;font-size:.8rem;table-layout:fixed;border:1px solid #ddb38d;background:#fffdfb}
.watch-pop-table th,.watch-pop-table td{border:1px solid #e7c4a5;padding:7px 8px;text-align:center;vertical-align:middle;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.watch-pop-table th{position:sticky;top:0;background:#fff1e3;color:#5c341c;z-index:1;border-bottom:2px solid #cf9c73}
.watch-pop-table tbody tr:nth-child(even){background:rgba(255,244,232,.92)}
.watch-pop-table tbody tr:hover{background:rgba(238,155,47,.08)}
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
@media(max-width:900px){.lists{grid-template-columns:1fr}.cfg-grid{grid-template-columns:1fr}.cfg-fields{grid-template-columns:1fr 1fr}.reproc-grid{grid-template-columns:1fr}.recover-grid{grid-template-columns:1fr 1fr 1fr}}
@media(max-width:1020px){.hist-filters{grid-template-columns:1fr 1fr 1fr}.audit-filters{grid-template-columns:1fr 1fr}.audit-summary{grid-template-columns:1fr 1fr 1fr}.watch-summary{grid-template-columns:1fr 1fr}.recover-grid{grid-template-columns:1fr 1fr 1fr}}
@media(max-width:640px){.top-right{flex-direction:column;align-items:flex-end}.hist-filters{grid-template-columns:1fr}.audit-filters{grid-template-columns:1fr}.audit-summary{grid-template-columns:1fr 1fr}.watch-filters{grid-template-columns:1fr}.watch-summary{grid-template-columns:1fr 1fr}.recover-grid{grid-template-columns:1fr}.watch-pop-search{grid-template-columns:1fr}.watch-pop-close{position:static}}
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
    <button id="tabBtnAudit" type="button" class="tab-btn" onclick="switchTab('audit')">Conferência</button>
    <button id="tabBtnWatch" type="button" class="tab-btn" onclick="switchTab('watch')">Prazos</button>
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
            <label>Limite de mensagens</label>
            <input id="limit" type="number" value="100" min="1" max="1000"/>
          </div>
        </div>
        <div class="muted" style="margin-top:6px">Usa as mensagens mais recentes com label do Botana, remarca para reprocessamento e já executa a leitura em seguida.</div>
        <label class="cb"><input id="unread" type="checkbox" checked/>Marcar como não lido</label>
        <div class="btns">
          <button id="reprocessBtn" onclick="reprocess()">Reprocessar agora</button>
        </div>
        <div class="action-box">
          <div class="action-head">
            <span class="action-title" id="manualActionTitle">Ações manuais</span>
            <span id="manualActionBadge" class="status-pill off"><span>•</span><span>Aguardando</span></span>
          </div>
          <div id="manualActionMsg" class="proc-line">Nenhuma ação manual em andamento.</div>
          <div id="manualActionDetail" class="action-detail">Use o botão acima para remarcar as mensagens e executar a leitura no mesmo fluxo.</div>
          <div class="proc-progress">
            <div class="proc-track"><div id="manualActionBar" class="proc-fill"></div></div>
            <div id="manualActionProgress" class="action-progress">Progresso: -</div>
          </div>
        </div>
      </article>
    </section>

    <section class="card recover-card" style="margin-top:10px">
      <h3>Recuperar e-mails sem leitura</h3>
      <div class="recover-grid">
        <div>
          <label>Data inicial</label>
          <input id="recoverDateFrom" type="date"/>
        </div>
        <div>
          <label>Data final</label>
          <input id="recoverDateTo" type="date"/>
        </div>
        <div>
          <label>NF inicial</label>
          <input id="recoverNfStart" type="text" placeholder="20247"/>
        </div>
        <div>
          <label>NF final</label>
          <input id="recoverNfEnd" type="text" placeholder="20481"/>
        </div>
        <div>
          <label>Limite de mensagens</label>
          <input id="recoverLimit" type="number" value="200" min="1" max="1000"/>
        </div>
        <div style="display:flex;align-items:end;justify-content:center">
          <button id="recoverBtn" onclick="recoverMissing()">Recuperar faltantes</button>
        </div>
      </div>
      <div class="muted recover-note">Busca apenas mensagens enviadas com XML que ainda não têm label do Botana. Use período e/ou faixa de NF; o progresso aparece em Ações manuais.</div>
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
        <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">
          <button type="button" class="sec hist-reset-btn" onclick="exportHistoryCsv()">Exportar CSV</button>
          <button type="button" class="sec hist-reset-btn" onclick="resetHistColumnWidths()">Resetar larguras</button>
        </div>
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

  <section id="tabAudit" class="tab-panel hidden">
    <section class="card" style="margin-top:10px">
      <h3 class="audit-title">Conferência de parcelas lançadas</h3>
      <div class="audit-filters">
        <div>
          <label>Modo</label>
          <select id="aMode" onchange="toggleAuditFilters()">
            <option value="mes">Mês do lançamento</option>
            <option value="nfs">Faixa de NF</option>
            <option value="todos">Tudo</option>
          </select>
        </div>
        <div>
          <label>Mês</label>
          <input id="aMonth" type="month"/>
        </div>
        <div>
          <label>NF inicial</label>
          <input id="aNfStart" type="text" placeholder="49001"/>
        </div>
        <div>
          <label>NF final</label>
          <input id="aNfEnd" type="text" placeholder="49100"/>
        </div>
        <div style="display:flex;align-items:end"><button id="auditRunBtn" onclick="loadParcelAudit()">Conferir parcelas</button></div>
      </div>
      <div class="audit-toolbar">
        <div class="muted audit-note">A conferência lê diretamente as planilhas, seleciona as NFs pelo filtro escolhido e compara o total esperado da NF com as parcelas registradas nas abas.</div>
        <div id="auditStatus" class="audit-state">Pronto para conferir.</div>
      </div>
      <div class="audit-summary">
        <div class="k"><div id="auditK1" class="n">0</div><div class="t">NFs verificadas</div></div>
        <div class="k"><div id="auditK2" class="n">0</div><div class="t">Com divergência</div></div>
        <div class="k"><div id="auditK3" class="n">0</div><div class="t">Parcelas esperadas</div></div>
        <div class="k"><div id="auditK4" class="n">0</div><div class="t">Parcelas lançadas</div></div>
        <div class="k"><div id="auditK5" class="n">0</div><div class="t">Duplicadas</div></div>
      </div>
      <div class="table-wrap" style="margin-top:10px">
        <table class="audit-table">
          <colgroup>
            <col class="audit-col-status"/>
            <col class="audit-col-nf"/>
            <col class="audit-col-cliente"/>
            <col class="audit-col-sm"/>
            <col class="audit-col-sm"/>
            <col class="audit-col-sm"/>
            <col class="audit-col-sm"/>
            <col class="audit-col-date"/>
            <col class="audit-col-aba"/>
          </colgroup>
          <thead>
            <tr>
              <th>Status</th>
              <th>NF</th>
              <th>Cliente</th>
              <th>Esperadas</th>
              <th>Lançadas</th>
              <th>Faltando</th>
              <th>Duplicadas</th>
              <th>Últ. venc.</th>
              <th>Aba</th>
            </tr>
          </thead>
          <tbody id="aBody"><tr><td colspan="9">Sem dados</td></tr></tbody>
        </table>
      </div>
    </section>
  </section>

  <section id="tabWatch" class="tab-panel hidden">
    <section class="card" style="margin-top:10px">
      <h3 class="watch-title">Boletos e depósitos próximos do limite</h3>
      <div class="watch-filters">
        <div>
          <label>Boletos em até</label>
          <input id="wBoletoDays" type="number" min="1" max="7" value="7"/>
        </div>
        <div>
          <label>Depósitos há pelo menos</label>
          <input id="wDepositoDays" type="number" min="1" max="7" value="7"/>
        </div>
        <div style="display:flex;align-items:end;justify-content:center"><button id="watchRunBtn" onclick="loadDueWatch()">Atualizar relação</button></div>
      </div>
      <div class="watch-actions">
        <button type="button" class="sec" onclick="openWatchSearchModal()">Buscar boletos em aberto</button>
      </div>
      <div class="watch-toolbar">
        <div class="muted watch-note">A relação lê diretamente as planilhas e lista apenas títulos com `Status` vazio ou `A Receber`. Boletos futuros ficam em amarelo; itens que vencem hoje ou já passaram ficam em vermelho.</div>
        <div id="watchStatus" class="watch-state">Pronto para consultar.</div>
      </div>
      <div class="watch-summary">
        <div class="k"><div id="watchK1" class="n">0</div><div class="t">Total na relação</div></div>
        <div class="k"><div id="watchK2" class="n">0</div><div class="t">Boletos a vencer</div></div>
        <div class="k"><div id="watchK3" class="n">0</div><div class="t">Boletos no limite</div></div>
        <div class="k"><div id="watchK4" class="n">0</div><div class="t">Depósitos atrasados</div></div>
      </div>
      <div class="table-wrap" style="margin-top:10px">
        <table class="watch-table">
          <colgroup>
            <col class="watch-col-type"/>
            <col class="watch-col-status"/>
            <col class="watch-col-date"/>
            <col class="watch-col-days"/>
            <col class="watch-col-client"/>
            <col class="watch-col-nf"/>
            <col class="watch-col-value"/>
            <col class="watch-col-aba"/>
          </colgroup>
          <thead>
            <tr>
              <th>Tipo</th>
              <th>Situação</th>
              <th>Vencimento</th>
              <th>Dias úteis</th>
              <th>Cliente</th>
              <th>NF</th>
              <th>Valor</th>
              <th>Aba</th>
            </tr>
          </thead>
          <tbody id="wBody"><tr><td colspan="8">Sem dados</td></tr></tbody>
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
<div id="watchSearchModal" class="watch-pop" onclick="closeWatchSearchModal(event)">
  <div class="watch-pop-box" onclick="event.stopPropagation()">
    <div class="watch-pop-head">
      <h4>Buscar boletos em aberto</h4>
      <button type="button" class="watch-pop-close" onclick="closeWatchSearchModal()">Fechar</button>
    </div>
    <div class="watch-pop-search">
      <div class="watch-pop-field">
        <label>Nome do cliente</label>
        <input id="watchSearchInput" type="text" placeholder="Digite o nome completo ou parcial"/>
        <div id="watchSearchSuggestions" class="watch-pop-suggest"></div>
      </div>
      <button id="watchSearchBtn" type="button" onclick="searchOpenBoletos()">Buscar</button>
    </div>
    <div id="watchSearchState" class="watch-pop-state">Digite um nome para consultar boletos em aberto.</div>
    <div id="watchSearchResults" class="watch-pop-results">
      <div class="watch-pop-empty">Nenhuma busca executada ainda.</div>
    </div>
  </div>
</div>
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
  if(h==='hist'||h==='diag'||h==='main'||h==='audit'||h==='watch') return h;
  const q=new URLSearchParams(window.location.search||'');
  const t=String(q.get('tab')||'').trim().toLowerCase();
  if(t==='hist'||t==='diag'||t==='main'||t==='audit'||t==='watch') return t;
  return 'main';
}
function switchTab(tab){
  const next=(tab==='hist'||tab==='diag'||tab==='audit'||tab==='watch')?tab:'main';
  const changed=_activeTab!==next;
  const m=next==='main';
  const h=next==='hist';
  const a=next==='audit';
  const w=next==='watch';
  const d=next==='diag';
  document.getElementById('tabMain').classList.toggle('hidden',!m);
  document.getElementById('tabHist').classList.toggle('hidden',!h);
  document.getElementById('tabAudit').classList.toggle('hidden',!a);
  document.getElementById('tabWatch').classList.toggle('hidden',!w);
  document.getElementById('tabDiag').classList.toggle('hidden',!d);
  document.getElementById('tabBtnMain').classList.toggle('active',m);
  document.getElementById('tabBtnHist').classList.toggle('active',h);
  document.getElementById('tabBtnAudit').classList.toggle('active',a);
  document.getElementById('tabBtnWatch').classList.toggle('active',w);
  document.getElementById('tabBtnDiag').classList.toggle('active',d);
  _activeTab=next;
  if(next==='hist'){
    loadHistory().catch(()=>{});
  }
  if(next==='audit'){
    loadParcelAudit(false).catch(()=>{});
  }
  if(next==='watch'){
    loadDueWatch(false).catch(()=>{});
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
  const d=Number(it.duplicates||0);
  return `E-mails ${m} | Anexos ${a} | XML ${x} | Lançamentos ${l} | Duplicadas ${d}`;
}
function _reprocessView(action){
  const a=action||{};
  const requested=Math.max(0, Number(a.requested_limit||0));
  const total=Math.max(0, Number(a.progress_total||0));
  const current=Math.max(0, Number(a.progress_current||0));
  const visibleTotal=total>0?total:requested;
  const denom=Math.max(1, visibleTotal||1);
  const perc=visibleTotal>0?Math.max(0,Math.min(100,Math.round((Math.min(current,visibleTotal)/denom)*100))):0;
  const currentEmail=String(a.current_email||'').trim();
  const currentDate=String(a.current_date||'').trim();
  const currentSubject=String(a.current_subject||'').trim();
  return {requested,total,current,visibleTotal,perc,currentEmail,currentDate,currentSubject};
}
function _manualRequestedLimit(action,fallback){
  const requested=Math.max(0, Number(((action||{}).requested_limit)||0));
  if(requested>0)return requested;
  return Math.max(1, Number(fallback||100));
}
function updProcessing(proc,maxMessages,action){
  const runEl=document.getElementById('procRun');
  const nowEl=document.getElementById('procNow');
  const lastEl=document.getElementById('procLast');
  const barFill=document.getElementById('procBarFill');
  const barLabel=document.getElementById('procBarLabel');
  if(!runEl||!nowEl||!lastEl||!barFill||!barLabel) return;
  const p=proc||{};
  const a=action||{};
  const active=!!a.active;
  const kind=String(a.kind||'').trim();
  const phase=String(a.phase||'').trim();
  if(active&&kind==='reprocess'&&phase!=='processing'){
    const view=_reprocessView(a);
    const detailParts=[];
    if(view.currentEmail)detailParts.push(`E-mail atual ${view.currentEmail}`);
    if(view.currentDate)detailParts.push(`Data ${view.currentDate}`);
    if(view.currentSubject)detailParts.push(`Assunto ${view.currentSubject}`);
    runEl.textContent='Loop: reprocessamento manual em andamento';
    nowEl.textContent=detailParts.length?`Mensagem atual: ${detailParts.join(' | ')}`:'Mensagem atual: buscando e-mails para reprocessar';
    const last=p.last||{};
    const lastEnd=last.finished_at?_fmtDateTime(last.finished_at):'-';
    let statusTxt='-';
    if(last.ok===true) statusTxt='OK';
    else if(last.ok===false) statusTxt='Erro';
    lastEl.textContent=`Último ciclo automático: ${statusTxt} em ${lastEnd} | ${_fmtCycleShort(last)}`;
    barFill.style.width=String(view.perc)+'%';
    barLabel.textContent=`Reprocessamento: ${view.current}/${view.visibleTotal||'-'} (${view.perc}%) | Limite pedido ${view.requested||'-'} | Atualizadas ${Number(a.changed||0)} | Falhas ${Number(a.failed||0)}`;
    return;
  }
  if(active&&kind==='reprocess'&&phase==='processing'){
    const cur=p.current||{};
    const currentLimit=Math.max(1, Number(a.progress_total||_manualRequestedLimit(a,maxMessages)));
    const curStart=cur.started_at?_fmtDateTime(cur.started_at):'-';
    runEl.textContent='Loop: executando leitura do reprocessamento';
    nowEl.textContent=`Ciclo atual: inicio ${curStart} | ${_fmtCycleShort(cur)}`;
    const last=p.last||{};
    const lastEnd=last.finished_at?_fmtDateTime(last.finished_at):'-';
    let statusTxt='-';
    if(last.ok===true) statusTxt='OK';
    else if(last.ok===false) statusTxt='Erro';
    lastEl.textContent=`Ultimo ciclo automatico: ${statusTxt} em ${lastEnd} | ${_fmtCycleShort(last)}`;
    let curV=Number(cur.messages||0);
    if(!Number.isFinite(curV)||curV<0)curV=0;
    if(curV>currentLimit)curV=currentLimit;
    const perc=Math.max(0,Math.min(100,Math.round((curV/Math.max(1,currentLimit))*100)));
    barFill.style.width=String(perc)+'%';
    barLabel.textContent=`Relancamento: ${curV}/${currentLimit} (${perc}%) | XML ${Number(cur.xmls||0)} | Lancamentos ${Number(cur.launched||0)} | Duplicadas ${Number(cur.duplicates||0)} | Labels atualizadas ${Number(a.changed||0)} | Falhas ${Number(a.failed||0)}`;
    return;
  }
  if(active&&kind==='recover_missing'&&phase!=='processing'){
    const wanted=Math.max(1, Number(a.requested_limit||a.progress_total||maxMessages||100));
    const matched=Math.max(0, Number(a.matched||0));
    const inspected=Math.max(0, Number(a.inspected||0));
    const perc=Math.max(0,Math.min(100,Math.round((Math.min(matched,wanted)/Math.max(1,wanted))*100)));
    const detailParts=[];
    if(String(a.current_email||'').trim())detailParts.push(`E-mail atual ${String(a.current_email||'').trim()}`);
    if(String(a.current_date||'').trim())detailParts.push(`Data ${String(a.current_date||'').trim()}`);
    if(String(a.current_subject||'').trim())detailParts.push(`Assunto ${String(a.current_subject||'').trim()}`);
    runEl.textContent='Loop: recuperação manual em andamento';
    nowEl.textContent=detailParts.length?`Mensagem atual: ${detailParts.join(' | ')}`:'Mensagem atual: varrendo e-mails sem label do Botana';
    const last=p.last||{};
    const lastEnd=last.finished_at?_fmtDateTime(last.finished_at):'-';
    let statusTxt='-';
    if(last.ok===true) statusTxt='OK';
    else if(last.ok===false) statusTxt='Erro';
    lastEl.textContent=`Último ciclo automático: ${statusTxt} em ${lastEnd} | ${_fmtCycleShort(last)}`;
    barFill.style.width=String(perc)+'%';
    barLabel.textContent=`Recuperação: ${matched}/${wanted} encontradas | ${inspected} analisadas`;
    return;
  }
  if(active&&kind==='recover_missing'&&phase==='processing'){
    const cur=p.current||{};
    const currentLimit=Math.max(1, Number(a.progress_total||_manualRequestedLimit(a,maxMessages)));
    const matched=Math.max(0, Number(a.matched||0));
    const inspected=Math.max(0, Number(a.inspected||0));
    const curStart=cur.started_at?_fmtDateTime(cur.started_at):'-';
    runEl.textContent='Loop: executando leitura da recuperação';
    nowEl.textContent=`Ciclo atual: inicio ${curStart} | ${_fmtCycleShort(cur)}`;
    const last=p.last||{};
    const lastEnd=last.finished_at?_fmtDateTime(last.finished_at):'-';
    let statusTxt='-';
    if(last.ok===true) statusTxt='OK';
    else if(last.ok===false) statusTxt='Erro';
    lastEl.textContent=`Último ciclo automático: ${statusTxt} em ${lastEnd} | ${_fmtCycleShort(last)}`;
    let curV=Number(cur.messages||0);
    if(!Number.isFinite(curV)||curV<0)curV=0;
    if(curV>currentLimit)curV=currentLimit;
    const perc=Math.max(0,Math.min(100,Math.round((curV/Math.max(1,currentLimit))*100)));
    barFill.style.width=String(perc)+'%';
    barLabel.textContent=`Recuperação: ${curV}/${currentLimit} (${perc}%) | Encontradas ${matched} | Analisadas ${inspected} | XML ${Number(cur.xmls||0)} | Lançamentos ${Number(cur.launched||0)} | Duplicadas ${Number(cur.duplicates||0)}`;
    return;
  }
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
function _setManualBadge(kind,label){
  const el=document.getElementById('manualActionBadge');
  if(!el)return;
  el.className='status-pill '+kind;
  el.innerHTML='<span>•</span><span>'+label+'</span>';
}
function updManualAction(action,processing,maxMessages){
  const a=action||{};
  const msgEl=document.getElementById('manualActionMsg');
  const detailEl=document.getElementById('manualActionDetail');
  const titleEl=document.getElementById('manualActionTitle');
  const barEl=document.getElementById('manualActionBar');
  const progressEl=document.getElementById('manualActionProgress');
  const repBtn=document.getElementById('reprocessBtn');
  const recoverBtn=document.getElementById('recoverBtn');
  if(!msgEl||!detailEl||!titleEl||!barEl||!progressEl)return;
  const active=!!a.active;
  const kind=String(a.kind||'').trim();
  const phase=String(a.phase||'').trim();
  const label=String(a.label||'Ações manuais').trim()||'Ações manuais';
  titleEl.textContent=label;
  if(active&&kind==='run_now'){
    const cur=((processing||{}).current)||{};
    const maxV=Math.max(1, Number(maxMessages||100));
    let curV=Number(cur.messages||0);
    if(!Number.isFinite(curV)||curV<0)curV=0;
    const perc=Math.max(8,Math.min(95,Math.round((Math.min(curV,maxV)/maxV)*100)));
    barEl.style.width=String(perc)+'%';
    progressEl.textContent=`Mensagens lidas: ${curV}/${maxV} | Anexos ${Number(cur.attachments||0)} | XML ${Number(cur.xmls||0)} | Lançamentos ${Number(cur.launched||0)} | Duplicadas ${Number(cur.duplicates||0)}`;
    msgEl.textContent=String(a.message||'Execução manual em andamento.');
    detailEl.textContent=String(a.detail||'Leitura de e-mails e lançamentos em andamento.');
    _setManualBadge('ok','Em andamento');
  }else if(active&&kind==='reprocess'){
    if(phase==='processing'){
      const cur=((processing||{}).current)||{};
      const maxV=Math.max(1, Number(a.progress_total||_manualRequestedLimit(a,maxMessages)));
      let curV=Number(cur.messages||0);
      if(!Number.isFinite(curV)||curV<0)curV=0;
      const perc=Math.max(8,Math.min(95,Math.round((Math.min(curV,maxV)/Math.max(1,maxV))*100)));
      barEl.style.width=String(perc)+'%';
      progressEl.textContent=`Mensagens lidas: ${curV}/${maxV} | Anexos ${Number(cur.attachments||0)} | XML ${Number(cur.xmls||0)} | Lançamentos ${Number(cur.launched||0)} | Duplicadas ${Number(cur.duplicates||0)} | Labels atualizadas ${Number(a.changed||0)} | Falhas ${Number(a.failed||0)}`;
      msgEl.textContent=String(a.message||'Reprocessamento em andamento.');
      detailEl.textContent=String(a.detail||'Leitura e relançamento em andamento.');
      _setManualBadge('ok','Lendo');
    }else{
      const view=_reprocessView(a);
      const currentParts=[];
      if(view.currentEmail)currentParts.push(`E-mail atual: ${view.currentEmail}`);
      if(view.currentDate)currentParts.push(`Data: ${view.currentDate}`);
      barEl.style.width=String(view.visibleTotal>0?Math.max(6,view.perc):18)+'%';
      progressEl.textContent=`Mensagens: ${view.current}/${view.visibleTotal||'-'} | Limite pedido ${view.requested||'-'} | Atualizadas ${Number(a.changed||0)} | Falhas ${Number(a.failed||0)}${currentParts.length?` | ${currentParts.join(' | ')}`:''}`;
      msgEl.textContent=String(a.message||'Reprocessamento em andamento.');
      detailEl.textContent=String(a.detail||'Atualizando a label do Botana para Reprocessado.');
      _setManualBadge('ok','Remarcando');
    }
  }else if(active&&kind==='recover_missing'){
    const matched=Math.max(0, Number(a.matched||0));
    const inspected=Math.max(0, Number(a.inspected||0));
    if(phase==='processing'){
      const cur=((processing||{}).current)||{};
      const maxV=Math.max(1, Number(a.progress_total||_manualRequestedLimit(a,maxMessages)));
      let curV=Number(cur.messages||0);
      if(!Number.isFinite(curV)||curV<0)curV=0;
      const perc=Math.max(8,Math.min(95,Math.round((Math.min(curV,maxV)/Math.max(1,maxV))*100)));
      barEl.style.width=String(perc)+'%';
      progressEl.textContent=`Mensagens lidas: ${curV}/${maxV} | Encontradas ${matched} | Analisadas ${inspected} | Anexos ${Number(cur.attachments||0)} | XML ${Number(cur.xmls||0)} | Lançamentos ${Number(cur.launched||0)} | Duplicadas ${Number(cur.duplicates||0)}`;
      msgEl.textContent=String(a.message||'Recuperação em andamento.');
      detailEl.textContent=String(a.detail||'Leitura e lançamento das mensagens encontradas em andamento.');
      _setManualBadge('ok','Lendo');
    }else{
      const wanted=Math.max(1, Number(a.requested_limit||a.progress_total||maxMessages||100));
      const perc=Math.max(8,Math.min(95,Math.round((Math.min(matched,wanted)/Math.max(1,wanted))*100)));
      const currentParts=[];
      if(String(a.current_email||'').trim())currentParts.push(`E-mail atual: ${String(a.current_email||'').trim()}`);
      if(String(a.current_date||'').trim())currentParts.push(`Data: ${String(a.current_date||'').trim()}`);
      barEl.style.width=String(perc)+'%';
      progressEl.textContent=`Encontradas ${matched}/${wanted} | Analisadas ${inspected}${currentParts.length?` | ${currentParts.join(' | ')}`:''}`;
      msgEl.textContent=String(a.message||'Recuperação em andamento.');
      detailEl.textContent=String(a.detail||'Varrendo e-mails sem label do Botana.');
      _setManualBadge('ok','Varrendo');
    }
  }else{
    const status=String(a.status||'idle');
    const finished=String(a.finished_at||'').trim();
    const finishedText=finished?`Última atualização: ${_fmtDateTime(finished)}`:'Use o botão acima para reprocessar as mensagens e executar a leitura no mesmo fluxo.';
    barEl.style.width=status==='success'&&Number(a.progress_total||0)>0?'100%':'0%';
    progressEl.textContent=status==='success'||status==='error'
      ? `Progresso final: ${Number(a.progress_current||0)}/${Number(a.progress_total||0)||'-'} | Limite pedido ${Number(a.requested_limit||0)||'-'} | Atualizadas ${Number(a.changed||0)} | Falhas ${Number(a.failed||0)}`
      : 'Progresso: -';
    msgEl.textContent=String(a.message||'Nenhuma ação manual em andamento.');
    detailEl.textContent=String(a.detail||finishedText);
    if(status==='success')_setManualBadge('off','Concluída');
    else if(status==='error')_setManualBadge('err','Com erro');
    else _setManualBadge('off','Aguardando');
  }
  if(repBtn){
    repBtn.disabled=active;
    repBtn.textContent=active&&kind==='reprocess'?(phase==='processing'?'Lendo...':'Reprocessando...'):'Reprocessar agora';
  }
  if(recoverBtn){
    recoverBtn.disabled=active;
    recoverBtn.textContent=active&&kind==='recover_missing'?(phase==='processing'?'Lendo...':'Varrendo...'):'Recuperar faltantes';
  }
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
      manual_action:(j.manual_action||{}),
      diagnostic:(j.diagnostic||{})
    };
    document.getElementById('details').textContent=JSON.stringify(diag,null,2);
    _nextRemain=Number((j.scheduler&&j.scheduler.next_in_seconds)||0);
    _tickNext();
    updAccount(j.account||{});
    setPill(ok,reading);
    updDaily(j.daily_report||{});
    updProcessing(j.processing||{}, Number(j.max_messages||100), j.manual_action||{});
    updManualAction(j.manual_action||{}, j.processing||{}, Number(j.max_messages||100));
  }catch(err){
    document.getElementById('details').textContent=JSON.stringify({erro:String((err&&err.message)||err||'Falha ao atualizar estado')},null,2);
  }
}
async function startLoop(){await api('/api/start',{method:'POST'});refresh();}
async function stopLoop(){await api('/api/stop',{method:'POST'});refresh();}
async function runNow(){
  const btn=document.getElementById('runNowBtn');
  if(btn){btn.disabled=true;btn.textContent='Iniciando...';}
  const msgEl=document.getElementById('manualActionMsg');
  const detailEl=document.getElementById('manualActionDetail');
  if(msgEl)msgEl.textContent='Solicitação enviada. Iniciando execução manual...';
  if(detailEl)detailEl.textContent='O painel vai atualizar automaticamente quando o backend aceitar a execução.';
  try{
    const j=await api('/api/run-now',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({account:'principal'})});
    if(j&&j.message&&msgEl)msgEl.textContent=String(j.message);
    await refresh();
  }catch(err){
    if(msgEl)msgEl.textContent='Falha ao iniciar a execução manual.';
    if(detailEl)detailEl.textContent=String(err&&err.message||err);
    alert('Erro ao executar agora: '+(err&&err.message||err));
    await refresh();
  }
}
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
async function reprocess(){
  const btn=document.getElementById('reprocessBtn');
  if(btn){btn.disabled=true;btn.textContent='Iniciando...';}
  const msgEl=document.getElementById('manualActionMsg');
  const detailEl=document.getElementById('manualActionDetail');
  if(msgEl)msgEl.textContent='Solicitacao enviada. Buscando mensagens para reprocessar...';
  if(detailEl)detailEl.textContent='As labels serao atualizadas primeiro; em seguida o Botana vai reler essas mensagens e tentar relancar na planilha.';
  try{
    const payload={account:'principal',max_messages:Number(document.getElementById('limit').value||100),mark_unread:document.getElementById('unread').checked};
    const j=await api('/api/reprocess',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
    if(j&&j.friendly&&msgEl)msgEl.textContent=String(j.friendly);
    await refresh();
  }catch(err){
    if(msgEl)msgEl.textContent='Falha ao iniciar o reprocessamento.';
    if(detailEl)detailEl.textContent=String(err&&err.message||err);
    alert('Erro ao reprocessar: '+(err&&err.message||err));
    await refresh();
  }
}
async function recoverMissing(){
  const btn=document.getElementById('recoverBtn');
  if(btn){btn.disabled=true;btn.textContent='Iniciando...';}
  const msgEl=document.getElementById('manualActionMsg');
  const detailEl=document.getElementById('manualActionDetail');
  if(msgEl)msgEl.textContent='Solicitação enviada. Varrendo Gmail em busca de mensagens sem label do Botana...';
  if(detailEl)detailEl.textContent='O Botana vai procurar mensagens dentro do período e/ou faixa de NF informados e reler somente as que ainda não foram marcadas.';
  try{
    const payload={
      max_messages:Number(document.getElementById('recoverLimit').value||200),
      nf_start:String(document.getElementById('recoverNfStart').value||'').trim(),
      nf_end:String(document.getElementById('recoverNfEnd').value||'').trim(),
      date_from:String(document.getElementById('recoverDateFrom').value||'').trim(),
      date_to:String(document.getElementById('recoverDateTo').value||'').trim(),
    };
    const j=await api('/api/recover-missing',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
    if(j&&j.friendly&&msgEl)msgEl.textContent=String(j.friendly);
    await refresh();
  }catch(err){
    if(msgEl)msgEl.textContent='Falha ao iniciar a recuperação.';
    if(detailEl)detailEl.textContent=String(err&&err.message||err);
    alert('Erro ao recuperar faltantes: '+(err&&err.message||err));
    await refresh();
  }
}
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
let _auditLoadSeq=0;
let _watchLoadSeq=0;
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
function _historyParams(){
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
  return p;
}
function exportHistoryCsv(){
  try{
    const p=_historyParams();
    window.location.assign(_url('/api/history/export?'+p.toString()));
  }catch(err){
    console.warn('Erro ao exportar histÃ³rico:',err);
  }
}
async function loadHistory(silent=false){
  try{
    const p=_historyParams();
    const j=await api('/api/history?'+p.toString());
    const items=(j&&Array.isArray(j.items))?j.items:[];
    _renderHistory(items);
  }catch(err){
    if(!silent)console.warn('Erro ao carregar histórico:',err);
    const body=document.getElementById('hBody');
    if(body)body.innerHTML='<tr><td colspan="9">Erro de rede: '+_esc(String(err&&err.message||err))+'</td></tr>';
  }
}
function toggleAuditFilters(){
  const mode=((document.getElementById('aMode')||{}).value||'mes').trim();
  const monthEl=document.getElementById('aMonth');
  const startEl=document.getElementById('aNfStart');
  const endEl=document.getElementById('aNfEnd');
  const useMonth=mode==='mes';
  const useNf=mode==='nfs';
  if(monthEl)monthEl.disabled=!useMonth;
  if(startEl)startEl.disabled=!useNf;
  if(endEl)endEl.disabled=!useNf;
}
function _fmtAuditList(values){
  const arr=(Array.isArray(values)?values:[]).map(v=>_compactSpaces(v)).filter(Boolean);
  return arr.length?arr.join(', '):'-';
}
function _fmtAuditDate(v){
  const txt=String(v||'').trim();
  if(!txt)return '-';
  if(/^\\d{4}-\\d{2}-\\d{2}$/.test(txt)){
    const [y,m,d]=txt.split('-');
    return `${d}/${m}/${y}`;
  }
  if(/^\\d{2}\\/\\d{2}\\/\\d{4}$/.test(txt))return txt;
  return _fmtDateTime(txt);
}
function _setAuditStatus(message,loading=false){
  const el=document.getElementById('auditStatus');
  if(!el)return;
  el.textContent=String(message||'Pronto para conferir.');
  el.classList.toggle('loading',!!loading);
}
function _setAuditLoading(active,message=''){
  const btn=document.getElementById('auditRunBtn');
  if(btn){
    btn.disabled=!!active;
    btn.textContent=active?'Conferindo...':'Conferir parcelas';
  }
  _setAuditStatus(message|| (active?'Conferindo planilhas...':'Pronto para conferir.'),active);
}
function _setAuditSummary(summary){
  const s=summary||{};
  [['auditK1','nfs_verificadas'],['auditK2','nfs_com_divergencia'],['auditK3','parcelas_esperadas'],['auditK4','parcelas_lancadas'],['auditK5','parcelas_duplicadas']].forEach(([id,key])=>{
    const el=document.getElementById(id);
    if(el)el.textContent=String(s[key]||0);
  });
}
function _renderParcelAudit(items){
  const body=document.getElementById('aBody');
  if(!body)return;
  const arr=Array.isArray(items)?items:[];
  body.innerHTML='';
  if(!arr.length){
    body.innerHTML='<tr><td colspan="9">Nenhuma NF encontrada para os filtros selecionados</td></tr>';
    return;
  }
  arr.forEach(it=>{
    const tr=document.createElement('tr');
    if(it.status==='erro')tr.classList.add('audit-row-erro');
    else if(it.status==='aviso')tr.classList.add('audit-row-aviso');
    const clienteView=_compactClienteLabel(it.cliente,it.descricao);
    const duplicadasTxt=Number(it.qtd_duplicada||0)>0?_fmtAuditList(it.parcelas_duplicadas):'0';
    const local=_fmtLocal(it.local_lancamento||it.aba||'-');
    const ultimoVenc=_fmtAuditDate(it.ultimo_vencimento||it.ultimo_lancamento);
    const deleteCandidates=Number(it.delete_candidates||0);
    const statusLabel=_esc(it.status_label||'-');
    const auditKey=_esc(it.audit_key||'');
    const nfValue=_esc(it.nf||'-');
    const statusCell=(it.status&&it.status!=='ok')
      ? `<button type="button" class="audit-status audit-status-btn ${_esc(it.status||'ok')}" title="Clique para excluir linhas excedentes/duplicadas desta NF direto da planilha" onclick="deleteAuditRows('${auditKey}','${nfValue}','${statusLabel}',${deleteCandidates})">${statusLabel}</button>`
      : `<span class="audit-status ${_esc(it.status||'ok')}">${statusLabel}</span>`;
    tr.innerHTML=`<td>${statusCell}</td><td title="${nfValue}">${nfValue}</td><td title="${_esc(_compactSpaces(it.descricao||it.cliente||'-'))}">${_esc(clienteView)}</td><td>${_esc(String(it.qtd_esperada||0))}</td><td>${_esc(String(it.qtd_lancada||0))}</td><td>${_esc(String(it.qtd_faltando||0))}</td><td title="${_esc(duplicadasTxt)}">${_esc(String(it.qtd_duplicada||0))}${Number(it.qtd_duplicada||0)>0?` - ${_esc(duplicadasTxt)}`:''}</td><td title="${_esc(ultimoVenc)}">${_esc(ultimoVenc)}</td><td title="${_esc(local)}">${_esc(local)}</td>`;
    body.appendChild(tr);
  });
}
async function deleteAuditRows(auditKey,nf,statusLabel,deleteCandidates){
  const key=String(auditKey||'').trim();
  const nfView=String(nf||'-').trim()||'-';
  const statusView=String(statusLabel||'-').trim()||'-';
  const count=Math.max(0, Number(deleteCandidates||0));
  if(!key){
    alert('Não foi possível identificar a NF selecionada.');
    return;
  }
  if(count<=0){
    alert(`A NF ${nfView} está com status ${statusView}, mas não há linhas pendentes removíveis automaticamente na planilha.`);
    return;
  }
  const msg=`Tem certeza que deseja excluir ${count} linha(s) excedente(s)/duplicada(s) da NF ${nfView} direto da planilha?\nEssa ação remove apenas as linhas identificadas como sobra na Conferência.`;
  if(!confirm(msg))return;
  try{
    const r=await api('/api/conferencia-parcelas/delete',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({audit_key:key})});
    alert(String((r&&r.message)||'Linhas removidas com sucesso.'));
    await loadParcelAudit(false);
  }catch(e){
    alert('Erro ao excluir na planilha: '+e.message);
  }
}
async function loadParcelAudit(silent=false){
  const reqId=++_auditLoadSeq;
  const showLoading=!silent||_activeTab==='audit';
  if(showLoading){
    _setAuditLoading(true,'Conferindo planilhas...');
    const body=document.getElementById('aBody');
    if(body)body.innerHTML='<tr><td colspan="9">Conferindo planilhas...</td></tr>';
  }
  try{
    const p=new URLSearchParams();
    const mode=((document.getElementById('aMode')||{}).value||'mes').trim()||'mes';
    const monthValue=((document.getElementById('aMonth')||{}).value||'').trim();
    const nfStart=((document.getElementById('aNfStart')||{}).value||'').trim();
    const nfEnd=((document.getElementById('aNfEnd')||{}).value||'').trim();
    p.set('filtro',mode);
    if(mode==='mes'&&monthValue)p.set('mes',monthValue);
    if(mode==='nfs'&&nfStart)p.set('nf_inicio',nfStart);
    if(mode==='nfs'&&nfEnd)p.set('nf_fim',nfEnd);
    const j=await api('/api/conferencia-parcelas?'+p.toString());
    if(reqId!==_auditLoadSeq)return;
    _setAuditSummary(j&&j.summary||{});
    _renderParcelAudit(j&&j.items||[]);
    const meta=(j&&j.meta)||{};
    const loadedAt=String(meta.loaded_at||'').trim();
    const linhas=Number(meta.linhas_lidas||0);
    const abas=Number(meta.abas_lidas||0);
    const statusMsg=loadedAt
      ? `Conferência atualizada em ${_fmtDateTime(loadedAt)} | ${linhas} linhas lidas em ${abas} abas`
      : 'Conferência atualizada.';
    _setAuditLoading(false,statusMsg);
  }catch(err){
    if(reqId!==_auditLoadSeq)return;
    if(!silent)console.warn('Erro ao carregar conferência:',err);
    _setAuditSummary({});
    const body=document.getElementById('aBody');
    if(body)body.innerHTML='<tr><td colspan="9">Erro de rede: '+_esc(String(err&&err.message||err))+'</td></tr>';
    _setAuditLoading(false,'Falha ao conferir as planilhas.');
  }
}
function _setWatchSummary(summary){
  const s=summary||{};
  [['watchK1','total_itens'],['watchK2','boletos_a_vencer'],['watchK3','boletos_vencidos'],['watchK4','depositos_atrasados']].forEach(([id,key])=>{
    const el=document.getElementById(id);
    if(el)el.textContent=String(s[key]||0);
  });
}
function _resetWatchSummary(){
  _setWatchSummary({total_itens:0,boletos_a_vencer:0,boletos_vencidos:0,depositos_atrasados:0});
}
function _setWatchStatus(message,loading=false){
  const el=document.getElementById('watchStatus');
  if(!el)return;
  el.textContent=String(message||'Pronto para consultar.');
  el.classList.toggle('loading',!!loading);
}
function _setWatchLoading(active,message=''){
  const btn=document.getElementById('watchRunBtn');
  if(btn){
    btn.disabled=!!active;
    btn.textContent=active?'Atualizando...':'Atualizar relação';
  }
  _setWatchStatus(message||(active?'Lendo planilhas...':'Pronto para consultar.'),active);
}
function _renderDueWatch(items){
  const body=document.getElementById('wBody');
  if(!body)return;
  const arr=Array.isArray(items)?items:[];
  body.innerHTML='';
  if(!arr.length){
    body.innerHTML='<tr><td colspan="8">Nenhum boleto ou depósito pendente encontrado para os limites escolhidos</td></tr>';
    return;
  }
  arr.forEach(it=>{
    const tr=document.createElement('tr');
    if(it.status==='erro')tr.classList.add('watch-row-erro');
    else tr.classList.add('watch-row-aviso');
    const clienteView=_compactClienteLabel(it.cliente,it.descricao);
    const local=_compactSpaces(it.local||it.aba||'-');
    tr.innerHTML=`<td>${_esc(it.tipo_label||'-')}</td><td><span class="watch-badge ${_esc(it.status||'aviso')}">${_esc(it.status_label||'-')}</span></td><td title="${_esc(_fmtAuditDate(it.vencimento))}">${_esc(_fmtAuditDate(it.vencimento))}</td><td title="${_esc(it.dias_label||'-')}">${_esc(it.dias_label||'-')}</td><td title="${_esc(_compactSpaces(it.descricao||it.cliente||'-'))}">${_esc(clienteView)}</td><td title="${_esc(it.nf||'-')}">${_esc(it.nf||'-')}</td><td title="${_esc(_fmtMoney(it.valor))}">${_esc(_fmtMoney(it.valor))}</td><td title="${_esc(local)}">${_esc(local)}</td>`;
    body.appendChild(tr);
  });
}
async function loadDueWatch(silent=false){
  const reqId=++_watchLoadSeq;
  const showLoading=!silent||_activeTab==='watch';
  const boletoInput=document.getElementById('wBoletoDays');
  const depositoInput=document.getElementById('wDepositoDays');
  let boletoDays=Math.max(1,Math.min(7,Number((boletoInput&&boletoInput.value)||7)||7));
  let depositoDays=Math.max(1,Math.min(7,Number((depositoInput&&depositoInput.value)||7)||7));
  if(boletoInput)boletoInput.value=String(boletoDays);
  if(depositoInput)depositoInput.value=String(depositoDays);
  if(showLoading){
    _resetWatchSummary();
    _setWatchLoading(true,'Lendo planilhas...');
    const body=document.getElementById('wBody');
    if(body)body.innerHTML='<tr><td colspan="8">Lendo planilhas...</td></tr>';
  }
  try{
    const p=new URLSearchParams();
    p.set('boleto_dias',String(boletoDays));
    p.set('deposito_dias',String(depositoDays));
    const j=await api('/api/prazos?'+p.toString());
    if(reqId!==_watchLoadSeq)return;
    _setWatchSummary(j&&j.summary||{});
    _renderDueWatch(j&&j.items||[]);
    const meta=(j&&j.meta)||{};
    const loadedAt=String(meta.loaded_at||'').trim();
    const linhas=Number(meta.linhas_lidas||0);
    const abas=Number(meta.abas_lidas||0);
    const statusMsg=loadedAt
      ? `Relação atualizada em ${_fmtDateTime(loadedAt)} | ${linhas} linhas lidas em ${abas} abas`
      : 'Relação atualizada.';
    _setWatchLoading(false,statusMsg);
  }catch(err){
    if(reqId!==_watchLoadSeq)return;
    _setWatchSummary({});
    const body=document.getElementById('wBody');
    if(body)body.innerHTML='<tr><td colspan="8">Erro de rede: '+_esc(String(err&&err.message||err))+'</td></tr>';
    _setWatchLoading(false,'Falha ao ler as planilhas.');
  }
}
function _setWatchSearchState(message,loading=false){
  const el=document.getElementById('watchSearchState');
  if(!el)return;
  el.textContent=String(message||'Digite um nome para consultar boletos em aberto.');
  el.classList.toggle('loading',!!loading);
}
function _renderWatchSearchResults(items,message){
  const box=document.getElementById('watchSearchResults');
  if(!box)return;
  const arr=Array.isArray(items)?items:[];
  if(!arr.length){
    box.innerHTML=`<div class="watch-pop-empty">${_esc(String(message||'Não existem pendências para a busca informada.'))}</div>`;
    return;
  }
  const rows=arr.map((it)=>{
    const clienteView=_compactClienteLabel(it.cliente,it.descricao);
    const local=_compactSpaces(it.local||it.aba||'-');
    const rowClass=String(it.status||'')==='erro'?'watch-row-erro':'watch-row-aviso';
    return `<tr class="${_esc(rowClass)}"><td title="${_esc(_compactSpaces(it.descricao||it.cliente||'-'))}">${_esc(clienteView)}</td><td title="${_esc(it.nf||'-')}">${_esc(it.nf||'-')}</td><td title="${_esc(_fmtAuditDate(it.vencimento))}">${_esc(_fmtAuditDate(it.vencimento))}</td><td><span class="watch-badge ${_esc(it.status||'aviso')}">${_esc(it.status_label||'-')}</span></td><td title="${_esc(it.dias_label||'-')}">${_esc(it.dias_label||'-')}</td><td title="${_esc(_fmtMoney(it.valor))}">${_esc(_fmtMoney(it.valor))}</td><td title="${_esc(local)}">${_esc(local)}</td></tr>`;
  }).join('');
  box.innerHTML=`<table class="watch-pop-table"><thead><tr><th>Cliente</th><th>NF</th><th>Vencimento</th><th>Situação</th><th>Prazo</th><th>Valor</th><th>Aba</th></tr></thead><tbody>${rows}</tbody></table>`;
  _setWatchSearchState(message||`${arr.length} boleto(s) em aberto encontrados.`,false);
}
let _watchSearchCatalog=[];
function _watchSearchKey(value){
  return String(value||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\\s+/g,' ').trim().toUpperCase();
}
function _setWatchSearchCatalog(items){
  const arr=Array.isArray(items)?items:[];
  const unique=[];
  const seen=new Set();
  arr.forEach((item)=>{
    const name=_compactSpaces(item||'');
    const key=_watchSearchKey(name);
    if(!key||seen.has(key))return;
    seen.add(key);
    unique.push(name);
  });
  _watchSearchCatalog=unique;
  const input=document.getElementById('watchSearchInput');
  _renderWatchSearchSuggestions(String((input&&input.value)||''));
}
function _filterWatchSearchSuggestions(query){
  const key=_watchSearchKey(query);
  if(!key)return [];
  const tokens=key.split(' ').filter(Boolean);
  return _watchSearchCatalog
    .map((name,idx)=>({name,key:_watchSearchKey(name),idx}))
    .filter((item)=>item.key&&tokens.every((token)=>item.key.includes(token)))
    .sort((a,b)=>{
      const aStart=a.key.startsWith(key)?0:1;
      const bStart=b.key.startsWith(key)?0:1;
      if(aStart!==bStart)return aStart-bStart;
      const aPos=a.key.indexOf(key);
      const bPos=b.key.indexOf(key);
      if(aPos!==bPos)return aPos-bPos;
      if(a.name.length!==b.name.length)return a.name.length-b.name.length;
      return a.idx-b.idx;
    })
    .slice(0,12)
    .map((item)=>item.name);
}
function _renderWatchSearchSuggestions(query){
  const el=document.getElementById('watchSearchSuggestions');
  if(!el)return;
  const arr=_filterWatchSearchSuggestions(query);
  if(!arr.length){
    el.innerHTML='';
    el.classList.remove('show');
    return;
  }
  el.innerHTML=arr.map((item)=>`<button type="button" class="watch-pop-suggest-item" data-value="${_esc(item)}" onclick="useWatchSearchSuggestion(this.dataset.value)">${_esc(item)}</button>`).join('');
  el.classList.add('show');
}
function useWatchSearchSuggestion(value){
  const input=document.getElementById('watchSearchInput');
  if(!input)return;
  input.value=String(value||'');
  _renderWatchSearchSuggestions('');
  input.focus();
  searchOpenBoletos();
}
async function loadWatchSearchSuggestions(){
  try{
    const j=await api('/api/prazos/search-suggestions');
    _setWatchSearchCatalog(j&&j.items||[]);
  }catch(err){
    console.warn('Erro ao carregar autocomplete da busca de prazos:',err);
  }
}
function openWatchSearchModal(){
  const modal=document.getElementById('watchSearchModal');
  if(modal)modal.classList.add('show');
  _setWatchSearchState('Digite um nome para consultar boletos em aberto.',false);
  _renderWatchSearchResults([], 'Nenhuma busca executada ainda.');
  loadWatchSearchSuggestions().catch(()=>{});
  const input=document.getElementById('watchSearchInput');
  if(input){setTimeout(()=>{input.focus();_renderWatchSearchSuggestions(input.value||'');},20);}
}
function closeWatchSearchModal(ev){
  if(ev&&ev.target&&ev.currentTarget&&ev.target!==ev.currentTarget)return;
  const modal=document.getElementById('watchSearchModal');
  if(modal)modal.classList.remove('show');
  _renderWatchSearchSuggestions('');
}
async function searchOpenBoletos(){
  const input=document.getElementById('watchSearchInput');
  const btn=document.getElementById('watchSearchBtn');
  const query=String((input&&input.value)||'').trim();
  if(!query){
    _renderWatchSearchSuggestions('');
    _renderWatchSearchResults([], 'Informe um nome para consultar.');
    return;
  }
  if(btn){btn.disabled=true;btn.textContent='Buscando...';}
  _renderWatchSearchSuggestions('');
  _setWatchSearchState('Consultando boletos em aberto nas planilhas...',true);
  try{
    const j=await api('/api/prazos/search?nome='+encodeURIComponent(query));
    _setWatchSearchCatalog(j&&j.suggestions||[]);
    _renderWatchSearchResults(j&&j.items||[], String((j&&j.message)||'Busca concluída.'));
  }catch(err){
    _renderWatchSearchResults([], 'Falha ao consultar boletos em aberto.');
    _setWatchSearchState(String(err&&err.message||err),false);
  }finally{
    if(btn){btn.disabled=false;btn.textContent='Buscar';}
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
['limit'].forEach(id=>{const el=document.getElementById(id);if(!el)return;el.addEventListener('keydown',(e)=>{if(e.key==='Enter'){e.preventDefault();reprocess();}});});
['recoverDateFrom','recoverDateTo','recoverNfStart','recoverNfEnd','recoverLimit'].forEach(id=>{const el=document.getElementById(id);if(!el)return;el.addEventListener('keydown',(e)=>{if(e.key==='Enter'){e.preventDefault();recoverMissing();}});});
document.querySelectorAll('#hAt,#hVenc,#hNf,#hCliente,#hAba,#hLimit').forEach(el=>{el.addEventListener('keydown',(e)=>{if(e.key==='Enter'){e.preventDefault();loadHistory();}});});
document.querySelectorAll('#aMode,#aMonth,#aNfStart,#aNfEnd').forEach(el=>{el.addEventListener('keydown',(e)=>{if(e.key==='Enter'){e.preventDefault();loadParcelAudit();}});});
document.querySelectorAll('#wBoletoDays,#wDepositoDays').forEach(el=>{el.addEventListener('keydown',(e)=>{if(e.key==='Enter'){e.preventDefault();loadDueWatch();}});});
['watchSearchInput'].forEach(id=>{const el=document.getElementById(id);if(!el)return;el.addEventListener('keydown',(e)=>{if(e.key==='Enter'){e.preventDefault();searchOpenBoletos();}});});
['watchSearchInput'].forEach(id=>{const el=document.getElementById(id);if(!el)return;el.addEventListener('input',()=>{_renderWatchSearchSuggestions(el.value||'');});el.addEventListener('focus',()=>{_renderWatchSearchSuggestions(el.value||'');});});
window.addEventListener('keydown',(e)=>{
  if(e.key!=='Escape')return;
  const modal=document.getElementById('watchSearchModal');
  if(modal&&modal.classList.contains('show')){
    e.preventDefault();
    closeWatchSearchModal();
  }
});
window.addEventListener('hashchange',()=>{const t=_tabFromLocation();if(t!==_activeTab)switchTab(t);});
const _auditMonthEl=document.getElementById('aMonth');
if(_auditMonthEl&&!_auditMonthEl.value){
  try{_auditMonthEl.value=new Date().toISOString().slice(0,7);}catch(_){}
}
toggleAuditFilters();
refresh();loadHistory();switchTab(_tabFromLocation());setInterval(refresh,3000);setInterval(_tickNext,1000);setInterval(()=>{if(_activeTab==='hist')loadHistory(true);},10000);
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
            if self.client_address[0] in ["127.0.0.1", "::1", "localhost"] and parsed_path in ["/api/relatorio-nfs", "/api/conferencia-parcelas", "/api/prazos", "/api/prazos/search", "/api/prazos/search-suggestions", "/api/clean-sheets", "/api/clean-sheets/log"]:
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
                    friendly = "Monitoramento pausado. Use o painel para reprocessar mensagens ou inicie o loop."
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
                        "manual_action": _manual_action_snapshot(),
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

            if parsed.path == "/api/conferencia-parcelas":
                qs = parse_qs(parsed.query or "")
                filtro = (qs.get("filtro", [""])[0] or "mes").strip()
                mes = (qs.get("mes", [""])[0] or "").strip()
                nf_inicio = (qs.get("nf_inicio", [""])[0] or "").strip()
                nf_fim = (qs.get("nf_fim", [""])[0] or "").strip()
                try:
                    resultado = _gerar_conferencia_parcelas(filtro, mes, nf_inicio, nf_fim)
                    return _json_response(self, 200, {"ok": True, **resultado})
                except Exception as e:
                    return _json_response(self, 500, {"ok": False, "message": str(e)})

            if parsed.path == "/api/prazos":
                qs = parse_qs(parsed.query or "")
                try:
                    boleto_dias = int((qs.get("boleto_dias", ["7"])[0] or "7").strip())
                except Exception:
                    boleto_dias = 7
                try:
                    deposito_dias = int((qs.get("deposito_dias", ["7"])[0] or "7").strip())
                except Exception:
                    deposito_dias = 7
                try:
                    resultado = _gerar_relacao_pendencias(boleto_dias, deposito_dias)
                    return _json_response(self, 200, {"ok": True, **resultado})
                except Exception as e:
                    return _json_response(self, 500, {"ok": False, "message": str(e)})

            if parsed.path == "/api/prazos/search-suggestions":
                try:
                    return _json_response(self, 200, {"ok": True, "items": _load_watch_search_suggestions()})
                except Exception as e:
                    return _json_response(self, 500, {"ok": False, "message": str(e)})

            if parsed.path == "/api/prazos/search":
                qs = parse_qs(parsed.query or "")
                nome = (qs.get("nome", [""])[0] or "").strip()
                if not nome:
                    return _json_response(self, 400, {"ok": False, "message": "Informe um nome para consultar."})
                try:
                    resultado = _buscar_boletos_em_aberto_por_nome(nome)
                    return _json_response(self, 200, {"ok": True, **resultado})
                except Exception as e:
                    return _json_response(self, 500, {"ok": False, "message": str(e)})

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
            if parsed.path == "/api/history/export":
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
                file_name = f"historico_botana_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
                return _csv_response(self, 200, file_name, _history_csv_rows(items))
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
                    try:
                        max_messages = max(1, min(1000, int(data.get("max_messages", 0) or 0)))
                    except Exception:
                        max_messages = 0
                    started, info = _start_run_now_background(max_messages_override=max_messages or None)
                    if not started:
                        msg = _manual_action_busy_message() or str((info or {}).get("message") or "Não foi possível iniciar a execução manual.")
                        return _json_response(self, 409, {"ok": False, "message": msg, "action": info})
                    friendly = "Execucao manual iniciada."
                    if max_messages:
                        friendly = f"Execucao manual iniciada com limite de {max_messages} mensagens."
                    return _json_response(self, 202, {"ok": True, "started": True, "message": friendly, "action": info})
                if parsed.path == "/api/reprocess":
                    if not _can_operate(user):
                        return _json_response(self, 403, {"ok": False, "message": "Sem permissao"})
                    try:
                        max_messages = max(1, min(1000, int(data.get("max_messages", 100))))
                    except Exception:
                        max_messages = 100
                    mark_unread = bool(data.get("mark_unread", True))
                    started, info = _start_reprocess_background(max_messages=max_messages, mark_unread=mark_unread)
                    if not started:
                        msg = _manual_action_busy_message() or str((info or {}).get("message") or "Nao foi possivel iniciar o reprocessamento.")
                        return _json_response(self, 409, {"ok": False, "message": msg, "action": info})
                    friendly = f"Reprocessamento iniciado para ate {max_messages} mensagens mais recentes; a leitura sera executada em seguida."
                    return _json_response(self, 202, {"ok": True, "started": True, "friendly": friendly, "action": info})
                if parsed.path == "/api/recover-missing":
                    if not _can_operate(user):
                        return _json_response(self, 403, {"ok": False, "message": "Sem permissao"})
                    try:
                        max_messages = max(1, min(1000, int(data.get("max_messages", 200))))
                    except Exception:
                        max_messages = 200
                    nf_start = str(data.get("nf_start", "") or "").strip()
                    nf_end = str(data.get("nf_end", "") or "").strip()
                    date_from = str(data.get("date_from", "") or "").strip()
                    date_to = str(data.get("date_to", "") or "").strip()
                    started, info = _start_recover_missing_background(
                        max_messages=max_messages,
                        nf_start=nf_start,
                        nf_end=nf_end,
                        date_from=date_from,
                        date_to=date_to,
                    )
                    if not started:
                        msg = str((info or {}).get("message") or "").strip()
                        if not msg:
                            msg = _manual_action_busy_message() or "Nao foi possivel iniciar a recuperacao."
                        status_code = 400 if "Informe" in msg else 409
                        return _json_response(self, status_code, {"ok": False, "message": msg, "action": info})
                    filtros = _describe_recovery_filters(nf_start=nf_start, nf_end=nf_end, date_from=date_from, date_to=date_to) or "os filtros informados"
                    friendly = f"Recuperacao iniciada para ate {max_messages} mensagens sem label do Botana em {filtros}."
                    return _json_response(self, 202, {"ok": True, "started": True, "friendly": friendly, "action": info})
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

                if parsed.path == "/api/conferencia-parcelas/delete":
                    if not _can_operate(user):
                        return _json_response(self, 403, {"ok": False, "message": "Sem permissão"})
                    auditKey = str(data.get("audit_key", "") or "").strip()
                    if not auditKey:
                        return _json_response(self, 400, {"ok": False, "message": "NF da conferência não informada"})
                    resultado = _delete_audit_rows(auditKey)
                    statusCode = 200 if resultado.get("ok") else 400
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
