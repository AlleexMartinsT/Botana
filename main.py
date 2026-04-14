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
from googleapiclient.discovery import build
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
_REPROCESS_LOOKBACK_DAYS = 14
_EMAIL_CACHE = {"email": "", "error": "", "pending": False, "at": 0.0}
_NEXT_RUN_AT = 0.0
_GMAIL_SERVICE_LOCK = threading.Lock()
_IS_READING = False
_PROCESS_STATS_LOCK = threading.Lock()
_PROCESS_EXEC_LOCK = threading.Lock()
_AUDIT_DELETE_CACHE_LOCK = threading.Lock()
_AUDIT_DELETE_CACHE = {
    "at": 0.0,
    "items": {},
}
_AUDIT_MONTH_MISSING_CACHE_LOCK = threading.Lock()
_AUDIT_MONTH_MISSING_CACHE = {}
_AUDIT_MONTH_MISSING_CACHE_TTL = 10 * 60
_AUDIT_MONTH_GAP_LIMIT = 12
_AUDIT_MONTH_MAX_GAP_CHECKS = 18
_AUDIT_SHEET_CACHE_LOCK = threading.Lock()
_AUDIT_SHEET_CACHE = {
    "MVA": {"at": 0.0, "rows": [], "meta": {}},
    "EH": {"at": 0.0, "rows": [], "meta": {}},
}
_AUDIT_SHEET_CACHE_TTL = 12
_AUDIT_META_CACHE_LOCK = threading.Lock()
_AUDIT_META_CACHE = {}
_AUDIT_META_CACHE_TTL = 20
_AUDIT_JOB_LOCK = threading.Lock()
_AUDIT_JOB_SNAPSHOTS = {}
_AUDIT_JOBS = {}
_AUDIT_JOBS_BY_REQUEST = {}
_AUDIT_JOB_SNAPSHOT_TTL = 20 * 60
_AUDIT_JOB_STATE_TTL = 30 * 60
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
    "launched": 0,
    "duplicates": 0,
    "messages_read": 0,
    "subject_mismatch_count": 0,
    "subject_mismatch_notes": [],
    "current_email": "",
    "current_subject": "",
    "current_date": "",
    "requested_limit": 0,
    "window_oldest_date": "",
    "window_newest_date": "",
    "window_selected": 0,
    "continue_after_id": "",
    "continue_remaining": 0,
    "requested_nf_count": 0,
    "found_nf_numbers": [],
    "missing_nf_numbers": [],
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
                "launched": 0,
                "duplicates": 0,
                "messages_read": 0,
                "subject_mismatch_count": 0,
                "subject_mismatch_notes": [],
                "current_email": "",
                "current_subject": "",
                "current_date": "",
                "requested_limit": 0,
                "window_oldest_date": "",
                "window_newest_date": "",
                "window_selected": 0,
                "continue_after_id": "",
                "continue_remaining": 0,
                "requested_nf_count": 0,
                "found_nf_numbers": [],
                "missing_nf_numbers": [],
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
        _MANUAL_ACTION["window_oldest_date"] = ""
        _MANUAL_ACTION["window_newest_date"] = ""
        _MANUAL_ACTION["window_selected"] = 0
        _MANUAL_ACTION["continue_after_id"] = ""
        _MANUAL_ACTION["continue_remaining"] = 0
        _MANUAL_ACTION["requested_nf_count"] = 0
        _MANUAL_ACTION["found_nf_numbers"] = []
        _MANUAL_ACTION["missing_nf_numbers"] = []
        _MANUAL_ACTION["matched"] = 0
        _MANUAL_ACTION["inspected"] = 0
        _MANUAL_ACTION["launched"] = 0
        _MANUAL_ACTION["duplicates"] = 0
        _MANUAL_ACTION["messages_read"] = 0
        _MANUAL_ACTION["subject_mismatch_count"] = 0
        _MANUAL_ACTION["subject_mismatch_notes"] = []
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
            else "Acompanhe o progresso no painel."
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
                    else "Acompanhe o progresso no painel."
                ),
            )
            ok, msg = executar_um_ciclo(max_messages_override=max_messages_override)
            proc = _process_snapshot().get("last", {})
            detail = f"E-mails lidos: {int(proc.get('messages', 0) or 0)}"
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


def _start_reprocess_background(max_messages: int, mark_unread: bool = False, continue_after_id: str = "") -> tuple[bool, dict]:
    snap = _manual_action_snapshot()
    if bool(snap.get("active")):
        if not snap.get("message"):
            label = str(snap.get("label") or "Acao manual").strip() or "Acao manual"
            snap["message"] = f"{label} ja esta em andamento."
        return False, snap
    continue_after_id = str(continue_after_id or "").strip()
    initial_detail = (
        f"At\u00e9 {int(max_messages)} mensagens mais recentes com label do Botana "
        f"nos \u00faltimos {_REPROCESS_LOOKBACK_DAYS} dias ser\u00e3o remarcadas e relidas neste ciclo."
    )
    if continue_after_id:
        initial_detail = (
            f"Continua\u00e7\u00e3o do reprocessamento: at\u00e9 {int(max_messages)} mensagens mais antigas "
            f"dentro da janela de {_REPROCESS_LOOKBACK_DAYS} dias ser\u00e3o remarcadas e relidas neste ciclo."
        )
    started, snap = _manual_action_begin(
        "reprocess",
        "Reprocessamento",
        "Reprocessamento iniciado.",
        detail=initial_detail,
        progress_total=int(max_messages),
        requested_limit=int(max_messages),
        continue_after_id=continue_after_id,
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
                detail=(
                    "Atualizando a label do Botana nas mensagens mais recentes antes de reler os e-mails."
                    if not continue_after_id
                    else "Atualizando a label do Botana no próximo lote mais antigo antes de reler os e-mails."
                ),
                phase="marking",
            )
            result = _reprocess_recent(
                max_messages=max_messages,
                mark_unread=mark_unread,
                progress_cb=_progress,
                continue_after_id=continue_after_id,
            )
            targets = list(result.get("targets") or [])
            changed = int(result.get("changed", 0) or 0)
            failed = int(result.get("failed", 0) or 0)
            matched = int(result.get("matched", 0) or 0)
            window_oldest_date = str(result.get("window_oldest_date", "") or "").strip()
            window_newest_date = str(result.get("window_newest_date", "") or "").strip()
            window_selected = int(result.get("window_selected", 0) or 0)
            next_continue_after_id = str(result.get("continue_after_id", "") or "").strip()
            continue_remaining = int(result.get("continue_remaining", 0) or 0)
            window_text = ""
            if window_oldest_date and window_newest_date:
                window_text = f"Lote selecionado: {window_selected or matched} mensagens, de {window_newest_date} até {window_oldest_date}."
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
                    window_oldest_date=window_oldest_date,
                    window_newest_date=window_newest_date,
                    window_selected=window_selected,
                    continue_after_id=next_continue_after_id,
                    continue_remaining=continue_remaining,
                    message="Preparação concluída. Iniciando leitura dos e-mails reprocessados.",
                    detail=(window_text or ""),
                )
                ok, msg = executar_um_ciclo(
                    max_messages_override=len(targets),
                    messages_override=targets,
                    preserve_reprocess_label=True,
                )
                proc = _process_snapshot().get("last", {})
                detail = (
                    f"Falhas ao marcar: {failed} | "
                    f"E-mails lidos: {int(proc.get('messages', 0) or 0)}"
                )
                if window_text:
                    detail = f"{detail} | {window_text}"
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
                    window_oldest_date=window_oldest_date,
                    window_newest_date=window_newest_date,
                    window_selected=window_selected,
                    continue_after_id=next_continue_after_id,
                    continue_remaining=continue_remaining,
                )
                return
            if matched <= 0:
                friendly = (
                    f"Nenhuma mensagem com label do Botana foi encontrada para reprocessar "
                    f"nos \u00faltimos {_REPROCESS_LOOKBACK_DAYS} dias."
                )
            elif changed <= 0:
                friendly = "Nenhuma mensagem foi marcada para reprocessar."
            else:
                friendly = f"Reprocessamento concluido: {changed} de {matched} mensagens preparadas."
            detail = f"Falhas: {failed}"
            if window_text:
                detail = f"{detail} | {window_text}"
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
                window_oldest_date=window_oldest_date,
                window_newest_date=window_newest_date,
                window_selected=window_selected,
                continue_after_id=next_continue_after_id,
                continue_remaining=continue_remaining,
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
    seen_regions = set()
    for marker in ("VENCIMENTO", "VENCTO", "DATA DE VENCIMENTO", "VALOR DOCUMENTO"):
        for match in re.finditer(marker, upper):
            start = max(0, match.start() - 12)
            end = min(len(text), match.start() + 420)
            key = (start, end)
            if key in seen_regions:
                continue
            seen_regions.add(key)
            regions.append(text[start:end])
    if not regions:
        regions.append(text[:520])
    for region in regions:
        for date_match in re.finditer(r"\b(\d{2}/\d{2}/\d{4})\b", region):
            vencimento = _normalize_ddmmyyyy(date_match.group(1))
            if not vencimento:
                continue
            tail = region[date_match.end() : date_match.end() + 180]
            value_match = re.search(r"\b(\d{1,3}(?:\.\d{3})*,\d{2})\b", tail)
            if not value_match:
                continue
            valor = _safe_money_text(value_match.group(1))
            if valor > 0:
                return vencimento, valor
    combined = re.search(
        r"\b(\d{2}/\d{2}/\d{4})\b[\s\S]{0,120}?\b(\d{1,3}(?:\.\d{3})*,\d{2})\b",
        text,
    )
    if combined:
        vencimento = _normalize_ddmmyyyy(combined.group(1))
        valor = _safe_money_text(combined.group(2))
        if vencimento and valor > 0:
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
    candidatos = _boletos_for_xml(payload, boleto_infos)
    if not candidatos:
        return payload
    if not all(str(item.get("vencimento") or "").strip() and float(item.get("valor") or 0.0) > 0 for item in candidatos):
        return payload
    valor_total = float(payload.get("valorTotal") or 0.0)
    source = str(payload.get("parcelas_source") or "").strip().lower()
    if parcelas:
        if source != "fat" or len(parcelas) != 1 or len(candidatos) <= 1:
            return payload
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
    else:
        soma_boletos = round(sum(float(item.get("valor") or 0.0) for item in candidatos), 2)
        if len(candidatos) == 1:
            boleto_valor = round(float(candidatos[0].get("valor") or 0.0), 2)
            if valor_total > 0 and abs(boleto_valor - valor_total) > 0.05:
                logger.warning(
                    "NF %s sem parcelas no XML teve 1 boleto PDF, mas o valor %.2f difere do total %.2f. Mantendo XML.",
                    payload.get("nf"),
                    boleto_valor,
                    valor_total,
                )
                return payload
        elif valor_total > 0 and abs(soma_boletos - valor_total) > 0.05:
            logger.warning(
                "NF %s sem parcelas no XML teve %d boletos PDF, mas a soma %.2f difere do total %.2f. Mantendo XML.",
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
    reason = "o XML veio apenas com fatura total" if source == "fat" else "o XML veio sem parcelas"
    logger.info(
        "NF %s inferiu %d parcelas a partir dos PDFs de boleto porque %s.",
        payload.get("nf"),
        len(parcelas_inferidas),
        reason,
    )
    return payload


def _message_payment_hint(message_meta: dict | None) -> str:
    if not isinstance(message_meta, dict):
        return ""
    text = " ".join(
        str(message_meta.get(key, "") or "").strip()
        for key in ("subject", "snippet")
    )
    key = f" {_normalize_ascii_key(text)} "
    has_boleto = any(token in key for token in (" BOLETO ", " BLT ", " BOLET0 ", " BOLETT "))
    has_deposito = any(token in key for token in (" DEPOSITO ", " DEP CX ", " DEP BR ", " CONTA CAIXA ", " CONTA BRADESCO "))
    if has_deposito and not has_boleto:
        return "deposito"
    if has_boleto and not has_deposito:
        return "boleto"
    return ""


def _should_ignore_cash_sale_xml(dados_xml: dict | None, message_meta: dict | None = None) -> bool:
    payload = dict(dados_xml or {})
    nat_op = str(payload.get("naturezaOperacao") or "").strip().upper()
    if "VISTA" not in nat_op and "VENDA A VISTA" not in nat_op:
        return False
    payment_hint = _message_payment_hint(message_meta)
    if payment_hint in {"deposito", "boleto"}:
        return False
    return True


def _message_date_fallback(message_meta: dict | None) -> str:
    if not isinstance(message_meta, dict):
        return ""
    raw = str(message_meta.get("date", "") or "").strip()
    if not raw:
        return ""
    try:
        dt = parsedate_to_datetime(raw)
        if dt.tzinfo is not None:
            dt = dt.astimezone()
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return _normalize_ddmmyyyy(raw)


def _infer_deposito_parcela_from_preview(dados_xml: dict, message_meta: dict | None, boleto_infos: list[dict]) -> dict:
    payload = dict(dados_xml or {})
    if list(payload.get("parcelas") or []):
        return payload
    if list(boleto_infos or []):
        return payload
    if _message_payment_hint(message_meta) != "deposito":
        return payload
    valor_total = float(payload.get("valorTotal") or 0.0)
    if valor_total <= 0:
        return payload
    vencimento = (
        str(payload.get("dataEmissao") or "").strip()
        or str(payload.get("vencimento") or "").strip()
        or _message_date_fallback(message_meta)
    )
    if not vencimento:
        return payload
    parcela = {
        "numero": 1,
        "numParcela": "1ª Parcela",
        "vencimento": vencimento,
        "valor": valor_total,
    }
    payload["parcelas"] = [parcela]
    payload["qtdParcelas"] = 1
    payload["parcelas_source"] = "preview_deposito"
    payload["vencimento"] = vencimento
    payload["numParcela"] = parcela["numParcela"]
    payload["valorParcela"] = valor_total
    try:
        payload["anoVencimento"] = datetime.strptime(vencimento, "%d/%m/%Y").strftime("%Y")
    except Exception:
        pass
    logger.info(
        "NF %s inferiu lançamento como depósito a partir do assunto/corpo do e-mail.",
        payload.get("nf"),
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
            message_meta = {
                "subject": str(m.get("subject", "") or "").strip(),
                "snippet": str(m.get("snippet", "") or "").strip(),
                "date": str(m.get("date", "") or "").strip(),
            }
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
                            # Ignora vendas à vista, exceto quando o próprio e-mail indica depósito/boleto,
                            # porque nesses casos o fallback ainda pode montar a parcela corretamente.
                            nat_op = dados.get("naturezaOperacao", "").strip().upper()
                            dest_nome = dados.get("destinatario", "")
                            dest_cnpj = re.sub(r"\D+", "", str(dados.get("cnpjDestinatario") or ""))
                            payment_hint = _message_payment_hint(message_meta)
                            if _should_ignore_cash_sale_xml(dados, message_meta):
                                # Checa se a mensagem jÃ¡ foi processada no relatÃ³rio atual:
                                if dados.get('nf') not in consolidarRelatorioTMP():
                                    escreverRelatorio(f"{_now()} - NF {dados.get('nf')} ignorada (venda Ã  vista).")
                                    continue
                                else: logger.info(f"{cor_ciano}NF {dados['nf']} jÃ¡ registrada no relatÃ³rio, nÃ£o duplicando a mensagem de ignorada.{reset}")
                                continue
                            if ("VISTA" in nat_op or "VENDA A VISTA" in nat_op) and payment_hint in {"deposito", "boleto"}:
                                logger.info(
                                    "NF %s mantida para processamento apesar de venda à vista porque o e-mail indica %s.",
                                    dados.get("nf"),
                                    payment_hint,
                                )
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
                dados_xml = _infer_deposito_parcela_from_preview(dados_xml, message_meta, xml_boletos)
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
    empresa_filter: str = "todos",
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
    empresa = _normalize_audit_empresa(empresa_filter)
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
                empresa_item = _infer_audit_empresa_from_values(sheet_title, aba, nf_value=nf)
                local = "/".join([x for x in (sheet_title, aba) if x]) or "Botana/RelatÃ³rio"

                if empresa != "todos" and empresa_item != empresa:
                    continue

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
                    "empresa": empresa_item,
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


def _normalize_audit_empresa(value: str) -> str:
    key = _normalize_ascii_key(str(value or "").strip())
    if key == "MVA":
        return "MVA"
    if key == "EH":
        return "EH"
    return "todos"


def _audit_request_signature(filtro: str, mes: str, nf_inicio: str, nf_fim: str, empresa: str) -> str:
    payload = {
        "filtro": str(filtro or "mes").strip().lower(),
        "mes": str(mes or "").strip(),
        "nf_inicio": re.sub(r"\D+", "", str(nf_inicio or ""))[:12],
        "nf_fim": re.sub(r"\D+", "", str(nf_fim or ""))[:12],
        "empresa": _normalize_audit_empresa(empresa),
    }
    raw = json.dumps(payload, ensure_ascii=True, sort_keys=True)
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()


def _audit_companies_for_filter(empresa_filter: str) -> list[str]:
    empresa = _normalize_audit_empresa(empresa_filter)
    if empresa == "todos":
        return [company for company in ("MVA", "EH") if PLANILHAS.get(company)]
    return [empresa] if PLANILHAS.get(empresa) else []


def _infer_audit_empresa_from_values(*values: str, nf_value: str = "") -> str:
    for raw in values:
        key = _normalize_ascii_key(raw)
        if not key:
            continue
        if re.search(r"(^|\b)MVA(\b|$)", key):
            return "MVA"
        if re.search(r"(^|\b)EH(\b|$)", key):
            return "EH"
    nf_digits = re.sub(r"\D+", "", str(nf_value or ""))
    try:
        nf_num = int(nf_digits)
    except Exception:
        nf_num = 0
    if nf_num >= 40000:
        return "MVA"
    if 19000 <= nf_num < 40000:
        return "EH"
    return "todos"


def _audit_merge_meta_parts(metas: list[dict], empresa_filter: str = "todos") -> dict:
    empresa = _normalize_audit_empresa(empresa_filter)
    out = {
        "loaded_at": datetime.now().isoformat(),
        "planilhas_lidas": 0,
        "abas_lidas": 0,
        "linhas_lidas": 0,
        "source": "planilhas",
        "empresa": empresa,
        "empresas_lidas": [],
    }
    seen = set()
    latest_loaded_at = ""
    latest_loaded_ms = 0
    for meta in list(metas or []):
        if not isinstance(meta, dict):
            continue
        out["planilhas_lidas"] += int(meta.get("planilhas_lidas") or 0)
        out["abas_lidas"] += int(meta.get("abas_lidas") or 0)
        out["linhas_lidas"] += int(meta.get("linhas_lidas") or 0)
        loaded_at = str(meta.get("loaded_at") or "").strip()
        companies = list(meta.get("empresas_lidas") or [])
        if not companies and str(meta.get("empresa") or "").strip():
            companies = [str(meta.get("empresa") or "").strip()]
        for company in companies:
            normalized = str(company or "").strip().upper()
            if normalized and normalized not in seen:
                seen.add(normalized)
                out["empresas_lidas"].append(normalized)
        try:
            loaded_ms = int(datetime.fromisoformat(loaded_at).timestamp() * 1000) if loaded_at else 0
        except Exception:
            loaded_ms = 0
        if loaded_ms >= latest_loaded_ms:
            latest_loaded_ms = loaded_ms
            latest_loaded_at = loaded_at
    if latest_loaded_at:
        out["loaded_at"] = latest_loaded_at
    return out


def _audit_drive_metadata(empresa_filter: str = "todos", force_refresh: bool = False) -> dict:
    companies = _audit_companies_for_filter(empresa_filter)
    now = time.time()
    items = []
    fallback_key = "__fallback__:" + ",".join(companies)
    with _AUDIT_META_CACHE_LOCK:
        cache_snapshot = dict(_AUDIT_META_CACHE)
    if not force_refresh and cache_snapshot:
        fallback_entry = dict(cache_snapshot.get(fallback_key) or {})
        fallback_at = float(fallback_entry.get("at", 0.0) or 0.0)
        if fallback_at and (now - fallback_at) <= _AUDIT_META_CACHE_TTL:
            payload = dict(fallback_entry.get("payload") or {})
            if payload:
                return payload
        fresh = True
        for company in companies:
            for year, file_id in (PLANILHAS.get(company) or {}).items():
                if not file_id:
                    continue
                entry = dict(cache_snapshot.get(str(file_id)) or {})
                cache_at = float(entry.get("at", 0.0) or 0.0)
                if not cache_at or (now - cache_at) > _AUDIT_META_CACHE_TTL:
                    fresh = False
                    break
                item = dict(entry.get("item") or {})
                if item:
                    items.append(item)
            if not fresh:
                break
        if fresh and items:
            signature_src = json.dumps(sorted(items, key=lambda item: (item.get("empresa"), item.get("ano"), item.get("id"))), ensure_ascii=True, sort_keys=True)
            return {
                "loaded_at": datetime.now().isoformat(),
                "items": items,
                "signature": hashlib.sha1(signature_src.encode("utf-8")).hexdigest(),
            }

    creds = Credentials.from_service_account_file(
        GOOGLE_CREDENTIALS_SHEETS,
        scopes=["https://www.googleapis.com/auth/drive.metadata.readonly"],
    )
    try:
        drive = build("drive", "v3", credentials=creds, cache_discovery=False)
        cache_updates = {}
        items = []
        for company in companies:
            for year, file_id in (PLANILHAS.get(company) or {}).items():
                if not file_id:
                    continue
                response = (
                    drive.files()
                    .get(fileId=str(file_id), fields="id,name,modifiedTime,version")
                    .execute()
                )
                item = {
                    "empresa": company,
                    "ano": str(year or "").strip(),
                    "id": str(response.get("id") or file_id).strip(),
                    "name": str(response.get("name") or "").strip(),
                    "modifiedTime": str(response.get("modifiedTime") or "").strip(),
                    "version": str(response.get("version") or "").strip(),
                }
                items.append(item)
                cache_updates[str(file_id)] = {"at": now, "item": item}
        with _AUDIT_META_CACHE_LOCK:
            _AUDIT_META_CACHE.update(cache_updates)
        signature_src = json.dumps(sorted(items, key=lambda item: (item.get("empresa"), item.get("ano"), item.get("id"))), ensure_ascii=True, sort_keys=True)
        return {
            "loaded_at": datetime.now().isoformat(),
            "items": items,
            "signature": hashlib.sha1(signature_src.encode("utf-8")).hexdigest(),
            "mode": "drive",
        }
    except Exception as exc:
        logger.warning("Metadado do Drive indisponível para a conferência. Usando snapshot local curto: %s", exc)
        fallback = {
            "loaded_at": datetime.now().isoformat(),
            "items": [],
            "signature": f"fallback:{','.join(companies)}:{int(now // _AUDIT_META_CACHE_TTL)}",
            "mode": "fallback",
        }
        with _AUDIT_META_CACHE_LOCK:
            _AUDIT_META_CACHE[fallback_key] = {"at": now, "payload": dict(fallback)}
        return fallback


def _audit_gc_expired_jobs():
    cutoff_jobs = time.time() - _AUDIT_JOB_STATE_TTL
    cutoff_snapshots = time.time() - _AUDIT_JOB_SNAPSHOT_TTL
    with _AUDIT_JOB_LOCK:
        for job_id, state in list(_AUDIT_JOBS.items()):
            updated_at = float(state.get("updated_at_ts", 0.0) or 0.0)
            if updated_at and updated_at < cutoff_jobs:
                _AUDIT_JOBS.pop(job_id, None)
        for request_key, job_id in list(_AUDIT_JOBS_BY_REQUEST.items()):
            if job_id not in _AUDIT_JOBS:
                _AUDIT_JOBS_BY_REQUEST.pop(request_key, None)
        for request_key, snapshot in list(_AUDIT_JOB_SNAPSHOTS.items()):
            updated_at = float(snapshot.get("updated_at_ts", 0.0) or 0.0)
            if updated_at and updated_at < cutoff_snapshots:
                _AUDIT_JOB_SNAPSHOTS.pop(request_key, None)


def _audit_job_public_state(state: dict | None) -> dict:
    data = dict(state or {})
    return {
        "job_id": str(data.get("job_id") or "").strip(),
        "status": str(data.get("status") or "idle").strip(),
        "message": str(data.get("message") or "").strip(),
        "started_at": str(data.get("started_at") or "").strip(),
        "updated_at": str(data.get("updated_at") or "").strip(),
        "finished_at": str(data.get("finished_at") or "").strip(),
        "request_signature": str(data.get("request_signature") or "").strip(),
        "sheet_signature": str(data.get("sheet_signature") or "").strip(),
        "result": dict(data.get("result") or {}) if data.get("result") else None,
        "current_result": dict(data.get("current_result") or {}) if data.get("current_result") else None,
    }

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


def _load_audit_sheet_rows(force_refresh: bool = False, empresa_filter: str = "todos") -> tuple[list[dict], dict]:
    empresa = _normalize_audit_empresa(empresa_filter)
    if empresa == "todos":
        rows = []
        metas = []
        for company in ("MVA", "EH"):
            if not PLANILHAS.get(company):
                continue
            company_rows, company_meta = _load_audit_sheet_rows(force_refresh=force_refresh, empresa_filter=company)
            rows.extend(company_rows)
            if company_meta:
                metas.append(company_meta)
        meta = {
            "loaded_at": datetime.now().isoformat(),
            "planilhas_lidas": sum(int(meta.get("planilhas_lidas") or 0) for meta in metas),
            "abas_lidas": sum(int(meta.get("abas_lidas") or 0) for meta in metas),
            "linhas_lidas": len(rows),
            "source": "planilhas",
            "empresa": "todos",
            "empresas_lidas": [str(meta.get("empresa") or "").strip() for meta in metas if str(meta.get("empresa") or "").strip()],
        }
        return rows, meta

    if not force_refresh:
        with _AUDIT_SHEET_CACHE_LOCK:
            cache_entry = dict(_AUDIT_SHEET_CACHE.get(empresa) or {})
            cache_at = float(cache_entry.get("at", 0.0) or 0.0)
            if cache_at and (time.time() - cache_at) <= _AUDIT_SHEET_CACHE_TTL:
                cached_rows = [dict(item or {}) for item in list(cache_entry.get("rows") or [])]
                cached_meta = dict(cache_entry.get("meta") or {})
                if cached_meta:
                    return cached_rows, cached_meta

    creds = Credentials.from_service_account_file(
        GOOGLE_CREDENTIALS_SHEETS,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    gc = gspread.authorize(creds)
    rows = []
    planilhas_lidas = 0
    abas_lidas = 0

    for ano, planilha_id in (PLANILHAS.get(empresa) or {}).items():
        if not planilha_id:
            continue
        try:
            planilha = gc.open_by_key(planilha_id)
            planilhas_lidas += 1
        except Exception as exc:
            logger.warning("Falha ao abrir planilha %s %s para conferencia: %s", empresa, ano, exc)
            continue
        for worksheet in planilha.worksheets():
            abas_lidas += 1
            try:
                linhas = _audit_sheet_values(worksheet)
            except Exception as exc:
                logger.warning("Falha ao ler aba %s/%s: %s", getattr(planilha, "title", empresa), worksheet.title, exc)
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
                        "group_key": f"{empresa}:{ano}:{nf_digits}",
                        "planilha_id": str(planilha_id or "").strip(),
                        "sheet_type": str(empresa or "").strip(),
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
        "empresa": empresa,
        "empresas_lidas": [empresa],
    }
    with _AUDIT_SHEET_CACHE_LOCK:
        _AUDIT_SHEET_CACHE[empresa] = {
            "at": time.time(),
            "rows": [dict(item or {}) for item in rows],
            "meta": dict(meta),
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


def _audit_row_can_delete_duplicate_extra(item: dict) -> bool:
    row_number = max(0, _audit_safe_int((item or {}).get("row_number"), 0))
    return row_number > 1


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
            if _audit_row_is_safe_delete(extra) or _audit_row_can_delete_duplicate_extra(extra):
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


def _audit_delete_candidate_ref(item: dict) -> dict:
    planilha_id, worksheet_title, row_number = _audit_row_ref(item)
    return {
        "planilha_id": str(planilha_id or "").strip(),
        "worksheet_title": str(worksheet_title or "").strip(),
        "row_number": max(0, _audit_safe_int(row_number, 0)),
        "nf": str((item or {}).get("nf") or "").strip(),
        "parcela": str((item or {}).get("parcela") or "").strip(),
    }


def _audit_delete_cache_replace(entries: dict):
    payload = {}
    for raw_key, refs in (entries or {}).items():
        cache_key = str(raw_key or "").strip()
        if not cache_key:
            continue
        clean_refs = []
        seen = set()
        for ref in list(refs or []):
            if not isinstance(ref, dict):
                continue
            planilha_id = str(ref.get("planilha_id") or "").strip()
            worksheet_title = str(ref.get("worksheet_title") or "").strip()
            row_number = max(0, _audit_safe_int(ref.get("row_number"), 0))
            ref_key = (planilha_id, worksheet_title, row_number)
            if not planilha_id or not worksheet_title or row_number <= 1 or ref_key in seen:
                continue
            seen.add(ref_key)
            clean_refs.append(
                {
                    "planilha_id": planilha_id,
                    "worksheet_title": worksheet_title,
                    "row_number": row_number,
                    "nf": str(ref.get("nf") or "").strip(),
                    "parcela": str(ref.get("parcela") or "").strip(),
                }
            )
        payload[cache_key] = clean_refs
    with _AUDIT_DELETE_CACHE_LOCK:
        current = dict(_AUDIT_DELETE_CACHE.get("items") or {})
        current.update(payload)
        _AUDIT_DELETE_CACHE["at"] = time.time()
        _AUDIT_DELETE_CACHE["items"] = current


def _audit_delete_cache_get(audit_keys: list[str]) -> tuple[list[dict], list[str], bool]:
    with _AUDIT_DELETE_CACHE_LOCK:
        cache_at = float(_AUDIT_DELETE_CACHE.get("at", 0.0) or 0.0)
        cache_items = dict(_AUDIT_DELETE_CACHE.get("items") or {})
    expired = bool(cache_at and (time.time() - cache_at > 30 * 60))
    refs = []
    missing = []
    seen = set()
    for raw_key in list(audit_keys or []):
        key = str(raw_key or "").strip()
        if not key:
            continue
        items = list(cache_items.get(key) or [])
        if not items:
            missing.append(key)
            continue
        for ref in items:
            ref_key = (
                str(ref.get("planilha_id") or "").strip(),
                str(ref.get("worksheet_title") or "").strip(),
                max(0, _audit_safe_int(ref.get("row_number"), 0)),
            )
            if ref_key in seen:
                continue
            seen.add(ref_key)
            refs.append(dict(ref))
    return refs, missing, expired


def _delete_audit_rows(audit_keys) -> dict:
    keys = []
    raw_values = [audit_keys] if isinstance(audit_keys, str) else list(audit_keys or [])
    for raw_key in raw_values:
        key = str(raw_key or "").strip()
        if key and key not in keys:
            keys.append(key)
    if not keys:
        return {"ok": False, "message": "NF da conferência não informada."}

    refs, missing, expired = _audit_delete_cache_get(keys)
    if expired and not refs:
        return {
            "ok": False,
            "message": "A conferência carregada expirou. Clique em Conferir parcelas novamente antes de limpar novas linhas.",
            "deleted": 0,
        }
    if not refs:
        return {
            "ok": False,
            "message": "As linhas removíveis não estão mais no snapshot atual da Conferência. Recarregue a aba e tente novamente.",
            "deleted": 0,
            "missing": missing,
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
    ranges_by_sheet = {}

    for item in refs:
        planilha_id = str(item.get("planilha_id") or "").strip()
        worksheet_title = str(item.get("worksheet_title") or "").strip()
        row_number = max(0, _audit_safe_int(item.get("row_number"), 0))
        if not planilha_id or not worksheet_title or row_number <= 1:
            continue
        if planilha_id not in book_cache:
            book_cache[planilha_id] = gc.open_by_key(planilha_id)
        cache_key = (planilha_id, worksheet_title)
        if cache_key not in sheet_cache:
            sheet_cache[cache_key] = book_cache[planilha_id].worksheet(worksheet_title)
        ranges_by_sheet.setdefault(cache_key, []).append((row_number, item))

    for cache_key, entries in ranges_by_sheet.items():
        worksheet = sheet_cache[cache_key]
        clear_ranges = [f"A{row_number}:I{row_number}" for row_number, _ in sorted(entries, key=lambda pair: pair[0])]
        for _ in range(3):
            try:
                worksheet.batch_clear(clear_ranges)
                for row_number, item in entries:
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

    nfs = sorted({str(item.get("nf") or "").strip() for item in refs if str(item.get("nf") or "").strip()})
    nf_msg = ", ".join(nfs[:4]) + ("..." if len(nfs) > 4 else "")
    if deleted:
        with _AUDIT_DELETE_CACHE_LOCK:
            cache_items = dict(_AUDIT_DELETE_CACHE.get("items") or {})
            for key in keys:
                cache_items.pop(key, None)
            _AUDIT_DELETE_CACHE["items"] = cache_items
    return {
        "ok": bool(deleted),
        "message": (
            f"{len(deleted)} linha(s) limpa(s) na planilha para {len(keys)} NF(s): {nf_msg}."
            if deleted
            else "Nenhuma linha foi limpa na planilha."
        ),
        "deleted": len(deleted),
        "items": deleted,
        "missing": missing,
    }


def _gerar_conferencia_parcelas_from_rows(
    filtro: str,
    mes: str,
    nf_inicio: str,
    nf_fim: str,
    linhas: list[dict],
    meta: dict | None = None,
    empresa: str = "todos",
    include_missing_diagnostics: bool = True,
) -> dict:
    filtro_normalizado = str(filtro or "mes").strip().lower()
    if filtro_normalizado not in {"mes", "nfs", "todos"}:
        filtro_normalizado = "mes"
    empresa_normalizada = _normalize_audit_empresa(empresa)
    meta = dict(meta or {})
    grupos = {}
    for item in linhas:
        group_key = str(item.get("group_key") or "").strip()
        if not group_key:
            continue
        grupos.setdefault(group_key, []).append(item)
    nfs_existentes_em_qualquer_mes = {
        _audit_safe_int((item or {}).get("nf_num"), 0)
        for item in list(linhas or [])
        if _audit_safe_int((item or {}).get("nf_num"), 0) > 0
    }

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
    audit_delete_cache = {}
    missing_reason_cache = {}
    missing_reason_state = {}
    resumo = {
        "nfs_verificadas": 0,
        "nfs_ok": 0,
        "nfs_com_divergencia": 0,
        "parcelas_esperadas": 0,
        "parcelas_lancadas": 0,
        "parcelas_duplicadas": 0,
    }
    nfs_vistas_no_escopo = set()
    diagnose_missing_nfs = include_missing_diagnostics and (
        filtro_normalizado == "nfs"
        and nf_inicio_num is not None
        and nf_fim_num is not None
        and (nf_fim_num - nf_inicio_num) <= 40
    )

    for _, nf_items in grupos.items():
        if not _matches_scope(nf_items):
            continue

        nf = str((nf_items[0] or {}).get("nf") or "").strip()
        nf_num = _audit_safe_int((nf_items[0] or {}).get("nf_num"), 0)
        if nf_num > 0:
            nfs_vistas_no_escopo.add(nf_num)
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
        audit_key = str((nf_items[0] or {}).get("group_key") or "").strip()
        audit_delete_cache[audit_key] = [_audit_delete_candidate_ref(item) for item in delete_candidates]
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
                "audit_key": audit_key,
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

    if (
        filtro_normalizado == "nfs"
        and nf_inicio_num is not None
        and nf_fim_num is not None
        and nf_fim_num >= nf_inicio_num
    ):
        for nf_num in range(nf_inicio_num, nf_fim_num + 1):
            if nf_num in nfs_vistas_no_escopo:
                continue
            reason_hint = ""
            if diagnose_missing_nfs:
                reason_hint = missing_reason_cache.get(nf_num)
                if reason_hint is None:
                    reason_hint = _audit_missing_nf_reason(nf_num, state=missing_reason_state)
                    missing_reason_cache[nf_num] = reason_hint
            itens_saida.append(
                {
                    "audit_key": f"missing:{nf_num}",
                    "nf": str(nf_num),
                    "cliente": "-",
                    "descricao": "-",
                    "qtd_esperada": 0,
                    "qtd_lancada": 0,
                    "qtd_bruta": 0,
                    "qtd_faltando": 1,
                    "qtd_excedente": 0,
                    "qtd_duplicada": 0,
                    "parcelas_duplicadas": [],
                    "vencimentos": [],
                    "ultimo_vencimento": "-",
                    "aba": "-",
                    "local_lancamento": "-",
                    "status": "erro",
                    "status_label": "NF ausente",
                    "reason_hint": reason_hint,
                    "delete_candidates": 0,
                    "can_delete_rows": False,
                }
            )
            resumo["nfs_verificadas"] += 1
            resumo["nfs_com_divergencia"] += 1

    if include_missing_diagnostics and filtro_normalizado == "mes" and str(mes or "").strip():
        for missing_item in _audit_missing_nf_candidates_for_month(
            mes,
            month_rows=linhas,
            existing_nfs=nfs_existentes_em_qualquer_mes,
        ):
            nf_num = _audit_safe_int(missing_item.get("nf"), 0)
            if nf_num <= 0 or nf_num in nfs_vistas_no_escopo:
                continue
            nfs_vistas_no_escopo.add(nf_num)
            itens_saida.append(
                {
                    "audit_key": f"gmail-missing:{nf_num}",
                    "nf": str(nf_num),
                    "cliente": "-",
                    "descricao": "-",
                    "qtd_esperada": 0,
                    "qtd_lancada": 0,
                    "qtd_bruta": 0,
                    "qtd_faltando": 1,
                    "qtd_excedente": 0,
                    "qtd_duplicada": 0,
                    "parcelas_duplicadas": [],
                    "vencimentos": [],
                    "ultimo_vencimento": "-",
                    "aba": "-",
                    "local_lancamento": "-",
                    "status": "erro",
                    "status_label": "NF ausente",
                    "reason_hint": str(missing_item.get("reason_hint") or "").strip(),
                    "delete_candidates": 0,
                    "can_delete_rows": False,
                }
            )
            resumo["nfs_verificadas"] += 1
            resumo["nfs_com_divergencia"] += 1

    itens_saida.sort(
        key=lambda item: (
            0 if item.get("status") == "erro" else 1 if item.get("status") == "aviso" else 2,
            -int(re.sub(r"\D+", "", str(item.get("nf") or "0")) or 0),
        )
    )
    _audit_delete_cache_replace(audit_delete_cache)

    return {
        "filtro": filtro_normalizado,
        "mes": mes,
        "nf_inicio": nf_inicio_num,
        "nf_fim": nf_fim_num,
        "empresa": empresa_normalizada,
        "summary": resumo,
        "meta": meta,
        "items": itens_saida,
    }


def _gerar_conferencia_parcelas(filtro: str, mes: str, nf_inicio: str, nf_fim: str, empresa: str = "todos") -> dict:
    empresa_normalizada = _normalize_audit_empresa(empresa)
    linhas, meta = _load_audit_sheet_rows(empresa_filter=empresa_normalizada)
    return _gerar_conferencia_parcelas_from_rows(
        filtro,
        mes,
        nf_inicio,
        nf_fim,
        linhas,
        meta=meta,
        empresa=empresa_normalizada,
        include_missing_diagnostics=True,
    )


def _audit_update_job_state(job_id: str, **updates):
    job_key = str(job_id or "").strip()
    if not job_key:
        return
    with _AUDIT_JOB_LOCK:
        state = dict(_AUDIT_JOBS.get(job_key) or {})
        if not state:
            return
        state.update(updates)
        state["updated_at"] = datetime.now().isoformat()
        state["updated_at_ts"] = time.time()
        _AUDIT_JOBS[job_key] = state


def _audit_job_worker(job_id: str):
    with _AUDIT_JOB_LOCK:
        state = dict(_AUDIT_JOBS.get(job_id) or {})
    if not state:
        return
    filtro = str(state.get("filtro") or "mes").strip()
    mes = str(state.get("mes") or "").strip()
    nf_inicio = str(state.get("nf_inicio") or "").strip()
    nf_fim = str(state.get("nf_fim") or "").strip()
    empresa = _normalize_audit_empresa(state.get("empresa"))
    request_signature = str(state.get("request_signature") or "").strip()
    sheet_signature = str(state.get("sheet_signature") or "").strip()
    companies = _audit_companies_for_filter(empresa)
    rows_by_company = {}
    metas_by_company = {}
    try:
        for idx, company in enumerate(companies):
            _audit_update_job_state(job_id, status="running", message=f"Lendo planilhas {company}...")
            company_rows, company_meta = _load_audit_sheet_rows(force_refresh=True, empresa_filter=company)
            rows_by_company[company] = list(company_rows or [])
            metas_by_company[company] = dict(company_meta or {})
            partial_companies = companies[: idx + 1]
            partial_rows = []
            partial_metas = []
            for partial_company in partial_companies:
                partial_rows.extend(list(rows_by_company.get(partial_company) or []))
                if metas_by_company.get(partial_company):
                    partial_metas.append(dict(metas_by_company.get(partial_company) or {}))
            partial_result = _gerar_conferencia_parcelas_from_rows(
                filtro,
                mes,
                nf_inicio,
                nf_fim,
                partial_rows,
                meta=_audit_merge_meta_parts(partial_metas, empresa_filter=empresa if empresa != "todos" else "todos"),
                empresa=empresa if empresa != "todos" else "todos",
                include_missing_diagnostics=False,
            )
            stage_label = (
                f"{company} pronta"
                if empresa != "todos"
                else (f"{company} pronta. {companies[idx + 1]} carregando em segundo plano..." if idx + 1 < len(companies) else "Planilhas carregadas. Finalizando diagnósticos...")
            )
            _audit_update_job_state(
                job_id,
                current_result=partial_result,
                message=stage_label,
            )

        final_rows = []
        final_metas = []
        for company in companies:
            final_rows.extend(list(rows_by_company.get(company) or []))
            if metas_by_company.get(company):
                final_metas.append(dict(metas_by_company.get(company) or {}))
        _audit_update_job_state(job_id, status="running", message="Finalizando diagnósticos da conferência...")
        final_result = _gerar_conferencia_parcelas_from_rows(
            filtro,
            mes,
            nf_inicio,
            nf_fim,
            final_rows,
            meta=_audit_merge_meta_parts(final_metas, empresa_filter=empresa if empresa != "todos" else "todos"),
            empresa=empresa if empresa != "todos" else "todos",
            include_missing_diagnostics=True,
        )
        finished_at = datetime.now().isoformat()
        snapshot = {
            "request_signature": request_signature,
            "sheet_signature": sheet_signature,
            "result": final_result,
            "updated_at": finished_at,
            "updated_at_ts": time.time(),
        }
        with _AUDIT_JOB_LOCK:
            _AUDIT_JOB_SNAPSHOTS[request_signature] = snapshot
            state = dict(_AUDIT_JOBS.get(job_id) or {})
            state.update(
                {
                    "status": "done",
                    "message": "Conferência pronta.",
                    "result": final_result,
                    "current_result": final_result,
                    "finished_at": finished_at,
                    "updated_at": finished_at,
                    "updated_at_ts": time.time(),
                }
            )
            _AUDIT_JOBS[job_id] = state
    except Exception as exc:
        logger.exception("Falha ao montar job da conferência: %s", exc)
        _audit_update_job_state(
            job_id,
            status="error",
            message=str(exc),
            finished_at=datetime.now().isoformat(),
        )


def _start_audit_job(filtro: str, mes: str, nf_inicio: str, nf_fim: str, empresa: str = "todos") -> dict:
    _audit_gc_expired_jobs()
    empresa_normalizada = _normalize_audit_empresa(empresa)
    request_signature = _audit_request_signature(filtro, mes, nf_inicio, nf_fim, empresa_normalizada)
    metadata = _audit_drive_metadata(empresa_normalizada)
    sheet_signature = str(metadata.get("signature") or "").strip()
    metadata_mode = str(metadata.get("mode") or "").strip()
    with _AUDIT_JOB_LOCK:
        snapshot = dict(_AUDIT_JOB_SNAPSHOTS.get(request_signature) or {})
        if snapshot and str(snapshot.get("sheet_signature") or "").strip() == sheet_signature and snapshot.get("result"):
            return {
                "job_id": "",
                "status": "done",
                "message": "Conferência pronta a partir do snapshot." if metadata_mode == "drive" else "Conferência pronta a partir do snapshot local recente.",
                "result": dict(snapshot.get("result") or {}),
                "current_result": dict(snapshot.get("result") or {}),
                "request_signature": request_signature,
                "sheet_signature": sheet_signature,
                "cached": True,
            }

        existing_job_id = str(_AUDIT_JOBS_BY_REQUEST.get(request_signature) or "").strip()
        existing_state = dict(_AUDIT_JOBS.get(existing_job_id) or {}) if existing_job_id else {}
        if existing_state and str(existing_state.get("sheet_signature") or "").strip() == sheet_signature and str(existing_state.get("status") or "").strip() in {"running", "done"}:
            return _audit_job_public_state(existing_state)

        job_id = secrets.token_hex(8)
        now_iso = datetime.now().isoformat()
        state = {
            "job_id": job_id,
            "status": "running",
            "message": "Iniciando conferência..." if metadata_mode == "drive" else "Iniciando conferência com snapshot local curto...",
            "started_at": now_iso,
            "updated_at": now_iso,
            "updated_at_ts": time.time(),
            "finished_at": "",
            "filtro": str(filtro or "mes").strip(),
            "mes": str(mes or "").strip(),
            "nf_inicio": str(nf_inicio or "").strip(),
            "nf_fim": str(nf_fim or "").strip(),
            "empresa": empresa_normalizada,
            "request_signature": request_signature,
            "sheet_signature": sheet_signature,
            "metadata": dict(metadata or {}),
            "result": None,
            "current_result": None,
        }
        _AUDIT_JOBS[job_id] = state
        _AUDIT_JOBS_BY_REQUEST[request_signature] = job_id
    thread = threading.Thread(target=_audit_job_worker, args=(job_id,), daemon=True)
    thread.start()
    return _audit_job_public_state(state)


def _get_audit_job(job_id: str) -> dict:
    _audit_gc_expired_jobs()
    key = str(job_id or "").strip()
    if not key:
        raise ValueError("Job da conferência não informado.")
    with _AUDIT_JOB_LOCK:
        state = dict(_AUDIT_JOBS.get(key) or {})
    if not state:
        raise ValueError("Job da conferência não encontrado ou expirado.")
    return _audit_job_public_state(state)


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


def _load_watch_search_suggestions(empresa_filter: str = "todos") -> list[str]:
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
        linhas, _ = _load_audit_sheet_rows(empresa_filter=_normalize_audit_empresa(empresa_filter))
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


def _gerar_relacao_pendencias(boletos_dias: int, depositos_dias: int, empresa_filter: str = "todos") -> dict:
    boleto_limit = max(1, min(7, _audit_safe_int(boletos_dias, 7)))
    deposito_limit = max(1, min(7, _audit_safe_int(depositos_dias, 7)))
    empresa = _normalize_audit_empresa(empresa_filter)
    linhas, meta = _load_audit_sheet_rows(empresa_filter=empresa)
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
                "empresa": _normalize_report_text(str(item.get("sheet_type") or "").strip()),
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
        "meta": {**meta, "loaded_at": datetime.now().isoformat(), "empresa": empresa},
        "limits": {"boletos_dias": boleto_limit, "depositos_dias": deposito_limit},
        "items": itens,
    }


def _buscar_boletos_em_aberto_por_nome(nome: str, empresa_filter: str = "todos") -> dict:
    nome_busca = _normalize_report_text(str(nome or "").strip())
    termo = _normalize_ascii_key(nome_busca)
    if not termo:
        raise ValueError("Informe um nome para buscar.")
    empresa = _normalize_audit_empresa(empresa_filter)
    linhas, meta = _load_audit_sheet_rows(empresa_filter=empresa)
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
                "empresa": _normalize_report_text(str(item.get("sheet_type") or "").strip()),
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
        "meta": {**meta, "loaded_at": datetime.now().isoformat(), "empresa": empresa},
        "items": itens,
        "suggestions": suggestions[:500],
    }


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


def _normalize_recovery_mode(value: str) -> str:
    mode = str(value or "").strip().lower()
    if mode in {"period", "periodo", "período", "date", "data"}:
        return "period"
    if mode in {"range", "faixa", "nfs"}:
        return "range"
    if mode in {"list", "lista", "manual", "escolha"}:
        return "list"
    return ""


def _parse_nf_selection_list(value) -> list[int]:
    if isinstance(value, (list, tuple, set)):
        raw_values = [str(item or "").strip() for item in value]
    else:
        raw_values = [str(value or "").strip()]
    itens = []
    vistos = set()
    for raw in raw_values:
        for token in re.findall(r"\d{1,5}", raw):
            try:
                numero = int(token)
            except Exception:
                continue
            if numero <= 0 or numero in vistos:
                continue
            vistos.add(numero)
            itens.append(numero)
    return itens


def _resolve_recovery_mode(
    mode: str = "",
    nf_start: str = "",
    nf_end: str = "",
    date_from: str = "",
    date_to: str = "",
    nf_list=None,
) -> str:
    mode_norm = _normalize_recovery_mode(mode)
    if mode_norm:
        return mode_norm
    if _parse_nf_selection_list(nf_list):
        return "list"
    start_nf, end_nf = _parse_nf_filter_range(nf_start, nf_end)
    if start_nf is not None or end_nf is not None:
        return "range"
    if _parse_iso_date_input(date_from) or _parse_iso_date_input(date_to):
        return "period"
    return ""


def _recover_nf_query_terms(nf_start: str = "", nf_end: str = "", nf_list=None) -> str:
    numeros = _parse_nf_selection_list(nf_list)
    if numeros:
        if len(numeros) == 1:
            return str(numeros[0])
        if len(numeros) <= 25:
            return "{" + " ".join(str(numero) for numero in numeros) + "}"
        return ""
    start_nf, end_nf = _parse_nf_filter_range(nf_start, nf_end)
    if start_nf is None or end_nf is None:
        return ""
    if start_nf == end_nf:
        return str(start_nf)
    if end_nf - start_nf <= 20:
        return "{" + " ".join(str(numero) for numero in range(start_nf, end_nf + 1)) + "}"
    return ""


def _recover_search_query(
    mode: str = "",
    nf_start: str = "",
    nf_end: str = "",
    date_from: str = "",
    date_to: str = "",
    nf_list=None,
) -> str:
    mode_norm = _resolve_recovery_mode(mode=mode, nf_start=nf_start, nf_end=nf_end, date_from=date_from, date_to=date_to, nf_list=nf_list)
    extra_parts = []
    if mode_norm == "period":
        start_date = _parse_iso_date_input(date_from)
        end_date = _parse_iso_date_input(date_to)
        if start_date:
            extra_parts.append(f"after:{start_date.strftime('%Y/%m/%d')}")
        if end_date:
            extra_parts.append(f"before:{(end_date + timedelta(days=1)).strftime('%Y/%m/%d')}")
    else:
        nf_query = _recover_nf_query_terms(nf_start=nf_start, nf_end=nf_end, nf_list=nf_list)
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


def _extract_nf_numbers_from_xml_filename(filename: str) -> list[int]:
    text = str(filename or "").strip()
    upper_name = text.upper()
    if ".XML" not in upper_name:
        return []
    values = set(_extract_nf_numbers_from_text(filename))
    for chave in re.findall(r"(\d{44})", text):
        try:
            numero = int(chave[25:34])
        except Exception:
            continue
        if numero > 0:
            values.add(numero)
    return sorted(values)


def _extract_xml_hint_numbers(text: str) -> list[int]:
    values = []
    for raw in re.findall(r"\bXMLNF\s*0*(\d{3,})\b", str(text or "").upper()):
        try:
            values.append(int(raw))
        except Exception:
            continue
    return sorted(set(values))


def _format_nf_number_list(values, limit: int = 8) -> str:
    numeros = []
    vistos = set()
    for item in list(values or []):
        try:
            numero = int(item)
        except Exception:
            continue
        if numero <= 0 or numero in vistos:
            continue
        vistos.add(numero)
        numeros.append(numero)
    if not numeros:
        return ""
    if len(numeros) <= max(1, int(limit or 8)):
        return ", ".join(str(numero) for numero in numeros)
    head = ", ".join(str(numero) for numero in numeros[: max(1, int(limit or 8))])
    return f"{head} e mais {len(numeros) - max(1, int(limit or 8))}"


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


def _describe_recovery_filters(
    mode: str = "",
    nf_start: str = "",
    nf_end: str = "",
    date_from: str = "",
    date_to: str = "",
    nf_list=None,
) -> str:
    mode_norm = _resolve_recovery_mode(mode=mode, nf_start=nf_start, nf_end=nf_end, date_from=date_from, date_to=date_to, nf_list=nf_list)
    parts = []
    if mode_norm == "list":
        numeros = _parse_nf_selection_list(nf_list)
        if numeros:
            parts.append("NFs " + ", ".join(str(numero) for numero in numeros))
    elif mode_norm == "range":
        start_nf, end_nf = _parse_nf_filter_range(nf_start, nf_end)
        if start_nf is not None and end_nf is not None:
            parts.append(f"NFs {start_nf} a {end_nf}" if start_nf != end_nf else f"NF {start_nf}")
    elif mode_norm == "period":
        start_date = _parse_iso_date_input(date_from)
        end_date = _parse_iso_date_input(date_to)
        if start_date and end_date:
            parts.append(f"período {start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')}")
        elif start_date:
            parts.append(f"a partir de {start_date.strftime('%d/%m/%Y')}")
        elif end_date:
            parts.append(f"até {end_date.strftime('%d/%m/%Y')}")
    return " | ".join(parts)


def _preview_from_message_headers(
    from_raw: str = "",
    subject: str = "",
    date_raw: str = "",
    snippet: str = "",
    timestamp: int = 0,
) -> dict:
    _, email_addr = parseaddr(str(from_raw or "").strip())
    email_view = email_addr or str(from_raw or "").strip()
    date_view = str(date_raw or "").strip()
    stamp = int(timestamp or 0)
    parsed_date = None
    if date_raw:
        try:
            dt = parsedate_to_datetime(str(date_raw))
            if dt.tzinfo is not None:
                dt = dt.astimezone()
            date_view = dt.strftime("%d/%m/%Y %H:%M")
            parsed_date = dt.date()
            if not stamp:
                stamp = int(dt.timestamp() * 1000)
        except Exception:
            pass
    return {
        "email": email_view,
        "subject": str(subject or "").strip(),
        "snippet": str(snippet or "").strip(),
        "date": date_view,
        "date_raw": str(date_raw or "").strip(),
        "timestamp": stamp,
        "_parsed_date": parsed_date,
    }


def _preview_from_message_item(item: dict | None) -> dict:
    payload = dict(item or {})
    return _preview_from_message_headers(
        from_raw=str(payload.get("from", "") or "").strip(),
        subject=str(payload.get("subject", "") or "").strip(),
        date_raw=str(payload.get("date", "") or "").strip(),
        snippet=str(payload.get("snippet", "") or "").strip(),
    )


def _preview_matches_recovery_filters(
    preview: dict,
    mode: str = "",
    nf_start: int | None = None,
    nf_end: int | None = None,
    nf_list=None,
    start_date=None,
    end_date=None,
) -> bool:
    mode_norm = _normalize_recovery_mode(mode)
    if mode_norm == "period":
        parsed_date = preview.get("_parsed_date")
        if parsed_date is None:
            return False
        if start_date and parsed_date < start_date:
            return False
        if end_date and parsed_date > end_date:
            return False
        return True
    texto = " ".join(
        x for x in (
            str(preview.get("subject", "") or "").strip(),
            str(preview.get("snippet", "") or "").strip(),
        ) if x
    )
    numeros = _extract_nf_numbers_from_text(texto)
    if not numeros:
        return False
    if mode_norm == "range":
        if nf_start is None or nf_end is None:
            return False
        return any(nf_start <= numero <= nf_end for numero in numeros)
    if mode_norm == "list":
        wanted = set(_parse_nf_selection_list(nf_list))
        return any(numero in wanted for numero in numeros)
    return False


def _message_attachment_text(service, msg_id: str) -> str:
    if not msg_id:
        return ""
    try:
        payload = (
            service.users()
            .messages()
            .get(userId="me", id=msg_id, format="full")
            .execute()
            .get("payload", {})
            or {}
        )
    except Exception as exc:
        logger.warning("Falha ao carregar anexos da mensagem %s para filtro de NF: %s", msg_id, exc)
        return ""
    parts = list(payload.get("parts", []) or [])
    if not parts:
        parts = [payload]
    names = []
    stack = list(parts)
    while stack:
        part = stack.pop()
        children = list(part.get("parts", []) or [])
        if children:
            stack.extend(children)
        filename = str(part.get("filename", "") or "").strip()
        if filename:
            names.append(filename)
            for numero in _extract_nf_numbers_from_xml_filename(filename):
                names.append(f"XMLNF {numero}")
    return " ".join(names).strip()


def _recovery_match_details(
    service,
    msg_id: str,
    preview: dict,
    mode: str = "",
    nf_start: int | None = None,
    nf_end: int | None = None,
    nf_list=None,
    start_date=None,
    end_date=None,
    attachment_text: str = "",
) -> dict:
    def _warning_from_numbers(preview_numbers, attachment_numbers) -> str:
        if not (preview_numbers or attachment_numbers):
            return ""
        preview_label = ", ".join(str(numero) for numero in preview_numbers) if preview_numbers else "nenhuma NF"
        attachment_label = ", ".join(str(numero) for numero in attachment_numbers) if attachment_numbers else "nenhuma NF"
        if preview_label == attachment_label:
            return ""
        return (
            f"Assunto/snippet indicavam {preview_label}, "
            f"mas os anexos/XML indicavam {attachment_label}."
        )

    mode_norm = _normalize_recovery_mode(mode)
    preview_match = _preview_matches_recovery_filters(
        preview,
        mode=mode,
        nf_start=nf_start,
        nf_end=nf_end,
        nf_list=nf_list,
        start_date=start_date,
        end_date=end_date,
    )
    if mode_norm == "period":
        return {
            "matched": bool(preview_match),
            "used_attachment_match": False,
            "attachment_text": str(attachment_text or "").strip(),
            "warning_note": "",
        }
    text = str(attachment_text or "").strip()
    if not text:
        text = _message_attachment_text(service, msg_id)
    if not text:
        if preview_match:
            return {
                "matched": True,
                "used_attachment_match": False,
                "attachment_text": "",
                "warning_note": "",
            }
        return {
            "matched": False,
            "used_attachment_match": False,
            "attachment_text": "",
            "warning_note": "",
        }
    preview_text = " ".join(
        x for x in (
            str(preview.get("subject", "") or "").strip(),
            str(preview.get("snippet", "") or "").strip(),
        ) if x
    )
    preview_numbers = _extract_nf_numbers_from_text(preview_text)
    xml_attachment_numbers = _extract_xml_hint_numbers(text)
    xml_match = False
    if xml_attachment_numbers:
        if mode_norm == "range":
            xml_match = bool(
                nf_start is not None
                and nf_end is not None
                and any(nf_start <= numero <= nf_end for numero in xml_attachment_numbers)
            )
        elif mode_norm == "list":
            wanted = set(_parse_nf_selection_list(nf_list))
            xml_match = any(numero in wanted for numero in xml_attachment_numbers)
    attachment_numbers = _extract_nf_numbers_from_text(text)
    merged_preview = dict(preview or {})
    merged_preview["subject"] = " ".join(
        x for x in (
            str(preview.get("subject", "") or "").strip(),
            text,
        ) if x
    )
    attachment_match = _preview_matches_recovery_filters(
        merged_preview,
        mode=mode,
        nf_start=nf_start,
        nf_end=nf_end,
        nf_list=nf_list,
        start_date=start_date,
        end_date=end_date,
    )
    warning_note = _warning_from_numbers(preview_numbers, xml_attachment_numbers or attachment_numbers)
    if preview_match and xml_attachment_numbers and not xml_match:
        return {
            "matched": False,
            "used_attachment_match": False,
            "attachment_text": text,
            "warning_note": warning_note,
        }
    if preview_match and attachment_numbers and not attachment_match:
        return {
            "matched": False,
            "used_attachment_match": False,
            "attachment_text": text,
            "warning_note": warning_note,
        }
    if not preview_match and xml_attachment_numbers and xml_match:
        return {
            "matched": True,
            "used_attachment_match": True,
            "attachment_text": text,
            "warning_note": warning_note,
        }
    if preview_match:
        return {
            "matched": True,
            "used_attachment_match": False,
            "attachment_text": text,
            "warning_note": warning_note if attachment_match else "",
        }
    return {
        "matched": bool(attachment_match),
        "used_attachment_match": bool(attachment_match),
        "attachment_text": text,
        "warning_note": warning_note if attachment_match else "",
    }


def _audit_missing_nf_reason(nf_num: int, state: dict | None = None) -> str:
    try:
        nf_value = int(nf_num or 0)
    except Exception:
        return ""
    if nf_value <= 0:
        return ""
    ctx = state if isinstance(state, dict) else {}
    if ctx.get("disabled"):
        return ""
    attachment_text_cache = ctx.setdefault("attachment_text_cache", {})
    service = ctx.get("service")
    if service is None:
        try:
            service = _get_gmail_service_locked(timeout=0.5)
        except Exception:
            ctx["disabled"] = True
            return ""
        ctx["service"] = service
    query = _join_gmail_query(build_sent_xml_query(filter_mode="", extra_query=""), str(nf_value))
    page_token = None
    pages = 0
    seen_ids = set()
    while pages < 2 and len(seen_ids) < 12:
        try:
            batch, next_page_token = buscarMessagesEnviadosPagina(
                service,
                max_results=8,
                page_token=page_token,
                query=query,
            )
        except Exception as exc:
            logger.warning("Falha ao diagnosticar NF ausente %s no Gmail: %s", nf_value, exc)
            ctx["disabled"] = True
            return ""
        pages += 1
        if not batch and not next_page_token:
            break
        for item in batch:
            msg_id = str(item.get("id", "")).strip()
            if not msg_id or msg_id in seen_ids:
                continue
            seen_ids.add(msg_id)
            preview = _preview_from_message_item(item)
            attachment_text = str(attachment_text_cache.get(msg_id, "") or "").strip()
            match_info = _recovery_match_details(
                service,
                msg_id,
                preview,
                mode="list",
                nf_list=[nf_value],
                attachment_text=attachment_text,
            )
            attachment_text_cache[msg_id] = str(match_info.get("attachment_text", "") or "").strip()
            warning_note = str(match_info.get("warning_note", "") or "").strip()
            date_view = str(preview.get("date", "") or "").strip()
            subject_view = str(preview.get("subject", "") or "").strip()
            context = " | ".join(part for part in (date_view, subject_view) if part)
            if warning_note:
                base = f"{warning_note}"
                if context:
                    base = f"{base} E-mail: {context}."
                if bool(match_info.get("matched")):
                    return f"{base} A recuperacao consegue localizar esta NF pelos anexos/XML."
                return base
            if bool(match_info.get("matched")):
                if context:
                    return f"E-mail com XML localizado no Gmail para esta NF. {context}."
                return "E-mail com XML localizado no Gmail para esta NF."
        if not next_page_token:
            break
        page_token = next_page_token
    return ""


def _audit_month_gap_candidates(month_rows, month_key: str) -> list[int]:
    if not month_rows or not str(month_key or "").strip():
        return []
    groups = {}
    for item in list(month_rows or []):
        if str((item or {}).get("scope_month") or "").strip() != month_key:
            continue
        nf_num = _audit_safe_int((item or {}).get("nf_num"), 0)
        if nf_num <= 0:
            continue
        group_key = (
            str((item or {}).get("sheet_type") or "").strip(),
            str((item or {}).get("sheet_year") or "").strip(),
        )
        groups.setdefault(group_key, set()).add(nf_num)

    out = []
    seen = set()
    for numbers in groups.values():
        ordered = sorted(int(n) for n in numbers if int(n or 0) > 0)
        for prev_nf, next_nf in reversed(list(zip(ordered, ordered[1:]))):
            gap = int(next_nf) - int(prev_nf) - 1
            if gap <= 0 or gap > _AUDIT_MONTH_GAP_LIMIT:
                continue
            for nf_num in range(int(next_nf) - 1, int(prev_nf), -1):
                if nf_num in seen:
                    continue
                seen.add(nf_num)
                out.append(nf_num)
                if len(out) >= _AUDIT_MONTH_MAX_GAP_CHECKS:
                    return out
    return out


def _audit_missing_nf_candidates_for_month(mes: str, month_rows=None, existing_nfs=None) -> list[dict]:
    month_key = str(mes or "").strip()
    if not re.match(r"^\d{4}-\d{2}$", month_key):
        return []
    seen_nfs = {int(n) for n in (existing_nfs or set()) if int(n or 0) > 0}
    candidates = [nf for nf in _audit_month_gap_candidates(month_rows or [], month_key) if nf not in seen_nfs]
    if not candidates:
        return []

    signature_src = ",".join(str(nf) for nf in candidates)
    signature = hashlib.sha1(signature_src.encode("utf-8")).hexdigest()
    cache_key = f"{month_key}:{signature}"
    now = time.time()
    with _AUDIT_MONTH_MISSING_CACHE_LOCK:
        cache_entry = dict(_AUDIT_MONTH_MISSING_CACHE.get(cache_key) or {})
        if cache_entry and (now - float(cache_entry.get("at", 0.0) or 0.0) <= _AUDIT_MONTH_MISSING_CACHE_TTL):
            return list(cache_entry.get("items") or [])

    state = {}
    out = []
    for nf_num in candidates:
        reason_hint = _audit_missing_nf_reason(nf_num, state=state)
        if not reason_hint:
            continue
        out.append({"nf": nf_num, "reason_hint": reason_hint})

    with _AUDIT_MONTH_MISSING_CACHE_LOCK:
        _AUDIT_MONTH_MISSING_CACHE.clear()
        _AUDIT_MONTH_MISSING_CACHE[cache_key] = {
            "at": now,
            "items": list(out),
        }
    return out


def _message_matches_recovery_filters(
    service,
    msg_id: str,
    preview: dict,
    mode: str = "",
    nf_start: int | None = None,
    nf_end: int | None = None,
    nf_list=None,
    start_date=None,
    end_date=None,
    attachment_text: str = "",
) -> bool:
    return bool(
        _recovery_match_details(
            service,
            msg_id,
            preview,
            mode=mode,
            nf_start=nf_start,
            nf_end=nf_end,
            nf_list=nf_list,
            start_date=start_date,
            end_date=end_date,
            attachment_text=attachment_text,
        ).get("matched")
    )


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
        preview.update(
            _preview_from_message_headers(
                from_raw=from_raw,
                subject=subject,
                date_raw=date_raw,
                snippet="",
                timestamp=int(internal_ts or 0),
            )
        )
    except Exception as exc:
        logger.warning("Falha ao carregar cabecalhos da mensagem %s: %s", msg_id, exc)
    return preview


def _selected_preview_window(items: list[dict]) -> dict:
    selected = [dict(item.get("preview") or {}) for item in list(items or []) if isinstance(item, dict)]
    selected = [item for item in selected if item]
    if not selected:
        return {"selected": 0, "oldest_date": "", "newest_date": ""}
    with_ts = [item for item in selected if int(item.get("timestamp", 0) or 0) > 0]
    if with_ts:
        newest = max(with_ts, key=lambda item: int(item.get("timestamp", 0) or 0))
        oldest = min(with_ts, key=lambda item: int(item.get("timestamp", 0) or 0))
    else:
        newest = selected[0]
        oldest = selected[-1]
    return {
        "selected": len(selected),
        "oldest_date": str(oldest.get("date", "") or "").strip(),
        "newest_date": str(newest.get("date", "") or "").strip(),
    }


def _find_missing_messages_for_nf_list(
    service,
    nf_values: list[int],
    max_messages: int,
    page_size: int,
    progress_cb=None,
) -> dict:
    requested = _parse_nf_selection_list(nf_values)
    requested_total = len(requested)
    per_nf_limit = 3
    message_limit = max(1, min(int(max_messages or 1), requested_total * per_nf_limit))
    page_limit = max(4, min(10, _manual_scan_page_limit(50, page_size)))
    targets = []
    target_ids = set()
    inspected_ids = set()
    inspected = 0
    pages = 0
    found_nf_numbers = []
    missing_nf_numbers = []
    attachment_text_cache = {}
    subject_mismatch_notes = []
    subject_mismatch_count = 0
    criteria_desc = _describe_recovery_filters(mode="list", nf_list=requested) or "NFs selecionadas"
    if callable(progress_cb):
        progress_cb(
            phase="searching",
            progress_current=0,
            progress_total=requested_total,
            matched=0,
            inspected=0,
            current_email="",
            current_subject="",
            current_date="",
            requested_nf_count=requested_total,
            found_nf_numbers=[],
            missing_nf_numbers=requested[:],
            message="Buscando as NFs selecionadas no Gmail.",
            detail=f"Criterios: {criteria_desc}.",
        )
    for idx, nf in enumerate(requested, start=1):
        query = _join_gmail_query(build_sent_xml_query(filter_mode="", extra_query=""), str(nf))
        page_token = None
        local_pages = 0
        local_matches = 0
        found_for_nf = False
        while local_pages < page_limit and local_matches < per_nf_limit and len(targets) < message_limit:
            batch, next_page_token = buscarMessagesEnviadosPagina(
                service,
                max_results=page_size,
                page_token=page_token,
                query=query,
            )
            pages += 1
            local_pages += 1
            if not batch and not next_page_token:
                break
            for item in batch:
                msg_id = str(item.get("id", "")).strip()
                if not msg_id:
                    continue
                preview = _preview_from_message_item(item)
                current_email = str(preview.get("email", "")).strip()
                current_subject = str(preview.get("subject", "")).strip()
                current_date = str(preview.get("date", "")).strip()
                if msg_id not in inspected_ids:
                    inspected_ids.add(msg_id)
                    inspected += 1
                if callable(progress_cb):
                    progress_cb(
                        phase="searching",
                        progress_current=len(found_nf_numbers),
                        progress_total=requested_total,
                        matched=len(targets),
                        inspected=inspected,
                        current_email=current_email,
                        current_subject=current_subject,
                        current_date=current_date,
                        requested_nf_count=requested_total,
                        found_nf_numbers=found_nf_numbers[:],
                        missing_nf_numbers=[numero for numero in requested if numero not in set(found_nf_numbers)],
                        message=f"Buscando NF {nf} ({idx} de {requested_total}).",
                        detail=f"{inspected} mensagem(ns) analisada(s) ate agora.",
                    )
                attachment_text = attachment_text_cache.get(msg_id, "")
                match_info = _recovery_match_details(
                    service,
                    msg_id,
                    preview,
                    mode="list",
                    nf_list=[nf],
                    attachment_text=attachment_text,
                )
                attachment_text_cache[msg_id] = str(match_info.get("attachment_text", "") or "").strip()
                warning_note = str(match_info.get("warning_note", "") or "").strip()
                if warning_note:
                    subject_mismatch_count += 1
                    if len(subject_mismatch_notes) < 8:
                        note_parts = [f"NF {nf}"]
                        if current_date:
                            note_parts.append(current_date)
                        if current_subject:
                            note_parts.append(current_subject)
                        context = " | ".join(note_parts)
                        subject_mismatch_notes.append(
                            f"{context} -> {warning_note}" if context else warning_note
                        )
                if not bool(match_info.get("matched")):
                    continue
                found_for_nf = True
                if nf not in found_nf_numbers:
                    found_nf_numbers.append(nf)
                if msg_id in target_ids:
                    local_matches += 1
                    if local_matches >= per_nf_limit:
                        break
                    continue
                target_ids.add(msg_id)
                local_matches += 1
                targets.append(
                    {
                        "id": msg_id,
                        "threadId": str(item.get("threadId", "")).strip(),
                        "labelIds": list(item.get("labelIds", []) or []),
                        "snippet": str(item.get("snippet", "") or ""),
                        "subject": str(item.get("subject", "") or "").strip(),
                        "date": str(item.get("date", "") or "").strip(),
                        "from": str(item.get("from", "") or "").strip(),
                        "matched_nf": nf,
                    }
                )
                if callable(progress_cb):
                    progress_cb(
                        phase="searching",
                        progress_current=len(found_nf_numbers),
                        progress_total=requested_total,
                        matched=len(targets),
                        inspected=inspected,
                        current_email=current_email,
                        current_subject=current_subject,
                        current_date=current_date,
                        requested_nf_count=requested_total,
                        found_nf_numbers=found_nf_numbers[:],
                        missing_nf_numbers=[numero for numero in requested if numero not in set(found_nf_numbers)],
                        message=f"NFs localizadas: {len(found_nf_numbers)} de {requested_total}.",
                        detail=f"NF {nf} localizada em {local_matches} mensagem(ns).",
                    )
                if local_matches >= per_nf_limit or len(targets) >= message_limit:
                    break
            if local_matches >= per_nf_limit or len(targets) >= message_limit or not next_page_token:
                break
            page_token = next_page_token
        if not found_for_nf:
            missing_nf_numbers.append(nf)
            if callable(progress_cb):
                progress_cb(
                    phase="searching",
                    progress_current=len(found_nf_numbers),
                    progress_total=requested_total,
                    matched=len(targets),
                    inspected=inspected,
                    current_email="",
                    current_subject="",
                    current_date="",
                    requested_nf_count=requested_total,
                    found_nf_numbers=found_nf_numbers[:],
                    missing_nf_numbers=missing_nf_numbers[:],
                    message=f"NFs localizadas: {len(found_nf_numbers)} de {requested_total}.",
                    detail=f"A NF {nf} nao apareceu nas mensagens verificadas.",
                )
    return {
        "ok": True,
        "matched": len(targets),
        "inspected": inspected,
        "pages": pages,
        "query": "consultas individuais por NF",
        "criteria": criteria_desc,
        "mode": "list",
        "targets": targets,
        "subject_mismatch_count": subject_mismatch_count,
        "subject_mismatch_notes": subject_mismatch_notes,
        "requested_nf_count": requested_total,
        "found_nf_numbers": found_nf_numbers,
        "missing_nf_numbers": missing_nf_numbers,
    }


def _reprocess_recent_query() -> str:
    return f"newer_than:{int(_REPROCESS_LOOKBACK_DAYS)}d"


def _reprocess_recent(max_messages: int, mark_unread: bool = False, progress_cb=None, continue_after_id: str = "") -> dict:
    service = _get_gmail_service_locked()
    wanted = max(1, min(1000, int(max_messages)))
    messages_raw = listar_mensagens_com_labels_botana(
        service,
        max_results=1000,
        query=_reprocess_recent_query(),
    )
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
    start_index = 0
    continue_after_id = str(continue_after_id or "").strip()
    if continue_after_id:
        for idx, item in enumerate(mensagens_com_meta):
            if str(item.get("id", "")).strip() == continue_after_id:
                start_index = idx + 1
                break
    messages = mensagens_com_meta[start_index : start_index + wanted]
    window_info = _selected_preview_window(messages)
    selected_oldest_id = str((messages[-1] or {}).get("id", "")).strip() if messages else ""
    remaining_after = max(0, len(mensagens_com_meta) - (start_index + len(messages)))
    changed = 0
    failed = 0
    targets = []
    if callable(progress_cb):
        detail_parts = ["Atualizando a label do Botana para preparar a releitura."]
        if window_info.get("oldest_date") and window_info.get("newest_date"):
            detail_parts.append(
                f"Lote selecionado: {int(window_info.get('selected', 0) or 0)} mensagens, de {window_info.get('newest_date')} até {window_info.get('oldest_date')}."
            )
        if continue_after_id:
            detail_parts.append("Continuação do reprocessamento a partir do lote anterior.")
        progress_cb(
            progress_current=0,
            progress_total=len(messages),
            changed=changed,
            failed=failed,
            current_email="",
            current_subject="",
            current_date="",
            message=f"{len(messages)} mensagens mais recentes encontradas para reprocessar.",
            detail=" ".join(detail_parts),
            window_oldest_date=str(window_info.get("oldest_date", "") or "").strip(),
            window_newest_date=str(window_info.get("newest_date", "") or "").strip(),
            window_selected=int(window_info.get("selected", 0) or 0),
            continue_after_id=selected_oldest_id if remaining_after > 0 else "",
            continue_remaining=remaining_after,
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
                message="Reprocessamento em andamento.",
                detail=(
                    f"Data atual: {current_date}"
                    if current_date
                    else "Atualizando a label do Botana para preparar a releitura."
                ),
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
                message="Reprocessamento em andamento.",
                detail=(
                    f"Falhas: {failed}"
                    + (f" | Data atual: {current_date}" if current_date else "")
                ),
            )
    return {
        "ok": True,
        "matched": len(messages),
        "changed": changed,
        "failed": failed,
        "mark_unread": bool(mark_unread),
        "window_oldest_date": str(window_info.get("oldest_date", "") or "").strip(),
        "window_newest_date": str(window_info.get("newest_date", "") or "").strip(),
        "window_selected": int(window_info.get("selected", 0) or 0),
        "continue_after_id": selected_oldest_id if remaining_after > 0 else "",
        "continue_remaining": remaining_after,
        "targets": targets,
    }


def _find_missing_messages(
    max_messages: int,
    mode: str = "",
    nf_start: str = "",
    nf_end: str = "",
    date_from: str = "",
    date_to: str = "",
    nf_list=None,
    progress_cb=None,
) -> dict:
    mode_norm = _resolve_recovery_mode(mode=mode, nf_start=nf_start, nf_end=nf_end, date_from=date_from, date_to=date_to, nf_list=nf_list)
    start_nf, end_nf = _parse_nf_filter_range(nf_start, nf_end)
    start_date = _parse_iso_date_input(date_from)
    end_date = _parse_iso_date_input(date_to)
    nf_values = _parse_nf_selection_list(nf_list)
    if mode_norm == "period":
        if not start_date and not end_date:
            raise ValueError("Informe ao menos uma data para recuperar e-mails por período.")
    elif mode_norm == "range":
        if start_nf is None and end_nf is None:
            raise ValueError("Informe uma faixa de NF para recuperar e-mails por intervalo.")
    elif mode_norm == "list":
        if not nf_values:
            raise ValueError("Adicione ao menos uma NF na lista manual para recuperar e-mails.")
    else:
        raise ValueError("Escolha um período, uma faixa de NF ou uma lista manual para recuperar e-mails.")
    service = _get_gmail_service_locked()
    page_size = max(1, min(500, int(_RUNTIME_SETTINGS.get("gmail_page_size", 50) or 50)))
    if mode_norm == "list":
        return _find_missing_messages_for_nf_list(
            service,
            nf_values=nf_values,
            max_messages=max_messages,
            page_size=page_size,
            progress_cb=progress_cb,
        )
    wanted = max(1, min(1000, int(max_messages)))
    page_limit = _manual_scan_page_limit(wanted, page_size)
    query = _recover_search_query(mode=mode_norm, nf_start=nf_start, nf_end=nf_end, date_from=date_from, date_to=date_to, nf_list=nf_values)
    targets = []
    seen_ids = set()
    inspected = 0
    pages = 0
    page_token = None
    attachment_text_cache = {}
    subject_mismatch_notes = []
    subject_mismatch_count = 0
    criteria_desc = _describe_recovery_filters(mode=mode_norm, nf_start=nf_start, nf_end=nf_end, date_from=date_from, date_to=date_to, nf_list=nf_values) or "filtros informados"
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
            message="Varrendo Gmail em busca das mensagens escolhidas.",
            detail=f"Critérios: {criteria_desc}.",
        )
    while pages < page_limit and len(targets) < wanted:
        batch, next_page_token = buscarMessagesEnviadosPagina(
            service,
            max_results=page_size,
            page_token=page_token,
            query=query,
        )
        pages += 1
        if not batch and not next_page_token:
            break
        for item in batch:
            msg_id = str(item.get("id", "")).strip()
            if not msg_id or msg_id in seen_ids:
                continue
            seen_ids.add(msg_id)
            preview = _preview_from_message_item(item)
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
                    message=f"Analisando mensagens: {inspected} verificadas.",
                    detail=f"Encontradas {len(targets)} dentro dos filtros. Página {pages} de até {page_limit}.",
                )
            attachment_text = attachment_text_cache.get(msg_id, "")
            match_info = _recovery_match_details(
                service,
                msg_id,
                preview,
                mode=mode_norm,
                nf_start=start_nf,
                nf_end=end_nf,
                nf_list=nf_values,
                start_date=start_date,
                end_date=end_date,
                attachment_text=attachment_text,
            )
            attachment_text_cache[msg_id] = str(match_info.get("attachment_text", "") or "").strip()
            warning_note = str(match_info.get("warning_note", "") or "").strip()
            if warning_note:
                subject_mismatch_count += 1
                if len(subject_mismatch_notes) < 8:
                    note_parts = []
                    nf_hits = _extract_nf_numbers_from_text(warning_note)
                    if nf_hits:
                        note_parts.append("NF " + ", ".join(str(numero) for numero in nf_hits))
                    if current_date:
                        note_parts.append(current_date)
                    if current_subject:
                        note_parts.append(current_subject)
                    context = " | ".join(note_parts)
                    subject_mismatch_notes.append(
                        f"{context} -> {warning_note}" if context else warning_note
                    )
            if not bool(match_info.get("matched")):
                continue
            targets.append(
                {
                    "id": msg_id,
                    "threadId": str(item.get("threadId", "")).strip(),
                    "labelIds": list(item.get("labelIds", []) or []),
                    "snippet": str(item.get("snippet", "") or ""),
                    "subject": str(item.get("subject", "") or "").strip(),
                    "date": str(item.get("date", "") or "").strip(),
                    "from": str(item.get("from", "") or "").strip(),
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
        "mode": mode_norm,
        "targets": targets,
        "subject_mismatch_count": subject_mismatch_count,
        "subject_mismatch_notes": subject_mismatch_notes,
    }


def _start_recover_missing_background(
    max_messages: int,
    mode: str = "",
    nf_start: str = "",
    nf_end: str = "",
    date_from: str = "",
    date_to: str = "",
    nf_list=None,
) -> tuple[bool, dict]:
    mode_norm = _resolve_recovery_mode(mode=mode, nf_start=nf_start, nf_end=nf_end, date_from=date_from, date_to=date_to, nf_list=nf_list)
    criteria_desc = _describe_recovery_filters(mode=mode_norm, nf_start=nf_start, nf_end=nf_end, date_from=date_from, date_to=date_to, nf_list=nf_list)
    requested_nf_values = _parse_nf_selection_list(nf_list) if mode_norm == "list" else []
    initial_requested_nf_count = len(requested_nf_values)
    if not criteria_desc:
        return False, {"message": "Escolha um período, uma faixa de NF ou uma lista manual para recuperar e-mails."}
    snap = _manual_action_snapshot()
    if bool(snap.get("active")):
        if not snap.get("message"):
            label = str(snap.get("label") or "Acao manual").strip() or "Acao manual"
            snap["message"] = f"{label} ja esta em andamento."
        return False, snap
    started, snap = _manual_action_begin(
        "recover_missing",
        "Recuperação de e-mails",
        "Recuperação iniciada.",
        detail=f"Buscando mensagens com XML em {criteria_desc}.",
        progress_total=initial_requested_nf_count if initial_requested_nf_count > 0 else 0,
        requested_limit=int(max_messages),
        matched=0,
        inspected=0,
        requested_nf_count=initial_requested_nf_count,
        found_nf_numbers=[],
        missing_nf_numbers=requested_nf_values[:],
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
                    message="Interrompendo ciclo automático para recuperar e-mails.",
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
                detail=f"Varrendo Gmail em busca de mensagens com XML em {criteria_desc}.",
                phase="searching",
                matched=0,
                inspected=0,
                progress_total=initial_requested_nf_count if initial_requested_nf_count > 0 else 0,
                requested_nf_count=initial_requested_nf_count,
                found_nf_numbers=[],
                missing_nf_numbers=requested_nf_values[:],
            )
            result = _find_missing_messages(
                max_messages=max_messages,
                mode=mode_norm,
                nf_start=nf_start,
                nf_end=nf_end,
                date_from=date_from,
                date_to=date_to,
                nf_list=nf_list,
                progress_cb=_progress,
            )
            targets = list(result.get("targets") or [])
            matched = int(result.get("matched", 0) or 0)
            inspected = int(result.get("inspected", 0) or 0)
            subject_mismatch_count = int(result.get("subject_mismatch_count", 0) or 0)
            subject_mismatch_notes = list(result.get("subject_mismatch_notes") or [])
            found_nf_numbers = _parse_nf_selection_list(result.get("found_nf_numbers") or [])
            missing_nf_numbers = _parse_nf_selection_list(result.get("missing_nf_numbers") or [])
            requested_nf_count = int(result.get("requested_nf_count", initial_requested_nf_count) or 0)
            if targets:
                nf_progress_note = ""
                if requested_nf_count > 0:
                    nf_progress_note = f"NFs localizadas: {len(found_nf_numbers)} de {requested_nf_count}."
                    if missing_nf_numbers:
                        nf_progress_note = f"{nf_progress_note} Nao localizadas: {_format_nf_number_list(missing_nf_numbers)}."
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
                    detail=(
                        f"O Botana vai reler {len(targets)} mensagens encontradas em {criteria_desc}."
                        + (f" {nf_progress_note}" if nf_progress_note else "")
                    ),
                    requested_nf_count=requested_nf_count,
                    found_nf_numbers=found_nf_numbers,
                    missing_nf_numbers=missing_nf_numbers,
                )
                ok, msg = executar_um_ciclo(
                    max_messages_override=len(targets),
                    messages_override=targets,
                )
                proc = _process_snapshot().get("last", {})
                launched = int(proc.get("launched", 0) or 0)
                duplicates = int(proc.get("duplicates", 0) or 0)
                messages_read = int(proc.get("messages", 0) or 0)
                detail = (
                    f"Encontradas: {matched} | "
                    f"Analisadas: {inspected} | "
                    f"E-mails lidos: {messages_read} | "
                    f"Lançadas: {launched} | "
                    f"Duplicadas: {duplicates}"
                )
                if subject_mismatch_count > 0:
                    detail = f"{detail} | Divergência assunto/anexos: {subject_mismatch_count}"
                if ok:
                    if subject_mismatch_count > 0:
                        msg = (
                            "Recuperação concluída com aviso: "
                            "o assunto do e-mail não batia com os anexos em "
                            f"{subject_mismatch_count} mensagem{'s' if subject_mismatch_count != 1 else ''}."
                        )
                    elif launched > 0:
                        msg = (
                            f"Recuperação concluída: {launched} "
                            f"lançamento{'s' if launched != 1 else ''} "
                            "adicionado"
                            f"{'s' if launched != 1 else ''} na planilha."
                        )
                    elif duplicates > 0:
                        msg = (
                            "Recuperação concluída sem novas linhas: "
                            "as mensagens encontradas já estavam lançadas."
                        )
                    else:
                        msg = (
                            "Recuperação concluída sem lançamentos: "
                            "os e-mails foram lidos, mas nada novo entrou na planilha."
                        )
                if resume_loop:
                    restarted = iniciar_verificacao()
                    detail = f"{detail} | Loop automático {'retomado' if restarted else 'não retomado'}."
                if requested_nf_count > 0:
                    detail = f"{detail} | NFs localizadas: {len(found_nf_numbers)}/{requested_nf_count}"
                    if missing_nf_numbers:
                        detail = f"{detail} | NFs nao localizadas: {_format_nf_number_list(missing_nf_numbers)}"
                    if ok and missing_nf_numbers:
                        if subject_mismatch_count > 0:
                            msg = (
                                f"Recuperacao parcial com aviso: {len(found_nf_numbers)} de {requested_nf_count} NF(s) localizadas. "
                                f"Nao localizadas: {_format_nf_number_list(missing_nf_numbers)}. "
                                "Houve divergencia entre assunto/PDF e XML em pelo menos uma mensagem."
                            )
                        else:
                            msg = f"Recuperacao parcial: {len(found_nf_numbers)} de {requested_nf_count} NF(s) localizadas."
                            if launched > 0:
                                msg = (
                                    f"{msg} {launched} lancamento{'s' if launched != 1 else ''} "
                                    f"adicionado{'s' if launched != 1 else ''} na planilha."
                                )
                            elif duplicates > 0:
                                msg = f"{msg} As NFs localizadas ja estavam lancadas."
                            else:
                                msg = f"{msg} Nada novo entrou na planilha."
                            msg = f"{msg} Nao localizadas: {_format_nf_number_list(missing_nf_numbers)}."
                _manual_action_finish(
                    ok,
                    msg,
                    detail=detail,
                    progress_current=messages_read,
                    progress_total=len(targets),
                    matched=matched,
                    inspected=inspected,
                    launched=launched,
                    duplicates=duplicates,
                    messages_read=messages_read,
                    subject_mismatch_count=subject_mismatch_count,
                    subject_mismatch_notes=subject_mismatch_notes,
                    requested_limit=int(max_messages),
                    requested_nf_count=requested_nf_count,
                    found_nf_numbers=found_nf_numbers,
                    missing_nf_numbers=missing_nf_numbers,
                )
                return
            detail = f"Nenhuma mensagem com XML combinou com {criteria_desc}. Analisadas: {inspected}."
            if missing_nf_numbers:
                detail = f"{detail} NFs nao localizadas: {_format_nf_number_list(missing_nf_numbers)}."
            if resume_loop:
                restarted = iniciar_verificacao()
                detail = f"{detail} | Loop automático {'retomado' if restarted else 'não retomado'}."
            message = (
                "Nenhum e-mail foi encontrado para os filtros informados."
                if not missing_nf_numbers
                else (
                    "Nenhum e-mail foi encontrado para as NFs selecionadas. "
                    f"Nao localizadas: {_format_nf_number_list(missing_nf_numbers)}."
                )
            )
            if subject_mismatch_count > 0 and missing_nf_numbers:
                message = (
                    "Nenhum e-mail foi aceito para as NFs selecionadas. "
                    f"Nao localizadas: {_format_nf_number_list(missing_nf_numbers)}. "
                    "Os e-mails encontrados tinham divergencia entre assunto/PDF e XML."
                )
            elif subject_mismatch_count > 0:
                message = (
                    "Nenhum e-mail foi aceito para os filtros informados. "
                    "Os e-mails encontrados tinham divergencia entre assunto/PDF e XML."
                )
            _manual_action_finish(
                True,
                message,
                detail=detail,
                progress_current=matched,
                progress_total=requested_nf_count if requested_nf_count > 0 else 0,
                matched=matched,
                inspected=inspected,
                subject_mismatch_count=subject_mismatch_count,
                subject_mismatch_notes=subject_mismatch_notes,
                requested_limit=int(max_messages),
                requested_nf_count=requested_nf_count,
                found_nf_numbers=found_nf_numbers,
                missing_nf_numbers=missing_nf_numbers,
            )
        except Exception as exc:
            logger.exception("Falha na recuperação de e-mails em background: %s", exc)
            if resume_loop:
                try:
                    iniciar_verificacao()
                except Exception:
                    logger.exception("Falha ao retomar loop automatico apos erro na recuperação.")
            _manual_action_finish(False, "Erro na recuperação de e-mails.", detail=str(exc), requested_limit=int(max_messages))

    threading.Thread(target=_worker, daemon=True, name="botana-recover-emails").start()
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
<p id="loginHint">Entre com usuário e senha para continuar</p>
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
function _isPopupMode(){try{return new URLSearchParams(window.location.search||'').get('popup')==='1';}catch(_){return false;}}
function backToHub(){
  try{
    const ref=document.referrer?new URL(document.referrer):null;
    if(ref&&ref.origin){window.location.assign(ref.origin+'/');return;}
  }catch(_){}
  window.location.assign(new URL('/',window.location.origin).toString());
}
function initHubBackLogin(){
  const b=document.getElementById('hubBackLogin');
  const hint=document.getElementById('loginHint');
  if(_isPopupMode()){
    if(b)b.classList.add('hidden');
    if(hint)hint.textContent='Entre com usuário e senha para validar o acesso no Hub.';
    return;
  }
  if(!b)return;
  if(_BASE_PREFIX)b.classList.remove('hidden');
}
async function login(){
  const u=document.getElementById('u').value||'';
  const p=document.getElementById('p').value||'';
  const m=document.getElementById('m');
  const b=document.getElementById('b');
  const inputU=document.getElementById('u');
  const inputP=document.getElementById('p');
  b.disabled=true;
  m.textContent='Validando acesso';
  try{
    const r=await fetch(_url('/api/login'),{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:u,password:p})});
    const j=await r.json();
    if(r.ok&&j.ok){
      if(_isPopupMode()){
        if(inputU)inputU.disabled=true;
        if(inputP)inputP.disabled=true;
        m.style.color='#2e7d32';
        m.textContent='Login confirmado. Esta janela pode ser fechada.';
        try{if(window.opener)window.opener.postMessage({type:'botana-login-ok'},'*');}catch(_){}
        setTimeout(()=>{try{window.close();}catch(_){}},700);
        return;
      }
      window.location.href=_url('/');
      return;
    }
    m.style.color='#9c2c1d';
    m.textContent=j.message||'Usuário ou senha inválidos';
  }catch(_){
    m.style.color='#9c2c1d';
    m.textContent='Falha ao conectar com o servidor';
  }finally{
    if(!_isPopupMode()||m.textContent!=='Login confirmado. Esta janela pode ser fechada.'){
      b.disabled=false;
    }
  }
}
['u','p'].forEach(id=>{document.getElementById(id).addEventListener('keydown',(e)=>{if(e.key==='Enter')login();});});
initHubBackLogin();
</script></body></html>"""
def _render_server_html() -> str:
    return """<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Botana - Painel</title>
<link rel="preconnect" href="https://unpkg.com" crossorigin />
<link rel="stylesheet" href="https://unpkg.com/tabulator-tables@6.3.1/dist/css/tabulator.min.css" />
<style>
@import url('https://fonts.googleapis.com/css2?family=Lexend:wght@300;400;600;700;800&display=swap');
:root{--o:#da7a1c;--o2:#ee9b2f;--b:#4a2b18;--bg:#f8efe6;--br:#e4c6a7}
*{box-sizing:border-box}
body{margin:0;min-height:100vh;font-family:'Lexend',Arial,sans-serif;background:linear-gradient(160deg,rgba(41,22,11,.78),rgba(95,56,28,.72)),url('/assets/store-bg') center/cover fixed;display:flex;justify-content:center;align-items:center;padding:12px;color:#2a1b12}
.app{width:min(1150px,100%);border-radius:18px;overflow:visible;border:1px solid rgba(231,200,168,.9);background:linear-gradient(180deg,rgba(255,250,246,.96),rgba(255,245,235,.92));box-shadow:0 24px 60px rgba(21,11,6,.35)}
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
.title-help-row{display:flex;align-items:center;justify-content:center;gap:8px;margin-bottom:6px}
.title-help-row h3{margin:0;text-align:center}
.reproc-grid{display:grid;grid-template-columns:minmax(180px,240px);gap:8px;align-items:end;justify-content:center;justify-items:center}
.reproc-grid > div{display:flex;flex-direction:column;align-items:center}
.reproc-grid > div label{text-align:center}
.reproc-grid > div input,.reproc-grid > div select{width:min(220px,100%);text-align:center}
.reproc-card .muted{text-align:center}
.help-tip{position:relative;display:inline-flex;align-items:center;justify-content:center;width:20px;height:20px;border-radius:999px;border:1px solid #d6b18f;background:#fff7ef;color:#7a4d30;font-size:.8rem;font-weight:800;line-height:1;cursor:help;user-select:none;z-index:30}
.help-tip:hover,.help-tip:focus-visible{background:#fff1e3;outline:none}
.help-tip-bubble{position:absolute;right:0;left:auto;top:calc(100% + 8px);transform:translateY(-4px);width:min(320px,calc(100vw - 40px));max-width:min(320px,calc(100vw - 40px));padding:10px 12px;border-radius:12px;border:1px solid #e2b58d;background:#fffaf5;color:#6b4128;font-size:.78rem;font-weight:500;line-height:1.45;box-shadow:0 10px 24px rgba(21,11,6,.16);opacity:0;pointer-events:none;transition:opacity .18s ease,transform .18s ease;z-index:40;text-align:left}
.help-tip:hover .help-tip-bubble,.help-tip:focus-visible .help-tip-bubble{opacity:1;transform:translateY(0);pointer-events:auto}
.recover-shell{max-width:920px;margin:0 auto;display:grid;gap:12px}
.recover-grid{display:grid;grid-template-columns:minmax(160px,220px) minmax(0,1fr);grid-template-areas:"mode filter" "action action";gap:10px 14px;align-items:start}
.recover-group,.recover-mode-box,.recover-range-box,.recover-period-box,.recover-list-box,.recover-action-box{width:100%;display:flex;flex-direction:column;justify-content:center;align-items:center;padding:12px;border:1px solid rgba(211,172,139,.72);border-radius:14px;background:rgba(255,252,247,.88)}
.recover-mode-box{grid-area:mode;max-width:220px;justify-self:center}
.recover-period-box,.recover-range-box,.recover-list-box{grid-area:filter;max-width:100%;justify-self:stretch}
.recover-action-box{grid-area:action;max-width:360px;justify-self:center}
.recover-grid.mode-list .recover-list-box{max-width:760px}
.recover-grid.mode-list .recover-action-box{max-width:360px}
.recover-grid.mode-period .recover-period-box{max-width:460px}
.recover-grid.mode-range .recover-range-box{max-width:420px}
.recover-grid label,.recover-mode-box label,.recover-range-box label,.recover-period-box label,.recover-list-box label,.recover-action-box label{width:100%;text-align:center}
.recover-grid input,.recover-grid select,.recover-mode-box select,.recover-range-box input,.recover-period-box input,.recover-list-box input,.recover-action-box input{width:100%;max-width:220px;text-align:center}
.recover-grid .hidden{display:none !important}
.recover-range-fields,.recover-period-fields{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;width:100%;max-width:460px}
.recover-list-entry{display:flex;gap:8px;justify-content:center;align-items:center;flex-wrap:wrap;width:100%;max-width:500px}
.recover-list-entry input{width:min(260px,100%)}
.recover-list-tags{display:flex;gap:8px;flex-wrap:wrap;justify-content:center;align-items:center;min-height:46px;padding:10px;border:1px dashed rgba(110,58,27,.26);border-radius:14px;background:rgba(255,250,245,.82);width:100%;max-width:500px}
.recover-tag{display:inline-flex;align-items:center;gap:6px;padding:6px 10px;border-radius:999px;background:rgba(110,58,27,.12);border:1px solid rgba(110,58,27,.22);font-size:12px;font-weight:700;color:#5b3118}
.recover-tag button{border:none;background:transparent;color:inherit;font-weight:800;cursor:pointer;padding:0;line-height:1}
.recover-action-row{display:flex;gap:12px;justify-content:center;align-items:center;flex-wrap:wrap;width:100%}
.recover-action-row button{min-width:220px}
.cb{margin-top:8px;display:inline-flex;align-items:center;gap:8px}
.action-box{margin-top:10px;border:1px solid #d8b391;border-radius:10px;background:#fffaf5;padding:9px;display:grid;gap:6px}
.action-head{display:flex;justify-content:space-between;align-items:center;gap:8px}
.action-title{font-size:.84rem;font-weight:700;color:#5b321c}
.action-detail{font-size:.78rem;color:#6c4a35}
.action-progress{font-size:.78rem;color:#6b4128}
.hist-filters{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,180px));gap:10px 12px;align-items:end;justify-content:center;max-width:1320px;margin:0 auto}
.hist-filters > div{display:flex;flex-direction:column;justify-content:center;align-items:center;width:min(180px,100%);min-height:72px}
.hist-filters > div label{width:100%;text-align:center}
.hist-filters > div input,.hist-filters > div select{width:min(180px,100%);text-align:center}
.hist-run-wrap{display:flex;flex-direction:column;justify-content:center;align-items:center;gap:4px;width:min(180px,100%);min-height:72px}
.hist-run-wrap label{width:100%;text-align:center;visibility:hidden}
.hist-run-wrap button{width:min(180px,100%)}
.table-wrap{width:100%;overflow:auto;border:1px solid #d9af86;border-radius:10px;background:#fffdfb;box-shadow:inset 0 0 0 1px rgba(217,175,134,.22)}
.hist-toolbar{margin-top:8px;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}
.hist-reset-btn{padding:6px 10px;font-size:.78rem}
.hist-table{width:100%;min-width:1240px;border-collapse:collapse;font-size:.8rem;table-layout:fixed;border:1px solid #ddb38d}
.hist-table th,.hist-table td{border:1px solid #e7c4a5;padding:7px 8px;text-align:center;vertical-align:middle;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.hist-table th{position:sticky;top:0;background:#fff1e3;color:#5c341c;z-index:1;padding-right:18px;border-bottom:2px solid #cf9c73;overflow:visible}
.hist-table th.sortable{cursor:pointer;user-select:none}
.hist-table th.sortable:after{content:""}
.hist-table tbody tr:nth-child(even){background:rgba(255,244,232,.92)}
.hist-table tbody tr:hover{background:rgba(238,155,47,.08)}
.hist-table.audit-tone-mva tbody tr:nth-child(even){background:rgba(255,244,232,.92)}
.hist-table.audit-tone-eh tbody tr:nth-child(even){background:rgba(240,247,255,.96)}
.hist-table.audit-tone-todos tbody tr.audit-row-tone-mva{background:rgba(255,244,232,.92)}
.hist-table.audit-tone-todos tbody tr.audit-row-tone-eh{background:rgba(240,247,255,.96)}
.hist-table.audit-tone-mva tbody tr:hover{background:rgba(255,232,208,.96)!important}
.hist-table.audit-tone-eh tbody tr:hover{background:rgba(224,239,255,.96)!important}
.hist-table.audit-tone-todos tbody tr.audit-row-tone-mva:hover{background:rgba(255,232,208,.96)!important}
.hist-table.audit-tone-todos tbody tr.audit-row-tone-eh:hover{background:rgba(224,239,255,.96)!important}
.hist-table th.is-resizing{background:#ffe5cf}
.col-resizer{position:absolute;top:0;right:-6px;width:12px;height:100%;cursor:col-resize;user-select:none;touch-action:none;z-index:4}
.col-resizer::after{content:"";position:absolute;top:7px;bottom:7px;left:5px;width:2px;background:#c68551;border-radius:999px;opacity:.78}
.audit-filters{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,180px));gap:10px 12px;align-items:end;justify-content:center;max-width:1120px;margin:0 auto}
.audit-filters > div{display:flex;flex-direction:column;justify-content:center;align-items:center;width:min(180px,100%);min-height:72px}
.audit-filters > div label{width:100%;text-align:center}
.audit-filters > div input,.audit-filters > div select{width:min(180px,100%);text-align:center}
.audit-source-wrap{display:flex;flex-direction:column;justify-content:center;align-items:center;gap:6px;width:min(180px,100%);min-height:72px}
.audit-source-group{display:inline-flex;align-items:center;justify-content:center;gap:10px;flex-wrap:nowrap;width:min(180px,100%)}
.audit-source-btn{position:relative;display:inline-flex;align-items:center;justify-content:center;width:42px;height:42px;border-radius:12px;border:1px solid #d6b18f;background:#fffdfb;color:#6b4126;cursor:pointer;transition:transform .15s ease,box-shadow .15s ease,border-color .15s ease,background .15s ease}
.audit-source-btn:hover{transform:translateY(-1px);box-shadow:0 6px 14px rgba(92,52,28,.12)}
.audit-source-btn:focus-visible{outline:none;border-color:#a96024;box-shadow:0 0 0 3px rgba(218,122,28,.18)}
.audit-source-btn.active{background:linear-gradient(180deg,#fff0df,#ffe0bf);border-color:#cf8a4c;color:#5b3118;box-shadow:0 7px 16px rgba(92,52,28,.14)}
.audit-source-btn svg{width:18px;height:18px;display:block}
.audit-source-badge{position:absolute;right:4px;bottom:4px;display:inline-flex;align-items:center;justify-content:center;min-width:15px;height:15px;padding:0 4px;border-radius:999px;background:#7a4d30;color:#fff;font-size:.54rem;font-weight:800;line-height:1}
.audit-source-btn[data-empresa="mva"] .audit-source-badge{background:#d66f17}
.audit-source-btn[data-empresa="eh"] .audit-source-badge{background:#2d7a78}
.audit-source-btn[data-empresa="todos"] .audit-source-badge{background:#6b4126}
.audit-run-wrap{display:flex;flex-direction:column;justify-content:center;align-items:center;gap:4px;width:min(180px,100%);min-height:72px}
.audit-run-wrap label{width:100%;text-align:center;visibility:hidden}
.audit-run-wrap button{width:min(180px,100%)}
.audit-title{text-align:center}
.audit-toolbar{margin-top:8px;display:flex;flex-direction:column;justify-content:center;align-items:center;gap:6px}
.audit-state{min-height:20px;text-align:center;font-size:.83rem;color:#6b4126}
.audit-state.loading{color:#a25b18;font-weight:700}
.audit-summary{margin-top:10px;display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:10px}
.audit-summary .k{border:1px solid #e2b58d;border-radius:10px;background:linear-gradient(180deg,#fff7ef,#fff1e3);padding:10px;text-align:center}
.audit-summary .n{font-size:1.3rem;font-weight:800;color:#7a3d11}
.audit-summary .t{font-size:.8rem;color:#6b4126}
.audit-table{width:100%;min-width:1040px;border-collapse:collapse;font-size:.8rem;table-layout:fixed;border:1px solid #ddb38d}
.audit-table th,.audit-table td{border:1px solid #e7c4a5;padding:7px 8px;text-align:center;vertical-align:middle;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.audit-table th{position:sticky;top:0;background:#fff1e3;color:#5c341c;z-index:1;border-bottom:2px solid #cf9c73}
.audit-table th.sortable{cursor:pointer;user-select:none}
.audit-table th.sortable:after{content:""}
.audit-table th.sortable.asc,.audit-table th.sortable.desc{background:#fde8d2}
.audit-col-status{width:90px}
.audit-col-nf{width:84px}
.audit-col-cliente{width:340px}
.audit-col-sm{width:82px}
.audit-col-date{width:118px}
.audit-col-aba{width:110px}
.audit-table tbody tr{background:#fffdfb}
.audit-table.audit-tone-mva tbody tr:nth-child(even){background:rgba(255,244,232,.92)}
.audit-table.audit-tone-eh tbody tr:nth-child(even){background:rgba(240,247,255,.96)}
.audit-table.audit-tone-todos tbody tr.audit-row-tone-mva{background:rgba(255,244,232,.92)}
.audit-table.audit-tone-todos tbody tr.audit-row-tone-eh{background:rgba(240,247,255,.96)}
.audit-table.audit-tone-mva tbody tr:hover{background:rgba(255,232,208,.96)!important}
.audit-table.audit-tone-eh tbody tr:hover{background:rgba(224,239,255,.96)!important}
.audit-table.audit-tone-todos tbody tr.audit-row-tone-mva:hover{background:rgba(255,232,208,.96)!important}
.audit-table.audit-tone-todos tbody tr.audit-row-tone-eh:hover{background:rgba(224,239,255,.96)!important}
.audit-status{display:inline-flex;align-items:center;justify-content:center;padding:3px 8px;border-radius:999px;font-size:.72rem;font-weight:700;border:1px solid transparent}
.audit-status-btn{cursor:pointer;transition:transform .15s ease, box-shadow .15s ease}
.audit-status-btn:hover{transform:translateY(-1px);box-shadow:0 4px 10px rgba(92,52,28,.12)}
.audit-status-btn[disabled]{cursor:default;opacity:.78;box-shadow:none;transform:none}
.audit-table th input{ text-align:center }
.audit-table th input::placeholder{ text-align:center }
.audit-table th.audit-col-aba-head,.audit-table td.audit-cell-local{font-size:.72rem}
.audit-status.ok{background:#e9f8ec;color:#1c6a32;border-color:#87c69a}
.audit-status.aviso{background:#fff3dd;color:#8b5a00;border-color:#e7bf6e}
.audit-status.erro{background:#fde7ea;color:#a61d2d;border-color:#dc3545}
.audit-row-aviso{background:rgba(255,193,7,.12)!important}
.audit-row-aviso:hover{background:rgba(255,193,7,.2)!important}
.audit-row-erro{background:rgba(220,53,69,.14)!important}
.audit-row-erro:hover{background:rgba(220,53,69,.22)!important}
.audit-row-local-pending{background:rgba(240,198,79,.10)!important}
.audit-row-local-pending:hover{background:rgba(240,198,79,.16)!important}
.audit-row-local-pending td{border-bottom:3px dashed #f0c64f!important}
.audit-row-local-removed{background:rgba(240,198,79,.16)!important}
.audit-row-local-removed:hover{background:rgba(240,198,79,.22)!important}
.audit-row-local-removed td{border-bottom:3px solid #f0c64f!important}
.audit-tabulator,.panel-tabulator{display:none}
.audit-tabulator.active,.panel-tabulator.active{display:block}
#auditTableTabulator,#histTableTabulator,#watchTableTabulator{border:1px solid #ddb38d;border-radius:10px;overflow:hidden;background:#fffdfb}
#auditTableTabulator .tabulator,
#histTableTabulator .tabulator,
#watchTableTabulator .tabulator{border:none;background:#fffdfb;font-size:.8rem;color:#3f2819}
#auditTableTabulator .tabulator-header,
#histTableTabulator .tabulator-header,
#watchTableTabulator .tabulator-header{border-bottom:1px solid #e7c4a5;background:#fff1e3}
#auditTableTabulator .tabulator-col,
#auditTableTabulator .tabulator-header .tabulator-col,
#histTableTabulator .tabulator-col,
#histTableTabulator .tabulator-header .tabulator-col,
#watchTableTabulator .tabulator-col,
#watchTableTabulator .tabulator-header .tabulator-col{background:transparent;border-right:1px solid #efe0d0;color:#5c341c;font-weight:800}
#auditTableTabulator .tabulator-col-title,
#histTableTabulator .tabulator-col-title,
#watchTableTabulator .tabulator-col-title{display:block;width:100%;text-align:center}
#auditTableTabulator .tabulator-header-filter input,
#histTableTabulator .tabulator-header-filter input,
#watchTableTabulator .tabulator-header-filter input{text-align:center}
#auditTableTabulator .tabulator-header-filter input::placeholder,
#histTableTabulator .tabulator-header-filter input::placeholder,
#watchTableTabulator .tabulator-header-filter input::placeholder{text-align:center}
#auditTableTabulator .tabulator-row,
#histTableTabulator .tabulator-row,
#watchTableTabulator .tabulator-row{border-bottom:1px solid #f0e0cf;background:#fffdfb}
#auditTableTabulator.audit-tone-mva .tabulator-row:nth-child(even),
#histTableTabulator.audit-tone-mva .tabulator-row:nth-child(even),
#watchTableTabulator.audit-tone-mva .tabulator-row:nth-child(even){background:rgba(255,244,232,.92)}
#auditTableTabulator.audit-tone-eh .tabulator-row:nth-child(even),
#histTableTabulator.audit-tone-eh .tabulator-row:nth-child(even),
#watchTableTabulator.audit-tone-eh .tabulator-row:nth-child(even){background:rgba(240,247,255,.96)}
#auditTableTabulator.audit-tone-todos .tabulator-row.audit-row-tone-mva,
#histTableTabulator.audit-tone-todos .tabulator-row.audit-row-tone-mva,
#watchTableTabulator.audit-tone-todos .tabulator-row.audit-row-tone-mva{background:rgba(255,244,232,.92)!important}
#auditTableTabulator.audit-tone-todos .tabulator-row.audit-row-tone-eh,
#histTableTabulator.audit-tone-todos .tabulator-row.audit-row-tone-eh,
#watchTableTabulator.audit-tone-todos .tabulator-row.audit-row-tone-eh{background:rgba(240,247,255,.96)!important}
#auditTableTabulator.audit-tone-mva .tabulator-row:hover,
#auditTableTabulator.audit-tone-mva .tabulator-row.tabulator-selectable:hover,
#histTableTabulator.audit-tone-mva .tabulator-row:hover,
#histTableTabulator.audit-tone-mva .tabulator-row.tabulator-selectable:hover,
#watchTableTabulator.audit-tone-mva .tabulator-row:hover,
#watchTableTabulator.audit-tone-mva .tabulator-row.tabulator-selectable:hover{background:rgba(255,232,208,.96)!important}
#auditTableTabulator.audit-tone-eh .tabulator-row:hover,
#auditTableTabulator.audit-tone-eh .tabulator-row.tabulator-selectable:hover,
#histTableTabulator.audit-tone-eh .tabulator-row:hover,
#histTableTabulator.audit-tone-eh .tabulator-row.tabulator-selectable:hover,
#watchTableTabulator.audit-tone-eh .tabulator-row:hover,
#watchTableTabulator.audit-tone-eh .tabulator-row.tabulator-selectable:hover{background:rgba(224,239,255,.96)!important}
#auditTableTabulator.audit-tone-todos .tabulator-row.audit-row-tone-mva:hover,
#auditTableTabulator.audit-tone-todos .tabulator-row.audit-row-tone-mva.tabulator-selectable:hover,
#histTableTabulator.audit-tone-todos .tabulator-row.audit-row-tone-mva:hover,
#histTableTabulator.audit-tone-todos .tabulator-row.audit-row-tone-mva.tabulator-selectable:hover,
#watchTableTabulator.audit-tone-todos .tabulator-row.audit-row-tone-mva:hover,
#watchTableTabulator.audit-tone-todos .tabulator-row.audit-row-tone-mva.tabulator-selectable:hover{background:rgba(255,232,208,.96)!important}
#auditTableTabulator.audit-tone-todos .tabulator-row.audit-row-tone-eh:hover,
#auditTableTabulator.audit-tone-todos .tabulator-row.audit-row-tone-eh.tabulator-selectable:hover,
#histTableTabulator.audit-tone-todos .tabulator-row.audit-row-tone-eh:hover,
#histTableTabulator.audit-tone-todos .tabulator-row.audit-row-tone-eh.tabulator-selectable:hover,
#watchTableTabulator.audit-tone-todos .tabulator-row.audit-row-tone-eh:hover,
#watchTableTabulator.audit-tone-todos .tabulator-row.audit-row-tone-eh.tabulator-selectable:hover{background:rgba(224,239,255,.96)!important}
#auditTableTabulator .tabulator-cell,
#histTableTabulator .tabulator-cell,
#watchTableTabulator .tabulator-cell{border-right:1px solid #f3e8dc;padding:8px 9px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;text-align:center}
#auditTableTabulator .tabulator-header .tabulator-col:last-child .tabulator-col-content,
#histTableTabulator .tabulator-header .tabulator-col:last-child .tabulator-col-content,
#watchTableTabulator .tabulator-header .tabulator-col:last-child .tabulator-col-content{font-size:.72rem}
#auditTableTabulator .tabulator-header .tabulator-col:last-child .tabulator-header-filter input,
#histTableTabulator .tabulator-header .tabulator-col:last-child .tabulator-header-filter input,
#watchTableTabulator .tabulator-header .tabulator-col:last-child .tabulator-header-filter input{font-size:.72rem}
#auditTableTabulator .tabulator-row .tabulator-cell:last-child,
#histTableTabulator .tabulator-row .tabulator-cell:last-child,
#watchTableTabulator .tabulator-row .tabulator-cell:last-child{font-size:.72rem}
#auditTableTabulator .tabulator-footer,
#histTableTabulator .tabulator-footer,
#watchTableTabulator .tabulator-footer{border-top:1px solid #e7c4a5;background:#fff8f1;color:#6b4126;font-size:12px;font-weight:700}
#auditTableTabulator .tabulator-page,
#histTableTabulator .tabulator-page,
#watchTableTabulator .tabulator-page{border:1px solid #d9d0c5;background:#fff;color:#384658}
#auditTableTabulator .tabulator-page.active,
#histTableTabulator .tabulator-page.active,
#watchTableTabulator .tabulator-page.active{background:var(--o);color:#2b1408;border-color:var(--o)}
#auditTableTabulator .tabulator-row.audit-row-aviso,
#watchTableTabulator .tabulator-row.watch-row-aviso{background:rgba(255,193,7,.12)!important}
#auditTableTabulator .tabulator-row.audit-row-aviso:hover,
#watchTableTabulator .tabulator-row.watch-row-aviso:hover{background:rgba(255,193,7,.2)!important}
#auditTableTabulator .tabulator-row.audit-row-erro,
#watchTableTabulator .tabulator-row.watch-row-erro,
#histTableTabulator .tabulator-row.dup-row{background:rgba(220,53,69,.14)!important}
#auditTableTabulator .tabulator-row.audit-row-erro:hover,
#watchTableTabulator .tabulator-row.watch-row-erro:hover,
#histTableTabulator .tabulator-row.dup-row:hover{background:rgba(220,53,69,.22)!important}
#auditTableTabulator .tabulator-row.audit-row-local-pending{background:rgba(240,198,79,.10)!important}
#auditTableTabulator .tabulator-row.audit-row-local-pending:hover{background:rgba(240,198,79,.16)!important}
#auditTableTabulator .tabulator-row.audit-row-local-removed{background:rgba(240,198,79,.16)!important}
#auditTableTabulator .tabulator-row.audit-row-local-removed:hover{background:rgba(240,198,79,.22)!important}
#auditTableTabulator .tabulator-row.audit-row-local-removed .tabulator-cell{border-bottom:3px solid #f0c64f!important}
.watch-title{text-align:center}
.watch-filters{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,180px));gap:10px 12px;align-items:end;justify-content:center;max-width:960px;margin:0 auto}
.watch-filters > div{display:flex;flex-direction:column;justify-content:center;align-items:center;width:min(180px,100%);min-height:72px}
.watch-filters > div label{width:100%;text-align:center}
.watch-filters > div input{width:min(180px,100%);text-align:center}
.watch-run-wrap{display:flex;flex-direction:column;justify-content:center;align-items:center;gap:4px;width:min(180px,100%);min-height:72px}
.watch-run-wrap label{width:100%;text-align:center;visibility:hidden}
.watch-run-wrap button{width:min(180px,100%)}
.watch-actions{display:flex;justify-content:center;align-items:center;gap:8px;flex-wrap:wrap;margin-top:14px}
.watch-toolbar{margin-top:8px;display:flex;flex-direction:column;justify-content:center;align-items:center;gap:6px}
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
.watch-table.audit-tone-mva tbody tr:nth-child(even){background:rgba(255,244,232,.92)}
.watch-table.audit-tone-eh tbody tr:nth-child(even){background:rgba(240,247,255,.96)}
.watch-table.audit-tone-todos tbody tr.audit-row-tone-mva{background:rgba(255,244,232,.92)}
.watch-table.audit-tone-todos tbody tr.audit-row-tone-eh{background:rgba(240,247,255,.96)}
.watch-table.audit-tone-mva tbody tr:hover{background:rgba(255,232,208,.96)!important}
.watch-table.audit-tone-eh tbody tr:hover{background:rgba(224,239,255,.96)!important}
.watch-table.audit-tone-todos tbody tr.audit-row-tone-mva:hover{background:rgba(255,232,208,.96)!important}
.watch-table.audit-tone-todos tbody tr.audit-row-tone-eh:hover{background:rgba(224,239,255,.96)!important}
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
.ov{position:fixed;inset:0;z-index:99999;display:none;align-items:center;justify-content:center;background:rgba(22,10,5,.78);backdrop-filter:blur(3px)}
.ov.show{display:flex}
.ovb{width:min(440px,92vw);border-radius:14px;border:1px solid #f0c89d;background:linear-gradient(180deg,#fff6ec,#ffe8d4);text-align:center;padding:18px}
.cnt{margin-top:12px;font-size:2.4rem;font-weight:800;color:#b05714}
.continue-pop{position:fixed;inset:0;z-index:99997;display:none;align-items:center;justify-content:center;background:rgba(22,10,5,.72);backdrop-filter:blur(3px);padding:18px}
.continue-pop.show{display:flex}
.continue-pop-box{width:min(520px,92vw);border-radius:16px;border:1px solid #f0c89d;background:linear-gradient(180deg,#fff7ef,#ffe8d4);padding:18px;display:grid;gap:12px;text-align:center;box-shadow:0 18px 42px rgba(20,10,4,.22)}
.continue-pop-title{margin:0;font-size:1rem;color:#5c341c}
.continue-pop-text{font-size:.88rem;color:#6b4126}
.continue-pop-window{border:1px solid #efc9a3;border-radius:12px;background:#fff8f1;padding:10px;font-size:.83rem;color:#714224}
.continue-pop-fields{display:flex;gap:12px;justify-content:center;align-items:end;flex-wrap:wrap}
.continue-pop-fields > div{display:flex;flex-direction:column;align-items:center}
.continue-pop-fields label{text-align:center}
.continue-pop-fields input{width:min(110px,100%);text-align:center}
.continue-pop-actions{display:flex;gap:10px;justify-content:center;align-items:center;flex-wrap:wrap}
@media(max-width:900px){.lists{grid-template-columns:1fr}.cfg-grid{grid-template-columns:1fr}.cfg-fields{grid-template-columns:1fr 1fr}.reproc-grid{grid-template-columns:1fr}.recover-grid{grid-template-columns:1fr;grid-template-areas:"mode" "filter" "action"}.recover-mode-box,.recover-period-box,.recover-range-box,.recover-list-box,.recover-action-box{max-width:none}.recover-range-fields,.recover-period-fields{grid-template-columns:1fr 1fr}}
@media(max-width:1020px){.hist-filters{grid-template-columns:1fr 1fr 1fr}.audit-filters{grid-template-columns:1fr 1fr}.audit-summary{grid-template-columns:1fr 1fr 1fr}.watch-summary{grid-template-columns:1fr 1fr}.recover-range-fields,.recover-period-fields{grid-template-columns:1fr}}
@media(max-width:640px){.top-right{flex-direction:column;align-items:flex-end}.hist-filters{grid-template-columns:1fr}.audit-filters{grid-template-columns:1fr}.audit-summary{grid-template-columns:1fr 1fr}.watch-filters{grid-template-columns:1fr}.watch-summary{grid-template-columns:1fr 1fr}.recover-range-fields,.recover-period-fields{grid-template-columns:1fr}.recover-action-row{flex-direction:column;align-items:center}.watch-pop-search{grid-template-columns:1fr}.watch-pop-close{position:static}.continue-pop-fields,.continue-pop-actions{flex-direction:column;align-items:center}.help-tip-bubble{right:-8px;width:min(280px,calc(100vw - 24px));max-width:min(280px,calc(100vw - 24px))}}
</style>
<script src="https://unpkg.com/tabulator-tables@6.3.1/dist/js/tabulator.min.js"></script>
</head><body>
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
        <div class="title-help-row">
          <h3>Reprocessar e-mails</h3>
          <span class="help-tip" tabindex="0" aria-label="Ajuda do reprocessamento">?
            <span class="help-tip-bubble">Usa por padrão as mensagens mais recentes com label do Botana nas últimas duas semanas, remarca para reprocessamento e já executa a leitura em seguida. Para algo mais antigo, use Recuperar e-mails.</span>
          </span>
        </div>
        <div class="reproc-grid">
          <div>
            <label>Limite de mensagens</label>
            <input id="limit" type="number" value="100" min="1" max="1000"/>
          </div>
        </div>
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
      <div class="title-help-row">
        <h3>Recuperar e-mails</h3>
        <span class="help-tip" tabindex="0" aria-label="Ajuda da recuperação">?
          <span class="help-tip-bubble">Procura mensagens com XML pelos filtros informados e tenta lançar no financeiro mesmo que o e-mail já tenha label do Botana. Use período, faixa de NF ou monte uma lista manual; duplicidades continuam bloqueadas pelo writer.</span>
        </span>
      </div>
      <div class="recover-shell">
        <div id="recoverGrid" class="recover-grid mode-period">
          <div class="recover-mode-box">
            <label>Modo</label>
            <select id="recoverMode" onchange="toggleRecoverFilters()">
              <option value="period">Período</option>
              <option value="range">Faixa de NF</option>
              <option value="list">Escolha manual</option>
            </select>
          </div>
          <div id="recoverPeriodBox" class="recover-period-box">
            <label>Período</label>
            <div class="recover-period-fields">
              <input id="recoverDateFrom" type="date"/>
              <input id="recoverDateTo" type="date"/>
            </div>
          </div>
          <div id="recoverRangeBox" class="recover-range-box hidden">
            <label>Faixa de NF</label>
            <div class="recover-range-fields">
              <input id="recoverNfStart" type="text" inputmode="numeric" maxlength="5" placeholder="20247"/>
              <input id="recoverNfEnd" type="text" inputmode="numeric" maxlength="5" placeholder="20481"/>
            </div>
          </div>
          <div id="recoverListBox" class="recover-list-box hidden">
            <label>NFs escolhidas</label>
            <div class="recover-list-entry">
              <input id="recoverListInput" type="text" inputmode="numeric" maxlength="64" placeholder="20247, 20344" oninput="sanitizeRecoverNfInput(this)" onkeydown="recoverListKeydown(event)"/>
              <button type="button" class="sec" onclick="addRecoverNf()">Adicionar NF</button>
            </div>
            <div id="recoverListTags" class="recover-list-tags"><span class="muted">Nenhuma NF adicionada.</span></div>
          </div>
          <div class="recover-action-box">
            <div class="recover-action-row">
              <button id="recoverBtn" onclick="recoverEmails()">Recuperar</button>
            </div>
          </div>
        </div>
      </div>
    </section>
  </section>

  <section id="tabHist" class="tab-panel hidden">
    <section class="card" style="margin-top:10px">
      <h3>Histórico de processamento e lançamentos</h3>
      <div class="hist-filters">
        <div class="audit-source-wrap">
          <label>Origem</label>
          <div id="historyEmpresaGroup" class="audit-source-group" data-value="mva" role="group" aria-label="Origem do histórico">
            <button type="button" class="audit-source-btn active" data-empresa="mva" title="Somente MVA" aria-label="Somente MVA" onclick="setHistoryEmpresa('mva')">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 20V8l5 3V8l5 3V6l6 4v10"/><path d="M3 20h18"/><path d="M8 20v-4"/><path d="M13 20v-5"/><path d="M18 20v-3"/></svg>
              <span class="audit-source-badge" aria-hidden="true">M</span>
            </button>
            <button type="button" class="audit-source-btn" data-empresa="eh" title="Somente EH" aria-label="Somente EH" onclick="setHistoryEmpresa('eh')">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 20V9l4-2 4 2V6l6 3v11"/><path d="M4 20h16"/><path d="M8 20v-4"/><path d="M12 20v-5"/><path d="M17 20v-3"/></svg>
              <span class="audit-source-badge" aria-hidden="true">E</span>
            </button>
            <button type="button" class="audit-source-btn" data-empresa="todos" title="MVA + EH" aria-label="MVA + EH" onclick="setHistoryEmpresa('todos')">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="4" y="7" width="10" height="10" rx="2"/><rect x="10" y="5" width="10" height="12" rx="2"/></svg>
              <span class="audit-source-badge" aria-hidden="true">+</span>
            </button>
          </div>
        </div>
        <div><label>Data / Horário</label><input id="hAt" type="text" placeholder="06/04/2026 12:30 ou 2026-04-06"/></div>
        <div><label>Vencimento</label><input id="hVenc" type="date"/></div>
        <div><label>NF</label><input id="hNf" type="text" placeholder="49001"/></div>
        <div><label>Cliente</label><input id="hCliente" type="text" placeholder="Nome curto ou completo"/></div>
        <div><label>Aba</label><input id="hAba" type="text" placeholder="Janeiro ou MVA/Janeiro"/></div>
        <div><label>Limite</label><input id="hLimit" type="number" min="10" max="2000" value="300"/></div>
        <div class="hist-run-wrap"><label aria-hidden="true">&nbsp;</label><button onclick="loadHistory()">Aplicar filtros</button></div>
      </div>
      <div class="hist-toolbar">
        <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">
          <button type="button" class="sec hist-reset-btn" onclick="resetHistColumnWidths()">Resetar larguras</button>
        </div>
      </div>
      <div class="table-wrap" style="margin-top:10px">
        <div id="histTableTabulator" class="panel-tabulator"></div>
        <table id="histTableLegacy" class="hist-table">
          <colgroup>
            <col id="histCol-at" style="width:150px"/>
            <col id="histCol-venc" style="width:110px"/>
            <col id="histCol-doc" style="width:100px"/>
            <col id="histCol-cliente" style="width:240px"/>
            <col id="histCol-parcela" style="width:90px"/>
            <col id="histCol-vparcela" style="width:130px"/>
            <col id="histCol-vtotal" style="width:130px"/>
            <col id="histCol-local" style="width:150px"/>
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
            </tr>
          </thead>
          <tbody id="hBody"><tr><td colspan="8">Sem dados</td></tr></tbody>
        </table>
      </div>
    </section>
  </section>

  <section id="tabAudit" class="tab-panel hidden">
    <section class="card" style="margin-top:10px">
      <div class="title-help-row">
        <h3 class="audit-title">Conferência de parcelas lançadas</h3>
        <span class="help-tip" tabindex="0" aria-label="Ajuda da conferência">?
          <span class="help-tip-bubble">A conferência lê diretamente as planilhas, seleciona as NFs pelo filtro escolhido e compara o total esperado da NF com as parcelas registradas nas abas.</span>
        </span>
      </div>
      <div class="audit-filters">
        <div>
          <label>Modo</label>
          <select id="aMode" onchange="toggleAuditFilters()">
            <option value="mes">Mês do lançamento</option>
            <option value="nfs">Faixa de NF</option>
            <option value="todos">Tudo</option>
          </select>
        </div>
        <div id="auditMonthField">
          <label>Mês</label>
          <input id="aMonth" type="month"/>
        </div>
        <div class="audit-source-wrap">
          <label>Origem</label>
          <div id="auditEmpresaGroup" class="audit-source-group" data-value="mva" role="group" aria-label="Origem da conferência">
            <button type="button" class="audit-source-btn active" data-empresa="mva" title="Somente MVA" aria-label="Somente MVA" onclick="setAuditEmpresa('mva')">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 20V8l5 3V8l5 3V6l6 4v10"/><path d="M3 20h18"/><path d="M8 20v-4"/><path d="M13 20v-5"/><path d="M18 20v-3"/></svg>
              <span class="audit-source-badge" aria-hidden="true">M</span>
            </button>
            <button type="button" class="audit-source-btn" data-empresa="eh" title="Somente EH" aria-label="Somente EH" onclick="setAuditEmpresa('eh')">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 20V9l4-2 4 2V6l6 3v11"/><path d="M4 20h16"/><path d="M8 20v-4"/><path d="M12 20v-5"/><path d="M17 20v-3"/></svg>
              <span class="audit-source-badge" aria-hidden="true">E</span>
            </button>
            <button type="button" class="audit-source-btn" data-empresa="todos" title="MVA + EH" aria-label="MVA + EH" onclick="setAuditEmpresa('todos')">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="4" y="7" width="10" height="10" rx="2"/><rect x="10" y="5" width="10" height="12" rx="2"/></svg>
              <span class="audit-source-badge" aria-hidden="true">+</span>
            </button>
          </div>
        </div>
        <div id="auditNfStartField" class="hidden">
          <label>NF inicial</label>
          <input id="aNfStart" type="text" placeholder="49001"/>
        </div>
        <div id="auditNfEndField" class="hidden">
          <label>NF final</label>
          <input id="aNfEnd" type="text" placeholder="49100"/>
        </div>
        <div class="audit-run-wrap"><label aria-hidden="true">&nbsp;</label><button id="auditRunBtn" onclick="loadParcelAudit()">Conferir parcelas</button></div>
      </div>
      <div class="audit-toolbar">
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
        <div id="auditTableTabulator" class="audit-tabulator"></div>
        <table id="auditTableLegacy" class="audit-table">
          <colgroup>
            <col class="audit-col-status"/>
            <col class="audit-col-nf"/>
            <col class="audit-col-cliente"/>
            <col class="audit-col-sm"/>
            <col class="audit-col-sm"/>
            <col class="audit-col-sm"/>
            <col class="audit-col-date"/>
            <col class="audit-col-aba"/>
          </colgroup>
          <thead>
            <tr>
              <th class="sortable" data-key="status">Status</th>
              <th class="sortable" data-key="nf">NF</th>
              <th class="sortable" data-key="cliente">Cliente</th>
              <th class="sortable" data-key="parcelas">Parc.</th>
              <th class="sortable" data-key="faltando">Faltando</th>
              <th class="sortable" data-key="duplicada">Duplicadas</th>
              <th class="sortable" data-key="vencimento">Últ. venc.</th>
              <th class="sortable audit-col-aba-head" data-key="local">Aba</th>
            </tr>
          </thead>
          <tbody id="aBody"><tr><td colspan="8">Sem dados</td></tr></tbody>
        </table>
      </div>
    </section>
  </section>

  <section id="tabWatch" class="tab-panel hidden">
    <section class="card" style="margin-top:10px">
      <div class="title-help-row">
        <h3 class="watch-title">Boletos e depósitos próximos do limite</h3>
        <span class="help-tip" tabindex="0" aria-label="Ajuda dos prazos">?
          <span class="help-tip-bubble">A relação lê diretamente as planilhas e lista apenas títulos com Status vazio ou A Receber. Boletos futuros ficam em amarelo; itens que vencem hoje ou já passaram ficam em vermelho.</span>
        </span>
      </div>
      <div class="watch-filters">
        <div>
          <label>Boletos em até</label>
          <input id="wBoletoDays" type="number" min="1" max="7" value="7"/>
        </div>
        <div>
          <label>Depósitos há pelo menos</label>
          <input id="wDepositoDays" type="number" min="1" max="7" value="7"/>
        </div>
        <div class="audit-source-wrap">
          <label>Origem</label>
          <div id="watchEmpresaGroup" class="audit-source-group" data-value="mva" role="group" aria-label="Origem dos prazos">
            <button type="button" class="audit-source-btn active" data-empresa="mva" title="Somente MVA" aria-label="Somente MVA" onclick="setWatchEmpresa('mva')">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 20V8l5 3V8l5 3V6l6 4v10"/><path d="M3 20h18"/><path d="M8 20v-4"/><path d="M13 20v-5"/><path d="M18 20v-3"/></svg>
              <span class="audit-source-badge" aria-hidden="true">M</span>
            </button>
            <button type="button" class="audit-source-btn" data-empresa="eh" title="Somente EH" aria-label="Somente EH" onclick="setWatchEmpresa('eh')">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M5 20V9l4-2 4 2V6l6 3v11"/><path d="M4 20h16"/><path d="M8 20v-4"/><path d="M12 20v-5"/><path d="M17 20v-3"/></svg>
              <span class="audit-source-badge" aria-hidden="true">E</span>
            </button>
            <button type="button" class="audit-source-btn" data-empresa="todos" title="MVA + EH" aria-label="MVA + EH" onclick="setWatchEmpresa('todos')">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="4" y="7" width="10" height="10" rx="2"/><rect x="10" y="5" width="10" height="12" rx="2"/></svg>
              <span class="audit-source-badge" aria-hidden="true">+</span>
            </button>
          </div>
        </div>
        <div class="watch-run-wrap"><label aria-hidden="true">&nbsp;</label><button id="watchRunBtn" onclick="loadDueWatch()">Atualizar relação</button></div>
      </div>
      <div class="watch-actions">
        <button type="button" class="sec" onclick="openWatchSearchModal()">Buscar boletos em aberto</button>
      </div>
      <div class="watch-toolbar">
        <div id="watchStatus" class="watch-state">Pronto para consultar.</div>
      </div>
      <div class="watch-summary">
        <div class="k"><div id="watchK1" class="n">0</div><div class="t">Total na relação</div></div>
        <div class="k"><div id="watchK2" class="n">0</div><div class="t">Boletos a vencer</div></div>
        <div class="k"><div id="watchK3" class="n">0</div><div class="t">Boletos no limite</div></div>
        <div class="k"><div id="watchK4" class="n">0</div><div class="t">Depósitos atrasados</div></div>
      </div>
      <div class="table-wrap" style="margin-top:10px">
        <div id="watchTableTabulator" class="panel-tabulator"></div>
        <table id="watchTableLegacy" class="watch-table">
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
<div id="continueReprocessModal" class="continue-pop" onclick="closeContinueReprocessModal(event)">
  <div class="continue-pop-box" onclick="event.stopPropagation()">
    <h4 class="continue-pop-title">Continuar reprocessando?</h4>
    <div id="continueReprocessText" class="continue-pop-text">O último lote terminou. Você pode continuar do próximo lote mais antigo.</div>
    <div id="continueReprocessWindow" class="continue-pop-window">Janela do lote: -</div>
    <div class="continue-pop-fields">
      <div>
        <label>Quantidade a mais</label>
        <input id="continueReprocessQty" type="number" min="1" max="1000" value="100"/>
      </div>
    </div>
    <div class="continue-pop-actions">
      <button id="continueReprocessYesBtn" type="button" onclick="continueReprocessFromPrompt()">Sim, continuar</button>
      <button type="button" class="sec" onclick="closeContinueReprocessModal()">Não, concluir</button>
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
let _lastManualActionState={};
let _lastReprocessPromptKey='';
let _dismissedReprocessPromptKey='';
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
let _recoverNfSelections=[];
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
  if(next!=='audit'){
    _stopAuditJobPolling();
  }
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
function _recoverDigits(value){return String(value||'').replace(/\\D+/g,'').slice(0,5);}
function _parseRecoverNfInput(value){
  const seen=new Set();
  return (String(value||'').match(/\\d{1,5}/g)||[])
    .map((item)=>String(item||'').slice(0,5))
    .filter((item)=>item&&!seen.has(item)&&(seen.add(item),true));
}
function sanitizeRecoverNfInput(el){
  if(!el)return;
  const values=_parseRecoverNfInput(el.value||'');
  el.value=values.join(', ');
}
function _renderRecoverNfTags(){
  const box=document.getElementById('recoverListTags');
  if(!box)return;
  if(!_recoverNfSelections.length){
    box.innerHTML='<span class="muted">Nenhuma NF adicionada.</span>';
    return;
  }
  box.innerHTML=_recoverNfSelections.map((nf)=>`<span class="recover-tag">${_esc(nf)} <button type="button" title="Remover" onclick="removeRecoverNf('${_esc(nf)}')">×</button></span>`).join('');
}
function addRecoverNf(){
  const input=document.getElementById('recoverListInput');
  const values=_parseRecoverNfInput(input&&input.value||'');
  if(!values.length){
    if(input)input.focus();
    return;
  }
  values.forEach((nf)=>{if(!_recoverNfSelections.includes(nf))_recoverNfSelections.push(nf);});
  if(input){input.value='';input.focus();}
  _renderRecoverNfTags();
}
function removeRecoverNf(nf){
  _recoverNfSelections=_recoverNfSelections.filter((item)=>item!==String(nf||''));
  _renderRecoverNfTags();
}
function recoverListKeydown(ev){
  if(ev.key!=='Enter')return;
  ev.preventDefault();
  addRecoverNf();
}
function _focusRecoverField(id){
  const el=document.getElementById(id);
  if(!el||el.classList.contains('hidden')||el.closest('.hidden'))return false;
  try{el.focus();}catch(_){return false;}
  return true;
}
function handleRecoverFieldKeydown(ev){
  if(ev.key!=='Enter')return;
  ev.preventDefault();
  const mode=String((document.getElementById('recoverMode')||{}).value||'period');
  const currentId=String((ev&&ev.target&&ev.target.id)||'').trim();
  if(currentId==='recoverMode'){
    if(mode==='range'){
      if(_focusRecoverField('recoverNfStart'))return;
    }else if(mode==='list'){
      if(_focusRecoverField('recoverListInput'))return;
    }else{
      if(_focusRecoverField('recoverDateFrom'))return;
    }
    recoverEmails();
    return;
  }
  if(mode==='period'){
    if(currentId==='recoverDateFrom'&&_focusRecoverField('recoverDateTo'))return;
    recoverEmails();
    return;
  }
  if(mode==='range'){
    if(currentId==='recoverNfStart'&&_focusRecoverField('recoverNfEnd'))return;
    recoverEmails();
    return;
  }
  if(mode==='list'){
    if(currentId==='recoverListInput'){
      addRecoverNf();
      return;
    }
    recoverEmails();
  }
}
function toggleRecoverFilters(){
  const mode=String((document.getElementById('recoverMode')||{}).value||'period');
  const grid=document.getElementById('recoverGrid');
  const period=document.getElementById('recoverPeriodBox');
  const range=document.getElementById('recoverRangeBox');
  const list=document.getElementById('recoverListBox');
  if(grid){
    grid.classList.remove('mode-period','mode-range','mode-list');
    grid.classList.add(`mode-${mode}`);
  }
  if(period)period.classList.toggle('hidden',mode!=='period');
  if(range)range.classList.toggle('hidden',mode!=='range');
  if(list)list.classList.toggle('hidden',mode!=='list');
}
function _reprocessPromptKey(action){
  const a=action||{};
  return [String(a.finished_at||''),String(a.continue_after_id||''),String(a.window_oldest_date||''),String(a.window_newest_date||'')].join('|');
}
function closeContinueReprocessModal(ev,markDismissed=true){
  if(ev&&ev.target&&ev.currentTarget&&ev.target!==ev.currentTarget)return;
  const modal=document.getElementById('continueReprocessModal');
  if(modal)modal.classList.remove('show');
  const key=_reprocessPromptKey(_lastManualActionState||{});
  if(markDismissed&&key)_dismissedReprocessPromptKey=key;
}
function _openContinueReprocessModal(action){
  const modal=document.getElementById('continueReprocessModal');
  if(!modal)return;
  const qty=document.getElementById('continueReprocessQty');
  const requested=Math.max(1, Number((action&&action.requested_limit)||100));
  if(qty)qty.value=String(requested);
  const windowEl=document.getElementById('continueReprocessWindow');
  const textEl=document.getElementById('continueReprocessText');
  const windowNewest=String((action&&action.window_newest_date)||'').trim();
  const windowOldest=String((action&&action.window_oldest_date)||'').trim();
  const remaining=Math.max(0, Number((action&&action.continue_remaining)||0));
  if(windowEl)windowEl.textContent=windowNewest&&windowOldest?`Janela do lote concluído: ${windowNewest} até ${windowOldest}`:'Janela do lote concluído: -';
  if(textEl)textEl.textContent=remaining>0?`Ainda existem ${remaining} mensagens mais antigas disponíveis para continuar do ponto onde esse lote parou.`:'O último lote terminou. Você pode continuar do próximo lote mais antigo.';
  modal.classList.add('show');
}
function _maybePromptContinueReprocess(action){
  const a=action||{};
  if(String(a.kind||'').trim()!=='reprocess')return;
  if(!!a.active)return;
  if(String(a.status||'').trim()!=='success')return;
  if(Math.max(0, Number(a.continue_remaining||0))<=0)return;
  const key=_reprocessPromptKey(a);
  if(!key||key===_lastReprocessPromptKey||key===_dismissedReprocessPromptKey)return;
  _lastReprocessPromptKey=key;
  _openContinueReprocessModal(a);
}
async function continueReprocessFromPrompt(){
  const btn=document.getElementById('continueReprocessYesBtn');
  const qtyEl=document.getElementById('continueReprocessQty');
  const extra=Math.max(1, Math.min(1000, Number((qtyEl&&qtyEl.value)||100)||100));
  const continueAfterId=String((_lastManualActionState&&_lastManualActionState.continue_after_id)||'').trim();
  if(!continueAfterId){
    closeContinueReprocessModal();
    return;
  }
  if(btn){btn.disabled=true;btn.textContent='Continuando...';}
  try{
    await api('/api/reprocess',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({max_messages:extra,continue_after_id:continueAfterId})});
    closeContinueReprocessModal();
    await refresh();
  }catch(err){
    alert('Erro ao continuar o reprocessamento: '+(err&&err.message||err));
  }finally{
    if(btn){btn.disabled=false;btn.textContent='Sim, continuar';}
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
  const windowOldestDate=String(a.window_oldest_date||'').trim();
  const windowNewestDate=String(a.window_newest_date||'').trim();
  const windowSelected=Math.max(0, Number(a.window_selected||0));
  return {requested,total,current,visibleTotal,perc,currentEmail,currentDate,currentSubject,windowOldestDate,windowNewestDate,windowSelected};
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
    if(view.currentDate)detailParts.push(`Data ${view.currentDate}`);
    if(view.windowOldestDate&&view.windowNewestDate)detailParts.push(`Janela ${view.windowNewestDate} até ${view.windowOldestDate}`);
    runEl.textContent='Loop: reprocessamento manual em andamento';
    nowEl.textContent=detailParts.length?`Mensagem atual: ${detailParts.join(' | ')}`:'Mensagem atual: buscando e-mails para reprocessar';
    const last=p.last||{};
    const lastEnd=last.finished_at?_fmtDateTime(last.finished_at):'-';
    let statusTxt='-';
    if(last.ok===true) statusTxt='OK';
    else if(last.ok===false) statusTxt='Erro';
    lastEl.textContent=`Último ciclo automático: ${statusTxt} em ${lastEnd}`;
    barFill.style.width=String(view.perc)+'%';
    barLabel.textContent=`Reprocessamento: ${view.current}/${view.visibleTotal||'-'} (${view.perc}%) | Limite pedido ${view.requested||'-'} | Falhas ${Number(a.failed||0)}`;
    return;
  }
  if(active&&kind==='reprocess'&&phase==='processing'){
    const cur=p.current||{};
    const currentLimit=Math.max(1, Number(a.progress_total||_manualRequestedLimit(a,maxMessages)));
    const curStart=cur.started_at?_fmtDateTime(cur.started_at):'-';
    runEl.textContent='Loop: executando leitura do reprocessamento';
    nowEl.textContent=`Ciclo atual: início ${curStart}`;
    const last=p.last||{};
    const lastEnd=last.finished_at?_fmtDateTime(last.finished_at):'-';
    let statusTxt='-';
    if(last.ok===true) statusTxt='OK';
    else if(last.ok===false) statusTxt='Erro';
    lastEl.textContent=`Último ciclo automático: ${statusTxt} em ${lastEnd}`;
    let curV=Number(cur.messages||0);
    if(!Number.isFinite(curV)||curV<0)curV=0;
    if(curV>currentLimit)curV=currentLimit;
    const perc=Math.max(0,Math.min(100,Math.round((curV/Math.max(1,currentLimit))*100)));
    barFill.style.width=String(perc)+'%';
    barLabel.textContent=`Relançamento: ${curV}/${currentLimit} (${perc}%) | Falhas ${Number(a.failed||0)}`;
    return;
  }
  if(active&&kind==='recover_missing'&&phase!=='processing'){
    const matched=Math.max(0, Number(a.matched||0));
    const inspected=Math.max(0, Number(a.inspected||0));
    const perc=Math.max(12,Math.min(88,Math.round((Math.min(inspected,200)/200)*100)));
    const detailParts=[];
    if(String(a.current_email||'').trim())detailParts.push(`E-mail atual ${String(a.current_email||'').trim()}`);
    if(String(a.current_date||'').trim())detailParts.push(`Data ${String(a.current_date||'').trim()}`);
    if(String(a.current_subject||'').trim())detailParts.push(`Assunto ${String(a.current_subject||'').trim()}`);
    runEl.textContent='Loop: recuperação manual em andamento';
    nowEl.textContent=detailParts.length?`Mensagem atual: ${detailParts.join(' | ')}`:'Mensagem atual: varrendo e-mails com XML';
    const last=p.last||{};
    const lastEnd=last.finished_at?_fmtDateTime(last.finished_at):'-';
    let statusTxt='-';
    if(last.ok===true) statusTxt='OK';
    else if(last.ok===false) statusTxt='Erro';
    lastEl.textContent=`Último ciclo automático: ${statusTxt} em ${lastEnd}`;
    barFill.style.width=String(perc)+'%';
    barLabel.textContent=`Recuperação: ${matched} encontradas | ${inspected} analisadas`;
    return;
  }
  if(active&&kind==='recover_missing'&&phase==='processing'){
    const cur=p.current||{};
    const currentLimit=Math.max(1, Number(a.progress_total||_manualRequestedLimit(a,maxMessages)));
    const matched=Math.max(0, Number(a.matched||0));
    const inspected=Math.max(0, Number(a.inspected||0));
    const curStart=cur.started_at?_fmtDateTime(cur.started_at):'-';
    runEl.textContent='Loop: executando leitura da recuperação';
    nowEl.textContent=`Ciclo atual: início ${curStart}`;
    const last=p.last||{};
    const lastEnd=last.finished_at?_fmtDateTime(last.finished_at):'-';
    let statusTxt='-';
    if(last.ok===true) statusTxt='OK';
    else if(last.ok===false) statusTxt='Erro';
    lastEl.textContent=`Último ciclo automático: ${statusTxt} em ${lastEnd}`;
    let curV=Number(cur.messages||0);
    if(!Number.isFinite(curV)||curV<0)curV=0;
    if(curV>currentLimit)curV=currentLimit;
    const perc=Math.max(0,Math.min(100,Math.round((curV/Math.max(1,currentLimit))*100)));
    barFill.style.width=String(perc)+'%';
    barLabel.textContent=`Recuperação: ${curV}/${currentLimit} (${perc}%) | Encontradas ${matched} | Analisadas ${inspected}`;
    return;
  }
  const reading=!!p.reading;
  const running=!!p.running;
  if(reading) runEl.textContent='Loop: executando ciclo agora';
  else if(running) runEl.textContent='Loop: ativo (aguardando proximo ciclo)';
  else runEl.textContent='Loop: pausado';

  const cur=p.current||{};
  const curStart=cur.started_at?_fmtDateTime(cur.started_at):'-';
  nowEl.textContent=`Ciclo atual: início ${curStart}`;

  const last=p.last||{};
  const lastEnd=last.finished_at?_fmtDateTime(last.finished_at):'-';
  let statusTxt='-';
  if(last.ok===true) statusTxt='OK';
  else if(last.ok===false) statusTxt='Erro';
  let msg=`Último ciclo: ${statusTxt} em ${lastEnd}`;
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
    progressEl.textContent=`Mensagens lidas: ${curV}/${maxV}`;
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
      progressEl.textContent=`Mensagens lidas: ${curV}/${maxV} | Falhas ${Number(a.failed||0)}`;
      msgEl.textContent=String(a.message||'Reprocessamento em andamento.');
      detailEl.textContent=String(a.detail||'Leitura e relançamento em andamento.');
      _setManualBadge('ok','Lendo');
    }else{
      const view=_reprocessView(a);
      const currentParts=[];
      if(view.currentDate)currentParts.push(`Data: ${view.currentDate}`);
      if(view.windowOldestDate&&view.windowNewestDate)currentParts.push(`Janela: ${view.windowNewestDate} até ${view.windowOldestDate}`);
      barEl.style.width=String(view.visibleTotal>0?Math.max(6,view.perc):18)+'%';
      progressEl.textContent=`Mensagens: ${view.current}/${view.visibleTotal||'-'} | Limite pedido ${view.requested||'-'} | Falhas ${Number(a.failed||0)}${currentParts.length?` | ${currentParts.join(' | ')}`:''}`;
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
      progressEl.textContent=`Mensagens lidas: ${curV}/${maxV} | Encontradas ${matched} | Analisadas ${inspected}`;
      msgEl.textContent=String(a.message||'Recuperação em andamento.');
      detailEl.textContent=String(a.detail||'Leitura e lançamento das mensagens encontradas em andamento.');
      _setManualBadge('ok','Lendo');
    }else{
      const perc=Math.max(12,Math.min(88,Math.round((Math.min(inspected,200)/200)*100)));
      const currentParts=[];
      if(String(a.current_email||'').trim())currentParts.push(`E-mail atual: ${String(a.current_email||'').trim()}`);
      if(String(a.current_date||'').trim())currentParts.push(`Data: ${String(a.current_date||'').trim()}`);
      barEl.style.width=String(perc)+'%';
      progressEl.textContent=`Encontradas ${matched} | Analisadas ${inspected}${currentParts.length?` | ${currentParts.join(' | ')}`:''}`;
      msgEl.textContent=String(a.message||'Recuperação em andamento.');
      detailEl.textContent=String(a.detail||'Varrendo e-mails com XML pelos filtros escolhidos.');
      _setManualBadge('ok','Varrendo');
    }
  }else{
    const status=String(a.status||'idle');
    const finished=String(a.finished_at||'').trim();
    const windowParts=[];
    if(String(a.window_newest_date||'').trim()&&String(a.window_oldest_date||'').trim())windowParts.push(`Janela ${String(a.window_newest_date||'').trim()} até ${String(a.window_oldest_date||'').trim()}`);
    const finishedText=finished?`Última atualização: ${_fmtDateTime(finished)}`:'Use o botão acima para reprocessar as mensagens e executar a leitura no mesmo fluxo.';
    barEl.style.width=status==='success'&&Number(a.progress_total||0)>0?'100%':'0%';
    progressEl.textContent=status==='success'||status==='error'
      ? `Progresso final: ${Number(a.progress_current||0)}/${Number(a.progress_total||0)||'-'} | Limite pedido ${Number(a.requested_limit||0)||'-'} | Falhas ${Number(a.failed||0)}${windowParts.length?` | ${windowParts.join(' | ')}`:''}`
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
    recoverBtn.textContent=active&&kind==='recover_missing'?(phase==='processing'?'Lendo...':'Buscando...'):'Recuperar';
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
    _lastManualActionState=(j.manual_action||{});
    if(_lastManualActionState&&_lastManualActionState.active)closeContinueReprocessModal(null,false);
    _maybePromptContinueReprocess(_lastManualActionState);
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
    const payload={account:'principal',max_messages:Number(document.getElementById('limit').value||100)};
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
async function recoverEmails(){
  const btn=document.getElementById('recoverBtn');
  const recoverMaxMessages=1000;
  if(btn){btn.disabled=true;btn.textContent='Iniciando...';}
  const msgEl=document.getElementById('manualActionMsg');
  const detailEl=document.getElementById('manualActionDetail');
  const mode=String(document.getElementById('recoverMode').value||'period');
  if(msgEl)msgEl.textContent='Solicitação enviada. Procurando e-mails para recuperar...';
  if(detailEl)detailEl.textContent='O Botana vai localizar as mensagens pelos filtros escolhidos e tentar lançar o que encontrar no financeiro.';
  try{
    const payload={
      max_messages:recoverMaxMessages,
      mode:mode,
      nf_start:_recoverDigits(document.getElementById('recoverNfStart').value||''),
      nf_end:_recoverDigits(document.getElementById('recoverNfEnd').value||''),
      date_from:String(document.getElementById('recoverDateFrom').value||'').trim(),
      date_to:String(document.getElementById('recoverDateTo').value||'').trim(),
      nf_list:_recoverNfSelections.slice(),
    };
    const j=await api('/api/recover-emails',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
    if(j&&j.friendly&&msgEl)msgEl.textContent=String(j.friendly);
    await refresh();
  }catch(err){
    if(msgEl)msgEl.textContent='Falha ao iniciar a recuperação.';
    if(detailEl)detailEl.textContent=String(err&&err.message||err);
    alert('Erro ao recuperar e-mails: '+(err&&err.message||err));
    await refresh();
  }
}
async function recoverMissing(){return recoverEmails();}
function _fmtDateTime(v){
  if(!v)return '-';
  try{
    const d=new Date(v);
    if(Number.isNaN(d.getTime()))return String(v)==='-'?'-':String(v);
    return d.toLocaleString('pt-BR');
  }catch(_){
    return String(v);
  }
}
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
let _auditDeleteQueue=new Map();
let _auditDeleteTimer=null;
let _auditDeleteInFlight=false;
const _auditDeleteDelayMs=3000;
let _auditItems=[];
let _auditSort={key:'',dir:'asc'};
let _auditTable=null;
let _historyTable=null;
let _watchTable=null;
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
  if(_historyHasTabulator()){
    const host=document.getElementById('histTableTabulator');
    try{if(_historyTable)_historyTable.destroy();}catch(_){}
    if(host)host.innerHTML='';
    _historyTable=null;
    _ensureHistoryTabulator(_prepareHistoryRows(_histItems));
    _toggleHistoryRenderMode(true);
    _syncHistoryToneMode();
    return;
  }
  _histColWidths={..._histColDefaults};
  _applyHistColWidths();
  _saveHistColWidths();
}
function _initHistoryColumnResize(){
  if(_historyHasTabulator())return;
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
function _historyHasTabulator(){return typeof window.Tabulator==='function';}
function _getHistoryEmpresa(){
  const group=document.getElementById('historyEmpresaGroup');
  return _normalizeAuditEmpresaClient(group&&group.dataset&&group.dataset.value||'mva');
}
function _syncHistoryEmpresaButtons(){
  const active=_getHistoryEmpresa();
  document.querySelectorAll('#historyEmpresaGroup .audit-source-btn[data-empresa]').forEach((btn)=>{
    btn.classList.toggle('active',String(btn.dataset.empresa||'')===active);
  });
  _syncHistoryToneMode();
}
function _syncHistoryToneMode(){
  const toneClass=`audit-tone-${_getHistoryEmpresa()}`;
  ['audit-tone-mva','audit-tone-eh','audit-tone-todos'].forEach((cls)=>{
    const legacy=document.getElementById('histTableLegacy');
    const host=document.getElementById('histTableTabulator');
    if(legacy)legacy.classList.remove(cls);
    if(host)host.classList.remove(cls);
  });
  const legacy=document.getElementById('histTableLegacy');
  const host=document.getElementById('histTableTabulator');
  if(legacy)legacy.classList.add(toneClass);
  if(host)host.classList.add(toneClass);
}
function setHistoryEmpresa(value,reload=true){
  const group=document.getElementById('historyEmpresaGroup');
  if(!group)return;
  const next=_normalizeAuditEmpresaClient(value);
  const prev=_getHistoryEmpresa();
  group.dataset.value=next;
  _syncHistoryEmpresaButtons();
  if(reload!==false&&prev!==next&&_activeTab==='hist'){
    loadHistory(false).catch(()=>{});
  }
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
function _inferEmpresaKeyFromValue(empresa,local,nf){
  const explicit=_normalizeAuditEmpresaClient(empresa||'todos');
  if(explicit==='mva'||explicit==='eh')return explicit;
  const scope=String(local||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase();
  if(/(^|\b)MVA(\b|$)/.test(scope))return 'mva';
  if(/(^|\b)EH(\b|$)/.test(scope))return 'eh';
  const nfNum=Number(String(nf||'').replace(/\\D+/g,''))||0;
  if(nfNum>=40000)return 'mva';
  if(nfNum>=19000&&nfNum<40000)return 'eh';
  return '';
}
function _toneClassForEmpresa(activeEmpresa,empresaKey){
  const active=_normalizeAuditEmpresaClient(activeEmpresa||'todos');
  const key=String(empresaKey||'').trim().toLowerCase();
  if(active!=='todos')return '';
  if(key==='mva')return 'audit-row-tone-mva';
  if(key==='eh')return 'audit-row-tone-eh';
  return '';
}
function _toggleHistoryRenderMode(useTabulator){
  const host=document.getElementById('histTableTabulator');
  const legacy=document.getElementById('histTableLegacy');
  if(host)host.classList.toggle('active',!!useTabulator);
  if(legacy)legacy.style.display=useTabulator?'none':'table';
}
function _prepareHistoryRows(items){
  let arr=(Array.isArray(items)?items:[]).filter((it)=>it&&it.type==='boleto_lancado');
  arr=arr.map((it)=>{
    const nf=String(it.nf||it.numero||'').trim();
    const local=_fmtLocal(it.local_lancamento);
    const empresaKey=_inferEmpresaKeyFromValue(it.empresa,it.local_lancamento||local,nf);
    return {
      ...it,
      _nf_doc:nf?`NF ${nf}`:'-',
      _local_view:local,
      _cliente_view:_compactClienteLabel(it.cliente,it.descricao),
      _cliente_title:_compactSpaces(it.descricao||it.cliente||'-'),
      _empresa_key:empresaKey,
      _tone_class:_toneClassForEmpresa(_getHistoryEmpresa(),empresaKey),
    };
  });
  return _historyHasTabulator()?arr:_sortHist(arr);
}
function _historyClienteCellFormatter(cell){
  const data=cell.getRow().getData()||{};
  const emit=String(data.cnpj_emit||'-');
  const title=_esc(String(data._cliente_title||'-'));
  return `<div class="cell-menu"><button class="cell-btn" title="${title}" onclick="_toggleMenu(event,this)">${_esc(data._cliente_view||'-')}</button><div class="cell-pop"><button data-cnpj="${_esc(emit)}" onclick="_showCnpj(event,this)">Copiar CNPJ emitente</button></div></div>`;
}
function _historyDocFormatter(cell){
  const data=cell.getRow().getData()||{};
  const base=_esc(String(data._nf_doc||'-'));
  return `${base}${data.duplicata?'<span class="dup-badge">DUPLICADA</span>':''}`;
}
function _historyColumns(){
  return [
    {title:'Data',field:'at',hozAlign:'center',headerHozAlign:'center',width:160,sorter:'string',formatter:(cell)=>_fmtDateTime(cell.getValue()||'-'),headerFilter:'input'},
    {title:'Venc.',field:'vencimento',hozAlign:'center',headerHozAlign:'center',width:112,sorter:(a,b)=>_auditDateSortValue(a)-_auditDateSortValue(b),formatter:(cell)=>_fmtAuditDate(cell.getValue()||'-')||'-',headerFilter:'input'},
    {title:'NF',field:'nf',hozAlign:'center',headerHozAlign:'center',width:110,sorter:'number',headerFilter:'input',formatter:_historyDocFormatter},
    {title:'Cliente',field:'_cliente_view',hozAlign:'center',headerHozAlign:'center',minWidth:280,widthGrow:4,headerFilter:'input',formatter:_historyClienteCellFormatter},
    {title:'Parc.',field:'parcela',hozAlign:'center',headerHozAlign:'center',width:96,headerFilter:'input'},
    {title:'Parcela',field:'valor_parcela',hozAlign:'center',headerHozAlign:'center',width:132,sorter:'number',formatter:(cell)=>_fmtMoney(cell.getValue())},
    {title:'Total',field:'valor_total',hozAlign:'center',headerHozAlign:'center',width:132,sorter:'number',formatter:(cell)=>_fmtMoney(cell.getValue())},
    {title:'Aba',field:'_local_view',hozAlign:'center',headerHozAlign:'center',minWidth:130,widthGrow:1,headerFilter:'input'},
  ];
}
function _ensureHistoryTabulator(initialRows){
  if(!_historyHasTabulator())return null;
  if(_historyTable)return _historyTable;
  _historyTable=new Tabulator('#histTableTabulator',{
    data:Array.isArray(initialRows)?initialRows:[],
    layout:'fitDataStretch',
    responsiveLayout:false,
    pagination:'local',
    paginationSize:14,
    paginationCounter:'rows',
    movableColumns:true,
    resizableColumns:true,
    placeholder:'Sem dados para os filtros selecionados',
    columns:_historyColumns(),
    rowFormatter:function(row){
      const el=row.getElement();
      const data=row.getData()||{};
      el.classList.remove('dup-row','audit-row-tone-mva','audit-row-tone-eh');
      if(data._tone_class)el.classList.add(String(data._tone_class));
      if(data.duplicata)el.classList.add('dup-row');
      el.title=String(data._cliente_title||'').trim();
    },
  });
  return _historyTable;
}
function _getSortValue(it,key){
  if(key==='at')return it.at||'';
  if(key==='venc')return it.vencimento||'';
  if(key==='doc')return it._nf_doc||it._doc||'';
  if(key==='cliente')return it._cliente_view||it.descricao||it.cliente||'';
  if(key==='parcela')return it.parcela||'';
  if(key==='vparcela')return Number(it.valor_parcela||0);
  if(key==='vtotal')return Number(it.valor_total||0);
  if(key==='local')return it._local_view||it._local||'';
  return '';
}
function _sortHist(items){const k=_histSort.key;const dir=_histSort.dir==='asc'?1:-1;return [...items].sort((a,b)=>{const va=_getSortValue(a,k);const vb=_getSortValue(b,k);if(va<vb)return -1*dir;if(va>vb)return 1*dir;return 0;});}
function _renderHistory(items){
  _histItems=Array.isArray(items)?items:[];
  const arr=_prepareHistoryRows(_histItems);
  _syncHistoryToneMode();
  if(_historyHasTabulator()){
    const table=_ensureHistoryTabulator(arr);
    if(table){
      _toggleHistoryRenderMode(true);
      try{table.setData(arr);}catch(_){}
      return;
    }
  }
  _toggleHistoryRenderMode(false);
  const body=document.getElementById('hBody');
  if(!body)return;
  body.innerHTML='';
  if(!arr.length){body.innerHTML='<tr><td colspan="8">Sem dados para os filtros selecionados</td></tr>';return;}
  arr.forEach(it=>{
    const tr=document.createElement('tr');
    if(it._tone_class)tr.classList.add(String(it._tone_class));
    if(it.duplicata)tr.classList.add('dup-row');
    const emit=String(it.cnpj_emit||'-');
    const dupTag=it.duplicata?'<span class="dup-badge">DUPLICADA</span>':'';
    const menu=`<div class=\"cell-menu\"><button class=\"cell-btn\" title=\"${_esc(it._cliente_title||'-')}\" onclick=\"_toggleMenu(event,this)\">${_esc(it._cliente_view||'-')}</button><div class=\"cell-pop\"><button data-cnpj=\"${_esc(emit)}\" onclick=\"_showCnpj(event,this)\">Copiar CNPJ emitente</button></div></div>`;
    tr.innerHTML=`<td title=\"${_esc(_fmtDateTime(it.at))}\">${_fmtDateTime(it.at)}</td><td title=\"${_esc(_fmtAuditDate(it.vencimento||'-'))}\">${_esc(_fmtAuditDate(it.vencimento||'-'))}</td><td title=\"${_esc(it._nf_doc)}\">${_esc(it._nf_doc)}${dupTag}</td><td>${menu}</td><td title=\"${_esc(it.parcela||'-')}\">${_esc(it.parcela||'-')}</td><td title=\"${_esc(_fmtMoney(it.valor_parcela))}\">${_fmtMoney(it.valor_parcela)}</td><td title=\"${_esc(_fmtMoney(it.valor_total))}\">${_fmtMoney(it.valor_total)}</td><td title=\"${_esc(it._local_view)}\">${_esc(it._local_view)}</td>`;
    body.appendChild(tr);
  });
}
function _setSort(key){
  if(_historyHasTabulator())return;
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
_bindAuditSortHeaders();
function _historyParams(){
  const p=new URLSearchParams();
  const vAt=((document.getElementById('hAt')||{}).value||'').trim();
  const vVenc=((document.getElementById('hVenc')||{}).value||'').trim();
  const vNf=((document.getElementById('hNf')||{}).value||'').trim();
  const vCliente=((document.getElementById('hCliente')||{}).value||'').trim();
  const vAba=((document.getElementById('hAba')||{}).value||'').trim();
  const vLimit=Number((document.getElementById('hLimit')||{}).value||300);
  const vEmpresa=_getHistoryEmpresa();
  if(vAt)p.set('at',vAt);
  if(vVenc)p.set('venc',vVenc);
  if(vNf)p.set('nf',vNf);
  if(vCliente)p.set('cliente',vCliente);
  if(vAba)p.set('aba',vAba);
  p.set('empresa',vEmpresa);
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
    if(_historyHasTabulator()&&_historyTable){
      try{_historyTable.setData([]);}catch(_){}
      _toggleHistoryRenderMode(true);
      return;
    }
    const body=document.getElementById('hBody');
    if(body)body.innerHTML='<tr><td colspan="8">Erro de rede: '+_esc(String(err&&err.message||err))+'</td></tr>';
  }
}
let _auditJobPollTimer=null;
let _auditActiveRequestKey='';
let _auditActiveJobId='';
function _normalizeAuditEmpresaClient(value){
  const key=String(value||'todos').trim().toLowerCase();
  if(key==='mva'||key==='eh')return key;
  return 'todos';
}
function _getAuditEmpresa(){
  const group=document.getElementById('auditEmpresaGroup');
  return _normalizeAuditEmpresaClient(group&&group.dataset&&group.dataset.value||'mva');
}
function _syncAuditEmpresaButtons(){
  const active=_getAuditEmpresa();
  document.querySelectorAll('.audit-source-btn[data-empresa]').forEach((btn)=>{
    btn.classList.toggle('active',String(btn.dataset.empresa||'')===active);
  });
  _syncAuditToneMode();
}
function setAuditEmpresa(value,reload=true){
  const group=document.getElementById('auditEmpresaGroup');
  if(!group)return;
  const next=_normalizeAuditEmpresaClient(value);
  const prev=_getAuditEmpresa();
  group.dataset.value=next;
  _syncAuditEmpresaButtons();
  if(reload!==false&&prev!==next&&_activeTab==='audit'){
    loadParcelAudit(false).catch(()=>{});
  }
}
function _readAuditUiState(){
  const mode=((document.getElementById('aMode')||{}).value||'mes').trim()||'mes';
  const monthValue=((document.getElementById('aMonth')||{}).value||'').trim();
  const nfStart=((document.getElementById('aNfStart')||{}).value||'').trim();
  const nfEnd=((document.getElementById('aNfEnd')||{}).value||'').trim();
  return {mode,monthValue,nfStart,nfEnd,empresa:_getAuditEmpresa()};
}
function _buildAuditQuery(state,empresa){
  const p=new URLSearchParams();
  const mode=String(state&&state.mode||'mes').trim()||'mes';
  const monthValue=String(state&&state.monthValue||'').trim();
  const nfStart=String(state&&state.nfStart||'').trim();
  const nfEnd=String(state&&state.nfEnd||'').trim();
  const company=_normalizeAuditEmpresaClient(empresa||state&&state.empresa||'todos');
  p.set('filtro',mode);
  if(mode==='mes'&&monthValue)p.set('mes',monthValue);
  if(mode==='nfs'&&nfStart)p.set('nf_inicio',nfStart);
  if(mode==='nfs'&&nfEnd)p.set('nf_fim',nfEnd);
  p.set('empresa',company);
  return p;
}
function _auditFetchKey(state,empresa){
  return JSON.stringify({
    filtro:String(state&&state.mode||'mes'),
    mes:String(state&&state.monthValue||''),
    nf_inicio:String(state&&state.nfStart||''),
    nf_fim:String(state&&state.nfEnd||''),
    empresa:_normalizeAuditEmpresaClient(empresa||state&&state.empresa||'todos'),
  });
}
function _stopAuditJobPolling(){
  if(_auditJobPollTimer){
    clearTimeout(_auditJobPollTimer);
    _auditJobPollTimer=null;
  }
}
function _clearAuditFetchCache(){
  _stopAuditJobPolling();
  _auditActiveJobId='';
  _auditActiveRequestKey='';
}
function _buildAuditLoadedMessage(meta,prefix){
  const info=(meta&&typeof meta==='object')?meta:{};
  const loadedAt=String(info.loaded_at||'').trim();
  const linhas=Number(info.linhas_lidas||0);
  const abas=Number(info.abas_lidas||0);
  const head=loadedAt?`${String(prefix||'Conferência atualizada.')} em ${_fmtDateTime(loadedAt)}`:String(prefix||'Conferência atualizada.');
  return `${head} | ${linhas} linhas lidas em ${abas} abas`;
}
function _applyAuditPayload(payload){
  const data=(payload&&typeof payload==='object')?payload:{};
  const result=(data.current_result&&typeof data.current_result==='object')?data.current_result:((data.result&&typeof data.result==='object')?data.result:null);
  if(!result)return false;
  _setAuditSummary(result.summary||{});
  _renderParcelAudit(result.items||[]);
  return true;
}
function _scheduleAuditJobPoll(jobId,requestKey,delay=900){
  _stopAuditJobPolling();
  if(!jobId||requestKey!==_auditActiveRequestKey)return;
  _auditJobPollTimer=window.setTimeout(()=>{
    _auditJobPollTimer=null;
    _pollAuditJob(jobId,requestKey).catch((err)=>console.warn('Falha ao acompanhar job da conferência:',err));
  },delay);
}
async function _pollAuditJob(jobId,requestKey){
  if(!jobId||requestKey!==_auditActiveRequestKey)return;
  const data=await api('/api/conferencia-parcelas/job?id='+encodeURIComponent(jobId));
  if(requestKey!==_auditActiveRequestKey)return;
  const rendered=_applyAuditPayload(data);
  const message=String(data&&data.message||'Atualizando conferência...').trim()||'Atualizando conferência...';
  if(String(data&&data.status||'').trim()==='done'){
    _auditActiveJobId='';
    _stopAuditJobPolling();
    _setAuditLoading(false, rendered ? _buildAuditLoadedMessage((data.result||{}).meta||{},'Conferência atualizada') : message);
    return;
  }
  if(String(data&&data.status||'').trim()==='error'){
    _auditActiveJobId='';
    _stopAuditJobPolling();
    _setAuditLoading(false, message || 'Falha ao conferir as planilhas.');
    if(!rendered){
      _setAuditSummary({});
      if(_auditHasTabulator()&&_auditTable){
        _auditTable.setData([]);
        _toggleAuditRenderMode(true);
      }
      const body=document.getElementById('aBody');
      if(body)body.innerHTML='<tr><td colspan="8">Erro de rede: '+_esc(message)+'</td></tr>';
    }
    return;
  }
  if(rendered)_setAuditLoading(false, message);
  else _setAuditLoading(true, message);
  _scheduleAuditJobPoll(jobId,requestKey, rendered ? 1200 : 700);
}
function toggleAuditFilters(){
  const mode=((document.getElementById('aMode')||{}).value||'mes').trim();
  const monthEl=document.getElementById('aMonth');
  const startEl=document.getElementById('aNfStart');
  const endEl=document.getElementById('aNfEnd');
  const monthWrap=document.getElementById('auditMonthField');
  const startWrap=document.getElementById('auditNfStartField');
  const endWrap=document.getElementById('auditNfEndField');
  const useMonth=mode==='mes';
  const useNf=mode==='nfs';
  if(monthEl)monthEl.disabled=!useMonth;
  if(startEl)startEl.disabled=!useNf;
  if(endEl)endEl.disabled=!useNf;
  if(monthWrap)monthWrap.classList.toggle('hidden',!useMonth);
  if(startWrap)startWrap.classList.toggle('hidden',!useNf);
  if(endWrap)endWrap.classList.toggle('hidden',!useNf);
}
function _focusAuditField(id){
  const el=document.getElementById(id);
  if(!el||el.disabled||el.classList.contains('hidden')||el.closest('.hidden'))return false;
  try{el.focus();}catch(_){return false;}
  return true;
}
function handleAuditFieldKeydown(ev){
  if(ev.key!=='Enter')return;
  ev.preventDefault();
  const mode=String((document.getElementById('aMode')||{}).value||'mes').trim();
  const currentId=String((ev&&ev.target&&ev.target.id)||'').trim();
  if(currentId==='aMode'){
    if(mode==='nfs'){
      if(_focusAuditField('aNfStart'))return;
    }else if(mode==='mes'){
      if(_focusAuditField('aMonth'))return;
    }
    loadParcelAudit();
    return;
  }
  if(mode==='mes'){
    loadParcelAudit();
    return;
  }
  if(mode==='nfs'){
    if(currentId==='aNfStart'&&_focusAuditField('aNfEnd'))return;
    loadParcelAudit();
    return;
  }
  loadParcelAudit();
}
function _fmtAuditList(values){
  const arr=(Array.isArray(values)?values:[]).map(v=>_compactSpaces(v)).filter(Boolean);
  return arr.length?arr.join(', '):'-';
}
function _fmtAuditDate(v){
  const txt=String(v||'').trim();
  if(!txt||txt==='-'||txt.toLowerCase()==='invalid date')return '-';
  if(/^\\d{4}-\\d{2}-\\d{2}$/.test(txt)){
    const [y,m,d]=txt.split('-');
    return `${d}/${m}/${y}`;
  }
  if(/^\\d{2}\\/\\d{2}\\/\\d{4}$/.test(txt))return txt;
  return _fmtDateTime(txt);
}
function _auditHasTabulator(){return typeof window.Tabulator==='function';}
function _toggleAuditRenderMode(useTabulator){
  const host=document.getElementById('auditTableTabulator');
  const legacy=document.getElementById('auditTableLegacy');
  if(host)host.classList.toggle('active',!!useTabulator);
  if(legacy)legacy.style.display=useTabulator?'none':'table';
}
function _auditResolveRowTarget(target){
  if(target&&typeof target.getData==='function')return target;
  if(target&&target.row&&typeof target.row.getData==='function')return target.row;
  return null;
}
function _auditInferEmpresaKey(it){
  const explicit=_normalizeAuditEmpresaClient((it&&(
    it.empresa||it.sheet_type||it.empresa_key
  ))||'todos');
  if(explicit==='mva'||explicit==='eh')return explicit;
  const local=String((it&&(it.local_lancamento||it.aba||it._local_view||it.local_view))||'').toUpperCase();
  if(/(^|\\b)MVA(\\b|$)/.test(local))return 'mva';
  if(/(^|\\b)EH(\\b|$)/.test(local))return 'eh';
  const nfNum=Number(it&&it.nf||0);
  if(nfNum>=40000)return 'mva';
  if(nfNum>=19000&&nfNum<40000)return 'eh';
  return '';
}
function _auditToneClassForEmpresa(empresaKey){
  const active=_getAuditEmpresa();
  const key=String(empresaKey||'').trim().toLowerCase();
  if(active!=='todos')return '';
  if(key==='mva')return 'audit-row-tone-mva';
  if(key==='eh')return 'audit-row-tone-eh';
  return '';
}
function _syncAuditToneMode(){
  const toneClass=`audit-tone-${_getAuditEmpresa()}`;
  ['audit-tone-mva','audit-tone-eh','audit-tone-todos'].forEach((cls)=>{
    const legacy=document.getElementById('auditTableLegacy');
    const host=document.getElementById('auditTableTabulator');
    if(legacy)legacy.classList.remove(cls);
    if(host)host.classList.remove(cls);
  });
  const legacy=document.getElementById('auditTableLegacy');
  const host=document.getElementById('auditTableTabulator');
  if(legacy)legacy.classList.add(toneClass);
  if(host)host.classList.add(toneClass);
}
function _mapAuditRow(it){
  const reasonHint=_compactSpaces(it.reason_hint||'');
  const local=_fmtLocal(it.local_lancamento||it.aba||'-');
  const ultimoVenc=_fmtAuditDate(it.ultimo_vencimento||it.ultimo_lancamento);
  const duplicadasTxt=Number(it.qtd_duplicada||0)>0?_fmtAuditList(it.parcelas_duplicadas):'0';
  const deleteCandidates=Number(it.delete_candidates||0);
  const esperadas=Number(it.qtd_esperada||0);
  const lancadas=Number(it.qtd_lancada||0);
  const empresaKey=_auditInferEmpresaKey(it);
  return Object.assign({},it,{
    _nf_num:Number(it.nf||0),
    _status_rank:_auditStatusRank(it.status||''),
    _cliente_view:_compactClienteLabel(it.cliente,it.descricao),
    _parcelas_view:`${esperadas} / ${lancadas}`,
    _duplicadas_view:duplicadasTxt,
    _local_view:local,
    _ultimo_venc_view:ultimoVenc,
    _ultimo_venc_sort:_auditDateSortValue(it.ultimo_vencimento||it.ultimo_lancamento||''),
    _status_title:reasonHint || (deleteCandidates>0 ? 'Clique para limpar linhas excedentes/duplicadas desta NF direto na planilha' : ''),
    _delete_enabled:!!(it.status&&it.status!=='ok'&&deleteCandidates>0),
    _empresa_key:empresaKey,
    _tone_class:_auditToneClassForEmpresa(empresaKey),
    _local_pending:false,
    _local_removed:false,
  });
}
function _auditStatusCellFormatter(cell){
  const data=cell.getRow().getData()||{};
  const label=String(data.status_label||'-');
  const title=String(data._status_title||'').trim();
  if(data._delete_enabled){
    const btn=document.createElement('button');
    btn.type='button';
    btn.className=`audit-status audit-status-btn ${data.status||'ok'}`;
    btn.textContent=label;
    btn.disabled=!!data._local_removed||!!data._local_pending;
    btn.title=title||'Clique para limpar linhas excedentes/duplicadas desta NF direto na planilha';
    btn.addEventListener('click',async(ev)=>{
      ev.preventDefault();
      ev.stopPropagation();
      await deleteAuditRows({row:cell.getRow()},data.audit_key,data.nf,data.status_label,data.delete_candidates);
    });
    return btn;
  }
  const span=document.createElement('span');
  span.className=`audit-status ${data.status||'ok'}`;
  span.textContent=label;
  if(title)span.title=title;
  return span;
}
function _captureAuditViewport(){
  return {
    x: window.scrollX||window.pageXOffset||0,
    y: window.scrollY||window.pageYOffset||0,
  };
}
function _restoreAuditViewport(viewport){
  if(!viewport)return;
  const x=Number(viewport.x||0);
  const y=Number(viewport.y||0);
  const apply=()=>{
    try{window.scrollTo({left:x,top:y,behavior:'auto'});}
    catch(_){window.scrollTo(x,y);}
  };
  apply();
  try{window.requestAnimationFrame(()=>window.requestAnimationFrame(apply));}
  catch(_){setTimeout(apply,0);}
}
async function _setAuditTabulatorData(table,rows){
  if(!table)return;
  const nextRows=Array.isArray(rows)?rows:[];
  const viewport=_captureAuditViewport();
  try{
    if(typeof table.clearHeaderFilter==='function')table.clearHeaderFilter();
  }catch(_){}
  try{
    if(typeof table.clearSort==='function')table.clearSort();
  }catch(_){}
  await table.setData(nextRows);
  try{
    if(typeof table.setPage==='function')table.setPage(1);
  }catch(_){}
  try{
    table.redraw(true);
  }catch(_){}
  _restoreAuditViewport(viewport);
}
function _ensureAuditTabulator(initialRows){
  if(!_auditHasTabulator())return null;
  if(_auditTable)return _auditTable;
  const seedRows=Array.isArray(initialRows)?initialRows:[];
  _auditTable=new Tabulator('#auditTableTabulator',{
    data:seedRows,
    layout:'fitDataStretch',
    responsiveLayout:false,
    pagination:'local',
    paginationSize:14,
    paginationCounter:'rows',
    movableColumns:true,
    resizableColumns:true,
    placeholder:'Nenhuma NF encontrada para os filtros selecionados',
    columns:[
      {title:'Status',field:'status_label',hozAlign:'center',headerHozAlign:'center',width:112,sorter:function(a,b,aRow,bRow){return (aRow.getData()._status_rank||0)-(bRow.getData()._status_rank||0);},formatter:_auditStatusCellFormatter},
      {title:'NF',field:'nf',hozAlign:'center',headerHozAlign:'center',width:90,sorter:'number',headerFilter:'input'},
      {title:'Cliente',field:'_cliente_view',hozAlign:'center',headerHozAlign:'center',minWidth:360,widthGrow:5,headerFilter:'input'},
      {title:'Parc.',field:'_parcelas_view',hozAlign:'center',headerHozAlign:'center',width:104,sorter:function(a,b,aRow,bRow){
        const aData=aRow.getData()||{};
        const bData=bRow.getData()||{};
        const esperadasDiff=Number(aData.qtd_esperada||0)-Number(bData.qtd_esperada||0);
        if(esperadasDiff!==0)return esperadasDiff;
        return Number(aData.qtd_lancada||0)-Number(bData.qtd_lancada||0);
      }},
      {title:'Faltando',field:'qtd_faltando',hozAlign:'center',headerHozAlign:'center',width:96,sorter:'number'},
      {title:'Duplicadas',field:'qtd_duplicada',hozAlign:'center',headerHozAlign:'center',width:112,sorter:'number',formatter:function(cell){
        const data=cell.getRow().getData()||{};
        const count=Number(data.qtd_duplicada||0);
        const txt=String(data._duplicadas_view||'').trim();
        if(count>0&&txt&&txt!=='0'&&txt!=='-')return `<span title="${_esc(txt)}">${count} - ${_esc(txt)}</span>`;
        return String(count||0);
      }},
      {title:'Últ. venc.',field:'_ultimo_venc_view',hozAlign:'center',headerHozAlign:'center',width:118,sorter:function(a,b,aRow,bRow){return (aRow.getData()._ultimo_venc_sort||0)-(bRow.getData()._ultimo_venc_sort||0);}},
      {title:'Aba',field:'_local_view',hozAlign:'center',headerHozAlign:'center',minWidth:120,widthGrow:1,headerFilter:'input'},
    ],
    rowFormatter:function(row){
      const el=row.getElement();
      const data=row.getData()||{};
      el.classList.remove('audit-row-aviso','audit-row-erro','audit-row-local-pending','audit-row-local-removed','audit-row-tone-mva','audit-row-tone-eh');
      if(data._tone_class)el.classList.add(String(data._tone_class));
      if(data._local_removed)el.classList.add('audit-row-local-removed');
      else if(data._local_pending)el.classList.add('audit-row-local-pending');
      else if(data.status==='erro')el.classList.add('audit-row-erro');
      else if(data.status==='aviso')el.classList.add('audit-row-aviso');
      const reason=String(data.reason_hint||'').trim();
      if(reason)el.title=reason;
    },
  });
  return _auditTable;
}
function _initAuditSortHeaders(){
  if(_auditHasTabulator())return;
  const keys=['status','nf','cliente','parcelas','faltando','duplicada','vencimento','local'];
  const ths=document.querySelectorAll('.audit-table thead th');
  ths.forEach((th,idx)=>{
    const key=keys[idx];
    if(!key)return;
    th.classList.add('sortable');
    th.dataset.key=key;
  });
}
function _auditStatusRank(status){
  const txt=String(status||'').trim().toLowerCase();
  if(txt==='erro')return 0;
  if(txt==='aviso')return 1;
  return 2;
}
function _auditDateSortValue(value){
  const txt=String(value||'').trim();
  if(!txt||txt==='-')return 0;
  if(/^\\d{4}-\\d{2}-\\d{2}$/.test(txt))return Date.parse(`${txt}T00:00:00`)||0;
  if(/^\\d{2}\\/\\d{2}\\/\\d{4}$/.test(txt)){
    const [d,m,y]=txt.split('/');
    return Date.parse(`${y}-${m}-${d}T00:00:00`)||0;
  }
  const ms=Date.parse(txt);
  return Number.isFinite(ms)?ms:0;
}
function _getAuditSortValue(it,key){
  if(key==='status')return _auditStatusRank(it.status||'');
  if(key==='nf')return Number(it.nf||0);
  if(key==='cliente')return _compactSpaces(it.cliente||it.descricao||'');
  if(key==='parcelas')return (Number(it.qtd_esperada||0)*1000)+Number(it.qtd_lancada||0);
  if(key==='faltando')return Number(it.qtd_faltando||0);
  if(key==='duplicada')return Number(it.qtd_duplicada||0);
  if(key==='vencimento')return _auditDateSortValue(it.ultimo_vencimento||it.ultimo_lancamento||'');
  if(key==='local')return _compactSpaces(it.local_lancamento||it.aba||'');
  return '';
}
function _sortAudit(items){
  const arr=[...(Array.isArray(items)?items:[])];
  const key=String(_auditSort.key||'').trim();
  if(!key)return arr;
  const dir=_auditSort.dir==='desc'?-1:1;
  return arr.sort((a,b)=>{
    const va=_getAuditSortValue(a,key);
    const vb=_getAuditSortValue(b,key);
    if(va<vb)return -1*dir;
    if(va>vb)return 1*dir;
    const na=Number(a&&a.nf||0);
    const nb=Number(b&&b.nf||0);
    if(na<nb)return 1;
    if(na>nb)return -1;
    return 0;
  });
}
function _setAuditSort(key){
  if(_auditHasTabulator())return;
  const next=String(key||'').trim();
  if(!next)return;
  const ths=document.querySelectorAll('.audit-table th.sortable');
  ths.forEach(th=>{th.classList.remove('asc');th.classList.remove('desc');});
  if(_auditSort.key===next)_auditSort.dir=_auditSort.dir==='asc'?'desc':'asc';
  else{_auditSort.key=next;_auditSort.dir='asc';}
  const active=document.querySelector(`.audit-table th.sortable[data-key="${next}"]`);
  if(active)active.classList.add(_auditSort.dir);
  _renderParcelAudit(_auditItems);
}
function _bindAuditSortHeaders(){
  if(_auditHasTabulator())return;
  _initAuditSortHeaders();
  document.querySelectorAll('.audit-table th.sortable').forEach(th=>{
    if(th.dataset.sortReady==='1')return;
    th.dataset.sortReady='1';
    th.addEventListener('click',()=>_setAuditSort(th.dataset.key||''));
  });
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
  _auditItems=Array.isArray(items)?items:[];
  const viewport=_captureAuditViewport();
  if(_auditHasTabulator()){
    const rows=_auditItems.map(_mapAuditRow);
    const isFirstBuild=!_auditTable;
    const table=_ensureAuditTabulator(rows);
    _toggleAuditRenderMode(!!table);
    if(table){
      if(isFirstBuild)_restoreAuditViewport(viewport);
      if(!isFirstBuild){
        _setAuditTabulatorData(table,rows).catch((err)=>{
          console.warn('Falha ao atualizar grade da conferência:',err);
        });
      }
      return;
    }
  }
  _toggleAuditRenderMode(false);
  const body=document.getElementById('aBody');
  if(!body)return;
  const arr=_sortAudit(_auditItems);
  body.innerHTML='';
  if(!arr.length){
    body.innerHTML='<tr><td colspan="8">Nenhuma NF encontrada para os filtros selecionados</td></tr>';
    _restoreAuditViewport(viewport);
    return;
  }
  arr.forEach(it=>{
    const tr=document.createElement('tr');
    const toneClass=_auditToneClassForEmpresa(_auditInferEmpresaKey(it));
    if(toneClass)tr.classList.add(toneClass);
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
    const reasonHint=_compactSpaces(it.reason_hint||'');
    const statusTitle=reasonHint || (deleteCandidates>0 ? 'Clique para limpar linhas excedentes/duplicadas desta NF direto na planilha' : '');
    if(reasonHint)tr.title=reasonHint;
    const statusCell=(it.status&&it.status!=='ok'&&deleteCandidates>0)
      ? `<button type="button" class="audit-status audit-status-btn ${_esc(it.status||'ok')}" title="${_esc(statusTitle)}" onclick="deleteAuditRows(this,'${auditKey}','${nfValue}','${statusLabel}',${deleteCandidates})">${statusLabel}</button>`
      : `<span class="audit-status ${_esc(it.status||'ok')}"${statusTitle?` title="${_esc(statusTitle)}"`:''}>${statusLabel}</span>`;
    tr.innerHTML=`<td>${statusCell}</td><td title="${_esc(reasonHint||nfValue)}">${nfValue}</td><td title="${_esc(reasonHint||_compactSpaces(it.descricao||it.cliente||'-'))}">${_esc(clienteView)}</td><td>${_esc(`${Number(it.qtd_esperada||0)} / ${Number(it.qtd_lancada||0)}`)}</td><td>${_esc(String(it.qtd_faltando||0))}</td><td title="${_esc(duplicadasTxt)}">${_esc(String(it.qtd_duplicada||0))}${Number(it.qtd_duplicada||0)>0?` - ${_esc(duplicadasTxt)}`:''}</td><td title="${_esc(ultimoVenc)}">${_esc(ultimoVenc)}</td><td class="audit-cell-local" title="${_esc(reasonHint||local)}">${_esc(local)}</td>`;
    body.appendChild(tr);
  });
  _restoreAuditViewport(viewport);
}
function _markAuditRowDeletedLocal(btn,message){
  const auditRow=_auditResolveRowTarget(btn);
  if(auditRow){
    const data=auditRow.getData&&auditRow.getData()||{};
    try{
      auditRow.update({
        _local_pending:false,
        _local_removed:true,
        _delete_enabled:false,
        _status_title:String(message||'Linhas excedentes/duplicadas já limpas desta NF.'),
      });
      if(typeof auditRow.reformat==='function')auditRow.reformat();
      else if(typeof auditRow.normalizeHeight==='function')auditRow.normalizeHeight();
    }catch(_){}
    return;
  }
  const row=btn&&btn.closest?btn.closest('tr'):null;
  if(row){
    row.classList.remove('audit-row-erro','audit-row-aviso','audit-row-local-pending');
    row.classList.add('audit-row-local-removed');
  }
  if(btn){
    btn.disabled=true;
    btn.title=String(message||'Linhas excedentes/duplicadas já limpas desta NF.');
  }
}
function _markAuditRowPendingLocal(btn,message){
  const auditRow=_auditResolveRowTarget(btn);
  if(auditRow){
    try{
      auditRow.update({
        _local_removed:false,
        _local_pending:true,
        _delete_enabled:true,
        _status_title:String(message||'Limpeza em fila para envio em lote.'),
      });
      if(typeof auditRow.reformat==='function')auditRow.reformat();
      else if(typeof auditRow.normalizeHeight==='function')auditRow.normalizeHeight();
    }catch(_){}
    return;
  }
  const row=btn&&btn.closest?btn.closest('tr'):null;
  if(row){
    row.classList.remove('audit-row-local-removed');
    row.classList.add('audit-row-local-pending');
  }
  if(btn){
    btn.disabled=true;
    btn.title=String(message||'Limpeza em fila para envio em lote.');
  }
}
function _restoreAuditRowDeleteLocal(btn,message){
  const auditRow=_auditResolveRowTarget(btn);
  if(auditRow){
    const data=auditRow.getData&&auditRow.getData()||{};
    const reasonHint=_compactSpaces(data.reason_hint||'');
    const defaultTitle=reasonHint || 'Clique para limpar linhas excedentes/duplicadas desta NF direto na planilha';
    try{
      auditRow.update({
        _local_pending:false,
        _local_removed:false,
        _delete_enabled:!!(data.status&&data.status!=='ok'&&Number(data.delete_candidates||0)>0),
        _status_title:String(message||defaultTitle),
      });
      if(typeof auditRow.reformat==='function')auditRow.reformat();
      else if(typeof auditRow.normalizeHeight==='function')auditRow.normalizeHeight();
    }catch(_){}
    return;
  }
  const row=btn&&btn.closest?btn.closest('tr'):null;
  if(row){
    row.classList.remove('audit-row-local-pending');
  }
  if(btn){
    btn.disabled=false;
    btn.title=String(message||'Clique para limpar linhas excedentes/duplicadas desta NF direto na planilha');
  }
}
function _scheduleAuditDeleteFlush(){
  if(_auditDeleteTimer)clearTimeout(_auditDeleteTimer);
  _auditDeleteTimer=window.setTimeout(()=>{
    _auditDeleteTimer=null;
    _flushAuditDeleteQueue().catch((err)=>console.warn('Falha no lote de limpeza da conferência:',err));
  },_auditDeleteDelayMs);
}
async function _flushAuditDeleteQueue(){
  if(_auditDeleteInFlight){
    if(_auditDeleteQueue.size)_scheduleAuditDeleteFlush();
    return;
  }
  const queued=Array.from(_auditDeleteQueue.values());
  if(!queued.length)return;
  _auditDeleteQueue.clear();
  _auditDeleteInFlight=true;
  _setAuditStatus(`Aplicando ${queued.length} limpeza(s) em lote na planilha...`,true);
  try{
    const payload={audit_keys:queued.map(item=>item.auditKey)};
    const r=await api('/api/conferencia-parcelas/delete',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
    const successMsg=String((r&&r.message)||`${queued.length} limpeza(s) aplicadas com sucesso.`);
    _clearAuditFetchCache();
    queued.forEach(item=>_markAuditRowDeletedLocal(item.btn,successMsg));
    _setAuditStatus(successMsg,false);
    alert(successMsg);
  }catch(e){
    queued.forEach(item=>_restoreAuditRowDeleteLocal(item.btn));
    _setAuditStatus('Falha ao aplicar a limpeza em lote da conferência.',false);
    alert('Erro ao limpar na planilha: '+e.message);
  }finally{
    _auditDeleteInFlight=false;
    if(_auditDeleteQueue.size){
      _setAuditStatus(`${_auditDeleteQueue.size} limpeza(s) ainda estão na fila. Novo envio em até 3s.`,true);
      _scheduleAuditDeleteFlush();
    }
  }
}
async function deleteAuditRows(btn,auditKey,nf,statusLabel,deleteCandidates){
  const key=String(auditKey||'').trim();
  const nfView=String(nf||'-').trim()||'-';
  const statusView=String(statusLabel||'-').trim()||'-';
  const count=Math.max(0, Number(deleteCandidates||0));
  const auditRow=_auditResolveRowTarget(btn);
  const rowData=auditRow&&auditRow.getData&&auditRow.getData()||null;
  if(!key){
    alert('Não foi possível identificar a NF selecionada.');
    return;
  }
  if(rowData&&rowData._local_removed){
    _setAuditStatus(`A limpeza da NF ${nfView} já foi aplicada localmente nesta carga.`,false);
    return;
  }
  if(rowData&&rowData._local_pending){
    _setAuditStatus(`A NF ${nfView} já está na fila de limpeza em lote.`,true);
    return;
  }
  if(count<=0){
    alert(`A NF ${nfView} está com status ${statusView}, mas não há linhas pendentes removíveis automaticamente na planilha.`);
    return;
  }
  const msg=`Tem certeza que deseja limpar ${count} linha(s) excedente(s)/duplicada(s) da NF ${nfView} direto na planilha?\nEssa ação apaga apenas o conteúdo das linhas identificadas como sobra na Conferência, sem reordenar o restante.`;
  if(!confirm(msg))return;
  if(_auditDeleteQueue.has(key)){
    if(btn)_markAuditRowPendingLocal(btn,'Essa NF já está na fila de limpeza em lote.');
    _setAuditStatus(`${_auditDeleteQueue.size} limpeza(s) em fila. Envio em até 3s.`,true);
    return;
  }
  _auditDeleteQueue.set(key,{auditKey:key,btn:btn,nf:nfView});
  _markAuditRowPendingLocal(btn,`NF ${nfView} adicionada à fila de limpeza em lote.`);
  _setAuditStatus(`${_auditDeleteQueue.size} limpeza(s) em fila. Envio em até 3s.`,true);
  _scheduleAuditDeleteFlush();
}
async function loadParcelAudit(silent=false){
  const reqId=++_auditLoadSeq;
  const showLoading=!silent||_activeTab==='audit';
  _stopAuditJobPolling();
  try{
    const runBtn=document.getElementById('auditRunBtn');
    if(runBtn&&document.activeElement===runBtn)runBtn.blur();
  }catch(_){}
  if(showLoading){
    _setAuditLoading(true,'Conferindo planilhas...');
  }
  try{
    const state=_readAuditUiState();
    const requestKey=_auditFetchKey(state,state.empresa);
    _auditActiveRequestKey=requestKey;
    const p=_buildAuditQuery(state,state.empresa);
    const data=await api('/api/conferencia-parcelas/start?'+p.toString());
    if(reqId!==_auditLoadSeq||requestKey!==_auditActiveRequestKey)return;
    const rendered=_applyAuditPayload(data);
    const message=String(data&&data.message||'Conferindo planilhas...').trim()||'Conferindo planilhas...';
    if(String(data&&data.status||'').trim()==='done'){
      _auditActiveJobId='';
      _setAuditLoading(false, rendered ? _buildAuditLoadedMessage((data.result||{}).meta||{}, 'Conferência atualizada') : message);
      return;
    }
    _auditActiveJobId=String(data&&data.job_id||'').trim();
    if(rendered)_setAuditLoading(false, message);
    else _setAuditLoading(true, message);
    if(reqId!==_auditLoadSeq)return;
    _scheduleAuditJobPoll(_auditActiveJobId,requestKey, rendered ? 1000 : 500);
  }catch(err){
    if(reqId!==_auditLoadSeq)return;
    if(!silent)console.warn('Erro ao carregar conferência:',err);
    _auditActiveJobId='';
    _setAuditSummary({});
    if(_auditHasTabulator()&&_auditTable){
      _auditTable.setData([]);
      _toggleAuditRenderMode(true);
    }
    const body=document.getElementById('aBody');
    if(body)body.innerHTML='<tr><td colspan="8">Erro de rede: '+_esc(String(err&&err.message||err))+'</td></tr>';
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
function _getWatchEmpresa(){
  const group=document.getElementById('watchEmpresaGroup');
  return _normalizeAuditEmpresaClient(group&&group.dataset&&group.dataset.value||'mva');
}
function _syncWatchEmpresaButtons(){
  const active=_getWatchEmpresa();
  document.querySelectorAll('#watchEmpresaGroup .audit-source-btn[data-empresa]').forEach((btn)=>{
    btn.classList.toggle('active',String(btn.dataset.empresa||'')===active);
  });
  _syncWatchToneMode();
}
function _syncWatchToneMode(){
  const toneClass=`audit-tone-${_getWatchEmpresa()}`;
  ['audit-tone-mva','audit-tone-eh','audit-tone-todos'].forEach((cls)=>{
    const legacy=document.getElementById('watchTableLegacy');
    const host=document.getElementById('watchTableTabulator');
    if(legacy)legacy.classList.remove(cls);
    if(host)host.classList.remove(cls);
  });
  const legacy=document.getElementById('watchTableLegacy');
  const host=document.getElementById('watchTableTabulator');
  if(legacy)legacy.classList.add(toneClass);
  if(host)host.classList.add(toneClass);
}
function setWatchEmpresa(value,reload=true){
  const group=document.getElementById('watchEmpresaGroup');
  if(!group)return;
  const next=_normalizeAuditEmpresaClient(value);
  const prev=_getWatchEmpresa();
  group.dataset.value=next;
  _syncWatchEmpresaButtons();
  if(reload!==false&&prev!==next&&_activeTab==='watch'){
    loadDueWatch(false).catch(()=>{});
  }
}
function _toggleWatchRenderMode(useTabulator){
  const host=document.getElementById('watchTableTabulator');
  const legacy=document.getElementById('watchTableLegacy');
  if(host)host.classList.toggle('active',!!useTabulator);
  if(legacy)legacy.style.display=useTabulator?'none':'table';
}
function _prepareWatchRows(items){
  return (Array.isArray(items)?items:[]).map((it)=>{
    const nf=String(it.nf||'').trim();
    const local=_compactSpaces(it.local||it.aba||'-');
    const empresaKey=_inferEmpresaKeyFromValue(it.empresa,it.local||it.aba||'',nf);
    return {
      ...it,
      _cliente_view:_compactClienteLabel(it.cliente,it.descricao),
      _cliente_title:_compactSpaces(it.descricao||it.cliente||'-'),
      _local_view:local,
      _empresa_key:empresaKey,
      _tone_class:_toneClassForEmpresa(_getWatchEmpresa(),empresaKey),
    };
  });
}
function _watchStatusFormatter(cell){
  const data=cell.getRow().getData()||{};
  return `<span class="watch-badge ${_esc(data.status||'aviso')}">${_esc(data.status_label||'-')}</span>`;
}
function _watchColumns(){
  return [
    {title:'Tipo',field:'tipo_label',hozAlign:'center',headerHozAlign:'center',width:92},
    {title:'Situação',field:'status_label',hozAlign:'center',headerHozAlign:'center',width:126,formatter:_watchStatusFormatter},
    {title:'Vencimento',field:'vencimento',hozAlign:'center',headerHozAlign:'center',width:116,sorter:(a,b)=>_auditDateSortValue(a)-_auditDateSortValue(b),formatter:(cell)=>_fmtAuditDate(cell.getValue()||'-')||'-'},
    {title:'Dias úteis',field:'dias_label',hozAlign:'center',headerHozAlign:'center',width:136},
    {title:'Cliente',field:'_cliente_view',hozAlign:'center',headerHozAlign:'center',minWidth:260,widthGrow:4,headerFilter:'input'},
    {title:'NF',field:'nf',hozAlign:'center',headerHozAlign:'center',width:88,sorter:'number',headerFilter:'input'},
    {title:'Valor',field:'valor',hozAlign:'center',headerHozAlign:'center',width:124,sorter:'number',formatter:(cell)=>_fmtMoney(cell.getValue())},
    {title:'Aba',field:'_local_view',hozAlign:'center',headerHozAlign:'center',minWidth:120,widthGrow:1,headerFilter:'input'},
  ];
}
function _ensureWatchTabulator(initialRows){
  if(!_historyHasTabulator())return null;
  if(_watchTable)return _watchTable;
  _watchTable=new Tabulator('#watchTableTabulator',{
    data:Array.isArray(initialRows)?initialRows:[],
    layout:'fitDataStretch',
    responsiveLayout:false,
    pagination:'local',
    paginationSize:14,
    paginationCounter:'rows',
    movableColumns:true,
    resizableColumns:true,
    placeholder:'Nenhum boleto ou depósito pendente encontrado para os limites escolhidos',
    columns:_watchColumns(),
    rowFormatter:function(row){
      const el=row.getElement();
      const data=row.getData()||{};
      el.classList.remove('watch-row-aviso','watch-row-erro','audit-row-tone-mva','audit-row-tone-eh');
      if(data._tone_class)el.classList.add(String(data._tone_class));
      if(data.status==='erro')el.classList.add('watch-row-erro');
      else el.classList.add('watch-row-aviso');
      el.title=String(data._cliente_title||'').trim();
    },
  });
  return _watchTable;
}
function _renderDueWatch(items){
  const arr=_prepareWatchRows(items);
  _syncWatchToneMode();
  if(_historyHasTabulator()){
    const table=_ensureWatchTabulator(arr);
    if(table){
      _toggleWatchRenderMode(true);
      try{table.setData(arr);}catch(_){}
      return;
    }
  }
  _toggleWatchRenderMode(false);
  const body=document.getElementById('wBody');
  if(!body)return;
  body.innerHTML='';
  if(!arr.length){
    body.innerHTML='<tr><td colspan="8">Nenhum boleto ou depósito pendente encontrado para os limites escolhidos</td></tr>';
    return;
  }
  arr.forEach(it=>{
    const tr=document.createElement('tr');
    if(it._tone_class)tr.classList.add(String(it._tone_class));
    if(it.status==='erro')tr.classList.add('watch-row-erro');
    else tr.classList.add('watch-row-aviso');
    tr.innerHTML=`<td>${_esc(it.tipo_label||'-')}</td><td><span class="watch-badge ${_esc(it.status||'aviso')}">${_esc(it.status_label||'-')}</span></td><td title="${_esc(_fmtAuditDate(it.vencimento))}">${_esc(_fmtAuditDate(it.vencimento))}</td><td title="${_esc(it.dias_label||'-')}">${_esc(it.dias_label||'-')}</td><td title="${_esc(it._cliente_title||'-')}">${_esc(it._cliente_view||'-')}</td><td title="${_esc(it.nf||'-')}">${_esc(it.nf||'-')}</td><td title="${_esc(_fmtMoney(it.valor))}">${_esc(_fmtMoney(it.valor))}</td><td title="${_esc(it._local_view||'-')}">${_esc(it._local_view||'-')}</td>`;
    body.appendChild(tr);
  });
}
async function loadDueWatch(silent=false){
  const reqId=++_watchLoadSeq;
  const showLoading=!silent||_activeTab==='watch';
  const boletoInput=document.getElementById('wBoletoDays');
  const depositoInput=document.getElementById('wDepositoDays');
  const empresa=_getWatchEmpresa();
  let boletoDays=Math.max(1,Math.min(7,Number((boletoInput&&boletoInput.value)||7)||7));
  let depositoDays=Math.max(1,Math.min(7,Number((depositoInput&&depositoInput.value)||7)||7));
  if(boletoInput)boletoInput.value=String(boletoDays);
  if(depositoInput)depositoInput.value=String(depositoDays);
  if(showLoading){
    _resetWatchSummary();
    _setWatchLoading(true,'Lendo planilhas...');
  }
  try{
    const p=new URLSearchParams();
    p.set('boleto_dias',String(boletoDays));
    p.set('deposito_dias',String(depositoDays));
    p.set('empresa',empresa);
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
    if(_historyHasTabulator()&&_watchTable){
      try{_watchTable.setData([]);}catch(_){}
      _toggleWatchRenderMode(true);
    }
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
    const j=await api('/api/prazos/search-suggestions?empresa='+encodeURIComponent(_getWatchEmpresa()));
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
    const j=await api('/api/prazos/search?nome='+encodeURIComponent(query)+'&empresa='+encodeURIComponent(_getWatchEmpresa()));
    _setWatchSearchCatalog(j&&j.suggestions||[]);
    _renderWatchSearchResults(j&&j.items||[], String((j&&j.message)||'Busca concluída.'));
  }catch(err){
    _renderWatchSearchResults([], 'Falha ao consultar boletos em aberto.');
    _setWatchSearchState(String(err&&err.message||err),false);
  }finally{
    if(btn){btn.disabled=false;btn.textContent='Buscar';}
  }
}
async function logout(){await fetch(_url('/api/logout'),{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'}).catch(()=>{});window.location.href=_url('/login');}
['mode','maxPages','pageSize','intervalMin'].forEach(id=>{const el=document.getElementById(id);if(!el)return;el.addEventListener('keydown',(e)=>{if(e.key==='Enter'){e.preventDefault();saveSettings();}});});
['limit'].forEach(id=>{const el=document.getElementById(id);if(!el)return;el.addEventListener('keydown',(e)=>{if(e.key==='Enter'){e.preventDefault();reprocess();}});});
['recoverDateFrom','recoverDateTo','recoverNfStart','recoverNfEnd','recoverMode'].forEach(id=>{const el=document.getElementById(id);if(!el)return;el.addEventListener('keydown',handleRecoverFieldKeydown);});
['recoverNfStart','recoverNfEnd'].forEach(id=>{const el=document.getElementById(id);if(!el)return;el.addEventListener('input',()=>{el.value=_recoverDigits(el.value||'');});});
document.querySelectorAll('#hAt,#hVenc,#hNf,#hCliente,#hAba,#hLimit').forEach(el=>{el.addEventListener('keydown',(e)=>{if(e.key==='Enter'){e.preventDefault();loadHistory();}});});
['aMode','aMonth','aNfStart','aNfEnd'].forEach(id=>{const el=document.getElementById(id);if(!el)return;el.addEventListener('keydown',handleAuditFieldKeydown);});
['aNfStart','aNfEnd'].forEach(id=>{const el=document.getElementById(id);if(!el)return;el.addEventListener('input',()=>{el.value=_recoverDigits(el.value||'');});});
document.querySelectorAll('#wBoletoDays,#wDepositoDays').forEach(el=>{el.addEventListener('keydown',(e)=>{if(e.key==='Enter'){e.preventDefault();loadDueWatch();}});});
['watchSearchInput'].forEach(id=>{const el=document.getElementById(id);if(!el)return;el.addEventListener('keydown',(e)=>{if(e.key==='Enter'){e.preventDefault();searchOpenBoletos();}});});
['continueReprocessQty'].forEach(id=>{const el=document.getElementById(id);if(!el)return;el.addEventListener('keydown',(e)=>{if(e.key==='Enter'){e.preventDefault();continueReprocessFromPrompt();}});});
['watchSearchInput'].forEach(id=>{const el=document.getElementById(id);if(!el)return;el.addEventListener('input',()=>{_renderWatchSearchSuggestions(el.value||'');});el.addEventListener('focus',()=>{_renderWatchSearchSuggestions(el.value||'');});});
window.addEventListener('keydown',(e)=>{
  if(e.key!=='Escape')return;
  const continueModal=document.getElementById('continueReprocessModal');
  if(continueModal&&continueModal.classList.contains('show')){
    e.preventDefault();
    closeContinueReprocessModal();
    return;
  }
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
setHistoryEmpresa(_getHistoryEmpresa(),false);
setAuditEmpresa(_getAuditEmpresa(),false);
setWatchEmpresa(_getWatchEmpresa(),false);
toggleRecoverFilters();
_renderRecoverNfTags();
toggleAuditFilters();
refresh();loadHistory();switchTab(_tabFromLocation());setInterval(refresh,3000);setInterval(_tickNext,1000);setInterval(()=>{if(_activeTab==='hist')loadHistory(true);},10000);
initHubBackButton();
</script></body></html>"""


def _render_audit_tabulator_preview_html() -> str:
    return """<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Botana - Preview Tabulator</title>
<link rel="preconnect" href="https://unpkg.com" crossorigin />
<link rel="stylesheet" href="https://unpkg.com/tabulator-tables@6.3.1/dist/css/tabulator.min.css" />
<script src="https://unpkg.com/tabulator-tables@6.3.1/dist/js/tabulator.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=Lexend:wght@300;400;600;700;800&display=swap');
:root{--o:#da7a1c;--o2:#ee9b2f;--b:#4a2b18;--bg:#f8efe6;--br:#e4c6a7;--line:#e7c4a5}
*{box-sizing:border-box}
body{margin:0;min-height:100vh;font-family:'Lexend',Arial,sans-serif;background:linear-gradient(160deg,#2f1a0f,#5c341c);padding:18px;color:#2a1b12}
.preview-shell{max-width:1180px;margin:0 auto;display:grid;gap:16px}
.preview-head{padding:22px 24px;border:1px solid rgba(231,200,168,.85);border-radius:20px;background:linear-gradient(180deg,rgba(255,250,246,.98),rgba(255,245,235,.94));box-shadow:0 24px 60px rgba(21,11,6,.28);text-align:center}
.preview-kicker{margin:0;color:#a65e20;font-size:12px;font-weight:800;letter-spacing:.08em;text-transform:uppercase}
.preview-head h1{margin:10px 0 0;font-size:34px;line-height:1.06;color:var(--b)}
.preview-head p{margin:12px auto 0;max-width:760px;color:#6c4a35;font-size:14px;line-height:1.65}
.preview-actions{margin-top:16px;display:flex;justify-content:center;gap:10px;flex-wrap:wrap}
.preview-actions a,.preview-actions button{display:inline-flex;align-items:center;justify-content:center;padding:10px 16px;border-radius:999px;border:1px solid #d7b393;background:#fff9f3;color:#5a311b;font-weight:800;text-decoration:none;cursor:pointer}
.preview-actions a:hover,.preview-actions button:hover{background:linear-gradient(90deg,var(--o),var(--o2));border-color:transparent;color:#2b1408}
.card{background:rgba(255,248,240,.95);border:1px solid #e7c8a8;border-radius:16px;padding:14px;box-shadow:0 8px 20px rgba(21,11,6,.08)}
.card h3{margin:0 0 8px;color:var(--b);font-size:1rem;text-align:center}
.audit-filters{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,180px));gap:10px;align-items:end;justify-content:center;max-width:980px;margin:0 auto}
.audit-filters>div{display:flex;flex-direction:column;justify-content:center;align-items:center}
.audit-filters>div label{width:100%;text-align:center;font-weight:700;color:#5c341c}
.audit-filters>div input,.audit-filters>div select,.audit-filters>div button{width:min(180px,100%);text-align:center;padding:9px;border:1px solid #d6b18f;border-radius:9px;background:#fffdfb;font-family:inherit}
.audit-filters button{border:0;background:linear-gradient(90deg,var(--o),var(--o2));color:#2b1408;font-weight:800;cursor:pointer}
.audit-state{min-height:20px;text-align:center;font-size:.84rem;color:#6b4126;margin-top:10px}
.audit-state.loading{color:#a25b18;font-weight:700}
.audit-summary{margin-top:12px;display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:10px}
.audit-summary .k{border:1px solid #e2b58d;border-radius:12px;background:linear-gradient(180deg,#fff7ef,#fff1e3);padding:10px;text-align:center}
.audit-summary .n{font-size:1.3rem;font-weight:800;color:#7a3d11}
.audit-summary .t{font-size:.8rem;color:#6b4126}
.preview-note{margin-top:10px;text-align:center;font-size:12px;color:#6c4a35}
#auditPreviewTableWrap{margin-top:12px;border:1px solid #ddb38d;border-radius:14px;overflow:hidden;background:#fffdfb}
#auditPreviewTable .tabulator{border:none;background:#fffdfb;font-size:.83rem;color:#3f2819}
#auditPreviewTable .tabulator-header{border-bottom:1px solid var(--line);background:#fff1e3}
#auditPreviewTable .tabulator-col,#auditPreviewTable .tabulator-header .tabulator-col{background:transparent;border-right:1px solid #efe0d0;color:#5c341c;font-weight:800}
#auditPreviewTable .tabulator-row{border-bottom:1px solid #f0e0cf;background:#fffdfb}
#auditPreviewTable .tabulator-row:nth-child(even){background:rgba(255,244,232,.92)}
#auditPreviewTable .tabulator-row:hover,#auditPreviewTable .tabulator-row.tabulator-selectable:hover{background:rgba(238,155,47,.08)}
#auditPreviewTable .tabulator-cell{border-right:1px solid #f3e8dc;padding:9px 10px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
#auditPreviewTable .tabulator-footer{border-top:1px solid var(--line);background:#fff8f1;color:#6b4126;font-size:12px;font-weight:700}
#auditPreviewTable .tabulator-page{border:1px solid #d9d0c5;background:#fff;color:#384658}
#auditPreviewTable .tabulator-page.active{background:var(--o);color:#2b1408;border-color:var(--o)}
#auditPreviewTable .tabulator-row.audit-row-aviso{background:rgba(255,193,7,.12)!important}
#auditPreviewTable .tabulator-row.audit-row-aviso:hover{background:rgba(255,193,7,.2)!important}
#auditPreviewTable .tabulator-row.audit-row-erro{background:rgba(220,53,69,.14)!important}
#auditPreviewTable .tabulator-row.audit-row-erro:hover{background:rgba(220,53,69,.22)!important}
#auditPreviewTable .tabulator-row.audit-row-local-pending{background:rgba(240,198,79,.10)!important}
#auditPreviewTable .tabulator-row.audit-row-local-removed td{border-bottom:3px solid #f0c64f!important}
.audit-status{display:inline-flex;align-items:center;justify-content:center;padding:4px 9px;border-radius:999px;font-size:.72rem;font-weight:800;border:1px solid transparent}
.audit-status-btn{cursor:pointer;transition:transform .15s ease, box-shadow .15s ease;background:none}
.audit-status-btn:hover{transform:translateY(-1px);box-shadow:0 4px 10px rgba(92,52,28,.12)}
.audit-status-btn[disabled]{cursor:default;opacity:.78;box-shadow:none;transform:none}
.audit-status.ok{background:#e9f8ec;color:#1c6a32;border-color:#87c69a}
.audit-status.aviso{background:#fff3dd;color:#8b5a00;border-color:#e7bf6e}
.audit-status.erro{background:#fde7ea;color:#a61d2d;border-color:#dc3545}
.muted-center{text-align:center;color:#6c4a35;font-size:.82rem}
@media(max-width:980px){.audit-summary{grid-template-columns:repeat(3,minmax(0,1fr))}}
@media(max-width:760px){body{padding:12px}.preview-head h1{font-size:28px}.audit-filters{grid-template-columns:1fr 1fr}.audit-summary{grid-template-columns:1fr 1fr}}
@media(max-width:560px){.audit-filters{grid-template-columns:1fr}.audit-summary{grid-template-columns:1fr}}
</style></head><body>
<main class="preview-shell">
  <section class="preview-head">
    <p class="preview-kicker">Preview isolado</p>
    <h1>Conferência com Tabulator</h1>
    <p>Este preview existe só para testar como a grade da aba Conferência ficaria usando um grid dinâmico. O painel principal continua igual e esta rota não substitui a tela atual.</p>
    <div class="preview-actions">
      <a id="previewBackLink" href="/">Voltar ao painel atual</a>
      <button type="button" onclick="loadParcelAudit()">Conferir agora</button>
    </div>
  </section>

  <section class="card">
    <h3>Conferência de parcelas lançadas</h3>
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
        <input id="aNfStart" type="text" inputmode="numeric" placeholder="49001"/>
      </div>
      <div>
        <label>NF final</label>
        <input id="aNfEnd" type="text" inputmode="numeric" placeholder="49100"/>
      </div>
      <div style="display:flex;align-items:end;justify-content:center"><button id="auditRunBtn" onclick="loadParcelAudit()">Conferir parcelas</button></div>
    </div>
    <div id="auditStatus" class="audit-state">Pronto para conferir.</div>
    <div class="audit-summary">
      <div class="k"><div id="auditK1" class="n">0</div><div class="t">NFs verificadas</div></div>
      <div class="k"><div id="auditK2" class="n">0</div><div class="t">Com divergência</div></div>
      <div class="k"><div id="auditK3" class="n">0</div><div class="t">Parcelas esperadas</div></div>
      <div class="k"><div id="auditK4" class="n">0</div><div class="t">Parcelas lançadas</div></div>
      <div class="k"><div id="auditK5" class="n">0</div><div class="t">Duplicadas</div></div>
    </div>
    <div class="preview-note">Preview Tabulator: ordenação, filtro por coluna, paginação local e mesma API usada pela aba atual.</div>
    <div id="auditPreviewTableWrap">
      <div id="auditPreviewTable"></div>
    </div>
    <div class="muted-center" style="margin-top:10px">A rota principal continua com a grade atual em HTML. Aqui é só para comparar o comportamento do grid.</div>
  </section>
</main>
<script>
const _PATH_RESERVED=new Set(['','login','logout','api','assets','static','store-image','favicon.ico','preview']);
function _basePrefix(){const p=String(window.location.pathname||'/');const segs=p.split('/').filter(Boolean);if(!segs.length)return '';const first=String(segs[0]||'').toLowerCase();if(_PATH_RESERVED.has(first))return '';return `/${segs[0]}`;}
const _BASE_PREFIX=_basePrefix();
function _url(path){const p=String(path||'');if(!p.startsWith('/'))return p;if(!_BASE_PREFIX)return p;return p.startsWith(`${_BASE_PREFIX}/`)||p===_BASE_PREFIX?p:`${_BASE_PREFIX}${p}`;}
async function api(path,opts){const r=await fetch(_url(path),opts);const j=await r.json().catch(()=>({}));if(r.status===401){window.location.href=_url('/login');throw new Error('não autenticado');}if(!r.ok){throw new Error(String((j&&j.message)||`HTTP ${r.status}`));}return j;}
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
function _fmtLocal(local){const s=String(local||'');if(!s)return '-';const parts=s.split('/');if(parts.length>=2)return parts.slice(-2).join('/');return s;}
function _fmtDateTime(v){const ms=Date.parse(String(v||''));if(!Number.isFinite(ms))return String(v||'-')||'-';return new Date(ms).toLocaleString('pt-BR');}
function _fmtAuditDate(v){
  const txt=String(v||'').trim();
  if(!txt||txt==='-'||txt.toLowerCase()==='invalid date')return '-';
  if(/^\\d{4}-\\d{2}-\\d{2}$/.test(txt)){const [y,m,d]=txt.split('-');return `${d}/${m}/${y}`;}
  if(/^\\d{2}\\/\\d{2}\\/\\d{4}$/.test(txt))return txt;
  return _fmtDateTime(txt);
}
function _auditDateSortValue(value){
  const txt=String(value||'').trim();
  if(!txt||txt==='-'||txt.toLowerCase()==='invalid date')return 0;
  if(/^\\d{4}-\\d{2}-\\d{2}$/.test(txt))return Date.parse(`${txt}T00:00:00`)||0;
  if(/^\\d{2}\\/\\d{2}\\/\\d{4}$/.test(txt)){const [d,m,y]=txt.split('/');return Date.parse(`${y}-${m}-${d}T00:00:00`)||0;}
  const ms=Date.parse(txt);
  return Number.isFinite(ms)?ms:0;
}
function _statusRank(status){const txt=String(status||'').trim().toLowerCase();if(txt==='erro')return 0;if(txt==='aviso')return 1;return 2;}
function _setAuditStatus(message,loading=false){
  const el=document.getElementById('auditStatus');
  const btn=document.getElementById('auditRunBtn');
  if(el){el.textContent=String(message||'Pronto para conferir.');el.classList.toggle('loading',!!loading);}
  if(btn){btn.disabled=!!loading;btn.textContent=loading?'Conferindo...':'Conferir parcelas';}
}
function _setAuditSummary(summary){
  const s=summary||{};
  [['auditK1','nfs_verificadas'],['auditK2','nfs_com_divergencia'],['auditK3','parcelas_esperadas'],['auditK4','parcelas_lancadas'],['auditK5','parcelas_duplicadas']].forEach(([id,key])=>{
    const el=document.getElementById(id);
    if(el)el.textContent=String(s[key]||0);
  });
}
function toggleAuditFilters(){
  const mode=((document.getElementById('aMode')||{}).value||'mes').trim();
  const monthEl=document.getElementById('aMonth');
  const startEl=document.getElementById('aNfStart');
  const endEl=document.getElementById('aNfEnd');
  if(monthEl)monthEl.disabled=mode!=='mes';
  if(startEl)startEl.disabled=mode!=='nfs';
  if(endEl)endEl.disabled=mode!=='nfs';
}
let _auditPreviewTable=null;
function _auditStatusFormatter(cell){
  const data=cell.getRow().getData()||{};
  const label=String(data.status_label||'-');
  const title=String(data.status_title||'').trim();
  if(data.delete_enabled){
    const btn=document.createElement('button');
    btn.type='button';
    btn.className=`audit-status audit-status-btn ${data.status||'ok'}`;
    btn.textContent=label;
    btn.disabled=!!data.local_removed||!!data.local_pending;
    btn.title=title||'Clique para limpar linhas excedentes/duplicadas desta NF direto na planilha';
    btn.addEventListener('click',async(ev)=>{ev.preventDefault();ev.stopPropagation();await deleteAuditRow(cell.getRow());});
    return btn;
  }
  const span=document.createElement('span');
  span.className=`audit-status ${data.status||'ok'}`;
  span.textContent=label;
  if(title)span.title=title;
  return span;
}
function _ensureAuditPreviewTable(){
  if(_auditPreviewTable)return _auditPreviewTable;
  _auditPreviewTable=new Tabulator('#auditPreviewTable',{
    data:[],
    layout:'fitDataStretch',
    responsiveLayout:false,
    pagination:'local',
    paginationSize:15,
    paginationCounter:'rows',
    movableColumns:true,
    resizableColumns:true,
    placeholder:'Nenhuma NF encontrada para os filtros selecionados.',
    initialSort:[{column:'status_rank',dir:'asc'},{column:'nf_num',dir:'desc'}],
    columns:[
      {title:'Status',field:'status_label',hozAlign:'center',headerHozAlign:'center',width:112,sorter:function(a,b,aRow,bRow){return (aRow.getData().status_rank||0)-(bRow.getData().status_rank||0);},formatter:_auditStatusFormatter},
      {title:'NF',field:'nf',hozAlign:'center',headerHozAlign:'center',width:92,sorter:'number',headerFilter:'input'},
      {title:'Cliente',field:'cliente_view',minWidth:300,widthGrow:4,headerFilter:'input'},
      {title:'Parc.',field:'parcelas_view',hozAlign:'center',headerHozAlign:'center',width:104,sorter:function(a,b,aRow,bRow){
        const aData=aRow.getData()||{};
        const bData=bRow.getData()||{};
        const esperadasDiff=Number(aData.qtd_esperada||0)-Number(bData.qtd_esperada||0);
        if(esperadasDiff!==0)return esperadasDiff;
        return Number(aData.qtd_lancada||0)-Number(bData.qtd_lancada||0);
      }},
      {title:'Faltando',field:'qtd_faltando',hozAlign:'center',headerHozAlign:'center',width:96,sorter:'number'},
      {title:'Duplicadas',field:'qtd_duplicada',hozAlign:'center',headerHozAlign:'center',width:112,sorter:'number',formatter:function(cell){const data=cell.getRow().getData()||{};const count=Number(data.qtd_duplicada||0);const view=String(data.duplicadas_view||'').trim();if(count>0&&view&&view!=='-'){return `<span title="${_esc(view)}">${count} - ${_esc(view)}</span>`;}return String(count||0);}},
      {title:'Últ. venc.',field:'ultimo_vencimento_view',hozAlign:'center',headerHozAlign:'center',width:118,sorter:function(a,b,aRow,bRow){return (aRow.getData().ultimo_vencimento_sort||0)-(bRow.getData().ultimo_vencimento_sort||0);}},
      {title:'Aba',field:'local_view',minWidth:180,widthGrow:2,headerFilter:'input'}
    ],
    rowFormatter:function(row){
      const el=row.getElement();
      const data=row.getData()||{};
      el.classList.remove('audit-row-aviso','audit-row-erro','audit-row-local-pending','audit-row-local-removed');
      if(data.local_removed)el.classList.add('audit-row-local-removed');
      else if(data.local_pending)el.classList.add('audit-row-local-pending');
      else if(data.status==='erro')el.classList.add('audit-row-erro');
      else if(data.status==='aviso')el.classList.add('audit-row-aviso');
      const reason=String(data.reason_hint||'').trim();
      if(reason)el.title=reason;
    }
  });
  return _auditPreviewTable;
}
function _mapAuditItems(items){
  return (Array.isArray(items)?items:[]).map((it)=>{
    const reason=String(it.reason_hint||'').trim();
    const local=String(it.local_lancamento||it.aba||'-');
    const lastDueRaw=String(it.ultimo_vencimento||it.ultimo_lancamento||'');
    const duplicadasView=Number(it.qtd_duplicada||0)>0?((Array.isArray(it.parcelas_duplicadas)?it.parcelas_duplicadas:[]).map(v=>_compactSpaces(v)).filter(Boolean).join(', ')||'-'):'-';
    const deleteCandidates=Math.max(0,Number(it.delete_candidates||0));
    const esperadas=Math.max(0,Number(it.qtd_esperada||0));
    const lancadas=Math.max(0,Number(it.qtd_lancada||0));
    return Object.assign({},it,{
      nf_num:Number(it.nf||0),
      status_rank:_statusRank(it.status||''),
      cliente_view:_compactClienteLabel(it.cliente,it.descricao),
      parcelas_view:`${esperadas} / ${lancadas}`,
      local_view:_fmtLocal(local),
      ultimo_vencimento_view:_fmtAuditDate(lastDueRaw),
      ultimo_vencimento_sort:_auditDateSortValue(lastDueRaw),
      duplicadas_view:duplicadasView,
      delete_enabled:!!(it.status&&it.status!=='ok'&&deleteCandidates>0),
      status_title:reason || (deleteCandidates>0?'Clique para limpar linhas excedentes/duplicadas desta NF direto na planilha':''),
      local_removed:false,
      local_pending:false
    });
  });
}
async function deleteAuditRow(row){
  const data=row&&typeof row.getData==='function'?row.getData():null;
  if(!data)return;
  const auditKey=String(data.audit_key||'').trim();
  const nfView=String(data.nf||'-').trim()||'-';
  const count=Math.max(0,Number(data.delete_candidates||0));
  if(!auditKey||count<=0){alert(`A NF ${nfView} não tem linhas removíveis automaticamente na planilha.`);return;}
  const msg=`Tem certeza que deseja limpar ${count} linha(s) excedente(s)/duplicada(s) da NF ${nfView} direto na planilha?\\nEssa ação apaga apenas o conteúdo das linhas identificadas como sobra na Conferência, sem reordenar o restante.`;
  if(!confirm(msg))return;
  await row.update({local_pending:true});
  _setAuditStatus(`Aplicando limpeza da NF ${nfView} na planilha...`,true);
  try{
    const res=await api('/api/conferencia-parcelas/delete',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({audit_key:auditKey})});
    await row.update({local_pending:false,local_removed:true});
    _setAuditStatus(String((res&&res.message)||`Limpeza da NF ${nfView} aplicada com sucesso.`),false);
  }catch(err){
    await row.update({local_pending:false,local_removed:false});
    _setAuditStatus('Falha ao limpar na planilha.',false);
    alert('Erro ao limpar na planilha: '+String(err&&err.message||err));
  }
}
async function loadParcelAudit(){
  _setAuditStatus('Conferindo planilhas...',true);
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
    const data=await api('/api/conferencia-parcelas?'+p.toString());
    _setAuditSummary((data&&data.summary)||{});
    const table=_ensureAuditPreviewTable();
    table.setData(_mapAuditItems((data&&data.items)||[]));
    const meta=(data&&data.meta)||{};
    const loadedAt=String(meta.loaded_at||'').trim();
    const linhas=Number(meta.linhas_lidas||0);
    const abas=Number(meta.abas_lidas||0);
    const statusMsg=loadedAt?`Conferência atualizada em ${_fmtDateTime(loadedAt)} | ${linhas} linhas lidas em ${abas} abas`:'Conferência atualizada.';
    _setAuditStatus(statusMsg,false);
  }catch(err){
    _setAuditSummary({});
    const table=_ensureAuditPreviewTable();
    table.setData([]);
    _setAuditStatus('Falha ao conferir as planilhas.',false);
    alert('Erro ao carregar conferência: '+String(err&&err.message||err));
  }
}
function handleAuditFieldKeydown(ev){
  if(ev.key!=='Enter')return;
  ev.preventDefault();
  const mode=String((document.getElementById('aMode')||{}).value||'mes').trim();
  const currentId=String((ev&&ev.target&&ev.target.id)||'').trim();
  if(currentId==='aMode'){
    if(mode==='nfs'){document.getElementById('aNfStart').focus();return;}
    if(mode==='mes'){document.getElementById('aMonth').focus();return;}
    loadParcelAudit();return;
  }
  if(mode==='mes'){loadParcelAudit();return;}
  if(mode==='nfs'&&currentId==='aNfStart'){document.getElementById('aNfEnd').focus();return;}
  loadParcelAudit();
}
document.getElementById('previewBackLink').setAttribute('href',_url('/'));
['aMode','aMonth','aNfStart','aNfEnd'].forEach(id=>{const el=document.getElementById(id);if(el)el.addEventListener('keydown',handleAuditFieldKeydown);});
try{document.getElementById('aMonth').value=new Date().toISOString().slice(0,7);}catch(_){}
toggleAuditFilters();
_ensureAuditPreviewTable();
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

            if parsed.path == "/preview/tabulator/conferencia":
                return _html_response(self, 200, _render_audit_tabulator_preview_html())
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
                empresa = (qs.get("empresa", [""])[0] or "todos").strip()
                try:
                    resultado = _gerar_conferencia_parcelas(filtro, mes, nf_inicio, nf_fim, empresa)
                    return _json_response(self, 200, {"ok": True, **resultado})
                except Exception as e:
                    return _json_response(self, 500, {"ok": False, "message": str(e)})

            if parsed.path == "/api/conferencia-parcelas/start":
                qs = parse_qs(parsed.query or "")
                filtro = (qs.get("filtro", [""])[0] or "mes").strip()
                mes = (qs.get("mes", [""])[0] or "").strip()
                nf_inicio = (qs.get("nf_inicio", [""])[0] or "").strip()
                nf_fim = (qs.get("nf_fim", [""])[0] or "").strip()
                empresa = (qs.get("empresa", [""])[0] or "todos").strip()
                try:
                    resultado = _start_audit_job(filtro, mes, nf_inicio, nf_fim, empresa)
                    return _json_response(self, 200, {"ok": True, **resultado})
                except Exception as e:
                    return _json_response(self, 500, {"ok": False, "message": str(e)})

            if parsed.path == "/api/conferencia-parcelas/job":
                qs = parse_qs(parsed.query or "")
                job_id = (qs.get("id", [""])[0] or "").strip()
                try:
                    resultado = _get_audit_job(job_id)
                    return _json_response(self, 200, {"ok": True, **resultado})
                except Exception as e:
                    return _json_response(self, 500, {"ok": False, "message": str(e)})

            if parsed.path == "/api/prazos":
                qs = parse_qs(parsed.query or "")
                empresa = (qs.get("empresa", ["todos"])[0] or "todos").strip()
                try:
                    boleto_dias = int((qs.get("boleto_dias", ["7"])[0] or "7").strip())
                except Exception:
                    boleto_dias = 7
                try:
                    deposito_dias = int((qs.get("deposito_dias", ["7"])[0] or "7").strip())
                except Exception:
                    deposito_dias = 7
                try:
                    resultado = _gerar_relacao_pendencias(boleto_dias, deposito_dias, empresa)
                    return _json_response(self, 200, {"ok": True, **resultado})
                except Exception as e:
                    return _json_response(self, 500, {"ok": False, "message": str(e)})

            if parsed.path == "/api/prazos/search-suggestions":
                try:
                    qs = parse_qs(parsed.query or "")
                    empresa = (qs.get("empresa", ["todos"])[0] or "todos").strip()
                    return _json_response(self, 200, {"ok": True, "items": _load_watch_search_suggestions(empresa)})
                except Exception as e:
                    return _json_response(self, 500, {"ok": False, "message": str(e)})

            if parsed.path == "/api/prazos/search":
                qs = parse_qs(parsed.query or "")
                nome = (qs.get("nome", [""])[0] or "").strip()
                empresa = (qs.get("empresa", ["todos"])[0] or "todos").strip()
                if not nome:
                    return _json_response(self, 400, {"ok": False, "message": "Informe um nome para consultar."})
                try:
                    resultado = _buscar_boletos_em_aberto_por_nome(nome, empresa)
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
                empresa = (qs.get("empresa", ["todos"])[0] or "todos").strip()
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
                    empresa_filter=empresa,
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
                empresa = (qs.get("empresa", ["todos"])[0] or "todos").strip()
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
                    empresa_filter=empresa,
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
                    mark_unread = bool(data.get("mark_unread", False))
                    continue_after_id = str(data.get("continue_after_id", "") or "").strip()
                    started, info = _start_reprocess_background(
                        max_messages=max_messages,
                        mark_unread=mark_unread,
                        continue_after_id=continue_after_id,
                    )
                    if not started:
                        msg = _manual_action_busy_message() or str((info or {}).get("message") or "Nao foi possivel iniciar o reprocessamento.")
                        return _json_response(self, 409, {"ok": False, "message": msg, "action": info})
                    friendly = (
                        f"Continuação do reprocessamento iniciada para até {max_messages} mensagens mais antigas dentro dos últimos {_REPROCESS_LOOKBACK_DAYS} dias; a leitura será executada em seguida."
                        if continue_after_id
                        else f"Reprocessamento iniciado para até {max_messages} mensagens mais recentes dos últimos {_REPROCESS_LOOKBACK_DAYS} dias; a leitura será executada em seguida."
                    )
                    return _json_response(self, 202, {"ok": True, "started": True, "friendly": friendly, "action": info})
                if parsed.path in ("/api/recover-emails", "/api/recover-missing"):
                    if not _can_operate(user):
                        return _json_response(self, 403, {"ok": False, "message": "Sem permissao"})
                    try:
                        max_messages = max(1, min(1000, int(data.get("max_messages", 1000))))
                    except Exception:
                        max_messages = 1000
                    mode = str(data.get("mode", "") or "").strip()
                    nf_start = str(data.get("nf_start", "") or "").strip()
                    nf_end = str(data.get("nf_end", "") or "").strip()
                    date_from = str(data.get("date_from", "") or "").strip()
                    date_to = str(data.get("date_to", "") or "").strip()
                    nf_list = data.get("nf_list", [])
                    started, info = _start_recover_missing_background(
                        max_messages=max_messages,
                        mode=mode,
                        nf_start=nf_start,
                        nf_end=nf_end,
                        date_from=date_from,
                        date_to=date_to,
                        nf_list=nf_list,
                    )
                    if not started:
                        msg = str((info or {}).get("message") or "").strip()
                        if not msg:
                            msg = _manual_action_busy_message() or "Nao foi possivel iniciar a recuperacao."
                        status_code = 400 if any(token in msg for token in ("Informe", "Escolha", "Adicione")) else 409
                        return _json_response(self, status_code, {"ok": False, "message": msg, "action": info})
                    filtros = _describe_recovery_filters(mode=mode, nf_start=nf_start, nf_end=nf_end, date_from=date_from, date_to=date_to, nf_list=nf_list) or "os filtros informados"
                    friendly = f"Recuperação iniciada para mensagens com XML em {filtros}."
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

                if parsed.path == "/api/conferencia-parcelas/delete":
                    if not _can_operate(user):
                        return _json_response(self, 403, {"ok": False, "message": "Sem permissão"})
                    auditKeys = []
                    for rawKey in list(data.get("audit_keys") or []):
                        key = str(rawKey or "").strip()
                        if key and key not in auditKeys:
                            auditKeys.append(key)
                    auditKey = str(data.get("audit_key", "") or "").strip()
                    if auditKey and auditKey not in auditKeys:
                        auditKeys.append(auditKey)
                    if not auditKeys:
                        return _json_response(self, 400, {"ok": False, "message": "NF da conferência não informada"})
                    resultado = _delete_audit_rows(auditKeys)
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
