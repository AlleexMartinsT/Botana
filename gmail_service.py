# gmail_service.py
import os
import base64
import time
import logging
from datetime import datetime, timedelta
from typing import List, Dict, Any, Tuple
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from config import GOOGLE_CREDENTIALS_GMAIL, DOWNLOAD_DIR


# Scopes: precisamos de modify para acrescentar labels (e opcionalmente marcar como lido)
SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]

logger = logging.getLogger("bot.gmail_service")
LABEL_NAME = "XML Processado Botana"
REPROCESS_LABEL_NAME = "XML Reprocessado Botana"
BOTANA_LABEL_PREFIXES = (LABEL_NAME, REPROCESS_LABEL_NAME)
_LABEL_CACHE = {"at": 0.0, "labels": []}
DEFAULT_SENT_XML_QUERY = "in:sent has:attachment filename:xml"


def _label_date(when=None) -> str:
    ref = when or datetime.now()
    return ref.strftime("%d/%m/%Y")


def build_botana_label_name(reprocessed: bool = False, when=None) -> str:
    prefix = REPROCESS_LABEL_NAME if reprocessed else LABEL_NAME
    return f"{prefix} - {_label_date(when)}"


def _list_labels(service, force: bool = False) -> List[Dict[str, Any]]:
    now = time.time()
    if not force and _LABEL_CACHE.get("labels") and (now - float(_LABEL_CACHE.get("at", 0.0)) < 60):
        return list(_LABEL_CACHE.get("labels") or [])
    labels = service.users().labels().list(userId="me").execute().get("labels", [])
    _LABEL_CACHE.update({"at": now, "labels": list(labels or [])})
    return list(labels or [])


def list_botana_labels(service) -> List[Dict[str, Any]]:
    labels = _list_labels(service)
    out = []
    for item in labels:
        name = str(item.get("name", "")).strip()
        if any(name.startswith(prefix) for prefix in BOTANA_LABEL_PREFIXES):
            out.append(item)
    return out


def list_botana_label_ids(service) -> List[str]:
    ids = []
    for item in list_botana_labels(service):
        label_id = str(item.get("id", "")).strip()
        if label_id:
            ids.append(label_id)
    return ids


def build_period_query(mode: str = "last_30_days", when=None) -> str:
    ref = when or datetime.now()
    mode_norm = str(mode or "last_30_days").strip().lower()
    today = ref.date()
    if mode_norm == "last_15_days":
        return "newer_than:15d"
    if mode_norm == "last_30_days":
        return "newer_than:30d"
    if mode_norm == "last_45_days":
        return "newer_than:45d"
    if mode_norm == "last_60_days":
        return "newer_than:60d"
    if mode_norm == "current_week":
        week_start = today - timedelta(days=today.weekday())
        return f"after:{week_start.strftime('%Y/%m/%d')}"
    first_this_month = today.replace(day=1)
    if mode_norm == "previous_month":
        prev_month_last = first_this_month - timedelta(days=1)
        prev_month_start = prev_month_last.replace(day=1)
        return (
            f"after:{prev_month_start.strftime('%Y/%m/%d')} "
            f"before:{first_this_month.strftime('%Y/%m/%d')}"
        )
    if mode_norm == "current_and_previous_month":
        prev_month_last = first_this_month - timedelta(days=1)
        prev_month_start = prev_month_last.replace(day=1)
        return f"after:{prev_month_start.strftime('%Y/%m/%d')}"
    return "newer_than:30d"


def build_sent_xml_query(filter_mode: str = "last_30_days", extra_query: str = "") -> str:
    parts = [DEFAULT_SENT_XML_QUERY]
    period = build_period_query(mode=filter_mode) if str(filter_mode or "").strip() else ""
    if period:
        parts.append(period)
    extra = str(extra_query or "").strip()
    if extra:
        parts.append(extra)
    return " ".join(part for part in parts if str(part).strip())

def _get_token_path(cred_path: str) -> str:
    return cred_path.replace(".json", "_token.json")

def getGmailService(cred_file: str = GOOGLE_CREDENTIALS_GMAIL):
    """
    Autentica e retorna um serviço Gmail (v1). Salva token em cred_file_token.json.
    """
    token_path = _get_token_path(cred_file)
    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(cred_file, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w", encoding="utf-8") as fh:
            fh.write(creds.to_json())

    service = build("gmail", "v1", credentials=creds)
    return service

def ensure_label(service, label_name: str = LABEL_NAME) -> str:
    """Retorna o id do rótulo, criando se necessário."""
    labels = service.users().labels().list(userId="me").execute().get("labels", [])
    for l in labels:
        if l.get("name", "").lower() == label_name.lower():
            return l["id"]

    body = {"name": label_name, "labelListVisibility": "labelShow", "messageListVisibility": "show"}
    created = service.users().labels().create(userId="me", body=body).execute()
    fresh = list(labels)
    fresh.append(created)
    _LABEL_CACHE.update({"at": time.time(), "labels": fresh})
    logger.info("Rótulo criado: %s (%s)", label_name, created.get("id"))
    return created.get("id")

def buscarMessagesEnviados(service, max_results: int = 15) -> List[Dict[str, Any]]:
    """
    Busca threads com mensagens enviadas contendo anexos XML
    e retorna todas as mensagens (enviadas e recebidas) dentro dessas threads.
    """
    q = "in:sent has:attachment filename:xml"

    try:
        resp = service.users().threads().list(userId="me", q=q, maxResults=max_results).execute()
        threads = resp.get("threads", []) or []
        results = []

        logger.info("Buscar: %d threads encontradas", len(threads))

        for t in threads:
            thread_id = t.get("id")
            try:
                thread = service.users().threads().get(userId="me", id=thread_id).execute()
                msgs = thread.get("messages", [])
                
                for msg in msgs:
                    payload = msg.get("payload", {})
                    parts = payload.get("parts", []) or []

                    # Verifica anexos dentro das partes aninhadas (caso Gmail codifique assim)
                    for part in parts:
                        if "parts" in part:
                            for sub in part["parts"]:
                                if sub.get("filename", "").lower().endswith(".xml"):
                                    break

                    results.append({
                        "id": msg["id"],
                        "threadId": msg["threadId"],
                        "labelIds": msg.get("labelIds", []),
                        "snippet": msg.get("snippet", ""),
                    })

            except Exception as e:
                logger.warning("Falha ao obter thread %s: %s", thread_id, e)
                continue
        return results

    except Exception as e:
        logger.exception("Erro ao listar threads: %s", e)
        return []


def buscarMessagesEnviadosPagina(
    service,
    max_results: int = 100,
    page_token: str | None = None,
    query: str | None = None,
    skip_label_ids=None,
) -> Tuple[List[Dict[str, Any]], str | None]:
    """
    Busca uma página de threads com mensagens enviadas contendo XML.
    Retorna (mensagens, next_page_token).
    """
    q = str(query or DEFAULT_SENT_XML_QUERY).strip() or DEFAULT_SENT_XML_QUERY
    try:
        req = service.users().threads().list(
            userId="me",
            q=q,
            maxResults=max_results,
            pageToken=page_token,
        )
        resp = req.execute()
        threads = resp.get("threads", []) or []
        next_token = resp.get("nextPageToken")
        results: List[Dict[str, Any]] = []
        skip_ids = {str(item).strip() for item in (skip_label_ids or []) if str(item).strip()}

        logger.info("Buscar página: %d threads encontradas", len(threads))

        for t in threads:
            thread_id = t.get("id")
            try:
                thread = service.users().threads().get(userId="me", id=thread_id, format="metadata").execute()
                msgs = thread.get("messages", [])

                for msg in msgs:
                    label_ids = [str(label_id).strip() for label_id in (msg.get("labelIds", []) or []) if str(label_id).strip()]
                    if "SENT" not in label_ids:
                        continue
                    if skip_ids and any(label_id in skip_ids for label_id in label_ids):
                        continue
                    results.append(
                        {
                            "id": msg["id"],
                            "threadId": msg["threadId"],
                            "labelIds": label_ids,
                            "snippet": msg.get("snippet", ""),
                        }
                    )
            except Exception as e:
                logger.warning("Falha ao obter thread %s: %s", thread_id, e)
                continue

        return results, next_token

    except Exception as e:
        logger.exception("Erro ao listar threads (página): %s", e)
        return [], None

def _flatten_parts(parts):
    """
    Retorna lista plana de partes que representam anexos (ou potenciais anexos) — contempla recursion.
    """
    found = []
    for p in parts or []:
        if p.get("parts"):
            found.extend(_flatten_parts(p.get("parts")))
        else:
            found.append(p)
    return found

def _guess_extension_from_mime(mime: str):
    if not mime:
        return ""
    mime = mime.lower()
    if "pdf" in mime:
        return ".pdf"
    if "xml" in mime:
        return ".xml"
    if "jpeg" in mime or "jpg" in mime:
        return ".jpg"
    if "png" in mime:
        return ".png"
    return ""

def baixar_anexos_de_mensagem(service, msg_id: str) -> List[str]:
    """
    Baixa todos os anexos "reais" de uma mensagem (arquivos com filename ou attachmentId)
    e salva no DOWNLOAD_DIR. Retorna lista de caminhos salvos.
    Antes: apenas baixava XMLs/partes com xml. Agora baixa PDFs também (ex: boleto, DANFE).
    """
    saved = []
    try:
        message = service.users().messages().get(userId="me", id=msg_id, format="full").execute()
    except Exception as e:
        logger.exception("Erro ao obter mensagem %s: %s", msg_id, e)
        return saved

    payload = message.get("payload", {}) or {}
    parts = payload.get("parts", []) or []
    # Caso mensagem não seja multipart, considere payload como uma única parte
    all_parts = _flatten_parts(parts) if parts else [payload]

    if not all_parts:
        logger.debug("Nenhuma parte encontrada na mensagem %s", msg_id)
        return saved

    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    for idx, part in enumerate(all_parts, start=1):
        filename = part.get("filename") or ""
        mime = (part.get("mimeType") or "").lower()
        body = part.get("body", {}) or {}

        # 🔍 Baixe apenas PDFs ou XMLs
        if not (filename.lower().endswith(".pdf") or filename.lower().endswith(".xml")):
            # também aceita anexos que o Gmail não nomeia, mas que têm MIME pdf/xml
            if "pdf" not in mime and "xml" not in mime:
                continue

        # garante extensão
        if not filename:
            ext = ".pdf" if "pdf" in mime else ".xml" if "xml" in mime else ".bin"
            filename = f"{msg_id}_{idx}{ext}"

        # evita sobrescrever
        file_path = os.path.join(DOWNLOAD_DIR, f"{msg_id}_{idx}_{filename}")

        try:
            data_b = None
            if body.get("data"):
                # partes pequenas inline (às vezes XML simples)
                data_b = _decode_base64_fixed(body.get("data", ""))
            elif body.get("attachmentId"):
                attach_id = body["attachmentId"]
                attach = service.users().messages().attachments().get(
                    userId="me", messageId=msg_id, id=attach_id
                ).execute()
                raw = attach.get("data")
                if not raw:
                    continue
                data_b = _decode_base64_fixed(raw or "")
            else:
                continue

            # grava o arquivo
            with open(file_path, "wb") as fh:
                fh.write(data_b)
            saved.append(file_path)
            time.sleep(0.1)

        except Exception as e:
            logger.exception("Erro ao baixar anexo (%s): %s", filename, e)
            if os.path.exists(file_path):
                os.remove(file_path)
            continue


    logger.debug("Baixados %d anexos para mensagem %s", len(saved), msg_id)
    return saved

def _decode_base64_fixed(data: str) -> bytes:
    """Decodifica base64 corrigindo padding e caracteres urlsafe."""
    if not data:
        return b""
    data = data.strip()
    # Corrige caracteres urlsafe (- e _)
    data = data.replace("-", "+").replace("_", "/")
    # Corrige padding ausente
    missing_padding = len(data) % 4
    if missing_padding:
        data += "=" * (4 - missing_padding)
    return base64.b64decode(data, validate=False)


def _botana_label_ids(service) -> Tuple[Dict[str, str], List[str]]:
    labels = _list_labels(service)
    id_to_name = {}
    botana_ids = []
    for item in labels:
        label_id = str(item.get("id", "")).strip()
        name = str(item.get("name", "")).strip()
        if not label_id or not name:
            continue
        id_to_name[label_id] = name
        if any(name.startswith(prefix) for prefix in BOTANA_LABEL_PREFIXES):
            botana_ids.append(label_id)
    return id_to_name, botana_ids


def listar_mensagens_com_labels_botana(service, max_results: int = 1000) -> List[Dict[str, Any]]:
    labels = list_botana_labels(service)
    if not labels:
        return []
    wanted = max(1, min(1000, int(max_results or 1000)))
    seen = set()
    out: List[Dict[str, Any]] = []
    for label in labels:
        label_id = str(label.get("id", "")).strip()
        label_name = str(label.get("name", "")).strip()
        if not label_id:
            continue
        page_token = None
        while len(out) < wanted:
            batch_size = min(500, wanted - len(out))
            req_kwargs = {"userId": "me", "labelIds": [label_id], "maxResults": batch_size}
            if page_token:
                req_kwargs["pageToken"] = page_token
            resp = service.users().messages().list(**req_kwargs).execute()
            batch = resp.get("messages", []) or []
            for item in batch:
                msg_id = str(item.get("id", "")).strip()
                if not msg_id or msg_id in seen:
                    continue
                seen.add(msg_id)
                out.append({"id": msg_id, "threadId": item.get("threadId", ""), "botana_label": label_name})
                if len(out) >= wanted:
                    break
            page_token = str(resp.get("nextPageToken", "")).strip()
            if not page_token or not batch:
                break
        if len(out) >= wanted:
            break
    return out


def marcar_mensagem_com_label(service, msg_id: str, label_name: str | None = None, existing_label_ids=None, reprocessed: bool | None = None, when=None):
    try:
        id_to_name, botana_ids = _botana_label_ids(service)
        if reprocessed is None:
            existing = set(existing_label_ids or [])
            reprocessed = any(
                str(id_to_name.get(label_id, "")).startswith(REPROCESS_LABEL_NAME)
                for label_id in existing
            )
        target_name = str(label_name or build_botana_label_name(reprocessed=bool(reprocessed), when=when)).strip()
        target_id = ensure_label(service, target_name)
        remove_ids = [label_id for label_id in botana_ids if label_id != target_id]
        body = {"addLabelIds": [target_id]}
        if remove_ids:
            body["removeLabelIds"] = remove_ids
        service.users().messages().modify(userId="me", id=msg_id, body=body).execute()
        return target_name
    except Exception as e:
        logger.exception("Falha ao marcar mensagem %s com label: %s", msg_id, e)
        return None


def marcar_mensagem_para_reprocessar(service, msg_id: str, when=None, mark_unread: bool = True):
    try:
        _, botana_ids = _botana_label_ids(service)
        target_name = build_botana_label_name(reprocessed=True, when=when)
        target_id = ensure_label(service, target_name)
        remove_ids = [label_id for label_id in botana_ids if label_id != target_id]
        add_ids = [target_id]
        if mark_unread:
            add_ids.append("UNREAD")
        body = {"addLabelIds": add_ids}
        if remove_ids:
            body["removeLabelIds"] = remove_ids
        service.users().messages().modify(userId="me", id=msg_id, body=body).execute()
        return target_name
    except Exception as e:
        logger.exception("Falha ao marcar mensagem %s para reprocessamento: %s", msg_id, e)
        return None
