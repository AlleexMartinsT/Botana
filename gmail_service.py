# gmail_service.py
import os
import base64
import time
import logging
from typing import List, Dict, Any
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from config import GOOGLE_CREDENTIALS_GMAIL, DOWNLOAD_DIR


# Scopes: precisamos de modify para acrescentar labels (e opcionalmente marcar como lido)
SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]

logger = logging.getLogger("bot.gmail_service")
LABEL_NAME = "XML Processado Botana"

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

def marcar_mensagem_com_label(service, msg_id: str, label_name: str = LABEL_NAME):
    try:
        label_id = ensure_label(service, label_name)
        body = {"addLabelIds": [label_id]}
        service.users().messages().modify(userId="me", id=msg_id, body=body).execute()
    except Exception as e:
        logger.exception("Falha ao marcar mensagem %s com label: %s", msg_id, e)
