import gspread
import logging
from datetime import datetime
import locale
import os
import re
import time
import unicodedata

from config import PLANILHAS
from logger_config import logger, cor_ciano, reset

# Garante que os meses saiam em portugues (ex: Fev/2025)
os.environ["LANG"] = "pt_BR.UTF-8"
try:
    locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")
except locale.Error:
    locale.setlocale(locale.LC_TIME, "ptb")  # fallback Windows

logger = logging.getLogger("bot.sheets_writer")


def apiCooldown():
    logger.warning("Limite da API atingido, aguardando 30 segundos...")
    time.sleep(30)


def _parse_date_any(date_str):
    """Tenta varios formatos e retorna datetime ou None."""
    if not date_str:
        return None
    formats = (
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%Y-%m-%dT%H:%M:%S%z",
        "%Y-%m-%dT%H:%M:%S",
        "%d-%m-%Y",
        "%d.%m.%Y",
    )
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except Exception:
            continue
    try:
        return datetime.fromisoformat(str(date_str).replace("Z", "+00:00"))
    except Exception:
        return None


def _result(ok, inserted, reason, **extra):
    return {
        "ok": bool(ok),
        "inserted": bool(inserted),
        "duplicate": bool(reason in {"duplicate", "duplicate_nf_full"}),
        "reason": str(reason or ""),
        "sheet_title": str(extra.get("sheet_title", "") or ""),
        "sheet_type": str(extra.get("sheet_type", "") or ""),
        "aba": str(extra.get("aba", "") or ""),
        "vencimento": str(extra.get("vencimento", "") or ""),
        "descricao": str(extra.get("descricao", "") or ""),
        "nf": str(extra.get("nf", "") or ""),
        "valor_total": float(extra.get("valor_total", 0.0) or 0.0),
        "qtd_parcelas": int(extra.get("qtd_parcelas", 0) or 0),
        "parcela": str(extra.get("parcela", "") or ""),
        "valor_parcela": float(extra.get("valor_parcela", 0.0) or 0.0),
        "valor_pago": str(extra.get("valor_pago", "") or ""),
        "status": str(extra.get("status", "") or ""),
    }


def _normalize_key_text(value):
    txt = unicodedata.normalize("NFKD", str(value or "")).encode("ascii", "ignore").decode("ascii")
    txt = re.sub(r"\s+", " ", txt).strip().upper()
    return txt


def _safe_float(value):
    try:
        if isinstance(value, str):
            vv = value.strip().replace("\xa0", " ")
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
        return float(value or 0)
    except Exception:
        return 0.0


def _extract_parcela_number(value):
    txt = _normalize_key_text(value)
    if not txt:
        return None
    patterns = (
        r"\b(\d+)\s*(?:A|O)?\s*PARC(?:ELA)?\b",
        r"\bPARC(?:ELA)?\s*(\d+)\b",
        r"\b(\d+)\s*/\s*\d+\b",
        r"\b(\d+)\s+DE\s+\d+\b",
        r"\b(\d+)\b",
    )
    for pattern in patterns:
        match = re.search(pattern, txt)
        if not match:
            continue
        try:
            return int(match.group(1))
        except Exception:
            continue
    return None


def _parcel_identity(parcela, vencimento="", valor_parcela=0):
    numero = _extract_parcela_number(parcela)
    if numero is not None:
        return f"parcela:{numero}"
    data = _parse_date_any(vencimento)
    venc = data.strftime("%d/%m/%Y") if data else str(vencimento or "").strip()
    valor = round(_safe_float(valor_parcela), 2)
    if venc or valor:
        return f"fallback:{venc}|{valor:.2f}"
    raw = _normalize_key_text(parcela)
    return f"raw:{raw}" if raw else ""


def _read_worksheet_values(worksheet):
    for _ in range(3):
        try:
            return worksheet.get_values()
        except gspread.exceptions.APIError as exc:
            if "429" in str(exc):
                apiCooldown()
                continue
            raise
    return []


def _has_matching_parcela(rows, nf, parcel_identity):
    for row in rows:
        if len(row) < 7 or str(row[2] or "").strip() != nf:
            continue
        row_identity = _parcel_identity(row[5], row[0], row[6])
        if parcel_identity and row_identity == parcel_identity:
            return True
    return False


def _next_financial_row(rows):
    for row_number, row in enumerate(rows[1:], start=2):
        if not any(str(cell or "").strip() for cell in row[:9]):
            return row_number
    return max(2, len(rows) + 1)


def atualizarPlanilha(planilha, dados, gc):
    """
    Atualiza a planilha Google Sheets com os dados extraidos do XML.
    Retorna um dict com o resultado para registro no historico.
    """

    vencimento_raw = dados.get("vencimento")
    if not vencimento_raw:
        logger.warning("XML sem data de vencimento - ignorado.")
        return _result(False, False, "missing_vencimento")

    dataVenc = _parse_date_any(vencimento_raw)
    if not dataVenc:
        logger.warning("Data invalida no XML: %s", vencimento_raw)
        return _result(False, False, "invalid_vencimento")

    venc_str = dataVenc.strftime("%d/%m/%Y")
    nomeAba = dataVenc.strftime("%b/%Y").capitalize()
    ano = str(dataVenc.year)

    nome_planilha_upper = planilha.title.upper() if hasattr(planilha, "title") else ""
    tipo = None

    for t, anos in PLANILHAS.items():
        for _, sheet_id in anos.items():
            if sheet_id and sheet_id in planilha.url:
                tipo = t
                break
        if tipo:
            break

    if not tipo:
        if "MVA" in nome_planilha_upper:
            tipo = "MVA"
        elif "HORIZONTE" in nome_planilha_upper or "EH" in nome_planilha_upper:
            tipo = "EH"

    if not tipo:
        logger.error("Tipo de planilha nao identificado a partir do titulo: %s", nome_planilha_upper)
        return _result(False, False, "unknown_sheet_type")

    id_correto = PLANILHAS.get(tipo, {}).get(ano)
    if id_correto:
        if id_correto not in planilha.url:
            try:
                planilha_nova = gc.open_by_key(id_correto)
                logger.info("Redirecionado para '%s' (%s) - Tipo %s", planilha_nova.title, ano, tipo)
                return atualizarPlanilha(planilha_nova, dados, gc)
            except Exception as e:
                logger.error("Erro ao abrir planilha de %s (%s): %s", ano, tipo, e)
                return _result(False, False, "redirect_open_error")
    else:
        logger.warning("Nenhuma planilha configurada para %s %s. Mantendo planilha atual.", tipo, ano)

    descricao = str(dados.get("descricao", "") or "").strip()
    if tipo in ("MVA", "EH") and "(BOT)" not in descricao.upper():
        descricao = f"{descricao} (Bot)"

    nf = str(dados.get("nf", "") or "").strip()
    qtd_parcelas = int(dados.get("qtdParcelas", 1) or 1)
    parcela = str(dados.get("numParcela", "1ª Parcela") or "1ª Parcela")
    valor_total = float(dados.get("valorTotal", 0) or 0)
    valor_parcela = float(dados.get("valorParcela", 0) or 0)
    valor_pago = ""
    status = ""

    try:
        aba = planilha.worksheet(nomeAba)
    except gspread.exceptions.WorksheetNotFound:
        logger.warning("Criando nova aba: %s", nomeAba)
        aba = planilha.add_worksheet(title=nomeAba, rows="100", cols="9")
        aba.append_row(
            [
                "Vencimento",
                "Descrição",
                "NF",
                "Valor Total",
                "Qtd Parcelas",
                "Parcela",
                "Valor Parcela",
                "Valor Pago",
                "Status",
            ]
        )

    linhas = _read_worksheet_values(aba)

    incoming_identity = _parcel_identity(parcela, venc_str, valor_parcela)
    nf_existing_identities = set()
    duplicado = False

    for linha in linhas:
        if len(linha) < 7 or str(linha[2] or "").strip() != nf:
            continue
        existing_identity = _parcel_identity(linha[5], linha[0], linha[6])
        if existing_identity:
            nf_existing_identities.add(existing_identity)
        if incoming_identity and existing_identity == incoming_identity:
            duplicado = True
            break

    nf_completo = bool(nf and qtd_parcelas > 0 and len(nf_existing_identities) >= qtd_parcelas)

    base_payload = {
        "sheet_title": str(getattr(planilha, "title", "") or ""),
        "sheet_type": tipo,
        "aba": nomeAba,
        "vencimento": venc_str,
        "descricao": descricao,
        "nf": nf,
        "valor_total": valor_total,
        "qtd_parcelas": qtd_parcelas,
        "parcela": parcela,
        "valor_parcela": valor_parcela,
        "valor_pago": valor_pago,
        "status": status,
    }

    if duplicado:
        logger.warning("NF %s (%s) ja existe em %s pela mesma parcela estrutural.", nf, venc_str, nomeAba)
        return _result(True, False, "duplicate", **base_payload)
    if nf_completo:
        logger.warning(
            "NF %s ja possui %d parcela(s) estruturalmente registrada(s) em %s. Bloqueando novo lancamento.",
            nf,
            len(nf_existing_identities),
            nomeAba,
        )
        return _result(True, False, "duplicate_nf_full", **base_payload)

    novaLinha = [
        venc_str,
        descricao,
        nf,
        valor_total,
        qtd_parcelas,
        parcela,
        valor_parcela,
        valor_pago,
        status,
    ]

    for _ in range(3):
        try:
            target_range = f"A{_next_financial_row(linhas)}:I{_next_financial_row(linhas)}"
            aba.update(
                [novaLinha],
                range_name=target_range,
                value_input_option="USER_ENTERED",
            )
            linhas_confirmadas = _read_worksheet_values(aba)
            if not _has_matching_parcela(linhas_confirmadas, nf, incoming_identity):
                logger.error(
                    "NF %s nao foi confirmada na aba %s apos o append; historico nao sera gravado.",
                    nf,
                    nomeAba,
                )
                return _result(False, False, "append_unverified", **base_payload)
            logger.info(f"{cor_ciano}NF {nf} registrada em '{planilha.title}' / aba '{nomeAba}'{reset}")
            return _result(True, True, "inserted", **base_payload)
        except gspread.exceptions.APIError as e:
            if "429" in str(e):
                apiCooldown()
                continue
            raise e

    return _result(False, False, "append_failed", **base_payload)
