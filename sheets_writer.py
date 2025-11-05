import gspread
import logging
from datetime import datetime
import locale, os
import time
from googleapiclient.errors import HttpError
from config import PLANILHAS
from logger_config import logger, cor_ciano, reset

# Garante que os meses saiam em português (ex: Fev/2025)
os.environ["LANG"] = "pt_BR.UTF-8"
try:
    locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")
except locale.Error:
    locale.setlocale(locale.LC_TIME, "ptb")  # fallback Windows

logger = logging.getLogger("bot.sheets_writer")

def apiCooldown():
    logger.warning("⏳ Limite da API atingido, aguardando 30 segundos...")
    time.sleep(30)

def _parse_date_any(date_str):
    """Tenta vários formatos e retorna datetime ou None."""
    if not date_str:
        return None
    formats = (
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%Y-%m-%dT%H:%M:%S%z",
        "%Y-%m-%dT%H:%M:%S",
        "%d-%m-%Y",
        "%d.%m.%Y"
    )
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except Exception:
            continue
    # tentativa final com fromisoformat (aceita Z -> +00:00)
    try:
        return datetime.fromisoformat(date_str.replace("Z", "+00:00"))
    except Exception:
        return None


def atualizarPlanilha(planilha, dados, gc):
    """
    Atualiza a planilha Google Sheets com os dados extraídos do XML.
    Cria automaticamente a aba do mês/ano caso não exista.
    Se a data de vencimento for de outro ano (ex: 2026), muda para a planilha correspondente.
    """

    vencimento_raw = dados.get("vencimento")
    if not vencimento_raw:
        logger.warning("⚠️ XML sem data de vencimento — ignorado.")
        return

    dataVenc = _parse_date_any(vencimento_raw)
    if not dataVenc:
        logger.warning(f"⚠️ Data inválida no XML: {vencimento_raw}")
        return

    # padroniza para DD/MM/YYYY
    venc_str = dataVenc.strftime("%d/%m/%Y")
    nomeAba = dataVenc.strftime("%b/%Y").capitalize()
    ano = str(dataVenc.year)

    # Detecta tipo de planilha (MVA ou EH)
    nome_planilha_upper = planilha.title.upper() if hasattr(planilha, "title") else ""
    tipo = None

    for t, anos in PLANILHAS.items():
        for ano_cadastrado, sheet_id in anos.items():
            if sheet_id and sheet_id in planilha.url:
                tipo = t
                break
        if tipo:
            break

    # fallback por nome
    if not tipo:
        if "MVA" in nome_planilha_upper:
            tipo = "MVA"
        elif "HORIZONTE" in nome_planilha_upper or "EH" in nome_planilha_upper:
            tipo = "EH"

    if not tipo:
        logger.error(f"❌ Tipo de planilha não identificado a partir do título: {nome_planilha_upper}")
        return

    # ================================
    # 🔁 Redireciona automaticamente se o ano mudou
    # ================================
    id_correto = PLANILHAS.get(tipo, {}).get(ano)

    if id_correto:
        if id_correto not in planilha.url:
            try:
                planilha_nova = gc.open_by_key(id_correto)
                logger.info(f"✅ Redirecionado para '{planilha_nova.title}' ({ano}) - Tipo {tipo}")
                return atualizarPlanilha(planilha_nova, dados, gc)  # recursão segura
            except Exception as e:
                logger.error(f"❌ Erro ao abrir planilha de {ano} ({tipo}): {e}")
                return
    else:
        logger.warning(f"⚠️ Nenhuma planilha configurada para {tipo} {ano}. Mantendo planilha atual.")

    # ================================
    # 📝 Insere linha na planilha atual
    # ================================
    descricao = dados.get("descricao", "")
    if tipo in ("MVA", "EH"):
        if "(BOT)" not in descricao.upper():
            descricao = f"{descricao} (Bot)"

    # Tenta acessar a aba, se não existir cria
    try:
        aba = planilha.worksheet(nomeAba)
    except gspread.exceptions.WorksheetNotFound:
        logger.warning(f"🆕 Criando nova aba: {nomeAba}")
        aba = planilha.add_worksheet(title=nomeAba, rows="100", cols="9")
        aba.append_row([
            "Vencimento", "Descrição", "NF", "Valor Total", "Qtd Parcelas",
            "Parcela", "Valor Parcela", "Valor Pago", "Status"
        ])

    # Tenta obter todas as linhas (com retry por API limit)
    for _ in range(3):
        try:
            linhas = aba.get_all_values()
            break
        except gspread.exceptions.APIError as e:
            if "429" in str(e):
                apiCooldown()
                continue
            else:
                raise e

    # Evita duplicados — compara Vencimento + NF + Parcela + Descrição
    duplicado = any(
        len(linha) >= 6 and
        linha[0] == venc_str and
        linha[2] == str(dados.get("nf", "")) and
        linha[5] == dados.get("numParcela", "1ª Parcela") and
        linha[1] == descricao
        for linha in linhas
    )

    if duplicado:
        logger.warning(f"⚠️ NF {dados.get('nf')} ({venc_str}) já existe em {nomeAba}.")
        return

    novaLinha = [
        venc_str,
        descricao,
        dados.get("nf", ""),
        f"R$ {float(dados.get('valorTotal', 0)):.2f}",
        dados.get("qtdParcelas", 1),
        dados.get("numParcela", "1ª Parcela"),
        f"R$ {float(dados.get('valorParcela', 0)):.2f}",
        "",
        ""
    ]

    # Insere no Google Sheets
    for _ in range(3):
        try:
            aba.append_row(novaLinha, value_input_option="USER_ENTERED")
            logger.info(f"{cor_ciano}✅ NF {dados.get('nf')} registrada em '{planilha.title}' / aba '{nomeAba}'{reset}")
            break
        except gspread.exceptions.APIError as e:
            if "429" in str(e):
                apiCooldown()
                continue
            else:
                raise e
