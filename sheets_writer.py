import gspread
import logging
from datetime import datetime
import locale, os
import time
from googleapiclient.errors import HttpError

# Garante que os meses saiam em português (ex: Fev/2025)
os.environ["LANG"] = "pt_BR.UTF-8"
try:
    locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")
except locale.Error:
    locale.setlocale(locale.LC_TIME, "ptb")  # fallback Windows

logger = logging.getLogger("bot.sheets_writer")

def apiCooldown():
    print("⏳ Limite da API atingido, aguardando 30 segundos...")
    time.sleep(30)

def atualizarPlanilha(planilha, dados):
    """
    Atualiza a planilha Google Sheets com os dados extraídos do XML.
    Cria automaticamente a aba do mês/ano caso não exista.
    """

    vencimento = dados.get("vencimento")
    if not vencimento:
        print("⚠️ XML sem data de vencimento — ignorado.")
        return

    try:
        dataVenc = datetime.strptime(vencimento, "%Y-%m-%d")
    except ValueError:
        print(f"⚠️ Data inválida no XML: {vencimento}")
        return

    # Exemplo: "Nov/2025"
    nomeAba = dataVenc.strftime("%b/%Y").capitalize()

    # Tenta acessar a aba, se não existir cria
    try:
        aba = planilha.worksheet(nomeAba)
    except gspread.exceptions.WorksheetNotFound:
        print(f"🆕 Criando nova aba: {nomeAba}")
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

    # Evita duplicados
    duplicado = any(
        linha[0] == dataVenc.strftime("%d/%m/%Y")
        and linha[2] == str(dados["nf"])
        for linha in linhas if len(linha) >= 3
    )

    if duplicado:
        print(f"⚠️ NF {dados['nf']} ({vencimento}) já existe em {nomeAba}.")
        return

    nome_planilha = planilha.title.upper()
    descricao = dados.get("descricao", "")

    if "MVA" in nome_planilha or "EH" in nome_planilha:
        if "(BOT)" not in descricao.upper():
            descricao = f"{descricao} (Bot)"
    
    # Nova linha com todos os campos
    novaLinha = [
        dataVenc.strftime("%d/%m/%Y"),
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

            # Nome da planilha completa e aba
            nome_planilha = planilha.title
            empresa = "MVA" if "MVA" in nome_planilha.upper() else "EH"
            nome_aba = nomeAba
            
            logger.info(f"✅ NF {dados['nf']} registrada em '{nome_planilha}' / aba '{nome_aba}'\n")
            break
        except gspread.exceptions.APIError as e:
            if "429" in str(e):
                apiCooldown()
                continue
            else:
                raise e
