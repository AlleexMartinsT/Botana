# reporter.py
import os
import time
from datetime import datetime

from config import RELATORIO_DIR, RELATORIO_DIR as relatorioDir
from config import RELATORIO_DIR as REL_DIR  # redundância para compatibilidade
from config import RELATORIO_DIR as RELATORIO_DIR_dup

# Variáveis de relatório de sessão
eventosProcessados = []
eventosIgnorados = []
historicoEventos = set()

RELATORIO_TXT = "relatorio_status.txt"
RELATORIO_TEMP = "relatorio_temp.tmp"
ultimoRelatorio = {"Conta Principal": None, "Conta NFe": None}

def limparRelatoriosAntigos():
    agora = datetime.now()
    for arquivo in os.listdir(relatorioDir):
        caminho = os.path.join(relatorioDir, arquivo)
        if os.path.isfile(caminho):
            modificacao = datetime.fromtimestamp(os.path.getmtime(caminho))
            if (agora - modificacao).days > 7:
                os.remove(caminho)

def obterArquivoRelatorio():
    dataHoje = datetime.now().strftime("%Y-%m-%d")
    return os.path.join(relatorioDir, f"relatorio_{dataHoje}.txt")

def escreverRelatorio(texto):
    arquivoRelatorio = obterArquivoRelatorio()
    try:
        with open(arquivoRelatorio, "a", encoding="utf-8") as f:
            f.write(texto + "\n")
    except PermissionError:
        with open(arquivoRelatorio + ".tmp", "a", encoding="utf-8") as f:
            f.write(texto + "\n")

def registrarEvento(tipo, fornecedor, conta):
    if fornecedor.strip() in ["-", ""]:
        return
    if any(x in fornecedor.upper() for x in [
        "ELETRONICA HORIZONTE COMERCIO DE PRODUTOS ELETRONICOS LTDA",
        "MVA COMERCIO DE PRODUTOS ELETRONICOS LTDA EPP"
    ]):
        return

    if tipo == "processado":
        eventosProcessados.append((fornecedor, conta))
    elif tipo == "ignorado":
        eventosIgnorados.append((fornecedor, conta))

def consolidarRelatorioTMP():
    """Move conteúdo do .tmp para o .txt quando possível."""
    if os.path.exists(RELATORIO_TEMP):
        try:
            with open(RELATORIO_TEMP, "r", encoding="utf-8") as f_temp:
                dados_temp = f_temp.read()
            with open(RELATORIO_TXT, "a", encoding="utf-8") as f_rel:
                f_rel.write(dados_temp)
            os.remove(RELATORIO_TEMP)
            print("Relatório temporário consolidado com sucesso.")
        except PermissionError:
            pass
