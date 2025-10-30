# reporter.py
import os
import time
from pathlib import Path
from datetime import datetime

from config import RELATORIO_DIR as relatorioDir

# Variáveis de relatório de sessão
eventosProcessados = []
eventosIgnorados = []
historicoEventos = set()

RELATORIO_TXT = "relatorio_status.txt"
RELATORIO_TEMP = "relatorio_temp.tmp"
ultimoRelatorio = {"Conta Principal": None}

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
    """
    Lê o relatório atual e retorna um conjunto de NFs já registradas.
    """
    relatorio_path = Path(relatorioDir) / f"relatorio_{datetime.now():%Y-%m-%d}.txt"

    if not relatorio_path.exists():
        return set()  # se não existe, ainda não há nada registrado

    nfs_existentes = set()
    with open(relatorio_path, "r", encoding="utf-8") as f:
        for linha in f:
            if "NF" in linha:
                # extrai o número da NF (qualquer sequência numérica após 'NF')
                partes = linha.split("NF")
                if len(partes) > 1:
                    nf_num = partes[1].split()[0]  # pega o que vem logo depois
                    nfs_existentes.add(nf_num.strip())

    return nfs_existentes
