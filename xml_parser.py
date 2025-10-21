import xml.etree.ElementTree as ET
import datetime, re
from config import CNPJ_MVA, CNPJ_EH

def extrairDadosXML(caminhoXML):
    tree = ET.parse(caminhoXML)
    root = tree.getroot()
    ns = {"ns": "http://www.portalfiscal.inf.br/nfe"}

    ide = root.find(".//ns:ide", ns)
    emit = root.find(".//ns:emit", ns)
    dest = root.find(".//ns:dest", ns)
    fat = root.findall(".//ns:dup", ns)
    total = root.find(".//ns:ICMSTot", ns)

    dados = {
        "nf": ide.findtext("ns:nNF", default="", namespaces=ns),
        "emitente": emit.findtext("ns:xNome", default="", namespaces=ns),
        "cnpjEmitente": emit.findtext("ns:CNPJ", default="", namespaces=ns),
        "destinatario": dest.findtext("ns:xNome", default="", namespaces=ns),
        "valorTotal": float(total.findtext("ns:vNF", default="0", namespaces=ns)),
        "parcelas": [],
    }
    nat_op = ide.findtext("ns:natOp", default="", namespaces=ns).strip().upper()
    
    # Ignora se o destinatário for nossa própria empresa (pelo CNPJ)
    cnpj_dest = dest.findtext("ns:CNPJ", default="", namespaces=ns)
    cnpj_dest = re.sub(r"\D", "", cnpj_dest) 

    forma_pag = str(dados.get("formaPagamento", "")).strip()
    if (
        "VISTA" in nat_op
        or "VENDA A VISTA" in nat_op
        or forma_pag in ["01", "03", "04"]
    ):
        print(f"[DEBUG IGNORE RESULT] NF {dados['nf']} ignorada (venda à vista / cartão).")
        return None

    if cnpj_dest in (re.sub(r"\D", "", CNPJ_MVA), re.sub(r"\D", "", CNPJ_EH)):
        print(f"[DEBUG IGNORE RESULT] NF {dados['nf']} ignorada (destinatário é o nosso: {cnpj_dest})")
        return None

    # Extrai parcelas (duplicatas/faturas)
    for i, dup in enumerate(fat, start=1):
        venc = dup.findtext("ns:dVenc", default="", namespaces=ns)
        try:
            venc = datetime.datetime.strptime(venc, "%Y-%m-%d").strftime("%d/%m/%Y")
        except Exception:
            venc = venc or ""
        valor = float(dup.findtext("ns:vDup", default="0", namespaces=ns))
        dados["parcelas"].append({
            "numero": i,
            "numParcela": f"{i}ª Parcela",
            "vencimento": venc,
            "valor": valor
        })

    # Quantidade total de parcelas
    dados["qtdParcelas"] = len(dados["parcelas"]) or 1

    # Ano de vencimento (para definir planilha)
    if dados["parcelas"]:
        ano = datetime.datetime.strptime(dados["parcelas"][0]["vencimento"], "%d/%m/%Y").year
    else:
        ano = datetime.datetime.now().year
    dados["anoVencimento"] = str(ano)

    # Descrição = nome do destinatário + número da NF
    dados["descricao"] = f"{dados['destinatario']} BLT {dados['nf']}"

    # Campos simplificados para preenchimento na planilha
    if dados["parcelas"]:
        p = dados["parcelas"][0]
        try:
            dados["vencimento"] = datetime.datetime.strptime(p["vencimento"], "%d/%m/%Y").strftime("%Y-%m-%d")
        except Exception:
            dados["vencimento"] = ""
        dados["numParcela"] = p["numParcela"]
        dados["valorParcela"] = p["valor"]
    else:
        dados["vencimento"] = ""
        dados["numParcela"] = "1ª Parcela"
        dados["valorParcela"] = dados["valorTotal"]
    return dados
