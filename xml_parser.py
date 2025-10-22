import xml.etree.ElementTree as ET
import datetime, re
from config import CNPJ_MVA, CNPJ_EH

def extrairDadosXML(caminhoXML):
    tree = ET.parse(caminhoXML)
    root = tree.getroot()
    # Corrige se for um nfeProc (envolve a NFe dentro)
    if root.tag.endswith("nfeProc"):
        root = root.find(".//ns:NFe", {"ns": "http://www.portalfiscal.inf.br/nfe"})
    ns = {"ns": "http://www.portalfiscal.inf.br/nfe"}

    ide = root.find(".//ns:ide", ns)
    emit = root.find(".//ns:emit", ns)
    dest = root.find(".//ns:dest", ns)
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
    cnpj_dest = re.sub(r"\D", "", cnpj_dest or "")  # 👈 ajuste: evita erro se None

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

    fat = root.findall(".//ns:dup", ns)
    fat_fatura = root.find(".//ns:fat", ns)

    if fat:
        # Caso normal — há duplicatas
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
    else:
        # ⚠️ Fallback: usa <fat> se não houver <dup>
        if fat_fatura is not None:
            valor = float(fat_fatura.findtext("ns:vLiq", default="0", namespaces=ns))
            emissao = ide.findtext("ns:dhEmi", default="", namespaces=ns)
            try:
                data_emissao = datetime.datetime.fromisoformat(emissao.replace("Z", "+00:00"))
                # Define vencimento como 30 dias após emissão
                venc = (data_emissao + datetime.timedelta(days=30)).strftime("%d/%m/%Y")
            except Exception:
                venc = ""
            dados["parcelas"].append({
                "numero": 1,
                "numParcela": "1ª Parcela",
                "vencimento": venc,
                "valor": valor
            })

    # Quantidade total de parcelas
    dados["qtdParcelas"] = len(dados["parcelas"]) or 1

    # Ano de vencimento (para definir planilha)
    if dados["parcelas"]:
        try:
            ano = datetime.datetime.strptime(dados["parcelas"][0]["vencimento"], "%d/%m/%Y").year
        except Exception:
            ano = datetime.datetime.now().year  # 👈 ajuste: se vencimento vier vazio
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
