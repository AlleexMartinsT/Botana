import xml.etree.ElementTree as ET
import datetime, re
from config import CNPJ_MVA, CNPJ_EH

def _normalize_date_to_ddmmyyyy(date_raw):
    """Tenta normalizar várias entradas de data para 'DD/MM/YYYY'. Retorna '' se falhar."""
    if not date_raw:
        return ""
    # já no formato DD/MM/YYYY?
    candidates = [
        date_raw.strip(),
        date_raw.strip().replace(".", "/"),
        date_raw.strip().replace("-", "/")
    ]
    formats = ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%S", "%d-%m-%Y", "%d.%m.%Y")
    for cand in candidates:
        for fmt in formats:
            try:
                dt = datetime.datetime.strptime(cand, fmt)
                return dt.strftime("%d/%m/%Y")
            except Exception:
                continue
    # última tentativa com fromisoformat
    try:
        dt = datetime.datetime.fromisoformat(date_raw.replace("Z", "+00:00"))
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return ""

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
        "cnpjEmitente": re.sub(r"\D", "", emit.findtext("ns:CNPJ", default="", namespaces=ns) or ""),
        "destinatario": dest.findtext("ns:xNome", default="", namespaces=ns),
        "valorTotal": float(total.findtext("ns:vNF", default="0", namespaces=ns) or 0),
        "parcelas": [],
        "naturezaOperacao": ide.findtext("ns:natOp", default="", namespaces=ns).strip().upper(),
    }

    nat_op = ide.findtext("ns:natOp", default="", namespaces=ns).strip().upper()

    # Ignora se o destinatário for nossa própria empresa (pelo CNPJ)
    cnpj_dest = dest.findtext("ns:CNPJ", default="", namespaces=ns)
    cnpj_dest = re.sub(r"\D", "", cnpj_dest or "")  # 👈 ajuste: evita erro se None

    forma_pag = str(dados.get("formaPagamento", "")).strip()
    if ( "VISTA" in nat_op or "VENDA A VISTA" in nat_op or forma_pag in ["01", "03", "04"]):
        return dados

    if cnpj_dest in (re.sub(r"\D", "", CNPJ_MVA), re.sub(r"\D", "", CNPJ_EH)):
        return dados

    fat = root.findall(".//ns:dup", ns)
    fat_fatura = root.find(".//ns:fat", ns)

    if fat:
        # Caso normal — há duplicatas
        for i, dup in enumerate(fat, start=1):
            venc_raw = dup.findtext("ns:dVenc", default="", namespaces=ns)
            venc = _normalize_date_to_ddmmyyyy(venc_raw)
            valor = float(dup.findtext("ns:vDup", default="0", namespaces=ns) or 0)
            dados["parcelas"].append({
                "numero": i,
                "numParcela": f"{i}ª Parcela",
                "vencimento": venc,   # agora em DD/MM/YYYY
                "valor": valor
            })
    else:
        # ⚠️ Fallback: usa <fat> se não houver <dup>
        if fat_fatura is not None:
            valor = float(fat_fatura.findtext("ns:vLiq", default="0", namespaces=ns) or 0)
            emissao = ide.findtext("ns:dhEmi", default="", namespaces=ns)
            venc = ""
            try:
                data_emissao = datetime.datetime.fromisoformat(emissao.replace("Z", "+00:00"))
                venc = (data_emissao + datetime.timedelta(days=30)).strftime("%d/%m/%Y")
            except Exception:
                # tenta normalizar emissao mesmo que não tenha Z
                venc = _normalize_date_to_ddmmyyyy(emissao)
                if not venc:
                    venc = ""
            dados["parcelas"].append({
                "numero": 1,
                "numParcela": "1ª Parcela",
                "vencimento": venc,
                "valor": valor
            })

    # Quantidade total de parcelas
    dados["qtdParcelas"] = len(dados["parcelas"]) or 1

    # Ano de vencimento (para definir planilha) — pega o ano da primeira parcela quando possível
    if dados["parcelas"]:
        try:
            ano = datetime.datetime.strptime(dados["parcelas"][0]["vencimento"], "%d/%m/%Y").year
        except Exception:
            ano = datetime.datetime.now().year
    else:
        ano = datetime.datetime.now().year
    dados["anoVencimento"] = str(ano)

    # Descrição default = nome do destinatário + número da NF
    dados["descricao"] = f"{dados['destinatario']} BLT {dados['nf']}"

    # Campos simplificados para preenchimento na planilha:
    if dados["parcelas"]:
        p = dados["parcelas"][0]
        dados["vencimento"] = p["vencimento"]          # em DD/MM/YYYY
        dados["numParcela"] = p["numParcela"]
        dados["valorParcela"] = p["valor"]
    else:
        dados["vencimento"] = ""
        dados["numParcela"] = "1ª Parcela"
        dados["valorParcela"] = dados["valorTotal"]

    return dados
