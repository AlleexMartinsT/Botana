import datetime
import re
import xml.etree.ElementTree as ET

from config import CNPJ_EH, CNPJ_MVA


def _normalize_date_to_ddmmyyyy(date_raw):
    """Tenta normalizar varias entradas de data para DD/MM/YYYY."""
    if not date_raw:
        return ""

    candidates = [
        str(date_raw).strip(),
        str(date_raw).strip().replace(".", "/"),
        str(date_raw).strip().replace("-", "/"),
    ]
    formats = (
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%Y-%m-%dT%H:%M:%S%z",
        "%Y-%m-%dT%H:%M:%S",
        "%d-%m-%Y",
        "%d.%m.%Y",
    )

    for cand in candidates:
        for fmt in formats:
            try:
                dt = datetime.datetime.strptime(cand, fmt)
                return dt.strftime("%d/%m/%Y")
            except Exception:
                continue

    try:
        dt = datetime.datetime.fromisoformat(str(date_raw).replace("Z", "+00:00"))
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return ""


def extrairDadosXML(caminhoXML):
    tree = ET.parse(caminhoXML)
    root = tree.getroot()

    if root.tag.endswith("nfeProc"):
        root = root.find(".//ns:NFe", {"ns": "http://www.portalfiscal.inf.br/nfe"})

    ns = {"ns": "http://www.portalfiscal.inf.br/nfe"}

    ide = root.find(".//ns:ide", ns)
    emit = root.find(".//ns:emit", ns)
    dest = root.find(".//ns:dest", ns)
    total = root.find(".//ns:ICMSTot", ns)

    cnpj_emit = re.sub(r"\D", "", emit.findtext("ns:CNPJ", default="", namespaces=ns) or "")
    cnpj_dest = re.sub(r"\D", "", dest.findtext("ns:CNPJ", default="", namespaces=ns) or "")

    dados = {
        "nf": ide.findtext("ns:nNF", default="", namespaces=ns),
        "emitente": emit.findtext("ns:xNome", default="", namespaces=ns),
        "cnpjEmitente": cnpj_emit,
        "destinatario": dest.findtext("ns:xNome", default="", namespaces=ns),
        "cnpjDestinatario": cnpj_dest,
        "valorTotal": float(total.findtext("ns:vNF", default="0", namespaces=ns) or 0),
        "parcelas": [],
        "parcelas_source": "none",
        "naturezaOperacao": ide.findtext("ns:natOp", default="", namespaces=ns).strip().upper(),
    }

    nat_op = ide.findtext("ns:natOp", default="", namespaces=ns).strip().upper()
    forma_pag = str(dados.get("formaPagamento", "")).strip()

    if "VISTA" in nat_op or "VENDA A VISTA" in nat_op or forma_pag in ["01", "03", "04"]:
        return dados

    if cnpj_dest in (re.sub(r"\D", "", CNPJ_MVA), re.sub(r"\D", "", CNPJ_EH)):
        return dados

    fat = root.findall(".//ns:dup", ns)
    fat_fatura = root.find(".//ns:fat", ns)

    if fat:
        dados["parcelas_source"] = "dup"
        for i, dup in enumerate(fat, start=1):
            venc_raw = dup.findtext("ns:dVenc", default="", namespaces=ns)
            venc = _normalize_date_to_ddmmyyyy(venc_raw)
            valor = float(dup.findtext("ns:vDup", default="0", namespaces=ns) or 0)
            dados["parcelas"].append(
                {
                    "numero": i,
                    "numParcela": f"{i}ª Parcela",
                    "vencimento": venc,
                    "valor": valor,
                }
            )
    else:
        if fat_fatura is not None:
            dados["parcelas_source"] = "fat"
            valor = float(fat_fatura.findtext("ns:vLiq", default="0", namespaces=ns) or 0)
            emissao = ide.findtext("ns:dhEmi", default="", namespaces=ns)
            venc = ""
            try:
                data_emissao = datetime.datetime.fromisoformat(str(emissao).replace("Z", "+00:00"))
                venc = (data_emissao + datetime.timedelta(days=30)).strftime("%d/%m/%Y")
            except Exception:
                venc = _normalize_date_to_ddmmyyyy(emissao)
                if not venc:
                    venc = ""

            dados["parcelas"].append(
                {
                    "numero": 1,
                    "numParcela": "1ª Parcela",
                    "vencimento": venc,
                    "valor": valor,
                }
            )

    dados["qtdParcelas"] = len(dados["parcelas"]) or 1

    if dados["parcelas"]:
        try:
            ano = datetime.datetime.strptime(dados["parcelas"][0]["vencimento"], "%d/%m/%Y").year
        except Exception:
            ano = datetime.datetime.now().year
    else:
        ano = datetime.datetime.now().year
    dados["anoVencimento"] = str(ano)

    dados["descricao"] = f"{dados['destinatario']} BLT {dados['nf']}"

    if dados["parcelas"]:
        p = dados["parcelas"][0]
        dados["vencimento"] = p["vencimento"]
        dados["numParcela"] = p["numParcela"]
        dados["valorParcela"] = p["valor"]
    else:
        dados["vencimento"] = ""
        dados["numParcela"] = "1ª Parcela"
        dados["valorParcela"] = dados["valorTotal"]

    return dados
