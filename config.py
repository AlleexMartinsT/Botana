import os
from dotenv import load_dotenv

# Carregar variáveis de ambiente do arquivo .env dentro de secrets/
dotenv_path = os.path.join(os.path.dirname(__file__), "secrets", ".env")
load_dotenv(dotenv_path)


# Caminhos
BASE_DIR = os.path.dirname(__file__)
SECRETS_DIR = os.path.join(BASE_DIR, "secrets")
DOWNLOAD_DIR = os.path.join(BASE_DIR, "xmls_baixados")
RELATORIO_DIR = os.path.join(BASE_DIR, "relatorios")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(RELATORIO_DIR, exist_ok=True)

# Credenciais
GOOGLE_CREDENTIALS_GMAIL = os.path.join(SECRETS_DIR, os.getenv("GOOGLE_CREDENTIALS_GMAIL"))
GOOGLE_CREDENTIALS_SHEETS = os.path.join(SECRETS_DIR, os.getenv("GOOGLE_CREDENTIALS_SHEETS"))

# Planilhas
PLANILHAS = {
    "MVA": {
        "2025": os.getenv("SHEET_MVA_2025"),
        "2026": os.getenv("SHEET_MVA_2026")
    },
    "EH": {
        "2025": os.getenv("SHEET_EH_2025"),
        "2026": os.getenv("SHEET_EH_2026")
    }
}

# CNPJs
CNPJ_MVA = os.getenv("CNPJ_MVA")
CNPJ_EH = os.getenv("CNPJ_EH")

# Intervalo
INTERVALO = int(os.getenv("INTERVALO", "600"))
