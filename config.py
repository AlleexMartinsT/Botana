import os
from dotenv import load_dotenv

from sheet_registry import sheet_ids_from_environment

# Carrega variaveis de ambiente do arquivo .env dentro de secrets/
dotenv_path = os.path.join(os.path.dirname(__file__), "secrets", ".env")
load_dotenv(dotenv_path)

# Caminhos
BASE_DIR = os.path.dirname(__file__)
SECRETS_DIR = os.path.join(BASE_DIR, "secrets")
DOWNLOAD_DIR = os.path.join(BASE_DIR, "xmls_baixados")
RELATORIO_DIR = os.path.join(BASE_DIR, "relatorios")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
os.makedirs(RELATORIO_DIR, exist_ok=True)


def _env_or_default(key: str, default_value: str) -> str:
    value = os.getenv(key)
    if value is None:
        return default_value
    value = str(value).strip()
    return value if value else default_value


# Credenciais
GOOGLE_CREDENTIALS_GMAIL = os.path.join(
    SECRETS_DIR,
    _env_or_default("GOOGLE_CREDENTIALS_GMAIL", "credentials_gmail.json"),
)
GOOGLE_CREDENTIALS_SHEETS = os.path.join(
    SECRETS_DIR,
    _env_or_default("GOOGLE_CREDENTIALS_SHEETS", "credentials_sheets.json"),
)

# Planilhas
PLANILHAS = sheet_ids_from_environment(os.environ)

# CNPJs
CNPJ_MVA = os.getenv("CNPJ_MVA")
CNPJ_EH = os.getenv("CNPJ_EH")

# Intervalo
INTERVALO = int(os.getenv("INTERVALO", "1800"))
