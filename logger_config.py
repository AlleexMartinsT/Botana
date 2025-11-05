# logger_config.py

import logging
import colorlog
from colorlog.escape_codes import escape_codes

# Handler colorido
handler = colorlog.StreamHandler()
handler.setFormatter(colorlog.ColoredFormatter(
    "%(log_color)s%(asctime)s [%(levelname)s] %(message)s",
    log_colors={
        "INFO": "green",
        "WARNING": "yellow",
        "ERROR": "red",
        "DEBUG": "blue"
    }
))

# Configuração base
logging.basicConfig(level=logging.INFO, handlers=[handler])

# Cria o logger principal
logger = logging.getLogger("bot.main")

# Cores manuais (pra mensagens específicas)
cor_ciano = escape_codes["cyan"]
cor_roxo = escape_codes["purple"]
cor_vermelho = escape_codes["red"]
reset = escape_codes["reset"]
