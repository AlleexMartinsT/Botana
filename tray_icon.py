"""
Tray App do Botana
-----------------------
Mostra ícone na bandeja com menu:
  • Verificar agora
  • Abrir relatórios
  • Sair
Inclui indicador de status por cor:
  🔵 Azul = Ocioso
  🟢 Verde = Verificando
  🔴 Vermelho = Erro
E notificações do sistema.
"""

import os, sys
import threading
import time
import traceback
from pathlib import Path
import pystray
from PIL import Image, ImageDraw
from plyer import notification

from config import RELATORIO_DIR
from reporter import escreverRelatorio
from gmail_service import buscarMessagesEnviados

# =========================
# ÍCONE DINÂMICO
# =========================

def create_icon(color: str = "blue"):
    """Cria um ícone circular colorido com fundo transparente."""
    color_map = {
        "blue": (0, 128, 255, 255),
        "green": (0, 200, 0, 255),
        "red": (220, 0, 0, 255)
    }
    rgb = color_map.get(color, (0, 128, 255, 255))
    image = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.ellipse((8, 8, 56, 56), fill=rgb)
    return image

# =========================
# NOTIFICAÇÕES
# =========================
def notificar(titulo, mensagem):
    """Exibe uma notificação do sistema."""
    try:
        notification.notify(
            title=titulo,
            message=mensagem,
            timeout=5
        )
    except Exception as e:
        print(f"[Tray] Falha ao exibir notificação: {e}")

# =========================
# TRAY APP
# =========================
def run_tray(on_quit_callback, start_callback=None):
    """
    Inicia o ícone de bandeja com menus:
      - Verificar agora
      - Abrir relatórios
      - Sair
    """
    icon = pystray.Icon("Botana", title="Botana")
    status_lock = threading.Lock()
    status_color = {"value": "blue"}  # azul inicial (ocioso)

    def atualizar_cor(cor):
        with status_lock:
            status_color["value"] = cor
        icon.icon = create_icon(cor)
        icon.visible = True

    def executar_verificacao():
        """Executa verificação manual dos e-mails."""
        try:
            atualizar_cor("green")
            notificar("Botana", "Iniciando verificação manual...")
            print("\n[Manual] Verificação solicitada pelo usuário.")
            buscarMessagesEnviados("Conta Principal")
            atualizar_cor("blue")
            notificar("Botana", "✅ Verificação concluída com sucesso!")
            print("[Tray] Verificação concluída.")
        except Exception as e:
            msg = f"Erro durante verificação: {e}"
            print(f"[Tray] {msg}")
            traceback.print_exc()
            escreverRelatorio(f"[Tray] {msg}")
            atualizar_cor("red")
            notificar("Botana", f"❌ {msg}")
            time.sleep(10)
            atualizar_cor("blue")

    def verificar_agora(icon, item):
        """Aciona a verificação manual ou inicia o loop do main."""
        if start_callback:
            # Se o main.py passou um callback, inicia o loop principal
            print("[Tray] Iniciando loop principal via callback.")
            atualizar_cor("green")
            threading.Thread(target=start_callback, daemon=True).start()
        else:
            # Caso contrário, executa verificação manual única
            threading.Thread(target=executar_verificacao, daemon=True).start()

    def abrir_relatorios(icon, item):
        caminho = Path(RELATORIO_DIR).resolve()
        if not caminho.exists():
            os.makedirs(caminho, exist_ok=True)
        if sys.platform == "win32":
            os.startfile(caminho)
        elif sys.platform == "darwin":
            os.system(f"open '{caminho}'")
        else:
            os.system(f"xdg-open '{caminho}'")

    def sair(icon, item):
        notificar("Botana", "Encerrando o aplicativo...")
        icon.visible = False
        icon.stop()
        on_quit_callback()

    # Menu
    menu = pystray.Menu(
        pystray.MenuItem("Verificar agora", verificar_agora),
        pystray.MenuItem("Abrir relatórios", abrir_relatorios),
        pystray.MenuItem("Sair", sair)
    )

    icon.icon = create_icon("blue")
    icon.menu = menu
    print("[Tray] Ícone iniciado. Clique com o botão direito para opções.")
    icon.run()


# =========================
# TESTE MANUAL
# =========================

if __name__ == "__main__":
    def sair():
        print("Encerrando manualmente.")
        sys.exit(0)
    run_tray(sair)
