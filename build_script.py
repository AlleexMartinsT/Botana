"""
build_script.py — gera o executável base para empacotamento no Setup,
incluindo navegador Chromium do Playwright para funcionar no ambiente empacotado.
"""
import subprocess
import sys
import os
import shutil
from pathlib import Path
import glob

APP_NAME = "Botana"
MAIN_SCRIPT = "main.py"

print("\n=== [1] Iniciando build do Botana ===\n")

# === Localizações ===
BASE_DIR = Path(__file__).resolve().parent
DIST_DIR = BASE_DIR / "dist" / APP_NAME
MS_PLAYWRIGHT_DIR = Path(os.getenv("USERPROFILE", "")) / "AppData" / "Local" / "ms-playwright"

# === Verifica se Chromium está instalado ===
print("[*] Verificando se Playwright Chromium está instalado...")
if not MS_PLAYWRIGHT_DIR.exists():
    print("⚠️  Playwright Chromium não encontrado.")
    print("➡️  Execute antes do build:  python -m playwright install chromium\n")
else:
    print(f"✅ Chromium encontrado em: {MS_PLAYWRIGHT_DIR}")

# === Monta comando do PyInstaller ===
command = [
    sys.executable,
    "-m", "PyInstaller",
    "--noconfirm",
    "--onedir",
    "--noconsole",
    "--name", APP_NAME,
    "--add-data", f"secrets{os.pathsep}secrets",
    MAIN_SCRIPT
]

print("[*] Executando PyInstaller...\n")
subprocess.run(command, check=True)

# === Copia automaticamente o Chromium do Playwright ===
print("\n[2] Incluindo navegador Chromium do Playwright no pacote...")

# === Copia navegador do Playwright ===
if MS_PLAYWRIGHT_DIR.exists():
    print(f"\n[2] Copiando Chromium para pasta do app...")
    dest_base = DIST_DIR / "_internal" / "playwright"
    browsers_dest = dest_base / "driver" / "package" / ".local-browsers"
    browsers_dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(MS_PLAYWRIGHT_DIR, browsers_dest, dirs_exist_ok=True)
    print(f"✅ Chromium incluído em: {browsers_dest}")

    # Corrige versão do Chromium (renomeia automaticamente)
    for folder in browsers_dest.glob("chromium-*"):
        # Pega o número da versão atual instalada
        ver = folder.name.split("-")[-1]
        # O Playwright pode procurar por uma versão antiga — criamos um alias simbólico ou cópia
        expected_folder = browsers_dest / "chromium-1187"
        if not expected_folder.exists():
            shutil.copytree(folder, expected_folder)
            print(f"📦 Criado alias 'chromium-1187' apontando para versão real ({ver})")

else:
    print("⚠️  Chromium não foi copiado (pasta ms-playwright não encontrada).")

# === Finalização ===
print("\n✅ Build concluído com sucesso!")
print(f"📂 Pasta gerada: {DIST_DIR}")
print("\n💡 Dica: rode o app e verifique se o navegador interno abre corretamente.\n")
