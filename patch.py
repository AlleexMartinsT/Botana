import codecs
import re

file_path = r'c:\Users\TI\Desktop\Botana\main.py'
with codecs.open(file_path, 'r', 'utf-8') as f:
    content = f.read()

new_api = """                if parsed.path == "/api/clean-sheets":
                    if not _can_operate(user):
                        return _json_response(self, 403, {"ok": False, "message": "Sem permissão"})
                    try:
                        import correcao_planilhas
                        correcao_planilhas.iniciar_assistente_em_background()
                        return _json_response(self, 200, {"ok": True, "friendly": "Assistente de Limpeza iniciado com sucesso."})
                    except Exception as e:
                        return _json_response(self, 500, {"ok": False, "friendly": f"Falha ao iniciar o assistente: {e}"})

                return _json_response(self, 404, {"ok": False, "message": "N"""

new_content = content.replace('                return _json_response(self, 404, {"ok": False, "message": "N', new_api)

with codecs.open(file_path, 'w', 'utf-8') as f:
    f.write(new_content)

print("Patch aplicado com sucesso em main.py!")
