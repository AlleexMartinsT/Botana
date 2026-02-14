# Botana

Botana é um serviço de processamento de e-mails/XML que pode rodar:

- em modo local (tray)
- em modo servidor HTTP (integração com FinanceHub)

## Requisitos

- Python 3.12+
- Git
- Credenciais em `secrets/` conforme `config.py`

## Instalação local

```powershell
cd <BOTANA_DIR>
python -m venv .venv
.\.venv\Scripts\python -m pip install --upgrade pip
.\.venv\Scripts\python -m pip install -r requirements.txt
```

## Execução via HUB (modo servidor)

```powershell
cd <BOTANA_DIR>
.\.venv\Scripts\python main.py --server --host 127.0.0.1 --port 8865
```

Sem iniciar o loop automático:

```powershell
.\.venv\Scripts\python main.py --server --host 127.0.0.1 --port 8865 --no-loop
```

## Endpoints de saúde/controle

- `GET /api/state`
- `POST /api/start`
- `POST /api/stop`
- `POST /api/run-now`

## Integração com FinanceHub

No `instances.json` do Hub, use:

- `instance_type: "botana"`
- `backend_url: "http://127.0.0.1:8865"`
- `app_dir: "C:\\Botana"`
- `route_prefix: "botana"`
- `repo_url: "https://github.com/AlleexMartinsT/Botana.git"`
- `auto_clone_missing: true`
- `start_args: ["main.py","--server","--host","127.0.0.1","--port","8865"]`
