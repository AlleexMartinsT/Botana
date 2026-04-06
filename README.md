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

## Histórico no painel

- A aba `Histórico` exibe uma grade compacta, em linha única por registro, para facilitar leitura dentro do Hub.
- A coluna `Cliente` usa uma versão resumida da `Descrição`, preservando o número do `BLT` quando existir.
- As colunas `Descrição`, `Valor Pago` e `Status` não aparecem mais na grade principal.
- Os títulos ficam centralizados, sem marcadores visuais de ordenação, e as divisórias da grade ficaram mais evidentes.
- As colunas podem ser reajustadas arrastando a divisória do cabeçalho, com opção de resetar as larguras no próprio painel.
- Os filtros visíveis agora acompanham a grade: `Data/Horário`, `Vencimento`, `NF`, `Cliente` e `Aba`.
- O botão `Excluir` remove somente o registro do histórico/relatório (`relatorios/relatorio_*.txt`). Ele não apaga a linha original da planilha.
