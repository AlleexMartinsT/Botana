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
- `POST /api/reprocess`
- `GET /api/conferencia-parcelas`

## Integração com FinanceHub

No `instances.json` do Hub, use:

- `instance_type: "botana"`
- `backend_url: "http://127.0.0.1:8865"`
- `app_dir: "C:\\Botana"`
- `route_prefix: "botana"`
- `repo_url: "https://github.com/AlleexMartinsT/Botana.git"`
- `auto_clone_missing: true`
- `start_args: ["main.py","--server","--host","127.0.0.1","--port","8865"]`

## Ações manuais no painel

- Os botões `Executar agora` e `Marcar para reprocessar` agora iniciam a ação em background e devolvem resposta visual imediata.
- O card `Ações manuais` mostra estado, mensagem, detalhe e progresso da ação atual.
- Durante uma ação manual, os dois botões ficam desabilitados para evitar execuções concorrentes.
- A execução manual usa os contadores de leitura/lotes já exibidos no painel para mostrar avanço do ciclo.
- Se o loop automático estiver ativo, `Executar agora` interrompe esse ciclo, roda a execução manual e retoma o monitoramento no fim.
- Se o loop automático estiver ativo, `Marcar para reprocessar` também interrompe esse ciclo, reprocessa e retoma o monitoramento no fim.
- O reprocessamento agora usa apenas `Limite de mensagens`, buscando as mensagens mais recentes ainda marcadas com a label do Botana.
- As labels do Gmail passaram a incluir a data no formato `DD/MM/AAAA`; no fluxo normal viram `XML Processado Botana - 07/04/2026` e, ao reprocessar, mudam para `XML Reprocessado Botana - 07/04/2026`.
- Durante o reprocessamento, o painel mostra quantas mensagens já foram tratadas, quantas falharam e o e-mail/data da mensagem atual.
- As duas barras de progresso passam a refletir o reprocessamento ativo, incluindo o limite pedido no painel.
- As mensagens de status do ciclo manual e automático no painel usam acentuação PT-BR correta.
- Os cards `Configuração do Gmail`, `Autenticação` e `Reprocessar e-mails` usam altura natural no grid principal, sem forçar a mesma altura entre si.

## Histórico no painel

- A aba `Histórico` exibe uma grade compacta, em linha única por registro, para facilitar leitura dentro do Hub.
- A coluna `Cliente` usa uma versão resumida da `Descrição`, preservando o número do `BLT` quando existir.
- As colunas `Descrição`, `Valor Pago` e `Status` não aparecem mais na grade principal.
- Os títulos ficam centralizados, sem marcadores visuais de ordenação, e as divisórias da grade ficaram mais evidentes.
- As colunas podem ser reajustadas arrastando a divisória do cabeçalho, com opção de resetar as larguras no próprio painel.
- Os filtros visíveis agora acompanham a grade: `Data/Horário`, `Vencimento`, `NF`, `Cliente` e `Aba`.
- Registros com NF/parcela duplicadas ficam destacados em vermelho para facilitar a identificação.
- O botão `Excluir` remove somente o registro do histórico/relatório (`relatorios/relatorio_*.txt`). Ele não apaga a linha original da planilha.

## Conferência de parcelas

- A aba `Conferência` resume por NF se a quantidade de parcelas lançadas bate com a quantidade esperada registrada no XML.
- A varredura pode ser feita por `Mês do lançamento`, por `Faixa de NF` ou em `Tudo`.
- A conferência passa a ler também o fallback `relatorio_*.txt.tmp`, então a busca reflete lançamentos recém-gravados mesmo quando o arquivo principal está ocupado.
- O painel destaca NFs com parcelas faltando, duplicadas ou acima da quantidade esperada, com totais agregados no topo.
- A tabela da conferência não mostra mais a coluna `Parcelas`.
- A conferência usa os eventos `HIST_JSON` gravados em `relatorios/relatorio_*.txt`, sem apagar ou regravar linhas da planilha.
