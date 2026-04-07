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
- `POST /api/recover-missing`
- `GET /api/conferencia-parcelas`
- `GET /api/prazos`
- `GET /api/prazos/search`
- `GET /api/prazos/search-suggestions`

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

- O botão `Reprocessar agora` inicia a ação em background, devolve resposta visual imediata e dispara a leitura no mesmo fluxo.
- O card `Recuperar e-mails sem leitura` busca mensagens enviadas com XML que ainda não tenham label do Botana, usando período e/ou faixa de NF, e já relê essas mensagens na sequência.
- O card `Ações manuais` mostra estado, mensagem, detalhe e progresso da ação atual.
- Os resumos de `Executar agora`, `Reprocessar agora` e `Recuperar faltantes` mostram também quantas parcelas caíram como `Duplicadas`, para diferenciar releitura bem-sucedida de lançamento realmente novo.
- Durante a ação manual, o botão fica desabilitado para evitar execuções concorrentes.
- Se o loop automático estiver ativo, `Reprocessar agora` interrompe esse ciclo, remarca as mensagens, executa a leitura manual e retoma o monitoramento no fim.
- O reprocessamento agora usa apenas `Limite de mensagens`, buscando as mensagens mais recentes ainda marcadas com a label do Botana.
- Depois de remarcar as labels, o Botana relê exatamente as mensagens selecionadas nesse reprocessamento, em vez de depender só do lote padrão do ciclo automático.
- A recuperação de faltantes usa os mesmos contadores visuais de `Ações manuais`, mostrando quantas mensagens foram analisadas, quantas combinaram com os filtros e qual e-mail está sendo varrido no momento.
- As labels do Gmail passaram a incluir a data no formato `DD/MM/AAAA`; no fluxo normal viram `XML Processado Botana - 07/04/2026` e, ao reprocessar, mudam para `XML Reprocessado Botana - 07/04/2026`.
- Durante o reprocessamento, o painel mostra a fase atual, quantas mensagens já foram tratadas, quantas falharam e o e-mail/data da mensagem atual.
- As duas barras de progresso passam a refletir o reprocessamento ativo, incluindo o limite pedido no painel.
- As mensagens de status do ciclo manual e automático no painel usam acentuação PT-BR correta.
- Os cards `Configuração do Gmail`, `Autenticação` e `Reprocessar e-mails` usam altura natural no grid principal, sem forçar a mesma altura entre si.

## Leitura do Gmail

- O ciclo normal agora aplica de verdade o `Período` configurado no painel ao montar a busca do Gmail.
- A leitura automática passou a usar paginação real, respeitando `Máx páginas` e `Tamanho da página`, em vez de ficar presa ao primeiro lote retornado pela API.
- Mensagens que já tenham qualquer label do Botana são puladas no ciclo normal, evitando releitura desnecessária e liberando espaço para e-mails ainda não tratados.
- O limite efetivo de leitura automática fica alinhado ao menos ao produto de `Máx páginas x Tamanho da página`, para não manter um corte antigo menor do que o configurado na UI.
- Quando o XML vier só com a fatura total (`fat`) mas o e-mail trouxer vários boletos PDF, o Botana tenta extrair vencimento e valor de cada boleto para reconstruir as parcelas antes de lançar no financeiro.

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

- A aba `Conferência` lê diretamente as planilhas e resume por NF se a quantidade registrada bate com a quantidade esperada nas parcelas.
- A varredura pode ser feita por `Mês do lançamento`, por `Faixa de NF` ou em `Tudo`.
- No filtro por mês, a aba seleciona as NFs relacionadas ao mês escolhido e confere a NF inteira nas planilhas, em vez de depender do histórico.
- O painel destaca NFs com parcelas faltando, duplicadas ou acima da quantidade esperada, com totais agregados no topo.
- A tabela da conferência não mostra mais a coluna `Parcelas`.
- A conferência atualiza automaticamente ao abrir a aba, mostra estado de carregamento e informa quando a leitura das planilhas terminou.
- O cabeçalho da conferência fica centralizado e as colunas `Status`, `NF`, `Esperadas`, `Lançadas`, `Faltando` e `Duplicadas` usam largura mais compacta.

## Prazos no painel

- A aba `Prazos` lê diretamente as planilhas e lista apenas títulos com `Status` vazio ou `A Receber`.
- Linhas marcadas como `BAIXADO`, `BAIXADA`, `ESTORNADO` ou `ESTORNADA` em qualquer coluna relevante da linha, inclusive quando o texto cair fora da coluna de `Status`, são ignoradas na aba `Prazos`.
- O painel separa boletos próximos do vencimento e depósitos atrasados, usando dias úteis e filtros customizáveis de `1` a `7`.
- Por padrão, boletos entram quando vencem em até `7` dias úteis e depósitos entram quando estão atrasados há pelo menos `7` dias úteis.
- Linhas amarelas representam boletos que ainda vão vencer; linhas vermelhas representam itens que vencem hoje ou já passaram da data.
- A aba atualiza ao abrir, mostra estado de carregamento e resume totais de boletos a vencer, boletos no limite e depósitos atrasados.
- Os botões de atualização da relação e da busca por nome ficam em linhas separadas, com espaçamento próprio no card da aba `Prazos`.
- O botão `Buscar boletos em aberto` abre um popup próprio, fora do front-end nativo do navegador, para consultar por nome do cliente.
- O campo `Nome do cliente` do popup fica centralizado e usa uma lista própria de sugestões dentro do modal, em vez do `datalist` nativo do navegador.
- Se não houver pendências para o nome buscado, o popup mostra apenas a mensagem informando que não existem boletos em aberto.
- O autocomplete mistura os nomes salvos localmente em `%APPDATA%\\Botana\\watch_search_names.txt` com os nomes de clientes que ainda têm boletos em aberto nas planilhas, então buscas parciais como `LOCAR` passam a sugerir nomes maiores correspondentes.

## Exportação do histórico

- A aba `Histórico` agora tem um botão `Exportar CSV`.
- A exportação respeita os mesmos filtros e o mesmo limite aplicados na consulta atual da grade.

## Ajustes de Prazos

- A coluna `Valor` prioriza `Valor da Parcela` e aceita células formatadas como moeda no Google Sheets; se vier vazio, cai para `Valor Total`.
- Os filtros da aba `Prazos` ficam empilhados, centralizados e com campos menores.
- Ao recarregar a aba `Prazos`, os totais são zerados antes da nova leitura para não manter números antigos na tela.

## Duplicidade estrutural

- O writer do Botana agora bloqueia novo lançamento quando a NF já tem todas as parcelas esperadas preenchidas, mesmo que a nova descrição venha como `DEP` em vez de `BLT`.
- A conferência também normaliza o número da parcela (`2ª`, `2º`, `Parcela 2`, `2/2`) para não deixar duplicidades estruturais passarem só por diferença de texto.
- Na aba `Conferência`, casos marcados como `Duplicada` passam a aparecer como divergência em vermelho.
