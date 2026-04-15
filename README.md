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
- `POST /api/recover-emails`
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

- Quando o Hub abrir `/botana/login?popup=1`, o login do Botana fica em modo reduzido: não mostra o botão interno de voltar ao Hub e, após autenticar, confirma o acesso e fecha a janela sem redirecionar para o painel completo.

## Ações manuais no painel

- O botão `Reprocessar agora` inicia a ação em background, devolve resposta visual imediata e dispara a leitura no mesmo fluxo.
- O card `Recuperar e-mails` procura mensagens com XML pelos filtros informados e já executa a leitura em seguida para tentar lançar no financeiro.
- A recuperação agora tem três modos: `Período`, `Faixa de NF` e `Escolha manual`, em que você monta uma lista própria de NFs como `20247` e `20344`.
- O layout do card `Recuperar e-mails` ficou centralizado e agrupado em blocos, e o filtro ativo agora se recompõe automaticamente conforme o `Modo` escolhido.
- O card `Recuperar e-mails` ficou mais responsivo: em telas menores os blocos passam a empilhar melhor, e no modo `Faixa de NF` o `Enter` no primeiro campo avança para o segundo antes de confirmar a recuperação.
- Os campos preenchíveis de `Recuperar e-mails` ficaram mais compactos, o botão foi isolado em uma linha própria, e o campo visual de `Limite de mensagens` saiu da UI.
- A recuperação não depende mais de ausência de label; ela pode reler mensagens já marcadas pelo Botana, e o bloqueio de duplicidade continua acontecendo no writer da planilha.
- Quando o e-mail vier com boleto PDF válido mas o XML estiver sem parcelas, o Botana agora reconstrói a parcela diretamente do boleto e consegue lançar casos como a `NF 20571`.
- Na recuperação por NF, o Botana também passou a considerar o nome dos anexos quando o assunto/snippet vierem inconsistentes, então casos como `NF 20451` continuam encontráveis mesmo se o assunto do e-mail estiver errado.
- Quando essa recuperação precisar confiar nos anexos/XML porque o assunto do e-mail não bate com a NF real, o Botana agora marca esse aviso no estado final para o Hub alertar o usuário.
- Na recuperação por NF ou faixa, o Botana não aceita mais a mensagem só porque o assunto bate; se os anexos/XML apontarem outra NF, a mensagem é descartada para evitar falso positivo de `já lançada`.
- Quando isso acontecer, a NF continua marcada como não localizada na recuperação e o estado final passa a registrar a divergência para o Hub alertar que o assunto/PDF não bate com o XML.
- O card `Ações manuais` mostra estado, mensagem, detalhe e progresso da ação atual.
- Os resumos de `Executar agora`, `Reprocessar agora` e `Recuperar e-mails` priorizam progresso, falhas e janela do lote, sem expor contadores técnicos de anexos, XML, lançamentos ou duplicidades no painel.
- Durante a ação manual, o botão fica desabilitado para evitar execuções concorrentes.
- Se o loop automático estiver ativo, `Reprocessar agora` interrompe esse ciclo, remarca as mensagens, executa a leitura manual e retoma o monitoramento no fim.
- O reprocessamento agora usa apenas `Limite de mensagens`, buscando por padrão as mensagens mais recentes ainda marcadas com a label do Botana nas últimas duas semanas.
- A opção `Marcar como não lido` foi removida do painel porque o reprocessamento não depende do estado de leitura do Gmail para funcionar.
- A explicação desse fluxo no card `Reprocessar e-mails` fica sob um ícone pequeno de `?`, em vez de texto fixo ocupando espaço no painel.
- Os cards `Reprocessar e-mails` e `Recuperar e-mails` agora mostram esse `?` ao lado do próprio título, em vez de deixar a ajuda solta abaixo dos campos.
- O balão de ajuda desses `?` abre alinhado ao ícone e foi ajustado para não ser cortado pela lateral do painel.
- Para mensagens mais antigas do que essa janela padrão, o caminho recomendado é `Recuperar e-mails`.
- O reprocessamento monta o lote olhando todas as labels do Botana dentro dessa janela, para que labels datadas mais novas não fiquem escondidas atrás da label antiga `XML Processado Botana`.
- Depois de remarcar as labels, o Botana relê exatamente as mensagens selecionadas nesse reprocessamento, em vez de depender só do lote padrão do ciclo automático.
- A recuperação de faltantes usa a mesma área visual de `Ações manuais`, mostrando o andamento da varredura e quantas mensagens combinaram com os filtros.
- As labels do Gmail passaram a incluir a data no formato `DD/MM/AAAA`; no fluxo normal viram `XML Processado Botana - 07/04/2026` e, ao reprocessar, mudam para `XML Reprocessado Botana - 07/04/2026`.
- Durante o reprocessamento, o painel mostra a fase atual, quantas mensagens já foram tratadas, quantas falharam e a data do item atual, sem repetir label anterior nem o e-mail completo na linha de resumo.
- O reprocessamento também informa a janela real coberta pelo lote selecionado, mostrando de qual data mais recente até qual data mais antiga ele conseguiu ir com o limite pedido.
- Ao concluir um reprocessamento com mensagens mais antigas ainda disponíveis, o painel abre um popup automático perguntando se você quer continuar do próximo lote; se confirmar, basta informar a quantidade adicional.
- As duas barras de progresso passam a refletir o reprocessamento ativo, incluindo o limite pedido no painel.
- As mensagens de status do ciclo manual e automático no painel usam acentuação PT-BR correta.
- No resumo de `Últimos processados`, linhas consecutivas idênticas do ciclo deixam de aparecer em flood, passam a ser agrupadas com contador (`2x`, `3x` etc.) e a acentuação dos resumos do ciclo é normalizada antes de entrar na lista.
- As mensagens de ignorar NF por `venda à vista` também passaram a sair com acentuação correta no relatório/painel.
- Os cards `Configuração do Gmail`, `Autenticação` e `Reprocessar e-mails` usam altura natural no grid principal, sem forçar a mesma altura entre si.

- No modo `Escolha manual`, a recuperação passou a medir o progresso pelo total de NFs pedidas, em vez de começar com um teto genérico como `0/1000`.
- Quando a busca manual localiza apenas parte das NFs solicitadas, o resultado final informa quais NFs foram encontradas e quais ficaram faltando.
- A recuperação manual por lista não dispara mais erro interno de variável local ao iniciar a busca; o fluxo agora conclui normalmente e mostra se a NF foi lançada ou barrada como duplicada.

## Leitura do Gmail

- O ciclo normal agora aplica de verdade o `Período` configurado no painel ao montar a busca do Gmail.
- A leitura automática passou a usar paginação real, respeitando `Máx páginas` e `Tamanho da página`, em vez de ficar presa ao primeiro lote retornado pela API.
- Mensagens que já tenham qualquer label do Botana são puladas no ciclo normal, evitando releitura desnecessária e liberando espaço para e-mails ainda não tratados.
- O limite efetivo de leitura automática fica alinhado ao menos ao produto de `Máx páginas x Tamanho da página`, para não manter um corte antigo menor do que o configurado na UI.
- Quando o XML vier só com a fatura total (`fat`) mas o e-mail trouxer vários boletos PDF, o Botana tenta extrair vencimento e valor de cada boleto para reconstruir as parcelas antes de lançar no financeiro.
- Esse fallback de boletos PDF também cobre XMLs sem parcelas, desde que os boletos do e-mail tragam vencimento e valor compatíveis com a NF.
- Quando o XML vier sem parcelas e o e-mail indicar `DEPÓSITO` no assunto ou no corpo, o Botana usa esse sinal para montar um lançamento único de `DEP`, reaproveitando o valor total do XML e a data de emissão como fallback de vencimento.
- Mesmo quando o XML vier como `VENDA À VISTA`, o Botana não ignora mais a mensagem se o próprio e-mail indicar `DEPÓSITO` ou `BOLETO`; nesses casos ele deixa o fallback completar a parcela antes de decidir o lançamento.

## Histórico no painel

- A aba `Histórico` exibe uma grade compacta, em linha única por registro, para facilitar leitura dentro do Hub.
- A coluna `Cliente` usa uma versão resumida da `Descrição`, preservando o número do `BLT` quando existir.
- As colunas `Descrição`, `Valor Pago` e `Status` não aparecem mais na grade principal.
- Os títulos ficam centralizados, sem marcadores visuais de ordenação, e as divisórias da grade ficaram mais evidentes.
- As colunas podem ser reajustadas arrastando a divisória do cabeçalho, com opção de resetar as larguras no próprio painel.
- Os filtros visíveis agora acompanham a grade: um campo de busca único que muda conforme o tipo selecionado (Data, Data/Hora, Vencimento, NF, Cliente, Aba ou Sem Filtro).
- A aba `Histórico` agora também tem o filtro de origem por ícone (`MVA`, `EH` ou ambas), seguindo a mesma lógica visual da `Conferência`.
- A grade real do `Histórico` agora usa `Tabulator`, com paginação local, filtros por coluna, resize de colunas e fallback para a tabela HTML antiga se a biblioteca não carregar.
- Quando uma nova consulta do `Histórico` começa, a grade anterior fica coberta por uma tela de loading, em vez de manter os dados velhos aparentes até a resposta chegar.
- O rótulo curto de `Cliente` passou a ignorar prefixos numéricos no começo da descrição, como CNPJs abreviados (`34.826.916 ...`), e quando o texto terminar só com o número do boleto ele mostra `(... BLT)` sem carregar esse número bruto para a grade.
- Registros com NF/parcela duplicadas ficam destacados em vermelho para facilitar a identificação.
- A aba `Histórico` não expõe mais exclusão direta; a limpeza de sobras e duplicidades deve ser feita pela `Conferência`, que atua sobre a planilha real.
- O toolbar do `Histórico` ficou mais enxuto e mantém apenas o botão de resetar larguras da grade.

## Conferência de parcelas

- A aba `Conferência` lê diretamente as planilhas e resume por NF se a quantidade registrada bate com a quantidade esperada nas parcelas.
- A varredura pode ser feita por `Mês do lançamento`, por `Faixa de NF` ou em `Tudo`.
- No filtro por mês, a aba seleciona as NFs relacionadas ao mês escolhido e confere a NF inteira nas planilhas, em vez de depender do histórico.
- O painel destaca NFs com parcelas faltando, duplicadas ou acima da quantidade esperada, com totais agregados no topo.
- Ao clicar no badge de `Status` de uma NF com divergência, o Botana pede confirmação e limpa direto na planilha apenas as linhas excedentes/duplicadas identificadas automaticamente para aquela NF, sem reordenar as demais linhas.
- Depois da limpeza pela `Conferência`, a linha não é recarregada imediatamente; ela fica marcada localmente em amarelo, com sublinhado, e o badge deixa de ser clicável para evitar uma segunda remoção acidental antes da próxima conferência.
- Cliques em sequência no `Status` da Conferência entram em uma fila de até `3s` e são enviados em lote, usando o snapshot recém-carregado da aba em vez de reler todas as planilhas a cada exclusão.
- A tabela da conferência não mostra mais a coluna `Parcelas`.
- A conferência atualiza automaticamente ao abrir a aba, mostra estado de carregamento fora da tabela e informa quando a leitura das planilhas terminou.
- No filtro por `Faixa de NF`, a `Conferência` também passa a incluir linhas sintéticas de `NF ausente` quando algum número do intervalo não existe na planilha, com `-` nos campos textuais para não poluir a leitura da grade.
- No filtro por `Mês do lançamento`, a `Conferência` também cruza divergências conhecidas do Gmail e pode promover NFs ausentes para o topo quando assunto/PDF e XML apontarem números diferentes.
- No modo mensal, essa conferência cruza apenas lacunas curtas entre NFs já vistas na planilha e consulta o Gmail de forma pontual/cacheada para evitar leituras longas e erros `502`.
- Leituras seguidas de `Conferência` e `Prazos` reaproveitam por alguns segundos o snapshot recém-lido das planilhas, reduzindo rate limit do Google Sheets sem mudar a lógica da tela.
- A `Conferência` agora tem um filtro de origem por ícone (`MVA`, `EH` ou ambas): por padrão ela abre em `MVA`, deixa `EH` aquecendo em segundo plano e reutiliza esse cache quando você troca a origem.
- A zebra da `Conferência` acompanha a origem selecionada: `MVA` usa branco com laranja claro, `EH` usa branco com azul claro, e em `Todas` cada linha passa a herdar a cor da empresa de origem.
- A `Conferência` agora roda em job/snapshot próprio: a aba pode mostrar um resultado parcial assim que as planilhas da origem principal terminam de ler e continua finalizando os diagnósticos em segundo plano.
- Ao iniciar uma nova `Conferência`, a grade anterior fica coberta por uma tela de loading até o primeiro lote novo do job entrar na tela, evitando a leitura visual do resultado anterior como se ele ainda estivesse válido.
- Quando a `Drive API` estiver disponível nas credenciais do serviço, a `Conferência` reaproveita snapshots enquanto `modifiedTime/version` das planilhas não mudarem; no ambiente atual, se a `Drive API` estiver desativada, ela cai automaticamente para um snapshot local curto só para evitar releitura imediata ao sair e voltar da aba.
- Durante as atualizações parciais/finais da `Conferência`, a aba agora preserva a posição atual da tela para não ficar puxando a viewport para cima enquanto o job ainda está carregando.
- Os campos de `NF inicial` e `NF final` da `Conferência` só aparecem quando o modo ativo é `Faixa de NF`, evitando espaço morto na linha de filtros.
- No modo mensal, a `Conferência` não marca mais como `NF ausente` uma NF que já exista em outro mês da mesma empresa; isso corrige casos como a `49511`, cujo e-mail é de abril mas o lançamento foi para `Mai/2026` por causa do vencimento.
- Quando a `NF ausente` puder ser explicada pelo Gmail, a linha mostra esse motivo no hover; badges de `Status` só ficam clicáveis quando há linhas reais para limpar.
- As colunas da `Conferência` podem ser clicadas para ordenar a grade em ordem crescente ou decrescente sem perder o conteúdo já carregado.
- No filtro `Faixa de NF` da `Conferência`, o primeiro `Enter` em `NF inicial` agora avança para `NF final`, e só depois confirma a busca.
- Os textos explicativos longos de `Conferência` e `Prazos` passaram a ficar em `?` ao lado do título da seção, para a interface ficar mais limpa.
- A linha de filtros da `Conferência` agora usa blocos com largura e espaçamento uniformes, incluindo `Origem` e `Conferir parcelas`, para evitar desalinhamento visual entre os controles.
- O botão `Conferir parcelas` agora reserva a mesma altura-base dos campos com label, então ele fica alinhado com `Modo`, `Mês` e `Origem` em vez de subir visualmente na linha.
- O cabeçalho da conferência fica centralizado e as colunas `Status`, `NF`, `Parc.`, `Faltando` e `Duplicadas` usam largura mais compacta; `Parc.` mostra `esperadas / lançadas` na mesma célula.
- Os textos e campos da grade da `Conferência` ficam centralizados; a coluna `Aba` ficou mais compacta, com fonte menor, para liberar mais espaço visual para `Cliente` e para as demais colunas principais.
- Duplicatas extras da mesma parcela voltam a ficar limpáveis pelo badge de `Status` mesmo quando a sobra já estiver `Pago` ou `BAIXADO`, desde que exista outra linha mantida como referência da parcela.

- Existe uma rota isolada de preview em `/preview/tabulator/conferencia` para comparar uma versÃ£o Tabulator da `ConferÃªncia`, com paginaÃ§Ã£o local, filtros por coluna e ordenaÃ§Ã£o dinÃ¢mica, sem substituir a aba atual do painel.

- A aba real da `ConferÃªncia` agora usa `Tabulator` no front para ordenaÃ§Ã£o nativa, filtros por coluna, paginaÃ§Ã£o local e resize de colunas, com fallback automÃ¡tico para a grade HTML antiga se a biblioteca nÃ£o carregar.
- Na primeira carga da `Conferência`, o grid Tabulator já nasce com os dados do lote atual e limpa filtros/ordenação antigos antes de recarregar, evitando o estado em que os totais apareciam no topo mas a tabela ficava vazia.
- Quando a `Conferência` recebe atualizações do mesmo job em segundo plano, a grade preserva a página atual em vez de voltar para a página `1`; ela só reinicia a paginação quando o filtro principal muda.
- A grade Tabulator da `Conferência` agora mantém todas as colunas na mesma linha com scroll horizontal normal, sem criar linhas extras automáticas de detalhe como `Aba -` embaixo de cada registro.

## Prazos no painel

- A aba `Prazos` lê diretamente as planilhas e lista apenas títulos com `Status` vazio ou `A Receber`.
- Linhas marcadas como `BAIXADO`, `BAIXADA`, `ESTORNADO` ou `ESTORNADA` em qualquer coluna relevante da linha, inclusive quando o texto cair fora da coluna de `Status`, são ignoradas na aba `Prazos`.
- O painel separa boletos próximos do vencimento e depósitos atrasados, usando dias úteis e filtros customizáveis de `1` a `7`.
- Por padrão, boletos entram quando vencem em até `7` dias úteis e depósitos entram quando estão atrasados há pelo menos `7` dias úteis.
- Linhas amarelas representam boletos que ainda vão vencer; linhas vermelhas representam itens que vencem hoje ou já passaram da data.
- A aba atualiza ao abrir, mostra estado de carregamento fora da tabela e resume totais de boletos a vencer, boletos atrasados e depósitos atrasados.
- A aba `Prazos` agora também tem o filtro de origem por ícone (`MVA`, `EH` ou ambas), no mesmo padrão da `Conferência`.
- Os filtros de `Prazos` agora se dividem em uma caixa `Filtrar em dias:` com os campos compactos de `Boletos` e `Depósitos`, alinhada na mesma linha de `Origem` e do botão `Atualizar relação`.
- A grade principal de `Prazos` agora usa `Tabulator`, com paginação local, filtros por coluna e fallback para a tabela HTML antiga se a biblioteca não carregar.
- Ao recarregar `Prazos`, a relação anterior fica coberta por uma tela de loading enquanto a nova leitura das planilhas está em andamento.
- A grade de `Prazos` também preserva a página atual quando a mesma consulta é atualizada, evitando reset desnecessário da paginação durante recargas equivalentes.
- Os botões de atualização da relação e da busca por nome ficam em linhas separadas, com espaçamento próprio no card da aba `Prazos`.
- O botão `Buscar boletos em aberto` abre um popup próprio, fora do front-end nativo do navegador, para consultar por nome do cliente.
- O campo `Nome do cliente` do popup fica centralizado e usa uma lista própria de sugestões dentro do modal, em vez do `datalist` nativo do navegador.
- Se não houver pendências para o nome buscado, o popup mostra apenas a mensagem informando que não existem boletos em aberto.
- O autocomplete mistura os nomes salvos localmente em `%APPDATA%\\Botana\\watch_search_names.txt` com os nomes de clientes que ainda têm boletos em aberto nas planilhas, então buscas parciais como `LOCAR` passam a sugerir nomes maiores correspondentes.
- A busca por nome dentro de `Prazos` passa a respeitar a origem selecionada no painel, então `MVA`, `EH` e `Todas` usam o mesmo recorte tanto na grade quanto no popup.

## Ajustes de Prazos

- A coluna `Valor` prioriza `Valor da Parcela` e aceita células formatadas como moeda no Google Sheets; se vier vazio, cai para `Valor Total`.
- Os filtros da aba `Prazos` ficam empilhados, centralizados e com campos menores.
- Ao recarregar a aba `Prazos`, os totais são zerados antes da nova leitura para não manter números antigos na tela.

## Duplicidade estrutural

- O writer do Botana agora bloqueia novo lançamento quando a NF já tem todas as parcelas esperadas preenchidas, mesmo que a nova descrição venha como `DEP` em vez de `BLT`.
- A conferência também normaliza o número da parcela (`2ª`, `2º`, `Parcela 2`, `2/2`) para não deixar duplicidades estruturais passarem só por diferença de texto.
- Na aba `Conferência`, casos marcados como `Duplicada` passam a aparecer como divergência em vermelho.
