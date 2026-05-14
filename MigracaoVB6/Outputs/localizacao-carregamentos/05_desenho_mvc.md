# Desenho da Solucao MVC - Localizacao Carregamentos

## Aderencia ao Romaneio
- Padroes existentes identificados:
  - Controller herdando `DefaultController`.
  - Sessao via `Session["Logado"]`, `Session["Patio"]`, `Session["UsuarioId"]`.
  - Repositorio com Dapper + `Config.StringConexao()`.
  - View Razor com tabela HTML + JS dedicado em `Content/js`.
- Reusos previstos:
  - Estrutura semelhante a `CarregamentoCargaSoltaController` e repositorio dedicado.
  - Padrao de Json GET/POST com retorno `{ success, message, ... }`.
- Pontos de extensao:
  - Endpoint de consulta de localizacao.
  - Endpoint opcional de movimentacao para `SAIDA` (se escopo funcional confirmar necessidade).

## Componentes propostos
- Controller:
  - `LocalizacaoCarregamentosController`
  - Acoes: `Index`, `Filtrar`, `MoverSaida` (condicional), `Detalhe`.
- ViewModel:
  - `LocalizacaoCarregamentosViewModel` (patio, filtro placa, totais, lista inicial opcional).
- DTOs:
  - `LocalizacaoCarregamentoItemDto` (dados da grid + flag tipo).
  - `LocalizacaoDetalheDto` (campos do painel superior).
  - `MoverSaidaRequest` e `MoverSaidaResult`.
- Interface de repositorio:
  - `ILocalizacaoCarregamentosRepositorio`.
- Repositorio:
  - `LocalizacaoCarregamentosRepositorio` com SQL parametrizado para:
    - consulta consolidada (conteiner + carga solta),
    - validacoes de saida,
    - persistencia de saida/historico.
- View:
  - `Views/LocalizacaoCarregamentos/Index.cshtml`.
- Scripts:
  - `Content/js/localizacao-carregamentos.js`.

## Fluxo front-back
1. `GET Index`: valida sessao, injeta patio e carrega view.
2. `GET Filtrar`: recebe `placa`, consulta dados e retorna lista + totais.
3. Front seleciona linha e mostra detalhe local.
4. (Opcional) `POST MoverSaida`: valida regra de negocio e persiste update+historico em transacao.

## Rotas e endpoints
- GET:
  - `/LocalizacaoCarregamentos/Index`
  - `/LocalizacaoCarregamentos/Filtrar?placa=...`
- POST:
  - `/LocalizacaoCarregamentos/MoverSaida`
- JSON/Ajax:
  - Respostas padronizadas com `success`, `message`, `data`, `totais`.

## Persistencia e transacao
- Fontes de leitura:
  - mesmas tabelas/views mapeadas no VB6.
- Operacoes de escrita:
  - somente para fluxo de `SAIDA` (se mantido no escopo).
- Estrategia transacional:
  - usar `BeginTransaction(ReadCommitted)` para `UPDATE TB_CNTR_BL` + `INSERT TB_HIST_SHIFTING`.
  - adicionar validacao de pre-condicao dentro da mesma transacao quando viavel.

## Adaptacao de comportamento VB6 para web
- Atalhos:
  - `ESC` para voltar ao Home.
  - `Enter` no filtro para filtrar.
- Lista temporaria:
  - nao ha staging nesta tela; apenas selecao de resultado.
- Mensagens:
  - substituir forms modais VB6 por alertas JS/sweetalert.
- Selecao de linha:
  - highlight de linha + bind para painel de detalhe.

## Decisoes tecnicas
- Decisao:
  - Implementar consulta por SQL parametrizado em repositorio dedicado.
- Motivo:
  - Evitar concatenacao do VB6 e manter regra proxima da origem.
- Impacto:
  - Maior seguranca e manutencao.

- Decisao:
  - Tratar diferenca de variantes (`Localizacao` vs `LocalizacaoS`) por flag/configuracao de comportamento.
- Motivo:
  - Preservar compatibilidade sem criar telas duplicadas.
- Impacto:
  - Exige validacao funcional para escolher query base e janela temporal oficial.
