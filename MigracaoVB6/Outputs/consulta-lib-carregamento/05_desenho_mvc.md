# Desenho da Solucao MVC - Consulta Lib. Carregamento

## Aderencia ao Romaneio
- Padroes existentes identificados:
  - Controllers herdando `DefaultController`.
  - Injeccao de interfaces de repositorio no construtor.
  - Responses JSON para Ajax com `success/message` e uso de `TempData` em fluxos post-back.
  - View Razor em `Views/<Modulo>/Index.cshtml` e JS dedicado em `Content/js`.
  - Acesso a dados via Dapper em classes de `Repositorios`.
- Reusos previstos:
  - Padrao de sessao (`Session["Logado"]`, `Session["Patio"]`).
  - Estrutura de mensagens e codigos de `DefaultController`.
  - Estrategia de endpoints Ajax como em `AssociacaoMarcantes`.
- Pontos de extensao:
  - Novo modulo para consulta/liberacao com repositorio especializado.
  - Consolidacao de chamadas de dominio em servico para regras extensas.

## Componentes propostos
- Controller:
  - `ConsultaLiberacaoCarregamentoController : DefaultController`.
- ViewModel:
  - `ConsultaLiberacaoCarregamentoViewModel` (filtros + dados da carga + status + listas + flags de bloqueio).
- DTOs:
  - `ConsultaLiberacaoFiltroRequest`.
  - `ConsultaLiberacaoResultadoDto`.
  - `CargaListaItemDto`.
  - `ProtocoloSimultaneoDto`.
  - `RegistrarLiberacaoRequest`.
  - `RegistrarLiberacaoResultDto`.
- Interface de repositorio:
  - `IConsultaLiberacaoCarregamentoRepositorio`.
- Repositorio:
  - `ConsultaLiberacaoCarregamentoRepositorio` (Dapper + queries parametrizadas).
- View:
  - `Views/ConsultaLiberacaoCarregamento/Index.cshtml`.
- Scripts:
  - `Content/js/consulta-liberacao-carregamento.js`.

## Fluxo front-back
1. `GET Index` carrega tela, combos basicos (tipo veiculo) e estado inicial.
2. `POST Filtrar` recebe protocolo/placa/cntr/bl, aplica validacoes, consulta dados e retorna JSON com payload completo para render.
3. `POST Registrar` executa pipeline de validacoes e registro; retorna resultado detalhado de sucesso/erro.
4. `POST Limpar` pode ser local no front (sem roundtrip) mantendo padrao web.
5. `GET/POST Documentos` (opcional) define tratamento equivalente ao executavel externo (link/integração futura).

## Rotas e endpoints
- GET:
  - `/ConsultaLiberacaoCarregamento/Index`
- POST:
  - `/ConsultaLiberacaoCarregamento/Filtrar`
  - `/ConsultaLiberacaoCarregamento/Registrar`
  - `/ConsultaLiberacaoCarregamento/CarregarSimultaneos`
- JSON/Ajax:
  - retorno estruturado para bind de dados, status e bandeiras de bloqueio (`bloqueioBL`, `bloqueioCntr`, `glmePendente`, `podeRegistrar`).

## Persistencia e transacao
- Fontes de leitura:
  - `VW_AGENDA_CS`, `TB_BL`, `TB_PATIOS`, `TB_CNTR_BL`, `TB_AMR_CNTR_BL`, `TB_GD_RESERVA`, `TB_REGISTRO_SAIDA_CS`, `TB_REGISTRO_SAIDA_CNTR`, `vw_invent_armazem_item`.
- Operacoes de escrita:
  - update de tipo de veiculo (`TB_AG_VEICULOS`) e updates de liberacao/ordem conforme regras.
- Estrategia transacional:
  - encapsular em transacao no repositorio/servico para operacoes multiupdate de registro.
  - chamadas externas (SISCARGA/SEFAZ) com tratamento de erro e rollback logico quando aplicavel.

## Adaptacao de comportamento VB6 para web
- Atalhos:
  - manter `ESC` para sair/voltar; mapear atalhos adicionais apenas se validados com usuarios.
- Lista temporaria:
  - nao aplicavel como staging local; usar retorno de consultas para listas (`listaCarga` e `simultaneos`).
- Mensagens:
  - substituir cascata de `MsgBox` por alertas web/modais padronizados e texto preservado.
- Selecao de linha:
  - tabelas/listas selecionaveis com destaque e estado em JS.

## Decisoes tecnicas
- Decisao:
  - criar modulo dedicado (controller/repositorio/view/js).
- Motivo:
  - fluxo de liberacao tem regras densas e sensiveis, nao encaixa como extensao pequena de tela existente.
- Impacto:
  - maior isolamento, melhor testabilidade e menor risco de regressao cruzada.

- Decisao:
  - usar SQL parametrizado no Dapper em vez de concatenacao VB6.
- Motivo:
  - seguranca, manutenibilidade e previsibilidade.
- Impacto:
  - reducao de risco tecnico sem mudar regra de negocio.

- Decisao:
  - separar validacoes em camadas (interface, dominio e tecnica), mantendo classificacao rastreavel.
- Motivo:
  - facilitar homologacao com usuario final e troubleshooting.
- Impacto:
  - documentacao e implementacao mais claras por tipo de regra.
