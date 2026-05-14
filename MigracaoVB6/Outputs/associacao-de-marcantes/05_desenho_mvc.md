# Template 05 - Desenho da Solucao MVC

## Aderencia ao Romaneio
- Padroes existentes identificados:
  - Controllers em `Romaneio/Controllers` com injecao por interface.
  - Repositorios com Dapper em `Romaneio/Repositorios`.
  - Interfaces em `Romaneio/Interfaces`.
  - Uso de `Session` para contexto do usuario/patio.
- Reusos previstos:
  - Modelo `Marcante` existente pode ser reaproveitado parcialmente.
  - Padrao de controller/repository visto em `MovimentacaoCSController` e `MovimentacaoCSRepositorio`.
- Pontos de extensao:
  - Novo repositorio especializado para associacao de marcantes (evita acoplar regra diferente em `MovimentacaoCS`).

## Componentes propostos
- Controller:
  - `AssociacaoMarcantesController` (GET Index + acoes Ajax para filtro/staging/salvar).
- ViewModel:
  - `AssociacaoMarcantesViewModel` (filtros, linha selecionada, lista de resultados, lista staging, flag EXP, mensagens).
- DTOs:
  - `AssociacaoMarcantesLinhaDto` (dados da grid + IDs tecnicos)
  - `AssociacaoMarcanteTempDto` (marcante/qtde)
  - `AssociacaoMarcantesFiltroDto` (lote/cntr)
- Interface de repositorio:
  - `IAssociacaoMarcantesRepositorio`
- Repositorio:
  - `AssociacaoMarcantesRepositorio` com Dapper parametrizado.
- View:
  - `Views/AssociacaoMarcantes/Index.cshtml`
- Scripts:
  - `Scripts/associacao-marcantes.js` (atalhos, selecao de linha, chamadas Ajax, validacoes de UX)

## Fluxo front-back
1. `GET /AssociacaoMarcantes/Index` carrega tela vazia e estrutura de filtro.
2. `POST /AssociacaoMarcantes/Filtrar` retorna grid e modo (`EXP` ou normal); selecao de linha dispara `CarregarTemporarios`.
3. `POST /AssociacaoMarcantes/AdicionarTemp`, `RemoverTemp` e `Salvar` executam regras no backend e retornam estado/mensagens para atualizar a tela.

## Rotas e endpoints
- GET:
  - `Index()`
- POST:
  - `Filtrar(filtro)`
  - `AdicionarTemp(payload)`
  - `RemoverTemp(payload)`
  - `Salvar(payload)`
- JSON/Ajax:
  - `ObterConteinersPorLote(lote)`
  - `CarregarTemporarios(autonumCsOuCexp, modoExp)`

## Persistencia e transacao
- Fontes de leitura:
  - `VW_WMS_SEM_MARCANTE`
  - `VW_WMS_SEM_MARCANTE_CEXP`
  - `TB_MARCANTES`
  - `TB_CARGA_SOLTA`
  - `TB_CARGA_EXPORTACAO`
  - `TB_TEMP_MARCANTE`
  - `TB_TEMP_MARCANTE_CEXP`
- Operacoes de escrita:
  - inserir/remover na temp (staging)
  - update final em `TB_MARCANTES`
- Estrategia transacional:
  - recomendada transacao no `Salvar` para o loop de updates por marcante.
  - recomendada validacao atomica antes de gravar, para evitar estado parcial.

## Adaptacao de comportamento VB6 para web
- Atalhos:
  - mapear F1/F2/F4/F5/ESC no JS.
  - tratar F5 para evitar refresh indevido apenas no contexto da tela.
- Lista temporaria:
  - manter staging no banco (`TB_TEMP_*`) para paridade inicial com legado.
  - alternativa futura (hipotese de evolucao): staging em sessao para reduzir dependencia de temp persistida.
- Mensagens:
  - manter textos funcionais do legado para homologacao por equivalencia.
- Selecao de linha:
  - guardar linha selecionada no estado da pagina (id tecnico oculto) e revalidar no backend.

## Inventario de aderencia ao padrao atual do Romaneio
- Aderente:
  - Controller + Repositorio + Interface + ViewModel + View Razor.
  - Dapper para acesso ao banco.
- Ajuste necessario:
  - parametrizar SQL (repo atual possui concatenacoes em alguns pontos; para esta tela recomenda-se elevar padrao de seguranca).
  - registrar interface no UnityConfig caso adocao de mapeamento explicito seja aplicada.
- Nao recomendado:
  - criar camada paralela fora da estrutura `Controllers/Interfaces/Repositorios/Models/Views`.

## Decisoes tecnicas
- Decisao: criar modulo MVC dedicado para Associacao de Marcantes.
- Motivo: regra especifica e staging proprio; evita misturar com `MovimentacaoCS`.
- Impacto: incremento controlado, com baixo risco de regressao cruzada.

- Decisao: preservar regra de fallback para CEXP.
- Motivo: comportamento de negocio explicito no legado.
- Impacto: maior fidelidade funcional na migracao.

- Decisao: manter temp tables na fase inicial.
- Motivo: menor diferenca comportamental com VB6.
- Impacto: exige validar limpeza e concorrencia.

## Criterios de aceite da futura migracao (base para implementacao)
- Filtro por lote/cntr retorna mesmo conjunto funcional do legado.
- Modo EXP e regras de compatibilidade de marcante funcionam igual ao VB6.
- Inclusao/remocao em staging reproduz lista temporaria esperada.
- Salvar respeita consistencia de quantidades e atualiza `TB_MARCANTES` corretamente.
- Atalhos F1/F2/F4/F5/ESC funcionam sem degradar usabilidade web.
- Mensagens criticas equivalentes ao legado sao apresentadas ao usuario.
