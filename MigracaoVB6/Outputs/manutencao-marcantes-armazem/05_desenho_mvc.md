# Desenho da solucao MVC - Manutencao Marcantes Armazem (legado AdminMarcante)

## Aderencia ao projeto Romaneio (padroes observados)

- Controllers em `Romaneio/Controllers` herdando `DefaultController` (quando aplicavel), com checagem de `Session["Logado"]` em acoes — **referencia**: `AssociacaoMarcantesController`.
- Injecao via **Unity** em `Global.asax.cs` (`RegisterType<I..., ...Repositorio>`).
- Repositorios em `Romaneio/Repositorios` usando **Dapper** + `Config.StringConexao()` + `Config.QueryTimeoutInSeconds()` — **referencia**: `AssociacaoMarcantesRepositorio`.
- Interfaces em `Romaneio/Interfaces/Repositories`.
- ViewModels e DTOs em `Romaneio/Models/ViewModels`.
- Views Razor em `Romaneio/Views/<Controller>/Index.cshtml`.
- Scripts em `Romaneio/Content/js/*.js` — **referencia**: `Content/js/associacao-marcantes.js`.

**Decisao**: criar modulo dedicado espelhando o isolamento de `AssociacaoMarcantes` (controller + interface + repositorio + view + js), sem misturar em `AssociacaoMarcantes` nem em `CarregamentoCargaSolta`, pois o escopo e **update pontual de volume** pos-divergencia.

## Componentes propostos (sem implementar)

| Camada | Nome sugerido |
|--------|---------------|
| Controller | `ManutencaoMarcantesArmazemController` (ou `ManutencaoMarcantesController` se preferir nome mais curto — alinhar com menu Romaneio) |
| Interface | `IManutencaoMarcantesArmazemRepositorio` |
| Repositorio | `ManutencaoMarcantesArmazemRepositorio` |
| ViewModel | `ManutencaoMarcantesArmazemViewModel` (filtro, cabecalho divergencia, linhas, totais, mensagens) |
| DTOs | `ManutencaoMarcantesArmazemDivergenciaDto`, `ManutencaoMarcantesArmazemLinhaDto` |
| View | `Views/ManutencaoMarcantesArmazem/Index.cshtml` |
| Script | `Content/js/manutencao-marcantes-armazem.js` |

> Nome da rota/pasta: usar slug alinhado ao menu Romaneio e ao slug desta migracao: `manutencao-marcantes-armazem`.

## ViewModel (campos principais)

- `int? Lote` — mapeia `TXTLOTE`.
- `string Embalagem`, `int? QuantidadeReferencia` — mapeia `txtEmbalagem` / `txtQtde`.
- `int? AutonumCs` — mapeia `txtAutonumCS` (interno).
- `List<ManutencaoMarcantesArmazemLinhaDto> Linhas` — grid.
- `int? MarcanteSelecionado`, `int? NovaQuantidade` — edicao.
- `int? TotalVolumesListados` — derivado no servidor ao filtrar/salvar (paridade com `TXTTOT`).
- Flags de UI: `bool TotalConfereReferencia` (para CSS verde/vermelho).

## Endpoints e rotas (proposta)

| Metodo | Acao | Contrato (alto nivel) |
|--------|------|------------------------|
| GET | `Index()` | Renderiza tela vazia ou estado inicial. |
| POST | `Filtrar(int lote)` | Retorna divergencia + linhas + totais; mensagem se nao encontrado. |
| POST | `Salvar(int marcanteAutonum, int novaQuantidade, int autonumCs)` | Executa update parametrizado; retorna lista atualizada + totais + sucesso/erro. |

**JSON**: mesmo padrao `{ success, message, ... }` usado em `AssociacaoMarcantesController`.

## Fluxo front-back

1. Pagina carrega (`GET Index`).
2. Usuario informa lote e aciona Filtrar -> `POST Filtrar` -> backend consulta `VW_WMS_DIVERGENCIA_QTDE` e, se ok, monta lista com o SQL equivalente a `Carrega_Lv`.
3. Usuario seleciona linha na tabela -> front armazena `marcanteAutonum` (e exibe formatado).
4. Usuario informa nova quantidade -> `POST Salvar` com validacao server-side espelhando `cmdSalvar_Click` + **melhorias recomendadas** (ver proxima secao).
5. Resposta atualiza grid e estado do total/cores.

## Persistencia e SQL (parametrizado na migracao)

- Leitura divergencia: `SELECT TOP 1 embalagem, QUANTIDADE_REAL, AUTONUM_CS FROM SGIPA..VW_WMS_DIVERGENCIA_QTDE WHERE LOTE = @Lote` — **Hipotese**: pode ser necessario `TOP 1` se a view retornar multiplos; o VB6 usa `SELECT *` e nao trata multiplas linhas explicitamente.
- Lista: reproduzir SQL do `Carrega_Lv` com parametros `@AutonumCs`.
- Update: `UPDATE SGIPA..TB_MARCANTES SET VOLUMES = @Volumes WHERE AUTONUM = @Autonum` — **obrigatorio** validar no servidor que o marcante pertence ao `AUTONUM_CS` filtrado (o VB6 **nao** revalida na gravacao — **lacuna de seguranca** a corrigir na web).

## Estrategia de consulta / edicao / remocao

- **Consulta**: somente leitura das views/tabelas citadas.
- **Edicao**: unico campo gravavel `VOLUMES`.
- **Remocao**: fora de escopo (nao existe no legado).

## Adaptacao VB6 -> Web

- **Revalidacao de pertenca** do marcante ao `AUTONUM_CS` antes do `UPDATE` — **Recomendacao de engenharia** (reforca negocio ausente no legado).
- **Transacao**: opcional (uma unica tabela); manter simples, mas envolver em try/catch e retorno de erro ao usuario.
- **Atalhos**: F4/F5/ESC em JS; tratar `F5` com `preventDefault` quando foco na pagina (padrao do projeto em outras telas).
- **Feedback**: adicionar mensagem de sucesso — **melhoria de UX** em relacao ao VB6 silencioso.

## Inventario de aderencia

- **Aderente**: Controller + Repositorio + Interface + ViewModel + Razor + `Content/js`.
- **Ajuste**: registrar `IManutencaoMarcantesArmazemRepositorio` no `Global.asax.cs` e incluir `.csproj` (quando implementar).
- **Nao recomendado**: logica SQL direta no controller.

## Decisoes tecnicas (proposta)

1. **Modulo dedicado** — baixo acoplamento, espelha complexidade moderada e risco de alteracao de `TB_MARCANTES`.
2. **Validacao server-side reforcada** — mitigar risco de manipulacao de payload.
3. **Paridade funcional primeiro** — textos e fluxo iguais; UX de sucesso explicito como extensao acordada.
