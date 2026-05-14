# Desenho da Solucao MVC - Carregamento Carga Solta (Armazem)

## Aderencia ao Romaneio
- Padroes existentes identificados:
  - Controllers herdando de `DefaultController` (que herda `BaseController`).
  - Injecao de dependencia via Unity em `Global.asax.cs` (`container.RegisterType<IXxxRepositorio, XxxRepositorio>`).
  - Acesso a dados via Dapper sobre `SqlConnection` usando `Helpers.Config.StringConexao()` e `Config.QueryTimeoutInSeconds()`.
  - ViewModels sob `Models/ViewModels/` e classes simples sob `Models/`.
  - View Razor por modulo em `Views/<Modulo>/Index.cshtml`.
  - JS por tela em `Content/js/<nome-tela>.js` (vide `consulta-lib-carregamento.js`).
  - Mensagens via `TempData["MensagemSucesso"]`/`TempData["MensagemErro"]` para fluxos post-back e via `JsonResult` `{ success, message, ... }` para Ajax.
  - Sessao com `Session["Logado"]` e `Session["Patio"]` (visto em `MovimentacaoCSController` e `AssociacaoMarcantesController`).
  - Estilo de layout `Bootstrap` com `card-custom`, `form-control`, `btn`, `glyphicon` (vide `MovimentacaoCS/Index.cshtml`).
- Reusos previstos:
  - `IDadosRepositorio` para listas auxiliares (armazens/embalagens) - apenas se houver demanda.
  - Padrao de deteccao de bipagem (delta de tempo) pode ser extraido como utilitario JS reusavel.
  - Estrategia transacional ja vista em `AssociacaoMarcantesRepositorio.Salvar` (`con.BeginTransaction(IsolationLevel.ReadCommitted)`) - replicar.
  - Funcao escalar `FC_VALIDA_SAIDA_CARGA` permanece no banco; o repositorio apenas a invoca via `SELECT dbo.FC_VALIDA_SAIDA_CARGA(...) AS RETORNO`.
- Pontos de extensao:
  - Modulo novo `CarregamentoCargaSolta` (controller/repositorio/view/js) sem mexer em modulos existentes.
  - Compatibilidade futura com a tela de "Carregamento Carga Solta - Patio" (variante `CarregaCSPatio.frm`) reusando partes do repositorio - registrar como evolucao posterior.

## Componentes propostos
- Controller:
  - `CarregamentoCargaSoltaController : DefaultController` em `Romaneio/Controllers/CarregamentoCargaSoltaController.cs`.
- ViewModel:
  - `CarregamentoCargaSoltaViewModel` em `Models/ViewModels/CarregamentoCargaSoltaViewModel.cs`.
    - `int? PATIO` (vem de `Session["Patio"]`), `bool BIPAGEM_OBRIGATORIA` (mapeia `flag_carregamento_marcante_bip`), `bool BIPAGEM_BLOQUEAR_APOS_VALIDAR` (`= 1`).
    - `List<VeiculoDto> VEICULOS`.
- DTOs:
  - `VeiculoDto { string PLACA_C; string DISPLAY; }`.
  - `OrdemCarregamentoDto { int AUTONUMCS; int LOTE; int ITEM; int QUANTIDADE; int QTDE_CARREGADA; int NUM_OC; int ORDEM_CARREG; string EMBALAGEM; string PLACA_C; string PLACA_CARRETA; string MODELO; }`.
  - `MarcanteCarregadoDto { int MARCANTE; int QTDE; }`.
  - `LocalArmazemDto { int MARCANTE; int QTDE; string DESCR_ARMAZEM; string POSICAO; }`.
  - `MarcanteDadosDto { int MARCANTE; int AUTONUMCS; int LOTE; int ITEM; int VOLUMES; string EMBALAGEM; string MERCADORIA; string MARCA; int AUTONUM_ARMAZEM; string DESCR_ARMAZEM; int? AUTONUMCNTR; string ID_CONTEINER; string POSICAO_YARD; int AUTONUM_CS_YARD; }`.
  - `CarregarItemRequest { int AUTONUMCS_OC; int ORDEM; string PLACA_C; int MARCANTE; int QTDE; int AUTONUM_CS_YARD; int AUTONUM_ARMAZEM; }`.
  - `CarregarItemResult { bool SUCESSO; bool PARCIAL; string MENSAGEM; }`.
  - `CancelarItemRequest { int MARCANTE; int AUTONUMCS_OC; int AUTONUM_CS_YARD; int AUTONUM_ARMAZEM; int QTDE; }`.
  - `CancelarItemResult { bool SUCESSO; string MENSAGEM; }`.
  - `DocumentoOcRequest { int LOTE_BL; }` -> `DocumentoOcResult { bool LTL; int? AUTONUM_AG_CS; int? LOTE_BL; string URL_DOCUMENTO; }` (URL composta do lado server para uma rota equivalente ao `ConsultaDocColetor.exe`).
- Interface de repositorio:
  - `ICarregamentoCargaSoltaRepositorio` em `Interfaces/Repositories/ICarregamentoCargaSoltaRepositorio.cs`.
- Repositorio:
  - `CarregamentoCargaSoltaRepositorio : ICarregamentoCargaSoltaRepositorio` em `Repositorios/CarregamentoCargaSoltaRepositorio.cs`.
- View:
  - `Views/CarregamentoCargaSolta/Index.cshtml` (uma unica tela; modal bootstrap embutida para "Locais").
- Scripts:
  - `Content/js/carregamento-carga-solta.js` (atalhos, ajax, anti-digitacao).

## Fluxo front-back
1. **GET `Index`**:
   - Verifica `Session["Logado"]`; se nao, redireciona para Home.
   - Le `Session["Patio"]` e `Session["FlagCarregamentoMarcanteBip"]` (precisara ser populada no login a partir de `TB_CAD_USUARIOS.flag_carregamento_marcante_bip`).
   - Carrega `VEICULOS` via repositorio (consulta unica).
   - Renderiza `Index.cshtml`.
2. **GET `ListarOrdens?placa=ABC1D23`**: retorna lista de OCs (`OrdemCarregamentoDto[]`) para a placa selecionada.
3. **GET `ListarCarregamentos?placa=ABC1D23&autonumCs=<id>`**: retorna marcantes ja em `'CAM'` para aquela OC e placa.
4. **GET `ListarLocais?autonumCs=<id>`**: retorna locais do armazem para a OC.
5. **GET `BuscarMarcante?marcante=000000123`**: retorna `MarcanteDadosDto`.
6. **POST `Carregar`** (JSON `CarregarItemRequest`): aplica todas as validacoes, chama `FC_VALIDA_SAIDA_CARGA`, executa transacao e retorna `CarregarItemResult`.
7. **POST `Cancelar`** (JSON `CancelarItemRequest`): valida, executa transacao e retorna `CancelarItemResult`.
8. **GET `Documento?lote=<bl>`**: faz o roteamento LTL/AG_CS e retorna `DocumentoOcResult` (com URL de redirecionamento ou base64 caso haja servico web disponivel) - decisao de UX: ver "Decisoes".
9. Pos `Carregar`/`Cancelar` o JS:
   - Atualiza tabela `Ordens` (re-fetch `ListarOrdens` mantendo selecao por `AUTONUMCS`).
   - Atualiza tabela `Carregamentos` (re-fetch `ListarCarregamentos`).
   - Limpa campos do marcante e devolve foco a `txtMarcante` (replicando comportamento VB6).
   - Em caso de `BIPAGEM_BLOQUEAR_APOS_VALIDAR`, desabilita `txtMarcante` apos validar e habilita novamente apos `Limpar`.

## Rotas e endpoints
- GET:
  - `/CarregamentoCargaSolta/Index`
  - `/CarregamentoCargaSolta/ListarVeiculos` (re-fetch para botao "Atualizar")
  - `/CarregamentoCargaSolta/ListarOrdens?placa=...`
  - `/CarregamentoCargaSolta/ListarCarregamentos?placa=...&autonumCs=...`
  - `/CarregamentoCargaSolta/ListarLocais?autonumCs=...`
  - `/CarregamentoCargaSolta/BuscarMarcante?marcante=...`
  - `/CarregamentoCargaSolta/Documento?lote=...`
- POST:
  - `/CarregamentoCargaSolta/Carregar`
  - `/CarregamentoCargaSolta/Cancelar`
- JSON/Ajax:
  - Padrao: `{ success: bool, message: string, dados: object }`.
  - Em endpoints de lista, retorna `{ success: true, items: [...] }`.

## Persistencia e transacao
- Fontes de leitura:
  - `SGIPA..VW_CAM_CARREGAMENTO` (combo de veiculos e ordens).
  - `SGIPA..VW_INVENT_ARMAZEM` (locais).
  - `SGIPA..TB_CARGA_SOLTA`, `SGIPA..TB_MARCANTES`, `SGIPA..TB_CARGA_SOLTA_YARD`, `SGIPA..TB_BL`, `SGIPA..TB_SOLICITACAO_LTL`, `SGIPA..TB_AG_CS`, `SGIPA..TB_ARMAZENS_IPA`, `SGIPA..TB_CNTR_BL`, `SGIPA..DTE_TB_EMBALAGENS`.
  - Funcao escalar `dbo.FC_VALIDA_SAIDA_CARGA(BL, ORDEM, AUTONUMCS)`.
- Operacoes de escrita:
  - `INSERT INTO SGIPA..TB_CARGA_SOLTA_YARD (AUTONUM_CS, ARMAZEM, YARD, ORIGEM, QUANTIDADE, MOTIVO)` (`'CAM'` para `Carregar` e `'CANCC'` para `Cancelar`, sempre `MOTIVO=8` e `ORIGEM='I'`).
  - `UPDATE SGIPA..TB_CARGA_SOLTA_YARD SET QUANTIDADE=0` (carga total) ou `SET QUANTIDADE=QUANTIDADE-@QTDE` (parcial).
  - `UPDATE SGIPA..TB_MARCANTES SET ORDEM=@ORDEM, PLACA_C=@PLACA_C` (associacao - manter dentro da transacao na MVC).
  - `UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CS_YARD=@NOVO_ID` (e opcionalmente `VOLUMES=@QTDE` em parcial).
  - `UPDATE SGIPA..tb_temp_marcante SET QTDE=@QTDE WHERE MARCANTE=@MARCANTE_TXT` (parcial).
  - `INSERT INTO SGIPA..TB_HIST_SHIFTING_CS (MARCANTE, ARMAZEM, YARD, DT_MOV, USUARIO)`.
  - Cancelar: `UPDATE TB_MARCANTES SET PLACA_C=NULL, AUTONUM_CS_YARD=@NOVO_ID`.
- Estrategia transacional:
  - Todas as operacoes de `Carregar` em uma unica `BeginTransaction(IsolationLevel.ReadCommitted)` no repositorio (replicar padrao de `AssociacaoMarcantesRepositorio.Salvar`).
  - Recuperacao do `AUTONUM` do novo registro `TB_CARGA_SOLTA_YARD` via clausula `OUTPUT INSERTED.AUTONUM` (corrige risco de concorrencia do legado).
  - Em caso de erro, rollback e retorno `{ success: false, message: "..." }` ao front.
  - Validacao com `FC_VALIDA_SAIDA_CARGA` ocorre antes do `BEGIN TRAN` para evitar lock desnecessario - se nao retorna `'OK'`, aborta e devolve a critica textual.

## Adaptacao de comportamento VB6 para web
- Atalhos:
  - JS captura `keydown` no `document` (com `e.preventDefault()` em `F1..F5` quando o foco esta na area da tela): `F1` abre modal Locais, `F2` aciona endpoint Documento, `F3` limpa, `F4` carrega, `F5` cancela, `Esc` fecha modal ou navega para Home (definir).
  - Garantir que atalhos so disparem quando a tela esta visivel e nenhum outro modal ativo.
- Lista temporaria:
  - Nao se aplica como staging local (a tela nao tem lista temporaria; a "lista" e a tabela `lvCm`, que reflete o estado real do banco).
- Mensagens:
  - Substituir `txtCritica` por banner Bootstrap (`#critica`) com classe `alert-success` ou `alert-danger`. Cores e auto-dismiss opcional.
  - `MsgBox` "Campo nao permite digitacao manual" -> alerta inline (toastr/alert) com mesmo texto.
- Selecao de linha:
  - `lvOC` -> `<table>` com classe `table-hover`. Clicar em linha aplica `selected` e dispara `ListarCarregamentos`. Navegacao por teclado (`ArrowUp`/`ArrowDown`) via JS.
- Anti-digitacao manual:
  - Replicar logica do VB6 em JS:
    - Manter `lastKeyDownAt` e `lastInputAt`.
    - Em cada `keydown`, registrar timestamp.
    - Em `Enter`, se `lastInputAt - inicioCaptura > 1500ms` ou nao houve sequencia rapida (heuristica), e `BIPAGEM_OBRIGATORIA = true`, exibir alerta e limpar.
  - Quando `BIPAGEM_BLOQUEAR_APOS_VALIDAR = true`, apos resposta de `BuscarMarcante` setar `disabled` no input.

## Decisoes tecnicas
- Decisao: criar modulo dedicado `CarregamentoCargaSolta` (controller/repositorio/view/js).
  - Motivo: tela tem regras especificas (transacoes, `FC_VALIDA_SAIDA_CARGA`, anti-digitacao) que nao se encaixam em modulos existentes.
  - Impacto: maior isolamento e testabilidade; reduz risco de regressao em telas adjacentes (`MovimentacaoCS`, `AssociacaoMarcantes`).

- Decisao: usar Dapper com SQL parametrizado (`@VARIAVEL`) em vez de concatenacao.
  - Motivo: seguranca (evita injection) e clareza.
  - Impacto: nenhum em comportamento; melhora robustez.

- Decisao: trazer o `UPDATE TB_MARCANTES SET ORDEM, PLACA_C` para dentro da transacao do Carregar.
  - Motivo: o legado fazia fora; isso permite estado inconsistente em caso de falha intermediaria.
  - Impacto: comportamento equivalente em sucesso, porem mais consistente em falhas. Risco baixo.

- Decisao: usar `OUTPUT INSERTED.AUTONUM` no INSERT em `TB_CARGA_SOLTA_YARD` ao inves de `SELECT MAX(AUTONUM)` apos.
  - Motivo: corrige risco de concorrencia (dois operadores na mesma OC).
  - Impacto: alinha com boas praticas SQL Server; nao altera funcionalmente o legado.

- Decisao: gravar historico tambem no `Cancelar` (linha em `TB_HIST_SHIFTING_CS` com `YARD='CANCC'`).
  - Motivo: rastreabilidade. O legado nao grava (parece omissao).
  - Impacto: pequena alteracao de comportamento, alinhada com auditoria.
  - Status: PROPOSTA - confirmar com time de operacao antes de implementar.

- Decisao: substituir `ConsultaDocColetor.exe` por uma rota web equivalente.
  - Motivo: navegador nao executa exe; impacto operacional.
  - Impacto/abordagens:
    - Curto prazo: endpoint `Documento` retorna URL/PDF baseado no servico `Bandeirantes.AnexoDiretorio.Service.dll` (ja referenciado pelo exe).
    - Medio prazo: integrar com tela existente do Romaneio que liste anexos do BL/AG_CS, se houver.
  - Status: pendente de definicao; tela operacional (carregar/cancelar) nao depende disso para funcionar.

- Decisao: persistir `flag_carregamento_marcante_bip` em `Session` no login do Romaneio.
  - Motivo: evita query repetida e replica comportamento VB6 (carregado uma vez no login).
  - Impacto: ajuste leve no `LoginController` (ler `TB_CAD_USUARIOS.flag_carregamento_marcante_bip` e gravar em `Session["FlagCarregamentoMarcanteBip"]`).

- Decisao: filtro de patio reaproveita `Session["Patio"]`.
  - Motivo: padrao do Romaneio (`MovimentacaoCSController` ja faz `Session["Patio"].ToString()`).
  - Impacto: alinhado.

- Decisao: nao replicar o `Adodc1` (datacontrol OLEDB) - usar Dapper como nas demais telas.
  - Motivo: simplificacao; o combo e estatico apos load + botao Refresh.
  - Impacto: 1 query a menos no canal duplo; consistente.

- Decisao: nao implementar `KeyPress=Enter -> SendKeys("{TAB}")`.
  - Motivo: HTML nativamente avanca foco com `Enter` em formularios; dispensa simulacao.
  - Impacto: comportamento equivalente sem hack.

## Estrutura de arquivos sugerida (resumo)
```
Romaneio/Romaneio/Romaneio/
  Controllers/CarregamentoCargaSoltaController.cs        [novo]
  Repositorios/CarregamentoCargaSoltaRepositorio.cs      [novo]
  Interfaces/Repositories/ICarregamentoCargaSoltaRepositorio.cs [novo]
  Models/ViewModels/CarregamentoCargaSoltaViewModel.cs   [novo]
  Views/CarregamentoCargaSolta/Index.cshtml              [novo]
  Content/js/carregamento-carga-solta.js                 [novo]
  Global.asax.cs                                         [editar: registrar interface no Unity]
  Romaneio.csproj                                        [editar: incluir os novos arquivos]
```

## Endpoints (contratos de payload)
- `GET ListarOrdens(placa)` -> `{ success: true, items: OrdemCarregamentoDto[] }`.
- `GET ListarCarregamentos(placa, autonumCs)` -> `{ success: true, items: MarcanteCarregadoDto[] }`.
- `GET ListarLocais(autonumCs)` -> `{ success: true, items: LocalArmazemDto[] }`.
- `GET BuscarMarcante(marcante)` -> `{ success: bool, dados: MarcanteDadosDto, message }` (404 se nao encontrado).
- `POST Carregar(CarregarItemRequest)` -> `{ success: bool, parcial: bool, message }`.
- `POST Cancelar(CancelarItemRequest)` -> `{ success: bool, message }`.
- `GET Documento(loteBl)` -> `{ success, ltl: bool, autonumAgCs, urlDocumento }` (ou redirect direto).

## Esqueleto de implementacao (referencia)

### Controller (esqueleto, nao executavel)
```csharp
public class CarregamentoCargaSoltaController : DefaultController
{
    private readonly ICarregamentoCargaSoltaRepositorio _repo;

    public CarregamentoCargaSoltaController(ICarregamentoCargaSoltaRepositorio repo) => _repo = repo;

    [HttpGet]
    public ActionResult Index()
    {
        if (Session["Logado"] == null) return RedirectToAction("Index", "Home");

        var view = new CarregamentoCargaSoltaViewModel
        {
            PATIO = int.TryParse(Session["Patio"]?.ToString(), out var p) ? (int?)p : null,
            BIPAGEM_OBRIGATORIA = (Session["FlagCarregamentoMarcanteBip"]?.ToString() ?? "0") != "0",
            BIPAGEM_BLOQUEAR_APOS_VALIDAR = (Session["FlagCarregamentoMarcanteBip"]?.ToString() ?? "0") == "1",
            VEICULOS = _repo.ListarVeiculos(/* patio */).ToList()
        };
        return View(view);
    }

    [HttpPost]
    public JsonResult Carregar(CarregarItemRequest request)
    {
        if (Session["Logado"] == null) return Json(new { success = false, message = "Sessao caiu" });
        var codUsuario = Convert.ToInt32(Session["IdUsuario"]);
        var result = _repo.Carregar(request, codUsuario);
        return Json(new { success = result.SUCESSO, parcial = result.PARCIAL, message = result.MENSAGEM });
    }
    // ... demais acoes seguem mesmo padrao
}
```

### Repositorio (transacao do Carregar - referencia)
```csharp
public CarregarItemResult Carregar(CarregarItemRequest req, int codUsuario)
{
    using (var con = new SqlConnection(Config.StringConexao()))
    {
        // 1) FC_VALIDA_SAIDA_CARGA antes da transacao
        var bl = con.QueryFirstOrDefault<int?>(
            "SELECT BL FROM SGIPA..TB_CARGA_SOLTA WHERE AUTONUM = @AUTONUMCS",
            new { req.AUTONUMCS_OC });
        if (bl == null) return new CarregarItemResult { SUCESSO = false, MENSAGEM = "Lote nao encontrado" };

        var validacao = con.ExecuteScalar<string>(
            "SELECT dbo.FC_VALIDA_SAIDA_CARGA(@BL, @ORDEM, @AUTONUMCS) AS RETORNO",
            new { BL = bl, req.ORDEM, AUTONUMCS = req.AUTONUMCS_OC });
        if (!string.Equals(validacao, "OK", StringComparison.OrdinalIgnoreCase))
            return new CarregarItemResult { SUCESSO = false, MENSAGEM = validacao };

        // 2) Transacao
        con.Open();
        using (var tx = con.BeginTransaction(IsolationLevel.ReadCommitted))
        {
            try
            {
                con.Execute(
                    "UPDATE SGIPA..TB_MARCANTES SET ORDEM=@ORDEM, PLACA_C=@PLACA_C WHERE AUTONUM=@MARCANTE",
                    new { req.ORDEM, req.PLACA_C, req.MARCANTE }, tx);

                // qtde original do marcante
                var qtdeOriginal = con.QueryFirstOrDefault<int>(
                    "SELECT VOLUMES FROM SGIPA..TB_MARCANTES WHERE AUTONUM=@MARCANTE",
                    new { req.MARCANTE }, tx);
                var parcial = req.QTDE < qtdeOriginal;

                if (!parcial)
                    con.Execute(
                        "UPDATE SGIPA..TB_CARGA_SOLTA_YARD SET QUANTIDADE=0 WHERE AUTONUM=@ID",
                        new { ID = req.AUTONUM_CS_YARD }, tx);
                else
                    con.Execute(
                        "UPDATE SGIPA..TB_CARGA_SOLTA_YARD SET QUANTIDADE=QUANTIDADE-@Q WHERE AUTONUM=@ID",
                        new { Q = req.QTDE, ID = req.AUTONUM_CS_YARD }, tx);

                var novoId = con.QuerySingle<int>(@"
                    INSERT INTO SGIPA..TB_CARGA_SOLTA_YARD
                        (AUTONUM_CS, ARMAZEM, YARD, ORIGEM, QUANTIDADE, MOTIVO)
                    OUTPUT INSERTED.AUTONUM
                    VALUES (@AUTONUMCS, @ARMAZEM, 'CAM', 'I', @Q, 8)",
                    new { req.AUTONUMCS_OC, req.AUTONUM_ARMAZEM, ARMAZEM = req.AUTONUM_ARMAZEM, Q = req.QTDE }, tx);

                if (!parcial)
                    con.Execute(
                        "UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CS_YARD=@NOVO WHERE AUTONUM=@MARCANTE",
                        new { NOVO = novoId, req.MARCANTE }, tx);
                else
                {
                    con.Execute(
                        "UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CS_YARD=@NOVO, VOLUMES=@Q WHERE AUTONUM=@MARCANTE",
                        new { NOVO = novoId, Q = req.QTDE, req.MARCANTE }, tx);
                    con.Execute(
                        "UPDATE SGIPA..tb_temp_marcante SET QTDE=@Q WHERE MARCANTE=@TXT",
                        new { Q = req.QTDE, TXT = req.MARCANTE.ToString("D9") }, tx);
                }

                con.Execute(@"INSERT INTO SGIPA..TB_HIST_SHIFTING_CS
                    (MARCANTE, ARMAZEM, YARD, DT_MOV, USUARIO)
                    VALUES (@MARCANTE, 0, 'CAM', GETDATE(), @USUARIO)",
                    new { req.MARCANTE, USUARIO = codUsuario }, tx);

                tx.Commit();
                return new CarregarItemResult
                {
                    SUCESSO = true,
                    PARCIAL = parcial,
                    MENSAGEM = parcial ? "Item parcialmente carregado" : "Item carregado"
                };
            }
            catch (Exception ex)
            {
                tx.Rollback();
                return new CarregarItemResult { SUCESSO = false, MENSAGEM = "Erro: " + ex.Message };
            }
        }
    }
}
```

> Trechos acima sao referencia para implementacao na fase do Agente 07. Nada e implementado neste estudo.
