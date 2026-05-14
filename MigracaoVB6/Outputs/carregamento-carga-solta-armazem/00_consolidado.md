# Consolidado - Carregamento Carga Solta (Armazem)

## Resumo executivo
- Tela: Carregamento Carga Solta (Coletor Armazem) - form `CarregaCS` em `ProjetoReferencia/Coletor/CarregaCS.frm`.
- Objetivo: registrar carregamento de marcantes de carga solta de importacao a partir do armazem em um veiculo (placa cavalo + carreta) ja agendado, com transacao em `TB_CARGA_SOLTA_YARD` + `TB_MARCANTES` + historico em `TB_HIST_SHIFTING_CS`. Suporta cancelamento (`F5`), consulta de locais no armazem (`F1`) e consulta de documentos (`F2`, via exe externo).
- Status do estudo: concluido (Agentes 01..06).
- Status de prontidao para implementacao: **PRONTO** para Agente 07, com 4 pendencias de validacao listadas em "Pendencias".

## Principais achados
- Achado 1: A tela alvo e exclusivamente do projeto Armazem (`Coletor.vbp` -> `ColetorArm.exe`); o arquivo similar `CarregaCSPatio.frm` pertence ao projeto Patio (`ColetorPatioSmall.vbp`) e NAO faz parte desta migracao.
- Achado 2: Toda a regra de bloqueio de saida e centralizada na funcao escalar `dbo.FC_VALIDA_SAIDA_CARGA(BL, ORDEM, AUTONUMCS)` que retorna `'OK'` ou texto de critica - manter como-esta.
- Achado 3: O `INSERT` em `TB_CARGA_SOLTA_YARD` no legado nao usa `OUTPUT INSERTED.AUTONUM`; ele faz um `SELECT MAX(AUTONUM) WHERE AUTONUM_CS=...` apos o insert - ponto que sera corrigido na MVC para evitar troca de IDs sob concorrencia.
- Achado 4: O campo `txtMarcante` tem logica de "anti-digitacao manual" baseada em `DateDiff` entre `KeyDown(Enter)` e a digitacao - so permite bipagem quando `TB_CAD_USUARIOS.flag_carregamento_marcante_bip <> 0`.
- Achado 5: `MOTIVO=8` fixo em todos os movimentos do form (carregar e cancelar). Codigo `'CAM'` representa "carregado em caminhao" e `'CANCC'` "cancelado".
- Achado 6: Carregamento parcial atualiza `TB_MARCANTES.VOLUMES` E `tb_temp_marcante.QTDE` - tabela compartilhada com a tela "Associacao de Marcantes" - exige cuidado.
- Achado 7: F5 (`cmdF5_Click`) NAO grava em `TB_HIST_SHIFTING_CS` - pode ser bug do legado; proposta de correcao na MVC.
- Achado 8: O `UPDATE TB_MARCANTES SET ORDEM=..., PLACA_C=...` ocorre antes de `BeginTrans` no F4 - na MVC sera trazido para dentro da transacao.

## Regras criticas
- Regra: O `AUTONUMCS` da carga vinculada ao marcante deve coincidir com o `AUTONUMCS` da OC selecionada (caso contrario "Carga Divergente da Ordem de carregamento").
- Regra: A quantidade carregada nao pode ultrapassar o `VOLUMES` original do marcante (`QTDE > VOLUMES` -> "Quantidade indisponivel").
- Regra: `dbo.FC_VALIDA_SAIDA_CARGA` deve retornar `'OK'` antes de qualquer escrita.
- Regra: Marcante com `LOCAL='CAM'` NAO pode ser carregado novamente; Marcante com `LOCAL <> 'CAM'` nao pode ser cancelado.
- Regra: Filtro de patio do operador (`Col_Patio`) define os veiculos visiveis. `Col_Patio=1` -> `PATIO IN (1,7)`.
- Regra: `Wflag_carregamento_marcante_bip <> 0` exige bipagem; `=1` desabilita o campo apos validar.
- Regra: Apos sucesso (carregar/cancelar), foco volta ao `txtMarcante`; apos limpar, idem.
- Regra: Acesso a tela exige permissao `COLETOR_ARM` para `COMMAND1(7)` no legado (a equivalencia web precisa ser definida).

## Decisoes de migracao
- Decisao: Criar modulo dedicado (`CarregamentoCargaSolta`).
  - Justificativa: regras especificas e fluxo transacional proprio, sem encaixe em modulos existentes.
- Decisao: Trazer o `UPDATE TB_MARCANTES (ORDEM, PLACA_C)` para dentro da transacao do `Carregar`.
  - Justificativa: consistencia em caso de falha entre o update e o insert da yard.
- Decisao: Usar `OUTPUT INSERTED.AUTONUM` no insert de `TB_CARGA_SOLTA_YARD`.
  - Justificativa: corrigir potencial troca de IDs sob concorrencia.
- Decisao: Gravar historico tambem no `Cancelar` (`YARD='CANCC'`).
  - Justificativa: rastreabilidade. Sujeito a confirmacao com operacao.
- Decisao: Substituir o `Adodc1` por consulta unica via Dapper no Index.
  - Justificativa: padrao do Romaneio; remove canal duplo de dados.
- Decisao: Persistir `flag_carregamento_marcante_bip` em `Session` no login do Romaneio.
  - Justificativa: replicar comportamento do legado (carrega uma vez no login).
- Decisao: Implementar `F2 Documento` em release posterior ou com placeholder.
  - Justificativa: integracao com `ConsultaDocColetor.exe` exige analise separada; fluxo principal nao depende disso.

## Pendencias de validacao
- Pendencia: Confirmar significado de `MOTIVO=8` fixo nos inserts de `TB_CARGA_SOLTA_YARD`.
  - Responsavel: time de operacao / DBA.
- Pendencia: Validar a hipotese de gravar historico no `F5` (cancelamento).
  - Responsavel: time de operacao.
- Pendencia: Validar comportamento de carregamento parcial e impacto em `tb_temp_marcante` (compartilhamento com `Associacao de Marcantes`).
  - Responsavel: time de operacao + dev responsavel pela tela `Associacao de Marcantes`.
- Pendencia: Confirmar permissionamento web equivalente a `Valida_Acesso_Botao` (`COLETOR_ARM`).
  - Responsavel: lider tecnico do Romaneio.
- Pendencia: Definir como tratar o `F2` (documentos) em ambiente web.
  - Responsavel: arquiteto + product owner.
- Pendencia: Confirmar disponibilidade da function `dbo.FC_VALIDA_SAIDA_CARGA` no banco do ambiente alvo.
  - Responsavel: DBA.

## Escopo da implementacao (fase atual)
- Incluido:
  - Tela web equivalente em `Views/CarregamentoCargaSolta/Index.cshtml`.
  - Combo de veiculo, lista de OCs, lista de carregamentos, modal Locais.
  - Bipagem do marcante com carga automatica de dados.
  - Carregar (`F4`) total e parcial, com transacao consistente.
  - Cancelar (`F5`).
  - Limpar (`F3`), Sair (`Esc`).
  - Atalhos F1..F5 + ESC.
  - Anti-digitacao manual (replicado para web).
  - Mensageria padronizada (banner sucesso/erro).
- Fora de escopo (release subsequente):
  - Integracao do `F2 DOC` com sistema de anexos web.
  - Telemetria/auditoria adicional alem do `TB_HIST_SHIFTING_CS`.
  - Tela equivalente para Patio (`CarregaCSPatio.frm`) - tema separado.
  - Migracao da function `FC_VALIDA_SAIDA_CARGA` para C# (mantemos no banco).
  - Substituicao da tabela compartilhada `tb_temp_marcante` (refator estrutural).

## Arquivos VB6 relacionados
- `ProjetoReferencia/Coletor/CarregaCS.frm` (form alvo, 1305 linhas).
- `ProjetoReferencia/Coletor/CarregaCS.frx` (binding `DataCombo1`).
- `ProjetoReferencia/Coletor/Principal.frm` (entrada do menu, `Command1(7)`).
- `ProjetoReferencia/Coletor/FrmAcessoAnt.frm` (login - define `Wflag_carregamento_marcante_bip`).
- `ProjetoReferencia/Coletor/Microled.bas` (`db`, `ConnectionString`, `Cod_Usuario`, `Nnull`, `PRSet`).
- `ProjetoReferencia/Coletor/mdlColetor.bas` (`Col_Patio`, `Col_Login`, `Valida_Acesso_Botao` com sistema `COLETOR_ARM`).
- `ProjetoReferencia/Coletor/Coletor.vbp` (composicao do projeto Armazem).

## Objetos de banco envolvidos
- Views:
  - `SGIPA..VW_CAM_CARREGAMENTO`
  - `SGIPA..VW_INVENT_ARMAZEM`
- Tabelas:
  - `SGIPA..TB_MARCANTES`
  - `SGIPA..TB_CARGA_SOLTA_YARD`
  - `SGIPA..TB_CARGA_SOLTA`
  - `SGIPA..TB_HIST_SHIFTING_CS`
  - `SGIPA..TB_BL`
  - `SGIPA..TB_SOLICITACAO_LTL`
  - `SGIPA..TB_AG_CS`
  - `SGIPA..TB_ARMAZENS_IPA`
  - `SGIPA..TB_CNTR_BL`
  - `SGIPA..DTE_TB_EMBALAGENS`
  - `SGIPA..tb_temp_marcante`
  - `SGIPA..TB_CAD_USUARIOS` (login - flag bipagem)
  - `SGIPA..TB_SYS_FUNCOES`/`TB_SYS_GRP_PERMISSOES`/`TB_SYS_USER_GRUPOS` (permissionamento)
- Functions:
  - `SGIPA..DBO.FC_VALIDA_SAIDA_CARGA(BL int, ORDEM int, AUTONUMCS int) -> varchar` (escalar; `'OK'` libera, qualquer outro retorno e critica).

## Proposta inicial de migracao MVC
- Stack: ASP.NET MVC C# + Dapper (padrao do Romaneio).
- Componentes novos:
  - `Controllers/CarregamentoCargaSoltaController.cs` (DefaultController).
  - `Repositorios/CarregamentoCargaSoltaRepositorio.cs`.
  - `Interfaces/Repositories/ICarregamentoCargaSoltaRepositorio.cs`.
  - `Models/ViewModels/CarregamentoCargaSoltaViewModel.cs` (com DTOs aninhados).
  - `Views/CarregamentoCargaSolta/Index.cshtml`.
  - `Content/js/carregamento-carga-solta.js`.
- Edicoes:
  - `Global.asax.cs` (registro Unity).
  - `Romaneio.csproj` (incluir arquivos).
  - `LoginController` (popular `Session["FlagCarregamentoMarcanteBip"]`).
- Endpoints (5 GET + 2 POST + 1 GET de documento):
  - `GET Index`, `GET ListarVeiculos`, `GET ListarOrdens`, `GET ListarCarregamentos`, `GET ListarLocais`, `GET BuscarMarcante`, `GET Documento`.
  - `POST Carregar`, `POST Cancelar`.
- Persistencia:
  - Transacao `ReadCommitted` em Carregar/Cancelar.
  - `OUTPUT INSERTED.AUTONUM` na escrita.
  - Function `FC_VALIDA_SAIDA_CARGA` chamada antes da transacao.

## Riscos e duvidas abertas
- Risco: `FC_VALIDA_SAIDA_CARGA` evoluir no banco e quebrar a tela sem mudanca de codigo (caixa-preta legada).
- Risco: Concorrencia entre dois operadores na mesma OC - mitigado com `OUTPUT INSERTED`.
- Risco: Compartilhamento de `tb_temp_marcante` com outra tela.
- Risco: Comportamento do leitor de codigo de barras (timing de Enter) varia por modelo.
- Risco: `ConsultaDocColetor.exe` nao tem equivalente direto na web.
- Risco: Permissionamento web pode nao replicar fielmente `Valida_Acesso_Botao` (`sistema='COLETOR_ARM'`).
- Duvida: O `MOTIVO=8` deve permanecer fixo? Existe parametrizacao por usuario/patio?
- Duvida: `'CANCC'` e o codigo correto de yard para cancelamento? Ele aparece em algum dashboard?
- Duvida: O carregamento parcial e usado de fato? Qual o caso de uso?
- Duvida: O update em `tb_temp_marcante` e necessario, ou foi resquicio? Pode ser removido na migracao?

## Criterios de aceite
- Carregamento total funciona ponta-a-ponta para uma OC valida com marcante valido (resultado: registro `'CAM'` com QTDE igual ao original; marcante religado).
- Carregamento parcial reduz `VOLUMES` no marcante e ajusta `tb_temp_marcante`.
- Cancelamento reverte um marcante carregado, libera `PLACA_C` e atualiza listas.
- A funcao `FC_VALIDA_SAIDA_CARGA` e respeitada (retorno != 'OK' bloqueia operacao).
- Mensagens funcionais preservadas (texto do legado mantido em essencia).
- Atalhos F1..F5 + ESC funcionando como no legado.
- Combo de veiculo, lista de ordens e lista de carregamentos exibem dados corretos para um operador com `Col_Patio` definido.
- Bipagem obrigatoria (`flag_bip <> 0`) bloqueia digitacao manual.
- Sessao expirada retorna `{ success: false, message: "Sessao caiu" }` em endpoints POST.
- Sem regressao em telas adjacentes (`MovimentacaoCS`, `AssociacaoMarcantes`).
- Codigo segue padrao do Romaneio (DefaultController, Dapper, Bootstrap, JS por tela).

## Lista dos arquivos preenchidos em `MigracaoVB6/Outputs/carregamento-carga-solta-armazem/`
- `01_diagnostico_vb6.md`
- `02_mapeamento_funcional.md`
- `03_mapeamento_visual.md`
- `04_regras_negocio.md`
- `05_desenho_mvc.md`
- `06_plano_implementacao.md`
- `00_consolidado.md`

## Proximas etapas
1. Validar com lider tecnico/DBA as 4 pendencias criticas (FC_VALIDA_SAIDA_CARGA, MOTIVO=8, historico no F5, permissionamento web).
2. Acionar Agente 07 (`MigracaoVB6/Agents/07_agente_execucao_mvc.md`) com este consolidado para iniciar a implementacao no projeto `Romaneio` seguindo o plano de 5 etapas.
3. Apos primeira versao funcional, agendar homologacao com operador real (cenarios CT-01..CT-05 e CE-01..CE-07 e CB-01..CB-07).
