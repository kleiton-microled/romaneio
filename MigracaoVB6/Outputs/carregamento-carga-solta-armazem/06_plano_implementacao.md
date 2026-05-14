# Plano de Implementacao - Carregamento Carga Solta (Armazem)

## Etapas

### Etapa 1 - Base backend (contratos e leitura)
- Escopo:
  - Criar `Models/ViewModels/CarregamentoCargaSoltaViewModel.cs` com VM principal e DTOs (`VeiculoDto`, `OrdemCarregamentoDto`, `MarcanteCarregadoDto`, `LocalArmazemDto`, `MarcanteDadosDto`, requests/results de Carregar/Cancelar/Documento).
  - Criar `Interfaces/Repositories/ICarregamentoCargaSoltaRepositorio.cs` com assinaturas das operacoes.
  - Criar `Repositorios/CarregamentoCargaSoltaRepositorio.cs` implementando inicialmente as queries de leitura (`ListarVeiculos`, `ListarOrdens`, `ListarCarregamentos`, `ListarLocais`, `BuscarMarcante`).
  - Adicionar registro no Unity em `Global.asax.cs`.
  - Atualizar `Romaneio.csproj` com os novos arquivos.
- Entregaveis:
  - 4 arquivos novos compilando sem alterar comportamento existente.
- Criterio de pronto:
  - `Romaneio` compila sem erros.
  - Queries testaveis manualmente via repositorio (chamadas unitarias diretas) retornam dados coerentes com as queries do VB6 (mesmas linhas para a mesma placa).

### Etapa 2 - Controller e endpoints (sem operacoes de escrita)
- Escopo:
  - Criar `Controllers/CarregamentoCargaSoltaController.cs` herdando `DefaultController`.
  - Implementar `Index`, `ListarVeiculos`, `ListarOrdens`, `ListarCarregamentos`, `ListarLocais`, `BuscarMarcante`, `Documento` (apenas pipeline LTL/AG_CS sem abrir nada ainda).
  - Adicionar checagens de `Session["Logado"]` e leitura de `Session["Patio"]` e `Session["FlagCarregamentoMarcanteBip"]` (com fallback `0` se a sessao nao existir).
  - Atualizar `LoginController` (ou equivalente) para popular `Session["FlagCarregamentoMarcanteBip"]` a partir de `TB_CAD_USUARIOS.flag_carregamento_marcante_bip`.
- Entregaveis:
  - Controller exposto e endpoints respondendo `application/json` no formato `{ success, items|dados|message }`.
- Criterio de pronto:
  - Endpoints navegaveis via Postman/Browser retornam payload correto.
  - Verificar com `Session` simulado (browser logado) acesso liberado e bloqueio quando deslogado.

### Etapa 3 - View e scripts
- Escopo:
  - Criar `Views/CarregamentoCargaSolta/Index.cshtml` seguindo o estilo de `Views/MovimentacaoCS/Index.cshtml` (cards, inputs, botoes Bootstrap).
  - Criar modal `#modalLocais` dentro da view.
  - Criar `Content/js/carregamento-carga-solta.js`:
    - Inicializa estado (`patio`, `flagBipObrigatoria`, `flagBlockApos`).
    - Bindings dos botoes/atalhos (F1..F5, ESC).
    - `loadOrdens(placa)`, `loadCarregamentos(placa, autonumCs)`, `loadLocais(autonumCs)`, `buscarMarcante(marcante)`.
    - Controle de selecao da OC (data attributes na linha).
    - Logica de anti-digitacao manual (timestamps de `keydown`/`Enter`).
    - Renderizacao da critica em banner.
- Entregaveis:
  - View com layout responsivo + JS funcional para fluxos de leitura.
- Criterio de pronto:
  - Selecionar veiculo carrega ordens.
  - Selecionar ordem carrega carregamentos.
  - F1 abre modal Locais com dados.
  - F2 chama endpoint Documento (placeholder).
  - F3 limpa campos.
  - Bipar marcante popula campos.
  - ESC fecha modal/sai.

### Etapa 4 - Persistencia: Carregar e Cancelar (transacionais)
- Escopo:
  - Implementar `Carregar` no repositorio com transacao `BeginTransaction(IsolationLevel.ReadCommitted)`:
    - Chamada `FC_VALIDA_SAIDA_CARGA` antes da transacao.
    - Update `TB_MARCANTES (ORDEM, PLACA_C)` dentro da transacao (decisao tecnica).
    - Update yard atual (zerar ou decrementar).
    - Insert `TB_CARGA_SOLTA_YARD` com `OUTPUT INSERTED.AUTONUM` (pega novo id atomicamente).
    - Update `TB_MARCANTES (AUTONUM_CS_YARD, [VOLUMES])`.
    - Update `tb_temp_marcante (QTDE)` em parcial.
    - Insert `TB_HIST_SHIFTING_CS`.
    - Commit/Rollback.
  - Implementar `Cancelar` no repositorio com transacao similar:
    - Update yard atual zerando.
    - Insert `TB_CARGA_SOLTA_YARD` com `YARD='CANCC'`.
    - Update `TB_MARCANTES (AUTONUM_CS_YARD, PLACA_C=NULL)`.
    - Insert opcional em `TB_HIST_SHIFTING_CS` com `YARD='CANCC'` (proposta - ver decisoes).
  - Endpoints `POST Carregar` e `POST Cancelar` no controller.
  - JS: chamar via `$.ajax` POST JSON, atualizar listas pos sucesso, mostrar critica em vermelho/verde.
- Entregaveis:
  - Operacoes de carregar/cancelar funcionais e transacionalmente seguras.
- Criterio de pronto:
  - F4 carrega marcante; F5 cancela; lvOC e lvCm refletem alteracao apos cada acao; mensagens corretas; transacoes consistentes em sucesso e em erro forcado.

### Etapa 5 - Refinamentos finais (UX, atalhos, integracao Documento)
- Escopo:
  - Anti-digitacao manual: refinar JS com testes em leitor real.
  - Bloqueio do `txtMarcante` apos validar quando `BIPAGEM_BLOQUEAR_APOS_VALIDAR=true`.
  - Atalhos F1..F5/ESC com `preventDefault`, considerando navegadores que reservam F1.
  - Tooltips/ajustes visuais (badge "PLACA_C ja vinculada", progress bar `n/m`).
  - Endpoint `Documento`: implementar adaptacao web (ver decisao em 05_desenho_mvc).
  - Logging dos erros de transacao (via mecanismo existente do projeto, se houver).
  - Permissionamento: garantir bloqueio quando `Session` nao indica permissao para `COMMAND1(7)` / `COLETOR_ARM` (alinhar com como o Romaneio trata permissionamento - confirmar se `Session["Logado"]` e suficiente ou se ha tabela de permissoes web).
- Entregaveis:
  - Tela "polida" para uso operacional; documentacao breve no `doc/` (opcional).
- Criterio de pronto:
  - Sem regressao em outras telas.
  - Cenarios negativos exibem mensagens corretas.
  - Performance aceitavel (queries < 1s em ambiente de homologacao com dados reais).

## Checklist tecnico
- [ ] Backend
  - [ ] ViewModel + DTOs criados
  - [ ] Interface de repositorio criada
  - [ ] Repositorio implementado com Dapper parametrizado
  - [ ] Registrado no Unity (`Global.asax.cs`)
  - [ ] Adicionado ao `Romaneio.csproj`
- [ ] Frontend
  - [ ] `Index.cshtml` no padrao Bootstrap das outras telas
  - [ ] `carregamento-carga-solta.js` em `Content/js/`
  - [ ] Atalhos F1..F5 + ESC funcionando
  - [ ] Modal Locais funcional
  - [ ] Banner de critica (sucesso/erro) padronizado
- [ ] Banco
  - [ ] Validacao manual: `dbo.FC_VALIDA_SAIDA_CARGA` retorna `OK` para casos esperados
  - [ ] `OUTPUT INSERTED.AUTONUM` em `TB_CARGA_SOLTA_YARD` testado
  - [ ] Conferir indices em `TB_MARCANTES.AUTONUM` e `TB_CARGA_SOLTA_YARD.AUTONUM_CS`
- [ ] Rotas
  - [ ] `/CarregamentoCargaSolta/Index`
  - [ ] `/CarregamentoCargaSolta/ListarVeiculos`
  - [ ] `/CarregamentoCargaSolta/ListarOrdens`
  - [ ] `/CarregamentoCargaSolta/ListarCarregamentos`
  - [ ] `/CarregamentoCargaSolta/ListarLocais`
  - [ ] `/CarregamentoCargaSolta/BuscarMarcante`
  - [ ] `/CarregamentoCargaSolta/Carregar`
  - [ ] `/CarregamentoCargaSolta/Cancelar`
  - [ ] `/CarregamentoCargaSolta/Documento`
- [ ] Validacoes
  - [ ] Marcante 9 digitos
  - [ ] OC selecionada antes de F1/F4
  - [ ] AUTONUMCS marcante = AUTONUMCS OC
  - [ ] LOCAL diferente de `'CAM'` para F4
  - [ ] LOCAL = `'CAM'` para F5
  - [ ] QTDE <= QTDE original
  - [ ] FC_VALIDA_SAIDA_CARGA = `'OK'`
  - [ ] Bipagem manual quando flag ativa
- [ ] Tratamento de erros
  - [ ] Rollback em qualquer falha de transacao
  - [ ] Mensagens funcionais preservadas
  - [ ] `try/catch` em endpoint POST com retorno padronizado
- [ ] Testes (manuais)
  - [ ] Cenario feliz: bipagem total
  - [ ] Cenario feliz: bipagem parcial
  - [ ] Cenario feliz: cancelamento
  - [ ] Cenario erro: marcante invalido
  - [ ] Cenario erro: AUTONUMCS divergente
  - [ ] Cenario erro: ja carregado / nao carregado
  - [ ] Cenario erro: FC_VALIDA_SAIDA_CARGA != OK
  - [ ] Cenario UX: flag bip ativada (rejeita digitacao)
  - [ ] Cenario UX: ESC fecha modal e nao a tela

## Plano de testes
- Cenarios felizes:
  - **CT-01** Carregamento total - bipar marcante valido com qtde original; F4 -> "Item carregado"; OC mostra `n/m` atualizado; carregamentos lista o marcante.
  - **CT-02** Carregamento parcial - reduzir `txtQtde` para metade; F4 -> "Item parcialmente carregado"; verificar `TB_MARCANTES.VOLUMES` reduzida e `tb_temp_marcante.QTDE` atualizada.
  - **CT-03** Cancelamento - bipar marcante ja em CAM; F5 -> "Item cancelado"; lista de carregamentos remove o item; `TB_MARCANTES.PLACA_C` volta para NULL.
  - **CT-04** F1 Locais - selecionar OC e abrir Locais; verificar lista do armazem.
  - **CT-05** F2 Documentos - selecionar lote e disparar Documento; verificar pipeline LTL vs AG_CS.
- Cenarios de erro:
  - **CE-01** Marcante invalido (sem AUTONUM_CARGA): "Marcante nao informado/invalido" ou campos vazios.
  - **CE-02** Marcante de outra carga: "Carga Divergente da Ordem de carregamento".
  - **CE-03** Marcante ja em CAM + F4: "Carga ja carregada".
  - **CE-04** Marcante fora de CAM + F5: "Carga ainda nao carregada".
  - **CE-05** QTDE > original: "Quantidade indisponivel".
  - **CE-06** FC_VALIDA_SAIDA_CARGA != OK: exibir o retorno textual.
  - **CE-07** Lote nao encontrado em TB_CARGA_SOLTA: "Lote nao encontrado".
- Cenarios de borda:
  - **CB-01** Operador com `flag_carregamento_marcante_bip <> 0` digitando manualmente: alerta "Campo nao permite digitacao manual"; campo limpo.
  - **CB-02** Operador com `flag = 1`: apos validar marcante, campo fica disabled ate F3.
  - **CB-03** Concorrencia: dois operadores carregando OCs diferentes simultaneamente - cada uma com `OUTPUT INSERTED.AUTONUM` consistente (sem troca de IDs).
  - **CB-04** Refresh do combo de veiculos durante operacao - manter selecao se a placa ainda existir.
  - **CB-05** ESC com modal aberto - fecha apenas o modal.
  - **CB-06** Sessao caiu durante POST - controller retorna `{ success: false, message: "Sessao caiu" }`.
  - **CB-07** `Col_Patio = 1` - validar que combo lista placas dos patios `1` e `7`.

## Riscos e mitigacoes
- Risco: `FC_VALIDA_SAIDA_CARGA` possui regras complexas; uma alteracao no banco pode quebrar a tela web sem mudanca de codigo.
  - Mitigacao: documentar a dependencia, evitar replicar logica em codigo C# e ter alerta operacional em caso de retorno inesperado (ex.: log).
- Risco: Concorrencia entre dois operadores na mesma OC.
  - Mitigacao: `OUTPUT INSERTED.AUTONUM` ao inves de `MAX`. Considerar `UPDLOCK` em leituras criticas se ainda houver problemas.
- Risco: Comportamento de leitor de codigo de barras varia por modelo.
  - Mitigacao: testar com modelos reais; permitir override por `flag` se necessario.
- Risco: `ConsultaDocColetor.exe` nao funciona em ambiente web.
  - Mitigacao: separar essa funcionalidade em release subsequente; entregar a tela funcional sem F2 inicialmente, ou com placeholder de "em construcao".
- Risco: Tabela `tb_temp_marcante` e compartilhada com outra tela; alteracao parcial pode interferir.
  - Mitigacao: validar com time da `Associacao de Marcantes` antes de aplicar.
- Risco: Falta de permissionamento web equivalente ao `Valida_Acesso_Botao` (`COLETOR_ARM`).
  - Mitigacao: confirmar com lider tecnico se a checagem de `Session["Logado"]` e suficiente; senao, replicar consulta `tb_sys_funcoes/tb_sys_grp_permissoes` no servidor.
- Risco: `Wflag_carregamento_marcante_bip` pode nao estar populada na sessao do Romaneio (se o login foi feito antes da migracao do flag).
  - Mitigacao: ler na primeira chamada e cachear (fallback para `0` quando ausente).
- Risco: Comportamento divergente em carregamento parcial (ja existe no legado e e raramente usado).
  - Mitigacao: confirmar com usuario final antes de homologar.
