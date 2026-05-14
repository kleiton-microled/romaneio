# Regras de Negocio - Localizacao Carregamentos

## Regras identificadas por acao

### Filtrar
- Regra (negocio): listar apenas carregamentos em andamento (`dt_gate_out is null`) e recentes (janela de dias).
- Evidencia: `Carrega_Grid1` em `Localizacao.frm` e `LocalizacaoS.frm`.
- Classificacao: Regra de negocio.

- Regra (validacao): quando placa informada, filtrar por `GN.PLACA` ou `GN.CARRETA`.
- Evidencia: condicao com `mskPlaca`.
- Classificacao: Regra de validacao.

- Regra (tecnica): separar resultado de conteiner e carga solta por `UNION ALL`, com campo flag para contagem.
- Evidencia: coluna `FLAG_CNTR`.
- Classificacao: Regra tecnica.

### Incluir
- Regra: nao ha inclusao direta de carregamento nesta tela (consulta/localizacao).
- Evidencia: inexistencia de insert de inclusao no fluxo principal.
- Classificacao: Regra de interface.

### Remover
- Regra: nao ha remocao de registros de carregamento nesta tela.
- Evidencia: inexistencia de delete no form.
- Classificacao: Regra de interface.

### Salvar/Persistir
- Regra (somente variante `S`): permitir movimentar para `SAIDA` apenas conteiner selecionado.
- Evidencia: `CmdSaida.Visible` condicionado por `FLAG_CNTR`.
- Classificacao: Regra de negocio.

- Regra (validacao): bloquear `SAIDA` se item ja estiver em `SAIDA`.
- Evidencia: `If Me.txtQuadra.text = "SAIDA" Then ... Exit Sub`.
- Classificacao: Regra de validacao.

- Regra (validacao): bloquear `SAIDA` se houver conteiner acima na pilha.
- Evidencia: busca em `VW_INVENT_SISTEMAS` para `YARD` da altura superior.
- Classificacao: Regra de validacao.

- Regra (persistencia): ao confirmar, atualizar `TB_CNTR_BL.YARD='SAIDA'` e inserir historico em `TB_HIST_SHIFTING`.
- Evidencia: `UPDATE ... tb_cntr_bl` + `INSERT ... TB_HIST_SHIFTING`.
- Classificacao: Regra tecnica.

### Sair/Navegar
- Regra: `ESC` encerra a tela sem confirmar.
- Evidencia: `Form_KeyDown`.
- Classificacao: Regra de interface.

## Validacoes
- Campo: `mskPlaca`
  - Condicao: se preenchido, aplica filtro em placa/carreta.
  - Mensagem: nao ha mensagem especifica; apenas restringe dataset.
- Campo: item selecionado para `SAIDA`
  - Condicao: nao pode estar em `SAIDA`.
  - Mensagem: "Conteiner ja posicionado em SAIDA".
- Campo: empilhamento
  - Condicao: nao pode existir conteiner no nivel acima.
  - Mensagem: "Existe um conteiner acima, impossivel movimentar (...)"

## Regras de consistencia
- Quantidade:
  - Totais de conteiner e carga solta calculados por contagem de linhas com flag.
- Integridade referencial:
  - Join por chaves de gate, ordem e carga/conteiner nas tabelas operacionais.
- Concorrencia:
  - Nao ha controle transacional explicito no VB6 para o par `UPDATE + INSERT` da saida.
  - Hipotese: pode haver janela de inconsistência em falha parcial de persistencia.

## Efeitos colaterais
- Atualizacoes indiretas:
  - `YARD='SAIDA'` impacta visao de inventario e rastreabilidade de patio.
- Reuso de tabela temporaria:
  - Nao evidenciado nesta tela.
- Dependencias externas:
  - Dependencia de estado global (`Col_Patio`, `Cod_Usuario`, tamanho de yard).

## Inventario de banco (obrigatorio)
- Tabelas envolvidas:
  - `operador..tb_gate_new`
  - `operador..tb_amr_gate`
  - `operador..TB_CAD_TRANSPORTADORAS`
  - `operador..tb_remocao`
  - `operador..tb_patios`
  - `operador..TB_YARD`
  - `sgipa..TB_CNTR_BL`
  - `sgipa..TB_CARGA_SOLTA`
  - `sgipa..TB_ORDEM_CARREGAMENTO`
  - `sgipa..tb_registro_saida_cs` (em uma variante)
  - `sgipa..tb_cad_usuarios`
  - `sgipa..TB_HIST_SHIFTING`
- Views envolvidas:
  - `sgipa..VW_LOCAL_OC`
  - `operador..VW_INVENT_SISTEMAS`
- Procedures/functions envolvidas:
  - Nao ha procedure nomeada; operacoes via SQL direto.
- Campos lidos (principais):
  - placa, carreta, dt_lib_ent_cam, dt_gate_in, id_conteiner/carga, yard, temp_rua, fantasia, usuario, patio.
- Campos gravados:
  - `TB_CNTR_BL.YARD`
  - `TB_HIST_SHIFTING` (`CNTR`, `ORIGEM`, `DESTINO`, `DATA`, `TIPO`, `USUARIO`, `MOTIVO`)
- Joins relevantes:
  - gate -> amr_gate -> ordem/carga/conteiner + transportadora + usuario.
- Filtros obrigatorios:
  - `dt_gate_out is null`
  - janela temporal (`getdate()-1`/`getdate()-3`)
  - `patio` por sessao.
- Transacoes/commits:
  - Ausentes explicitamente no VB6.
- Pontos de concorrencia:
  - risco de update sem historico (ou vice-versa) em erro entre comandos.

## Hipoteses e duvidas
- Hipotese:
  - `LocalizacaoS.frm` e a variante principal em operacao atual do patio.
- Evidencia parcial:
  - Presenca no `ColetorPatioSmall.vbp` e menu `PrincipalS`.
- Acao de validacao:
  - Confirmar com usuario final e release de producao atual.

- Hipotese:
  - Diferenca de origem dos dados de carga solta (`tb_registro_saida_cs` vs `VW_LOCAL_OC`) reflete evolucao historica.
- Evidencia parcial:
  - SQLs divergentes entre `Localizacao.frm` e `LocalizacaoS.frm`.
- Acao de validacao:
  - Validar com DBA qual regra deve prevalecer no MVC.
