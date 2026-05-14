# Diagnostico VB6 - Localizacao Carregamentos

## Identificacao da tela
- Nome: Localizacao Carregamentos (menu: "Localizacao Carregamentos" / "Localizacao - Carregamentos").
- Modulo: Coletor (VB6).
- Form principal: `Localizacao` (variantes em `LocalizacaoS.frm` e `Localizacao.frm`).
- Caminho: `ProjetoReferencia/Coletor`.

## Arquivos localizados
- `.frm`:
  - `ProjetoReferencia/Coletor/LocalizacaoS.frm` (versao com comando de "SAIDA" para conteiner).
  - `ProjetoReferencia/Coletor/Localizacao.frm` (versao sem botao de "SAIDA"; foco em consulta).
  - `ProjetoReferencia/Coletor/PrincipalS.frm` (abre `Localizacao.Show` no `Command1(Index=8)`).
  - `ProjetoReferencia/Coletor/PrincipalG.frm` (abre `Localizacao.Show` no `Command1(Index=8)`).
- `.frx`:
  - `ProjetoReferencia/Coletor/LocalizacaoS.frx`
  - `ProjetoReferencia/Coletor/Localizacao.frx`
- `.bas`:
  - `ProjetoReferencia/Coletor/Microled.bas` (dependencias globais inferidas por uso: `db`, `Nnull`, `ConnectionStringAdo`, `Executa`, banco).
  - `ProjetoReferencia/Coletor/mdlColetor.bas` (globais operacionais inferidas: `Col_Patio`, mensagens de sessao).
- `.cls`:
  - Nao foi encontrada evidencia direta para esta tela.
- Outros:
  - `ProjetoReferencia/Coletor/ColetorPatioSmall.vbp` (inclui `LocalizacaoS.frm`).
  - `ProjetoReferencia/Coletor/ColetorPatioNaoUsando.vbp` (inclui `Localizacao.frm`).

## Eventos mapeados
- Form_Load:
  - Busca descricao do patio em `operador..tb_patios` e preenche `TxtPatio`.
  - `LocalizacaoS`: chama `Carrega_Grid1` no load.
  - `Localizacao`: deixa carga automatica comentada no load e executa no activate.
- Form_Activate:
  - Recarrega a grid (`Carrega_Grid1`) em pelo menos uma variante.
- Form_KeyDown:
  - `ESC` fecha a tela (`Unload Me`).
- Clicks de botoes:
  - `Command1_Click`: aplica filtro por placa/carreta e recarrega a grid.
  - `CmdSaida_Click` (apenas `LocalizacaoS`): executa fluxo de movimentacao para `SAIDA`.
- Eventos de lista/grid:
  - `Grid1_Click` e `Grid1_RowColChange` chamam `popula`.
- Eventos de foco:
  - Sem foco automatico relevante alem da selecao de linha.

## SQLs e persistencia
- Consultas:
  - Query principal em `Carrega_Grid1` com `UNION ALL` entre:
    - bloco de conteiner (`tb_gate_new` + `tb_amr_gate` + `TB_CNTR_BL` + `TB_ORDEM_CARREGAMENTO` + `tb_remocao` + transportadora + usuario).
    - bloco de carga solta (`tb_gate_new` + `tb_amr_gate` + `TB_CARGA_SOLTA`/`tb_registro_saida_cs` ou `VW_LOCAL_OC` conforme variante).
  - Filtro opcional por placa/carreta com `mskPlaca`.
  - Filtro obrigatorio por patio (`Col_Patio`, com regra especial de patio 1/7 na variante `S`).
- Inserts:
  - `LocalizacaoS/CmdSaida_Click`: `INSERT` em `TB_HIST_SHIFTING`.
- Updates:
  - `LocalizacaoS/CmdSaida_Click`: `UPDATE sgipa..tb_cntr_bl SET yard='SAIDA'`.
- Deletes:
  - Nao ha evidencia direta.
- Views/Tabelas/Functions citadas:
  - `operador..tb_gate_new`
  - `operador..tb_amr_gate`
  - `sgipa..TB_CNTR_BL`
  - `sgipa..TB_ORDEM_CARREGAMENTO`
  - `operador..TB_CAD_TRANSPORTADORAS`
  - `operador..tb_remocao`
  - `sgipa..TB_CARGA_SOLTA`
  - `sgipa..DTE_TB_EMBALAGENS`
  - `sgipa..tb_registro_saida_cs`
  - `sgipa..VW_LOCAL_OC`
  - `sgipa..tb_cad_usuarios`
  - `operador..tb_patios`
  - `operador..VW_INVENT_SISTEMAS`
  - `operador..TB_YARD`
  - `sgipa..TB_HIST_SHIFTING`

## Dependencias tecnicas
- Variaveis globais:
  - `Col_Patio`, `Col_Tam_Pos_Yard`, `Cod_Usuario`, `Col_Mensagem`, `Col_Cancelado`, `Col_Imagem`.
- Funcoes utilitarias:
  - `Nnull`, `Executa`, `IIf`, `db.Execute`.
- Formularios auxiliares:
  - `Frmmensagem`, `frmMensagem2`.
- Controle de permissao:
  - Nao ha validacao de permissao explicita dentro do form; depende da entrada por menu/login.

## Atalhos e automatismos
- Atalhos:
  - `ESC` para sair.
- Foco/atualizacao automatica:
  - Recarregamento automatico da grid no load/activate.
- Comportamentos implicitos:
  - Contagem automatica por tipo (`TOTCNTR` e `TOTCS`) via coluna flag.
  - Detalhe da linha selecionada preenche bloco superior de campos somente leitura.

## Pontos criticos
- Riscos:
  - SQL montado por concatenacao (manutencao e risco de entrada).
  - Janela temporal fixa (`getdate()-1` / `getdate()-3`) pode ocultar casos fora da faixa.
  - Fluxo de `CmdSaida_Click` sem transacao explicita entre update e insert.
- Inconsistencias:
  - Duas variantes de form com diferencas funcionais (consulta pura vs consulta + saida).
  - Diferenca de fonte para carga solta (`tb_registro_saida_cs` em uma variante, `VW_LOCAL_OC` em outra).
- Hipoteses:
  - Hipotese: versao operacional atual para patio pequeno e `LocalizacaoS.frm` (presente no `ColetorPatioSmall.vbp` e menu principal S).
  - Hipotese: `Localizacao.frm` e usada em outro pacote/compilacao legado com escopo mais restrito.
