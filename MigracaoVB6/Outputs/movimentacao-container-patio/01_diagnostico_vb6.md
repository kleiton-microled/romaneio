# 01 — Diagnóstico VB6 — Movimentação Container: Pátio

## Identificação da tela

| Campo | Valor |
|--------|--------|
| Nome operacional (menu) | **Movimentação Conteiner** (`Principal.frm`, `PrincipalS.frm`, `PrincipalG.frm` — botão `Command1`, índice **4**) |
| Nome alinhado ao escopo deste estudo | **Movimentação Container: Pátio** (pacote **Coletor Pátio**, `ColetorPatioSmall.vbp`, `Title="Coletor Pátio"`) |
| Classe VB6 | `FrmMovimentacao` (`Attribute VB_Name = "FrmMovimentacao"`) |
| Forms físicos | Duas variantes de arquivo com **o mesmo `VB_Name`**: `FrmMovimentacaoS.frm` (pátio / layout compacto + aba Avarias) e `FrmMovimentacao.frm` (layout expandido, 3 abas principais sem aba “Avarias” no `SSTab` da versão analisada na árvore principal) |
| Evidência de compilação pátio | `ProjetoReferencia/Coletor/ColetorPatioSmall.vbp` referencia `Form=FrmMovimentacaoS.frm` |

**Observação:** O `.vbp` genérico `Coletor.vbp` presente no repositório **não lista** `FrmMovimentacao*.frm`; o menu `Principal.frm` mesmo assim chama `FrmMovimentacao.Show`. Tratar como **hipótese**: o `.vbp` local pode estar incompleto ou haver outro projeto de build — o código-fonte dos forms existe e é a fonte deste diagnóstico.

---

## Arquivos localizados

### `.frm` (núcleo e entrada)

| Arquivo | Papel |
|---------|--------|
| `ProjetoReferencia/Coletor/FrmMovimentacaoS.frm` | Variante **Pátio Small**: abas Moviment. / Histórico / Dados CNTR / **Avarias**; lógica principal de movimentação (`Atualiza_Posicao`, `Busca_Dados`, etc.) |
| `ProjetoReferencia/Coletor/FrmMovimentacao.frm` | Variante alternativa (layout maior; mesma classe `FrmMovimentacao`) — revisar diff pontual antes da implementação MVC se ambos forem usados em produção |
| `ProjetoReferencia/Coletor/PrincipalS.frm` | Menu **Coletor Pátio**: `Case 4` → `FrmMovimentacao.Show 1` |
| `ProjetoReferencia/Coletor/Principal.frm` | Menu: `Case 4` → `FrmMovimentacao.Show 1` |
| `ProjetoReferencia/Coletor/PrincipalG.frm` | Menu: `Case 4` → `FrmMovimentacao.Show 1` |
| `ProjetoReferencia/Coletor/Frmmensagem.frm` / `frmMensagem2.frm` | Mensagens e confirmações (`Frmmensagem`, `frmMensagem2`) |
| `ProjetoReferencia/Coletor/FrmVeiculoServ.frm` | Modal chamado em fluxos **SAÍDA** (EF=E) e **DTA** (evidência em `Atualiza_Posicao`) |

### `.frx`

- `FrmMovimentacaoS.frx` — bindings visuais (`Adodc`, `DataGrid`, `MSHFlexGrid`, etc.)
- `FrmMovimentacao.frx` — idem para a variante não-S

### `.bas` / módulos

| Módulo | Uso inferido |
|--------|----------------|
| `ProjetoReferencia/Coletor/Microled.bas` ou `ProjetoReferencia/Sgipa/Microled.bas` (conforme projeto) | `db`, `ConnectionString`, `ConnectionStringAdo`, `PRSet`, `Executa`, `Nnull`, etc. |
| `ProjetoReferencia/Coletor/mdlColetor.bas` | `Col_Mensagem`, `Col_Imagem`, `Col_Cancelado`, globais de sessão (`Col_Patio`, `Cod_Usuario`, …) |
| `ProjetoReferencia/Coletor/Segreg_IMO.bas` | `Verifica_Regras_Seg_Imo_Delta` — pós-movimentação com possível **estorno** via segunda chamada à SP |

### `.vbp`

- `ColetorPatioSmall.vbp` — **Coletor Pátio** + `FrmMovimentacaoS.frm`
- `ColetorPatioNaoUsando.vbp` — nome sugere pacote legado; não aprofundado neste estudo
- `Coletor.vbp` — não referencia explicitamente o form neste arquivo (ver hipótese acima)

---

## Controles relevantes (variante `FrmMovimentacaoS`)

- **Abas (`SSTab1`):** `Moviment.` | `Histórico` | `Dados CNTR` | `Avarias`
- **Identificação:** `CNTR` (atalho final — dispara `Busca_Cntr` no `Change`), `MskCntr` (máscara ID completo), `MskCntr.Tag` = sistema (`I` importação, `A` armazém IPA, `R` redex)
- **Pátio:** `txtPatio` (somente leitura; descrição vinda de `OPERADOR..TB_PATIOS`)
- **Posição atual / destino:** `TXTYardAtual`, `mskDestino`
- **Motivo:** `DataCombo2` → `OPERADOR..TB_cad_motivo` (`flag_ativo=1`)
- **Câmera (motivos com `flag_camera=1`):** `Frame2`, `DataCombo4` → `OPERADOR..TB_CAMERAS`, botões `Command3` (“Salvar” da câmera)
- **Ações principais:** `Command1` (“Salvar” — confirma movimentação), `Command2` (“Limpar”), `Command4` (“X” — fecha form), `Command5` (fecha `Frame5`)
- **Opções de grid auxiliar:** `Option1(0)` “ATUAL”, `Option1(1)` “DESTINO” — alternam critério de `Carrega_Grid2`
- **Grids:** `Grid1` (próximos movimentos), `Grid2` (unidades na pilha / vizinhança), `DataGrid1` (histórico shifting), `GridHist_Avaria`, `GridAvarias`
- **Checkbox:** `Check1` — caption **“LACRES CONFERIDOS”** (validação para histórico shifting não-GATEIN)
- **Campos diversos:** lacres IPA/Gate/Lona, tara, ISO, GMCI, IMO, Reefer (labels dinâmicos), blocos `Text1(0)` para mensagens contextuais

---

## Eventos mapeados (trechos verificados no código)

| Evento / rotina | Comportamento |
|-----------------|---------------|
| `Form_Load` | Prepara `Adodc3` (motivos), `Adodc31` (câmeras), descrição do pátio; **permissões** para aba Avarias (`TB_SYS_FUNCOES` código **998** / **999**) e botões incluir/excluir/finalizar |
| `Form_KeyPress` | Enter → `SendKeys "{tab}"` |
| `Form_Unload` | Chama `Unload Me` novamente (**risco de recursão**, padrão já visto em outras telas do legado) |
| `CNTR_Change` | Se não contém `_`, chama `Busca_Cntr` |
| `MskCntr_LostFocus` | `Busca_Dados` |
| `mskDestino_Change` / `LostFocus` | Recarrega grid 2; ajusta `DataCombo2.BoundText` conforme origem/destino (`SAIDA`, `GATEIN`, etc.) |
| `TXTYardAtual_LostFocus` | Mesma lógica de motivo padrão que `mskDestino_LostFocus` |
| `Option1_Click` | Atualiza label da pilha e `Carrega_Grid2` |
| `Command1_Click` | Valida motivo; se motivo exige câmera (`flag_camera=1`), exibe `Frame2` e **não** move ainda; senão `Atualiza_Posicao` |
| `Command3_Click` | Exige `DataCombo4`; oculta frame; `Atualiza_Posicao` |
| `Command2_Click` | `Limpa_Dados` + reinicia máscaras |
| `Command4_Click` | `Unload Me` |
| `SSTab1_Click` | Tab **1**: hist shifting + histórico de avarias; Tab **3**: listas e grids de avarias |

Rotinas internas principais: `Busca_Cntr`, `Busca_Dados`, `Carrega_Grid1`, `Carrega_Grid2`, `Limpa_Dados`, `Atualiza_Posicao`, `CarregaAvarias`, `cmdIncluirAvaria_Click`, `cmdExcluirAvaria_Click`, `cmdFinalizarAvaria_Click`, `Verifica_Finalizado`, `Valida_Tara`.

---

## SQL, funções e procedures

### Consultas / TVFs / views (amostra — não exaustiva)

- `SELECT * FROM operador.dbo.fnInvent_Sistemas_YardF('<final>', <Col_Patio>)` — resolve contêiner pelo final
- `SELECT * FROM operador.dbo.fnInvent_Sistemas_Col('<id>', <Col_Patio>)` — carrega dados do equipamento no pátio
- `SELECT ... FROM OPERADOR..VW_PROX_MOV_CNTR` + `VW_PROX_MOV_CNTR_COL` — grid de próximos movimentos (`Carrega_Grid1`)
- Consultas grandes a `SGIPA..VW_PROX_MOV_CNTR_ALL` e `OPERADOR..VW_PROX_MOV_CNTR_COL` com filtros por `PATIO`, `YARD LIKE`, regras de **pilha par/impar** (`Carrega_Grid2`)
- Validação pré-movimentação: `SELECT ... operador.dbo.fnMovimentaCntr(...)` — deve retornar **OK** na primeira coluna
- Bloqueio scanner: `SELECT SGIPA.DBO.FN_BLOQUEIO_SCANNER(<id>,'SAIDA')`
- Destinos especiais **BK** / **CLIA**: leituras/atualizações em `SGIPA..TB_CNTR_BL`, `OPERADOR..TB_GD_RESERVA`, `SGIPA..VW_AGG_AGENDAMENTOS`
- Avarias: `VW_EGMCI_AVARIAS_TTL`, `sgipa..dte_tb_avarias`, `TB_AVARIAS_CONTEINER`, `TB_ETAPAS_AVARIA`
- Histórico shifting na aba histórico: montagem dinâmica conforme `MskCntr.Tag` (`R`, `A`, default `I`) com `OPERADOR..TB_HIST_SHIFTING`

### Persistência da movimentação

- **Procedure:** `[Operador].[dbo].[pcMovimentaCntr]` com parâmetros: `Autonum`, `Sigla`, `Gmci`, `Sistema`, `YardOrigem`, `YardDestino`, `Tara`, lacres, `Camera_Posicionamento`, `Usuario`, `Motivo`
- **INSERT** explícitos na tela: divergência de lacres → `TB_AVARIAS_CONTEINER`; etapa de avarias → `TB_ETAPAS_AVARIA`; inclusão manual de avarias; exclusão de avarias
- **Segregação IMO:** se `Verifica_Regras_Seg_Imo_Delta` retorna texto, executa **nova** `pcMovimentaCntr` **invertendo** origem/destino (efeito de **rollback** da movimentação) e exibe mensagem

### Trecho com sintaxe Oracle (risco)

Em `Busca_Dados`, para `Sistema = "I"`, há SQL com `||`, `to_char`, tabelas `tb_agendamento_posicao` — **concatenação Oracle**. **Hipótese:** linked server ou ambiente híbrido; em SQL Server puro pode falhar. Deve ser validado em ambiente real.

---

## Dependências técnicas globais

- `Col_Patio` — filtro de pátio; regra **`Col_Patio = 1 Or Col_Patio = 7`** → `PATIO IN (1,7)` em várias queries
- `Col_Patio = 3` — ramificações para `tb_yard` (`yard_eq`) e `SUBSTRING` de pilha (posições 3–4 vs 2–3)
- `Cod_Usuario` — auditoria / permissões / inserts
- `ConnectionString` / `ConnectionStringAdo` — ADO
- Formulários modais: `FrmVeiculoServ`, `Frmmensagem`, `frmMensagem2`

---

## Atalhos e automatismos

- Enter emula Tab (`Form_KeyPress`)
- Resolução por **final** do ID no campo `CNTR` após digitar sem underscore
- Ajuste automático de **motivo** (`DataCombo2.BoundText`) conforme combinação yard atual / destino / SAÍDA

---

## Pontos críticos

| Tipo | Descrição |
|------|-----------|
| SQL injection / concatenação | Quase todos os SQL são montados por concatenação de strings |
| Transação | `Trata_Erro_Atualizacao` referencia `db.RollbackTrans`, mas não há evidência clara de `BeginTrans` no trecho analisado de `Atualiza_Posicao` — **hipótese:** transação dentro da SP ou código incompleto |
| `Form_Unload` | `Unload Me` duplicado |
| Complexidade `Carrega_Grid2` | Lógica extensa de LIKE / pilhas / UNION — alto risco de regressão na migração |
| Segregação IMO | Segunda chamada à SP para estorno — comportamento deve ser reproduzido ou substituído por transação única no backend |
| Mix Oracle/SQL Server | Query de agendamento de posicionamento |

---

## Inventário resumido de objetos de banco (evidência no código)

**Schemas principais:** `operador`, `sgipa`, `redex` (consultas condicionais).

**Tabelas / views / rotinas (referenciadas):**  
`TB_PATIOS`, `TB_cad_motivo`, `TB_CAMERAS`, `TB_YARD`, `TB_HIST_SHIFTING`, `TB_CNTR_BL`, `TB_GD_RESERVA`, `TB_ARMAZENS_IPA`, `TB_BL`, `TB_AMR_CNTR_BL`, `TB_REGISTRO_SAIDA_CNTR`, `TB_ORDEM_CARREGAMENTO`, `TB_AVARIAS_CONTEINER`, `TB_ETAPAS_AVARIA`, `dte_tb_avarias`, `dte_tb_tipos_conteiner`, views `VW_PROX_MOV_*`, `VW_EGMCI_AVARIAS_TTL`, `VW_SIGVIG_CNTR_ESTOQUE`, funções `fnInvent_Sistemas_*`, `fnMovimentaCntr`, `FN_BLOQUEIO_SCANNER`, SP `pcMovimentaCntr`, e objetos SIGVIG (`TB_RET_UN_SIGVIG`, etc.).

---

## Hipóteses explícitas

1. O nome amigável **“Movimentação Container: Pátio”** corresponde ao executável **Coletor Pátio** e ao form `FrmMovimentacaoS.frm`, embora o menu continue chamando `FrmMovimentacao.Show`.
2. O `.vbp` `Coletor.vbp` na pasta pode não refletir todos os forms usados em builds reais.
3. A query Oracle embutida em `Busca_Dados` pode depender de linked server ou estar obsoleta — requer validação em DB.
