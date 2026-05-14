# Regras de negocio, validacao, interface e tecnica — Container REEFER (Patio)

Legenda de classificacao:

- **RN** = Regra de negocio
- **RI** = Regra de interface
- **RT** = Regra tecnica
- **RV** = Regra de validacao
- **HP** = Hipotese (sem evidencia completa no codigo analisado)

## Catalogo de regras

| ID | Classificacao | Regra | Evidencia |
|----|---------------|--------|-----------|
| R01 | RN | Somente conteneres **reefer** elegiveis (`FLAG_REEFER='Y'`) e **nao** `FLAG_REEFER_DESLIGADO='N'` entram na busca por final (ReeferS). | `Busca_Cntr` SQL |
| R02 | HP | Na variante `Reefer.frm`, a primeira consulta de `Busca_Dados` **nao** aplica os filtros `FLAG_REEFER` na string ativa (linhas comentadas) — comportamento pode divergir do Patio. | `Reefer.frm` |
| R03 | RN | Patio do operador restringe consultas: em Patio, se `Col_Patio` for **1 ou 7**, usar **`PATIO IN (1,7)`** nas queries principais. | `ReeferS.frm` — `IIf(Col_Patio = 1 Or Col_Patio = 7, ...)` |
| R04 | RN | Resolver conteiner por **5 caracteres finais** (`SUBSTR(ID_CONTEINER,8,5)`); exige **match unico** para auto-preencher ID completo. | `Busca_Cntr` |
| R05 | RV | Se nenhum registro na view com ID completo: mensagem "Conteiner nao encontrado". | `Busca_Dados` |
| R06 | RV | Se `flag_reefer = N`: bloqueia com mensagem "nao e reefer". | `Busca_Dados` |
| R07 | RV | Se `flag_reefer_desligado = Y`: bloqueia "captado com carga DRY". | `Busca_Dados` |
| R08 | RN | Sistema de origem (`Sistema` = I/R/O) determina qual autonum preenche `lblIPA`, `lblRdx` ou `lblOp`. | `Busca_Dados` |
| R09 | RN | Para sistema **I** (IPA), exibir estado **plug off** a partir de `TB_CNTR_BL.DT_PLUG_OFF` e MIN/MAX (`tempmin`, `tempmax`) em ReeferS. | `Busca_Dados` + SQL dedicado |
| R10 | RI | Enquanto alerta "REEFER DESLIGADO" visivel, acao **PLUG OFF** nao deve rodar (`Command2_Click`). | `Command2_Click` |
| R11 | RN | **PLUG OFF** grava `DT_PLUG_OFF = GETDATE()` no BL IPA quando `lblIPA > 0`. | `Command2_Click` |
| R12 | HP | PLUG OFF **nao** trata `lblRdx` ou `lblOp` neste codigo — operador Redex/Outros pode precisar de outro fluxo. | Ausencia de branch |
| R13 | RV | Ao salvar, exige conteiner resolvido: `lblIPA`, `lblRdx` e `lblOp` nao podem ser todos zero. | `Valida_Dados` |
| R14 | RV | Campos obrigatorios na gravacao: `txtTempSup`, `txtTempRet`, `txtUmidade`, `txtVentilacao`, `txtEscala`. | `Valida_Dados` |
| R15 | RN | Se diferenca absoluta entre **supply** e **setpoint** > **2**, exibir confirmacao; usuario pode cancelar (`Col_Cancelado`). | `Valida_Dados` + `frmMensagem2` |
| R16 | RN | Gravacao insere linha em `TB_MONITORING` com timestamp `GETDATE()`, usuario `Cod_Usuario`, patio conforme `IIf(Col_Patio=7,1,Col_Patio)` (**mapeamento 7->1 na coluna PATIO**). | `Command1_Click` ReeferS |
| R17 | RN | Apos salvar, se IPA: `DT_PLUG_OFF` em `TB_CNTR_BL` **anulado** (NULL). | `Command1_Click` |
| R18 | RN | Apos salvar, se RDX: `DT_PLUG_OFF` em `REDEX..TB_PATIO` **anulado** (NULL) — **somente ReeferS**. | `Command1_Click` ReeferS |
| R19 | RT | Sem transacao explicita: INSERT e UPDATEs sequenciais independentes. | Codigo |
| R20 | RT | SQL montado por concatenacao — risco de injecao e quebra com caracteres especiais. | Padrao do legado |
| R21 | RI | `CNTR_Change` dispara busca quando mascara nao contem `_`. | `CNTR_Change` |
| R22 | RI | `MskCntr_LostFocus` sempre chama `Busca_Dados` (pode reconsultar ao tabular). | `MskCntr_LostFocus` |
| R23 | RT | `Consiste_Temp`: normaliza; se nao numerico, esvazia. | `Microled.bas` |
| R24 | RT | `Consiste_Umidade` / `Consiste_Ventilacao`: nao numerico vira `OFF` / `CLOSED`; zero tambem vira `OFF` / `CLOSED`. | `Microled.bas` |
| R25 | RV | `txtEscala` aceita apenas **F** ou **C** apos lost focus. | `txtEscala_LostFocus` |
| R26 | RT | **Bug**: ao carregar setpoints, se `Ventilacao` vazia, codigo atribui `"CLOSED"` a **`txtUmidadeSet`** em vez de `txtVentilacaoSet`. | `Busca_Dados` Reefer/ReeferS |
| R27 | RN | Historico (`Grid2`) filtra por **um** dos autonums IPA/RDX/OP conforme qual label foi preenchido. | `Carrega_Grid1` |
| R28 | RN | Aba Previsao — **Entradas previstas**: conteneres com `FLAG_TERMINAL=0`, `GMCI` nao nulo, `TEMP` nao vazio, `flag_desligado=0`, patio conforme R03. | `Carrega_Grid_Entradas_Previstas` |
| R29 | RN | Aba Previsao — **Estoque**: `FLAG_TERMINAL=1`, `FLAG_HISTORICO=0`, `DT_ENTRADA` nao nulo, `FLAG_DESLIGADO=0`, `TEMP` nao vazio, patio conforme R03; opcionalmente `dt_plug_off` preenchido; agendamento `GD.PERIODO_INICIAL`; posicionados `yard like CF% ou CONF%`. | `Carrega_Grid_Saidas_Previstas` |
| R30 | RI | Linhas com `dt_plug_off` recebem cor de fundo da celula de conteiner (`CellBackColor = Shape1.FillColor`). | Loop `GridSaidas` |
| R31 | RN | Aba **Un. Desligadas**: mesmo nucleo de estoque, porem **obrigatoriamente** `c.dt_plug_off is not null`, ordenado por `dt_plug_off desc`. | `Carrega_Grid_Desligados` |
| R32 | RT | SQL Server: uso problematico de `CASE col WHEN NULL` em coluna de atracacao — revisar semantica real. | SQL em `Carrega_Grid_Entradas_Previstas` |
| R33 | HP | `Label13.Caption` pode ser titulo dinamico em outro modulo nao analisado. | Designer vazio |

## Inventario de banco (consolidado)

### Tabelas

| Tabela | Leitura | Gravacao |
|--------|---------|----------|
| `OPERADOR..TB_MONITORING` | Sim (historico + joins min/max) | INSERT |
| `SGIPA..TB_CNTR_BL` | Sim (entradas/saidas/desligados + plug + min/max) | UPDATE `DT_PLUG_OFF` |
| `REDEX..TB_PATIO` | Nao diretamente na listagem deste form | UPDATE `DT_PLUG_OFF` (ReeferS pos-salvar) |
| `SGIPA..TB_CAD_USUARIOS` | Sim (login nome) | Nao |
| `SGIPA..DTE_TB_VIAGENS` | Sim | Nao |
| `SGIPA..DTE_TB_ARMAZENS` | Sim | Nao |
| `SGIPA..TB_TTIMEUNT` | Sim | Nao |
| `OPERADOR..TB_GD_RESERVA` | Sim | Nao |

### Views

| View | Uso |
|------|-----|
| `<Banco_Operador>VW_INVENT_SISTEMAS` | Localizar conteiner; flags reefer; sistema |

### Procedures / functions

- **Nenhuma** referenciada literalmente neste form.

### Campos lidos (amostra representativa)

- `VW_INVENT_SISTEMAS`: `autonum`, `data_ent_temp`, `sistema`, `temperature`, `scale`, `umidade`, `ventilacao`, `flag_reefer`, `flag_reefer_desligado`, `id_conteiner`, `patio`
- `TB_CNTR_BL`: `DT_PLUG_OFF`, `tempmin`, `tempmax`, `ID_CONTEINER`, `TAMANHO`, `TEMP`, `ESCALA`, `DT_ENTRADA`, `YARD`, `FLAG_TERMINAL`, `FLAG_HISTORICO`, `FLAG_DESLIGADO`, `GMCI`, `VIAGEM`, `AUTONUM_GD_RESERVA`, `AUTONUM` (join monitoring)
- `TB_MONITORING`: medidas, datas, usuario, chaves `AUTONUM_CNTR_*`

### Campos gravados

- `TB_MONITORING`: todos os campos do INSERT listado no diagnostico.
- `TB_CNTR_BL.DT_PLUG_OFF`: GETDATE() ou NULL.
- `TB_PATIO.DT_PLUG_OFF`: NULL (ReeferS).

### Joins e filtros relevantes

- Agregados `MIN(AUTONUM)` / `MAX(AUTONUM)` por `AUTONUM_CNTR_IPA` em `TB_MONITORING` para primeira e ultima leitura por conteiner IPA.
- `LEFT JOIN TB_GD_RESERVA` por `AUTONUM_GD_RESERVA`.

### Transacoes / concorrencia

- **Sem** transacao multi-passo.
- **Concorrencia**: duas estacoes podem gravar monitoramentos e updates de plug off em sequencia diferente — ultimo comando prevalece (**risco operacional**).

### Impacto reefer / energia / temperatura

- **Energia / status plug**: `DT_PLUG_OFF` alterado manualmente e revertido ao salvar leitura.
- **Temperatura**: historico de supply/retorno vs setpoint; alerta divergencia.
- **Hipotese**: Efeitos em **romaneio** ou **carregamento** seriam indiretos via integridade de `TB_CNTR_BL` — nao evidenciados neste form.

## Checklist interno (Agente 04)

- [x] Regras por acao
- [x] Validacoes com mensagens
- [x] Consistencia quantitativa (divergencia >2)
- [x] Transacao (ausencia)
- [x] Casos silenciosos / riscos
- [x] Hipoteses separadas
