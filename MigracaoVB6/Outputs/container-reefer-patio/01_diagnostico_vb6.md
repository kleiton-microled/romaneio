# Diagnostico VB6 — Container REEFER (Patio)

## Identificacao da tela (evidencia no codigo)

- **Nome funcional solicitado**: "Container - REEFER: Patio" (equivalente operacional ao botao do coletor de patio).
- **Caption do botao no menu (Coletor Patio)**: `Principal.Command1(Index=7)` com texto legivel como **"Conteiner - REEFER"** (acentos corrompidos no `.frm`; encoding ANSI misturado e comum no legado).
- **Titulo do produto / contexto Patio**: `ColetorPatioSmall.vbp` — `Title="Coletor Pátio"` e `ExeName32="ColetorPatio19032026.exe"`.
- **Nome interno VB do form**: `Reefer` (`Attribute VB_Name = "Reefer"`), implementado em **dois arquivos fisicos distintos**:
  - **Patio (este estudo)**: `ProjetoReferencia/Coletor/ReeferS.frm` (+ `ReeferS.frx`), referenciado em `ColetorPatioSmall.vbp` como `Form=ReeferS.frm`.
  - **Coletor armazem (referencia cruzada)**: `ProjetoReferencia/Coletor/Reefer.frm` (+ `Reefer.frx`); **nao** consta nas primeiras linhas de `Coletor.vbp` analisadas; existe no diretorio e compartilha o mesmo `VB_Name` com logica muito similar (divergencias documentadas abaixo).

### Entrada pelo menu (Patio)

| Arquivo | Trecho |
|---------|--------|
| `ProjetoReferencia/Coletor/PrincipalS.frm` | `Command1_Click` — `Case 7: Reefer.Show 1` — `Principal.Show` |

## Arquivos localizados

### Nucleo (Patio)

| Tipo | Arquivo |
|------|---------|
| `.frm` | `ProjetoReferencia/Coletor/ReeferS.frm` |
| `.frx` | `ProjetoReferencia/Coletor/ReeferS.frx` |
| `.vbp` | `ProjetoReferencia/Coletor/ColetorPatioSmall.vbp` |

### Variante Armazem (comparacao tecnica)

| Tipo | Arquivo |
|------|---------|
| `.frm` | `ProjetoReferencia/Coletor/Reefer.frm` |
| `.frx` | `ProjetoReferencia/Coletor/Reefer.frx` |

### Dependencias compartilhadas

| Tipo | Arquivo | Uso na tela |
|------|---------|-------------|
| `.bas` | `ProjetoReferencia/Coletor/Microled.bas` | `db`, `Sql`/`SQl`, `Banco_Operador`, `ConnectionStringAdo`, `CONVERTE_SQL`, `Nnull`, `Consiste_Temp`, `Consiste_Umidade`, `Consiste_Ventilacao` |
| `.bas` | `ProjetoReferencia/Coletor/mdlColetor.bas` | `Col_Patio`, `Col_Mensagem`, `Col_Cancelado`, `Cod_Usuario` (**Hipotese**: `Cod_Usuario` definido em outro modulo global — uso no INSERT sem declaracao neste `.frm`) |
| `.frm` | `ProjetoReferencia/Coletor/FrmmensagemS.frm` (Patio) / `Frmmensagem.frm` (Arm) | `Frmmensagem` / `FrmmensagemS` — mensagens modais (`Label1.Caption` + `Show vbModal`) |
| `.frm` | `ProjetoReferencia/Coletor/frmMensagem2S.frm` / `frmMensagem2.frm` | Confirmacao divergencia temperatura (`Col_Mensagem`, `frmMensagem2.Show 1`, `Col_Cancelado`) |

> **Hipotese**: No projeto Patio (`ColetorPatioSmall.vbp`), as referencias de form de mensagem podem ser as variantes `*S`; o codigo-fonte em `ReeferS.frm` referencia `Frmmensagem` e `frmMensagem2` **sem sufixo S** — confirmar em build real qual `.frm` e compilado no mesmo projeto (nomes podem resolver por compatibilidade de referencia ou copia).

### OCXs / controles ActiveX (declarativos no `.frm`)

- `MSMASK32.OCX` — `MaskEdBox` (`CNTR`, `MskCntr`)
- `TABCTL32.OCX` — `SSTab` (abas Monitoramento / Previsao / Un. Desligadas)
- `MSADODC.OCX` — `Adodc6`, `AdodcEntradas`, `AdodcSaidas`, `AdodcDesligados` (fonte de dados para grids; `Visible=False`)
- `MSHFLXGD.OCX` — `MSHFlexGrid` (`Grid2`, `GridEntradas`, `GridSaidas`, `GridDesligados`)

## Controles principais (ReeferS — aba Monitoramento)

| Nome (VB) | Tipo | Observacao |
|-----------|------|------------|
| `SSTab1` | Tab | 3 abas |
| `CNTR` | MaskEdBox | Final do conteiner (`###-#`) |
| `MskCntr` | MaskEdBox | ID completo (`>????######-#`) |
| `txtSetPoint`, `txtUmidadeSet`, `txtVentilacaoSet` | TextBox | Somente leitura (`Locked`); setpoints |
| `txtTempSup`, `txtTempRet`, `txtEscala`, `txtUmidade`, `txtVentilacao` | TextBox | Entrada de leituras |
| `txtGateIN` | TextBox | Leitura Gate IN (locked) |
| `txtMIN`, `txtMAX` | TextBox | **Somente ReeferS** — limites de temperatura de `TB_CNTR_BL` |
| `Command1` | Button | `SALVAR` |
| `Command2` | Button | `PLUG OFF` |
| `Command4` | Button | `X` fecha form |
| `Grid2` | MSHFlexGrid | Historico de monitoramento |
| `lblIPA`, `lblRdx`, `lblOp` | Label | Hidden — armazenam `AUTONUM` por sistema (I/R/O) |
| `lblReeferDesligado`, `lblPlugOFF` | Label | Alerta visual plug desligado |

## Abas 2 e 3 (Previsao / Un. Desligadas)

- `GridEntradas`, `AdodcEntradas`, `lblTotEnt`, `Label1(1)` — "UNIDADES PREVISTAS - ENTRADA PATIO"
- `GridSaidas`, `AdodcSaidas`, `lblTotSai`, `Check1..3`, `Shape1` — "UNIDADES EM ESTOQUE" + filtros
- `GridDesligados`, `AdodcDesligados`, `lblTotDesl`, `Label13` — unidades desligadas

## Eventos mapeados (ReeferS)

| Origem | Evento | Comportamento resumido |
|--------|--------|------------------------|
| `CNTR` | `Change` | Se mascara sem `_`, chama `Busca_Cntr` |
| `MskCntr` | `LostFocus` | `Busca_Dados` |
| `Command1` | `Click` | `Valida_Dados`; INSERT `TB_MONITORING`; limpa `DT_PLUG_OFF` em `TB_CNTR_BL` e/ou `REDEX..TB_PATIO`; recarrega grid |
| `Command2` | `Click` | Se nao esta "reefer desligado" visivel e `lblIPA>0`, grava `DT_PLUG_OFF=GETDATE()` em `SGIPA..TB_CNTR_BL` e reconsulta |
| `Command4` | `Click` | `Unload Me` |
| `SSTab1` | `Click` | Tab 1: carrega grids entradas/saidas; Tab 2: `Carrega_Grid_Desligados` |
| `Check1..3` | `Click` | Reexecuta `Carrega_Grid_Saidas_Previstas` |
| `Form` | `KeyDown` | `ESC` fecha; **bug**: usa `KeyAscii` em `KeyDown` (parametro correto e `KeyCode`) — Enter->Tab pode nao funcionar como pretendido |
| `txtEscala` | `LostFocus` | Normaliza `F`/`C` |
| `txtTempSup`, `txtTempRet`, `txtUmidade`, `txtVentilacao` | `LostFocus` | `UCase` + `Consiste_*` |

## Funcoes / subs no `.frm`

| Nome | Finalidade |
|------|------------|
| `Busca_Cntr` | Resolve conteiner pelo final (5 chars) em `VW_INVENT_SISTEMAS` com patio e flags reefer |
| `Busca_Dados` | Carrega dados do conteiner, setpoints, plug off, labels de sistema, grid historico |
| `Carrega_Grid1` | Historico `TB_MONITORING` por autonum IPA/RDX/OP |
| `Carrega_Grid_Entradas_Previstas` | Previsao entrada patio |
| `Carrega_Grid_Saidas_Previstas` | Estoque patio + destaque linhas com `dt_plug_off` |
| `Carrega_Grid_Desligados` | Lista com `dt_plug_off is not null` |
| `Limpa_Dados` / `Limpa_Monitora` | Reset de campos (**`Limpa_Monitora` nao referenciado** no trecho analisado — possivel codigo morto) |
| `Valida_Dados` | Obrigatoriedades + alerta divergencia supply vs setpoint (>2) |

## SQL e persistencia (extraido do legado)

### Consultas — inventario / busca conteiner

**ReeferS — `Busca_Dados` (filtro patio unificado 1 e 7)**

```sql
SELECT autonum, data_ent_temp, sistema, temperature, scale, umidade, ventilacao,
       flag_reefer, flag_reefer_desligado
FROM <Banco_Operador>VW_INVENT_SISTEMAS
WHERE PATIO <dynamic: IN (1,7) se Col_Patio em (1,7); senao = Col_Patio>
  AND id_conteiner = '<MskCntr>'
  AND FLAG_REEFER='Y' AND FLAG_REEFER_DESLIGADO='N'
```

**Reefer (arm) — `Busca_Dados`**: mesmo SELECT com `patio = Col_Patio` apenas; filtros `FLAG_REEFER` **comentados** na primeira query (comportamento diferente).

**`Busca_Cntr` — contagem / resolucao por final**

```sql
SELECT COUNT(1)
FROM <Banco_Operador>VW_INVENT_SISTEMAS
WHERE PATIO <mesma regra Patio 1/7>
  AND SUBSTR(ID_CONTEINER,8,5)='<CNTR>'
  AND FLAG_REEFER='Y' AND FLAG_REEFER_DESLIGADO='N'
```

> **Regra tecnica**: `SUBSTR(ID_CONTEINER,8,5)` pressupoe formato de `ID_CONTEINER` com sufixo de 5 caracteres na posicao 8 (**Hipotese**: prefixo de 7 caracteres + 5 do final).

### Consultas — historico monitoramento (`Carrega_Grid1`)

```sql
SELECT
  LTRIM(LEFT(CONVERT(CHAR,A.DT_MONITORING,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,A.DT_MONITORING,108),5)),
  A.TEMP_SUP + ' ' + A.ESCALA,
  A.TEMP_RET + ' ' + A.ESCALA,
  A.UMIDADE,
  A.VENTILACAO,
  B.USUARIO,
  A.AUTONUM
FROM OPERADOR..TB_MONITORING A
LEFT JOIN SGIPA..TB_CAD_USUARIOS B ON A.USUARIO = B.AUTONUM
WHERE (AUTONUM_CNTR_IPA = <lblIPA> OU AUTONUM_CNTR_RDX = <lblRdx> OU AUTONUM_CNTR_OP = <lblOp>)
ORDER BY A.AUTONUM DESC
```

### Insert — gravacao leitura (`Command1_Click`)

```sql
INSERT INTO OPERADOR..TB_MONITORING (
  DT_MONITORING, ID_CONTEINER,
  AUTONUM_CNTR_IPA, AUTONUM_CNTR_RDX, AUTONUM_CNTR_OP,
  TEMP_SETPOINT, UMIDADE_SETPOINT, VENTILACAO_SETPOINT,
  TEMP_SUP, TEMP_RET, VENTILACAO, UMIDADE, ESCALA,
  PATIO, USUARIO
) VALUES (
  GETDATE(), '<MskCntr>',
  <lblIPA>, <lblRdx>, <lblOp>,
  '<txtSetPoint>', '<txtUmidadeSet>', '<txtVentilacaoSet>',
  '<txtTempSup>', '<txtTempRet>', '<txtVentilacao>', '<txtUmidade>', '<txtEscala>',
  <IIf(Col_Patio=7,1,Col_Patio)>,
  <Cod_Usuario>
)
```

### Updates pos-gravacao (ReeferS)

```sql
UPDATE SGIPA..TB_CNTR_BL SET DT_PLUG_OFF=NULL WHERE AUTONUM=<lblIPA>  -- se lblIPA>0
UPDATE REDEX..TB_PATIO SET DT_PLUG_OFF=NULL WHERE AUTONUM_PATIO=<lblRdx> -- se lblRdx>0
```

**Reefer.frm (arm)**: apenas o primeiro UPDATE (IPA); **nao** atualiza `REDEX..TB_PATIO`.

### Update — PLUG OFF (`Command2_Click`)

```sql
UPDATE SGIPA..TB_CNTR_BL SET DT_PLUG_OFF=GETDATE() WHERE AUTONUM=<lblIPA>
```
- Condicao de tela: `lblReeferDesligado.Visible = False` e `lblIPA > 0`.
- **Hipotese**: Unidades sistema `R` ou `O` nao recebem plug off por este botao neste form (sem branch `lblRdx`/`lblOp`).

### Grids Previsao / Estoque / Desligados

SQL extenso em `Carrega_Grid_Entradas_Previstas`, `Carrega_Grid_Saidas_Previstas`, `Carrega_Grid_Desligados` — ver `02_mapeamento_funcional.md` e inventario em `04_regras_negocio.md`.

**Objetos envolvidos (nomes literais no SQL)**:

- `SGIPA..TB_CNTR_BL`, `SGIPA..DTE_TB_VIAGENS`, `SGIPA..DTE_TB_ARMAZENS`, `SGIPA..TB_TTIMEUNT`
- `OPERADOR..TB_MONITORING`, `OPERADOR..TB_GD_RESERVA`
- `REDEX..TB_PATIO` (somente ReeferS no fluxo de salvar)
- View dinamica: `<Banco_Operador>VW_INVENT_SISTEMAS`

## Transacoes

- Nao ha `BeginTrans` / `CommitTrans` / `RollbackTrans` no codigo deste `.frm`.
- Cada `db.Execute` e autonomo (**risco de consistencia parcial** se segunda instrucao falhar apos primeira).

## Pontos criticos / riscos tecnicos

1. **SQL injection / concatenacao**: todos os SQLs concatenam valores de tela e labels.
2. **`Form_KeyDown`**: uso de `KeyAscii` incorreto para o evento.
3. **`Busca_Dados` — ramo Else de ventilacao vazio**: atribui `"CLOSED"` a `txtUmidadeSet` em vez de `txtVentilacaoSet` (Reefer e ReeferS) — **bug evidente** no codigo.
4. **`Consiste_Temp` (Microled.bas)**: primeira linha usa `Replace(Temperatura,...)` mas em seguida sobrescreve com `Replace(Temperatura, "+", "")` sem encadear o resultado da virgula — possivel inconsistencia menor.
5. **CASE `WHEN NULL` em SQL Server**: expressao `CASE V.DT_ATRACACAO WHEN NULL THEN ...` **nao** captura NULL da forma idiomática (`IS NULL`); comportamento depende do motor/compatibilidade — validar resultado real.
6. **Procedimentos armazenados**: nenhum referenciado neste form; apenas SQL ad-hoc.

## Checklist interno (Agente 01)

- [x] Tela localizada no menu/fluxo (`PrincipalS` Case 7)
- [x] Form principal identificado (`Reefer` em `ReeferS.frm` para Patio)
- [x] Recursos `.frx` identificados
- [x] Dependencias `.bas` mapeadas
- [x] SQLs extraidos e classificados
- [x] Atalhos de teclado listados (ESC; Enter->Tab **questionavel** por bug)
- [x] Eventos por controle mapeados
- [x] Inconsistencias/risco identificados
