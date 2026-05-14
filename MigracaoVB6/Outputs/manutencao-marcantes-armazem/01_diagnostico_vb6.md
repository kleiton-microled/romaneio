# Diagnostico VB6 - Manutencao Marcantes (Armazem / divergencia)

## Identificacao da tela (evidencia no codigo)

- **Nome no menu (Coletor Armazem)**: botao `Principal.Command1` com `Index = 10`, `Caption` legivel como **"Manutencao - Marcantes"** (caracteres acentuados corrompidos no arquivo `.frm`, padrao ANSI/UTF-8 misto).
- **Titulo no form**: `Label1(0).Caption` = **"MANUTENCAO- MARCANTES"** (mesma observacao de encoding; no arquivo aparecem caracteres substitutos).
- **Nome interno VB**: `AdminMarcante` (`Attribute VB_Name = "AdminMarcante"`).
- **Arquivo principal**: `ProjetoReferencia/Coletor/AdminMarcante.frm`.
- **Projeto**: `Coletor.vbp` (executavel `ColetorArm.exe`), linha `Form=AdminMarcante.frm`.
- **Entrada pelo menu**: `Principal.frm` -> `Command1_Click` -> `Case 10` -> `AdminMarcante.Show 1`.
- **Relacao com o nome solicitado ("Manutencao - Marcantes Armazem")**:
  - O texto **"Armazem"** nao aparece literalmente no `Caption` do botao nem no titulo do form nos arquivos analisados.
  - A tela lista **ARMAZEM** e **POSICAO** a partir de `OPERADOR..VW_INVENT_ARMAZEM` e filtra cargas em contexto de **divergencia** (`SGIPA..VW_WMS_DIVERGENCIA_QTDE`), o que caracteriza manutencao operacional de volumes de marcantes ligada ao inventario/armazem.

## Arquivos localizados

### Nucleo

| Tipo | Arquivo |
|------|---------|
| `.frm` | `ProjetoReferencia/Coletor/AdminMarcante.frm` |
| `.frx` | **Nao existe** (`AdminMarcante.frx` ausente na pasta Coletor). |

### Entrada e permissao

| Tipo | Arquivo / trecho |
|------|-------------------|
| `.frm` | `ProjetoReferencia/Coletor/Principal.frm` — `Case 10: AdminMarcante.Show 1` |
| `.frm` | `Principal.Form_Activate` — `Valida_Acesso_Botao(Col_Login, "COMMAND1(10)")` desabilita o botao se falso; visibilidade `Command1(10).Visible = True` no ramo `Sistema_Coletor = "IPA"` (e `False` no ramo `REDEX`). |

### Dependencias compartilhadas (referenciadas pelo form)

| Tipo | Arquivo | Uso na tela |
|------|---------|-------------|
| `.bas` | `ProjetoReferencia/Coletor/Microled.bas` (referenciado no `Coletor.vbp` como `Microled`) | Conexao global `db`, string `Sql`, recordset `tb1`, funcao `Nnull` |
| `.bas` | `ProjetoReferencia/Coletor/mdlColetor.bas` | `Col_Mensagem` antes de `frmMensagem2.Show vbModal` |
| `.frm` | `ProjetoReferencia/Coletor/frmMensagem2.frm` | `frmMensagem2.Show vbModal` apos setar `Col_Mensagem`. |

> Observacao: existem outros forms de mensagem na mesma pasta (`frmMensagem2S.frm`, `frmMensagem2SA.frm`, etc.); a chamada desta tela e especificamente `frmMensagem2`.

## Controles do form (inventario)

| Nome (VB) | Tipo | Observacao |
|-----------|------|------------|
| `Label1(0)` | Label | Titulo da tela |
| `TXTLOTE` | TextBox | Entrada de lote; `MaxLength = 7` |
| `cmdFiltra` | CommandButton | `[F4]-FILTRA` |
| `TXTAUTONUMCS` | TextBox | **Hidden** (`Visible = False`) — armazena `AUTONUM_CS` |
| `txtEmbalagem` | TextBox | `Locked = True` — preenchido no filtro |
| `txtQtde` | TextBox | `Locked = True` — quantidade de referencia do filtro (`QUANTIDADE_REAL` da view) |
| `Label3` | Label | Texto "MARCANTES" |
| `lv` | `MSComctlLib.ListView` | Grid de marcantes (colunas configuradas no `Form_Load`) |
| `TXTTOT` | TextBox | `Locked = True` — soma das quantidades listadas; cor dinamica |
| `TXTMARCANTE` | TextBox | Marcante selecionado (preenchido pelo clique na lista) |
| `TXTQTDEN` | TextBox | Nova quantidade informada pelo usuario |
| `cmdSalvar` | CommandButton | `[F5]-SALVA` |
| `cmdEscape` | CommandButton | `[ESC]-VOLTAR` |
| `Label1(1..6)` | Label | Rotulos de colunas/campos (LOTE, QTDE, EMBALAGEM, TTL, MARCANTE, QTDE nova) |

## Eventos mapeados

| Origem | Evento | Comportamento |
|--------|--------|---------------|
| Form | `Form_Load` | Posiciona form (`Top=0`, `Left=0`); chama `Limpa`; define colunas do `lv`. |
| Form | `Form_KeyDown` | `ESC` -> `Unload Me`; `F4` -> `cmdFiltra_Click`; `F5` -> `cmdSalvar_Click`. |
| Form | `Form_Unload` | `Unload Me` (recursivo/redundante; efeito pratico de encerramento). |
| `cmdEscape` | `Click` | `Unload Me`. |
| `cmdFiltra` | `Click` | Valida lote; consulta `VW_WMS_DIVERGENCIA_QTDE`; preenche dados e `Carrega_Lv`. |
| `cmdSalvar` | `Click` | Valida marcante e quantidade; `UPDATE TB_MARCANTES`; limpa campos; recarrega lista; tenta reposicionar selecao na lista. |
| `lv` | `Click` | Copia marcante da linha selecionada para `TXTMARCANTE`. |

**Eventos ausentes no codigo analisado**: `Change`/`LostFocus` em lote ou quantidade; `Validate`/`QueryUnload`; tratamento de erro em torno de `db.Execute` (sem `On Error` no `cmdSalvar`).

## Funcoes e procedimentos (no proprio `.frm`)

| Nome | Finalidade |
|------|------------|
| `cmdFiltra_Click` | Filtra divergencia por lote e carrega lista. |
| `cmdSalvar_Click` | Atualiza `VOLUMES` do marcante e refresca grid. |
| `Carrega_Lv` | Monta SQL, abre `tb1`, popula `lv`, calcula `TXTTOT`, ajusta cor do total. |
| `Limpa` | Limpa campos de detalhe e `txtAutonumCS`; chama `Carrega_Lv`. |

## SQL e persistencia (texto extraido do legado)

### 1) Filtro inicial (`cmdFiltra_Click`)

```sql
SELECT * FROM SGIPA..VW_WMS_DIVERGENCIA_QTDE WHERE LOTE = <Val(TXTLOTE)>
```

- Campos lidos no recordset (apos `Open`): `embalagem`, `QUANTIDADE_REAL`, `AUTONUM_CS`.
- Atribuicoes: `txtEmbalagem` <- `Nnull(tb1!embalagem, 1)`; `TXTQTDE` <- `Nnull(tb1!QUANTIDADE_REAL, 0)`; `txtAutonumCS` <- `tb1!AUTONUM_CS`; depois `Carrega_Lv`.

### 2) Lista de marcantes (`Carrega_Lv`) — SQL ativo

```sql
SELECT C.ITEM,
       I.DESCR_ARMAZEM,
       I.POSICAO,
       M.VOLUMES AS QTDE,
       M.AUTONUM AS MARCANTE,
       M.AUTONUM_CARGA AS AUTONUMCS
FROM SGIPA..TB_MARCANTES M
INNER JOIN SGIPA..TB_CARGA_SOLTA C ON M.AUTONUM_CARGA = C.AUTONUM
LEFT JOIN (
    SELECT DESCR_ARMAZEM, POSICAO, MARCANTE
    FROM OPERADOR..VW_INVENT_ARMAZEM I
    WHERE AUTONUMCS = <txtAutonumCS>
      AND QTDE > 0
) I ON M.AUTONUM = I.MARCANTE
WHERE C.AUTONUM = <txtAutonumCS>
  AND M.DT_ASSOCIACAO IS NOT NULL
ORDER BY M.AUTONUM
```

- Existe bloco **comentado** com variante usando `C.BL` e join direto em `VW_INVENT_ARMAZEM` — **nao executado**; serve apenas como evidencia de evolucao ou tentativa anterior.

### 3) Gravacao (`cmdSalvar_Click`)

```sql
UPDATE SGIPA..TB_MARCANTES
SET VOLUMES = <Val(TXTQTDEN)>
WHERE autonum = <Val(txtMarcante)>
```

- Execucao: `db.Execute Sql` **sem** `BeginTrans` / `CommitTrans` / `RollbackTrans` no trecho analisado.

## Objetos de banco citados (somente os que aparecem no SQL ativo)

| Objeto | Papel |
|--------|-------|
| `SGIPA..VW_WMS_DIVERGENCIA_QTDE` | Filtro por `LOTE`; fornece referencia de quantidade e `AUTONUM_CS`. |
| `SGIPA..TB_MARCANTES` | Leitura de `VOLUMES`, `AUTONUM`, `AUTONUM_CARGA`, `DT_ASSOCIACAO`; **update** de `VOLUMES`. |
| `SGIPA..TB_CARGA_SOLTA` | Join por `AUTONUM`; filtro `C.AUTONUM = AUTONUM_CS`. |
| `OPERADOR..VW_INVENT_ARMAZEM` | Subconsulta para `DESCR_ARMAZEM`, `POSICAO`, `MARCANTE` com `QTDE > 0`. |

**Campos lidos (lista)**: `ITEM`, `DESCR_ARMAZEM`, `POSICAO`, `QTDE` (alias de `M.VOLUMES`), `MARCANTE`, `AUTONUMCS`.

**Campos gravados**: `TB_MARCANTES.VOLUMES` para o `AUTONUM` selecionado.

## Atalhos de teclado

| Tecla | Acao |
|-------|------|
| F4 | Filtrar (`cmdFiltra_Click`) |
| F5 | Salvar (`cmdSalvar_Click`) |
| ESC | Sair (`Unload Me`) — tambem via `cmdEscape` |

`KeyPreview = True` no form habilita `Form_KeyDown`.

## Dependencias globais e mensagens

- `db.Execute` — conexao ADO global (`Microled.bas`).
- `Sql`, `tb1` — globais tipicas do padrao Microled.
- `Nnull(campo, tipo)` — normalizacao de nulos.
- Mensagens via `Col_Mensagem` + `frmMensagem2` modal:
  - "Informe o lote"
  - "Lote nao encontrado como divergente" (encoding original corrompido no `.frm`)
  - "Selecione o marcante"
  - "Informe a quantidade"

## Pontos criticos / riscos tecnicos (evidencia + classificacao)

1. **SQL concatenado** com valores de tela (`Val(...)`) — **Regra tecnica / risco**: injecao e falha de tipo se entrada nao numerica (mitigado parcialmente por `Val()` retornar 0).
2. **Ausencia de transacao** na gravacao — **Regra tecnica**: operacao e um unico `UPDATE`; risco de estado inconsistente e baixo em relacao a multiplas tabelas, mas sem tratamento de erro.
3. **`cmdSalvar_Click`**: `LvAnt = Me.LV.SelectedItem` e depois `Me.LV.SelectedItem = LvAnt` — **Risco tecnico**: tipos `ListItem` vs inteiro; codigo pode depender de comportamento implicito do OCX ou de propriedade padrao; `On Error Resume Next` mascara falha de reposicionamento.
4. **`lv_Click`**: `ListItems(Val(Me.LV.SelectedItem))` — **Risco tecnico / Hipotese**: semantica depende do tipo retornado por `SelectedItem` e conversao `Val`; pode ser fragil entre versoes.
5. **`Limpa` + `Carrega_Lv`**: `Limpa` zera `txtAutonumCS` mas **nao** limpa `TXTLOTE`; se `TXTLOTE` nao estiver vazio, `Carrega_Lv` monta SQL com `AUTONUMCS` vazio — **Risco / possivel SQL invalido ou resultado inesperado**.
6. **Definicao da view `VW_WMS_DIVERGENCIA_QTDE`** — **Lacuna**: nao esta no repositorio; colunas e significado de `LOTE` vs `BL` dependem do banco.

## Checklist interno (Agente 01)

- [x] Tela localizada no fluxo de entrada (`Principal`, `Case 10`).
- [x] Form principal identificado (`AdminMarcante.frm`).
- [x] Recursos `.frx` — nao aplicavel (arquivo ausente).
- [x] Dependencias `.bas` mapeadas (`Microled`, `mdlColetor`).
- [x] SQLs extraidos e classificados (consulta + update).
- [x] Atalhos de teclado listados.
- [x] Eventos por controle mapeados.
- [x] Inconsistencias/risco identificados.
