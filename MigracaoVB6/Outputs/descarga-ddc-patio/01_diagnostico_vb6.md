# Diagnostico VB6 — Descarga DDC (Patio)

## Identificacao da tela (evidencia no codigo)

- **Nome funcional solicitado**: "Descarga DDC: Patio".
- **Entrada no menu do Coletor Patio**: `PrincipalS.frm`, `Command1` com `Caption = "Descarga DDC"` e `Index = 15`; no clique abre `FrmDescargaArmazem`.

```352:354:d:\git\kleiton\romaneio\ProjetoReferencia\Coletor\PrincipalS.frm
    Case 15
        FrmDescargaArmazem.Show 1
        Principal.Show
```

- **Projeto executavel Patio**: `ProjetoReferencia/Coletor/ColetorPatioSmall.vbp` — `Title="Coletor Pátio"` e inclusao explicita do form:

```35:35:d:\git\kleiton\romaneio\ProjetoReferencia\Coletor\ColetorPatioSmall.vbp
Form=FrmDescargaArmazem.frm
```

- **Nome interno VB do form**: `FrmDescargaArmazem` (`Attribute VB_Name = "FrmDescargaArmazem"`).
- **Titulo da janela (caption)**: `Caption = "Descarga Armazem"` (texto diferente do botao de menu "Descarga DDC").

```5:8:d:\git\kleiton\romaneio\ProjetoReferencia\Coletor\FrmDescargaArmazem.frm
Begin VB.Form FrmDescargaArmazem 
   BackColor       =   &H00BA9F8D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descarga Armazem"
```

- **Label superior no form**: `Label5.Caption = "CARREGAMENTO DDC"` (contradicao visual com "descarga"; registrar como esta no legado).

```534:537:d:\git\kleiton\romaneio\ProjetoReferencia\Coletor\FrmDescargaArmazem.frm
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "CARREGAMENTO DDC"
```

## Arquivos localizados

### Nucleo (Patio / DDC)

| Tipo | Arquivo |
|------|---------|
| `.frm` | `ProjetoReferencia/Coletor/FrmDescargaArmazem.frm` |
| `.frx` | `ProjetoReferencia/Coletor/FrmDescargaArmazem.frx` (recursos de binding/picture) |
| `.frm` | `ProjetoReferencia/Coletor/PrincipalS.frm` (menu que abre a tela) |
| `.vbp` | `ProjetoReferencia/Coletor/ColetorPatioSmall.vbp` |

### Referencia cruzada (nao e esta tela, mas usa nome parecido)

| Contexto | Evidencia |
|----------|-----------|
| "Descarga CD" / talie | `PrincipalS.frm` `Case 3` define `DD=True`, `CD=False` e abre `FrmTalieColetor` — **fluxo distinto** desta tela. |

### Dependencias compartilhadas

| Tipo | Arquivo | Uso na tela |
|------|---------|-------------|
| `.bas` | `ProjetoReferencia/Coletor/Microled.bas` | `Banco_Sgipa`, `Banco_Operador`, `ConnectionString`, `CONVERTE_SQL`, `PRSet`, `Executa`, `Cod_Usuario`, `Rst` (recordset global), `Banco_Em_Uso` |
| `.bas` | `ProjetoReferencia/Coletor/mdlColetor.bas` | Modulo de coletor (nao referenciado diretamente neste `.frm`; pode existir acoplamento global em runtime) |

### OCXs / controles ActiveX

- `MSAdoDc.ocx` — `Adodc1`, `Adodc2`, `Adodc3`, `Adodc4` (`Visible=False`).
- `MSDatLst.Ocx` — `DataCombo` (`cbPlaca`, `cbLote`, `DcEmbalagem`).
- `MSDatGrd.ocx` — `DataGrid` (`Grid1`).

## Controles principais (declarativos no `.frm`)

| Nome (VB) | Tipo | Observacao |
|-----------|------|------------|
| `Label5` | Label | Texto "CARREGAMENTO DDC" |
| `cbPlaca` | DataCombo | Lista placas + ID conteiner (`Adodc4`) |
| `cbLote` | DataCombo | Lotes (`Adodc3`) |
| `Command1` | CommandButton | `F4 - Carregar` |
| `Command2` | CommandButton | `F2 - Limpar` |
| `Frame1` | Frame | Area de itens (inicia `Enabled=False`) |
| `Grid1` | DataGrid | Itens de descarga (`Adodc1`) |
| `Command5` | CommandButton | `F3 - Novo` |
| `Command4` | CommandButton | `F6 - Excluir` |
| `Command3` | CommandButton | `F5 - Salvar` |
| `Option1` / `Option2` | OptionButton | `SIM` / `NAO` para "Entrega Finalizada ?" |
| `TxtQuantidade` | TextBox | `MaxLength=5` |
| `DcEmbalagem` | DataCombo | Embalagens (`Adodc2`) |
| `TxtGate`, `TxtOC`, `TxtLote`, `TxtAutonum`, `TxtCntr` | TextBox | Campos auxiliares; varios `Visible=False` |

## Eventos e fluxos mapeados

| Origem | Evento | Comportamento resumido |
|--------|--------|------------------------|
| `Form` | `Load` | `Carrega_Placa`; carrega `dte_tb_embalagens` em `Adodc2`; `TxtCntr=""` |
| `Form` | `KeyDown` | `ESC` fecha; `F4` `CARREGAR`; `F2` `Limpar`; `F3`/`F5`/`F6` exigem `Frame1.Enabled` |
| `Form` | `Unload` | Chama `Unload Me` de dentro do proprio `Unload` (**risco de recursao** — ver lacunas) |
| `cbPlaca` | `LostFocus` | Limpa estado; resolve `TxtGate` via `tb_gate_new`; carrega lotes ativos em `Adodc3` |
| `Command1` | `Click` | `CARREGAR` |
| `Command2` | `Click` | `Limpar` |
| `Command3` | `Click` | `Salvar` |
| `Command4` | `Click` | `Excluir` |
| `Command5` | `Click` | `Novo` |
| `Grid1` | `Click` | Seleciona item: preenche `TxtAutonum`, quantidade, embalagem, opcoes |

## Funcoes / subs no `.frm`

| Nome | Finalidade |
|------|------------|
| `Carrega_Grid1` | Monta `Adodc1` com itens de `tb_descarga_armazem` filtrados por `cntr` e `gate` |
| `Carrega_Placa` | Lista registros de saida com `flag_ddc=1` na ordem de carregamento |
| `CARREGAR` | Valida placa/lote; le `tb_registro_saida_cntr`; preenche `TxtOC`, `TxtCntr`; habilita `Frame1` e grid |
| `Limpar` | Reseta combos e campos auxiliares; desabilita `Frame1` |
| `Novo` | Prepara inclusao de item (zera `TxtAutonum`, limpa campos de item) |
| `Salvar` | INSERT/UPDATE em `tb_descarga_armazem`; atualiza `flag_ddc` em `tb_cntr_bl` |
| `Excluir` | DELETE com bloqueio se houver saida em `tb_amr_gate` |

## SQL e persistencia (extraido do legado)

### Lista de placas (combo inicial)

```590:600:d:\git\kleiton\romaneio\ProjetoReferencia\Coletor\FrmDescargaArmazem.frm
Sub Carrega_Placa()
    Dim SQl As String
    
    SQl = "select distinct a.autonum as autonum_rs, b.placa_c + ' (' + cc.id_conteiner + ')' as placa"
    SQl = SQl & " from sgipa..tb_registro_saida_cntr a"
    SQl = SQl & " inner join sgipa..tb_cntr_bl cc on a.cntr=cc.autonum"
    SQl = SQl & " inner join sgipa..tb_ordem_carregamento b on a.ordem_carreg = b.autonum"
    SQl = SQl & " where b.flag_saida=0 and b.flag_ddc=1"
    Adodc4.ConnectionString = ConnectionString
    Adodc4.RecordSource = SQl
    Adodc4.Refresh
```

### Gate aberto para a placa selecionada

```856:860:d:\git\kleiton\romaneio\ProjetoReferencia\Coletor\FrmDescargaArmazem.frm
        SQl = "select gn.autonum"
        SQl = SQl & " from sgipa..tb_registro_saida_cntr a"
        SQl = SQl & " inner join sgipa..tb_ordem_carregamento oc on a.ordem_carreg = oc.autonum"
        SQl = SQl & " inner join operador..tb_gate_new gn on oc.placa_c = gn.placa"
        SQl = SQl & " where a.autonum=" & cbPlaca.BoundText & " and gn.flag_gate_in=1 and isnull(gn.flag_gate_out,0)=0"
```

### Lotes ativos para a placa

```869:874:d:\git\kleiton\romaneio\ProjetoReferencia\Coletor\FrmDescargaArmazem.frm
        SQl = "select distinct bl.autonum as lote"
        SQl = SQl & " from sgipa..tb_registro_saida_cntr a"
        SQl = SQl & " inner join Sgipa..tb_amr_cntr_bl c on a.cntr = c.cntr"
        SQl = SQl & " inner join sgipa..tb_bl bl on c.bl = bl.autonum"
        SQl = SQl & " where a.autonum=" & cbPlaca.BoundText
        SQl = SQl & " and bl.flag_ativo=1"
```

### Grid de itens

```571:583:d:\git\kleiton\romaneio\ProjetoReferencia\Coletor\FrmDescargaArmazem.frm
Sub Carrega_Grid1()
    If Val(TxtGate.text) <> 0 And Val(TxtCntr) <> 0 Then
      SQl = "select a.quantidade, c.descr, A.FINALIZADO, a.autonum, A.EMBALAGEM "
      SQl = SQl & " from " & Banco_Sgipa & "tb_descarga_armazem a"
      SQl = SQl & " left join " & Banco_Operador & "tb_amr_gate b on a.autonum = b.id_descarga_armazem"
      SQl = SQl & " inner join " & Banco_Sgipa & "dte_tb_embalagens c on a.embalagem = c.code"
      SQl = SQl & " left join " & Banco_Operador & "tb_gate_new d on b.gate = d.autonum"
      SQl = SQl & " where "
      SQl = SQl & " a.cntr=" & TxtCntr
      SQl = SQl & " and a.gate=" & TxtGate
    Else
      SQl = "select null quantidade, null descr, null FINALIZADO, null autonum, null EMBALAGEM from dual"
```

### INSERT item (trecho principal)

```750:768:d:\git\kleiton\romaneio\ProjetoReferencia\Coletor\FrmDescargaArmazem.frm
        If TxtAutonum = "0" Then
            SQl = "insert into " & Banco_Sgipa & "tb_descarga_armazem ("
            If Banco_Em_Uso = "ORACLE" Then SQl = SQl & "autonum,"
            SQl = SQl & "cntr, quantidade, embalagem,  usuario, lote,gate,id_oc,finalizado"
            SQl = SQl & ") values ("
            If Banco_Em_Uso = "ORACLE" Then SQl = SQl & "" & Banco_Sgipa & "seq_tb_descarga_armazem.nextval,"
            SQl = SQl & "" & TxtCntr
            SQl = SQl & "," & TxtQuantidade
            SQl = SQl & ",'" & DcEmbalagem.BoundText & "'"
            SQl = SQl & "," & Cod_Usuario
            SQl = SQl & "," & TxtLote
            SQl = SQl & "," & TxtGate
            SQl = SQl & "," & TxtOC
            If Option1.Value = True Then
                SQl = SQl & ",1"
            Else
                SQl = SQl & ",0"
            End If
            SQl = SQl & ")"
```

### Atualizacao de status DDC no conteiner

```807:816:d:\git\kleiton\romaneio\ProjetoReferencia\Coletor\FrmDescargaArmazem.frm
        SQl = "select nvl(flag_ddc,0) flag_ddc from " & Banco_Sgipa & "tb_cntr_bl where autonum=" & TxtCntr
        PRSet Rst, SQl
        If Rst!Flag_DDC = 0 Then
            SQl = "update " & Banco_Sgipa & "tb_cntr_bl set flag_ddc=1,inicio_ddc=sysdate where autonum=" & TxtCntr
            Executa SQl
        End If
        If Option1.Value = True Then
            SQl = "update " & Banco_Sgipa & "tb_cntr_bl set flag_ddc=2 where autonum=" & TxtCntr
            Executa SQl
        End If
```

### Exclusao e bloqueio por saida

```635:647:d:\git\kleiton\romaneio\ProjetoReferencia\Coletor\FrmDescargaArmazem.frm
        SQl = "select sum(a.quantidade_saida) from " & Banco_Operador & "tb_amr_gate a"
        SQl = SQl & " where id_descarga_armazem=" & TxtAutonum
        PRSet Rs, SQl
        If Rs.Fields(0) <> 0 Then
            MsgBox "Jб consta saida para este item - Exclusгo Cancelada"
            Rs.Close
            Exit Sub
        End If
        Rs.Close
        
        
        SQl = "delete from " & Banco_Sgipa & "tb_descarga_armazem where autonum = " & TxtAutonum
```

## Transacoes / commits

- O legado usa `Executa` / `DB.Execute` por comando, **sem transacao explicita** envolvendo INSERT/UPDATEs concatenados em `Salvar`.
- **Hipotese**: `Executa` pode encapsular commit por instrucao; confirmar implementacao em `Microled.bas`.

## Atalhos de teclado

- `ESC`: fecha form.
- `F4`: `CARREGAR`.
- `F2`: `Limpar`.
- `F3`, `F5`, `F6`: acoes de item, bloqueadas com mensagem se `Frame1` desabilitado.

## Pontos criticos / lacunas tecnicas

1. **Handlers sem controle correspondente no designer**: existem `Private Sub cbConteiner_LostFocus`, `DataGrid1_Click`, `Command7_Click`, `MskPlaca_LostFocus`, `MaskEdBox1_Change` no codigo, mas **nao** ha `Begin VB... cbConteiner` / `DataGrid1` / `Command7` / `MskPlaca` na secao de desenho listada — **inconsistencia de compilacao possivel** ou form divergente do binario.
2. **`txtMarca` em `Salvar`**: referenciado na limpeza, **sem controle** no designer listado.
3. **`Form_Unload`**: contem `Unload Me` — recursao potencial.
4. **Mistura de sintaxe SQL**: `dual`, `nvl`, `seq_*.nextval`, `sysdate` vs `isnull` em outra query — depende de `CONVERTE_SQL` / banco em uso.
5. **Concatenacao de SQL** com valores de UI — risco de injecao e quebra por caracteres especiais.
6. **Logica de UPDATE de embalagem**: bloco aninhado `If Rs.Fields(0) <> 0 Then` duas vezes (linhas 784-796) torna o ramo `Else` da linha 794 **inalcancavel** quando ha saida — revisar com negocio.

## Imagem de referencia

- Nao foi localizada imagem da tela em `MigracaoVB6` nem anexo no repositorio para este estudo. Existe apenas `FrmDescargaArmazem.frx` (recursos binarios do form).
