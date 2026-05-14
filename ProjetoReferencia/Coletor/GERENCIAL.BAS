Attribute VB_Name = "Gerencial"
Public Calculando As Boolean

Public ValidaEntreposto As Boolean

Public rs_aux As New ADODB.Recordset
Public rs_aux2 As New ADODB.Recordset
Public LOCAL_REPORT As String
Public Tela_Pre_Calculo As Boolean
Public ORACLE_SQL_NULO As String 'PARA COMPATIBILIDADE DA FUNÇAO NVL COM SQL
Public sql_nulo As String
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public TEXTO As String
Global ListViewHeader_hWnd          As Long
Global lngOldProc                   As Long
Global Const GWL_WNDPROC            As Long = (-4)
Global Const WM_RBUTTONUP           As Long = &H205
Public WConnect As String
Public wBL_Master As Long
Public Pre_Calculo As Boolean
Public CalcAuto As Boolean
Public CalcAverba As Boolean
Public Importador_Anterior As Long
Public Indicador_Anterior As Long
Public Tipo_Registro_Saida As String
Public Gr_Vista As Boolean


Public LoteEntreposto As Long
Public BLentreposto As String

Public MinutosVersao As Byte
Public TentativaVersao As Byte

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Sub ReAssociaMarcantes(qLote As Long, QCntr As Long, qItem As Integer)
'Funcao para reassociar os marcantes que perderam referencia por cancelamento de Desova

Dim SQl As String
SQl = " SELECT C.AUTONUM AS AUTONUMCS,V.AUTONUM AS MARCANTE "
SQl = SQl & " FROM SGIPA..TB_CARGA_SOLTA C INNER JOIN SGIPA..VW_MARCANTES_REASSOCIAR V "
SQl = SQl & " ON C.BL=V.BL AND C.CNTR=V.CNTR AND C.ITEM=V.ITEM"
SQl = SQl & " WHERE "
SQl = SQl & " C.BL=" & qLote
SQl = SQl & " AND C.CNTR=" & QCntr
SQl = SQl & " AND C.ITEM=" & qItem
Set tbRAM = New ADODB.Recordset
tbRAM.Open SQl, db, 1, 1
While Not tbRAM.EOF
    SQl = "UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CARGA=" & tbRAM!AutonumCs & " WHERE AUTONUM=" & tbRAM!Marcante
    db.Execute SQl
    tbRAM.MoveNext
Wend
tbRAM.Close


Set tbRAM2 = New ADODB.Recordset
SQl = " select y.autonum,m.AUTONUM_CARGA from sgipa..tb_carga_solta_yard y inner join sgipa..tb_marcantes m on y.autonum=m.autonum_cs_yard where y.autonum_cs<>m.autonum_carga and m.autonum_carga>0 and m.bl=" & qLote
tbRAM2.Open SQl, db, 1, 1
While Not tbRAM2.EOF
    SQl = "update sgipa..tb_carga_solta_yard set autonum_cs=" & tbRAM2!autonum_carga & " where autonum=" & tbRAM2!autonum
    db.Execute SQl
    tbRAM2.MoveNext
Wend
tbRAM2.Close



End Sub

Public Sub Sendkeys(text$, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys text, wait
   Set WshShell = Nothing
End Sub


Sub AcertaDesovaCol(Lote As Long, CNTR As Long)
'desovas feitas parcialmente no coletor nao gravando autonum_carga na tb_marcantes

Dim SQl As String

SQl = "SELECT distinct S.AUTONUM,S.BL,S.CNTR,S.ITEM FROM TB_CARGA_SOLTA S "
SQl = SQl & " INNER JOIN TB_MARCANTES M ON"
SQl = SQl & " s.cntr = M.cntr"
SQl = SQl & " AND S.BL=M.BL"
SQl = SQl & " AND S.ITEM=M.ITEM"
SQl = SQl & " INNER JOIN TB_CNTR_BL C ON S.CNTR=C.AUTONUM"
SQl = SQl & " Where M.AUTONUM_CARGA = 0 And (c.DT_FIM_DESOVA Is Not Null or c.flag_desovado=1) "
If Lote > 0 Then
    SQl = SQl & " and s.bl=" & Lote
End If
If CNTR > 0 Then
    SQl = SQl & " and s.cntr=" & CNTR
End If
'Sql = Sql & " AND S.DT_SAIDA IS NULL"

Set tb1 = New ADODB.Recordset
tb1.Open SQl, db, 1, 1
While Not tb1.EOF
    SQl = "UPDATE TB_MARCANTES SET AUTONUM_CARGA=" & tb1!autonum & " WHERE AUTONUM_CARGA=0 AND BL=" & tb1!BL & " AND CNTR=" & tb1!CNTR & " AND ITEM=" & tb1!Item
    db.Execute SQl, N
    Debug.Print N
    tb1.MoveNext
Wend

End Sub



Sub Acerto_Estoque_DTE()
'If Dir(Me.txtArquivoOrigem.Text, vbArchive) = "" Then
'    MsgBox "Arquivo de Origem nao encontrado", vbExclamation
'    Exit Sub
'End If
'
'
'
'Dim Linha As String
'Dim Viagem As String
'Dim Conteiner As String
'Dim Tipo As String
'Dim FlagDta As Boolean
'Dim Envia As Boolean
'Dim Envia14 As Boolean
'Dim AutonumCntr As Long
'Dim Lote As Long
'
'Open Me.txtArquivoOrigem.Text For Input As #1
'If Right(txtDiretorioDestino, 1) <> "\" Then txtDiretorioDestino = txtDiretorioDestino & "\"
'Open Me.txtDiretorioDestino & "eve_23\log.txt" For Output As #2
'
'
'Me.lblLidos.Caption = 0
'
'While Not EOF(1)
'    FlagDta = False
'    Envia = False
'    Envia14 = False
'    AutonumCntr = 0
'    Lote = 0
'    MSG = " "
'
'    Line Input #1, Linha
'    Print #2, Left(Linha & Space(85), 85) & "   -   ";
'    Tipo = UCase(Trim(Mid$(Linha, 1, 16)))
'    If InStr(1, Tipo, "DTA") > 0 Then
'        FlagDta = True
'        DT = Mid$(Linha, 17, 8) & "/" & Mid$(Linha, 26, 4)
'    Else
'        DT = ""
'    End If
'    'Else
'    '    Dt = Mid$(Linha, 17, 7) & "-" & Mid$(Linha, 24, 1) & "/" & Mid$(Linha, 26, 4)
'    'End If
'    Viagem = Mid$(Linha, 20, 5) & "/" & Mid$(Linha, 26, 4)
'    Conteiner = Trim(Mid$(Linha, 31, 11))
'
'    If Len(Conteiner) = 11 Then
'        Conteiner = Left$(Conteiner, 10) & "-" & Right(Conteiner, 1)
'    End If
'
'    Set tb1 = New ADODB.Recordset
'    If Not FlagDta Then
'        tb1.Open "select autonum,viagem from tb_cntr_bl where viagem='" & Viagem & "' and id_conteiner='" & Conteiner & "' and flag_terminal=1", db, 1, 1
'    Else
'        tb1.Open "select autonum,viagem from tb_cntr_bl where dt='" & DT & "' and id_conteiner='" & Conteiner & "' and flag_terminal=1", db, 1, 1
'    End If
'    If Not tb1.EOF Then
'        AutonumCntr = tb1!Autonum
'        Viagem = tb1!Viagem
'    Else
'        MSG = "Conteiner nao encontrado"
'    End If
'    tb1.Close
'
'    If AutonumCntr > 0 Then
'        Sql = "SELECT A.BL FROM TB_AMR_CNTR_BL A INNER JOIN TB_BL B "
'        Sql = Sql & " ON A.BL=B.AUTONUM WHERE A.CNTR=" & AutonumCntr & " AND B.FLAG_ATIVO=1"
'        Set Tbl = New ADODB.Recordset
'        Tbl.Open Sql, db, 1, 1
'        If Not Tbl.EOF Then
'            Lote = Tbl!BL
'            'verifica o ultimo evento 23 enviado se está criticado
'            Sql = "Select flag_receita , obs_receita  from tb_auditoria_pacote where "
'            Sql = Sql & " autonum in (select max(autonum) from tb_auditoria_pacote where  bl=" & Lote & " and evento=23 )"
'            Set tbPAC = New ADODB.Recordset
'            tbPAC.Open Sql, db, 1, 1
'            If Not tbPAC.EOF Then
'                If tbPAC!flag_receita = 1 Then 'negado
'                    Envia = True
'                    If InStr(Nnull(tbPAC!Obs_Receita, 1), "JA POSSUI DOCUMENTO DE SAIDA") <> 0 Then Envia = False
'
'                    If InStr(Nnull(tbPAC!Obs_Receita, 1), "SEM ENTRADA") <> 0 Then
'                        Dim Rsn As New ADODB.Recordset
'                        Sql = "select a.* from " & Banco_Operador & "tb_amr_gate a"
'                        Sql = Sql & ", " & Banco_Operador & "tb_gate_new b"
'                        Sql = Sql & ", " & Banco_Operador & "tb_gate_opcoes c"
'                        Sql = Sql & " where"
'                        Sql = Sql & " a.funcao_gate = c.codigo_opcao"
'                        Sql = Sql & " and c.funcao_carga='E'"
'                        Sql = Sql & " and nvl(a.historico,0)=0"
'                        Sql = Sql & " and a.gate = b.autonum"
'                        Sql = Sql & " and a.cntr_ipa=" & AutonumCntr & " and b.patio=" &  PATIO_COletor & " "
'                        PRSet Rsn, Sql
'                        If Not Rsn.EOF Then
'                            If Not TempMonta_14("CNTR", "B", Rsn!gate, Rsn!Autonum, 0, Nnull(seqMov, 0), Me.txtDiretorioDestino & "eve_14\") Then
'                                MSG = "Evento 14 nao gerado - Erro na geracao do pacote"
'                                Envia = False
'                            Else
'                                MSG = "Evento 14 gerado"
'                            End If
'                        End If
'                    End If
'
'                    If Envia Then
'                        MSG = "Evento 23 gerado, critica anterior : " & Nnull(tbPAC!Obs_Receita, 1)
'                    Else
''                        msg= "Evento 23 nao gerado devido falha no envio do 14"
'                    End If
'
'                Else
'                    MSG = "Evento 23 não gerado, evento já estava aceito do lote " & Lote
'                End If
'            Else
'                Envia = True
'                MSG = "Pacote 23 gerado "
'            End If
'            tbPAC.Close
'
'            If Envia Then
'                Call tempGera_Evento23(Viagem, Lote, Me.txtDiretorioDestino & "eve_23\", True)
'                Me.lblEventosGerados.Caption = Val(Me.lblEventosGerados.Caption) + 1
'                Me.lblEventosGerados.Refresh
'            End If
'
'        End If
'
'    End If
'
'    Print #2, MSG
'
'    Me.lblLidos.Caption = Val(Me.lblLidos.Caption) + 1
'    Me.lblLidos.Refresh
'Wend
'
'Close (1)
'Close (2)

End Sub

 Sub Calcula_Impostos_Bl(Lote As Long, SEQ_GR As Long)

Dim Imposto_Total As Currency
Dim Fator As Double
Dim Lista As Double
Dim tb As New ADODB.Recordset
 
'calculo de imposto na tela de calculo

 PRSet tb, "Select autonum_lista from tb_bl where autonum=" & Lote
 If Not tb.EOF Then
    Lista = Nnull(tb!AUTONUM_LISTA, 0)
    SQl = Empty
    SQl = SQl & "delete "
    SQl = SQl & "FROM "
    SQl = SQl & "" & Banco_Sgipa & "TB_SERVICOS_FATURADOS_IMPOSTOS "
    SQl = SQl & "WHERE "
    SQl = SQl & "autonum_servico_faturado IN "
    If SEQ_GR = 0 Then SQl = SQl & "(select autonum FROM " & Banco_Sgipa & "TB_SERVICOS_FATURADOS WHERE BL = " & Lote & " and seq_gr IS NULL  )"
    If SEQ_GR <> 0 Then SQl = SQl & "(select autonum FROM " & Banco_Sgipa & "TB_SERVICOS_FATURADOS WHERE BL = " & Lote & " and seq_gr=" & SEQ_GR & "  )"
    Executa SQl
        
    SQl = " SELECT a.autonum, a.valor+a.adicional+a.desconto AS valor_servico, a.servico, c.autonum autonum_imposto, c.taxa"
    SQl = SQl & " ,(SELECT SUM(taxa) FROM " & Banco_Sgipa & "tb_lp_servicos_impostos d, " & Banco_Sgipa & "tb_cad_impostos e WHERE"
    SQl = SQl & " d.id_imposto = e.autonum AND d.id_servico = a.servico and d.id_tabela=" & Lista & " and d.calcular=1) Imposto_Total"
    SQl = SQl & ",a.vlr_sapiens_20, a.vlr_sapiens_40"
    SQl = SQl & " FROM"
    SQl = SQl & " " & Banco_Sgipa & "tb_servicos_faturados a"
    SQl = SQl & " ," & Banco_Sgipa & "tb_lp_servicos_impostos b"
    SQl = SQl & " ," & Banco_Sgipa & "tb_cad_impostos c"
    SQl = SQl & " WHERE "
    SQl = SQl & " b.id_tabela=" & Lista
    SQl = SQl & " and b.calcular=1"
    SQl = SQl & " and a.servico = b.id_servico"
    SQl = SQl & " AND b.id_imposto = c.autonum"
    If SEQ_GR = 0 Then SQl = SQl & " AND a.bl=" & Lote & " And a.seq_gr IS NULL  "
    If SEQ_GR <> 0 Then SQl = SQl & " AND a.bl=" & Lote & " And a.seq_gr = " & SEQ_GR
    SQl = SQl & " ORDER BY a.bl,a.servico"
    Set tb = New ADODB.Recordset
    tb.Open SQl, db, 3, 3
    Do While Not tb.EOF
        WValor = 0
        wvalor_20 = 0
        wvalor_40 = 0
        If Nnull(tb!Imposto_Total, 0) <> 0 Then
            Imposto_Total = Nnull(tb!valor_servico, 0)
            WValor = Round((Imposto_Total * tb!taxa) / 100, 2)
        
            If Nnull(tb!vlr_sapiens_20, 0) <> 0 Then wvalor_20 = Round((tb!vlr_sapiens_20 * tb!taxa) / 100, 2)
            If Nnull(tb!vlr_sapiens_40, 0) <> 0 Then wvalor_40 = Round((tb!vlr_sapiens_40 * tb!taxa) / 100, 2)
            If Nnull(tb!vlr_sapiens_20, 0) <> 0 And Nnull(tb!vlr_sapiens_40, 0) <> 0 Then
                wvalor_40 = WValor - wvalor_20
            End If

        Else
            WValor = 0
        End If
        SQl = "insert into " & Banco_Sgipa & "tb_servicos_faturados_impostos ("
        SQl = SQl & "autonum_servico_faturado,autonum_imposto,valor_imposto"
        SQl = SQl & ",valor_imposto_20,valor_imposto_40"
        SQl = SQl & ") values ("
        SQl = SQl & " " & tb!autonum
        SQl = SQl & " ," & tb!autonum_imposto
        SQl = SQl & " ," & PPonto(WValor) & ""
        SQl = SQl & " ," & PPonto(wvalor_20) & ""
        SQl = SQl & " ," & PPonto(wvalor_40) & ""
        SQl = SQl & ")"
        Executa SQl  'z4
        tb.MoveNext
    Loop

End If
tb.Close

End Sub

Sub Calcula_Desconto(Tipo As Integer, Lote As Long, SEQ_GR As Long, Id_Desc As Long)
Dim Tot1       As Currency
Dim Tot2       As Currency
Dim Tot3       As Currency
Dim Valor_Ger  As Currency
Dim ValorFinal1 As Currency
Dim ValorFinal2 As Currency
Dim Perc_Ger   As Currency
Dim Valor      As Currency
Dim Perc       As Currency
Dim ValorFinal As Currency
Dim DESCONTO   As Currency
Dim Lista   As Long
Dim Tbh As ADODB.Recordset
Dim TBD As ADODB.Recordset
Dim Tbg As ADODB.Recordset
Dim TBP As ADODB.Recordset
Dim tb As ADODB.Recordset
Dim Tb_Servicos As ADODB.Recordset
Dim Entrada As Boolean
Dim Filtro_Servico As String
Dim ValorTot As Currency
Dim Ini As Long
Dim Fim As Long
Dim z As Long
Dim SQl As String
Dim WSQL As String
Dim wReefer As Boolean


If Lote = 0 Then Exit Sub
SQl = "  SELECT "
SQl = SQl & " servico,descr,"
SQl = SQl & " vlr AS vlr,"
SQl = SQl & " desconto AS desconto,"
SQl = SQl & " adicional AS adicional, "
SQl = SQl & " liquido AS Liquido,"
SQl = SQl & " ordem_desc,"
SQl = SQl & " seq_gr,autonum,"
SQl = SQl & " TIPO,autonum_imposto"

SQl = SQl & "   FROM  "
SQl = SQl & "   (  "
SQl = SQl & "    (      "
SQl = SQl & "    SELECT a.servico,b.descr,a.valor AS vlr,a.desconto,a.adicional,(a.valor+ a.adicional+a.desconto) AS Liquido ,b.ordem_desc,a.seq_gr,a.autonum, 1 AS tipo, 0 as autonum_imposto"
SQl = SQl & "    FROM " & Banco_Sgipa & "TB_SERVICOS_FATURADOS a left join " & Banco_Sgipa & "TB_SERVICOS_IPA b"
SQl = SQl & "    on a.servico=b.autonum where   a.bl=" & Lote

If SEQ_GR <> 0 Then
  SQl = SQl & "    and  a.seq_gr=" & SEQ_GR
Else
  SQl = SQl & "    and isnull(a.seq_gr,0)=0"
End If
SQl = SQl & "    )  "
SQl = SQl & "    UNION     "
SQl = SQl & "    (         "
SQl = SQl & "    SELECT 0 as servico,c.descricao AS descr ,SUM(b.VALOR_imposto) AS vlr ,0 AS desconto ,0 AS adicional  ,SUM(b.VALOR_imposto) AS liquido ,0,a.seq_gr ,0 AS autonum,2 AS tipo, c.autonum as autonum_imposto"
SQl = SQl & "    FROM " & Banco_Sgipa & "tb_servicos_faturados a ," & Banco_Sgipa & "tb_servicos_faturados_impostos b ," & Banco_Sgipa & "tb_cad_impostos c         "
SQl = SQl & "    WHERE a.autonum = b.autonum_servico_faturado AND b.autonum_imposto = c.autonum AND a.bl=" & Lote
If SEQ_GR <> 0 Then
  SQl = SQl & "    and  a.seq_gr=" & SEQ_GR
Else
  SQl = SQl & "    and isnull(a.seq_gr,0)=0"
End If
SQl = SQl & "    GROUP BY     "
SQl = SQl & "    c.descricao,a.seq_gr,C.AUTONUM"
SQl = SQl & "    )  "
SQl = SQl & "    )  "
If Banco_Em_Uso = "SQLSERVER" Then SQl = SQl & " ALIAS_1 "
SQl = SQl & "    ORDER BY seq_gr,tipo,ordem_desc ,descr"
Tot3 = 0
PRSet Tbh, SQl
If Tbh.EOF Then Exit Sub
    While Not Tbh.EOF = True
        If Tbh!Tipo = 1 Then Tot3 = Tot3 + Tbh!Liquido
        Tbh.MoveNext
    Wend
Tbh.MoveFirst
i = 0
Tot1 = 0
Tot2 = 0
Valor_Ger = 0
Perc_Ger = 0
ValorTot = 0
ValorFinal1 = 0
ValorFinal2 = 0
SQl = Empty
SQl = "Select * from " & Banco_Sgipa & "tb_val_comercial where flag_aprovado = 0 and autonum = " & Id_Desc
PRSet TBD, SQl
If Not TBD.EOF Then
    SQl = "Select Count(1) as contar from " & Banco_Sgipa & "TB_DESC_COMERCIAL Where AUTONUM_VAL_COM=" & Id_Desc & " "
    PRSet Tbg, SQl
    WSQL = Empty
    WSQL = WSQL & "DELETE "
    WSQL = WSQL & "FROM "
    WSQL = WSQL & "" & Banco_Sgipa & "TB_SERVICOS_DESCONTOS_IMPOSTOS "
    WSQL = WSQL & "WHERE "
    WSQL = WSQL & "autonum_servico_DESCONTO IN "
    WSQL = WSQL & "(select autonum FROM " & Banco_Sgipa & "TB_SERVICOS_DESCONTOS WHERE AUTONUM_VAL_COM= " & Id_Desc
    WSQL = WSQL & ")"
    Executa WSQL
    WSQL = Empty
    WSQL = WSQL & "DELETE FROM " & Banco_Sgipa & "TB_SERVICOS_DESCONTOS WHERE AUTONUM_VAL_COM= " & Id_Desc
    Executa WSQL
    SQl = "Select DESC_VAL,DESC_PER from " & Banco_Sgipa & "TB_DESC_COMERCIAL Where  nvl(Servicos,0)=0 and (DESC_VAL>0 OR DESC_PER>0) AND AUTONUM_VAL_COM=" & Id_Desc & " and AllIn=0 and nvl(PRO_RATA,0)=0 order by autonum"
    PRSet TBD, SQl
    If Not TBD.EOF Then
       Do
           Valor_Ger = Valor_Ger + Nnull(TBD!desc_val, 0)
           Perc_Ger = Perc_Ger + Nnull(TBD!desc_per, 0)
           TBD.MoveNext
       Loop Until TBD.EOF
       If Valor_Ger > 0 And Tot3 > 0 Then
         Perc_Ger = Perc_Ger + ((Valor_Ger / Tot3) * 100)
         Valor_Ger = 0
       End If
    End If
    Tot1 = 0
    SQl = "Select DESC_VAL from " & Banco_Sgipa & "TB_DESC_COMERCIAL  Where  DESC_VAL>0 AND AUTONUM_VAL_COM=" & Id_Desc & " and AllIn=2 and nvl(PRO_RATA,0)=0  "
    PRSet TBD, SQl
    If Not TBD.EOF Then
        SQl = ""
        SQl = SQl & "        SELECT SUM (b.valor_imposto) AS liquido "
        SQl = SQl & "           FROM " & Banco_Sgipa & "tb_servicos_faturados a, "
        SQl = SQl & "                " & Banco_Sgipa & "tb_servicos_faturados_impostos b "
        SQl = SQl & "          WHERE a.autonum = b.autonum_servico_faturado "
        SQl = SQl & "        and   a.bl=" & Lote
        If SEQ_GR <> 0 Then
          SQl = SQl & "    and  a.seq_gr=" & SEQ_GR
        Else
          SQl = SQl & "    and nvl(a.seq_gr,0)=0"
        End If
        PRSet TBP, SQl
        ValorFinal = 0
        If Not TBP.EOF Then
           ValorFinal = Nnull(TBP!Liquido, 0)
        End If
        SQl = ""
        SQl = SQl & "SELECT SUM (a.valor + a.adicional + a.desconto) AS liquido "
        SQl = SQl & "           FROM " & Banco_Sgipa & "tb_servicos_faturados a "
        SQl = SQl & "          WHERE   a.bl=" & Lote
        If SEQ_GR <> 0 Then
          SQl = SQl & "    and  a.seq_gr=" & SEQ_GR
        Else
          SQl = SQl & "    and nvl(a.seq_gr,0)=0"
        End If
        PRSet TBP, SQl
        If Not TBP.EOF Then
           If Nnull(TBP!Liquido, 0) > 0 Then ValorFinal = ValorFinal / Nnull(TBP!Liquido, 0)
        End If
        Do
'         Valor_Ger = 0
          Valor_Ger = Valor_Ger + (Nnull(TBP!Liquido, 0) - (Nnull(TBD!desc_val, 0) / (1 + ValorFinal)))
          Valor_Ger = Fix(Valor_Ger * 100) / 100
          ValorFinal2 = Nnull(TBD!desc_val, 0)
          TBD.MoveNext
       Loop Until TBD.EOF
    End If
    While Not Tbh.EOF = True
        i = i + 1
        DESCONTO = 0
        Valor = 0
        Perc = 0
        Tot1 = Tot1 + Tbh!Liquido
        If Nnull(Tbh!Servico, 0) > 0 Then
            SQl = "Select DESC_VAL,DESC_PER  from " & Banco_Sgipa & "TB_DESC_COMERCIAL Where  Servicos=" & Tbh!Servico & " and (DESC_VAL>0 OR DESC_PER>0) AND AUTONUM_VAL_COM=" & Id_Desc & " and AllIn=0 and nvl(PRO_RATA,0)=0 order by autonum"
            PRSet TBD, SQl
            If Not TBD.EOF Then
               Do
                 Valor = Valor + Nnull(TBD!desc_val, 0)
                 Perc = Perc + Nnull(TBD!desc_per, 0)
                 TBD.MoveNext
               Loop Until TBD.EOF
               If Valor > 0 Then
                    SQl = ""
                    SQl = SQl & "        SELECT SUM (b.valor_imposto) AS liquido "
                    SQl = SQl & "           FROM " & Banco_Sgipa & "tb_servicos_faturados a, "
                    SQl = SQl & "                " & Banco_Sgipa & "tb_servicos_faturados_impostos b "
                    SQl = SQl & "          WHERE A.SERVICO=" & Tbh!Servico & " AND a.autonum = b.autonum_servico_faturado "
                    SQl = SQl & "        and   a.bl=" & Lote
                    If SEQ_GR <> 0 Then
                      SQl = SQl & "    and  a.seq_gr=" & SEQ_GR
                    Else
                      SQl = SQl & "    and nvl(a.seq_gr,0)=0"
                    End If
                   PRSet TBP, SQl
                   ValorFinal = 0
                    If Not TBP.EOF Then
                       ValorFinal = Nnull(TBP!Liquido, 0)
                    End If
                    If (Tbh!Liquido + ValorFinal) > 0 Then
                 '     Perc = Perc + ((Valor / (Tbh!Liquido + ValorFinal)) * 100)
                 '     Valor = 0
                    End If
               End If
            End If
            SQl = "Select Periodo,Pro_Rata,Flag_FreeTime  from " & Banco_Sgipa & "TB_DESC_COMERCIAL Where  Servicos=" & Tbh!Servico & " And AUTONUM_VAL_COM=" & Id_Desc & " and AllIn=0 and  Nvl(Periodo,0)>0 order by autonum"
            PRSet TBD, SQl
            If Not TBD.EOF Then
               Do
                 WSQL = "Select Parcela, To_Date(DataFimPeriodo,'dd/mm/yyyy')-To_Date(DataIniPeriodo,'dd/mm/yyyy') as dias "
                 WSQL = WSQL & "From Tb_Descricao_Calculo Where "
                 WSQL = WSQL & " Bl=" & Lote
                 If SEQ_GR <> 0 Then
                   WSQL = WSQL & " and seq_gr=" & SEQ_GR
                 Else
                   WSQL = WSQL & " and nvl(seq_gr,0)=0"
                 End If
                 WSQL = WSQL & " And Servico=" & Tbh!Servico
                 WSQL = WSQL & " And Periodos=" & TBD!Periodo
                 PRSet TBP, WSQL
                 If Not TBP.EOF Then
                    Do
                       If Nnull(TBP!Dias, 0) > 0 Then
                            Valor = Valor + (Nnull(TBP!Parcela, 0) / Nnull(TBP!Dias, 0)) * (Nnull(TBP!Dias, 0) - (Nnull(TBD!pro_rata, 0)))
                       End If
                       TBP.MoveNext
                    Loop Until TBP.EOF
                 End If
                 TBD.MoveNext
               Loop Until TBD.EOF
            End If
            SQl = "Select A.DESC_VAL,B.PERIODO_INI, B.PERIODO_FIM ,ALLIN from " & Banco_Sgipa & "TB_DESC_COMERCIAL A, " & Banco_Sgipa & "TB_ALL_IN B Where   A.AUTONUM=B.AUTONUM_DESC_COM AND  DESC_VAL>0 AND AUTONUM_VAL_COM=" & Id_Desc & " and AllIn=1 and nvl(PRO_RATA,0)=0 And B.Servicos=" & Tbh!Servico & " order by a.autonum"
            PRSet TBD, SQl
            If Not TBD.EOF Then
               Do
                 If Nnull(TBD!periodo_ini, 0) > 0 Then
                    Ini = Nnull(TBD!periodo_ini, 0)
                    Fim = Nnull(TBD!periodo_fim, 0)
                    If Ini > Fim Then Fim = Ini
                    WSQL = "Select max(periodos) as permax ,min(periodos) as perini  "
                    WSQL = WSQL & "From Tb_Descricao_Calculo Where "
                    WSQL = WSQL & " Bl=" & Lote
                    If SEQ_GR <> 0 Then
                      WSQL = WSQL & " and seq_gr=" & SEQ_GR
                    Else
                      WSQL = WSQL & " and nvl(seq_gr,0)=0"
                    End If
                    WSQL = WSQL & " And Servico=" & Tbh!Servico
                    PRSet TBP, WSQL
                    If Not TBP.EOF Then
                       If Ini <> Nnull(TBP!perini, 0) Or Fim <> Nnull(TBP!permax, 0) Then
                          For z = Ini To Fim
                             WSQL = "Select Parcela "
                             WSQL = WSQL & "From Tb_Descricao_Calculo Where "
                             WSQL = WSQL & " Bl=" & Lote
                             If SEQ_GR <> 0 Then
                               WSQL = WSQL & " and seq_gr=" & SEQ_GR
                             Else
                               WSQL = WSQL & " and nvl(seq_gr,0)=0"
                             End If
                             WSQL = WSQL & " And Servico=" & Tbh!Servico
                             WSQL = WSQL & " And Periodos=" & z
                             PRSet TBP, WSQL
                             If Not TBP.EOF Then
                                Do
                                   Valor = Valor + Nnull(TBP!Parcela, 0)
                                   TBP.MoveNext
                                Loop Until TBP.EOF
                             End If
                          Next z
                       Else
                          Valor = Valor + Nnull(Tbh!Liquido, 0)
                       End If
                    Else
                       Valor = Valor + Nnull(Tbh!Liquido, 0)
                    End If
                 Else
                   Valor = Valor + Nnull(Tbh!Liquido, 0)
                 End If
                 ValorTot = Nnull(TBD!desc_val, 0)
                 If ValorTot > Valor Then ValorTot = 0
                 TBD.MoveNext
               Loop Until TBD.EOF
            End If
            If Nnull(Tbg!Contar, 0) > 0 Then
                DESCONTO = Valor + Abs(Nnull(Tbh!DESCONTO, 0))
                DESCONTO = DESCONTO + (Nnull(Tbh!Liquido, 0) * ((Perc + Perc_Ger) / 100))
                If Valor_Ger > 0 Then
                   If (Nnull(Tbh!Liquido, 0) - DESCONTO) > Valor_Ger Then
                      DESCONTO = DESCONTO + Valor_Ger
                      Valor_Ger = 0
                   Else
                      Valor_Ger = Valor_Ger - ((Nnull(Tbh!Liquido, 0) - DESCONTO))
                      DESCONTO = DESCONTO + (Nnull(Tbh!Liquido, 0) - DESCONTO)
                   End If
                End If
                WSQL = "INSERT INTO " & Banco_Sgipa & "TB_SERVICOS_DESCONTOS"
                WSQL = WSQL & " (AUTONUM,BL,SEQ_GR,SERVICO,VALOR,DESCONTO,adicional,Autonum_Fat,Desconto_Ant,Autonum_Val_Com"
                WSQL = WSQL & ") VALUES ("
                WSQL = WSQL & " " & Banco_Sgipa & "SEQ_SERVICOS_DESCONTOS.NEXTVAL"
                WSQL = WSQL & "," & Lote
                If SEQ_GR <> 0 Then
                  WSQL = WSQL & "," & SEQ_GR
                Else
                  WSQL = WSQL & ",0"
                End If
                WSQL = WSQL & "," & Tbh!Servico
                WSQL = WSQL & "," & str(Nnull(Tbh!VLR, 0))
                If DESCONTO > 0 Then
                    If CLng(Fix(DESCONTO * 100)) - CLng(Fix((Tbh!Liquido) * 100)) > 0 Then
                       WSQL = WSQL & "," & str(Nnull((Tbh!Liquido) * -1, 0))
                    Else
                       WSQL = WSQL & "," & str(Nnull(DESCONTO * -1, 0))
                    End If
                Else
                  WSQL = WSQL & ",0"
                End If
                WSQL = WSQL & "," & str(Nnull(Tbh!Adicional, 0))
                WSQL = WSQL & "," & Tbh!autonum
                WSQL = WSQL & "," & str(Nnull(Tbh!DESCONTO, 0))
                WSQL = WSQL & "," & Id_Desc & ")"
                Executa (WSQL)
              End If
            End If
        Tbh.MoveNext
    Wend
    TxtTotal = FormatNumber(Tot1, 2)
    If ValorTot > 0 Then
        SQl = "Select distinct b.servicos from " & Banco_Sgipa & "TB_DESC_COMERCIAL A, " & Banco_Sgipa & "TB_ALL_IN B Where  A.AUTONUM=B.AUTONUM_DESC_COM AND  DESC_VAL>0 AND AUTONUM_VAL_COM=" & Id_Desc & " and AllIn=1 and nvl(PRO_RATA,0)=0 "
        PRSet TBD, SQl
        If Not TBD.EOF Then
           Do
                SQl = "select * from tb_servicos_descontos  "
                SQl = SQl & " WHERE bl=" & Lote & " and servico=" & TBD!Servicos & " "
                If SEQ_GR <> 0 Then
                   SQl = SQl & "And Seq_gr=" & SEQ_GR
                Else
                   SQl = SQl & "And Seq_gr=0"
                End If
                PRSet TBP, SQl
                If Not TBP.EOF Then
                   If Abs(TBP!DESCONTO) > ValorTot Then
                      SQl = "Update tb_servicos_descontos set desconto=desconto +" & Trim(str(ValorTot)) & " where autonum=" & TBP!autonum
                      If ValorTot > 0 Then Executa SQl
                      ValorTot = 0
                    Else
                      SQl = "Update tb_servicos_descontos set desconto=0  where autonum=" & TBP!autonum
                      Executa SQl
                      ValorTot = ValorTot - Abs(Nnull(TBP!DESCONTO, 0))
                   End If
                End If
                TBD.MoveNext
           Loop Until TBD.EOF
        End If
    End If
    SQl = "Select Autonum_Lista  from " & Banco_Sgipa & "TB_BL Where  autonum = " & Lote
    PRSet TBD, SQl
    If Not TBD.EOF Then
      Lista = TBD!AUTONUM_LISTA
    End If
    WSQL = Empty
    WSQL = WSQL & " SELECT a.autonum,a.autonum_fat,(a.valor + a.adicional + a.desconto) AS valor_servico, a.servico, c.autonum autonum_imposto, c.taxa"
    WSQL = WSQL & " ,(SELECT SUM(taxa) FROM " & Banco_Sgipa & "tb_lp_servicos_impostos d, " & Banco_Sgipa & "tb_cad_impostos e WHERE"
    WSQL = WSQL & " d.id_imposto = e.autonum AND d.id_servico = a.servico and d.id_tabela=" & Lista & " and d.calcular=1) Imposto_Total"
    WSQL = WSQL & " FROM"
    WSQL = WSQL & " " & Banco_Sgipa & "tb_servicos_descontos a "
    WSQL = WSQL & " ," & Banco_Sgipa & "tb_lp_servicos_impostos b "
    WSQL = WSQL & " ," & Banco_Sgipa & "tb_cad_impostos c "
    WSQL = WSQL & " WHERE "
    WSQL = WSQL & " b.id_tabela=" & Lista
    WSQL = WSQL & " and b.calcular=1"
    WSQL = WSQL & " and a.servico = b.id_servico "
    WSQL = WSQL & " AND b.id_imposto = c.autonum "
    WSQL = WSQL & " AND a.AUTONUM_VAL_COM= " & Id_Desc
    WSQL = WSQL & " ORDER BY a.bl,a.servico"
    PRSet tb, WSQL
    Do While Not tb.EOF
            WValor = 0
            If Nnull(tb!Imposto_Total, 0) <> 0 Then
                Imposto_Total = Nnull(tb!valor_servico, 0)
                WValor = Round((Imposto_Total * tb!taxa) / 100, 2)
            Else
                WValor = 0
            End If
         If Round(Nnull(tb!valor_servico, 0), 2) = 0 Then WValor = 0
         SQl = "Select * from " & Banco_Sgipa & "TB_SERVICOS_faturados_impostos Where Autonum_Imposto=" & tb!autonum_imposto & " and Autonum_Servico_Faturado=" & tb!Autonum_Fat
         PRSet TBD, SQl
         If Not TBD.EOF Then
           Valor = Nnull(TBD!valor_imposto, 0)
         End If
         SQl = "insert into " & Banco_Sgipa & "tb_servicos_descontos_impostos"
         SQl = SQl & " (autonum,autonum_servico_desconto,autonum_imposto,valor_imposto,valor_imposto_Ant) values ("
         SQl = SQl & " " & Banco_Sgipa & "SEQ_SERV_DESCONTOS_IMPOSTOS.nextval"
         SQl = SQl & " ," & tb!autonum
         SQl = SQl & " ," & tb!autonum_imposto
         SQl = SQl & " ," & PPonto(WValor)
         SQl = SQl & " ," & PPonto(Valor)
         SQl = SQl & ")"
         Executa SQl  'z4
         tb.MoveNext
     Loop
     tb.Close
     SQl = ""
     SQl = SQl & "SELECT a.autonum, (a.valor + a.adicional + a.desconto) AS liquido "
     SQl = SQl & "           FROM " & Banco_Sgipa & "tb_servicos_descontos a "
     SQl = SQl & "          WHERE   (a.valor + a.adicional + a.desconto)<=0 and a.AUTONUM_VAL_COM= " & Id_Desc
     PRSet TBP, SQl
     If Not TBP.EOF Then
        Do
          If Nnull(TBP!Liquido, 0) = 0 Then
             SQl = "Update tb_servicos_descontos_impostos set valor_imposto=0 WHERE   autonum_servico_desconto=" & TBP!autonum
             Executa SQl
          End If
          TBP.MoveNext
        Loop Until TBP.EOF
     End If
     
     If ValorFinal2 > 0 Then
        ValorFinal1 = 0
        SQl = ""
        SQl = SQl & "SELECT SUM (a.valor + a.adicional + a.desconto) AS liquido "
        SQl = SQl & "           FROM " & Banco_Sgipa & "tb_servicos_descontos a "
        SQl = SQl & "          WHERE   a.AUTONUM_VAL_COM= " & Id_Desc
        PRSet TBP, SQl
        If Not TBP.EOF Then
           ValorFinal1 = Nnull(TBP!Liquido, 0)
        End If
        SQl = ""
        SQl = SQl & "        SELECT SUM (b.valor_imposto) AS liquido,mAX(b.autonum) as autoimp"
        SQl = SQl & "           FROM " & Banco_Sgipa & "tb_servicos_descontos a, "
        SQl = SQl & "                " & Banco_Sgipa & "tb_servicos_descontos_impostos b "
        SQl = SQl & "          WHERE a.autonum = b.autonum_servico_desconto AND b.valor_imposto>0"
        SQl = SQl & "        and   a.AUTONUM_VAL_COM= " & Id_Desc
        PRSet TBP, SQl
        If Not TBP.EOF Then
           ValorFinal1 = ValorFinal1 + Nnull(TBP!Liquido, 0)
           If (Abs(ValorFinal2 - ValorFinal1) <> 0) And ValorFinal1 > 0 Then
              SQl = "Update tb_servicos_descontos_impostos set valor_imposto=valor_Imposto+" & Trim(str(ValorFinal2 - ValorFinal1)) & "  "
              SQl = SQl & "          WHERE   autonum=" & TBP!autoimp
              Executa SQl
           End If
        End If
     End If
     SQl = "  SELECT seq_gr,TIPO,ordem_desc,"
     SQl = SQl & " descr,servico,"
     SQl = SQl & " vlr AS vlr,"
     SQl = SQl & " desconto AS desconto,"
     SQl = SQl & " adicional AS adicional, "
     SQl = SQl & " liquido AS Liquido,"
     SQl = SQl & " autonum,"
     SQl = SQl & " autonum_imposto"
     SQl = SQl & "   FROM  "
     'Sql = Sql & "   (  "
     SQl = SQl & "    (      "
     SQl = SQl & "    SELECT a.servico,b.descr,a.valor AS vlr,a.desconto,a.adicional,(a.valor+ a.adicional+a.desconto) AS Liquido ,b.ordem_desc,a.seq_gr,a.autonum, 1 AS tipo, 0 as autonum_imposto"
     SQl = SQl & "    FROM " & Banco_Sgipa & "TB_SERVICOS_descontos a left join " & Banco_Sgipa & "TB_SERVICOS_IPA b "
     SQl = SQl & "    on a.servico=b.autonum where  a.AUTONUM_VAL_COM=  " & Id_Desc
     SQl = SQl & "    UNION     "
     SQl = SQl & "    SELECT 0 as servico,c.descricao AS descr ,SUM(b.VALOR_imposto) AS vlr ,0 AS desconto ,0 AS adicional  ,SUM(b.VALOR_imposto) AS liquido ,0,a.seq_gr ,0 AS autonum,2 AS tipo, c.autonum as autonum_imposto"
     SQl = SQl & "    FROM " & Banco_Sgipa & "tb_servicos_descontos a ," & Banco_Sgipa & "tb_servicos_descontos_impostos b ," & Banco_Sgipa & "tb_cad_impostos c         "
     SQl = SQl & "    WHERE a.autonum = b.autonum_servico_desconto AND b.autonum_imposto = c.autonum AND a.AUTONUM_VAL_COM= " & Id_Desc
     SQl = SQl & "    GROUP BY     "
     SQl = SQl & "    c.descricao,a.seq_gr,c.autonum "
     SQl = SQl & "    ) A "
     SQl = SQl & "    ORDER BY seq_gr,tipo,ordem_desc ,descr"
     PRSet Tbh, SQl
     i = 0
     Tot2 = 0
     While Not Tbh.EOF = True
         Tot2 = Tot2 + Tbh!Liquido
         Tbh.MoveNext
     Wend
     TxtTotalNew = FormatNumber(Tot2, 2)
Else
    While Not Tbh.EOF = True
        Tot1 = Tot1 + Tbh!Liquido
        TxtTotal = FormatNumber(Tot1, 2)
        Tbh.MoveNext
    Wend
End If





'rateia integracao sapiens
SQl = "select * from " & Banco_Sgipa & "tb_servicos_faturados where (isnull(vlr_sapiens_20,0) + isnull(vlr_sapiens_40,0)) > 0 and seq_gr IS NULL And BL = " & Lote
PRSet Rs, SQl
If Not Rs.EOF Then
     Do While Not Rs.EOF
        Valor = Nnull(Rs!Valor, 0) + Nnull(Rs!DESCONTO, 0) + Nnull(Rs!Adicional, 0)
           If Nnull(Rs!percent_sapiens_20, 0) = 100 Or Nnull(Rs!percent_sapiens_40, 0) = 100 Then
               If Nnull(Rs!percent_sapiens_20, 0) = 100 Then
                  SQl = "update " & Banco_Sgipa & "tb_servicos_faturados set percent_sapiens_40=0,vlr_sapiens_20=" & PPonto(Valor) & " where autonum=" & Rs!autonum
                Else
                  SQl = "update " & Banco_Sgipa & "tb_servicos_faturados set percent_sapiens_20=0,vlr_sapiens_40=" & PPonto(Valor) & " where autonum=" & Rs!autonum
                End If
                Executa SQl
           Else
                v20 = Round((Valor * Nnull(Rs!percent_sapiens_20, 0)) / 100, 2)
                v40 = Valor - v20
                SQl = "update " & Banco_Sgipa & "tb_servicos_faturados set vlr_sapiens_20=" & PPonto(v20)
                SQl = SQl & ", vlr_sapiens_40=" & PPonto(v40) & " where autonum=" & Rs!autonum
                Executa SQl
           End If
           Rs.MoveNext
     Loop
End If
Rs.Close




End Sub



Public Sub Exporta_CSV(CD As CommonDialog, Consulta As String, Optional Cabecalho As String = "")
    If Consulta = "" Then Exit Sub

    Dim LinhaExcel As String
    Dim c As Integer
    Dim RsExcel As New ADODB.Recordset
    Dim campo As String

    LinhaExcel = ""
    PRSet RsExcel, Consulta
    If RsExcel.EOF Then
        RsExcel.Close
        MsgBox "Nenhum registro encontrado para exportacao"
        Exit Sub
    End If
    
    CD.ShowSave
    If CD.FileName = "" Then Exit Sub
    If InStr(UCase(CD.FileName), ".CSV") = 0 Then
        Open CD.FileName & ".csv" For Output As #1
    Else
        Open CD.FileName For Output As #1
    End If
    
    If Cabecalho = "" Then
'        LinhaExcel = LinhaExcel & "Admissão;Lote;BL;Navio;Porto;Importador;Indicador;Item;Qtde;Embalagem"
'        LinhaExcel = LinhaExcel & ";Mercadoria;IMO;UN;Peso BL;Tipo Documento;Documento;Dias;TG;FMA;Terminal;Local;"
        LinhaExcel = ""
        For c = 0 To RsExcel.Fields.Count - 1
            campo = RsExcel.Fields(c).Name
            LinhaExcel = LinhaExcel & campo & ";"
        Next
    Else
        LinhaExcel = Cabecalho
    End If
    Print #1, UCase(LinhaExcel)
    
    Do While Not RsExcel.EOF
        LinhaExcel = ""
        For c = 0 To RsExcel.Fields.Count - 1
            campo = RsExcel.Fields(c).Name
            If Not IsNull(RsExcel.Fields(campo).Value) Then
                LinhaExcel = LinhaExcel & RsExcel.Fields(campo).Value
            Else
                LinhaExcel = LinhaExcel & ""
            End If
            LinhaExcel = LinhaExcel & ";"
        Next
        Print #1, UCase(LinhaExcel)
        RsExcel.MoveNext
    Loop
    RsExcel.Close
    Close #1
    MsgBox "Arquivo gerado no local especificado"
End Sub

Sub Verifica_Divergencia_Info_IMO(Lote As Long)

    Dim LotePesquisa As Long
    Dim SQl As String
    Dim ImoMaster As Boolean
    Dim Imo As Boolean
    ImoMaster = False
    Imo = False

    SQl = "select count(*) from sgipa.dbo.tb_bl where autonum=" & Lote
    If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then Exit Sub

    Do While True
        SQl = "select isnull(bl_master,0) from sgipa.dbo.tb_bl where autonum=" & Lote
        BLM = Nnull(db.Execute(SQl).Fields(0), 0)
        If BLM = 0 Then
            BLM = Lote
            Exit Do
        End If
        Lote = BLM
    Loop

    'Verifica se master é imo
    LotePesquisa = BLM
    GoSub Verifica_Imo
    ImoMaster = Imo
    
    
    
    'verifica existe algum submaster sem desconsolidacao
    Dim Desconsolidado_Total As Boolean
    Desconsolidado_Total = True
    SQl = "select"
    SQl = SQl & " blsm.autonum  as smlote  , blsm.grau  as smgrau , isnull(blsm.audit_desconsolidado,0)  as sm17,"
    SQl = SQl & " blc.autonum   as clote   , blc.grau   as cgrau  , isnull(blc.audit_desconsolidado,0)   as blc17,"
    SQl = SQl & " blcc.autonum  as cclote  , blcc.grau  as ccgrau , isnull(blcc.audit_desconsolidado,0)  as blcc17,"
    SQl = SQl & " blcc2.autonum as cc2lote , blcc2.grau as cc2grau, isnull(blcc2.audit_desconsolidado,0) as blcc217"
    SQl = SQl & " from sgipa.dbo.tb_bl blm"
    SQl = SQl & " left join sgipa.dbo.tb_bl blsm  on blm.autonum  = blsm.bl_master"
    SQl = SQl & " left join sgipa.dbo.tb_bl blc   on blsm.autonum = blc.bl_master"
    SQl = SQl & " left join sgipa.dbo.tb_bl blcc  on blc.autonum  = blcc.bl_master"
    SQl = SQl & " left join sgipa.dbo.tb_bl blcc2 on blcc.autonum = blcc2.bl_master"
    SQl = SQl & " where blm.autonum=" & BLM
    Dim Rs As New ADODB.Recordset
    PRSet Rs, SQl
    If Not Rs.EOF Then
        Do While Not Rs.EOF
            If Nnull(Rs!smgrau, 1) = "S" And Rs!sm17 = 0 Then Desconsolidado_Total = False
            If Nnull(Rs!cgrau, 1) = "S" And Rs!blc17 = 0 Then Desconsolidado_Total = False
            If Nnull(Rs!ccgrau, 1) = "S" And Rs!blcc17 = 0 Then Desconsolidado_Total = False
            If Nnull(Rs!cc2grau, 1) = "S" And Rs!blcc217 = 0 Then Desconsolidado_Total = False
            Rs.MoveNext
            If Not Desconsolidado_Total Then Exit Do
        Loop
        Rs.MoveFirst
        
        'se todos os sub-master foram desconsolidados, verifica se algum deles é imo
        If Desconsolidado_Total Then
            Imo = False
            Do While Not Rs.EOF
                LotePesquisa = Nnull(Rs!smlote, 0)
                If LotePesquisa <> 0 Then GoSub Verifica_Imo
                If Not Imo Then
                    LotePesquisa = Nnull(Rs!clote, 0)
                    If LotePesquisa <> 0 Then GoSub Verifica_Imo
                End If
                If Not Imo Then
                    LotePesquisa = Nnull(Rs!cclote, 0)
                    If LotePesquisa <> 0 Then GoSub Verifica_Imo
                End If
                If Not Imo Then
                    LotePesquisa = Nnull(Rs!cc2lote, 0)
                    If LotePesquisa <> 0 Then GoSub Verifica_Imo
                End If
                If Imo Then Exit Do
                Rs.MoveNext
            Loop
        
            'verifica se encontrou divergencia entre o master e os desconsolidados
            If ImoMaster <> Imo Then
                Dim MSG As String
                MSG = "Encontrada divergencia de informacao IMO"
                MSG = MSG & "BL Master : " & BLM & vbCr & vbLf
                MSG = MSG & "BL HOUSE  : " & LotePesquisa
            
                Dim Emails As String
                Emails = ""
                SQl = "SELECT * FROM SGIPA.DBO.TB_EMAIL_ALERTAS WHERE TIPO_ALERTA=9"
                Dim RSe As New ADODB.Recordset
                PRSet RSe, SQl
                Do While Not RSe.EOF
                    If Emails <> "" Then Emails = Emails & ";"
                    If Nnull(RSe!Email, 1) <> "" Then
                        Emails = Emails & RSe!Email
                    End If
                    RSe.MoveNext
                Loop
                RSe.Close
            
                If Emails <> "" Then
            
                    Envia Emails, "Divergencia de informação IMO durante a desconsolidação", MSG, 1, "", "", "", "CIABAND"
                    
                End If
            End If
            
        End If
    End If
    Rs.Close
    
    Exit Sub
    

Verifica_Imo:
    SQl = "select count(*)"
    SQl = SQl & " from sgipa.dbo.tb_amr_cntr_bl amr"
    SQl = SQl & " inner join sgipa.dbo.tb_cntr_bl cc on amr.cntr=cc.autonum"
    SQl = SQl & " where amr.bl=" & LotePesquisa
    SQl = SQl & " and (isnull(cc.imo1,0)<>0 or isnull(cc.imo2,0)<>0 or isnull(cc.imo3,0)<>0 or isnull(cc.imo4,0)<>0) "
    If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then Imo = True
    If Not Imo Then
        SQl = "select count(*)"
        SQl = SQl & " from sgipa.dbo.tb_carga_cntr"
        SQl = SQl & " where bl=" & LotePesquisa
        SQl = SQl & " and isnull(imo,0)<>0"
        If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then Imo = True
    End If
    Return



End Sub

Function Vincula_Tabela_Band_até_17092015(Lote As Long, Optional Atualizar As Boolean = True) As Long
'VINCULA TABELA DE PRECOS AO BL
'-----------------------------------------------------
Dim Data_Validade As Date
Dim Tbh As New ADODB.Recordset
Dim wImportador As Long
Dim wDespachante As Long
Dim wCaptador As Long
Dim wColoader As Long
Dim wNvocc As Long
Dim wVinculo As Long
Dim wBL As Long
Dim wViagem As String
Dim DtasIndeferida As Byte
Dim wHubPort As Byte
Dim Rs As New ADODB.Recordset
Dim Rs3 As New ADODB.Recordset

Dim tbh1 As New ADODB.Recordset
Dim wLista As Long
Dim wCargaPatio As Byte
Dim wLista_Anterior As Long
Dim wTipo_Vinculo As String
Dim Empresa_BL As Long
Dim Coloader As Boolean

Dim Codigo_Empresa As Byte


Coloader = False

wLista_Anterior = 0
wTipo_Vinculo = "A"

wBL = 0

SQl = "select nvl(a.flag_carga_patio,0) flag_carga_patio, nvl(a.autonum_lista,0) lista, nvl(a.flag_hubport,0) flag_hubport"
SQl = SQl & ", nvl(a.flag_dtas_indeferida,0) flag_dtas_indeferida, b.razao, nvl(a.bl_master,0) bl_master, a.viagem"
SQl = SQl & ", nvl(a.importador,0) importador,nvl(a.despachante,0) despachante ,nvl(a.captador,0) captador, a.codigo_empresa"
SQl = SQl & ", nvl(a.audit_av,0) audit_av, a.primeira_entrada, a.patio, a.tipo_documento, nvl(a.Flag_Redestinacao,0) as Redestinacao"
SQl = SQl & " from " & Banco_Sgipa & "tb_bl a left join  " & Banco_Sgipa & "tb_cad_parceiros b on a.captador = b.autonum where "
SQl = SQl & " a.autonum=" & Lote

PRSet Rs, SQl
Codigo_Empresa = Nnull(Rs!Codigo_Empresa, 0) + 1




'verifica tabela hub port
If Nnull(Rs!flag_hubport, 0) = 1 Then
    Dim TabelaHP As Integer
    SQl = "select tabela_hp from " & Banco_Sgipa & "tb_empresas where autonum = " & Codigo_Empresa
    TabelaHP = Nnull(db.Execute(SQl).Fields(0), 0)
    If TabelaHP <> 0 Then
        wVinculo = TabelaHP
        Rs.Close
        GoTo Padrao
    End If
End If
'***********************




SQl = "SELECT nvl(cod_empresa,0) cod_empresa FROM " & Banco_Operador & "TB_PATIOS WHERE autonum = " & Nnull(Rs!Patio, 0)
PRSet Rs3, SQl
Empresa_BL = Rs3!Cod_Empresa
Rs3.Close
If Nnull(Rs!redestinacao, 0) = 1 Then
   If Empresa_BL = 1 Then
      Empresa_BL = 2
   Else
      Empresa_BL = 1
   End If
End If

'busca data de validade pela data da entrada do conteiner
ERRO = 1
SQl = "select min(dt_entrada) dt_entrada from " & Banco_Sgipa & "tb_cntr_bl a, " & Banco_Sgipa & "tb_amr_cntr_bl b"
SQl = SQl & " where a.autonum = b.cntr"
SQl = SQl & " and b.bl=" & Lote
PRSet Rs3, SQl
If Not Rs3.EOF Then
    If Not IsNull(Rs3!Dt_Entrada) Then
        Data_Validade = Format$(Rs3!Dt_Entrada, "dd/mm/yyyy hh:mm:ss")
        ERRO = 0
    End If
End If
Rs3.Close
If ERRO = 1 Then
    SQl = "select min(dt_entrada) dt_entrada from " & Banco_Sgipa & "tb_carga_solta where bl=" & Lote
    PRSet Rs3, SQl
    If Not Rs3.EOF Then
        If Not IsNull(Rs3!Dt_Entrada) Then
            Data_Validade = Format$(Rs3!Dt_Entrada, "dd/mm/yyyy hh:mm:ss")
            ERRO = 0
        End If
    End If
    Rs3.Close
End If
If ERRO = 1 Then
    Data_Validade = Format$(Now, "dd/mm/yyyy hh:mm:ss")
End If


Dim Rs2 As New ADODB.Recordset
SQl = "select autonum_lista from " & Banco_Sgipa & "tb_lp_lotes where lote=" & Lote
PRSet Rs2, SQl
If Not Rs2.EOF Then
    wTipo_Vinculo = "M"
    wVinculo = Rs2!AUTONUM_LISTA
    Rs2.Close
    GoTo Padrao
End If
Rs2.Close

SQl = "select autonum_lista from " & Banco_Sgipa & "tb_lp_lotes where lote=" & Nnull(Rs!BL_Master, 0)
PRSet Rs2, SQl
If Not Rs2.EOF Then
    wTipo_Vinculo = "M"
    wVinculo = Rs2!AUTONUM_LISTA
    Rs2.Close
    GoTo Padrao
End If
Rs2.Close


wLista_Anterior = Rs!Lista
wViagem = Nnull(Rs!Viagem, 1)
wImportador = Rs!Importador
wDespachante = Rs!Despachante
wCaptador = Rs!Captador
wBL = Lote
DtasIndeferida = Rs!flag_dtas_indeferida
wLista = Rs!Lista

wVinculo = 0


'Verifica Margem de acordo com local de atracacao
Dim wMargem As String
wMargem = ""
SQl = "select margem from " & Banco_Sgipa & "DTE_TB_VIAGENS b, " & Banco_Sgipa & "DTE_TB_ARMAZENS a"
SQl = SQl & " where"
SQl = SQl & " b.arm_atracacao = a.code"
SQl = SQl & " and b.viagem='" & wViagem & "'"
PRSet Rs, SQl
If Not Rs.EOF Then
    wMargem = UCase(Nnull(Rs!Margem, 1))
End If
Rs.Close
If wMargem = "" Then wMargem = "MDIR"
'------------------------------------------------


'Verifica coloader
'-----------------------
Dim BLf As Long
Dim BLM As Long
Dim BLs As Long
Dim BLC As Long
Dim BL_Pai As Long
BLf = Lote  'Unico ou Filhote
SQl = "select bl_master from " & Banco_Sgipa & "tb_bl where autonum=" & BLf
BL_Pai = Nnull(db.Execute(SQl).Fields(0), 0)
If BL_Pai <> 0 Then
    BLM = BL_Pai    'master
    SQl = "select bl_master from " & Banco_Sgipa & "tb_bl where autonum=" & BLM
    BL_Pai = Nnull(db.Execute(SQl).Fields(0), 0)
    If BL_Pai <> 0 Then
        BLC = BLf       'coloader
        BLs = BLM       'sub master
        BLM = BL_Pai    'sub master
    End If
End If


If BLC <> 0 Then
    SQl = "select importador from " & Banco_Sgipa & "tb_bl where autonum=" & BLs
    wColoader = Nnull(db.Execute(SQl).Fields(0), 0)
    
    'procura tabela especifica  (nao olha o grupo)
    'coloader
    SQl = "select a.autonum"
    SQl = SQl & " from"
    SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
    SQl = SQl & " where"
    SQl = SQl & " a.flag_liberada=1"
    SQl = SQl & " and a.coloader=" & wColoader
    SQl = SQl & " and nvl(a.importador,0)=0"
    SQl = SQl & " and nvl(a.despachante,0)=0"
    SQl = SQl & " and nvl(a.captador,0)=0"
    SQl = SQl & " and nvl(a.flag_acordo,0)=0"
    SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
    SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
    PRSet Rs, SQl
    If Not Rs.EOF Then
        wVinculo = Rs!autonum
    End If
    Rs.Close

    If wVinculo = 0 Then
        SQl = "select a.autonumlista,a.tipo"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
        SQl = SQl & " where"
        SQl = SQl & " a.autonum_parceiro = " & wColoader
        SQl = SQl & " and a.tipo='S'"
        PRSet Rs, SQl
        Do While Not Rs.EOF And wVinculo = 0
            Tipo = Rs!Tipo
            wVinculo = Rs!autonumlista
            
            SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
            SQl = SQl & " flag_liberada=1"
            SQl = SQl & " and nvl(flag_acordo,0)=0"
            SQl = SQl & " and autonum=" & wVinculo
            SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
            SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
            PRSet Rs3, SQl
            If Not Rs3.EOF Then
                If Nnull(Rs3!Importador, 0) <> 0 Or Nnull(Rs3!Despachante, 0) <> 0 Or Nnull(Rs3!Captador, 0) <> 0 Then wVinculo = 0
            Else
                wVinculo = 0
            End If
            Rs3.Close
                
            Rs.MoveNext
        Loop
        Rs.Close
    End If

    If wVinculo = 0 Then            'coloader + importador + captador
        SQl = "select a.autonum"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
        SQl = SQl & " where"
        SQl = SQl & " a.flag_liberada=1"
        SQl = SQl & " and a.coloader=" & wColoader
        SQl = SQl & " and nvl(a.importador,0)=" & wImportador
        SQl = SQl & " and nvl(a.captador,0)=" & wCaptador
        SQl = SQl & " and nvl(a.despachante,0)=0"
        SQl = SQl & " and nvl(a.flag_acordo,0)=0"
        SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
        SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
        PRSet Rs, SQl
        If Not Rs.EOF Then
            wVinculo = Rs!autonum
        End If
        Rs.Close

        If wVinculo = 0 Then
            SQl = "select a.autonumlista,a.tipo"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
            SQl = SQl & " where"
            SQl = SQl & " a.autonum_parceiro = " & wColoader
            SQl = SQl & " and a.tipo='S'"
            PRSet Rs, SQl
            Do While Not Rs.EOF And wVinculo = 0
                wVinculo = Rs!autonumlista
                
                SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
                SQl = SQl & " flag_liberada=1"
                SQl = SQl & " and nvl(flag_acordo,0)=0"
                SQl = SQl & " and autonum=" & wVinculo
                SQl = SQl & " and nvl(despachante,0)=0"
                SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
                SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
                PRSet Rs3, SQl
                If Not Rs3.EOF Then
                   If Nnull(Rs3!Coloader, 0) = 0 Or Nnull(Rs3!Importador, 0) = 0 Or Nnull(Rs3!Captador, 0) = 0 Then wVinculo = 0
                   If wImportador <> Nnull(Rs3!Importador, 0) And wVinculo > 0 Then
                      'Verifica se importador faz parte do grupo da tabela
                      SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='I' and autonum_parceiro=" & wImportador
                      If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                   End If
                   If wVinculo <> 0 And wCaptador <> Nnull(Rs3!Captador, 0) Then
                       'Verifica se indicador faz parte do grupo da tabela
                       SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='C' and autonum_parceiro=" & wCaptador
                       If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                   End If
               Else
                  wVinculo = 0
               End If
               Rs3.Close
               Rs.MoveNext
            Loop
            Rs.Close
        End If
    End If

    If wVinculo = 0 Then            'coloader + importador
        SQl = "select a.autonum"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
        SQl = SQl & " where"
        SQl = SQl & " a.flag_liberada=1"
        SQl = SQl & " and a.coloader=" & wColoader
        SQl = SQl & " and nvl(a.importador,0)=" & wImportador
        SQl = SQl & " and nvl(a.despachante,0)=0"
        SQl = SQl & " and nvl(a.captador,0)=0"
        SQl = SQl & " and nvl(a.flag_acordo,0)=0"
        SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
        SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
        PRSet Rs, SQl
        If Not Rs.EOF Then
            wVinculo = Rs!autonum
        End If
        Rs.Close

        If wVinculo = 0 Then
            SQl = "select a.autonumlista,a.tipo"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
            SQl = SQl & " where"
            SQl = SQl & " a.autonum_parceiro = " & wColoader
            SQl = SQl & " and a.tipo='S'"
            PRSet Rs, SQl
            Do While Not Rs.EOF And wVinculo = 0
                wVinculo = Rs!autonumlista
                
                SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
                SQl = SQl & " flag_liberada=1"
                SQl = SQl & " and nvl(flag_acordo,0)=0"
                SQl = SQl & " and autonum=" & wVinculo
                SQl = SQl & " and nvl(despachante,0)=0"
                SQl = SQl & " and nvl(captador,0)=0"
                SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
                SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
                PRSet Rs3, SQl
                If Not Rs3.EOF Then
                   If Nnull(Rs3!Coloader, 0) = 0 Or Nnull(Rs3!Importador, 0) = 0 Then wVinculo = 0
                   If wImportador <> Nnull(Rs3!Importador, 0) And wVinculo > 0 Then
                      'Verifica se importador faz parte do grupo da tabela
                      SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='I' and autonum_parceiro=" & wImportador
                      If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                   End If
                Else
                    wVinculo = 0
                End If
                Rs3.Close
                    
                Rs.MoveNext
            Loop
            Rs.Close
        End If
    End If


    If wVinculo = 0 Then            'coloader + indicador
        SQl = "select a.autonum"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
        SQl = SQl & " where"
        SQl = SQl & " a.flag_liberada=1"
        SQl = SQl & " and a.coloader=" & wColoader
        SQl = SQl & " and nvl(a.captador,0)=" & wCaptador
        SQl = SQl & " and nvl(a.despachante,0)=0"
        SQl = SQl & " and nvl(a.importador,0)=0"
        SQl = SQl & " and nvl(a.flag_acordo,0)=0"
        SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
        SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
        PRSet Rs, SQl
        If Not Rs.EOF Then
            wVinculo = Rs!autonum
        End If
        Rs.Close

        If wVinculo = 0 Then
            SQl = "select a.autonumlista,a.tipo"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
            SQl = SQl & " where"
            SQl = SQl & " a.autonum_parceiro = " & wColoader
            SQl = SQl & " and a.tipo='S'"
            PRSet Rs, SQl
            Do While Not Rs.EOF And wVinculo = 0
                wVinculo = Rs!autonumlista
                
                SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
                SQl = SQl & " flag_liberada=1"
                SQl = SQl & " and nvl(flag_acordo,0)=0"
                SQl = SQl & " and autonum=" & wVinculo
                SQl = SQl & " and nvl(despachante,0)=0"
                SQl = SQl & " and nvl(importador,0)=0"
                SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
                SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
                PRSet Rs3, SQl
                If Not Rs3.EOF Then
                   If Nnull(Rs3!Coloader, 0) = 0 Or Nnull(Rs3!Captador, 0) = 0 Then wVinculo = 0
                   If wVinculo <> 0 And wCaptador <> Nnull(Rs3!Captador, 0) Then
                       'Verifica se indicador faz parte do grupo da tabela
                       SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='C' and autonum_parceiro=" & wCaptador
                       If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                   End If
                Else
                    wVinculo = 0
                End If
                Rs3.Close
                    
                Rs.MoveNext
            Loop
            Rs.Close
        End If
    End If

    If wVinculo = 0 Then GoTo Regra2

Else
    
Regra2:
    
    'indicador + importador
    SQl = "select a.autonum"
    SQl = SQl & " from"
    SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
    SQl = SQl & " where"
    SQl = SQl & " a.flag_liberada=1"
    SQl = SQl & " and nvl(a.importador,0)=" & wImportador
    SQl = SQl & " and nvl(a.coloader,0)=0"
    SQl = SQl & " and nvl(a.despachante,0)=0"
    SQl = SQl & " and nvl(a.captador,0)=" & wCaptador
    SQl = SQl & " and nvl(a.flag_acordo,0)=0"
    SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
    SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
    PRSet Rs, SQl
    If Not Rs.EOF Then
        wVinculo = Rs!autonum
    End If
    Rs.Close
    
    
            If wVinculo = 0 And wCaptador <> 0 And wImportador <> 0 Then
            SQl = "select a.autonumlista,a.tipo"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
            SQl = SQl & " where a.tipo='C' and "
            SQl = SQl & " a.autonum_parceiro = " & wCaptador
            PRSet Rs, SQl
            
            SQl = "select a.autonumlista,a.tipo"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
            SQl = SQl & " where a.tipo='I' and "
            SQl = SQl & " a.autonum_parceiro = " & wImportador
            PRSet Rs2, SQl
        
            Do While Not Rs.EOF And wVinculo = 0
                If Not Rs2.EOF Then
                    Do While Not Rs2.EOF And wVinculo = 0
                        If Rs!autonumlista = Rs2!autonumlista Then
                            wVinculo = Rs!autonumlista
                            SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
                            SQl = SQl & " flag_liberada=1"
                            SQl = SQl & " and autonum=" & wVinculo
                            SQl = SQl & " and nvl(coloader,0)=0"
                            SQl = SQl & " and nvl(despachante,0)=0"
                            SQl = SQl & " and nvl(flag_acordo,0)=0"
                            SQl = SQl & " and data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
                            SQl = SQl & " and (final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
                            PRSet Rs3, SQl
                            If Rs3.EOF Then
                                wVinculo = 0
                            End If
                            Rs3.Close
                        End If
                        Rs2.MoveNext
                    Loop
                    Rs2.MoveFirst
                End If
                Rs.MoveNext
            Loop
            Rs.Close
            Rs2.Close
        End If
    
    

    'importador
    If wVinculo = 0 Then
        SQl = "select a.autonum"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
        SQl = SQl & " where"
        SQl = SQl & " a.flag_liberada=1"
        SQl = SQl & " and nvl(a.importador,0)=" & wImportador
        SQl = SQl & " and nvl(a.coloader,0)=0"
        SQl = SQl & " and nvl(a.captador,0)=0"
        SQl = SQl & " and nvl(a.despachante,0)=0"
        SQl = SQl & " and nvl(a.flag_acordo,0)=0"
        SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
        SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
        PRSet Rs, SQl
        If Not Rs.EOF Then
            wVinculo = Rs!autonum
        End If
        Rs.Close
    End If

    'importador no grupo
    If wVinculo = 0 Then
        SQl = "select a.autonumlista,a.tipo"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
        SQl = SQl & " where"
        SQl = SQl & " a.autonum_parceiro = " & wImportador
        SQl = SQl & " and a.tipo='I'"
        PRSet Rs, SQl
        Do While Not Rs.EOF And wVinculo = 0
            SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
            SQl = SQl & " flag_liberada=1"
            SQl = SQl & " and nvl(flag_acordo,0)=0"
            SQl = SQl & " and autonum=" & Rs!autonumlista
            SQl = SQl & " and nvl(despachante,0)=0"
            SQl = SQl & " and nvl(captador,0)=0"
            SQl = SQl & " and nvl(coloader,0)=0"
            SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
            SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
            PRSet Rs3, SQl
            If Not Rs3.EOF Then
                wVinculo = Rs!autonumlista
            End If
            Rs3.Close
            Rs.MoveNext
        Loop
        Rs.Close
    End If
    
    'captador
    If wVinculo = 0 Then
        SQl = "select a.autonum"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
        SQl = SQl & " where"
        SQl = SQl & " a.flag_liberada=1"
        SQl = SQl & " and nvl(a.coloader,0)=0"
        SQl = SQl & " and nvl(a.importador,0)=0"
        SQl = SQl & " and nvl(a.captador,0)=" & wCaptador
        SQl = SQl & " and nvl(a.despachante,0)=0"
        SQl = SQl & " and nvl(a.flag_acordo,0)=0"
        SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
        SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
        PRSet Rs, SQl
        If Not Rs.EOF Then
            wVinculo = Rs!autonum
        End If
        Rs.Close
    End If

    'captador no grupo
    If wVinculo = 0 Then
        SQl = "select a.autonumlista,a.tipo"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
        SQl = SQl & " where"
        SQl = SQl & " a.autonum_parceiro = " & wCaptador
        SQl = SQl & " and a.tipo='C'"
        PRSet Rs, SQl
        Do While Not Rs.EOF And wVinculo = 0
            SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
            SQl = SQl & " flag_liberada=1"
            SQl = SQl & " and nvl(flag_acordo,0)=0"
            SQl = SQl & " and autonum=" & Rs!autonumlista
            SQl = SQl & " and nvl(despachante,0)=0"
            SQl = SQl & " and nvl(importador,0)=0"
            SQl = SQl & " and nvl(coloader,0)=0"
            SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
            SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
            PRSet Rs3, SQl
            If Not Rs3.EOF Then
                wVinculo = Rs!autonumlista
            End If
            Rs3.Close
            Rs.MoveNext
        Loop
        Rs.Close
    End If

    'despachante
    If wVinculo = 0 Then
        SQl = "select a.autonum"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
        SQl = SQl & " where"
        SQl = SQl & " a.flag_liberada=1"
        SQl = SQl & " and nvl(a.coloader,0)=0"
        SQl = SQl & " and nvl(a.importador,0)=0"
        SQl = SQl & " and nvl(a.captador,0)=0"
        SQl = SQl & " and nvl(a.despachante,0)=" & wDespachante
        SQl = SQl & " and nvl(a.flag_acordo,0)=0"
        SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
        SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
        PRSet Rs, SQl
        If Not Rs.EOF Then
            wVinculo = Rs!autonum
        End If
        Rs.Close
    End If

    'despachante no grupo
    If wVinculo = 0 Then
        SQl = "select a.autonumlista,a.tipo"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
        SQl = SQl & " where"
        SQl = SQl & " a.autonum_parceiro = " & wDespachante
        SQl = SQl & " and a.tipo='D'"
        PRSet Rs, SQl
        Do While Not Rs.EOF And wVinculo = 0
            SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
            SQl = SQl & " flag_liberada=1"
            SQl = SQl & " and nvl(flag_acordo,0)=0"
            SQl = SQl & " and autonum=" & Rs!autonumlista
            SQl = SQl & " and nvl(captador,0)=0"
            SQl = SQl & " and nvl(importador,0)=0"
            SQl = SQl & " and nvl(coloader,0)=0"
            SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
            SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
            PRSet Rs3, SQl
            If Not Rs3.EOF Then
                wVinculo = Rs!autonumlista
            End If
            Rs3.Close
            Rs.MoveNext
        Loop
        Rs.Close
    End If
End If


Padrao:
If wVinculo = 0 Then
    
'    Dim RsPadrao As New ADODB.recordset
'    Sql = "select tabela_padrao from " & Banco_Sgipa & "tb_empresas where autonum = " & Empresa_BL
'    PRSet RsPadrao, Sql
'    If Not RsPadrao.EOF Then
'        wVinculo = RsPadrao!Tabela_Padrao
'    Else
'        MsgBox "Tabela padrao nao encontrada - Comunique o T.I."
'        wVinculo = 0
'    End If
'    RsPadrao.Close

    Dim RsPadrao As New ADODB.Recordset
    SQl = "select tabela_padrao from " & Banco_Sgipa & "tb_empresas where autonum = " & Codigo_Empresa
    PRSet RsPadrao, SQl
    If Not RsPadrao.EOF Then
        wVinculo = RsPadrao!Tabela_Padrao
    Else
        MsgBox "Tabela padrao nao encontrada - Comunique o T.I."
        wVinculo = 0
    End If
    RsPadrao.Close

End If
If wVinculo <> 0 Then
    
    SQl = "update " & Banco_Sgipa & "tb_bl set autonum_lista=" & wVinculo & " ,DT_TERMINO_AVERBA=sysdate where autonum=" & Lote
    If Atualizar Then Executa SQl
    
    'Atualiza historico de vinculos
    SQl = "insert into " & Banco_Sgipa & "tb_hist_vinculos ("
    SQl = SQl & "lote"
    SQl = SQl & ",data_vinculo"
    SQl = SQl & ",importador"
    SQl = SQl & ",despachante"
    SQl = SQl & ",nvocc"
    SQl = SQl & ",indicador"
    SQl = SQl & ",vinculo_anterior"
    SQl = SQl & ",vinculo_atual"
    SQl = SQl & ",tipo_vinculo"
    SQl = SQl & ",usuario"
    SQl = SQl & ") values ("
    SQl = SQl & Lote
    SQl = SQl & ",sysdate"
    SQl = SQl & ",'" & wImportador & "'"
    SQl = SQl & ",'" & wDespachante & "'"
    SQl = SQl & ",'" & wNvocc & "'"
    SQl = SQl & ",'" & wCaptador & "'"
    SQl = SQl & ",'" & wLista_Anterior & "'"
    SQl = SQl & ",'" & wVinculo & "'"
    SQl = SQl & ",'" & wTipo_Vinculo & "'"
    SQl = SQl & ",'" & Cod_Usuario & "'"
    SQl = SQl & ")"
'    executa sql  'z4
    
End If
'-----------------------------------------------------

Vincula_Tabela_Band_até_17092015 = wVinculo

End Function

Function Vincula_Tabela_Band(Lote As Long, Optional FCL_LCL As String = "", Optional Atualizar As Boolean = False, Optional GravaTeste As Boolean = False) As Long

'VINCULA TABELA DE PRECOS AO BL
'-----------------------------------------------------
Dim Data_Validade As Date
Dim Tbh As New ADODB.Recordset
Dim wImportador As Long
Dim wDespachante As Long
Dim wCaptador As Long
Dim wColoader As Long
Dim wCoColoader As Long
Dim wCoColoader2 As Long
Dim wNvocc As Long
Dim wVinculo As Long
Dim wBL As Long
Dim wViagem As String
Dim DtasIndeferida As Byte
Dim wHubPort As Byte
Dim Rs As New ADODB.Recordset
Dim Rs3 As New ADODB.Recordset
Dim RsPadrao As New ADODB.Recordset
Dim tbh1 As New ADODB.Recordset
Dim wLista As Long
Dim wCargaPatio As Byte
Dim wLista_Anterior As Long
Dim wTipo_Vinculo As String
Dim Empresa_BL As Long
Dim Coloader As Boolean
Dim CoColoader As Boolean
Dim CoColoader2 As Boolean

Coloader = False

wLista_Anterior = 0
wTipo_Vinculo = "A"

wBL = 0

SQl = "select isnull(a.flag_carga_patio,0) flag_carga_patio, isnull(a.autonum_lista,0) lista, isnull(a.flag_hubport,0) flag_hubport"
SQl = SQl & ", isnull(a.flag_dtas_indeferida,0) flag_dtas_indeferida, b.razao, isnull(a.bl_master,0) bl_master, a.viagem"
SQl = SQl & ", isnull(a.importador,0) importador,isnull(a.despachante,0) despachante ,isnull(a.captador,0) captador, a.codigo_empresa"
SQl = SQl & ", isnull(a.audit_av,0) audit_av, a.primeira_entrada, a.patio, a.tipo_documento, isnull(a.Flag_Redestinacao,0) as Redestinacao"
SQl = SQl & ", isnull(a.destino_final,0) destino_final, isnull(a.hubport_FREEHAND,0) hubport_freehand"
SQl = SQl & " from " & Banco_Sgipa & "tb_bl a left join  " & Banco_Sgipa & "tb_cad_parceiros b on a.captador = b.autonum where "
SQl = SQl & " a.autonum=" & Lote

PRSet Rs, SQl


Dim Rs2 As New ADODB.Recordset
SQl = "select autonum_lista from " & Banco_Sgipa & "tb_lp_lotes where lote=" & Lote
PRSet Rs2, SQl
If Not Rs2.EOF Then
    wTipo_Vinculo = "M"
    wVinculo = Rs2!AUTONUM_LISTA
    Rs2.Close
    GoTo Padrao
End If
Rs2.Close



SQl = "SELECT isnull(cod_empresa,0) cod_empresa FROM " & Banco_Operador & "TB_PATIOS WHERE autonum = " & Nnull(Rs!Patio, 0)
PRSet Rs3, SQl
Empresa_BL = Rs3!Cod_Empresa
Rs3.Close

'If Nnull(Rs!redestinacao, 0) = 1 Then
'   If Empresa_BL = 1 Then
'      Empresa_BL = 2
'   Else
'      Empresa_BL = 1
'   End If
'End If


'busca data de validade pela data da entrada do conteiner
ERRO = 1
SQl = "select min(dt_entrada) dt_entrada from " & Banco_Sgipa & "tb_cntr_bl a, " & Banco_Sgipa & "tb_amr_cntr_bl b"
SQl = SQl & " where a.autonum = b.cntr"
SQl = SQl & " and b.bl=" & Lote
PRSet Rs3, SQl
If Not Rs3.EOF Then
    If Not IsNull(Rs3!Dt_Entrada) Then
        Data_Validade = Format$(Rs3!Dt_Entrada, "dd/mm/yyyy hh:mm:ss")
        ERRO = 0
    End If
End If
Rs3.Close
If ERRO = 1 Then
    SQl = "select min(dt_entrada) dt_entrada from " & Banco_Sgipa & "tb_carga_solta where bl=" & Lote
    PRSet Rs3, SQl
    If Not Rs3.EOF Then
        If Not IsNull(Rs3!Dt_Entrada) Then
            Data_Validade = Format$(Rs3!Dt_Entrada, "dd/mm/yyyy hh:mm:ss")
            ERRO = 0
        End If
    End If
    Rs3.Close
End If
If ERRO = 1 Then
    Data_Validade = Format$(Now, "dd/mm/yyyy hh:mm:ss")
End If



'Data_Validade = "11/03/2017"

Dim wHubport_RO As Byte
wHubport_RO = 0

'OS HubPort
If Nnull(Rs!flag_hubport, 0) = 1 Then
    If Nnull(Rs!hubport_freehand, 0) = 0 Then wHubport_RO = 1
    
    Dim TabelaHP As Integer
    Dim TabelaHPRO As Integer
    Dim GrupoHP As String
    Dim DiasFree As Integer
    
    'busca a tabela hubport no cadastro do indicador
    SQl = "select tabela_hubport_ro,tabela_hubport,grupo_freetime_hubport from " & Banco_Sgipa & "tb_cad_parceiros where autonum = " & Rs!Captador
    TabelaHPRO = Nnull(db.Execute(SQl).Fields(0), 0)
    TabelaHP = Nnull(db.Execute(SQl).Fields(1), 0)
    GrupoHP = Nnull(db.Execute(SQl).Fields(2), 1)
    If TabelaHP <> 0 Or GrupoHP <> "" Then
        If Nnull(Rs!hubport_freehand, 0) = 1 Then
            wVinculo = TabelaHP
        Else
            wVinculo = TabelaHPRO
        End If
        
        'busca dias freetime conforme o destino
        SQl = "select free_time from sgipa.dbo.TB_GRUPO_FREETIME_HUBPORT where grupo='" & GrupoHP & "' and destino=" & Nnull(Rs!destino_Final, 0)
        DiasFree = Nnull(db.Execute(SQl).Fields(0), 0)
        If DiasFree = 0 Then
            SQl = "select free_time from sgipa.dbo.TB_GRUPO_FREETIME_HUBPORT where grupo='" & GrupoHP & "' and destino is null"
            DiasFree = Nnull(db.Execute(SQl).Fields(0), 0)
        End If
        If DiasFree <> 0 Then
            'atualiza free time da tabela
            SQl = "update sgipa.dbo.tb_bl set free_time_hubport = " & DiasFree & " where autonum = " & Lote
            db.Execute SQl
            'verifica se o periodo esta dentro do free time
            If DateDiff("d", Data_Validade, Now) > DiasFree Or TabelaHP = 0 Then
                
                If TabelaHP = 0 Then
                    SQl = "update sgipa.dbo.tb_bl set nao_cobrar_hubport=1 where autonum=" & Lote
                    db.Execute SQl
                End If
                
                If wHubport_RO = 0 Then
                    'se periodo > free-time busca tabela geral do indicador
                    SQl = "select autonum from sgipa.dbo.tb_listas_precos where"
                    SQl = SQl & " captador=" & Rs!Captador & " and"
                    SQl = SQl & " isnull(hubport,0)=0 and"
                    SQl = SQl & " isnull(importador,0)=0 and"
                    SQl = SQl & " isnull(despachante,0)=0 and"
                    SQl = SQl & " isnull(coloader,0)=0 and"
                    SQl = SQl & " isnull(cocoloader,0)=0"
                    SQl = SQl & " and isnull(cocoloader2,0)=0"
                    SQl = SQl & " and data_inicio <= convert(datetime,'" & Format(Data_Validade, "dd/mm/yyyy") & "',103)"
                    SQl = SQl & " and (final_validade >= convert(datetime,'" & Format(Data_Validade, "dd/mm/yyyy") & "',103) or final_validade is null)"
                    SQl = SQl & " and flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(flag_acordo,0)=0"
                    wVinculo = Nnull(db.Execute(SQl).Fields(0), 0)
                    If wVinculo <> 0 Then
                        Rs.Close
                        GoTo Padrao
                    End If
                End If
            Else
                Rs.Close
                GoTo Padrao
            End If
        Else
            If TabelaHP = 0 Then
                SQl = "update sgipa.dbo.tb_bl set nao_cobrar_hubport=1 where autonum=" & Lote
                db.Execute SQl
            End If
            
            'se free-time nao informado vincular a geral do indicador
            SQl = "select autonum from sgipa.dbo.tb_listas_precos where"
            SQl = SQl & " captador=" & Rs!Captador & " and"
            SQl = SQl & " isnull(hubport,0)=0 and"
            SQl = SQl & " isnull(importador,0)=0 and"
            SQl = SQl & " isnull(despachante,0)=0 and"
            SQl = SQl & " isnull(coloader,0)=0 and"
            SQl = SQl & " isnull(cocoloader,0)=0"
            SQl = SQl & " and isnull(cocoloader2,0)=0"
            SQl = SQl & " and data_inicio <= convert(datetime,'" & Format(Data_Validade, "dd/mm/yyyy") & "',103)"
            SQl = SQl & " and (final_validade >= convert(datetime,'" & Format(Data_Validade, "dd/mm/yyyy") & "',103) or final_validade is null)"
            SQl = SQl & " and flag_liberada=1     and cod_empresa=" & Empresa_BL & " and isnull(flag_acordo,0)=0"
            wVinculo = Nnull(db.Execute(SQl).Fields(0), 0)
            If wVinculo <> 0 Then
                Rs.Close
                GoTo Padrao
            End If
        End If
    End If
End If
'***********************
wVinculo = 0






SQl = "select autonum_lista from " & Banco_Sgipa & "tb_lp_lotes where lote=" & Nnull(Rs!BL_Master, 0)
PRSet Rs2, SQl
If Not Rs2.EOF Then
    wTipo_Vinculo = "M"
    wVinculo = Rs2!AUTONUM_LISTA
    Rs2.Close
    GoTo Padrao
End If
Rs2.Close


wLista_Anterior = Rs!Lista
wViagem = Nnull(Rs!Viagem, 1)
wImportador = Rs!Importador
wDespachante = Rs!Despachante
wCaptador = Rs!Captador
wBL = Lote
DtasIndeferida = Rs!flag_dtas_indeferida
wLista = Rs!Lista

wVinculo = 0


'Verifica Margem de acordo com local de atracacao
Dim wMargem As String
wMargem = ""
SQl = "select margem from " & Banco_Sgipa & "DTE_TB_VIAGENS b, " & Banco_Sgipa & "DTE_TB_ARMAZENS a"
SQl = SQl & " where"
SQl = SQl & " b.arm_atracacao = a.code"
SQl = SQl & " and b.viagem='" & wViagem & "'"
PRSet Rs, SQl
If Not Rs.EOF Then
    wMargem = UCase(Nnull(Rs!Margem, 1))
End If
Rs.Close
If wMargem = "" Then wMargem = "MDIR"
'------------------------------------------------


'Verifica coloader
'-----------------------
Dim BLf As Long
Dim BLM As Long
Dim BLs As Long
Dim BLC As Long
Dim BL_Pai As Long
BLf = Lote  'Unico ou Filhote
SQl = "select bl_master from " & Banco_Sgipa & "tb_bl where autonum=" & BLf
BL_Pai = Nnull(db.Execute(SQl).Fields(0), 0)
If BL_Pai <> 0 Then
    BLM = BL_Pai    'master
    SQl = "select bl_master from " & Banco_Sgipa & "tb_bl where autonum=" & BLM
    BL_Pai = Nnull(db.Execute(SQl).Fields(0), 0)
    If BL_Pai <> 0 Then
        BLC = BLf       'coloader
        BLs = BLM       'sub master
        BLM = BL_Pai    'sub master
    End If
End If

If BLC <> 0 Then
    SQl = "select importador from " & Banco_Sgipa & "tb_bl where autonum=" & BLs
    wColoader = Nnull(db.Execute(SQl).Fields(0), 0)
End If




'Regras NOVAS de Vinculo bandeirantes - 2o. COCOLOADER - LCL
'1-) COLOADER,COCOLOADER,2o.COCOLOADER,INDICADOR,DESPACHANE,IMPORTADOR
'2-) COLOADER,COCOLOADER,2o.COCOLOADER,INDICADOR,IMPORTADOR
'3-) COLOADER,COCOLOADER,2o.COCOLOADER,INDICADOR,DESPACHANTE
'4-) COLOADER,COCOLOADER,2o.COCOLOADER,INDICADOR

'Verifica 2o.cocoloader
'-----------------------
SQl = "select isnull(sm.importador,0) as cocoloader "
SQl = SQl & " from sgipa.dbo.tb_bl c"
SQl = SQl & " LEFT JOIN sgipa.dbo.tb_bl sm on c.bl_master = sm.autonum"
SQl = SQl & " LEFT JOIN sgipa.dbo.tb_bl sm2 on sm.bl_master = sm2.autonum"
SQl = SQl & " LEFT JOIN sgipa.dbo.tb_bl sm3 on sm2.bl_master = sm3.autonum"
SQl = SQl & " LEFT JOIN sgipa.dbo.tb_bl m on sm3.bl_master = m.autonum"
SQl = SQl & " where c.autonum=" & Lote & " and m.grau='M'"
wCoColoader2 = Nnull(db.Execute(SQl).Fields(0), 0)

If wCoColoader2 <> 0 Then

    SQl = "select isnull(sm2.importador,0) as cocoloader "
    SQl = SQl & " from sgipa.dbo.tb_bl c"
    SQl = SQl & " LEFT JOIN sgipa.dbo.tb_bl sm on c.bl_master = sm.autonum"
    SQl = SQl & " LEFT JOIN sgipa.dbo.tb_bl sm2 on sm.bl_master = sm2.autonum"
    SQl = SQl & " LEFT JOIN sgipa.dbo.tb_bl sm3 on sm2.bl_master = sm3.autonum"
    SQl = SQl & " LEFT JOIN sgipa.dbo.tb_bl m on sm3.bl_master = m.autonum"
    SQl = SQl & " where c.autonum=" & Lote & " and m.grau='M'"
    wCoColoader = Nnull(db.Execute(SQl).Fields(0), 0)

    SQl = "select isnull(sm3.importador,0) as coloader "
    SQl = SQl & " from sgipa.dbo.tb_bl c"
    SQl = SQl & " LEFT JOIN sgipa.dbo.tb_bl sm on c.bl_master = sm.autonum"
    SQl = SQl & " LEFT JOIN sgipa.dbo.tb_bl sm2 on sm.bl_master = sm2.autonum"
    SQl = SQl & " LEFT JOIN sgipa.dbo.tb_bl sm3 on sm2.bl_master = sm3.autonum"
    SQl = SQl & " LEFT JOIN sgipa.dbo.tb_bl m on sm3.bl_master = m.autonum"
    SQl = SQl & " where c.autonum=" & Lote & " and m.grau='M'"
    wColoader = Nnull(db.Execute(SQl).Fields(0), 0)


    '1-) COLOADER,COCOLOADER,2o.cocoloader, INDICADOR,DESPACHANTE,IMPORTADOR
    SQl = "select a.autonum"
    SQl = SQl & " from"
    SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
    SQl = SQl & " where"
    SQl = SQl & " a.flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(a.hubport,0)=0"
    SQl = SQl & " and a.coloader=" & wColoader
    SQl = SQl & " and a.cocoloader=" & wCoColoader
    SQl = SQl & " and a.cocoloader2=" & wCoColoader2
    SQl = SQl & " and a.captador=" & wCaptador
    SQl = SQl & " and a.despachante=" & wDespachante
    SQl = SQl & " and a.importador=" & wImportador
    SQl = SQl & " and isnull(a.flag_acordo,0)=0"
    SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
    SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
    
    PRSet Rs, SQl
    If Not Rs.EOF Then
        wVinculo = Rs!autonum
    End If
    Rs.Close

    'NO GRUPO
    If wVinculo = 0 Then
        SQl = "select a.autonumlista,a.tipo"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
        SQl = SQl & " where"
        SQl = SQl & " a.autonum_parceiro = " & wCoColoader2
        SQl = SQl & " and a.tipo='2'"
        PRSet Rs, SQl
        Do While Not Rs.EOF And wVinculo = 0
            Tipo = Rs!Tipo
            wVinculo = Rs!autonumlista
            SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
            SQl = SQl & " flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(hubport,0)=0"
            SQl = SQl & " and isnull(flag_acordo,0)=0"
            SQl = SQl & " and autonum=" & wVinculo
            SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
            SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
            
            
          '  If wHubport_RO = 1 Then
          '      Sql = Sql & " and isnull(tb_listas_precos.hubport,0)=0"
          '  End If
            
            
            PRSet Rs3, SQl
            If Not Rs3.EOF Then
                If Nnull(Rs3!CoColoader2, 0) = 0 Or Nnull(Rs3!CoColoader, 0) = 0 Or Nnull(Rs3!Coloader, 0) = 0 Or Nnull(Rs3!Importador, 0) = 0 Or Nnull(Rs3!Despachante, 0) = 0 Or Nnull(Rs3!Captador, 0) = 0 Then wVinculo = 0
                If wVinculo <> 0 And wCoColoader <> 0 Then
                    'Verifica se cocoloader faz parte do grupo da tabela
                    SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='F' and autonum_parceiro=" & wCoColoader
                    If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                End If
                If wVinculo <> 0 And wColoader <> 0 Then
                    'Verifica se coloader faz parte do grupo da tabela
                    SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='S' and autonum_parceiro=" & wColoader
                    If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                End If
                If wVinculo <> 0 And wCaptador <> 0 Then
                    'Verifica se indicador faz parte do grupo da tabela
                    SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='C' and autonum_parceiro=" & wCaptador
                    If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                End If
                If wVinculo <> 0 And wDespachante <> 0 Then
                    'Verifica se despachante faz parte do grupo da tabela
                    SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='D' and autonum_parceiro=" & wDespachante
                    If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                End If
                If wVinculo <> 0 And wImportador <> 0 Then
                    'Verifica se importador faz parte do grupo da tabela
                    SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='I' and autonum_parceiro=" & wImportador
                    If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                End If
            Else
                wVinculo = 0
            End If
            Rs3.Close
            Rs.MoveNext
            If wVinculo <> 0 Then Exit Do
        Loop
        Rs.Close
    End If


    '2-) COLOADER,COCOLOADER,2o.cocoloader, INDICADOR,IMPORTADOR
    If wVinculo = 0 Then
        SQl = "select a.autonum"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
        SQl = SQl & " where"
        SQl = SQl & " a.flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(a.hubport,0)=0"
        SQl = SQl & " and a.coloader=" & wColoader
        SQl = SQl & " and a.cocoloader=" & wCoColoader
        SQl = SQl & " and a.cocoloader2=" & wCoColoader2
        SQl = SQl & " and isnull(a.captador,0)=" & wCaptador
        SQl = SQl & " and isnull(a.importador,0)=" & wImportador
        SQl = SQl & " and isnull(a.despachante,0)=0"
        SQl = SQl & " and isnull(a.flag_acordo,0)=0"
        SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
        SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
        PRSet Rs, SQl
        If Not Rs.EOF Then
            wVinculo = Rs!autonum
        End If
        Rs.Close

        'NO GRUPO
        If wVinculo = 0 Then
            SQl = "select a.autonumlista,a.tipo"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
            SQl = SQl & " where"
            SQl = SQl & " a.autonum_parceiro = " & wCoColoader2
            SQl = SQl & " and a.tipo='2'"
            PRSet Rs, SQl
            Do While Not Rs.EOF And wVinculo = 0
                Tipo = Rs!Tipo
                wVinculo = Rs!autonumlista
                SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
                SQl = SQl & " flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(hubport,0)=0"
                SQl = SQl & " and isnull(flag_acordo,0)=0"
                SQl = SQl & " and autonum=" & wVinculo
                SQl = SQl & " and isnull(despachante,0)=0"
                SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
                SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
                PRSet Rs3, SQl
                If Not Rs3.EOF Then
                    If Nnull(Rs3!CoColoader2, 0) = 0 Or Nnull(Rs3!CoColoader, 0) = 0 Or Nnull(Rs3!Coloader, 0) = 0 Or Nnull(Rs3!Importador, 0) = 0 Or Nnull(Rs3!Captador, 0) = 0 Then wVinculo = 0
                    If wVinculo <> 0 And wCoColoader <> 0 Then
                        'Verifica se cocoloader faz parte do grupo da tabela
                        SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='F' and autonum_parceiro=" & wCoColoader
                        If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                    End If
                    If wVinculo <> 0 And wColoader <> 0 Then
                        'Verifica se coloader faz parte do grupo da tabela
                        SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='S' and autonum_parceiro=" & wColoader
                        If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                    End If
                    If wVinculo <> 0 And wCaptador <> 0 Then
                        'Verifica se indicador faz parte do grupo da tabela
                        SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='C' and autonum_parceiro=" & wCaptador
                        If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                    End If
                    If wVinculo <> 0 And wImportador <> 0 Then
                        'Verifica se importador faz parte do grupo da tabela
                        SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='I' and autonum_parceiro=" & wImportador
                        If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                    End If
                Else
                    wVinculo = 0
                End If
                Rs3.Close
                Rs.MoveNext
                If wVinculo <> 0 Then Exit Do
            Loop
            Rs.Close
        End If

    End If

    '3-) COLOADER,COCOLOADER,2o. cocoloader,INDICADOR,DESPACHANTE
    If wVinculo = 0 Then
        SQl = "select a.autonum"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
        SQl = SQl & " where"
        SQl = SQl & " a.flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(a.hubport,0)=0"
        SQl = SQl & " and a.coloader=" & wColoader
        SQl = SQl & " and a.cocoloader=" & wCoColoader
        SQl = SQl & " and a.cocoloader2=" & wCoColoader2
        SQl = SQl & " and isnull(a.captador,0)=" & wCaptador
        SQl = SQl & " and isnull(a.despachante,0)=" & wDespachante
        SQl = SQl & " and isnull(a.importador,0)=0"
        SQl = SQl & " and isnull(a.flag_acordo,0)=0"
        SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
        SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
        PRSet Rs, SQl
        If Not Rs.EOF Then
            wVinculo = Rs!autonum
        End If
        Rs.Close

        'NO GRUPO
        If wVinculo = 0 Then
            SQl = "select a.autonumlista,a.tipo"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
            SQl = SQl & " where"
            SQl = SQl & " a.autonum_parceiro = " & wCoColoader2
            SQl = SQl & " and a.tipo='2'"
            PRSet Rs, SQl
            Do While Not Rs.EOF And wVinculo = 0
                Tipo = Rs!Tipo
                wVinculo = Rs!autonumlista
                SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
                SQl = SQl & " flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(hubport,0)=0"
                SQl = SQl & " and isnull(flag_acordo,0)=0"
                SQl = SQl & " and autonum=" & wVinculo
                SQl = SQl & " and isnull(importador,0)=0"
                SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
                SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
                PRSet Rs3, SQl
                If Not Rs3.EOF Then
                    If Nnull(Rs3!CoColoader2, 0) = 0 Or Nnull(Rs3!CoColoader, 0) = 0 Or Nnull(Rs3!Coloader, 0) = 0 Or Nnull(Rs3!Despachante, 0) = 0 Or Nnull(Rs3!Captador, 0) = 0 Then wVinculo = 0
                    If wVinculo <> 0 And wCoColoader <> 0 Then
                        'Verifica se coloader faz parte do grupo da tabela
                        SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='F' and autonum_parceiro=" & wCoColoader
                        If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                    End If
                    If wVinculo <> 0 And wColoader <> 0 Then
                        'Verifica se coloader faz parte do grupo da tabela
                        SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='S' and autonum_parceiro=" & wColoader
                        If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                    End If
                    If wVinculo <> 0 And wCaptador <> 0 Then
                        'Verifica se indicador faz parte do grupo da tabela
                        SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='C' and autonum_parceiro=" & wCaptador
                        If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                    End If
                    If wVinculo <> 0 And wDespachante <> 0 Then
                        'Verifica se despachante faz parte do grupo da tabela
                        SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='D' and autonum_parceiro=" & wDespachante
                        If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                    End If
                Else
                    wVinculo = 0
                End If
                Rs3.Close
                Rs.MoveNext
                If wVinculo <> 0 Then Exit Do
            Loop
            Rs.Close
        End If

    End If


    '4-) COLOADER,COCOLOADER,2o.cocoloader,INDICADOR
    If wVinculo = 0 Then
        SQl = "select a.autonum"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
        SQl = SQl & " where"
        SQl = SQl & " a.flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(a.hubport,0)=0"
        SQl = SQl & " and a.coloader=" & wColoader
        SQl = SQl & " and a.cocoloader=" & wCoColoader
        SQl = SQl & " and a.cocoloader2=" & wCoColoader2
        SQl = SQl & " and isnull(a.captador,0)=" & wCaptador
        SQl = SQl & " and isnull(a.despachante,0)=0"
        SQl = SQl & " and isnull(a.importador,0)=0"
        SQl = SQl & " and isnull(a.flag_acordo,0)=0"
        SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
        SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
        PRSet Rs, SQl
        If Not Rs.EOF Then
            wVinculo = Rs!autonum
        End If
        Rs.Close

        'NO GRUPO
        If wVinculo = 0 Then
            SQl = "select a.autonumlista,a.tipo"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
            SQl = SQl & " where"
            SQl = SQl & " a.autonum_parceiro = " & wCoColoader2
            SQl = SQl & " and a.tipo='2'"
            PRSet Rs, SQl
            Do While Not Rs.EOF And wVinculo = 0
                Tipo = Rs!Tipo
                wVinculo = Rs!autonumlista
                SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
                SQl = SQl & " flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(hubport,0)=0"
                SQl = SQl & " and isnull(flag_acordo,0)=0"
                SQl = SQl & " and autonum=" & wVinculo
                SQl = SQl & " and isnull(importador,0)=0"
                SQl = SQl & " and isnull(despachante,0)=0"
                SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
                SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
                PRSet Rs3, SQl
                If Not Rs3.EOF Then
                    If Nnull(Rs3!CoColoader2, 0) = 0 Or Nnull(Rs3!CoColoader, 0) = 0 Or Nnull(Rs3!Coloader, 0) = 0 Or Nnull(Rs3!Captador, 0) = 0 Then wVinculo = 0
                    If wVinculo <> 0 And wCoColoader <> 0 Then
                        'Verifica se coloader faz parte do grupo da tabela
                        SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='F' and autonum_parceiro=" & wCoColoader
                        If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                    End If
                    If wVinculo <> 0 And wColoader <> 0 Then
                        'Verifica se coloader faz parte do grupo da tabela
                        SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='S' and autonum_parceiro=" & wColoader
                        If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                    End If
                    If wVinculo <> 0 And wCaptador <> 0 Then
                        'Verifica se indicador faz parte do grupo da tabela
                        SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='C' and autonum_parceiro=" & wCaptador
                        If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                    End If
                Else
                    wVinculo = 0
                End If
                Rs3.Close
                Rs.MoveNext
                If wVinculo <> 0 Then Exit Do
            Loop
            Rs.Close
        End If



    End If

End If



If wVinculo = 0 Then
    
    'Regras NOVAS de Vinculo bandeirantes - 1o. COCOLOADER - LCL
    '1-) COLOADER,COCOLOADER,INDICADOR,DESPACHANE,IMPORTADOR
    '2-) COLOADER,COCOLOADER,INDICADOR,IMPORTADOR
    '3-) COLOADER,COCOLOADER,INDICADOR,DESPACHANTE
    '4-) COLOADER,COCOLOADER,INDICADOR
    
    'Verifica cocoloader
    '-----------------------
    SQl = "select isnull(sm.importador,0) as cocoloader "
    SQl = SQl & " from sgipa.dbo.tb_bl c"
    SQl = SQl & " LEFT JOIN sgipa.dbo.tb_bl sm on c.bl_master = sm.autonum"
    SQl = SQl & " LEFT JOIN sgipa.dbo.tb_bl sm2 on sm.bl_master = sm2.autonum"
    SQl = SQl & " LEFT JOIN sgipa.dbo.tb_bl m on sm2.bl_master = m.autonum"
    SQl = SQl & " where c.autonum=" & Lote & " and m.grau='M'"
    wCoColoader = Nnull(db.Execute(SQl).Fields(0), 0)
    
    If wCoColoader <> 0 Then
    
        SQl = "select isnull(sm2.importador,0) as cocoloader "
        SQl = SQl & " from sgipa.dbo.tb_bl c"
        SQl = SQl & " LEFT JOIN sgipa.dbo.tb_bl sm on c.bl_master = sm.autonum"
        SQl = SQl & " LEFT JOIN sgipa.dbo.tb_bl sm2 on sm.bl_master = sm2.autonum"
        SQl = SQl & " LEFT JOIN sgipa.dbo.tb_bl m on sm2.bl_master = m.autonum"
        SQl = SQl & " where c.autonum=" & Lote & " and m.grau='M'"
        wColoader = Nnull(db.Execute(SQl).Fields(0), 0)
    
        '1-) COLOADER,COCOLOADER,INDICADOR,DESPACHANTE,IMPORTADOR
        SQl = "select a.autonum"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
        SQl = SQl & " where"
        SQl = SQl & " a.flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(a.hubport,0)=0"
        SQl = SQl & " and a.coloader=" & wColoader
        SQl = SQl & " and a.cocoloader=" & wCoColoader
        SQl = SQl & " and a.captador=" & wCaptador
        SQl = SQl & " and a.despachante=" & wDespachante
        SQl = SQl & " and isnull(a.cocoloader2,0)=0"
        SQl = SQl & " and a.importador=" & wImportador
        SQl = SQl & " and isnull(a.flag_acordo,0)=0"
        SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
        SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
        PRSet Rs, SQl
        If Not Rs.EOF Then
            wVinculo = Rs!autonum
        End If
        Rs.Close
    
        'NO GRUPO
        If wVinculo = 0 Then
            SQl = "select a.autonumlista,a.tipo"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
            SQl = SQl & " where"
            SQl = SQl & " a.autonum_parceiro = " & wCoColoader
            SQl = SQl & " and a.tipo='F'"
            PRSet Rs, SQl
            Do While Not Rs.EOF And wVinculo = 0
                Tipo = Rs!Tipo
                wVinculo = Rs!autonumlista
                SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
                SQl = SQl & " flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(hubport,0)=0"
                SQl = SQl & " and isnull(flag_acordo,0)=0"
                SQl = SQl & " and isnull(cocoloader2,0)=0"
                SQl = SQl & " and autonum=" & wVinculo
                SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
                SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
                PRSet Rs3, SQl
                If Not Rs3.EOF Then
                    If Nnull(Rs3!CoColoader, 0) = 0 Or Nnull(Rs3!Coloader, 0) = 0 Or Nnull(Rs3!Importador, 0) = 0 Or Nnull(Rs3!Despachante, 0) = 0 Or Nnull(Rs3!Captador, 0) = 0 Then wVinculo = 0
                    If wVinculo <> 0 And wColoader <> 0 Then
                        'Verifica se coloader faz parte do grupo da tabela
                        SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='S' and autonum_parceiro=" & wColoader
                        If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                    End If
                    If wVinculo <> 0 And wCaptador <> 0 Then
                        'Verifica se indicador faz parte do grupo da tabela
                        SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='C' and autonum_parceiro=" & wCaptador
                        If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                    End If
                    If wVinculo <> 0 And wDespachante <> 0 Then
                        'Verifica se despachante faz parte do grupo da tabela
                        SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='D' and autonum_parceiro=" & wDespachante
                        If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                    End If
                    If wVinculo <> 0 And wImportador <> 0 Then
                        'Verifica se importador faz parte do grupo da tabela
                        SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='I' and autonum_parceiro=" & wImportador
                        If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                    End If
                Else
                    wVinculo = 0
                End If
                Rs3.Close
                Rs.MoveNext
                If wVinculo <> 0 Then Exit Do
            Loop
            Rs.Close
        End If
    
    
        '2-) COLOADER,COCOLOADER,INDICADOR,IMPORTADOR
        If wVinculo = 0 Then
            SQl = "select a.autonum"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
            SQl = SQl & " where"
            SQl = SQl & " a.flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(a.hubport,0)=0"
            SQl = SQl & " and a.coloader=" & wColoader
            SQl = SQl & " and a.cocoloader=" & wCoColoader
            SQl = SQl & " and isnull(a.captador,0)=" & wCaptador
            SQl = SQl & " and isnull(a.importador,0)=" & wImportador
            SQl = SQl & " and isnull(a.despachante,0)=0"
            SQl = SQl & " and isnull(a.cocoloader2,0)=0"
            SQl = SQl & " and isnull(a.flag_acordo,0)=0"
            SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
            SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
            PRSet Rs, SQl
            If Not Rs.EOF Then
                wVinculo = Rs!autonum
            End If
            Rs.Close
    
            'NO GRUPO
            If wVinculo = 0 Then
                SQl = "select a.autonumlista,a.tipo"
                SQl = SQl & " from"
                SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
                SQl = SQl & " where"
                SQl = SQl & " a.autonum_parceiro = " & wCoColoader
                SQl = SQl & " and a.tipo='F'"
                PRSet Rs, SQl
                Do While Not Rs.EOF And wVinculo = 0
                    Tipo = Rs!Tipo
                    wVinculo = Rs!autonumlista
                    SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
                    SQl = SQl & " flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(hubport,0)=0"
                    SQl = SQl & " and isnull(flag_acordo,0)=0"
                    SQl = SQl & " and autonum=" & wVinculo
                    SQl = SQl & " and isnull(despachante,0)=0"
                    SQl = SQl & " and isnull(cocoloader2,0)=0"
                    SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
                    SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
                    PRSet Rs3, SQl
                    If Not Rs3.EOF Then
                        If Nnull(Rs3!CoColoader, 0) = 0 Or Nnull(Rs3!Coloader, 0) = 0 Or Nnull(Rs3!Importador, 0) = 0 Or Nnull(Rs3!Captador, 0) = 0 Then wVinculo = 0
                        If wVinculo <> 0 And wColoader <> 0 Then
                            'Verifica se coloader faz parte do grupo da tabela
                            SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='S' and autonum_parceiro=" & wColoader
                            If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                        End If
                        If wVinculo <> 0 And wCaptador <> 0 Then
                            'Verifica se indicador faz parte do grupo da tabela
                            SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='C' and autonum_parceiro=" & wCaptador
                            If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                        End If
                        If wVinculo <> 0 And wImportador <> 0 Then
                            'Verifica se importador faz parte do grupo da tabela
                            SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='I' and autonum_parceiro=" & wImportador
                            If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                        End If
                    Else
                        wVinculo = 0
                    End If
                    Rs3.Close
                    Rs.MoveNext
                    If wVinculo <> 0 Then Exit Do
                Loop
                Rs.Close
            End If
            
        End If
    
        '3-) COLOADER,COCOLOADER,INDICADOR,DESPACHANTE
        If wVinculo = 0 Then
            SQl = "select a.autonum"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
            SQl = SQl & " where"
            SQl = SQl & " a.flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(a.hubport,0)=0"
            SQl = SQl & " and a.coloader=" & wColoader
            SQl = SQl & " and a.cocoloader=" & wCoColoader
            SQl = SQl & " and isnull(a.captador,0)=" & wCaptador
            SQl = SQl & " and isnull(a.despachante,0)=" & wDespachante
            SQl = SQl & " and isnull(a.importador,0)=0"
            SQl = SQl & " and isnull(a.cocoloader2,0)=0"
            SQl = SQl & " and isnull(a.flag_acordo,0)=0"
            SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
            SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
            PRSet Rs, SQl
            If Not Rs.EOF Then
                wVinculo = Rs!autonum
            End If
            Rs.Close
        
            'NO GRUPO
            If wVinculo = 0 Then
                SQl = "select a.autonumlista,a.tipo"
                SQl = SQl & " from"
                SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
                SQl = SQl & " where"
                SQl = SQl & " a.autonum_parceiro = " & wCoColoader
                SQl = SQl & " and a.tipo='F'"
                PRSet Rs, SQl
                Do While Not Rs.EOF And wVinculo = 0
                    Tipo = Rs!Tipo
                    wVinculo = Rs!autonumlista
                    SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
                    SQl = SQl & " flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(hubport,0)=0"
                    SQl = SQl & " and isnull(flag_acordo,0)=0"
                    SQl = SQl & " and autonum=" & wVinculo
                    SQl = SQl & " and isnull(importador,0)=0"
                    SQl = SQl & " and isnull(cocoloader2,0)=0"
                    SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
                    SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
                    PRSet Rs3, SQl
                    If Not Rs3.EOF Then
                        If Nnull(Rs3!CoColoader, 0) = 0 Or Nnull(Rs3!Coloader, 0) = 0 Or Nnull(Rs3!Despachante, 0) = 0 Or Nnull(Rs3!Captador, 0) = 0 Then wVinculo = 0
                        If wVinculo <> 0 And wColoader <> 0 Then
                            'Verifica se coloader faz parte do grupo da tabela
                            SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='S' and autonum_parceiro=" & wColoader
                            If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                        End If
                        If wVinculo <> 0 And wCaptador <> 0 Then
                            'Verifica se indicador faz parte do grupo da tabela
                            SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='C' and autonum_parceiro=" & wCaptador
                            If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                        End If
                        If wVinculo <> 0 And wDespachante <> 0 Then
                            'Verifica se despachante faz parte do grupo da tabela
                            SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='D' and autonum_parceiro=" & wDespachante
                            If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                        End If
                    Else
                        wVinculo = 0
                    End If
                    Rs3.Close
                    Rs.MoveNext
                    If wVinculo <> 0 Then Exit Do
                Loop
                Rs.Close
            End If
            
        End If
        
        
        '4-) COLOADER,COCOLOADER,INDICADOR
        If wVinculo = 0 Then
            SQl = "select a.autonum"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
            SQl = SQl & " where"
            SQl = SQl & " a.flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(a.hubport,0)=0"
            SQl = SQl & " and a.coloader=" & wColoader
            SQl = SQl & " and a.cocoloader=" & wCoColoader
            SQl = SQl & " and isnull(a.captador,0)=" & wCaptador
            SQl = SQl & " and isnull(a.despachante,0)=0"
            SQl = SQl & " and isnull(a.importador,0)=0"
            SQl = SQl & " and isnull(a.cocoloader2,0)=0"
            SQl = SQl & " and isnull(a.flag_acordo,0)=0"
            SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
            SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
            PRSet Rs, SQl
            If Not Rs.EOF Then
                wVinculo = Rs!autonum
            End If
            Rs.Close
    
            'NO GRUPO
            If wVinculo = 0 Then
                SQl = "select a.autonumlista,a.tipo"
                SQl = SQl & " from"
                SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
                SQl = SQl & " where"
                SQl = SQl & " a.autonum_parceiro = " & wCoColoader
                SQl = SQl & " and a.tipo='F'"
                PRSet Rs, SQl
                Do While Not Rs.EOF And wVinculo = 0
                    Tipo = Rs!Tipo
                    wVinculo = Rs!autonumlista
                    SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
                    SQl = SQl & " flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(hubport,0)=0"
                    SQl = SQl & " and isnull(flag_acordo,0)=0"
                    SQl = SQl & " and autonum=" & wVinculo
                    SQl = SQl & " and isnull(importador,0)=0"
                    SQl = SQl & " and isnull(despachante,0)=0"
                    SQl = SQl & " and isnull(cocoloader2,0)=0"
                    SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
                    SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
                    PRSet Rs3, SQl
                    If Not Rs3.EOF Then
                        If Nnull(Rs3!CoColoader, 0) = 0 Or Nnull(Rs3!Coloader, 0) = 0 Or Nnull(Rs3!Captador, 0) = 0 Then wVinculo = 0
                        If wVinculo <> 0 And wColoader <> 0 Then
                            'Verifica se coloader faz parte do grupo da tabela
                            SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='S' and autonum_parceiro=" & wColoader
                            If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                        End If
                        If wVinculo <> 0 And wCaptador <> 0 Then
                            'Verifica se indicador faz parte do grupo da tabela
                            SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='C' and autonum_parceiro=" & wCaptador
                            If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                        End If
                    Else
                        wVinculo = 0
                    End If
                    Rs3.Close
                    Rs.MoveNext
                    If wVinculo <> 0 Then Exit Do
                Loop
                Rs.Close
            End If
        
        
        
        End If
    
    End If

End If

If wVinculo = 0 Then

    If BLC <> 0 Then
'        sql = "select importador from " & Banco_Sgipa & "tb_bl where autonum=" & BLs
'        wColoader = Nnull(db.Execute(sql).Fields(0), 0)
        
        'procura tabela especifica  (nao olha o grupo)
        'coloader
        SQl = "select a.autonum"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
        SQl = SQl & " where"
        SQl = SQl & " a.flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(a.hubport,0)=0"
        SQl = SQl & " and a.coloader=" & wColoader
        SQl = SQl & " and isnull(a.importador,0)=0"
        SQl = SQl & " and isnull(a.despachante,0)=0"
        SQl = SQl & " and isnull(a.captador,0)=0"
        SQl = SQl & " and isnull(a.flag_acordo,0)=0"
        SQl = SQl & " and isnull(a.cocoloader,0)=0"
        SQl = SQl & " and isnull(a.cocoloader2,0)=0"
        SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
        SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
        PRSet Rs, SQl
        If Not Rs.EOF Then
            wVinculo = Rs!autonum
        End If
        Rs.Close
    
        If wVinculo = 0 Then
            SQl = "select a.autonumlista,a.tipo"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
            SQl = SQl & " where"
            SQl = SQl & " a.autonum_parceiro = " & wColoader
            SQl = SQl & " and a.tipo='S'"
            PRSet Rs, SQl
            Do While Not Rs.EOF And wVinculo = 0
                Tipo = Rs!Tipo
                wVinculo = Rs!autonumlista
                
                SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
                SQl = SQl & " flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(hubport,0)=0"
                SQl = SQl & " and isnull(flag_acordo,0)=0"
                SQl = SQl & " and isnull(cocoloader,0)=0"
                SQl = SQl & " and isnull(cocoloader2,0)=0"
                SQl = SQl & " and autonum=" & wVinculo
                SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
                SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
                PRSet Rs3, SQl
                If Not Rs3.EOF Then
                    If Nnull(Rs3!Importador, 0) <> 0 Or Nnull(Rs3!Despachante, 0) <> 0 Or Nnull(Rs3!Captador, 0) <> 0 Then wVinculo = 0
                Else
                    wVinculo = 0
                End If
                Rs3.Close
                    
                Rs.MoveNext
            Loop
            Rs.Close
        End If
    
        If wVinculo = 0 Then            'coloader + importador + captador
            SQl = "select a.autonum"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
            SQl = SQl & " where"
            SQl = SQl & " a.flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(a.hubport,0)=0"
            SQl = SQl & " and a.coloader=" & wColoader
            SQl = SQl & " and isnull(a.importador,0)=" & wImportador
            SQl = SQl & " and isnull(a.captador,0)=" & wCaptador
            SQl = SQl & " and isnull(a.despachante,0)=0"
            SQl = SQl & " and isnull(a.cocoloader,0)=0"
            SQl = SQl & " and isnull(a.cocoloader2,0)=0"
            SQl = SQl & " and isnull(a.flag_acordo,0)=0"
            SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
            SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
            PRSet Rs, SQl
            If Not Rs.EOF Then
                wVinculo = Rs!autonum
            End If
            Rs.Close
    
            If wVinculo = 0 Then
                SQl = "select a.autonumlista,a.tipo"
                SQl = SQl & " from"
                SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
                SQl = SQl & " where"
                SQl = SQl & " a.autonum_parceiro = " & wColoader
                SQl = SQl & " and a.tipo='S'"
                PRSet Rs, SQl
                Do While Not Rs.EOF And wVinculo = 0
                    wVinculo = Rs!autonumlista
                    
                    SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
                    SQl = SQl & " flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(hubport,0)=0"
                    SQl = SQl & " and isnull(flag_acordo,0)=0"
                    SQl = SQl & " and autonum=" & wVinculo
                    SQl = SQl & " and isnull(despachante,0)=0"
                    SQl = SQl & " and isnull(cocoloader,0)=0"
                    SQl = SQl & " and isnull(cocoloader2,0)=0"
                    SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
                    SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
                    PRSet Rs3, SQl
                    If Not Rs3.EOF Then
                       If Nnull(Rs3!Coloader, 0) = 0 Or Nnull(Rs3!Importador, 0) = 0 Or Nnull(Rs3!Captador, 0) = 0 Then wVinculo = 0
                       If wImportador <> Nnull(Rs3!Importador, 0) And wVinculo > 0 Then
                          'Verifica se importador faz parte do grupo da tabela
                          SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='I' and autonum_parceiro=" & wImportador
                          If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                       End If
                       If wVinculo <> 0 And wCaptador <> Nnull(Rs3!Captador, 0) Then
                           'Verifica se indicador faz parte do grupo da tabela
                           SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='C' and autonum_parceiro=" & wCaptador
                           If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                       End If
                   Else
                      wVinculo = 0
                   End If
                   Rs3.Close
                   Rs.MoveNext
                Loop
                Rs.Close
            End If
        End If
    
        If wVinculo = 0 Then            'coloader + importador
            SQl = "select a.autonum"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
            SQl = SQl & " where"
            SQl = SQl & " a.flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(a.hubport,0)=0"
            SQl = SQl & " and a.coloader=" & wColoader
            SQl = SQl & " and isnull(a.importador,0)=" & wImportador
            SQl = SQl & " and isnull(a.despachante,0)=0"
            SQl = SQl & " and isnull(a.captador,0)=0"
            SQl = SQl & " and isnull(a.cocoloader,0)=0"
            SQl = SQl & " and isnull(a.cocoloader2,0)=0"
            SQl = SQl & " and isnull(a.flag_acordo,0)=0"
            SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
            SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
            PRSet Rs, SQl
            If Not Rs.EOF Then
                wVinculo = Rs!autonum
            End If
            Rs.Close
    
            If wVinculo = 0 Then
                SQl = "select a.autonumlista,a.tipo"
                SQl = SQl & " from"
                SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
                SQl = SQl & " where"
                SQl = SQl & " a.autonum_parceiro = " & wColoader
                SQl = SQl & " and a.tipo='S'"
                PRSet Rs, SQl
                Do While Not Rs.EOF And wVinculo = 0
                    wVinculo = Rs!autonumlista
                    
                    SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
                    SQl = SQl & " flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(hubport,0)=0"
                    SQl = SQl & " and isnull(flag_acordo,0)=0"
                    SQl = SQl & " and autonum=" & wVinculo
                    SQl = SQl & " and isnull(despachante,0)=0"
                    SQl = SQl & " and isnull(captador,0)=0"
                    SQl = SQl & " and isnull(cocoloader,0)=0"
                    SQl = SQl & " and isnull(cocoloader2,0)=0"
                    SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
                    SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
                    PRSet Rs3, SQl
                    If Not Rs3.EOF Then
                       If Nnull(Rs3!Coloader, 0) = 0 Or Nnull(Rs3!Importador, 0) = 0 Then wVinculo = 0
                       If wImportador <> Nnull(Rs3!Importador, 0) And wVinculo > 0 Then
                          'Verifica se importador faz parte do grupo da tabela
                          SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='I' and autonum_parceiro=" & wImportador
                          If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                       End If
                    Else
                        wVinculo = 0
                    End If
                    Rs3.Close
                        
                    Rs.MoveNext
                Loop
                Rs.Close
            End If
        End If
    
    
        If wVinculo = 0 Then            'coloader + indicador
            SQl = "select a.autonum"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
            SQl = SQl & " where"
            SQl = SQl & " a.flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(a.hubport,0)=0"
            SQl = SQl & " and a.coloader=" & wColoader
            SQl = SQl & " and isnull(a.captador,0)=" & wCaptador
            SQl = SQl & " and isnull(a.despachante,0)=0"
            SQl = SQl & " and isnull(a.importador,0)=0"
            SQl = SQl & " and isnull(a.cocoloader,0)=0"
            SQl = SQl & " and isnull(a.cocoloader2,0)=0"
            SQl = SQl & " and isnull(a.flag_acordo,0)=0"
            SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
            SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
            PRSet Rs, SQl
            If Not Rs.EOF Then
                wVinculo = Rs!autonum
            End If
            Rs.Close
    
            If wVinculo = 0 Then
                SQl = "select a.autonumlista,a.tipo"
                SQl = SQl & " from"
                SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
                SQl = SQl & " where"
                SQl = SQl & " a.autonum_parceiro = " & wColoader
                SQl = SQl & " and a.tipo='S'"
                PRSet Rs, SQl
                Do While Not Rs.EOF And wVinculo = 0
                    wVinculo = Rs!autonumlista
                    
                    SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
                    SQl = SQl & " flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(hubport,0)=0"
                    SQl = SQl & " and isnull(flag_acordo,0)=0"
                    SQl = SQl & " and autonum=" & wVinculo
                    SQl = SQl & " and isnull(despachante,0)=0"
                    SQl = SQl & " and isnull(importador,0)=0"
                    SQl = SQl & " and isnull(cocoloader,0)=0"
                    SQl = SQl & " and isnull(cocoloader2,0)=0"
                    SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
                    SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
                    PRSet Rs3, SQl
                    If Not Rs3.EOF Then
                       If Nnull(Rs3!Coloader, 0) = 0 Or Nnull(Rs3!Captador, 0) = 0 Then wVinculo = 0
                       If wVinculo <> 0 And wCaptador <> Nnull(Rs3!Captador, 0) Then
                           'Verifica se indicador faz parte do grupo da tabela
                           SQl = "select count(*) from " & Banco_Sgipa & "tb_tp_grupos where autonumlista=" & wVinculo & " and tipo='C' and autonum_parceiro=" & wCaptador
                           If Nnull(db.Execute(SQl).Fields(0), 0) = 0 Then wVinculo = 0
                       End If
                    Else
                        wVinculo = 0
                    End If
                    Rs3.Close
                        
                    Rs.MoveNext
                Loop
                Rs.Close
            End If
        End If
    
        If wVinculo = 0 Then GoTo Regra2
    
    Else
        
Regra2:
        
        'indicador + importador
        SQl = "select a.autonum"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
        SQl = SQl & " where"
        SQl = SQl & " a.flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(a.hubport,0)=0"
        SQl = SQl & " and isnull(a.importador,0)=" & wImportador
        SQl = SQl & " and isnull(a.coloader,0)=0"
        SQl = SQl & " and isnull(a.despachante,0)=0"
        SQl = SQl & " and isnull(a.cocoloader,0)=0"
        SQl = SQl & " and isnull(a.cocoloader2,0)=0"
        SQl = SQl & " and isnull(a.captador,0)=" & wCaptador
        SQl = SQl & " and isnull(a.flag_acordo,0)=0"
        SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
        SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
        PRSet Rs, SQl
        If Not Rs.EOF Then
            wVinculo = Rs!autonum
        End If
        Rs.Close
        
        
                If wVinculo = 0 And wCaptador <> 0 And wImportador <> 0 Then
                SQl = "select a.autonumlista,a.tipo"
                SQl = SQl & " from"
                SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
                SQl = SQl & " where a.tipo='C' and "
                SQl = SQl & " a.autonum_parceiro = " & wCaptador
                PRSet Rs, SQl
                
                SQl = "select a.autonumlista,a.tipo"
                SQl = SQl & " from"
                SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
                SQl = SQl & " where a.tipo='I' and "
                SQl = SQl & " a.autonum_parceiro = " & wImportador
                PRSet Rs2, SQl
            
                Do While Not Rs.EOF And wVinculo = 0
                    If Not Rs2.EOF Then
                        Do While Not Rs2.EOF And wVinculo = 0
                            If Rs!autonumlista = Rs2!autonumlista Then
                                wVinculo = Rs!autonumlista
                                SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
                                SQl = SQl & " flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(hubport,0)=0"
                                SQl = SQl & " and autonum=" & wVinculo
                                SQl = SQl & " and isnull(coloader,0)=0"
                                SQl = SQl & " and isnull(cocoloader,0)=0"
                                SQl = SQl & " and isnull(cocoloader2,0)=0"
                                SQl = SQl & " and isnull(despachante,0)=0"
                                SQl = SQl & " and isnull(flag_acordo,0)=0"
                                SQl = SQl & " and data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
                                SQl = SQl & " and (final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
                                PRSet Rs3, SQl
                                If Rs3.EOF Then
                                    wVinculo = 0
                                End If
                                Rs3.Close
                            End If
                            Rs2.MoveNext
                        Loop
                        Rs2.MoveFirst
                    End If
                    Rs.MoveNext
                Loop
                Rs.Close
                Rs2.Close
            End If
        
        
    
        'importador
        If wVinculo = 0 Then
            SQl = "select a.autonum"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
            SQl = SQl & " where"
            SQl = SQl & " a.flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(a.hubport,0)=0"
            SQl = SQl & " and isnull(a.importador,0)=" & wImportador
            SQl = SQl & " and isnull(a.coloader,0)=0"
            SQl = SQl & " and isnull(a.cocoloader,0)=0"
            SQl = SQl & " and isnull(a.cocoloader2,0)=0"
            SQl = SQl & " and isnull(a.captador,0)=0"
            SQl = SQl & " and isnull(a.despachante,0)=0"
            SQl = SQl & " and isnull(a.flag_acordo,0)=0"
            SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
            SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
            PRSet Rs, SQl
            If Not Rs.EOF Then
                wVinculo = Rs!autonum
            End If
            Rs.Close
        End If
    
        'importador no grupo
        If wVinculo = 0 Then
            SQl = "select a.autonumlista,a.tipo"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
            SQl = SQl & " where"
            SQl = SQl & " a.autonum_parceiro = " & wImportador
            SQl = SQl & " and a.tipo='I'"
            PRSet Rs, SQl
            Do While Not Rs.EOF And wVinculo = 0
                SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
                SQl = SQl & " flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(hubport,0)=0"
                SQl = SQl & " and isnull(flag_acordo,0)=0"
                SQl = SQl & " and autonum=" & Rs!autonumlista
                SQl = SQl & " and isnull(despachante,0)=0"
                SQl = SQl & " and isnull(captador,0)=0"
                SQl = SQl & " and isnull(coloader,0)=0"
                SQl = SQl & " and isnull(cocoloader,0)=0"
                SQl = SQl & " and isnull(cocoloader2,0)=0"
                SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
                SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
                PRSet Rs3, SQl
                If Not Rs3.EOF Then
                    wVinculo = Rs!autonumlista
                End If
                Rs3.Close
                Rs.MoveNext
            Loop
            Rs.Close
        End If
        
        'captador
        If wVinculo = 0 Then
            SQl = "select a.autonum"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
            SQl = SQl & " where"
            SQl = SQl & " a.flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(a.hubport,0)=0"
            SQl = SQl & " and isnull(a.coloader,0)=0"
            SQl = SQl & " and isnull(a.cocoloader,0)=0"
            SQl = SQl & " and isnull(a.cocoloader2,0)=0"
            SQl = SQl & " and isnull(a.importador,0)=0"
            SQl = SQl & " and isnull(a.captador,0)=" & wCaptador
            SQl = SQl & " and isnull(a.despachante,0)=0"
            SQl = SQl & " and isnull(a.flag_acordo,0)=0"
            SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
            SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
            PRSet Rs, SQl
            If Not Rs.EOF Then
                wVinculo = Rs!autonum
            End If
            Rs.Close
        End If
    
        'captador no grupo
        If wVinculo = 0 Then
            SQl = "select a.autonumlista,a.tipo"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
            SQl = SQl & " where"
            SQl = SQl & " a.autonum_parceiro = " & wCaptador
            SQl = SQl & " and a.tipo='C'"
            PRSet Rs, SQl
            Do While Not Rs.EOF And wVinculo = 0
                SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
                SQl = SQl & " flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(hubport,0)=0"
                SQl = SQl & " and isnull(flag_acordo,0)=0"
                SQl = SQl & " and autonum=" & Rs!autonumlista
                SQl = SQl & " and isnull(despachante,0)=0"
                SQl = SQl & " and isnull(importador,0)=0"
                SQl = SQl & " and isnull(coloader,0)=0"
                SQl = SQl & " and isnull(cocoloader,0)=0"
                SQl = SQl & " and isnull(cocoloader2,0)=0"
                SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
                SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
                PRSet Rs3, SQl
                If Not Rs3.EOF Then
                    wVinculo = Rs!autonumlista
                End If
                Rs3.Close
                Rs.MoveNext
            Loop
            Rs.Close
        End If
    
        'despachante
        If wVinculo = 0 Then
            SQl = "select a.autonum"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_listas_precos a"
            SQl = SQl & " where"
            SQl = SQl & " a.flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(a.hubport,0)=0"
            SQl = SQl & " and isnull(a.coloader,0)=0"
            SQl = SQl & " and isnull(a.cocoloader,0)=0"
            SQl = SQl & " and isnull(a.cocoloader2,0)=0"
            SQl = SQl & " and isnull(a.importador,0)=0"
            SQl = SQl & " and isnull(a.captador,0)=0"
            SQl = SQl & " and isnull(a.despachante,0)=" & wDespachante
            SQl = SQl & " and isnull(a.flag_acordo,0)=0"
            SQl = SQl & " and a.data_inicio <= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss')"
            SQl = SQl & " and (a.final_validade >= to_date('" & Data_Validade & "','dd/mm/yyyy hh24:mi:ss') or a.final_validade is null)"
            PRSet Rs, SQl
            If Not Rs.EOF Then
                wVinculo = Rs!autonum
            End If
            Rs.Close
        End If
    
        'despachante no grupo
        If wVinculo = 0 Then
            SQl = "select a.autonumlista,a.tipo"
            SQl = SQl & " from"
            SQl = SQl & " " & Banco_Sgipa & "tb_tp_grupos a"
            SQl = SQl & " where"
            SQl = SQl & " a.autonum_parceiro = " & wDespachante
            SQl = SQl & " and a.tipo='D'"
            PRSet Rs, SQl
            Do While Not Rs.EOF And wVinculo = 0
                SQl = "select * from " & Banco_Sgipa & "tb_listas_precos where"
                SQl = SQl & " flag_liberada=1  and cod_empresa=" & Empresa_BL & "  and isnull(hubport,0)=0"
                SQl = SQl & " and isnull(flag_acordo,0)=0"
                SQl = SQl & " and autonum=" & Rs!autonumlista
                SQl = SQl & " and isnull(captador,0)=0"
                SQl = SQl & " and isnull(importador,0)=0"
                SQl = SQl & " and isnull(coloader,0)=0"
                SQl = SQl & " and isnull(cocoloader,0)=0"
                SQl = SQl & " and isnull(cocoloader2,0)=0"
                SQl = SQl & " and data_inicio <= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss')"
                SQl = SQl & " and (final_validade >= to_date('" & Format(Data_Validade, "dd/mm/yyyy hh:mm:ss") & "','dd/mm/yyyy hh24:mi:ss') or final_validade is null)"
                PRSet Rs3, SQl
                If Not Rs3.EOF Then
                    wVinculo = Rs!autonumlista
                End If
                Rs3.Close
                Rs.MoveNext
            Loop
            Rs.Close
        End If
    End If
End If



Padrao:
If wVinculo = 0 Then

    SQl = "select tabela_padrao from " & Banco_Sgipa & "tb_empresas where autonum = " & Empresa_BL
    PRSet RsPadrao, SQl
    If Not RsPadrao.EOF Then
        wVinculo = RsPadrao!Tabela_Padrao
    Else
        MsgBox "Tabela padrao nao encontrada - Comunique o T.I."
        wVinculo = 0
    End If
    RsPadrao.Close

End If

If wVinculo <> 0 Then
    'wVinculo = 4577
    SQl = "update " & Banco_Sgipa & "tb_bl set autonum_lista=" & wVinculo & " ,DT_TERMINO_AVERBA=sysdate where autonum=" & Lote
    If Atualizar Then Executa SQl
    
    'Atualiza historico de vinculos
    SQl = "insert into " & Banco_Sgipa & "tb_hist_vinculos ("
    SQl = SQl & "lote"
    SQl = SQl & ",data_vinculo"
    SQl = SQl & ",importador"
    SQl = SQl & ",despachante"
    SQl = SQl & ",nvocc"
    SQl = SQl & ",indicador"
    SQl = SQl & ",vinculo_anterior"
    SQl = SQl & ",vinculo_atual"
    SQl = SQl & ",tipo_vinculo"
    SQl = SQl & ",usuario"
    SQl = SQl & ") values ("
    SQl = SQl & Lote
    SQl = SQl & ",sysdate"
    SQl = SQl & ",'" & wImportador & "'"
    SQl = SQl & ",'" & wDespachante & "'"
    SQl = SQl & ",'" & wNvocc & "'"
    SQl = SQl & ",'" & wCaptador & "'"
    SQl = SQl & ",'" & wLista_Anterior & "'"
    SQl = SQl & ",'" & wVinculo & "'"
    SQl = SQl & ",'" & wTipo_Vinculo & "'"
    SQl = SQl & ",'" & Cod_Usuario & "'"
    SQl = SQl & ")"
'    executa sql  'z4
    
End If
'-----------------------------------------------------

Vincula_Tabela_Band = wVinculo

End Function



Public Function SID_Bandeirantes(Optional str As String, Optional Tipo As String, Optional Gr As String, Optional IDNOTA As Long) As String



On Error GoTo ERRO:
'TIPO - ND \ CLIENTE \ HCL
Dim j, w, l As Integer
Dim objDom As DOMDocument60
Set objDom = New DOMDocument60
Dim rsEmpresa As New ADODB.Recordset
objDom.resolveExternals = True
objDom.validateOnParse = True
objDom.async = False
Dim xmlhttp_N As New MSXML2.ServerXMLHTTP60
Set xmlhttp_N = New MSXML2.ServerXMLHTTP60
Dim xmlhttp As New MSXML2.ServerXMLHTTP60
Set xmlhttp = New MSXML2.ServerXMLHTTP60

SID_Bandeirantes = Empty

xmlhttp_N.setTimeouts 600000, 600000, 600000, 600000
xmlhttp.setTimeouts 600000, 600000, 600000, 600000


If Nnull(notaWSSID, 1) = "" Then
    SQl = "SELECT * FROM " & Banco_Sgipa & "TB_EMPRESAS WHERE AUTONUM = " & Nnull(Cod_Empresa, 0)
    PRSet rsEmpresa, SQl
    If Not rsEmpresa.EOF Then
        notaWSSID = Nnull(rsEmpresa!WS_SID, 1)
        notaUsuSID = Nnull(rsEmpresa!USU_SID, 1)
        notaSenSID = Nnull(rsEmpresa!SEN_SID, 1)
    End If
End If

If Nnull(notaWSSID, 1) = "" Then
    SID_Bandeirantes = "Parâmetro com o Endereço do Serviço SID não definido!"
    Exit Function
End If

'Call P_Select(RsParametro, "select * from " & Banco_Fatura & "parametro where cod_empresa = " & Cod_Empresa)
xmlhttp.Open "POST", "HTTP://" & Nnull(notaWSSID, 1) & "/sapiensweb/conector?SIS=CO&LOGIN=SID&ACAO=EXESENHA&NOMUSU=" & LCase(Nnull(notaUsuSID, 1)) & "&SENUSU=" & LCase(Nnull(notaSenSID, 1)), False
xmlhttp.send
Conexao = Mid(xmlhttp.responseText, 3, Len(xmlhttp.responseText) - 2)
If Not IsNumeric(Conexao) Then
    If Trim(Conexao) = "" Then
        SID_Bandeirantes = "Erro: Erro de Conexão no Serviço SID." & vbCrLf & " Não Obteve Retorno."
    Else
        SID_Bandeirantes = "Erro: Erro de Conexão no Serviço SID, favor Contate o suporte." & vbCrLf & Conexao
    End If
    Exit Function
End If

If Nnull(str, 1) <> "" Then
    objDom.loadXML str
    If objDom.parseError.reason = "" Then
        StrInstruction = str
        'Executa a AÇÃO
        xmlhttp.Open "POST", "HTTP://" & Nnull(notaWSSID, 1) & "/sapiensweb/conector?SIS=CO&LOGIN=SID&USER=" & Nnull(notaUsuSID, 1) & "&CONNECTION=" & Conexao & "&ACAO=SID.SRV.XML&" & str, False
        xmlhttp.send
        Resp_Exec_XML = xmlhttp.responseText
        
        'VERFICA SE HA ERRO DE TIME-OUT
        If InStr(1, Resp_Exec_XML, "esgotou o tempo limite") Then
            'FAZ NOVA CONEXAO
            xmlhttp_N.Open "POST", "HTTP://" & Nnull(notaWSSID, 1) & "/sapiensweb/conector?SIS=CO&LOGIN=SID&ACAO=EXESENHA&NOMUSU=" & Nnull(notaUsuSID, 1) & "&SENUSU=" & Nnull(notaSenSID, 1), False
            xmlhttp_N.send
            Conexao = Mid(xmlhttp_N.responseText, 3, Len(xmlhttp_N.responseText) - 2)
            If Not IsNumeric(Conexao) Then
                If Trim(Conexao) = "" Then
                    SID_Bandeirantes = "Erro: Erro de Conexão no Serviço SID." & vbCrLf & " Não Obteve Retorno."
                Else
                    SID_Bandeirantes = "Erro: Erro de Conexão no Serviço SID, favor Contate o suporte." & vbCrLf & Conexao
                End If
                Exit Function
            End If
            'EXECUTA A AÇÃO
            xmlhttp_N.Open "POST", "HTTP://" & Nnull(notaWSSID, 1) & "/sapiensweb/conector?SIS=CO&LOGIN=SID&USER=" & Nnull(notaUsuSID, 1) & "&CONNECTION=" & Conexao & "&ACAO=SID.SRV.XML&" & str, False
            xmlhttp_N.send
            Resp_Exec_XML = xmlhttp_N.responseText

        End If
            
        'VERIFICA SE HOUVE ERRO NO PROCESSO
        
        
        If InStr(1, Resp_Exec_XML, "ERRO:") > 0 Then
            If Nnull(Tipo, 1) = "CLIENTE" And InStr(1, UCase(Resp_Exec_XML), "JÁ CADASTRADO") > 0 Then

                
                Resp_Exec_XML = ObtemNumero2(Resp_Exec_XML)
                Resp_Exec_XML = "E" & Resp_Exec_XML
            
            Else
                Resp_Exec_XML = Mid(Resp_Exec_XML, InStr(1, Resp_Exec_XML, "ERRO:"), Len(Resp_Exec_XML) - InStr(1, Resp_Exec_XML, "ERRO:") - 1)
                Resp_Exec_XML = Mid(Resp_Exec_XML, 1, InStr(1, Resp_Exec_XML, "<") - 1)
            
            End If
        Else
            If Nnull(Tipo, 1) = "NF" Or Nnull(Tipo, 1) = "CLIENTE" Then
                Resp_Exec_XML = ObtemNumero2(Resp_Exec_XML)

'            ElseIf Nnull(Tipo, 1) = "CLIENTE" Then
'                Resp_Exec_XML = Mid(Resp_Exec_XML, InStr(1, Resp_Exec_XML, "<resultado>") + 11, 2)
            ElseIf Nnull(Tipo, 1) = "HCL" Then
                If InStr(1, Resp_Exec_XML, "OK") Then
                    Resp_Exec_XML = "OK"
                
                Else
                    Resp_Exec_XML = Mid(Resp_Exec_XML, InStr(1, Resp_Exec_XML, "</retorno>"))
                    Resp_Exec_XML = Mid(Resp_Exec_XML, InStr(1, Resp_Exec_XML, "<retorno>") + 9)
                    Resp_Exec_XML = Replace(Resp_Exec_XML, "</retorno></sidxml>", "")
                    Resp_Exec_XML = Replace(Resp_Exec_XML, "\\", "\")
                    Resp_Exec_XML = Replace(Resp_Exec_XML, "C:\", "\\" & Nnull(notaWSSID, 1) & "\")
                
                End If
            Else
                'Resp_Exec_XML = "<sidxml><resultado>OK</resultado><retorno>OK</retorno><retorno>C:\Senior\\Sapiens\web\html\tempimg\SAPIENSSRV_TMP_00788_3283598900000.JPG</retorno></sidxml>"
                
                Resp_Exec_XML = Mid(Resp_Exec_XML, InStr(1, Resp_Exec_XML, "</retorno>"))
                Resp_Exec_XML = Mid(Resp_Exec_XML, InStr(1, Resp_Exec_XML, "<retorno>") + 9)
                Resp_Exec_XML = Replace(Resp_Exec_XML, "</retorno></sidxml>", "")
                Resp_Exec_XML = Replace(Resp_Exec_XML, "\\", "\")
                Resp_Exec_XML = Replace(Resp_Exec_XML, "C:\", "\\" & Nnull(notaWSSID, 1) & "\")
            
            End If
        End If
    Else
        Resp_Exec_XML = "ERRO: Requisição Inválida!"
    End If
End If
    
    
SID_Bandeirantes = Resp_Exec_XML
    
'logoff
xmlhttp_N.Open "POST", "HTTP://" & Nnull(notaWSSID, 1) & "/sapiensweb/conector?SIS=CO&LOGIN=SID&USER=" & Nnull(notaUsuSID, 1) & "&CONNECTION=" & Conexao & "&ACAO=LOGOFF", False
xmlhttp_N.send
Set xmlhttp_N = Nothing
Set xmlhttp_N = New MSXML2.ServerXMLHTTP60
    
On Error GoTo Saida:
If Nnull(DirXMLIntegracao, 1) <> "" And Tipo = "NF" Then
    If Right(DirXMLIntegracao, 1) <> "\" Then DirXMLIntegracao = DirXMLIntegracao & "\"
    'If testaAcessoDir(DirXMLIntegracao) Then
    If Nnull(Gr, 0) > 0 Then
        nomeArq = "GR_" & Replace(Replace(Gr, ",", "_"), ";", "_") & "_" & Format(Now, "ddmmyyyymmhhss")
    Else
        nomeArq = "MANUAL_" & IDNOTA & "_" & Format(Now, "ddmmyyyymmhhss")
    End If
    Open DirXMLIntegracao & nomeArq & ".xml" For Output As #1
    Print #1, str
    Close #1
    'End If
End If

Exit Function
    
ERRO:
    SID_Bandeirantes = "ERRO:" & Err.Description
    Err.Description = Empty
    Exit Function
    
Saida:
    Err.Description = Empty

End Function
Public Function ObtemNumero(Valor As String) As String

Dim i As Integer
ObtemNumero = ""



For i = 1 To Len(Valor)
    If IsNumeric(Mid(Valor, i, 1)) = True Then
        ObtemNumero = ObtemNumero & Mid(Valor, i, 1)
    End If
Next i

End Function
Public Function ObtemNumero2(Valor As String) As String


Dim i As Integer
ObtemNumero2 = ""
Dim Achou As Boolean
Achou = False

For i = 1 To Len(Valor)
    If IsNumeric(Mid(Valor, i, 1)) = True Then
        ObtemNumero2 = ObtemNumero2 & Mid(Valor, i, 1)
        Achou = True
    Else
        If Achou Then Exit For
    End If
Next i

End Function
Public Sub Atualiza_Vencimento(Lote As String, Data1 As String, Data2 As String, BL_Reefer As Boolean, FormaPag As String)

Dim Tbh As ADODB.Recordset
Dim TbAux As ADODB.Recordset
Dim Tbs As ADODB.Recordset
Dim WSQL As String
Dim WSQL1 As String


If Lote <> "" Then
 SQl = "  Select   Max(Validade_Gr) as Data1 ,Max(DT_BASE_CALCULO_REEFER) as Data2"
 SQl = SQl & "  From tb_gr_bl a , tb_servicos_faturados b where a.seq_gr=b.seq_gr and a.bl=" & Lote
 SQl = SQl & "  and a.status_gr in('IM','GE') and (a.forma_pagamento=2 or (a.forma_pagamento=3 and (b.servico=52 or b.servico=71 or b.servico=57)))"
 PRSet Tbh, SQl
 If Not Tbh.EOF Then
       WSQL = ""
       WSQL1 = ""
       If Not IsNull(Tbh!Data2) And BL_Reefer Then
          WSQL = " To_date('" & Format(Tbh!Data2, "DD/MM/YYYY") & "','dd/mm/yyyy') "
       Else
          If Not IsNull(Tbh!Data1) Then
             WSQL = " To_date('" & Format(Tbh!Data1, "DD/MM/YYYY") & "','dd/mm/yyyy') "
          End If
       End If
       If WSQL <> "" Then
            SQl = "  Select   count(1) as contar From tb_servicos_faturados where seq_gr is null and (servico=52 or servico=57 or servico=71) AND bl=" & Lote
            PRSet Tbs, SQl
            If Tbs!Contar > 0 Then
               If IsDate(Data1) Then
                  If CVDate(Data1) > (Tbh!Data1) Then WSQL = " To_Date('" & Data1 & "','dd/mm/yyyy') "
               End If
               If IsDate(Data2) And BL_Reefer Then
                  If CVDate(Data2) > (Tbh!Data2) Then WSQL = " To_Date('" & Data2 & "','dd/mm/yyyy') "
               End If
            End If
            SQl = "  Select   COUNT(1) AS Contar  From tb_gr_bl a  where a.bl=" & Lote
            SQl = SQl & "  and a.status_gr in('IM','GE') and (a.forma_pagamento=2 )"
            PRSet Tbs, SQl
            If Tbs!Contar > 0 Then
               WSQL1 = "Update Tb_Cntr_Bl Set Dt_Fim_Periodo=" & WSQL & " where Autonum in (Select Cntr From Tb_Amr_Cntr_Bl Where Bl =" & Val(Lote) & ")"
               Executa WSQL1
               WSQL1 = "Update Tb_CARGA_SOLTA Set Dt_Fim_Periodo=" & WSQL & " where Bl =" & Val(Lote)
               Executa WSQL1
            Else
               WSQL = "Update Tb_Cntr_Bl Set Dt_Fim_Periodo=null where Autonum in (Select Cntr From Tb_Amr_Cntr_Bl Where Bl =" & Val(Lote) & ")"
               Executa WSQL
               WSQL = "Update Tb_CARGA_SOLTA Set Dt_Fim_Periodo=NULL  where Bl =" & Val(Lote)
               Executa WSQL
            End If
       Else
         If UCase(FormaPag) <> "FATURADO" Then
            If IsDate(Data1) Then
               WSQL = " To_Date('" & Data1 & "','dd/mm/yy') "
            End If
            If IsDate(Data2) And BL_Reefer Then
               WSQL = " To_date('" & Data2 & "','dd/mm/yy') "
            End If
            If WSQL <> "" Then
               WSQL1 = "Update Tb_Cntr_Bl Set Dt_Fim_Periodo=" & WSQL & " where Autonum in (Select Cntr From Tb_Amr_Cntr_Bl Where Bl =" & Val(Lote) & ")"
               Executa WSQL1
               WSQL1 = "Update Tb_CARGA_SOLTA Set Dt_Fim_Periodo=" & WSQL & " where Bl =" & Val(Lote)
               Executa WSQL1
            End If
         Else
            If WSQL = "" Then
               WSQL = "Update Tb_Cntr_Bl Set Dt_Fim_Periodo=null where Autonum in (Select Cntr From Tb_Amr_Cntr_Bl Where Bl =" & Val(Lote) & ")"
               Executa WSQL
               WSQL = "Update Tb_CARGA_SOLTA Set Dt_Fim_Periodo=NULL  where Bl =" & Val(Lote)
               Executa WSQL
            End If
         End If
       End If
   End If
 End If
End Sub

Public Sub Limpa_Tela(Tela As Form)

Dim Controle As Object
Dim i As Integer

On Error Resume Next
TxtautonumCntr = 0
        For i = 0 To Tela.Controls.Count - 1
            Set Controle = Tela.Controls(i)
            If TypeOf Controle Is CheckBox Then Controle.Value = 0
            If TypeOf Controle Is DataCombo Then Controle.text = ""
            If TypeOf Controle Is ComboBox Then Controle.text = ""
            If TypeOf Controle Is TextBox Then Controle.text = ""
            If TypeOf Controle Is DataGrid Then Controle.Clear
            If TypeOf Controle Is MSFlexGrid Then Controle.Clear
            If TypeOf Controle Is MSHFlexGrid Then Controle.Clear
        Next i
                          
End Sub
Public Function SID_CLI_Band(Optional CGC As String, Optional P_codcli, Optional P_Autonum As Long, Optional emailCalculo As String) As String
On Error GoTo ERRO:
Dim clicep As String
Dim Rs As New ADODB.Recordset
Dim RsCep As New ADODB.Recordset
Dim RsCad As New ADODB.Recordset
Dim w As Integer
Dim TIP As Integer
Dim rsFiliais As New ADODB.Recordset
Dim rsaux As New ADODB.Recordset

SID_CLI_Band = Empty
cod_cli = Empty

If Nnull(P_Autonum, 0) > 0 Then
    PRSet Rs, "Select * from " & Banco_Sgipa & "tb_cad_parceiros where autonum = " & Nnull(P_Autonum, 0)
Else
    PRSet Rs, "select * from " & Banco_Sgipa & "tb_cad_parceiros where autonum in(Select max(autonum) as autonum from " & Banco_Sgipa & "tb_cad_parceiros where cgc = '" & CGC & "' )"
End If
clicep = Val(ObtemNumero(Nnull(Rs!CEP, 1)))


If Not Rs.EOF Then
    DoEvents
    
    If Nnull(Cod_Empresa, 1) <> "1" And Nnull(Cod_Empresa, 1) <> "2" Then Cod_Empresa = "1"
    
    SQl = "SELECT * FROM " & Banco_Sgipa & "TB_EMPRESAS WHERE AUTONUM = " & Nnull(Cod_Empresa, 0)
    PRSet rsaux, SQl
    If Not rsaux.EOF Then
        notaWSSID = Nnull(rsaux!WS_SID, 1)
        notaUsuSID = Nnull(rsaux!USU_SID, 1)
        notaSenSID = Nnull(rsaux!SEN_SID, 1)
        notaCodemp = Nnull(rsaux!CODEMP, 0)
        notaCodfil = Nnull(rsaux!CODFIL, 0)
    Else
        SID_CLI_Band = "Falta de Parâmetros necessários à execução da Integração!" & vbCrLf & "Favor Editar o Cadastro da Empresa!"
        Exit Function
    End If
    
    StrInstruction = "<sidxml><param retorno='XML'/>"
    StrInstruction = StrInstruction & "<sid acao='SID.CLI.GRAVAR'>"
    StrInstruction = StrInstruction & "<param nome='NOMCLI' valor='" & Replace(Nnull(Rs!Razao, 1), "&", "E") & "'/>"
    StrInstruction = StrInstruction & "<param nome='APECLI' valor='" & Replace(IIf(Nnull(Rs!Fantasia, 1) = "", Nnull(Rs!Razao, 1), Nnull(Rs!Fantasia, 1)), "&", "E") & "'/>"
    
    If Nnull(ObtemNumero(Rs!CGC), 0) = 0 And (Nnull(Rs!Estado, 1) = "" Or Nnull(Rs!Estado, 1) = "EX") Then
        StrInstruction = StrInstruction & "<param nome='SIGUFS' valor='EX'/>"
    Else
        StrInstruction = StrInstruction & "<param nome='SIGUFS' valor='" & Nnull(Rs!Estado, 1) & "'/>"
    End If
    
    
    StrInstruction = StrInstruction & "<param nome='FONCLI' valor='" & Nnull(Rs!FONE1, 1) & "'/>"
    StrInstruction = StrInstruction & "<param nome='FONCL2' valor='" & Nnull(Rs!FONE2, 1) & "'/>"
    StrInstruction = StrInstruction & "<param nome='FAXCLI' valor='" & Nnull(Rs!Fax, 1) & "'/>"
    If Nnull(Rs!Email, 1) = "" And Nnull(emailCalculo, 1) <> "" Then
        StrInstruction = StrInstruction & "<param nome='INTNET' valor='" & Nnull(emailCalculo, 1) & "'/>"
    Else
        StrInstruction = StrInstruction & "<param nome='INTNET' valor='" & Nnull(Rs!Email, 1) & "'/>"
    End If
    StrInstruction = StrInstruction & "<param nome='CGCCPF' valor='" & ObtemNumero(Nnull(Rs!CGC, 1)) & "'/>"
    StrInstruction = StrInstruction & "<param nome='NOMCTO' valor='" & Nnull(Rs!CONTATO, 1) & "'/>"
    StrInstruction = StrInstruction & "<param nome='ENDCLI' valor='" & IIf(Nnull(Rs!logradouro, 1) = "", ".", Nnull(Rs!logradouro, 1)) & "'/>"
    StrInstruction = StrInstruction & "<param nome='BAICLI' valor='" & IIf(Nnull(Rs!BAIRRO, 1) = "", ".", Nnull(Rs!BAIRRO, 1)) & "'/>"
    StrInstruction = StrInstruction & "<param nome='CIDCLI' valor='" & IIf(Nnull(Rs!Cidade, 1) = "", ".", Nnull(Rs!Cidade, 1)) & "'/>"
    StrInstruction = StrInstruction & "<param nome='CEPCLI' valor='" & clicep & "'/>"
    StrInstruction = StrInstruction & "<param nome='CODPAI' valor='" & Left(Nnull(Rs!PAIS, 1), 4) & "'/>"
    StrInstruction = StrInstruction & "<param nome='INSEST' valor='" & Replace(Replace(Replace(Nnull(Rs!IE, 1), "/", ""), ".", ""), "-", "") & "'/>"
    StrInstruction = StrInstruction & "<param nome='INSMUN' valor='" & Replace(Replace(Replace(Nnull(Rs!IM, 1), "/", ""), ".", ""), "-", "") & "'/>"
    StrInstruction = StrInstruction & "<param nome='SITCLI' valor='A" & "'/>"
    StrInstruction = StrInstruction & "<param nome='TIPCLI' valor='" & IIf(Len(Replace(Replace(Replace(ObtemNumero(Nnull(Rs!CGC, 1)), "-", ""), "/", ""), ".", "")) = 14, "J", "F") & "'/>"
    If Nnull(ObtemNumero(Rs!CGC), 0) = 0 And (Nnull(Rs!Estado, 1) = "" Or Nnull(Rs!Estado, 1) = "EX") Then
        StrInstruction = StrInstruction & "<param nome='TIPMER' valor='E" & "'/>"
    End If
    StrInstruction = StrInstruction & "<param nome='NENCLI' valor='0" & "'/>"
    
    If Trim(Nnull(P_codcli, 0)) <> 0 Then
        StrInstruction = StrInstruction & "<param nome='CODCLI' valor='" & Nnull(P_codcli, 0) & "'/>"
    End If
    
    StrInstruction = StrInstruction & "</sid></sidxml>"
    cod_cli = SID_Bandeirantes(StrInstruction, "CLIENTE")
    
    If InStr(1, cod_cli, "ERRO:") > 0 Then
        SID_CLI_Band = cod_cli
        GoTo Sem_Definicao
    ElseIf Mid(cod_cli, 1, 1) = "E" Then
        cod_cli = Mid(cod_cli, 2)
        SID_CLI_Band = cod_cli
        GoTo Sem_Definicao
    Else
        SID_CLI_Band = cod_cli
        If Nnull(P_codcli, 0) <> 0 Then GoTo Sem_Definicao
    End If
    
    SQl = "SELECT DISTINCT RAZAO_SOCIAL, CODEMP, CODFIL FROM " & Banco_Sgipa & "TB_EMPRESAS "
    PRSet rsFiliais, SQl
    While Not rsFiliais.EOF
        DoEvents
        
        StrInstruction = Empty
        StrInstruction = "<sidxml><param retorno='XML'/>"
        StrInstruction = StrInstruction & "<sid acao='SID.CLI.GravarDefinicoes'>"
        StrInstruction = StrInstruction & "<param nome='CODCPG' valor='AV'/>"
        StrInstruction = StrInstruction & "<param nome='CODREP' valor='1'/>"
        StrInstruction = StrInstruction & "<param nome='LIMAPR' valor='S'/>"
        StrInstruction = StrInstruction & "<param nome='VLRLIM' valor='999999999'/>"
        StrInstruction = StrInstruction & "<param nome='CODFIL' valor='" & Nnull(rsFiliais!CODFIL, 0) & "'/>"
        StrInstruction = StrInstruction & "<param nome='CODEMP' valor='" & Nnull(rsFiliais!CODEMP, 0) & "'/>"
        StrInstruction = StrInstruction & "<param nome='CODCLI' valor='" & Nnull(cod_cli, 1) & "'/>"
        StrInstruction = StrInstruction & "<param nome='CATCLI' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='CRIRAT' valor='5'/>"
        StrInstruction = StrInstruction & "<param nome='EXILCP' valor='N'/>"
        StrInstruction = StrInstruction & "<param nome='SALDUP' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='SALCRE' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='SALOUT' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='DATLIM' valor='" & Format(Date, "DD/MM/YYYY") & "'/>"
        StrInstruction = StrInstruction & "<param nome='VLRPFA' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='DATMAC' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='VLRMAC' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='VLRUPE' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='DATUPE' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='VLRUPG' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='DATUPG' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='VLRUFA' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='DATUFA' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='QTDCHS' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='MEDATR' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='MAIATR' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='VLRATR' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='DATATR' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='VLRMFA' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='DATMFA' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='QTDTPC' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='VLRUPC' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='DATUPC' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='QTDPGT' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='ACEPAR' valor='S'/>"
        StrInstruction = StrInstruction & "<param nome='PERAQA' valor='999'/>"
        StrInstruction = StrInstruction & "<param nome='QTDMFP' valor='999'/>"
        StrInstruction = StrInstruction & "<param nome='JUNPED' valor='N'/>"
        StrInstruction = StrInstruction & "<param nome='INDAGR' valor='S'/>"
        StrInstruction = StrInstruction & "<param nome='PERCOM' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='PERDS2' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='PERDS1' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='CODCRP' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='CODRVE' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='QTDDCV' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='CODFPG' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='CODTPR' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='CRIEDV' valor='S'/>"
        StrInstruction = StrInstruction & "<param nome='CODRED' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='CODTRA' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='TOLDSC' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='PERDSC' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='RECDTM' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='RECMUL' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='RECDTJ' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='RECTJR' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='RECJMM' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='PORSI2' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='PORSI1' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='CODCRT' valor='99'/>"
        StrInstruction = StrInstruction & "<param nome='CODIN2' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='CODIN1' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='PORNA2' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='PORNA1' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='CCBCLI' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='CODAGE' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='CODBAN' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='CODFRJ' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='USUAGE' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='CTAFCR' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='CTAFDV' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='CTARED' valor='100302'/>"
        StrInstruction = StrInstruction & "<param nome='CTARCR' valor='200004'/>"
        StrInstruction = StrInstruction & "<param nome='PERFRE' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='PERSEG' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='PEREMB' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='PERENC' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='PEROUT' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='FVETNS' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='FVEFPG' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='FVECPG' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='PERISS' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='CODVEN' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='INDEXP' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='DATPAL' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='HORPAL' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='CLIFOB' valor='C'/>"
        StrInstruction = StrInstruction & "<param nome='CODTAB' valor=''/>"
        StrInstruction = StrInstruction & "<param nome='PERDS3' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='PERDS4' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='PEROF1' valor='0'/>"
        StrInstruction = StrInstruction & "<param nome='PEROF2' valor='0'/>"
        StrInstruction = StrInstruction & "</sid></sidxml>"
        If SID_Bandeirantes(StrInstruction, "HCL") <> "OK" Then
            SID_CLI_Band = "Não foi possivel criar as definições do cliente para a empresa [" & rsFiliais!razao_social & "]"
            Exit Function
        End If

        rsFiliais.MoveNext
    Wend
Sem_Definicao:
    
End If
'SID_CLI_Band = cod_cli
Exit Function

ERRO:
    SID_CLI_Band = Err.Description
    Err.Description = Empty
    
End Function

Function Exec_Xml_Sap(Optional Letra As String, Optional str As String)
Dim j, w, l As Integer
Dim objDom As DOMDocument60
Set objDom = New DOMDocument60
objDom.resolveExternals = True
objDom.validateOnParse = True
objDom.async = False
Dim xmlhttp_N As New MSXML2.ServerXMLHTTP60
Set xmlhttp_N = New MSXML2.ServerXMLHTTP60
Dim xmlhttp As New MSXML2.ServerXMLHTTP60
Set xmlhttp = New MSXML2.ServerXMLHTTP60

xmlhttp_N.setTimeouts 60000, 60000, 60000, 60000
xmlhttp.setTimeouts 60000, 60000, 60000, 60000

'Call P_Select(RsParametro, "select * from " & Banco_Fatura & "parametro where cod_empresa = " & Cod_Empresa)
PRSet RsParametro, "select * from " & Banco_Fatura & "parametro where cod_empresa = " & Cod_Empresa

xmlhttp.Open "POST", "HTTP://" & Nnull(RsParametro!servidor_sid, 1) & "/sapiensweb/conector.exe?SIS=CO&LOGIN=SID&ACAO=EXESENHA&NOMUSU=" & Nnull(RsParametro!LOGIN_SID, 1) & "&SENUSU=" & Nnull(RsParametro!SENHA_SID, 1), False

xmlhttp.send
Conexao = Mid(xmlhttp.responseText, 3, Len(xmlhttp.responseText) - 2)
If Not IsNumeric(Conexao) Then
    If Trim(Conexao) = "" Then
        Resp_Exec_XML = "Erro: Erro de Conexão no Serviço SID." & vbCrLf & " Não Obteve Retorno."
    Else
        Resp_Exec_XML = "Erro: Erro de Conexão no Serviço SID, favor Contate o suporte." & vbCrLf & Conexao
    End If
    Exit Function
End If

If Nnull(str, 1) <> "" Then
    objDom.loadXML str
    If objDom.parseError.reason = "" Then
        StrInstruction = str
            
        'Executa a AÇÃO
        xmlhttp_N.Open "POST", "HTTP://" & Nnull(RsParametro!servidor_sid, 1) & "/sapiensweb/conector.exe?SIS=CO&LOGIN=SID&USER=" & Nnull(RsParametro!LOGIN_SID, 1) & "&CONNECTION=" & Conexao & "&ACAO=SID.SRV.XML&" & str, False
        xmlhttp_N.send
        Resp_Exec_XML = xmlhttp_N.responseText
        'VERFICA SE HA ERRO DE TIME-OUT
        If InStr(1, Resp_Exec_XML, "esgotou o tempo limite") Then
            'FAZ NOVA CONEXAO
            xmlhttp_N.Open "POST", "HTTP://" & Nnull(RsParametro!servidor_sid, 1) & "/sapiensweb/conector.exe?SIS=CO&LOGIN=SID&ACAO=EXESENHA&NOMUSU=" & Nnull(RsParametro!LOGIN_SID, 1) & "&SENUSU=" & Nnull(RsParametro!SENHA_SID, 1), False
            xmlhttp_N.send
            Conexao = Mid(xmlhttp_N.responseText, 3, Len(xmlhttp_N.responseText) - 2)
            If Not IsNumeric(Conexao) Then
                If Trim(Conexao) = "" Then
                    Resp_Exec_XML = "Erro: Erro de Conexão no Serviço SID." & vbCrLf & " Não Obteve Retorno."
                Else
                    Resp_Exec_XML = "Erro: Erro de Conexão no Serviço SID, favor Contate o suporte." & vbCrLf & Conexao
                End If
                Exit Function
            End If
            'EXECUTA A AÇÃO
            xmlhttp_N.Open "POST", "HTTP://" & Nnull(RsParametro!servidor_sid, 1) & "/sapiensweb/conector.exe?SIS=CO&LOGIN=SID&USER=" & Nnull(RsParametro!LOGIN_SID, 1) & "&CONNECTION=" & Conexao & "&ACAO=SID.SRV.XML&" & str, False
            xmlhttp_N.send
            Resp_Exec_XML = xmlhttp_N.responseText
        End If
            
        'VERIFICA SE HOUVE ERRO NO PROCESSO
        If InStr(1, Resp_Exec_XML, "ERRO:") > 0 Then
            Resp_Exec_XML = Mid(Resp_Exec_XML, InStr(1, Resp_Exec_XML, "ERRO:"), Len(Resp_Exec_XML) - InStr(1, Resp_Exec_XML, "ERRO:") - 1)
            Resp_Exec_XML = Mid(Resp_Exec_XML, 1, InStr(1, Resp_Exec_XML, "<") - 1)
        Else
            Resp_Exec_XML = Mid(Resp_Exec_XML, InStr(1, Resp_Exec_XML, "<resultado>") + 11, 2)
        End If
    End If
End If
    
'logoff
xmlhttp_N.Open "POST", "HTTP://" & Nnull(RsParametro!servidor_sid, 1) & "/sapiensweb/conector.exe?SIS=CO&LOGIN=SID&USER=" & Nnull(RsParametro!LOGIN_SID, 1) & "&CONNECTION=" & Conexao & "&ACAO=LOGOFF", False
xmlhttp_N.send
Set xmlhttp_N = Nothing
Set xmlhttp_N = New MSXML2.ServerXMLHTTP60
    

End Function


Function Nota_sap(Optional Letra As String, Optional str As String, Optional Corpo As String, Optional ID As Long)
''''NFE - SANTOS
'''Dim Retorno As String
'''Dim Teste As New SoapClient30
'''Private Const WSTransp As String = "http://10.1.50.17/tecter/service.asmx?WSDL"
'''Teste.MSSoapInit (WSTransp)
'''Retorno = Teste.geranfvenda(ID, str, Corpo)
'''Num_Nota = Retorno

Dim j, w, l As Integer
Dim objDom As DOMDocument60
Set objDom = New DOMDocument60
objDom.resolveExternals = True
objDom.validateOnParse = True
objDom.async = False
Dim xmlhttp_N As New MSXML2.ServerXMLHTTP60
Set xmlhttp_N = New MSXML2.ServerXMLHTTP60
Dim xmlhttp As New MSXML2.ServerXMLHTTP60
Set xmlhttp = New MSXML2.ServerXMLHTTP60

xmlhttp_N.setTimeouts 60000, 60000, 60000, 60000
xmlhttp.setTimeouts 60000, 60000, 60000, 60000


'Call P_Select(RsParametro, "select * from " & Banco_Fatura & "parametro where cod_empresa = " & Cod_Empresa)
PRSet RsParametro, "select * from " & Banco_Fatura & "parametro where cod_empresa = " & Cod_Empresa

xmlhttp.Open "POST", "HTTP://" & Nnull(RsParametro!servidor_sid, 1) & "/sapiensweb/conector.exe?SIS=CO&LOGIN=SID&ACAO=EXESENHA&NOMUSU=" & Nnull(RsParametro!LOGIN_SID, 1) & "&SENUSU=" & Nnull(RsParametro!SENHA_SID, 1), False

xmlhttp.send
Conexao = Mid(xmlhttp.responseText, 3, Len(xmlhttp.responseText) - 2)
If Not IsNumeric(Conexao) Then
    If Trim(Conexao) = "" Then
        Erro_SID = "Erro: Erro de Conexão no Serviço SID." & vbCrLf & " Não Obteve Retorno."
    Else
        Erro_SID = "Erro: Erro de Conexão no Serviço SID, favor Contate o suporte." & vbCrLf & Conexao
    End If
    Exit Function
End If

Num_Nota = Empty
Erro_SID = Empty
If Nnull(str, 1) <> "" Then
    objDom.loadXML str
    If objDom.parseError.reason = "" Then
        StrInstruction = str
        'Executa a AÇÃO
        xmlhttp_N.Open "POST", "HTTP://" & Nnull(RsParametro!servidor_sid, 1) & "/sapiensweb/conector.exe?SIS=CO&LOGIN=SID&USER=" & Nnull(RsParametro!LOGIN_SID, 1) & "&CONNECTION=" & Conexao & "&ACAO=SID.SRV.XML", False
        Call xmlhttp_N.setRequestHeader("Content-type", "application/x-www-form-urlencoded")
        xmlhttp_N.send (str)
        Num_Nota = xmlhttp_N.responseText
        Resp = Num_Nota
        'VERFICA SE HA ERRO DE TIME-OUT
        If InStr(1, Num_Nota, "esgotou o tempo limite") Then
            'FAZ NOVA CONEXAO
            xmlhttp_N.Open "POST", "HTTP://" & Nnull(RsParametro!servidor_sid, 1) & "/sapiensweb/conector.exe?SIS=CO&LOGIN=SID&ACAO=EXESENHA&NOMUSU=" & Nnull(RsParametro!LOGIN_SID, 1) & "&SENUSU=" & Nnull(RsParametro!SENHA_SID, 1), False
            xmlhttp_N.send
            Conexao = Mid(xmlhttp_N.responseText, 3, Len(xmlhttp_N.responseText) - 2)
            If Not IsNumeric(Conexao) Then
                If Trim(Conexao) = "" Then
                    Erro_SID = "Erro: Erro de Conexão no Serviço SID." & vbCrLf & " Não Obteve Retorno."
                Else
                    Erro_SID = "Erro: Erro de Conexão no Serviço SID, favor Contate o suporte." & vbCrLf & Conexao
                End If
                Num_Nota = Empty
                Exit Function
            End If
            'EXECUTA A AÇÃO
            xmlhttp_N.Open "POST", "HTTP://" & Nnull(RsParametro!servidor_sid, 1) & "/sapiensweb/conector.exe?SIS=CO&LOGIN=SID&USER=" & Nnull(RsParametro!LOGIN_SID, 1) & "&CONNECTION=" & Conexao & "&ACAO=SID.SRV.XML", False
            Call xmlhttp_N.setRequestHeader("Content-type", "application/x-www-form-urlencoded")
            xmlhttp_N.send (str)
            Num_Nota = xmlhttp_N.responseText
            Resp = Num_Nota
        End If
            
        'VERIFICA SE HOUVE ERRO NO PROCESSO
        If InStr(1, Num_Nota, "ERRO:") > 0 Then
            Erro_SID = Mid(Num_Nota, InStr(1, Num_Nota, "ERRO:"), Len(Num_Nota) - InStr(1, Num_Nota, "ERRO:") - 1)
            Erro_SID = Mid(Erro_SID, 1, InStr(1, Erro_SID, "<") - 1)
            Num_Nota = Empty
        Else
            'COMPOE O NUMERO DA NOTA FISCAL RETORNADO DO SAPIENS
            Num_Nota = Empty
            For l = 1 To Len(Resp)
                If IsNumeric(Mid(Resp, l, 1)) = True Then
                    Num_Nota = Num_Nota & Mid(Resp, l, 1)
                End If
            Next l
        End If
    End If
End If
    
'logoff
xmlhttp_N.Open "POST", "HTTP://" & Nnull(RsParametro!servidor_sid, 1) & "/sapiensweb/conector.exe?SIS=CO&LOGIN=SID&USER=" & Nnull(RsParametro!LOGIN_SID, 1) & "&CONNECTION=" & Conexao & "&ACAO=LOGOFF", False
xmlhttp_N.send
Set xmlhttp_N = Nothing
Set xmlhttp_N = New MSXML2.ServerXMLHTTP60
    

End Function

Function Nota_sap_NFE(Optional Letra As String, Optional str As String, Optional Corpo As String, Optional ID As Long)
'''NFE - SANTOS

Dim rsRPS As New ADODB.Recordset
On Error GoTo trata:

Dim Teste As New MSSOAPLib30.SoapClient30
Dim WSTransp As String
WSTransp = Nnull(RsParametro!ws_nfe, 1)
Teste.MSSoapInit (WSTransp)
Teste.ConnectorProperty("Timeout") = 180000
Retorno = ""
Retorno = Teste.geranfvenda(ID, str, Corpo)
Retorno = UCase(Retorno)
RPS = False
If InStr(1, Retorno, "COM SUCESSO") > 0 Or InStr(1, Retorno, "NFE GERADA") > 0 Then
    For l = 1 To Len(Retorno)
        If IsNumeric(Mid(Retorno, l, 1)) = True Then
            Num_Nota = Num_Nota & Mid(Retorno, l, 1)
        End If
        
        SQl = "SELECT MAX(RPSNUM) AS RPS FROM " & Banco_Fatura & "RPSFAT WHERE NFE =" & Nnull(Num_Nota, 0)
        PRSet rsRPS, SQl
        If Not rsRPS.EOF Then
            StrInstruction = "<sidxml><param retorno='XML'/><sid acao='SID.SRV.ALTEMPFIL'>"
            StrInstruction = StrInstruction & "<param nome='CodEmp' valor='" & Empresa & "'/>"
            StrInstruction = StrInstruction & "<param nome='CodFil' valor='" & FIL & "'/></sid>"
            StrInstruction = StrInstruction & "<sid acao='SID.Tcr.Alterar'>"
            If Cod_Empresa = 1 Then
                StrInstruction = StrInstruction & "<param nome='NumTit' valor='" & Int(Num_Nota) & ":01" & "'/>"
            Else
                StrInstruction = StrInstruction & "<param nome='NumTit' valor='" & Int(Num_Nota) & "-01" & "'/>"
            End If
            StrInstruction = StrInstruction & "<param nome='CodTpt' valor='DUP'/>"
            StrInstruction = StrInstruction & "<param nome='USU_RPS' valor='" & Nnull(rsRPS!RPS, 0) & "'/>"
            StrInstruction = StrInstruction & "</sid></sidxml>"
            Resp_Exec_XML = Empty
            Call Exec_Xml_Sap("", StrInstruction)
        End If
        
        
    Next l
ElseIf InStr(1, Retorno, "AGUARDANDO ENVIO À PREFEITURA") > 0 Then
    RPS = True
    For l = 1 To Len(Retorno)
        If IsNumeric(Mid(Retorno, l, 1)) = True Then
            Num_Nota = Num_Nota & Mid(Retorno, l, 1)
        End If
    Next l
ElseIf InStr(1, Retorno, "GERADO, PORÉM NÃO FOI POSSÍVEL PROCESSAR O RETORNO DE SUA TRANSMISSÃO A PREFEITURA") > 0 Then
    RPS = True
    For l = 1 To Len(Retorno)
        If IsNumeric(Mid(Retorno, l, 1)) = True Then
            Num_Nota = Num_Nota & Mid(Retorno, l, 1)
        End If
    Next l
ElseIf InStr(1, Retorno, "ERRO") > 0 Then
    Erro_SID = Retorno
    Num_Nota = 0
Else
    If Nnull(Retorno, 1) = "" Then
        Erro_SID = "Sem Retorno"
        Num_Nota = 0
    Else
        Erro_SID = Retorno
        Num_Nota = 0
    End If
End If


Open App.Path & "\TESTE_NFE_" & Format(Now, "ddmmyyyyhhmmss") & ".txt" For Output As #1
Print #1, "Data:" & Now
Print #1, "Conexão:" & Nnull(RsParametro!ws_nfe, 1)
Print #1, "Corpo da Nota:" & Corpo
Print #1, "ID da TB_FATURANOTA:" & ID
Print #1, "XML inclusão:" & str
Print #1, "Retorno:" & Nnull(Retorno, 1)
Close #1
Exit Function

trata:
Retorno = Err.Description
MsgBox "A Transação Não Pode ser executada, veja a mensagem:" & Err.Description
Open App.Path & "\TESTE_NFE_" & Format(Now, "ddmmyyyyhhmmss") & ".txt" For Output As #1
Print #1, "Data:" & Now
Print #1, "Conexão:" & Nnull(RsParametro!ws_nfe, 1)
Print #1, "Corpo da Nota:" & Corpo
Print #1, "ID da TB_FATURANOTA:" & ID
Print #1, "XML inclusão:" & str
Print #1, "Retorno:" & Nnull(Retorno, 1)
Close #1


End Function

Function SID_HIST_CLI(P_codcli As String)

    StrInstruction = AlfaWeb & "&ACAO=EXESENHA&NOMUSU=" & Nnull(FAT_PAR!LOGIN_SID, 1) & "&SENUSU=" & Nnull(FAT_PAR!SENHA_SID, 1)
    i = 0
    Do
        Conexao = LiberaFaturamentosa_xml.Inet1.OpenURL(UCase(StrInstruction))
         
        i = i + 1
    Loop Until i < 3 Or IsNumeric(Conexao) = True

    If Not IsNumeric(Conexao) Then
        MsgBox "Erro: Erro de Conexão no Serviço SID, favor Contate o suporte." & vbCrLf & Conexao, vbCritical
        Call P_Execute("Update " & Banco_Fatura & "faturanota set cancelada = 1, dt_cancelada = sysdate, motivo_cancela = 'ERRO SID.' where  id =" & Posicao)
        Exit Function
    End If
    
    If Cod_Empresa = 1 Then
        Empresa = 2
        FIL = 1
    Else
        Empresa = 3
        FIL = 1
    End If


    'defini as filiais
    StrInstruction = AlfaWeb & "&CONNECTION=" & Replace(Replace(Conexao, Chr(13), ""), Chr(10), "")
    StrInstruction = StrInstruction & "&USER=" & Nnull(FAT_PAR!LOGIN_SID, 1)
    StrInstruction = StrInstruction & "&ACAO=SID.Srv.AltEmpFil&CodEmp=" & Empresa & "&CodFil=" & FIL
    
    i = 0
    Do
      Resp = LiberaFaturamentosa_xml.Inet1.OpenURL(UCase(StrInstruction))
      i = i + 1
    Loop Until i < 3 Or Mid(Resp, 3, 2) = "OK"

    If Mid(Resp, 3, 2) <> "OK" Then
        MsgBox "Erro: Erro de Conexão no Serviço SID, favor Contate o suporte." & vbCrLf & Resp, vbCritical
        Call P_Execute("Update " & Banco_Fatura & "faturanota set cancelada = 1, dt_cancelada = sysdate, motivo_cancela = 'ERRO SID.' where  id =" & Posicao)
        Exit Function
    End If





    StrInstruction = AlfaWeb & "&CONNECTION=" & Replace(Replace(Conexao, Chr(13), ""), Chr(10), "")
    StrInstruction = StrInstruction & "&USER=" & Nnull(FAT_PAR!LOGIN_SID, 1)
    StrInstruction = StrInstruction & "&ACAO=SID.Cli.GravarDefinicoes"
    StrInstruction = StrInstruction & "&CodCli=" & P_codcli
    StrInstruction = StrInstruction & "&CodRep=1"
    StrInstruction = StrInstruction & "&CODCPG=AV"
    StrInstruction = StrInstruction & "&VLRLIM=999999999"
    StrInstruction = StrInstruction & "&LIMAPR=S"
    'resp = Inet1.OpenURL(Replace(strInstruction, " ", "&nbsp;"))
    Resp = LiberaFaturamentosa_xml.Inet1.OpenURL(UCase(StrInstruction))

End Function
Sub Bloqueio_BL_Parceiro_old(Id_BL As Long)

Dim MotivoBloq As Integer
Dim Flag_Bloq_NVOCC As Byte
Dim Tbbl As ADODB.Recordset
Dim TbParceiro As ADODB.Recordset
Dim rsaux As ADODB.Recordset
Dim DesBloqueia As Boolean
Dim Bloqs As Boolean

DesBloqueia = False
MotivoBloq = 0
Flag_Bloq_NVOCC = 0
SQl = "SELECT "
SQl = SQl & "FLAG_NVOCC,PRIMEIRA_ENTRADA,FLAG_BLOQUEIO_MANUAL, BL_MASTER,"
SQl = SQl & "IMPORTADOR,DESPACHANTE FROM " & Banco_Sgipa & "TB_BL WHERE AUTONUM=" & Id_BL
PRSet Tbbl, SQl
If Not Tbbl.EOF Then
    If Nnull(Tbbl!PRIMEIrA_ENTRADA, 1) = "" And Nnull(Tbbl!BL_Master, 0) = 0 Then
        DesBloqueia = True
        If Nnull(Tbbl!Flag_Nvocc, 1) = 1 Then
            PRSet TbParceiro, "Select (flag_Fcl+Flag_Lcl) as flag_bloqueio_nvocc,motivo_bloqueio_nvocc from " & Banco_Sgipa & "tb_cad_parceiros where autonum =  " & Tbbl!Importador
            If Not TbParceiro.EOF Then
                If Nnull(TbParceiro!FLAG_BLOQUEIO_NVOCC, 0) > 0 Then
                    MotivoBloq = Nnull(TbParceiro!motivo_bloqueio_nvocc, 0)
                    Flag_Bloq_NVOCC = 1
                    DesBloqueia = False
                    GoSub Bloqueia
                End If
            End If
            TbParceiro.Close
        Else
            PRSet TbParceiro, "Select flag_bloqueio_importador,motivo_bloqueio_importador,motivo_bloqueio_nvocc ,(flag_Fcl+Flag_Lcl) as FLAG_BLOQUEIO_NVOCC from " & Banco_Sgipa & "tb_cad_parceiros where autonum =  " & Tbbl!Importador
            If Not TbParceiro.EOF Then
                If Nnull(TbParceiro!FLAG_BLOQUEIO_IMPORTADOR, 0) = 1 Then
                    MotivoBloq = Nnull(TbParceiro!MOTIVO_BLOQUEIO_IMPORTADOR, 0)
                    DesBloqueia = False
                    GoSub Bloqueia
                End If
                If Nnull(TbParceiro!FLAG_BLOQUEIO_IMPORTADOR, 0) = 0 And Nnull(TbParceiro!FLAG_BLOQUEIO_NVOCC, 0) > 0 Then
                    MotivoBloq = Nnull(TbParceiro!motivo_bloqueio_nvocc, 0)
                    DesBloqueia = False
                    GoSub Bloqueia
                End If
            End If
            TbParceiro.Close
        End If
        PRSet TbParceiro, "Select flag_bloqueio_despachante,motivo_bloqueio_despachante from " & Banco_Sgipa & "tb_cad_parceiros where autonum =  " & Tbbl!Despachante
        If Not TbParceiro.EOF Then
            If Nnull(TbParceiro!FLAG_BLOQUEIO_DESPACHANTE, 0) = 1 Then
                MotivoBloq = Nnull(TbParceiro!MOTIVO_BLOQUEIO_DESPACHANTE, 0)
                DesBloqueia = False
                GoSub Bloqueia
            End If
        End If
        TbParceiro.Close
    End If
    If DesBloqueia Then
       GoSub DesBloqueia
    End If
End If
Exit Sub

DesBloqueia:

SQl = " Select AUTONUM from " & Banco_Sgipa & "TB_HIST_BLOQUEIO Where  "
SQl = SQl & " OBS='BLOQUEIO MANUAL DE B/L' and "
SQl = SQl & " TIPO='MBL-L' AND "
SQl = SQl & " BL= " & Id_BL & "AND "
SQl = SQl & " AUDIT_CHEGADA=0 AND "
SQl = SQl & " AUDIT_ENVIO=0 AND "
SQl = SQl & " AUDIT_LIBERA=0 AND "
SQl = SQl & " FLAG_BLOQUEIA=1 AND "
SQl = SQl & " FLAG_ACEITO=0 AND "
SQl = SQl & " status='B'"
PRSet rsaux, SQl
If Not rsaux.EOF Then
   Call P_Bloqueio(Id_BL, 20, "L", Usuario_Sistema, wCpf_Usuario)
   SQl = "update " & Banco_Sgipa & "tb_hist_bloqueio set status='L' where autonum=" & rsaux!autonum
   Executa SQl  'z4
   SQl = "SELECT COUNT(*) FROM " & Banco_Sgipa & "TB_HIST_BLOQUEIO WHERE STATUS='B' and"
   SQl = SQl & " bl=" & Id_BL
   PRSet rsaux, SQl
   Bloqs = Nnull(rsaux.Fields(0), 0)
   If Bloqs = 0 Then
       Executa "UPDATE " & Banco_Sgipa & "TB_BL SET FLAG_BLOQUEIO_MANUAL = 0 WHERE AUTONUM = " & Id_BL
   End If
End If
Return


Bloqueia:
SQl = ""
SQl = SQl & "SELECT autonum "
SQl = SQl & "  FROM " & Banco_Sgipa & "tb_hist_bloqueio "
SQl = SQl & " WHERE obs = 'BLOQUEIO MANUAL DE B/L' "
SQl = SQl & "   AND tipo = 'MBL-L' "
SQl = SQl & "   AND bl = " & Id_BL & " "
SQl = SQl & "   AND cntr = 0 "
SQl = SQl & "   AND audit_chegada = 0 "
SQl = SQl & "   AND audit_envio = 0 "
SQl = SQl & "   AND audit_libera = 0 "
SQl = SQl & "   AND flag_bloqueia = 1 "
SQl = SQl & "   AND flag_aceito = 0 "
SQl = SQl & "   AND cod_motivo_bloqueio = " & MotivoBloq & " "
SQl = SQl & "   AND flag_nvocc = " & Flag_Bloq_NVOCC
SQl = SQl & "   AND status = 'B' "
PRSet rsaux, SQl
If rsaux.EOF Then
    Call P_Bloqueio(Id_BL, Trim(str(MotivoBloq)), "B", Usuario_Sistema, wCpf_Usuario)
    SQl = " INSERT INTO "
    SQl = SQl & " " & Banco_Sgipa & "TB_HIST_BLOQUEIO ( "
    SQl = SQl & " AUTONUM, "
    SQl = SQl & " DATA, "
    SQl = SQl & " HORA, "
    SQl = SQl & " OBS, "
    SQl = SQl & " TIPO, "
    SQl = SQl & " BL, "
    SQl = SQl & " CNTR, "
    SQl = SQl & " AUDIT_CHEGADA, "
    SQl = SQl & " AUDIT_ENVIO, "
    SQl = SQl & " AUDIT_LIBERA, "
    SQl = SQl & " FLAG_BLOQUEIA, "
    SQl = SQl & " FLAG_ACEITO, "
    SQl = SQl & " COD_MOTIVO_BLOQUEIO, "
    SQl = SQl & " USUARIO_SIS, "
    SQl = SQl & " flag_nvocc ,status"
    SQl = SQl & " ) VALUES ( "
    SQl = SQl & " " & Banco_Sgipa & "SEQ_HIST_BLOQUEIO.NEXTVAL, "
    SQl = SQl & " SYSDATE, "
    SQl = SQl & " '" & Format$(Now, "HH:MM:SS") & "', "
    SQl = SQl & " 'BLOQUEIO MANUAL DE B/L', "
    SQl = SQl & " 'MBL-L', "
    SQl = SQl & " " & Id_BL & ", "
    SQl = SQl & " '0', "
    SQl = SQl & " '0', "
    SQl = SQl & " '0', "
    SQl = SQl & " '0', "
    SQl = SQl & " '1', "
    SQl = SQl & " '0', "
    SQl = SQl & " '" & MotivoBloq & "', "
    SQl = SQl & " '" & Usuario_Sistema & "',"
    SQl = SQl & " " & Flag_Bloq_NVOCC
    SQl = SQl & ",'B') "
'   executa sql  'z4
    Executa "UPDATE " & Banco_Sgipa & "TB_BL SET FLAG_BLOQUEIO_MANUAL = 1 WHERE AUTONUM = '" & Id_BL & "'"
End If
Return

End Sub


                

Sub Verifica_Comissao_Cntr(Lote As Long, Optional Importador_de As Long = 0, Optional importador_para As Long = 0, Optional Indicador_de As Long = 0, Optional Indicador_para As Long = 0)
    Dim Rs2 As New ADODB.Recordset
    Dim Rs3 As New ADODB.Recordset

    SQL2 = "select c.autonum"
    SQL2 = SQL2 & " from " & Banco_Sgipa & "tb_bl a, " & Banco_Sgipa & "tb_amr_cntr_bl b, " & Banco_Sgipa & "tb_cntr_bl c"
    SQL2 = SQL2 & " where a.autonum = b.bl"
    SQL2 = SQL2 & " and b.cntr = c.autonum"
    SQL2 = SQL2 & " and a.autonum=" & Lote
    SQL2 = SQL2 & " and nvl(c.numerorp,0)<>0"
    PRSet Rs2, SQL2
    Do While Not Rs2.EOF
        SQL2 = "update " & Banco_Sgipa & "tb_cntr_bl set recalculo_comissao_cntr=1"
        SQL2 = SQL2 & " where autonum=" & Rs2!autonum
        Executa SQl  'z42
        Rs2.MoveNext
    Loop
    Rs2.Close
End Sub

Sub Envia_Email_Posicionamento(QualLote As Long, Flag_Executado As Integer)

Dim Mensagem$
Mensagem$ = ""
Dim Assunto$
Assunto$ = ""
            
'Se Existe GR/Calculo ou GR - Enviar E-mail
Set TbGr = New ADODB.Recordset
SQl = "Select seq_gr from " & Banco_Sgipa & "tb_gr_bl where bl=" & QualLote & " and status_gr in ('IM','GR')"
TbGr.Open SQl, db, 1, 1
If Not TbGr.EOF Then
    Mensagem$ = "Atenção, foi " & IIf(Flag_Executado = 1, "executado", "agendado") & " um posicionamento de contêiner do lote " & QualLote & " após fechamento de GR ! Ordem de servico: " & QualOS
    Assunto$ = "Posicionamento " & IIf(Flag_Executado = 1, "executado", "agendado") & " gerando Serviço Adicional - Gr gerada"
Else
    TbGr.Close
    Set TbGr = New ADODB.Recordset
    SQl = "Select autonum from tb_servicos_faturados where bl=" & QualLote & " and nvl(seq_gr,0)=0"
    TbGr.Open CONVERTE_SQL(SQl), db, 1, 1
    If Not TbGr.EOF Then
        Mensagem$ = "Atenção, foi " & IIf(Flag_Executado = 1, "executado", "agendado") & " um posicionamento de contêiner do lote " & QualLote & " após o cálculo ! Ordem de servico: " & QualOS
        Assunto$ = "Posicionamento " & IIf(Flag_Executado = 1, "executado", "agendado") & " gerando Serviço Adicional - Cálculo defasado."
    End If
End If
TbGr.Close
If Mensagem$ <> "" Then
    Set tb_email = New ADODB.Recordset
    tb_email.Open "Select * from " & Banco_Sgipa & "tb_email_vinculo", db, 1, 1
    While Not tb_email.EOF
        SQl = "INSERT INTO " & Banco_Sgipa & "tb_email_gr_faturada(autonum,para,endereco,assunto,"
        SQl = SQl & "mensagem,flag_vinculo) values (" & Banco_Sgipa & "seq_tb_email_gr_faturada.nextval,"
        SQl = SQl & "'" & Nnull(tb_email!Nome, 1) & "',"
        SQl = SQl & "'" & Nnull(tb_email!Email, 1) & "',"
        SQl = SQl & "'" & Assunto$ & "',"
        SQl = SQl & "'" & Mensagem$ & "',0)"
        Executa SQl  'z4
        tb_email.MoveNext
    Wend
    tb_email.Close
End If
End Sub
Public Function Nome_Empresa(Codigo, Tipo, Optional Fantasia As Boolean = False) As String

    Dim s As String
    s = ""
    If Tipo = "E" Then
        s = "select razao_social, nome_fantasia from " & Banco_Sgipa & "tb_empresas where autonum=" & Codigo
    ElseIf Tipo = "P" Then
        s = "select a.razao_social, nome_fantasia from " & Banco_Sgipa & "tb_empresas a, " & Banco_Operador & "tb_patios b"
        s = s & " where a.autonum = b.cod_empresa and b.autonum=" & Codigo
    End If
    If s <> "" Then
        Dim ss As New ADODB.Recordset
        PRSet ss, s
        If Not ss.EOF Then
            If Fantasia Then
                Nome_Empresa = ss!nome_fantasia
            Else
                Nome_Empresa = ss!razao_social
            End If
        Else
            Nome_Empresa = ""
        End If
        ss.Close
    Else
        Nome_Empresa = ""
    End If
        
End Function



Sub Bloqueia_Saida_Carga(Motivo As Byte, Lote As Long, Optional CNTR As Long = 0)

    SQl = "insert into " & Banco_Sgipa & "tb_liberacao_cargas"
    SQl = SQl & " (autonum,bl,dt_inclusao,usuario,cod_motivo,cntr"
    SQl = SQl & ") values ("
    SQl = SQl & "" & Banco_Sgipa & "seq_tb_liberacao_cargas.nextval"
    SQl = SQl & "," & Lote
    SQl = SQl & ",sysdate"
    SQl = SQl & "," & Cod_Usuario
    SQl = SQl & "," & Motivo
    If CNTR = 0 Then
        SQl = SQl & ",null"
    Else
        SQl = SQl & "," & CNTR
    End If
    SQl = SQl & ")"
    Executa SQl  'z4

End Sub

Private Sub Main()



 'Call Importa_BL("c:\microledtemp\xml\BANDEIRANTES_EDI_Import_Ref_163588MI.xml", "BANDEIRANTES_EDI_Import_Ref_163588MI")

Homologando = False

Dim nParametro As Double
Dim sParametro As String
nParametro = 212.2547
sParametro = Format$(nParametro, "00000.00000")
If InStr(sParametro, ".") > 0 Then
    MsgBox "Configurações Regionais - Formato numérico inválido para o sistema.", vbOKOnly + vbInformation, "SGIPA"
    nParametro = Empty
    sParametro = Empty
    End
Else
    nParametro = Empty
    sParametro = Empty
    Desconsolidando = False
    Inicio = True
    nSplash = 3
    If App.Title = "Atendimento a Despachantes" Then
        If DBConnect(db, , , , False) = True Then
'            sLicenciado = "CIA.BANDEIRANTES"
            frmSplash.Show 1
            Cadastrar_Menus 'inserido 05/04/2006 Ludmila
            Controlar_Acesso_Mnus
            Principal.Show
        Else
            End
        End If
    Else
'        sLicenciado = "TERMINAL DE CONTÊINERES DA MARGEM DIREITA DE SANTOS"
        Load Senha_E
        With Senha_E
            .lShowPatio = False
            .lShowServer = True
            .PreparaAcesso
            Senha_E.Show 1
            If Senha_E.lAcessou = True Then
               frmSplash.Show 1
               Load Principal
                
               If Cod_Usuario <> 1 Then
                 Cadastrar_Menus
                 Controlar_Acesso_Mnus
                End If
                
                Principal.Show
            Else
                End
            End If
        End With
    End If
End If
nParametro = Empty
sParametro = Empty
End Sub

Public Function CorrigeTexto(TEXTO As String) As String
TEXTO = Replace$(TEXTO, "'", Empty)
TEXTO = Replace$(TEXTO, Chr$(34), Empty)
TEXTO = Replace$(TEXTO, ">", Empty)
TEXTO = Replace$(TEXTO, "#", Empty)
TEXTO = Replace$(TEXTO, "@", Empty)
TEXTO = Replace$(TEXTO, "¨", Empty)
TEXTO = Replace$(TEXTO, "´", Empty)
TEXTO = Replace$(TEXTO, "`", Empty)
TEXTO = Replace$(TEXTO, "^", Empty)
TEXTO = Replace$(TEXTO, "~", Empty)
TEXTO = Replace$(TEXTO, "//", Empty)
TEXTO = Replace$(TEXTO, "\\", Empty)
TEXTO = Replace$(TEXTO, "|", Empty)
TEXTO = Replace$(TEXTO, "á", "A")
TEXTO = Replace$(TEXTO, "à", "A")
TEXTO = Replace$(TEXTO, "â", "A")
TEXTO = Replace$(TEXTO, "ã", "A")
TEXTO = Replace$(TEXTO, "ä", "A")
TEXTO = Replace$(TEXTO, "é", "E")
TEXTO = Replace$(TEXTO, "è", "E")
TEXTO = Replace$(TEXTO, "ê", "E")
TEXTO = Replace$(TEXTO, "ë", "E")
TEXTO = Replace$(TEXTO, "í", "I")
TEXTO = Replace$(TEXTO, "ì", "I")
TEXTO = Replace$(TEXTO, "î", "I")
TEXTO = Replace$(TEXTO, "ï", "I")
TEXTO = Replace$(TEXTO, "ó", "O")
TEXTO = Replace$(TEXTO, "ò", "O")
TEXTO = Replace$(TEXTO, "ô", "O")
TEXTO = Replace$(TEXTO, "õ", "O")
TEXTO = Replace$(TEXTO, "ö", "O")
TEXTO = Replace$(TEXTO, "ú", "U")
TEXTO = Replace$(TEXTO, "ù", "U")
TEXTO = Replace$(TEXTO, "û", "U")
TEXTO = Replace$(TEXTO, "ü", "U")
CorrigeTexto = Trim$(UCase$(TEXTO))
End Function




Public Sub GravaErros(Maquina As String, usuario As String, ERRO As String)
Dim nFile As Integer
Dim sFile As String
sFile = Trim(App.Path) & Trim(IIf(Right$(App.Path, 1) = "\", Empty, "\")) & Trim(Format$(Now, "DDMMYY")) & Trim("ERR.Log")
nFile = FreeFile
If Not Dir(sFile) = Empty Then
    Open sFile For Append As #nFile
Else
    Open sFile For Output As #nFile
End If
Print #nFile, Format$(Now, "dd/mm/yyyy") & "|" & Maquina & "|" & usuario & "|" & ERRO
Close nFile
nFile = Empty
sFile = Empty
End Sub


Public Function Proximo_Termo(Empresa As Byte) As Long

    Dim SeqN As String
    SeqN = "Seq_Termo_" & Empresa & "_" & Format(Year(Now), "0000")
    
If Banco_Em_Uso = "ORACLE" Then
    SqlT = "select * from ALL_sequences"
    SqlT = SqlT & " Where"
    SqlT = SqlT & " sequence_owner='SGIPA'"
    SqlT = SqlT & " and"
    SqlT = SqlT & " UPPER(sequence_name)='" & UCase(SeqN) & "'"
    
    Dim Rst As New ADODB.Recordset
    PRSet Rst, SqlT
    If Rst.EOF Then
        SqlT = "create sequence " & Banco_Sgipa & "" & SeqN & " start with 1"
        Executa SQl  'z4T
    End If
    Rst.Close
 End If
    
    
    SqlT = "select " & Banco_Sgipa & "" & SeqN & ".nextval from dual"
    PRSet Rst, SqlT
    Proximo_Termo = Rst.Fields(0)
    
End Function

Public Function ValidaHoje() As Boolean

    Dim oRS As ADODB.Recordset
    db.Close
    DBConnect db, , , , (TipodeBanco = 1)
    
    If Banco_Em_Uso = "ORACLE" Then
        PRSet oRS, "SELECT TO_CHAR(SYSDATE,'DD/MM/YYYY HH24:MI:SS') AS HOJE FROM DUAL", 0
    Else
        PRSet oRS, "SELECT CONVERT(VARCHAR(23), GETDATE(), 103) + ' ' + CONVERT(VARCHAR, GETDATE(), 108) AS HOJE", 0
    End If
    
    'ValidaHoje = Not (oRS.Fields("HOJE") <> Format$(Now, "DD/MM/YYYY HH"))
    ValidaHoje = Not (Abs(DateDiff("n", oRS.Fields("HOJE"), Now)) > 5) '5 MINUTOS
    oRS.Close
    Set oRS = Nothing

End Function
Public Function Maquina() As String
Dim sRET As String
sRET = String(30, "X")
GetComputerName sRET, 30
sRET = Left(sRET, 30)
Maquina = Trim(UCase(Replace(Replace(Replace(sRET, "X", Empty), Chr(0), Empty), Chr(1), Empty)))
sRET = Empty
End Function

Public Function CodAdmin() As Long
'Retorna sempre o código do usuário ADMIN, senha ADMIN
On Error GoTo CodAdminErr:
Dim oRS As ADODB.Recordset
Screen.MousePointer = vbHourglass
PRSet oRS, "SELECT AUTONUM FROM " & Banco_Sgipa & "TB_CAD_USUARIOS WHERE USUARIO = 'ADMIN' AND SENHA = 'ADMIN'", 0
If Not oRS.EOF Then CodAdmin = oRS.Fields(0).Value
oRS.Close
GoTo CodAdminExit:
CodAdminErr:
MsgBox Err.Description
Err.Clear
CodAdminExit:
Set oRS = Nothing
Screen.MousePointer = vbDefault
End Function

Public Function Eferiado(ByVal Valor As String, Optional ByVal ValidaDomingos As Boolean = False) As Boolean
Dim oTB As ADODB.Recordset
Dim lRET As Boolean
Valor = Format$(Valor, "DD/MM/YYYY")
If Not IsDate(Valor) = True Then
    MsgBox "DATA INVÁLIDA."
    lRET = True
Else
    PRSet oTB, "SELECT * FROM " & Banco_Operador & "TB_FERIADOS WHERE DATA = TO_DATE('" & Valor & "','DD/MM/YYYY')"
    If Not oTB.EOF Then
        lRET = True
        oTB.Close
    Else
        lRET = False
        oTB.Close
        If ValidaDomingos = True Then
            If VBA.Weekday(Valor) = VbDayOfWeek.vbSunday Then
               lRET = True
            End If
        End If
    End If
End If
Eferiado = lRET
Set oTB = Nothing
lRET = Empty
End Function






Sub Imprime_Grci(IdCntr As Long)
Set Tbh = New ADODB.Recordset
Tbh.Open "SELECT CONTROLE_GRCI FROM " & Banco_Sgipa & "TB_CNTR_BL WHERE AUTONUM=" & IdCntr, db, 3, 3
If Nnull(Tbh!CONTROLE_GRCI, 0) = 0 Then
    Set Tbg = New ADODB.Recordset
    Tbg.Open "SELECT MAX(CONTROLE_GRCI) AS QUAL FROM " & Banco_Sgipa & "TB_CNTR_BL", db, 3, 3
    Executa "UPDATE " & Banco_Sgipa & "TB_CNTR_BL SET CONTROLE_GRCI=" & Nnull(Tbg!Qual, 0) + 1 & " WHERE AUTONUM=" & IdCntr
End If
Set Tbh = New ADODB.Recordset
SQl = "SELECT A.ID_CONTEINER,A.GMCI,A.CONTROLE_GRCI,A.VIAGEM,N.DESCR AS NAVIO"
SQl = SQl & ",A.TAMANHO,A.REGIME,T.CODIGO AS TIPO"

'Sql = Sql & " FROM " & Banco_Sgipa & "TB_CNTR_BL A"
'Sql = Sql & " ," & Banco_Sgipa & "DTE_TB_VIAGENS V"
'Sql = Sql & " ," & Banco_Sgipa & "DTE_TB_NAVIOS N"
'Sql = Sql & " ," & Banco_Sgipa & "DTE_TB_TIPOS_CONTEINER T"
'Sql = Sql & " WHERE"
'Sql = Sql & " A.VIAGEM=V.VIAGEM"
'Sql = Sql & " AND V.NAVIO=N.CODE(+)"
'Sql = Sql & " AND A.TIPO=T.CODE(+)"
'Sql = Sql & " AND A.AUTONUM=" & Nnull(IdCntr, 0)

SQl = SQl & " FROM " & Banco_Sgipa & "TB_CNTR_BL A"
SQl = SQl & " INNER JOIN " & Banco_Sgipa & "DTE_TB_VIAGENS V ON A.VIAGEM=V.VIAGEM"
SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "DTE_TB_NAVIOS N ON V.NAVIO=N.CODE"
SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "DTE_TB_TIPOS_CONTEINER T ON A.TIPO=T.CODE"
SQl = SQl & " WHERE"
SQl = SQl & " A.AUTONUM=" & Nnull(IdCntr, 0)

Tbh.Open SQl, db, 3, 3

Set Tbbl = New ADODB.Recordset
SQl = "SELECT PO.DESCR AS PORTO_ORIGEM,PD.DESCR AS PORTO_DESTINO,"
SQl = SQl & "D.DESCR AS AGENCIA "

'Sql = Sql & " FROM " & Banco_Sgipa & "TB_BL A"
'Sql = Sql & " ," & Banco_Sgipa & "TB_AMR_CNTR_BL B"
'Sql = Sql & " ," & Banco_Sgipa & "DTE_TB_AGENCIAS D"
'Sql = Sql & " ," & Banco_Sgipa & "DTE_TB_PORTOS PO"
'Sql = Sql & " ," & Banco_Sgipa & "DTE_TB_PORTOS PD"
'Sql = Sql & " WHERE"
'Sql = Sql & " B.BL=A.AUTONUM"
'Sql = Sql & " AND A.PORTO_ORIGEM=PO.CODE(+)"
'Sql = Sql & " AND A.PORTO_DESTINO=PD.CODE(+)"
'Sql = Sql & " AND A.AGENCIA=D.CODE(+)"
'Sql = Sql & " AND B.CNTR=" & IdCntr

SQl = SQl & " FROM " & Banco_Sgipa & "TB_BL A"
SQl = SQl & " INNER JOIN " & Banco_Sgipa & "TB_AMR_CNTR_BL B ON B.BL=A.AUTONUM"
SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "DTE_TB_AGENCIAS D ON A.AGENCIA=D.CODE"
SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "DTE_TB_PORTOS PO ON A.PORTO_ORIGEM=PO.CODE"
SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "DTE_TB_PORTOS PD ON A.PORTO_DESTINO=PD.CODE"
SQl = SQl & " WHERE"
SQl = SQl & " B.CNTR=" & IdCntr

Tbbl.Open SQl, db, 3, 3

If Tbh.EOF Then Exit Sub

'With Principal.CrystalReport2
'    .ReportFileName = App.Path & "\" & "guiagrci.rpt"
'    .Formulas(0) = "AGENCIA='" & Nnull(TbBl!agencia, 1) & "'"
'    .Formulas(1) = "CONTAINER='" & Nnull(tbh!Id_Conteiner, 1) & "'"
'    .Formulas(2) = "CONTROLE='" & Format(Nnull(tbh!CONTROLE_GRCI, 1), "000000") & "'"
'    .Formulas(3) = "DIMENSAO='" & Nnull(tbh!Tamanho, 0) & "'"
'    .Formulas(4) = "GRCI='" & Nnull(tbh!GMCI, 1) & "'"
'    .Formulas(5) = "NAVIO='" & Nnull(tbh!Navio, 1) & "'"
'    .Formulas(6) = "PORTO='" & Nnull(TbBl!Porto_Origem, 1) & "'"
'    .Formulas(7) = "PORTOD='" & Nnull(TbBl!PORTO_DESTINO, 1) & "'"
'    .Formulas(8) = "REGIME='" & Nnull(tbh!REGIME, 1) & "'"
'    .Formulas(9) = "TIPO='" & Nnull(tbh!Tipo, 1) & "'"
'    .Formulas(10) = "VIAGEM='" & Nnull(tbh!Viagem, 1) & "'"
'    .Action = 1
'    For I = 0 To 10
'        .Formulas(I) = ""
'    Next I
'End With

Tbh.Close
Tbbl.Close
End Sub
Sub REL_GMCI(DT As String, GMCI As String, Viagem As String)

Dim Tbh As New ADODB.Recordset
Dim GmciAnt As String
'************ GMCI ************
'* ESC(0) *
Dim l(1 To 66) As String
For i = 1 To 66
    l(i) = ""
Next
Dim Filtro As String

Filtro = ""
Filtro = Filtro & " AND A.VIAGEM='" & Viagem & "'"

If GMCI <> "" Then
    Filtro = Filtro & " AND A.GMCI = '" & GMCI & "'"
End If

If DT <> "" Then
    Filtro = Filtro & " AND A.dt = '" & DT & "'"
End If


'ZOLD Set TBH = New ADODB.Recordset
SQl = "SELECT a.autonum,A.VIAGEM,A.GMCI,a.dt,A.TARA, C.NOMENAVIO AS NAVIO, N.BRUTO,F.DESCR AS P_ORIGEM, G.DESCR AS P_DESTINO,"
SQl = SQl & " B.DESCR AS TIPO_C, A.ID_CONTEINER, A.TAMANHO,  A.REGIME, A.LACRE_ORIGEM,"
SQl = SQl & " J.DESCR AS ARM_ATRACACAOP, C.DT_ATRACACAOP,  H.CODE, H.DESCR AS AGENCIA, a.patio "
'Sql = Sql & " FROM"
'Sql = Sql & " " & Banco_Sgipa & "TB_CNTR_BL A, " & Banco_Sgipa & "DTE_TB_TIPOS_CONTEINER B, " & Banco_Sgipa & "DTE_TB_VIAGENS C,"
'Sql = Sql & " " & Banco_Sgipa & "TB_BL E, " & Banco_Sgipa & "DTE_TB_PORTOS F, " & Banco_Sgipa & "DTE_TB_PORTOS G, " & Banco_Sgipa & "DTE_TB_AGENCIAS H,"
'Sql = Sql & " " & Banco_Sgipa & "DTE_TB_ARMAZENS J, " & Banco_Sgipa & "TB_AMR_CNTR_BL N WHERE A.TIPO = B.CODE(+) AND A.VIAGEM = C.VIAGEM AND"
'Sql = Sql & " A.AUTONUM = N.CNTR AND E.AUTONUM=N.BL AND E.PORTO_ORIGEM = F.CODE(+) AND"
'Sql = Sql & " E.PORTO_DESTINO = G.CODE(+)  AND C.ARM_ATRACACAOP = J.CODE(+)"
'Sql = Sql & " AND E.AGENCIA = H.CODE(+) AND A.GMCI IS NOT NULL " & Filtro & " ORDER BY A.GMCI"

SQl = SQl & " FROM "
SQl = SQl & " " & Banco_Sgipa & "TB_CNTR_BL A "
SQl = SQl & " LEFT JOIN  " & Banco_Sgipa & "DTE_TB_TIPOS_CONTEINER B ON A.TIPO = B.CODE "
SQl = SQl & " INNER JOIN " & Banco_Sgipa & "DTE_TB_VIAGENS C ON A.VIAGEM = C.VIAGEM "
SQl = SQl & " INNER JOIN  " & Banco_Sgipa & "TB_AMR_CNTR_BL N ON A.AUTONUM = N.CNTR  "
SQl = SQl & " INNER JOIN " & Banco_Sgipa & "TB_BL E ON E.AUTONUM=N.BL "
SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "DTE_TB_PORTOS F ON E.PORTO_ORIGEM = F.CODE "
SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "DTE_TB_PORTOS G ON  E.PORTO_DESTINO = G.CODE "
SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "DTE_TB_AGENCIAS H ON E.AGENCIA = H.CODE "
SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "DTE_TB_ARMAZENS J ON C.ARM_ATRACACAOP = J.CODE "


SQl = SQl & " WHERE A.GMCI IS NOT NULL " & Filtro & " ORDER BY A.GMCI"


'ZOLD TBH.Open Sql, Db, 3, 3
PRSet Tbh, SQl

If Tbh.EOF Then
    MsgBox "Nenhum Registro", vbCritical
    Exit Sub
End If
Open App.Path & "\GMCI.LST" For Output As #1
Do
    If Tbh!GMCI <> GmciAnt Then
        l(8) = s(60) & Tbh!GMCI
        Dim DigCodesp As Integer
        DigCodesp = Calcula_Digito_Codesp(Tbh!DT, Tbh!Id_Conteiner, Tbh!GMCI, Tbh!Tamanho)
        l(10) = s(60) & DigCodesp
        l(13) = s(3) & Left$(Tbh!Navio & s(24), 24) & s(1) & Left$(Tbh!P_ORIGEM & s(25), 25) & s(5) & Left$(Tbh!Id_Conteiner, 10) & s(11) & Right$(Tbh!Id_Conteiner, 1)
        
        l(15) = s(3) & Left$(Tbh!agencia & s(45), 45) & s(1) & Left$(Tbh!Code & s(7), 7) & s(2) & Left$(Tbh!Tamanho & s(2), 2) & s(5) & Tbh!TIPO_C
        Dim Manifesto As String
        Manifesto = ""
        Set TBH2 = New ADODB.Recordset
        TBH2.Open "SELECT MANIFESTO FROM " & Banco_Sgipa & "DTE_TB_MANIFESTOS WHERE VIAGEM='" & Tbh!Viagem & "'", db, 3, 3
        If Not TBH2.EOF Then Manifesto = Nnull(TBH2!Manifesto, 1)
        TBH2.Close
        l(17) = s(3) & Left$(Tbh!DT & s(15), 15) & s(2) & Left$(Manifesto & s(7), 7)
        
        If Not IsNull(Tbh!Dt_AtracacaoP) Then
            l(17) = l(17) & s(9) & s(2) & Mid$(Tbh!Dt_AtracacaoP, 1, 1) & s(1) & Mid$(Tbh!Dt_AtracacaoP, 2, 1) & s(2) & Mid$(Tbh!Dt_AtracacaoP, 4, 1) & s(1) & Mid$(Tbh!Dt_AtracacaoP, 5, 1) & s(2) & Mid$(Tbh!Dt_AtracacaoP, 7, 1) & s(1) & Mid$(Tbh!Dt_AtracacaoP, 8, 1) & s(1) & Mid$(Tbh!Dt_AtracacaoP, 9, 1) & s(1) & Mid$(Tbh!Dt_AtracacaoP, 10, 1)
        Else
            l(17) = l(17) & s(11) & s(17)
        End If
        Set TBH2 = New ADODB.Recordset
        TBH2.Open "Select sum (a.bruto) as somabruto from " & Banco_Sgipa & "TB_AMR_CNTR_BL a," & Banco_Sgipa & "TB_BL b where a.bl=b.autonum and b.flag_ativo=1 and a.cntr=" & Tbh!autonum, db, 3, 3
        l(17) = l(17) & s(6) & Tbh!Regime & s(5) & Format(Nnull(TBH2!somaBruto, 0) + Nnull(Tbh!Tara, 0), "###,###,##0.000")
        l(20) = s(3) & Left$(Tbh!P_Destino & s(25), 25)
        If Not IsNull(Tbh!Arm_AtracacaoP) Then l(24) = l(24) & s(41) & Mid$(Tbh!Arm_AtracacaoP, 1, 1) & s(1) & Mid$(Tbh!Arm_AtracacaoP, 2, 1) & s(1) & Mid$(Tbh!Arm_AtracacaoP, 3, 1)
        'l(25) = s(4) & Mid$(TBh!LACRE_ORIGEM, 1, 1) & s(1) & Mid$(TBh!LACRE_ORIGEM, 2, 1) & s(1) & Mid$(TBh!LACRE_ORIGEM, 3, 1) & s(1)
        'l(25) = l(25) & Mid$(TBh!LACRE_ORIGEM, 4, 1) & s(1) & Mid$(TBh!LACRE_ORIGEM, 5, 1) & s(1) & Mid$(TBh!LACRE_ORIGEM, 6, 1) & s(1)
        'l(25) = l(25) & Mid$(TBh!LACRE_ORIGEM, 7, 1) & s(1) & Mid$(TBh!LACRE_ORIGEM, 8, 1) & s(1) & Mid$(TBh!LACRE_ORIGEM, 9, 1) & s(1)
        'l(25) = l(25) & Mid$(TBh!LACRE_ORIGEM, 10, 1) & s(1) & Mid$(TBh!LACRE_ORIGEM, 11, 1)
        l(26) = s(65) & "IPA"
        l(27) = s(58) & Nome_Empresa(Tbh!Patio, "P", True) '"CIA.BANDEIRANTES"
        l(47) = s(3) & "NOME DO MOTORISTA:"
        l(50) = s(3) & "MAX GROSS:"
        l(51) = s(3) & "TARA     :"
        l(52) = s(3) & "ISO      :"
        For i = 7 To 66
            l(i) = UCase(l(i))
            Printer.Print l(i)
            'Print #1, l(i)
        Next
        Printer.EndDoc
        'Print #1, Chr$(12)
    End If
    GmciAnt = Tbh!GMCI
    Tbh.MoveNext
Loop Until Tbh.EOF
Close #1
Tbh.Close
End Sub
Public Sub REL_GMVI(DT As String, GMVI As String, Viagem As String, Controle As String, Optional Tipo As Integer = 1)
Dim SQl     As String
Dim Tbh As New ADODB.Recordset
Dim nRecs   As Long
SQl = Empty
SQl = SQl & " INSERT INTO "
SQl = SQl & " " & Banco_Sgipa & "TB_EMISSAO_GUIAS ( "
SQl = SQl & " GUIA, "
SQl = SQl & " CONTROLE, "
SQl = SQl & " TIPO, "
SQl = SQl & " DT_EMISSAO, "
SQl = SQl & " USUARIO "
SQl = SQl & " ) VALUES ( "
SQl = SQl & " '" & GMVI & "', "
SQl = SQl & " '" & Controle & "', "
SQl = SQl & " 'G" & IIf(Tipo = 1, "M", "R") & "VI', "
SQl = SQl & " SYSDATE, "
SQl = SQl & " '" & tbSenha!autonum & "' "
SQl = SQl & " ) "
Executa SQl, nRecs

Exit Sub


SQl = Empty
'************ GMVI ************
'* ESC(0) *
Dim GMVIant As String
Dim l(1 To 47) As String
For i = 1 To 47
    l(i) = ""
Next
Dim Filtro As String
Dim nControle As Long
Filtro = ""
Filtro = Filtro & " AND A.VIAGEM='" & Viagem & "'"

If GMVI <> "" Then
    Filtro = Filtro & " AND A.GMVI = '" & GMVI & "'"
End If

If DT <> "" Then
    Filtro = Filtro & " AND A.DT = '" & DT & "'"
End If

'ZOLD Set TBH = New ADODB.Recordset

SQl = "SELECT a.DT, a.viagem,A.GMVI,A.OBS_MANUSEIO, I.DESCR AS PORTO_O, H.NUMERO, A.QUANTIDADE, A.PESO_BRUTO,J.DESCR AS EMBALAGEM,h.patio"
SQl = SQl & " ,A.MARCA, A.MERCADORIA, A.GENERO, C.NOMENAVIO AS NAVIO, F.DESCR AS AGENCIA, F.CODE AS COD_AGENCIA, C.DT_ATRACACAOP, E.CODE AS ARM_ATRACACAOP"

'Sql = Sql & " FROM " & Banco_Sgipa & "TB_CARGA_SOLTA A, " & Banco_Sgipa & "DTE_TB_VIAGENS C,"
'Sql = Sql & " " & Banco_Sgipa & "DTE_TB_ARMAZENS E, " & Banco_Sgipa & "DTE_TB_AGENCIAS F, " & Banco_Sgipa & "TB_BL H, " & Banco_Sgipa & "DTE_TB_PORTOS I," & Banco_Sgipa & "DTE_TB_EMBALAGENS J "
'Sql = Sql & " WHERE A.VIAGEM = C.VIAGEM  AND"
'Sql = Sql & " H.AGENCIA = F.CODE(+) AND A.BL = H.AUTONUM AND H.PORTO_ORIGEM = I.CODE(+) AND"
'Sql = Sql & "  C.ARM_ATRACACAOP = E.CODE(+) AND A.EMBALAGEM=J.CODE(+) and A.GMVI IS NOT NULL"
'Sql = Sql & Filtro & " ORDER BY A.GMVI"


SQl = SQl & " FROM " & Banco_Sgipa & "TB_CARGA_SOLTA A INNER JOIN " & Banco_Sgipa & "DTE_TB_VIAGENS C ON  A.VIAGEM = C.VIAGEM "
SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "DTE_TB_ARMAZENS E ON  C.ARM_ATRACACAOP = E.CODE "

SQl = SQl & " INNER JOIN " & Banco_Sgipa & "TB_BL H ON A.BL = H.AUTONUM "
SQl = SQl & " LEFT JOIN  " & Banco_Sgipa & "DTE_TB_AGENCIAS F ON  H.AGENCIA = F.CODE "

SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "DTE_TB_PORTOS I ON H.PORTO_ORIGEM = I.CODE "
SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "DTE_TB_EMBALAGENS J ON  A.EMBALAGEM=J.CODE "
SQl = SQl & " WHERE"
SQl = SQl & " A.GMVI IS NOT NULL "
SQl = SQl & Filtro & " ORDER BY A.GMVI"




'ZOLD TBH.Open Sql, Db, 3, 3

PRSet Tbh, SQl


If Tbh.EOF Then
    MsgBox "Nehum registro", vbCritical
    Exit Sub
End If
'Open App.Path & "\gmvi.lst" For Output As #1
Do
    If GMVIant <> Tbh!GMVI Then
        Do
            X = InputBox("Informe o número de vias da GMVI:" & Tbh!GMVI, , "1")
            If X = "" Then Exit Sub
        Loop While Not IsNumeric(X)
        If Val(X) <> 0 Then
            For nControle = 1 To Val(X)
                If IsNull(Tbh!Dt_AtracacaoP) Then
                    Dt_AtracacaoP = "00/00/0000"
                Else
                    Dt_AtracacaoP = Format(Tbh!Dt_AtracacaoP, "DD/MM/YYYY")
                End If
                l(15) = s(2) & Left$(Tbh!DT & s(13), 13) & s(3) & Left$(Tbh!GMVI & s(13), 13) & s(3) & Format(nControle, "000") & s(5) & "010" & Format(nControle, "000") & "/2001"

                Dim Manifesto As String
                Manifesto = ""
                Set TBH2 = New ADODB.Recordset
                TBH2.Open "Select manifesto from " & Banco_Sgipa & "DTE_TB_MANIFESTOS where viagem=" & Tbh!Viagem, db, 3, 3
                If Not TBH2.EOF Then Manifsto = Nnull(TBH2!Manifesto, 1)
                l(18) = s(2) & Left$(Tbh!porto_o & s(25), 25) & s(22) & Manifesto
                l(20) = s(2) & Left$(Tbh!Numero & s(16), 16) & s(9) & Left$(Tbh!quantidade & s(6), 6) & s(18) & Format(Left$(Tbh!PESO_BRUTO & s(8), 8), "###,##0.000")
                l(22) = s(2) & Tbh!Marca
                l(24) = s(2) & Left$(Tbh!MERCADORIA & s(25), 25) & s(24) & Nnull(Tbh!embalagem, 1)
                l(27) = s(2) & Left$(Tbh!Navio & s(21), 21) & s(2) & Left$(Tbh!agencia & s(38), 38) & s(2) & Tbh!COD_AGENCIA
                l(29) = s(3) & Mid$(Dt_AtracacaoP, 1, 1) & s(1) & Mid$(Dt_AtracacaoP, 2, 1) & s(2) & Mid$(Dt_AtracacaoP, 4, 1) & s(1) & Mid$(Dt_AtracacaoP, 5, 1) & s(2) & Mid$(Dt_AtracacaoP, 7, 1) & s(1) & Mid$(Dt_AtracacaoP, 8, 1) & s(1) & Mid$(Dt_AtracacaoP, 9, 1) & s(1) & Mid$(Dt_AtracacaoP, 10, 1)
                l(45) = s(3) & "IPA - " & Nome_Empresa(Tbh!Patio, "P", True) & " PATIO " & Tbh!Patio
                l(47) = s(3) & Nnull(Tbh!obs_manuseio, 1)
            Next nControle
        End If
    End If
    GMVIant = Tbh!GMVI
    Tbh.MoveNext
Loop Until Tbh.EOF
Tbh.Close
End Sub

Sub REL_GMCI_GRCI(DT As String, GMCI As String, Viagem As String, Tipo As Integer)
Dim GmciAnt    As String
Dim nA         As Integer
Dim l(1 To 66) As String
Dim Filtro     As String
Dim Tbh        As ADODB.Recordset
Dim TBH2       As ADODB.Recordset
Dim DigCodesp  As Integer
Dim Manifesto  As String

For i = 1 To 66
    l(i) = ""
Next

Filtro = ""
Filtro = Filtro & " AND A.VIAGEM='" & Viagem & "'"

If GMCI <> "" Then
    Filtro = Filtro & " AND A.GMCI = '" & GMCI & "'"
End If

If DT <> "" Then
    Filtro = Filtro & " AND A.dt = '" & DT & "'"
End If

SQl = "SELECT a.autonum,A.VIAGEM,A.GMCI,a.dt,A.TARA, C.NOMENAVIO AS NAVIO, N.BRUTO,F.DESCR AS P_ORIGEM, G.DESCR AS P_DESTINO,"
SQl = SQl & " B.DESCR AS TIPO_C, A.ID_CONTEINER, A.TAMANHO,  A.REGIME, A.LACRE_ORIGEM,a.patio,"
SQl = SQl & " J.DESCR AS ARM_ATRACACAOP, C.DT_ATRACACAOP,  H.CODE, H.DESCR AS AGENCIA "

'Sql = Sql & " FROM "
'Sql = Sql & " " & Banco_Sgipa & "TB_CNTR_BL A, " & Banco_Sgipa & "DTE_TB_TIPOS_CONTEINER B, " & Banco_Sgipa & "DTE_TB_VIAGENS C,"
'Sql = Sql & " " & Banco_Sgipa & "TB_BL E, " & Banco_Sgipa & "DTE_TB_PORTOS F, " & Banco_Sgipa & "DTE_TB_PORTOS G, " & Banco_Sgipa & "DTE_TB_AGENCIAS H,"
'Sql = Sql & " " & Banco_Sgipa & "DTE_TB_ARMAZENS J, " & Banco_Sgipa & "TB_AMR_CNTR_BL N WHERE A.TIPO = B.CODE(+) AND A.VIAGEM = C.VIAGEM AND"
'Sql = Sql & " A.AUTONUM = N.CNTR AND E.AUTONUM=N.BL AND E.PORTO_ORIGEM = F.CODE(+) AND"
'Sql = Sql & " E.PORTO_DESTINO = G.CODE(+)  AND C.ARM_ATRACACAOP = J.CODE(+)"
'Sql = Sql & " AND E.AGENCIA = H.CODE(+) AND A.GMCI IS NOT NULL " & Filtro & " ORDER BY A.GMCI"


SQl = SQl & " FROM " & Banco_Sgipa & "TB_CNTR_BL A LEFT JOIN " & Banco_Sgipa & "DTE_TB_TIPOS_CONTEINER B ON A.TIPO = B.CODE "
SQl = SQl & " INNER JOIN " & Banco_Sgipa & "DTE_TB_VIAGENS C ON A.VIAGEM = C.VIAGEM "
SQl = SQl & " INNER JOIN " & Banco_Sgipa & "TB_AMR_CNTR_BL N ON A.AUTONUM = N.CNTR  "
SQl = SQl & " INNER JOIN " & Banco_Sgipa & "TB_BL E ON  E.AUTONUM=N.BL  "
SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "DTE_TB_PORTOS F ON E.PORTO_ORIGEM = F.CODE "
SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "DTE_TB_PORTOS G ON E.PORTO_DESTINO = G.CODE "
SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "DTE_TB_AGENCIAS H ON E.AGENCIA = H.CODE"
SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "DTE_TB_ARMAZENS J ON C.ARM_ATRACACAOP = J.CODE "
SQl = SQl & " WHERE A.GMCI IS NOT NULL " & Filtro & " ORDER BY A.GMCI"

PRSet Tbh, SQl, 0

If Tbh.EOF Then
    MsgBox "Nenhum Registro", vbCritical
    Exit Sub
End If

'Open "lpt1" For Output As #1

Do
    If Tbh!GMCI <> GmciAnt Then

        If Tipo = 1 Then l(3) = s(23) & "GUIA DE MOVIMENTACAO DE CONTAINER" & s(5) & Tbh!GMCI
        If Tipo = 2 Then l(3) = s(25) & "GUIA DE REMOCAO DE CONTAINER" & s(8) & Tbh!GMCI

        DigCodesp = Calcula_Digito_Codesp(Tbh!DT, Tbh!Id_Conteiner, Tbh!GMCI, Tbh!Tamanho)
        l(4) = s(32) & "I M P O R T A C A O" & s(10) & DigCodesp
        
        'l(13) = s(3) & Left$(tbh!Navio & s(24), 24) & s(1) & Left$(tbh!P_ORIGEM & s(25), 25) & s(5) & Left$(tbh!Id_Conteiner, 10) & s(11) & Right$(tbh!Id_Conteiner, 1)
        'l(3) = s(61) & tbh!GMCI
        'DigCodesp = Calcula_Digito_Codesp(tbh!DT, tbh!Id_Conteiner, tbh!GMCI, tbh!tamanho)
        'l(5) = s(61) & DigCodesp
        
        l(8) = s(2) & Left$(Tbh!Navio & s(25), 25) & s(1) & Left$(Tbh!P_ORIGEM & s(31), 31) & s(1) & Left$(Tbh!Id_Conteiner, 10) & s(7) & Right$(Tbh!Id_Conteiner, 1)
        l(10) = s(2) & Left$(Tbh!agencia & s(45), 45) & s(1) & Left$(Tbh!Code & s(7), 7) & s(4) & Left$(Tbh!Tamanho & s(2), 2) & s(6) & Tbh!TIPO_C
        Manifesto = ""
        
        Set TBH2 = New ADODB.Recordset
        TBH2.Open "SELECT MANIFESTO FROM " & Banco_Sgipa & "DTE_TB_MANIFESTOS WHERE VIAGEM='" & Tbh!Viagem & "'", db, 3, 3
        If Not TBH2.EOF Then Manifesto = Nnull(TBH2!Manifesto, 1)
        TBH2.Close
        l(12) = s(2) & Left$(Tbh!DT & s(15), 15) & s(2) & Left$(Manifesto & s(7), 7)
        If Not IsNull(Tbh!Dt_AtracacaoP) Then
            l(12) = l(12) & s(11) & Mid$(Format$(Tbh!Dt_AtracacaoP, "dd/mm/yyyy"), 1, 1) & s(1) & Mid$(Format$(Tbh!Dt_AtracacaoP, "dd/mm/yyyy"), 2, 1) & s(1) & Mid$(Format$(Tbh!Dt_AtracacaoP, "dd/mm/yyyy"), 4, 1) & s(1) & Mid$(Format$(Tbh!Dt_AtracacaoP, "dd/mm/yyyy"), 5, 1) & s(2) & Mid$(Format$(Tbh!Dt_AtracacaoP, "dd/mm/yyyy"), 7, 1) & s(1) & Mid$(Format$(Tbh!Dt_AtracacaoP, "dd/mm/yyyy"), 8, 1) & s(1) & Mid$(Format$(Tbh!Dt_AtracacaoP, "dd/mm/yyyy"), 9, 1) & s(1) & Mid$(Format$(Tbh!Dt_AtracacaoP, "dd/mm/yyyy"), 10, 1)
        Else
            l(12) = l(12) & s(23)
        End If
        
        PRSet TBH2, "Select sum (a.bruto) as somabruto from " & Banco_Sgipa & "TB_AMR_CNTR_BL a," & Banco_Sgipa & "TB_BL b where a.bl=b.autonum and b.flag_ativo=1 and a.cntr=" & Tbh!autonum & " ", 0
        
        l(12) = l(12) & s(3) & Tbh!Regime & s(9) & Format(Nnull(TBH2!somaBruto, 0) + Nnull(Tbh!Tara, 0), "###,###,##0.000")
        l(15) = s(2) & Left$(Tbh!P_Destino & s(30), 30)
        If Not IsNull(Tbh!Arm_AtracacaoP) Then l(15) = l(15) & " " & Mid$(Tbh!Arm_AtracacaoP, 1, 1) & s(1) & Mid$(Tbh!Arm_AtracacaoP, 2, 1) & s(1) & Mid$(Tbh!Arm_AtracacaoP, 3, 1)
        
        l(20) = s(55) & Nome_Empresa(Tbh!Patio, "P", True)
        l(22) = s(55) & "    PATIO " & Tbh!Patio
        
        
        l(42) = s(3) & "NOME DO MOTORISTA:"
        l(43) = s(3) & "MAX GROSS:"
        l(44) = s(3) & "TARA     :"
        l(45) = s(3) & "ISO      :"
        For i = 3 To 63
            l(i) = UCase(l(i))
           'Print #1, l(I)
        Next
        NARQUIVO = Empty
    End If
    GmciAnt = Tbh!GMCI
    Tbh.MoveNext
Loop Until Tbh.EOF


Open "lpt1:" For Output As #1
For i = 3 To 63
    Print #1, l(i)
Next
Print #1, Chr$(12)
Close #1
Tbh.Close
End Sub


Sub REL_GRVI(DT As String, GMVI As String, Viagem As String)
'************ GMVI ************
'* ESC(0) *
Dim GMVIant As String
Dim Tbh As New ADODB.Recordset

Dim l(1 To 47) As String

For i = 1 To 47
    l(i) = ""
Next

Dim Filtro As String

Filtro = ""
Filtro = Filtro & " AND A.VIAGEM='" & Viagem & "'"

If GMVI <> "" Then
    Filtro = Filtro & " AND A.GMVI = '" & GMVI & "'"
End If

If DT <> "" Then
    Filtro = Filtro & " AND A.DT = '" & DT & "'"
End If


'ZOLD Set TBH = New ADODB.Recordset
    
SQl = "SELECT a.DT, a.viagem,A.GMVI,A.OBS_MANUSEIO, I.DESCR AS PORTO_O, H.NUMERO, A.QUANTIDADE, A.PESO_BRUTO,J.DESCR AS EMBALAGEM,"
SQl = SQl & " A.MARCA, A.MERCADORIA, A.GENERO, C.NOMENAVIO AS NAVIO, F.DESCR AS AGENCIA, F.CODE AS COD_AGENCIA, C.DT_ATRACACAOP, E.CODE AS ARM_ATRACACAOP"
'Sql = Sql & " FROM " & Banco_Sgipa & "TB_CARGA_SOLTA A, " & Banco_Sgipa & "DTE_TB_VIAGENS C,"
'Sql = Sql & " " & Banco_Sgipa & "DTE_TB_ARMAZENS E, " & Banco_Sgipa & "DTE_TB_AGENCIAS F, " & Banco_Sgipa & "TB_BL H, " & Banco_Sgipa & "DTE_TB_PORTOS I," & Banco_Sgipa & "DTE_TB_EMBALAGENS J "
'Sql = Sql & " WHERE A.VIAGEM = C.VIAGEM  AND"
'Sql = Sql & " H.AGENCIA = F.CODE(+) AND A.BL = H.AUTONUM AND H.PORTO_ORIGEM = I.CODE(+) AND"
'Sql = Sql & "  C.ARM_ATRACACAOP = E.CODE(+) AND A.EMBALAGEM=J.CODE(+) and A.GMVI IS NOT NULL"
'Sql = Sql & Filtro & " ORDER BY A.GMVI"

SQl = SQl & " FROM " & Banco_Sgipa & "TB_CARGA_SOLTA A "
SQl = SQl & " INNER JOIN " & Banco_Sgipa & "DTE_TB_VIAGENS C ON A.VIAGEM = C.VIAGEM "
SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "DTE_TB_ARMAZENS E ON C.ARM_ATRACACAOP = E.CODE "
SQl = SQl & " INNER JOIN  " & Banco_Sgipa & "TB_BL H ON  A.BL = H.AUTONUM "
SQl = SQl & " LEFT JOIN  " & Banco_Sgipa & "DTE_TB_AGENCIAS F ON  H.AGENCIA = F.CODE "

SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "DTE_TB_PORTOS I ON H.PORTO_ORIGEM = I.CODE "
SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "DTE_TB_EMBALAGENS J ON  A.EMBALAGEM=J.CODE"
SQl = SQl & " WHERE "
SQl = SQl & " A.GMVI IS NOT NULL"
SQl = SQl & Filtro & " ORDER BY A.GMVI"

'ZOLD TBH.Open Sql, Db, 3, 3

PRSet Tbh, SQl

If Tbh.EOF Then
    MsgBox "Nehum registro", vbCritical
    Exit Sub
End If
Open "c:\microledtemp\gmvi.lst" For Output As #1
Do
    If GMVIant <> Tbh!GMVI Then
        Do
            X = InputBox("Informe o número de vias da GMVI:" & Tbh!GMVI, , "1")
            If X = "" Then Exit Sub
        Loop While Not IsNumeric(X)
        If Val(X) <> 0 Then
            For Controle = 1 To Val(X)
                If IsNull(Tbh!Dt_AtracacaoP) Then
                    Dt_AtracacaoP = "00/00/0000"
                Else
                    Dt_AtracacaoP = Format(Tbh!Dt_AtracacaoP, "DD/MM/YYYY")
                End If
                l(15) = s(2) & Left$(Tbh!DT & s(13), 13) & s(3) & Left$(Tbh!GMVI & s(13), 13) & s(3) & Format(Controle, "000") & s(5) & "010" & Format(Controle, "000") & "/2001"
                
                Dim Manifesto As String
                Manifesto = ""
                Set TBH2 = New ADODB.Recordset
                TBH2.Open "Select manifesto from " & Banco_Sgipa & "DTE_TB_MANIFESTOS where viagem=" & Tbh!Viagem, db, 3, 3
                If Not TBH2.EOF Then Manifsto = Nnull(TBH2!Manifesto, 1)
                
                l(18) = s(2) & Left$(Tbh!porto_o & s(25), 25) & s(22) & Manifesto
                l(20) = s(2) & Left$(Tbh!Numero & s(16), 16) & s(9) & Left$(Tbh!quantidade & s(6), 6) & s(18) & Format(Left$(Tbh!PESO_BRUTO & s(8), 8), "###,##0.000")
                l(22) = s(2) & Tbh!Marca
                l(24) = s(2) & Left$(Tbh!MERCADORIA & s(25), 25) & s(24) & Nnull(Tbh!embalagem, 1)
                l(27) = s(2) & Left$(Tbh!Navio & s(21), 21) & s(2) & Left$(Tbh!agencia & s(38), 38) & s(2) & Tbh!COD_AGENCIA
                l(29) = s(3) & Mid$(Dt_AtracacaoP, 1, 1) & s(1) & Mid$(Dt_AtracacaoP, 2, 1) & s(2) & Mid$(Dt_AtracacaoP, 4, 1) & s(1) & Mid$(Dt_AtracacaoP, 5, 1) & s(2) & Mid$(Dt_AtracacaoP, 7, 1) & s(1) & Mid$(Dt_AtracacaoP, 8, 1) & s(1) & Mid$(Dt_AtracacaoP, 9, 1) & s(1) & Mid$(Dt_AtracacaoP, 10, 1)
                If CodEmpSis = "BAN" Then
                    l(45) = s(3) & "IPA - BANDEIRANTES"
                Else
                    l(45) = s(3) & "IPA - CIA.BANDEIRANTES SABOO"
                End If
                If CodEmpSis = "BAN" Then
                    l(45) = s(3) & "IPA - BANDEIRANTES"
                Else
                    If Cod_Empresa = 1 Then l(45) = s(3) & "IPA - CIA.BANDEIRANTES SABOO"
                    If Cod_Empresa = 2 Then l(45) = s(3) & "IPA - CIA.BANDEIRANTES "
                End If
                l(47) = s(3) & Nnull(Tbh!obs_manuseio, 1)
                For i = 10 To 47
                    l(i) = UCase(l(i))
                    Printer.Print l(i)
                Next
                Printer.EndDoc
            Next Controle
        End If
    End If
    GMVIant = Tbh!GMVI
    Tbh.MoveNext
Loop Until Tbh.EOF
Close #1
Tbh.Close
a = Shell("C:\WINDOWS\NOTEPAD.EXE " & "c:\microledtemp\GMVI.LST", vbMaximizedFocus)
End Sub







Public Function BloqueioBLCNTR(Lote As Long, CNTR As Long) As String
Dim RsBloq  As ADODB.Recordset
Dim sRET    As String
sRET = ""
If Lote <> 0 Then
    PRSet RsBloq, "SELECT FLAG_BLOQUEIO, FLAG_BLOQUEIO_MANUAL FROM " & Banco_Sgipa & "TB_BL WHERE AUTONUM = " & Lote
    If Not RsBloq.EOF Then
        If RsBloq.Fields("FLAG_BLOQUEIO") = 1 Or RsBloq.Fields("FLAG_BLOQUEIO_MANUAL") = 1 Then
            sRET = "BL Bloqueado"
        End If
    End If
    RsBloq.Close
End If
If CNTR <> 0 Then
    PRSet RsBloq, "SELECT id_conteiner,FLAG_BLOQUEIO, FLAG_BLOQUEIO_MANUAL,REGIME FROM " & Banco_Sgipa & "TB_CNTR_BL WHERE autonum = " & CNTR
    If Not RsBloq.EOF Then
        If RsBloq.Fields("FLAG_BLOQUEIO") = 1 Or RsBloq.Fields("FLAG_BLOQUEIO_MANUAL") = 1 Then
            sRET = "Conteiner " & RsBloq!Id_Conteiner & " Bloqueado"
        End If
    End If
    If RsBloq!Regime = "HH" Then
        RsBloq.Close
        'verifica cntr parte lote
        If sRET = "" Then
            SQl = "select "
            SQl = SQl & "(SELECT id_conteiner FROM " & Banco_Sgipa & "TB_CNTR_BL WHERE autonum = " & CNTR & ") id_conteiner"
            SQl = SQl & ",a.autonum,a.numero from " & Banco_Sgipa & "tb_bl a where a.flag_ativo=1 and a.autonum in "
            SQl = SQl & "("
            SQl = SQl & " select b.bl from " & Banco_Sgipa & "tb_amr_cntr_bl b where b.cntr=" & CNTR & " "
            SQl = SQl & ")  "
            SQl = SQl & " and (nvl(a.flag_bloqueio,0)=1 or nvl(a.flag_bloqueio_manual,0)=1)"
            PRSet RsBloq, SQl
            If Not RsBloq.EOF Then
                sRET = "Conteiner (" & RsBloq!Id_Conteiner & ") Part Lote Bloqueado no BL " & RsBloq!Numero & " (" & RsBloq!autonum & ")"
            End If
            RsBloq.Close
        End If
    Else
        RsBloq.Close
    End If
End If


BloqueioBLCNTR = sRET
sRET = Empty
Set RsBloq = Nothing
End Function



Public Function RepInsertLine(nMODO As Integer) 'As Crystal.CrystalReport
Dim SQl   As String
'Dim OCR   As Crystal.CrystalReport
SQl = Empty

If Banco_Em_Uso = "ORACLE" Then
SQl = "SELECT "
SQl = SQl & "RPAD(A.AUTONUM||' - '||A.DESCR,50) AS LISTA, "
SQl = SQl & "DECODE(E.RAZAO,NULL,'',E.RAZAO||'('||E.CGC||')') AS IMPORTADOR, "
SQl = SQl & "DECODE(G.RAZAO,NULL,'',G.RAZAO||'('||G.CGC||')') AS DESPACHANTE, "
SQl = SQl & "DECODE(F.RAZAO,NULL,'',F.RAZAO||'('||F.CGC||')') AS INDICADOR, "
SQl = SQl & "A.OBS AS OBSERVACOES, "
SQl = SQl & "DECODE(A.FLAG_LIBERADA,1,'LIBERADA','NÃO LIBERADA') AS LIBERADA, "
SQl = SQl & "DECODE(A.FLAG_GENERICA,1,'GENÉRICA','ESPECÍFICA') AS GENERICA, "
Else
SQl = "SELECT ((  CAST(A.AUTONUM AS VARCHAR) + ' - ' + CAST(A.DESCR AS VARCHAR)  +  "
SQl = SQl & " REPLICATE(' ', ( 50  -  LEN(CAST(A.AUTONUM AS VARCHAR) + ' - ' + CAST(A.DESCR AS VARCHAR))) ))) , "
SQl = SQl & "CASE WHEN E.RAZAO IS NULL THEN NULL ELSE CAST(E.RAZAO AS VARCHAR) + '(' + CAST(E.CGC AS VARCHAR) + ')'  END As Importador, "
SQl = SQl & "CASE  WHEN G.RAZAO IS NULL THEN NULL ELSE CAST(G.RAZAO AS VARCHAR) + '(' + CAST(G.CGC AS VARCHAR) + ')' END As Despachante, "
SQl = SQl & "CASE WHEN F.RAZAO IS NULL THEN NULL ELSE CAST(F.RAZAO AS VARCHAR) + '(' + CAST(F.CGC AS VARCHAR) + ')'  END As Indicador, "
SQl = SQl & "A.OBS AS OBSERVACOES, "
SQl = SQl & "CASE A.FLAG_LIBERADA  WHEN 1 THEN 'LIBERADA' Else 'NÃO LIBERADA'  END As LIBERADA,"
SQl = SQl & "CASE A.FLAG_GENERICA WHEN 1 THEN 'GENÉRICA'  Else 'ESPECÍFICA'   END As GENERICA, "

End If


SQl = SQl & "A.DATA_INICIO AS INICIO, "
SQl = SQl & "D.USUARIO AS USUARIO_SIS, "
SQl = SQl & "A.USUARIO_REDE AS USUARIO_REDE, "
SQl = SQl & "A.MAQUINA_REDE AS MAQUINA, "
SQl = SQl & "A.DIAS_APOS_GR, "
SQl = SQl & "H.DESCR AS FORMA_PAGAMENTO "


'Sql = Sql & "FROM "
'Sql = Sql & "" & Banco_Sgipa & "TB_LISTAS_PRECOS A, "
'Sql = Sql & "" & Banco_Sgipa & "TB_LISTA_PRECO_SERVICOS_FIXOS B, "
'Sql = Sql & "" & Banco_Sgipa & "TB_LISTA_P_S_PERIODO C, "
'Sql = Sql & "" & Banco_Sgipa & "TB_CAD_USUARIOS D, "
'Sql = Sql & "" & Banco_Sgipa & "TB_CAD_PARCEIROS E, "
'Sql = Sql & "" & Banco_Sgipa & "TB_CAD_PARCEIROS F, "
'Sql = Sql & "" & Banco_Sgipa & "TB_CAD_PARCEIROS G, "
'Sql = Sql & "" & Banco_Sgipa & "TB_FORMAS_PAGAMENTO H "
'Sql = Sql & "WHERE "
'Sql = Sql & "A.AUTONUM = B.LISTA(+) "
'Sql = Sql & "AND A.AUTONUM = C.LISTA(+) "

'If nMODO = 3 Then
'    Sql = Sql & "AND (B.AUTONUM_VINCULADO <> 0 OR C.AUTONUM_VINCULADO <> 0) "
'ElseIf nMODO = 2 Then
'    Sql = Sql & "AND (B.AUTONUM_VINCULADO = 0 OR C.AUTONUM_VINCULADO IS NULL) "
'End If
'Sql = Sql & "AND A.IMPORTADOR = E.AUTONUM(+) "
'Sql = Sql & "AND A.CAPTADOR = F.AUTONUM(+) "
'Sql = Sql & "AND A.DESPACHANTE = G.AUTONUM(+) "
'Sql = Sql & "AND A.FORMA_PAGAMENTO = H.AUTONUM(+) "
'Sql = Sql & "AND A.USUARIO_SIS = D.AUTONUM(+) "

SQl = SQl & " FROM "
 SQl = SQl & " " & Banco_Sgipa & "TB_LISTAS_PRECOS A "
    SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "TB_LISTA_PRECO_SERVICOS_FIXOS B ON A.AUTONUM = B.LISTA "
    SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "TB_LISTA_P_S_PERIODO C ON A.AUTONUM = C.LISTA "
    SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "TB_CAD_USUARIOS D ON A.USUARIO_SIS = D.AUTONUM "
    SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "TB_CAD_PARCEIROS E ON A.IMPORTADOR = E.AUTONUM "
    SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "TB_CAD_PARCEIROS F ON A.CAPTADOR = F.AUTONUM "
    SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "TB_CAD_PARCEIROS G ON A.DESPACHANTE = G.AUTONUM "
    SQl = SQl & " LEFT JOIN " & Banco_Sgipa & "TB_FORMAS_PAGAMENTO H ON  A.FORMA_PAGAMENTO = H.AUTONUM "
    SQl = SQl & " WHERE "
    If nMODO = 3 Then
        SQl = SQl & "(B.AUTONUM_VINCULADO <> 0 OR C.AUTONUM_VINCULADO <> 0) "
    ElseIf nMODO = 2 Then
        SQl = SQl & "(B.AUTONUM_VINCULADO = 0 OR C.AUTONUM_VINCULADO IS NULL) "
    End If

SQl = SQl & " GROUP BY "
SQl = SQl & "A.AUTONUM, "
SQl = SQl & "A.DESCR, "
SQl = SQl & "A.OBS, "
SQl = SQl & "A.FLAG_LIBERADA, "
SQl = SQl & "A.FLAG_GENERICA, "
SQl = SQl & "A.DATA_INICIO, "
SQl = SQl & "D.USUARIO, "
SQl = SQl & "A.USUARIO_REDE, "
SQl = SQl & "A.MAQUINA_REDE, "
SQl = SQl & "A.DIAS_APOS_GR, "
SQl = SQl & "H.DESCR, "
SQl = SQl & "E.RAZAO, "
SQl = SQl & "E.CGC, "
SQl = SQl & "G.RAZAO, "
SQl = SQl & "G.CGC, "
SQl = SQl & "F.RAZAO, "
SQl = SQl & "F.CGC "



'Set OCR = PreparaReport(App.Path & IIf(Right(App.Path, 1) = "\", Empty, "\") & "TCOB0000.RPT", sql, , True)
'Set RepInsertLine = OCR
''set ocr = nothing
SQl = Empty
End Function
Public Function FichaFMA(Lote As Long) As Integer
'Retornos: 0-Nada, 1-Conteiner, 2-Carga Solta, 3-Ambos
Screen.MousePointer = vbHourglass
Dim SQl    As String
Dim oTB    As ADODB.Recordset
Dim oSG    As ADODB.Recordset
Dim sCNTRs As String
Dim lCS    As Boolean
Dim lCN    As Boolean
Dim nRET   As Integer
Dim nUser  As Long
PRSet tbSenha, ""
nUser = tbSenha.Fields("AUTONUM")
PRSet oTB, "SELECT COUNT(*) as cnt FROM " & Banco_Sgipa & "TB_BL A, " & Banco_Sgipa & "TB_AMR_CNTR_BL B, " & Banco_Sgipa & "TB_CNTR_BL C WHERE A.AUTONUM = B.BL AND B.CNTR = C.AUTONUM AND A.AUTONUM = " & Lote & " ", 0
lCN = (oTB!cnt > 0)
oTB.Close

PRSet oTB, "SELECT COUNT(*) as cnt FROM " & Banco_Sgipa & "TB_BL A, " & Banco_Sgipa & "TB_CARGA_SOLTA B WHERE A.AUTONUM = B.BL AND A.AUTONUM = " & Lote & " ", 0
lCS = (oTB!cnt > 0)
oTB.Close

If lCS = False And lCN = False Then nRET = 0
If lCS = False And lCN = True Then nRET = 1
If lCS = True And lCN = False Then nRET = 2
If lCS = True And lCN = True Then nRET = 3

If nRET = 1 Or nRET = 3 Then
    Executa "DELETE FROM " & Banco_Sgipa & "TB_TEMP_FMA WHERE USUARIO = " & tbSenha.Fields("AUTONUM") & " AND MAQUINA = '" & Maquina & "' "
    SQl = Empty
    SQl = SQl & " SELECT "
    If Banco_Em_Uso = "ORACLE" Then
        SQl = SQl & " TO_CHAR(MAX(A.SEQ_FMA),'0000')||'/'||MAX(A.ANO_FMA) AS FMA, "
    Else
        SQl = SQl & " RIGHT('0000' + CONVERT(VARCHAR,MAX(A.SEQ_FMA)),4)||'/'||CAST(MAX(A.ANO_FMA) AS VARCHAR) AS FMA, "
    End If
    
    
    SQl = SQl & " MAX(C.PATIO) AS PATIO, "
    SQl = SQl & " MAX(D.NOMENAVIO) AS NAVIO, "
    SQl = SQl & " MAX(G.NACIONALIDADE) AS NACION, "
    SQl = SQl & " MAX(D.DT_ATRACACAO) AS ATRACACAO, "
    SQl = SQl & " MAX(E.DESCR) AS ARMADOR, "
    SQl = SQl & " MAX(F.DESCR) AS PORTO, "
    SQl = SQl & " MAX(H.MARCA) AS MARCA, "
    SQl = SQl & " MAX(H.MERCADORIA) AS ESPECIE, "
    SQl = SQl & " MAX(C.ID_CONTEINER) AS SIGLA, "
    SQl = SQl & " MAX(I.DESCR) AS EMBALAGEM, "
    SQl = SQl & " COUNT(C.ID_CONTEINER) AS QUANTIDADE, "
    SQl = SQl & " MAX(H.PESO_BRUTO) AS PESOM, "
    SQl = SQl & " MAX(J.RAZAO) AS CONSIGNATARIO, "
    SQl = SQl & " MAX(C.DT_ENTRADA) AS ENTRADA, "
    SQl = SQl & " MAX(A.OBS_FMA) AS OBS, "
    SQl = SQl & " MAX(A.NUMERO) AS NUMERO, "
    SQl = SQl & " A.AUTONUM AS LOTE "
    
    SQl = SQl & " , sum(h.quantidade_real) as quantidade_real"
    SQl = SQl & " , sum(H.PESO_BRUTO) AS PESO_bruto "
    
    
    SQl = SQl & " FROM "
    SQl = SQl & " TB_BL A, "
    SQl = SQl & " TB_AMR_CNTR_BL B, "
    SQl = SQl & " TB_CNTR_BL C, "
    SQl = SQl & " DTE_TB_VIAGENS D, "
    SQl = SQl & " DTE_TB_ARMAZENS E, "
    SQl = SQl & " DTE_TB_PORTOS F, "
    SQl = SQl & " DTE_TB_NAVIOS G, "
    SQl = SQl & " TB_CARGA_CNTR H, "
    SQl = SQl & " DTE_TB_EMBALAGENS I, "
    SQl = SQl & " TB_CAD_PARCEIROS J "
    SQl = SQl & " WHERE "
    SQl = SQl & " A.AUTONUM = B.BL "
    SQl = SQl & " AND B.CNTR = C.AUTONUM "
    SQl = SQl & " AND C.VIAGEM = D.VIAGEM "
    SQl = SQl & " AND D.ARM_ATRACACAO = E.CODE "
    SQl = SQl & " AND A.PORTO_ORIGEM = F.CODE "
    SQl = SQl & " AND D.NAVIO = G.CODE "
    SQl = SQl & " AND A.AUTONUM = H.BL and c.id_conteiner = h.id_conteiner "
    SQl = SQl & " AND H.EMBALAGEM = I.CODE "
    SQl = SQl & " AND A.IMPORTADOR = J.AUTONUM "
    SQl = SQl & " AND A.AUTONUM = " & Lote & " "
    SQl = SQl & " GROUP BY "
    SQl = SQl & " A.AUTONUM "
    
    PRSet oTB, SQl, 0
    
    SQl = Empty
    
    
    If Not oTB.EOF = True Then
        SQl = Empty
        SQl = " SELECT distinct"
        SQl = SQl & " C.ID_CONTEINER AS SIGLA "
        SQl = SQl & " FROM "
        SQl = SQl & " TB_BL A, "
        SQl = SQl & " TB_AMR_CNTR_BL B, "
        SQl = SQl & " TB_CNTR_BL C, "
        SQl = SQl & " DTE_TB_VIAGENS D, "
        SQl = SQl & " DTE_TB_ARMAZENS E, "
        SQl = SQl & " DTE_TB_PORTOS F, "
        SQl = SQl & " DTE_TB_NAVIOS G, "
        SQl = SQl & " TB_CARGA_CNTR H, "
        SQl = SQl & " DTE_TB_EMBALAGENS I, "
        SQl = SQl & " TB_CAD_PARCEIROS J "
        SQl = SQl & " WHERE "
        SQl = SQl & " A.AUTONUM = B.BL "
        SQl = SQl & " AND B.CNTR = C.AUTONUM "
        SQl = SQl & " AND C.VIAGEM = D.VIAGEM "
        SQl = SQl & " AND D.ARM_ATRACACAO = E.CODE "
        SQl = SQl & " AND A.PORTO_ORIGEM = F.CODE "
        SQl = SQl & " AND D.NAVIO = G.CODE "
        SQl = SQl & " AND A.AUTONUM = H.BL "
        SQl = SQl & " AND H.EMBALAGEM = I.CODE "
        SQl = SQl & " AND A.IMPORTADOR = J.AUTONUM "
        SQl = SQl & " AND A.AUTONUM = " & Lote & " "
        PRSet oSG, SQl, 0
        
        sCNTRs = Empty
        
        While Not oSG.EOF = True
            sCNTRs = sCNTRs & oSG.Fields("SIGLA") & " || "
            SQl = Empty
            SQl = SQl & "INSERT INTO "
            SQl = SQl & "" & Banco_Sgipa & "TB_TEMP_FMA ( "
            SQl = SQl & "FMA, "
            SQl = SQl & "PATIO, "
            SQl = SQl & "NAVIO, "
            SQl = SQl & "NACION, "
            SQl = SQl & "ATRACAC, "
            SQl = SQl & "ARMADOR, "
            SQl = SQl & "PORTO, "
            SQl = SQl & "MARCA, "
            SQl = SQl & "ESPECIE, "
            SQl = SQl & "EMBALAG, "
            SQl = SQl & "QUANTDD, "
            SQl = SQl & "PESO_MN, "
            SQl = SQl & "CONSIGN, "
            SQl = SQl & "ENTRADA, "
            SQl = SQl & "OBS, "
            SQl = SQl & "LOTE, "
            SQl = SQl & "USUARIO, "
            SQl = SQl & "Maquina,"
            SQl = SQl & "CONTEINER) VALUES ( "
            SQl = SQl & "'" & oTB.Fields("FMA") & "', "
            SQl = SQl & "" & oTB.Fields("PATIO") & ", "
            SQl = SQl & "'" & oTB.Fields("NAVIO") & "', "
            SQl = SQl & "'" & oTB.Fields("NACION") & "', "
            SQl = SQl & "TO_DATE('" & Format$(oTB.Fields("ATRACACAO"), "DD/MM/YYYY") & "','DD/MM/YYYY'), "
            SQl = SQl & "'" & oTB.Fields("ARMADOR") & "', "
            SQl = SQl & "'" & oTB.Fields("PORTO") & "', "
            SQl = SQl & "'" & oTB.Fields("MARCA") & "', "
            SQl = SQl & "'" & oTB.Fields("ESPECIE") & "', "
            SQl = SQl & "'" & oTB.Fields("EMBALAGEM") & "', "
            
    '        Sql = Sql & "" & oTB.Fields("QUANTIDADE") & ", "
    '        Sql = Sql & "" & PPonto(oTB.Fields("PESOM")) & ", "
            SQl = SQl & "" & oTB.Fields("QUANTIDADE_real") & ", "
            SQl = SQl & "" & PPonto(oTB.Fields("PESO_bruto")) & ", "
            
            
            SQl = SQl & "'" & oTB.Fields("CONSIGNATARIO") & "', "
            SQl = SQl & "TO_DATE('" & Format$(oTB.Fields("ENTRADA"), "DD/MM/YYYY") & "','DD/MM/YYYY'), "
          '  Sql = Sql & "'" & Trim(sCNTRs & oTB.Fields("OBS")) & "', "
            SQl = SQl & "'" & oTB.Fields("OBS") & "', "
            SQl = SQl & "" & oTB.Fields("LOTE") & ", "
            SQl = SQl & "" & tbSenha.Fields("AUTONUM") & ", "
            SQl = SQl & "'" & Maquina & "', "
             SQl = SQl & "'" & oSG.Fields("SIGLA") & "'"
            SQl = SQl & ")"
            Executa SQl
            oSG.MoveNext
        Wend
    End If
    oTB.Close
End If



If nRET = 2 Or nRET = 3 Then
    Executa "DELETE FROM " & Banco_Sgipa & "TB_TEMP_FMA_CS WHERE USUARIO = " & tbSenha.Fields("AUTONUM") & " AND MAQUINA = '" & Maquina & "' "
    SQl = Empty
    SQl = "  SELECT "
    If Banco_Em_Uso = "ORACLE" Then
        SQl = SQl & "  TO_CHAR(MAX(A.SEQ_FMA),'0000')||'/'||MAX(A.ANO_FMA) AS FMA, "
    Else
        SQl = SQl & " RIGHT('0000' + CONVERT(VARCHAR,MAX(A.SEQ_FMA)),4)||'/'||cast(MAX(A.ANO_FMA) as varchar) AS FMA, "
    End If
        
    SQl = SQl & "  MAX(C.PATIO) AS PATIO, "
    SQl = SQl & "  MAX(D.NOMENAVIO) AS NAVIO, "
    SQl = SQl & "  MAX(G.NACIONALIDADE) AS NACION, "
    SQl = SQl & "  MAX(D.DT_ATRACACAO) AS ATRACACAO, "
    SQl = SQl & "  MAX(E.DESCR) AS ARMADOR, "
    SQl = SQl & "  MAX(F.DESCR) AS PORTO, "
    SQl = SQl & "  MAX(C.MARCA) AS MARCA, "
    SQl = SQl & "  MAX(C.MERCADORIA) AS ESPECIE, "
    SQl = SQl & "  MAX('CONTEINER') AS EMBALAGEM, "
    SQl = SQl & "  MAX(C.ITEM) AS SIGLA, "
    SQl = SQl & "  MAX(I.DESCR) AS EMBALAGEM, "
    SQl = SQl & "  SUM(C.QUANTIDADE_real) AS QUANTIDADE_real, "
    SQl = SQl & "  sum(C.PESO_BRUTO) AS PESO_bruto, "
    SQl = SQl & "  MAX(J.RAZAO) AS CONSIGNATARIO, "
    SQl = SQl & "  MAX(B.DT_ENTRADA) AS ENTRADA, "
    SQl = SQl & "  MAX(A.OBS_FMA) AS OBS, "
    SQl = SQl & "  A.AUTONUM AS LOTE "
     SQl = SQl & "  FROM "
        SQl = SQl & "  TB_BL A"
        SQl = SQl & "  LEFT JOIN (SELECT A.AUTONUM AS CS, MAX(C.DATA_GATE_IN) AS DT_ENTRADA FROM TB_CARGA_SOLTA A "
        SQl = SQl & "            INNER JOIN TB_CARGA_GATE_CS B ON A.AUTONUM = B.AUTONUMCS "
        SQl = SQl & "            INNER JOIN " & Banco_Operador & "TB_GATE C ON B.AUTONUMGATE = C.AUTONUM "
        SQl = SQl & "            WHERE  CODIGO_PESAGEM = 1 AND NVL(A.CNTR,0)=0 "
        SQl = SQl & "              GROUP BY A.AUTONUM) B "
        SQl = SQl & "  ON A.AUTONUM = B.CS "
        
        SQl = SQl & "  INNER JOIN TB_CARGA_SOLTA C ON A.AUTONUM = C.BL "
        SQl = SQl & "  LEFT JOIN DTE_TB_VIAGENS D ON  C.VIAGEM = D.VIAGEM "
        SQl = SQl & "  LEFT JOIN DTE_TB_ARMAZENS E ON D.ARM_ATRACACAO = E.CODE "
        SQl = SQl & "  LEFT JOIN DTE_TB_PORTOS F ON  A.PORTO_ORIGEM = F.CODE "
        SQl = SQl & "  LEFT JOIN DTE_TB_NAVIOS G ON D.NAVIO = G.CODE "
        SQl = SQl & "  LEFT JOIN DTE_TB_EMBALAGENS I ON C.EMBALAGEM = I.CODE "
        SQl = SQl & "  LEFT JOIN TB_CAD_PARCEIROS J ON A.IMPORTADOR = J.AUTONUM"
        SQl = SQl & "  WHERE"
        SQl = SQl & "  A.AUTONUM = " & Lote & " "
        SQl = SQl & "  GROUP BY "
        SQl = SQl & "  A.AUTONUM"
    'FIM NOVO CODIGO

    PRSet oTB, SQl, 0
    SQl = Empty
    If Not oTB.EOF = True Then
        sCNTRs = Empty
        SQl = Empty
        SQl = SQl & "  SELECT "
        SQl = SQl & "  MAX(C.MERCADORIA) AS MERCADORIA "
         SQl = SQl & "  FROM "
            SQl = SQl & "  TB_BL A"
            SQl = SQl & "  LEFT JOIN (SELECT A.AUTONUM AS CS, MAX(C.DATA_GATE_IN) AS DT_ENTRADA "
            SQl = SQl & "            FROM TB_CARGA_SOLTA A INNER JOIN TB_CARGA_GATE_CS B ON A.AUTONUM = B.AUTONUMCS "
            SQl = SQl & "            INNER JOIN " & Banco_Operador & "TB_GATE C ON B.AUTONUMGATE = C.AUTONUM  "
            SQl = SQl & "            WHERE  CODIGO_PESAGEM = 1 AND NVL(A.CNTR,0) = 0 "
            SQl = SQl & "            GROUP BY A.AUTONUM) B ON A.AUTONUM = B.CS "
            SQl = SQl & "  LEFT JOIN TB_CARGA_SOLTA C ON A.AUTONUM = C.BL "
            SQl = SQl & "  LEFT JOIN DTE_TB_VIAGENS D  ON C.VIAGEM = D.VIAGEM "
            SQl = SQl & "  LEFT JOIN DTE_TB_ARMAZENS E ON D.ARM_ATRACACAO = E.CODE "
            SQl = SQl & "  LEFT JOIN DTE_TB_PORTOS F ON A.PORTO_ORIGEM = F.CODE "
            SQl = SQl & "  LEFT JOIN DTE_TB_NAVIOS G ON D.NAVIO = G.CODE "
            SQl = SQl & "  LEFT JOIN DTE_TB_EMBALAGENS I ON C.EMBALAGEM = I.CODE "
            SQl = SQl & "  LEFT JOIN TB_CAD_PARCEIROS J ON  A.IMPORTADOR = J.AUTONUM"
            SQl = SQl & "  WHERE "
            SQl = SQl & "  A.AUTONUM = " & Lote & " "
            SQl = SQl & "  GROUP BY "
            SQl = SQl & "  A.AUTONUM"
        
        PRSet oSG, SQl, 0
        While Not oSG.EOF = True
            sCNTRs = sCNTRs & oSG.Fields("MERCADORIA") & " || "
            oSG.MoveNext
        Wend
        oSG.Close
        SQl = Empty
        SQl = SQl & "INSERT INTO "
        SQl = SQl & "" & Banco_Sgipa & "TB_TEMP_FMA_CS ( "
        SQl = SQl & "FMA, "
        SQl = SQl & "PATIO, "
        SQl = SQl & "NAVIO, "
        SQl = SQl & "NACION, "
        SQl = SQl & "ATRACAC, "
        SQl = SQl & "ARMADOR, "
        SQl = SQl & "PORTO, "
        SQl = SQl & "MARCA, "
        SQl = SQl & "ESPECIE, "
        SQl = SQl & "EMBALAG, "
        SQl = SQl & "QUANTDD, "
        SQl = SQl & "PESO_MN, "
        SQl = SQl & "CONSIGN, "
        SQl = SQl & "ENTRADA, "
        SQl = SQl & "OBS, "
        SQl = SQl & "LOTE, "
        SQl = SQl & "USUARIO, "
        SQl = SQl & "Maquina ) VALUES ( "
        SQl = SQl & "'" & oTB.Fields("FMA") & "', "
        SQl = SQl & "" & Nnull(oTB.Fields("PATIO"), 0) & ", "
        SQl = SQl & "'" & oTB.Fields("NAVIO") & "', "
        SQl = SQl & "'" & Nnull(oTB.Fields("NACION"), 1) & "', "
        If Not IsNull(oTB.Fields("ATRACACAO")) Then
            SQl = SQl & "TO_DATE('" & Format$(oTB.Fields("ATRACACAO"), "DD/MM/YYYY") & "','DD/MM/YYYY'), "
        Else
            SQl = SQl & "NULL,"
        End If
        SQl = SQl & "'" & oTB.Fields("ARMADOR") & "', "
        SQl = SQl & "'" & oTB.Fields("PORTO") & "', "
        SQl = SQl & "'" & oTB.Fields("MARCA") & "', "
        SQl = SQl & "'" & oTB.Fields("ESPECIE") & "', "
        SQl = SQl & "'" & oTB.Fields("EMBALAGEM") & "', "
        
        SQl = SQl & "" & oTB.Fields("QUANTIDADE_REAL") & ", "
'        Sql = Sql & "" & PPonto(Nnull(oTB.Fields("PESOM"), 0)) & ","
        SQl = SQl & "" & PPonto(Nnull(oTB.Fields("PESO_bruto"), 0)) & ","
        
        SQl = SQl & "'" & oTB.Fields("CONSIGNATARIO") & "', "
        If Not IsNull(oTB.Fields("ENTRADA")) Then
            SQl = SQl & "TO_DATE('" & Format$(oTB.Fields("ENTRADA"), "DD/MM/YYYY") & "','DD/MM/YYYY'), "
        Else
            SQl = SQl & "NULL,"
        End If
        SQl = SQl & "'" & Trim(sCNTRs & oTB.Fields("OBS")) & "', "
        SQl = SQl & "" & oTB.Fields("LOTE") & ", "
        SQl = SQl & "" & tbSenha.Fields("AUTONUM") & ", "
        SQl = SQl & "'" & Maquina & "' "
        SQl = SQl & ")"
        Executa SQl
        
        SQl = Empty
        SQl = SQl & "INSERT INTO "
        SQl = SQl & "" & Banco_Sgipa & "TB_TEMP_FMA ( "
        SQl = SQl & "FMA, "
        SQl = SQl & "PATIO, "
        SQl = SQl & "NAVIO, "
        SQl = SQl & "NACION, "
        SQl = SQl & "ATRACAC, "
        SQl = SQl & "ARMADOR, "
        SQl = SQl & "PORTO, "
        SQl = SQl & "MARCA, "
        SQl = SQl & "ESPECIE, "
        SQl = SQl & "EMBALAG, "
        SQl = SQl & "QUANTDD, "
        SQl = SQl & "PESO_MN, "
        SQl = SQl & "CONSIGN, "
        SQl = SQl & "ENTRADA, "
        SQl = SQl & "OBS, "
        SQl = SQl & "LOTE, "
        SQl = SQl & "USUARIO, "
        SQl = SQl & "Maquina ) VALUES ( "
        SQl = SQl & "'" & oTB.Fields("FMA") & "', "
        SQl = SQl & "" & Nnull(oTB.Fields("PATIO"), 0) & ", "
        SQl = SQl & "'" & oTB.Fields("NAVIO") & "', "
        SQl = SQl & "'" & oTB.Fields("NACION") & "', "
        If Not IsNull(oTB.Fields("ATRACACAO")) Then
            SQl = SQl & "TO_DATE('" & Format$(oTB.Fields("ATRACACAO"), "DD/MM/YYYY") & "','DD/MM/YYYY'), "
        Else
            SQl = SQl & "NULL,"
        End If
        SQl = SQl & "'" & oTB.Fields("ARMADOR") & "', "
        SQl = SQl & "'" & oTB.Fields("PORTO") & "', "
        SQl = SQl & "'" & oTB.Fields("MARCA") & "', "
        SQl = SQl & "'" & oTB.Fields("ESPECIE") & "', "
        SQl = SQl & "'" & oTB.Fields("EMBALAGEM") & "', "
'        Sql = Sql & "" & PPonto(Nnull(oTB.Fields("QUANTIDADE"), 0)) & ", "
        SQl = SQl & "" & PPonto(Nnull(oTB.Fields("QUANTIDADE_real"), 0)) & ", "
'        Sql = Sql & "" & PPonto(Nnull(oTB.Fields("PESOM"), 0)) & ", "
        SQl = SQl & "" & PPonto(Nnull(oTB.Fields("PESO_bruto"), 0)) & ", "
        SQl = SQl & "'" & oTB.Fields("CONSIGNATARIO") & "', "
        
        If Not IsNull(oTB.Fields("ENTRADA")) Then
            SQl = SQl & "TO_DATE('" & Format$(oTB.Fields("ENTRADA"), "DD/MM/YYYY") & "','DD/MM/YYYY'), "
        Else
            SQl = SQl & "NULL,"
        End If
        
        SQl = SQl & "'" & Trim(sCNTRs & oTB.Fields("OBS")) & "', "
        SQl = SQl & "" & oTB.Fields("LOTE") & ", "
        SQl = SQl & "" & tbSenha.Fields("AUTONUM") & ", "
        SQl = SQl & "'" & Maquina & "' "
        SQl = SQl & ")"
        Executa SQl
    End If
    oTB.Close
End If
Set oTB = Nothing
Set oSG = Nothing
SQl = Empty
sCNTRs = Empty
FichaFMA = nRET
nRET = Empty
Screen.MousePointer = vbDefault
End Function


Public Sub EnviaMensagem(Optional TEXTO As String)
'Principal.StatusBar1.Panels(2).Text = Empty
'Principal.StatusBar1.Panels(2).Text = Trim(StrConv(TEXTO, vbProperCase))
'If Not TEXTO = Empty Then
'    Screen.MousePointer = vbHourglass
'Else
'    Screen.MousePointer = vbDefault
'End If
'Principal.StatusBar1.Refresh
End Sub

Public Sub MostraFormulario(ByVal Formulário As Form, Optional Modal As FormShowConstants = vbModal)
Dim sMENUS() As String
Dim sMENU2() As String
Dim t As Integer
Dim N As Integer

    
    'pega hora do servidor
    Dim oRS As ADODB.Recordset
'    PRSet oRS, "SELECT TO_CHAR(SYSDATE,'DD/MM/YYYY HH24:MI:SS') AS HOJE FROM DUAL", 0
'
'    If (Abs(DateDiff("n", oRS.Fields("HOJE"), Now)) > 5) Then '5 MINUTOS
'        MsgBox "Data/Hora Local esta divergente em mais de 5 minutos com o servidor" & vbCr & "Servidor = " & oRS!hoje & vbCr & "Local = " & Now & vbCr & "Por Favor, Corrija o Problema.", vbOKOnly + vbInformation, "Operação Inválida."
'        oRS.Close
'        Exit Sub
'    End If
'    oRS.Close
    
    
    
'    If Not ValidaHoje = True And Usuario_Sistema <> "MICROLED" Then
'        MsgBox "Data Local está Diferente Daquela No Servidor. Por Favor, Corrija o Problema.", vbOKOnly + vbInformation, "Operação Inválida."
'        Exit Sub
'    Else
        If Usuario_Sistema = "MICROLED" Then
            '=== DBA
            'ELIMINAR REGISTROS DE OBJETOS / NÍVEIS JÁ EXISTENTES
    '        Executa "DELETE FROM " & Banco_Sgipa & "TB_LV_ACESSOS"
    '        Executa "DELETE FROM " & Banco_Sgipa & "TB_OBJETOS"
            'CRIAR TABELAS, INDICES, SEQUENCIAIS E GATILHOS:
    '        CriaObjetos
            'PARA ZERAR AS SEQUENCES, QUANDO FOR O CASO:
    '        Executa "DROP SEQUENCE " & Banco_Sgipa & "SEQ_TB_LV_ACESSOS"
    '        Executa "CREATE SEQUENCE " & Banco_Sgipa & "SEQ_TB_LV_ACESSOS START WITH 1 NOMAXVALUE MINVALUE 1 NOCYCLE CACHE 2 NOORDER"
    '        Executa "DROP SEQUENCE " & Banco_Sgipa & "SEQ_TB_OBJETOS"
    '        Executa "CREATE SEQUENCE " & Banco_Sgipa & "SEQ_TB_OBJETOS START WITH 1 NOMAXVALUE MINVALUE 1 NOCYCLE CACHE 2 NOORDER"
    '        Executa "DROP SEQUENCE " & Banco_Sgipa & "SEQ_TB_USUARIOS"
    '        Executa "CREATE SEQUENCE " & Banco_Sgipa & "SEQ_TB_USUARIOS START WITH 1 NOMAXVALUE MINVALUE 1 NOCYCLE CACHE 2 NOORDER"
            '=== END DBA
            'COLETA DE DADOS PARA OBJETOS DE ACESSO CONTROLADO
            If TypeOf Formulário Is MDIForm Then
'                n = Coletar_Menus(Formulário, sMENUS)
'                For t = 0 To n
'                    If Not n = Empty Then
'                        If Not sMENUS(t) = Empty Then
'                            sMENU2 = Split(sMENUS(t), "|")
'                            GravaObjetoDeAcessoRestrito UCase(sMENU2(0)), UCase(sMENU2(1)), 0
'                        End If
'                    End If
'                Next
                Formulário.Show
            Else
                'GravaObjetoDeAcessoRestrito UCase(Formulário.Name), UCase(Formulário.Caption), 1
                If Not Formulário.MDIChild = True Then
                    If Not Modal = vbModal Then
                        Formulário.Show
                    Else
                        Formulário.Show
                    End If
                Else
                    Formulário.Show
                End If
            End If
        Else
            If Not TypeOf Formulário Is MDIForm Then
                If Not Formulário.MDIChild = True Then
                    If Not Modal = vbModal Then
                        Formulário.Show
                    Else
                        If Not Formulário.Visible = True Then
                            Formulário.Show
                        End If
                    End If
                Else
                    Formulário.Show
                End If
            Else
                Formulário.Show
            End If
        End If
'    End If
'End If

t = Empty
N = Empty
Erase sMENUS
Erase sMENU2
End Sub



Public Sub GaranteObjetos()
Dim oTB As ADODB.Recordset
Dim SQl As String
Dim ob() As String
Dim nm() As String
Dim N   As Integer
Dim nn  As Variant
If Usuario_Sistema = "MICROLED" Then
SQl = Empty
ReDim ob(196)
ReDim nm(196)
End If
Erase ob
Erase nm
Set oTB = Nothing
SQl = Empty
N = Empty
End Sub
'A RAZÃO DESTA FUNÇÃO É APENAS VALIDAR SE É POSSÍVEL OU NÃO EDITAR OS COMPONENTES DO
'LOTE, NA TELA DE CAPTAÇÃO. RETORNA UM BOOLEAN, INDICANDO SE PODE OU NÃO LIBERAR A EDIÇÃO.
Public Function LoteSolicitado(Lote As Long, Optional Desconsolidacao As Boolean = False) As Boolean
On Error GoTo LEErr:
Dim oRS  As ADODB.Recordset
Dim SQl  As String
Dim lTMP As Boolean
'VERIFICA SE HÁ PACOTES DE SOLICITAÇÃO PARA ALGUMA DAS CARGAS DO LOTE OU DE SEUS
'PARTE-LOTE. NÃO SERÁ UTILIZADO O FLAG_RECEITA DOS PACOTES, POIS SE HOUVER UM CANCELAMENTO
'DA SOLICITAÇÃO, OU A LIBERAÇÃO DO PACOTE DE ENVIO, ESTE PODERIA RETORNAR UM NULO/ZERO,
'IMPEDINDO A EDIÇÃO, QUE NESTE CASO É PERMITIDA.
'PARA RESOLVER, EM SEU LUGAR O QUE É CONSISTIDO É O AUDIT_ST DAS TABELAS DE CS/CNTR.
SQl = Empty
SQl = SQl & " SELECT "
SQl = SQl & " AUDIT_ST, "
SQl = SQl & " AUTONUM AS CN, "
SQl = SQl & " 0 AS CS "
SQl = SQl & " FROM "
SQl = SQl & " " & Banco_Sgipa & "TB_CNTR_BL "
SQl = SQl & " WHERE "
SQl = SQl & " NOT AUDIT_ST IS NULL "
SQl = SQl & " AND AUTONUM IN "
SQl = SQl & " ("
SQl = SQl & " SELECT "
SQl = SQl & " CNTR "
SQl = SQl & " FROM "
SQl = SQl & " " & Banco_Sgipa & "TB_AMR_CNTR_BL "
SQl = SQl & " WHERE "
SQl = SQl & " BL = '" & Lote & "'"
SQl = SQl & " ) "
SQl = SQl & " UNION ALL "
SQl = SQl & " SELECT "
SQl = SQl & " AUDIT_ST, "
SQl = SQl & " 0 AS CN, "
SQl = SQl & " AUTONUM AS CS "
SQl = SQl & " FROM "
SQl = SQl & " " & Banco_Sgipa & "TB_CARGA_SOLTA "
SQl = SQl & " WHERE "
SQl = SQl & " NOT AUDIT_ST IS NULL "
SQl = SQl & " AND BL = '" & Lote & "' "
SQl = SQl & " OR BL IN("
SQl = SQl & " SELECT "
SQl = SQl & " BL "
SQl = SQl & " FROM "
SQl = SQl & " " & Banco_Sgipa & "TB_AMR_CNTR_BL "
SQl = SQl & " WHERE "
SQl = SQl & " CNTR "
SQl = SQl & " IN ( "
SQl = SQl & " SELECT "
SQl = SQl & " CNTR "
SQl = SQl & " FROM "
SQl = SQl & " " & Banco_Sgipa & "TB_AMR_CNTR_BL "
SQl = SQl & " WHERE "
SQl = SQl & " BL = '" & Lote & "' "
SQl = SQl & " ) "
SQl = SQl & " ) "
''clipboard.SetText SQL  -  DESCOMENTAR PARA TESTAR A INSTRUÇÃO
PRSet oRS, SQl, 0
'RETORNOU REGISTROS? ELES TÊM SOLICITAÇÃO EFETUADA, CASO O CAMPO
'AUDIT_ST NÃO SEJA ZERO OU NULO.
 Do While Not oRS.EOF = True
    If Not lTMP = True Then lTMP = Not (Nnull(oRS.Fields("AUDIT_ST"), 0) = 0)
    If lTMP = True Then Exit Do
    oRS.MoveNext
Loop
oRS.Close
If lTMP = True Then
    If Desconsolidacao = True Then
        'DESCONSOLIDADO: SÓ VERIFICA SE HOUVER SOLICITAÇÕES.
        'CASO O VALOR DO FLAG SEJA 1, AS ALTERAÇÕES DEVERÃO SER POSSÍVEIS, INDEPENDENTE DE
        'QUAISQUER OUTROS FATORES.
        'A FUNÇÃO DEVE SEMPRE RETORNAR FALSE NESTES CASOS.
        SQl = Empty
        SQl = SQl & " SELECT "
        SQl = SQl & " NVL(AUDIT_DESCONSOLIDADO,0) AS QUAL "
        SQl = SQl & " FROM "
        SQl = SQl & " " & Banco_Sgipa & "TB_BL "
        SQl = SQl & " WHERE "
        SQl = SQl & " AUTONUM = (SELECT BL_MASTER FROM " & Banco_Sgipa & "TB_BL WHERE AUTONUM = '" & Lote & "') "
        ''clipboard.SetText SQL  -  DESCOMENTAR PARA TESTAR A INSTRUÇÃO
        PRSet oRS, SQl, 0
        'CASO RETORNE ALGUM REGISTRO, VERIFICAR O FLAG_DESCONSOLIDADO.
        'CASO CONTRÁRIO, O BL NÃO EXISTIRIA E HÁ ALGUM PROBLEMA NA CAPTAÇÃO. RETORNA FALSE
        'POR PADRÃO, OU SEJA, COMO NÃO HÁ UM BL, NÃO HÁ SOLICITAÇÃO.
        If Not oRS.EOF = True Then
            lTMP = (Nnull(oRS.Fields("QUAL"), 0) <> 0)
        Else
            lTMP = False
        End If
        oRS.Close
        'DESCONSOLIDADO: SÓ VERIFICA SE HOUVER SOLICITAÇÕES.
        'CASO O VALOR DO FLAG SEJA 1, AS ALTERAÇÕES DEVERÃO SER POSSÍVEIS, INDEPENDENTE DE
        'QUAISQUER OUTROS FATORES.
        'A FUNÇÃO DEVE SEMPRE RETORNAR FALSE NESTES CASOS.
        SQl = Empty
        SQl = SQl & " SELECT "
        SQl = SQl & " NVL(AUDIT_DESCONSOLIDADO,0) AS QUAL "
        SQl = SQl & " FROM "
        SQl = SQl & " " & Banco_Sgipa & "TB_BL "
        SQl = SQl & " WHERE "
        SQl = SQl & " AUTONUM = '" & Lote & "' "
        ''clipboard.SetText SQL  -  DESCOMENTAR PARA TESTAR A INSTRUÇÃO
        PRSet oRS, SQl, 0
        'CASO RETORNE ALGUM REGISTRO, VERIFICAR O FLAG_DESCONSOLIDADO.
        'CASO CONTRÁRIO, O BL NÃO EXISTIRIA E HÁ ALGUM PROBLEMA NA CAPTAÇÃO. RETORNA FALSE
        'POR PADRÃO, OU SEJA, COMO NÃO HÁ UM BL, NÃO HÁ SOLICITAÇÃO.
        If Not oRS.EOF = True Then
            If Not lTMP = True Then lTMP = (Nnull(oRS.Fields("QUAL"), 0) <> 0)
        Else
            lTMP = False
        End If
        oRS.Close
    End If
End If
LoteSolicitado = lTMP
GoTo LEExit:
LEErr:
MsgBox Err.Description
Err.Clear
LEExit:
SQl = Empty
lTMP = Empty
Set oRS = Nothing
End Function








Public Function LoteBloqueado(BL As Long) As String
Dim oRS     As ADODB.Recordset
Dim sBL     As String
'BLOQUEIOS POR B/L
sSql = Empty
sSql = sSql & "SELECT "
sSql = sSql & "NVL(A.FLAG_BLOQUEIO,0) AS FLAG_BLOQUEIO, "
sSql = sSql & "NVL(A.FLAG_BLOQUEIO_MANUAL,0) AS FLAG_BLOQUEIO_MANUAL, "
sSql = sSql & "NVL(A.NUMERO, ' ') AS NUMERO, "
sSql = sSql & "'[BLOQUEADO EM: ' || TO_CHAR(B.DATA,'DD/MM/RR')||' '||B.HORA|| ']     [TIPO DE BLOQUEIO: ' || B.TIPO || ']     [MOTIVO: ' || B.MOTIVO || ']     [BLOQUEADO POR: '|| B.USUARIO_SIS ||']' AS OBS "
sSql = sSql & "FROM "
sSql = sSql & "" & Banco_Sgipa & "TB_BL A, "
sSql = sSql & "(SELECT * FROM " & Banco_Sgipa & "TB_HIST_BLOQUEIO WHERE FLAG_BLOQUEIA = 1) B "
sSql = sSql & "WHERE "
sSql = sSql & "A.AUTONUM = " & BL & " "
sSql = sSql & "and  A.flag_ativo=1 "
sSql = sSql & "AND A.AUTONUM =  B.BL "
PRSet oRS, (sSql), 0
sBL = Empty
While Not oRS.EOF = True
    If (oRS.Fields("FLAG_BLOQUEIO") <> 0) Or _
       (oRS.Fields("FLAG_BLOQUEIO_MANUAL") <> 0) Then
        If Cod_Empresa = 1 Then sBL = sBL & oRS.Fields("NUMERO") & IIf((oRS.Fields("FLAG_BLOQUEIO_MANUAL") <> 0), " - BLOQUEIO IPA", " - BLOQUEIO DTE") & vbCrLf & vbTab & Replace(Nnull(oRS.Fields("OBS"), 1), "     ", Empty) & vbCrLf
        If Cod_Empresa = 2 Then sBL = sBL & oRS.Fields("NUMERO") & IIf((oRS.Fields("FLAG_BLOQUEIO_MANUAL") <> 0), " - BLOQUEIO IPA", " - BLOQUEIO DTE") & vbCrLf & vbTab & Replace(Nnull(oRS.Fields("OBS"), 1), "     ", Empty) & vbCrLf
    End If
    oRS.MoveNext
Wend
oRS.Close
LoteBloqueado = sBL
Set oRS = Nothing
sBL = Empty
End Function

Public Function ConteinerBloqueado(CNTR As Long) As String
Dim oRS     As ADODB.Recordset
Dim sOBS    As String
'BLOQUEIOS POR CONTEINER
sSql = Empty
sSql = sSql & " SELECT "
sSql = sSql & " C.ID_CONTEINER AS CONTEINER, "
sSql = sSql & " C.FLAG_BLOQUEIO AS BCM, "
sSql = sSql & " C.FLAG_BLOQUEIO_MANUAL AS BC, "
If Banco_Em_Uso = "ORACLE" Then
    sSql = sSql & " '[BLOQUEADO EM: ' || D.DATA || ' ' || D.HORA || ']     [TIPO DE BLOQUEIO: ' || D.TIPO || ']     [MOTIVO: ' || D.MOTIVO || ']     [BLOQUEADO POR: '|| D.USUARIO_SIS ||']' AS OBS "
Else
    sSql = sSql & " '[BLOQUEADO EM: ' + cast(D.DATA as varchar) + ' ' + cast(D.HORA as varchar) + ']     [TIPO DE BLOQUEIO: ' + D.TIPO + ']     [MOTIVO: ' + D.MOTIVO + ']     [BLOQUEADO POR: '+ D.USUARIO_SIS +']' AS OBS "
End If
sSql = sSql & " FROM "
sSql = sSql & " " & Banco_Sgipa & "TB_CNTR_BL C, "
sSql = sSql & " (SELECT * FROM " & Banco_Sgipa & "TB_HIST_BLOQUEIO WHERE FLAG_BLOQUEIA = 1) D "
sSql = sSql & " WHERE "
sSql = sSql & " C.AUTONUM = D.CNTR "
sSql = sSql & " AND C.AUTONUM = " & CNTR & " "
PRSet oRS, (sSql), 0
While Not oRS.EOF = True
    If (Nnull(oRS.Fields("BC"), 0) <> 0) Or _
       (Nnull(oRS.Fields("BCM"), 0) <> 0) Then
        sOBS = sOBS & vbCrLf
        If Cod_Empresa = 1 Then sOBS = sOBS & Nnull(oRS.Fields("CONTEINER"), 1) & IIf((Nnull(oRS.Fields("BCM"), 0) <> 0), " - BLOQUEIO IPA ", " - BLOQUEIO DTE") & vbCrLf & vbTab & Replace(Nnull(oRS.Fields("OBS"), 1), "     ", vbCrLf & vbTab)
        If Cod_Empresa = 2 Then sOBS = sOBS & Nnull(oRS.Fields("CONTEINER"), 1) & IIf((Nnull(oRS.Fields("BCM"), 0) <> 0), " - BLOQUEIO IPA", " - BLOQUEIO DTE") & vbCrLf & vbTab & Replace(Nnull(oRS.Fields("OBS"), 1), "     ", vbCrLf & vbTab)
    End If
    oRS.MoveNext
Wend
oRS.Close
ConteinerBloqueado = sOBS
Set oRS = Nothing
sOBS = Empty
End Function

Public Function PreparaReport(Arquivo As String, Optional Query As String, Optional Seleção As String, Optional Barra As Boolean = False, Optional SubRep As String) 'As Crystal.CrystalReport
Dim t    As Integer
'Dim oRPT As Crystal.CrystalReport
'Screen.MousePointer = vbHourglass
'Set oRPT = Principal.CrystalReport1
'If oRPT.SelectionFormula <> Empty Then oRPT.SelectionFormula = Empty
''limpareport oRPT
'Set oRPT = Principal.CrystalReport1
'If Dir(Arquivo) <> Empty Then
'    oRPT.ReportFileName = Arquivo
'    oRPT.Connect = Connect
'    If Query <> Empty Then oRPT.SqlQuery = Query
'    oRPT.WindowShowGroupTree = Barra
'    oRPT.WindowShowCancelBtn = True
'    oRPT.WindowShowCloseBtn = True
'    oRPT.WindowShowExportBtn = True
'    oRPT.WindowShowNavigationCtls = True
'    oRPT.WindowShowPrintBtn = True
'    oRPT.WindowShowPrintSetupBtn = True
'    oRPT.WindowShowProgressCtls = True
'    oRPT.WindowShowRefreshBtn = True
'    oRPT.WindowShowSearchBtn = True
'    oRPT.WindowShowZoomCtl = True
'    oRPT.Connect = Connect
'    If Seleção <> Empty And Query = "" Then oRPT.ReplaceSelectionFormula Seleção
'    oRPT.Connect = Connect
'    If oRPT.GetNSubreports > 0 Then
'        For t = 1 To oRPT.GetNSubreports
'          If UCase(oRPT.GetNthSubreportName(t - 1)) <> "LOG.RPT" And InStr(UCase(oRPT.GetNthSubreportName(t - 1)), "RODAPE") = 0 Then
'            oRPT.SubreportToChange = oRPT.GetNthSubreportName(t - 1)
'            oRPT.Connect = Connect
'            End If
'            If UCase(SubRep) = UCase(oRPT.GetNthSubreportName(t - 1)) Then
'               If Seleção <> Empty Then
'                oRPT.SelectionFormula = Seleção
'              Else
'               If Query <> Empty Then oRPT.SqlQuery = Query
'              End If
'            End If
'        Next t
'    End If
'    oRPT.SubreportToChange = Empty
'    oRPT.PageZoom 150
'    Set PreparaReport = oRPT
'Else
'    Set PreparaReport = Nothing
'End If
't = Empty
'Set oRPT = Nothing
'Screen.MousePointer = vbDefault
End Function

'Função auxiliar
Public Sub limpareport() '(ByRef obj As Crystal.CrystalReport)
Screen.MousePointer = vbHourglass
Dim t As Integer
'If Not obj Is Nothing Then
'    obj.WindowShowCancelBtn = False
'    obj.WindowShowCloseBtn = False
'    obj.WindowShowExportBtn = False
'    obj.WindowShowGroupTree = False
'    obj.WindowShowNavigationCtls = False
'    obj.WindowShowPrintBtn = False
'    obj.WindowShowPrintSetupBtn = False
'    obj.WindowShowProgressCtls = False
'    obj.WindowShowRefreshBtn = False
'    obj.WindowShowSearchBtn = False
'    obj.WindowShowZoomCtl = False
'    obj.PageZoom 0
'    For t = 0 To 255
'        obj.Formulas(t) = ""
'        obj.Formulas(t) = Empty
'        obj.SortFields(t) = Empty
'    Next t
'    obj.SqlQuery = Empty
'    If Not obj.ReportFileName = Empty Then
'        obj.ReplaceSelectionFormula Empty
'    End If
'    obj.Connect = Empty
'    obj.ReportFileName = Empty
'    obj.Destination = crptToWindow
'    Set obj = Nothing
'End If
'For t = 0 To 255
'    Principal.CrystalReport1.Formulas(t) = Empty
'    Principal.CrystalReport2.Formulas(t) = Empty
'Next t
'Principal.CrystalReport1.WindowShowCancelBtn = False
'Principal.CrystalReport1.WindowShowCloseBtn = False
'Principal.CrystalReport1.WindowShowExportBtn = False
'Principal.CrystalReport1.WindowShowGroupTree = False
'Principal.CrystalReport1.WindowShowNavigationCtls = False
'Principal.CrystalReport1.WindowShowPrintBtn = False
'Principal.CrystalReport1.WindowShowPrintSetupBtn = False
'Principal.CrystalReport1.WindowShowProgressCtls = False
'Principal.CrystalReport1.WindowShowRefreshBtn = False
'Principal.CrystalReport1.WindowShowSearchBtn = False
'Principal.CrystalReport1.WindowShowZoomCtl = False
'Principal.CrystalReport1.PageZoom 0
'Principal.CrystalReport1.ReportFileName = Empty
'Principal.CrystalReport1.SqlQuery = Empty
'Principal.CrystalReport1.Connect = Empty
'If Not Principal.CrystalReport2.ReportFileName = Empty Then
'    Principal.CrystalReport2.ReplaceSelectionFormula Empty
'End If
'If Not Principal.CrystalReport1.ReportFileName = Empty Then
'    Principal.CrystalReport1.ReplaceSelectionFormula Empty
'End If
'Principal.CrystalReport2.WindowShowCancelBtn = False
'Principal.CrystalReport2.WindowShowCloseBtn = False
'Principal.CrystalReport2.WindowShowExportBtn = False
'Principal.CrystalReport2.WindowShowGroupTree = False
'Principal.CrystalReport2.WindowShowNavigationCtls = False
'Principal.CrystalReport2.WindowShowPrintBtn = False
'Principal.CrystalReport2.WindowShowPrintSetupBtn = False
'Principal.CrystalReport2.WindowShowProgressCtls = False
'Principal.CrystalReport2.WindowShowRefreshBtn = False
'Principal.CrystalReport2.WindowShowSearchBtn = False
'Principal.CrystalReport2.WindowShowZoomCtl = False
'Principal.CrystalReport2.PageZoom 0
'Principal.CrystalReport2.SqlQuery = Empty
'Principal.CrystalReport2.Connect = Empty
'Principal.CrystalReport2.ReportFileName = Empty
'For t = 0 To 255
'    Principal.CrystalReport1.Formulas(t) = ""
'    Principal.CrystalReport2.Formulas(t) = ""
'    Principal.CrystalReport1.Formulas(t) = Empty
'    Principal.CrystalReport2.Formulas(t) = Empty
'Next t
'Principal.CrystalReport2.Destination = crptToWindow
'Principal.CrystalReport1.Destination = crptToWindow
't = Empty
'Screen.MousePointer = vbDefault
End Sub







'ALTERADO EM 27/07/2005
Public Function GravaMotorista(Optional CNH As String, Optional Motorista As String, Optional RG As String, Optional Emissao_Rg As String, Optional Orgao As String, Optional Estrangeiro As Long, Optional Modelo As String, Optional Cor As String, Optional CPF As String) As Boolean
On Error GoTo GMErr:
Dim oRST As ADODB.Recordset
Dim lRET As Boolean
Dim WSQL As String
Dim nRec As Long
WSQL = Empty
CNH = Trim$(Replace$(Replace$(Replace$(UCase$(CNH), ".", Empty), "-", Empty), "/", Empty))
Motorista = Trim$(Replace$(Replace$(Replace$(UCase$(Motorista), "'", Empty), Chr$(34), Empty), "/", Empty))
RG = Trim$(Replace$(Replace$(Replace$(UCase$(RG), ".", Empty), "-", Empty), "/", Empty))
If Not IsDate(Emissao_Rg) = True Then Emissao_Rg = Format$(Now, "DD/MM/YYYY HH:NN:SS")
Orgao = Trim$(Replace$(Replace$(Replace$(UCase$(Orgao), ".", Empty), "-", Empty), "/", Empty))
Estrangeiro = Nnull(Estrangeiro, 0)
Modelo = Trim$(Replace$(Replace$(Replace$(UCase$(Modelo), "'", Empty), Chr$(34), Empty), "/", Empty))
Cor = Trim$(Replace$(Replace$(Replace$(UCase$(Cor), "'", Empty), Chr$(34), Empty), "/", Empty))
CPF = Trim$(Replace$(Replace$(Replace$(UCase$(CPF), ".", Empty), "-", Empty), "/", Empty))
PRSet oRST, "SELECT AUTONUM FROM " & Banco_Operador & "TB_MOTORISTAS WHERE CNH = '" & CNH & "'"
If oRST.EOF Then
    WSQL = WSQL & "INSERT INTO "
    WSQL = WSQL & "" & Banco_Operador & "TB_MOTORISTAS ("
    WSQL = WSQL & "autonum"
    WSQL = WSQL & ",NOME, "
    If Estrangeiro = 0 Then
        WSQL = WSQL & "CPF, "
    End If
    WSQL = WSQL & "CNH, "
    WSQL = WSQL & "RG, "
    WSQL = WSQL & "ORGAO, "
    WSQL = WSQL & "EMISSAO_RG, "
    WSQL = WSQL & "TRANSPORTADORA, "
    WSQL = WSQL & "FLAG_ESTRANGEIRO "
    WSQL = WSQL & ") VALUES ( "
    WSQL = WSQL & "" & Banco_Operador & "seq_tb_motoristas.nextval"
    WSQL = WSQL & ",'" & Motorista & "', "
    If Estrangeiro = 0 Then
        WSQL = WSQL & "'" & CPF & "', "
    End If
    WSQL = WSQL & "'" & CNH & "', "
    WSQL = WSQL & "'" & RG & "', "
    WSQL = WSQL & "'" & Orgao & "', "
    WSQL = WSQL & "TO_DATE('" & Emissao_Rg & "','dd/mm/yyyy'), "
    WSQL = WSQL & "'0', "
    WSQL = WSQL & "'" & Estrangeiro & "' "
    WSQL = WSQL & ")"
Else
    WSQL = WSQL & "UPDATE "
    WSQL = WSQL & "" & Banco_Operador & "TB_MOTORISTAS "
    WSQL = WSQL & "SET "
    WSQL = WSQL & "NOME = '" & Motorista & "', "
    If Estrangeiro = 0 Then
        WSQL = WSQL & "CPF = '" & CPF & "', "
    End If
    WSQL = WSQL & "CNH = '" & CNH & "', "
    WSQL = WSQL & "RG = '" & RG & "', "
    WSQL = WSQL & "ORGAO = '" & Orgao & "', "
    WSQL = WSQL & "EMISSAO_RG = TO_DATE('" & Emissao_Rg & "','dd/mm/yyyy'), "
    WSQL = WSQL & "FLAG_ESTRANGEIRO = '" & Estrangeiro & "' "
    WSQL = WSQL & "WHERE "
    WSQL = WSQL & "AUTONUM = " & oRST!autonum & " "
End If

'Executa wSQL, nRec

Executa WSQL

'GravaMotorista = Not (nRec = 0)
oRST.Close
GoTo GMExit:
GMErr:
MsgBox Err.Description
Err.Clear

GMExit:
Set oRST = Nothing
lRET = Empty
nRec = Empty
WSQL = Empty
End Function

Public Sub MostraChave(AutonumBL As Long)
 
 On Error GoTo MCErr:
Dim nErro          As Integer
Dim SQl            As String
Dim wPOL           As String
Dim wPoD           As String
Dim NumeroPresenca As String
Dim SMSG           As String
Dim oRS()          As ADODB.Recordset
ReDim oRS(5)
PRSet oRS(0), "SELECT A.CE,A.PRIMEIRA_ENTRADA,A.NUMERO,A.VIAGEM,A.PORTO_ORIGEM,A.PORTO_DESTINO,A.ARMADOR,C.NAVIO,to_date(A.DT_BL,'dd/mm/yy') as DATABL FROM " & Banco_Sgipa & "TB_BL A," & Banco_Sgipa & "DTE_TB_VIAGENS C WHERE A.VIAGEM=C.VIAGEM AND A.AUTONUM=" & AutonumBL & " ", 0
If Not IsDate(oRS(0)!PRIMEIrA_ENTRADA) Or Format(oRS(0)!PRIMEIrA_ENTRADA, "yyyymmdd") >= "20080331" Then
    NumeroPresenca = "CEMERCANTE" & "31032008" & Mid$(Nnull(oRS(0)!CE, 1) & Space(15), 1, 15) & Space(3)
Else
  If IsNull(oRS(0)!Databl) Then
    nErro = 1
    SMSG = "ATENÇÃO:" & vbCrLf & "A Data do Lote se Encontra em Branco." & vbCrLf & "O Sistema não pôde fornecer o Numero de Presença."
  Else
    PRSet oRS(1), "SELECT DESCR, TRIGRAMA FROM " & Banco_Sgipa & "DTE_TB_PORTOS WHERE code = " & Nnull(oRS(0)!Porto_Origem, 0) & " ", 0
    If oRS(1).EOF Then
       nErro = 1
       SMSG = "Não foi encontrado nenhum porto cadastrado com o código " & Nnull(oRS(0)!Porto_Origem, 0) & "." & vbCrLf & "Por favor, verifique."
    Else
        If Not IsNull(oRS(1)!TRIGRAMA) Then wPOL = oRS(1)!TRIGRAMA
        If wPOL = "" Then
            nErro = 2
            SMSG = "Porto " & Nnull(oRS(0)!Porto_Origem, 0) & " ( " & Nnull(oRS(1)!Descr, 1) & " ) está cadastrado, porém não consta o trigrama do mesmo. " & vbCrLf & "Por favor, verifique."
        Else
            wPoD = ""
            PRSet oRS(2), "select descr, trigrama from " & Banco_Sgipa & "DTE_TB_PORTOS where code=" & Nnull(oRS(0)!porto_destino, 0) & " ", 0
            If oRS(2).EOF Then
                nErro = 3
                SMSG = "Não foi encontrado nenhum porto cadastrado com o código " & Nnull(oRS(0)!porto_destino, 0) & "."
            Else
                If Not IsNull(oRS(2)!TRIGRAMA) Then wPoD = oRS(2)!TRIGRAMA
                If wPoD = "" Then
                    nErro = 4
                    SMSG = "Porto " & Nnull(oRS(0)!porto_destino, 0) & " ( " & Nnull(oRS(2)!Descr, 1) & " ) está cadastrado, porém não consta o trigrama do mesmo. " & vbCrLf & "Por favor, verifique."
                Else
                    If wBLInformado = "" Then
                        NumeroPresenca = Right$(Format(Val(oRS(0)!Armador), "0000"), 4) & wPOL & wPoD & Format(CVDate(oRS(0)!Databl), "ddmmyyyy") & Left$(Format(Val(oRS(0)!Navio), "0000000"), 7) & Trim(Right$(oRS(0)!Numero, 11))
                    Else
                        NumeroPresenca = Right$(Format(Val(oRS(0)!Armador), "0000"), 4) & wPOL & wPoD & Format(CVDate(wDataInformada), "ddmmyyyy") & Left$(Format(Val(oRS(0)!Navio), "0000000"), 7) & Right$(wBLInformado, 11)
                    End If
                End If
            End If
            oRS(2).Close
        End If
    End If
    oRS(1).Close
  End If
End If
oRS(0).Close
If Not nErro > 0 Then
    MsgBox NumeroPresenca
Else
    MsgBox "ATENÇÃO:" & vbCrLf & "Foi encontrado um problema de processo [Falha de n°" & Format$(nErro, "00") & "]" & vbCrLf & vbCrLf & SMSG & vbCrLf & vbCrLf, vbOKOnly + vbInformation, "Cálculo da Chave de Presença"
End If
GoTo MCExit:
MCErr:
MsgBox Err.Description
Err.Clear
MCExit:
Erase oRS
nErro = Empty
SQl = Empty
wPOL = Empty
wPoD = Empty
NumeroPresenca = Empty
SMSG = Empty
End Sub




Public Function Verifica_Carga_Liberada(Motivo As Byte, Lote As Long, Optional CNTR As Long = 0) As Boolean
    
'    Verifica_Carga_Liberada = False
'vinculo
    Dim RsTestar As New ADODB.Recordset

    SQl = "select * from " & Banco_Sgipa & "tb_liberacao_cargas where"
    SQl = SQl & " bl=" & Lote
    SQl = SQl & " and cod_motivo=" & Motivo
    If CNTR <> 0 Then
        SQl = SQl & " and cntr=" & CNTR
    End If
    PRSet RsTestar, SQl
    If RsTestar.EOF Then
        Verifica_Carga_Liberada = False
    Else
        If RsTestar!Status = 1 Then
            Verifica_Carga_Liberada = True
        Else
            Verifica_Carga_Liberada = False
        End If
    End If
    RsTestar.Close

End Function

Public Sub TrocaCommand(Tela As Form)
  Dim i As Integer
  
  On Error Resume Next
  Tela.CmdNovo.Enabled = Not (Tela.CmdNovo.Enabled)
  Tela.CmdEditar.Enabled = Not (Tela.CmdEditar.Enabled)
  Tela.cmdGravar.Enabled = Not (Tela.cmdGravar.Enabled)
  Tela.cmdCancelar.Enabled = Not (Tela.cmdCancelar.Enabled)
  Tela.CmdExcluir.Enabled = Not (Tela.CmdExcluir.Enabled)
  Tela.CmdSair.Enabled = Not (Tela.CmdSair.Enabled)
  Tela.CmdFiltro.Enabled = Not (Tela.CmdFiltro.Enabled)
  Tela.DataGrid1.Enabled = Not (Tela.DataGrid1.Enabled)
  Tela.flexProcesso.Enabled = Not Tela.flexProcesso.Enabled
  For i = 0 To 3
    Tela.cmdNavega(i).Enabled = Not (Tela.cmdNavega(i).Enabled)
  Next i
End Sub

Public Function ControlarEdicao(Tela As Form, Tipo As String, Num As Integer, Rs As Recordset)
Dim Controle As Object
Dim i As Integer
Dim Oper As Boolean
Dim oldStyle
Dim strConteudoAnterior As String

On Error Resume Next

If Tipo = "Novo" Then Oper = True
If Tipo = "Editar" Then Oper = True
If Tipo = "Gravar" Then Oper = False
If Tipo = "Cancelar" Then Oper = False
Tela.fraDados1.Enabled = Oper
Tela.fraDados2.Enabled = Oper

For i = 0 To Tela.Controls.Count - 1
  Set Controle = Tela.Controls(i)
  Err.Number = 0
          
          


          
  If Controle.Tag <> "" And Controle.HelpContextID = Num Then
    If Err.Number = 0 Then
      If Not TypeOf Controle Is CommandButton Then
        Controle.Enabled = Oper
      End If
      
      If Rs.Fields(Controle.Tag).Type >= 200 Then
        Controle.MaxLength = Rs.Fields(Controle.Tag).DefinedSize
      Else
        If Rs.Fields(Controle.Tag).Type = 201 Then
          Controle.MaxLength = 0
        Else
          If Rs(Controle.Tag).Type = 7 _
             Or Rs(Controle.Tag).Type = 133 _
             Or Rs(Controle.Tag).Type = 134 _
             Or Rs(Controle.Tag).Type = 135 Then
            Controle.MaxLength = Rs.Fields(Controle.Tag).DefinedSize
          Else
            If Rs(Controle.Tag).Type = 5 _
               Or Rs(Controle.Tag).Type = 6 _
               Or Rs(Controle.Tag).Type = 131 Then
              Controle.MaxLength = 0
              If TypeOf Controle Is CheckBox Then
                
'MsgBox Controle.Tag
                
                Controle.Value = IIf(Nnull(Rs.Fields(Controle.Tag), 0).Value = 0, vbUnchecked, vbChecked)
              End If
            Else
              Controle.MaxLength = Rs.Fields(Controle.Tag).DefinedSize
            End If
          End If
        End If
      End If

      If Tipo = "Novo" Then
        If Rs.Fields(Controle.Tag).Type >= 200 Or Rs.Fields(Controle.Tag).Type = 201 Then
          If TypeOf Controle Is Label Then Controle.Caption = ""
          If TypeOf Controle Is TextBox Then Controle.text = ""
          If TypeOf Controle Is DataCombo Then Controle.text = ""
          If TypeOf Controle Is ComboBox Then Controle.text = ""
          If TypeOf Controle Is MaskEdBox Then
            strConteudoAnterior = Controle.Mask
            Controle.Mask = Empty
            Controle.text = Empty
            Controle.Mask = strConteudoAnterior
          End If
        Else
          If Rs(Controle.Tag).Type = 7 _
             Or Rs(Controle.Tag).Type = 133 _
             Or Rs(Controle.Tag).Type = 134 _
             Or Rs(Controle.Tag).Type = 135 Then
            If TypeOf Controle Is Label Then Controle.Caption = Date
            If TypeOf Controle Is TextBox Then Controle.text = ""
            If TypeOf Controle Is MaskEdBox Then
              strConteudoAnterior = Controle.Mask
              Controle.Mask = Empty
              Controle.text = Empty
              Controle.Mask = strConteudoAnterior
            End If
          Else
            If TypeOf Controle Is CheckBox Then Controle.Value = 0
            If TypeOf Controle Is Label Then Controle.Caption = "0"
            If TypeOf Controle Is TextBox Then Controle.text = ""
            If TypeOf Controle Is DataCombo Then Controle.text = ""
            If TypeOf Controle Is ComboBox Then Controle.ListIndex = 0
            If TypeOf Controle Is MaskEdBox Then
              strConteudoAnterior = Controle.Mask
              Controle.Mask = Empty
              Controle.text = Empty
              Controle.Mask = strConteudoAnterior
            End If
          End If
        End If
      End If
      Err.Number = 0
    End If
  End If
  
  Err.Number = 0
  
  If Controle.Tag <> "" Then
    If Err.Number = 0 Then
      If Tipo = "Novo" Then
        If TypeOf Controle Is Label Then Controle.Caption = ""
      End If
    End If
  End If
Next

On Error GoTo 0
End Function

Public Sub Desenhar_Tela_Padrao(Tela As Form)
On Error Resume Next
Tela.Frame1.Left = (Tela.Width - Tela.Frame1.Width) / 2
Tela.Frame1.Top = (Tela.Height - Tela.Frame1.Height) / 2
Tela.lblTITULO.Top = Tela.Frame1.Top - 300
Tela.lblTITULO.Left = Tela.Frame1.Left
Tela.lblTITULO.Width = Tela.Frame1.Width
Tela.lblTITULO.Caption = UCase(Tela.Caption)
Tela.lblTITULO.Height = 300
Tela.lblTITULO.Font = "Lucida Console"
Tela.lblTITULO.FontBold = False
Tela.lblTITULO.FontSize = 9
End Sub

Public Function CarregaConjuntoControles(Tela As Form, _
                                         ControlContainer As Control, _
                                         ByVal Rs As Recordset, _
                                         Optional CarregarRecordset As Boolean = False, _
                                         Optional ValCampoChave As String = Empty)
Dim Controle As Object
Dim i As Integer
Dim Tipo As String
Dim Zeros As String
Dim oldMask As String
Dim varSplit As Variant
Dim strCampoChave As String
Dim strTabela As String
Dim Num As Variant

If CarregarRecordset Then
  varSplit = Split(ControlContainer.Tag, ";")
  SQl = "SELECT * FROM " & varSplit(0)
  SQl = SQl & " WHERE " & varSplit(1) & " = " & ValCampoChave
  
  PRSet Rs, SQl
End If

On Error Resume Next
If Not Rs.EOF And Not Rs.BOF Then
  For i = 0 To Tela.Controls.Count - 1
    Set Controle = Tela.Controls(i)
    Err.Number = 0
    If Controle.Container.Name & "(" & Controle.Container.Index & ")" = ControlContainer.Name & "(" & ControlContainer.Index & ")" Then
      If Controle.Tag <> "" Then
        If Err.Number = 0 Or (TypeOf Controle Is Label) Then
          If Rs.Fields(Controle.Tag).Type >= 200 Or Rs.Fields(Controle.Tag).Type = 201 Then
            Tipo = "T"
          Else
            If Rs(Controle.Tag).Type = 7 _
               Or Rs(Controle.Tag).Type = 133 _
               Or Rs(Controle.Tag).Type = 134 _
               Or Rs(Controle.Tag).Type = 135 Then
              Tipo = "D"
            Else
              Tipo = "V"
            End If
          End If
          
          If Tipo = "T" Then
            If TypeOf Controle Is DataCombo Then
              Controle.BoundText = Nnull(Rs.Fields(Controle.Tag), 1)
            Else
              If TypeOf Controle Is Label Then
                Controle.Caption = Nnull(Rs.Fields(Controle.Tag), 1)
              Else
                Controle.text = Nnull(Rs.Fields(Controle.Tag), 1)
              End If
            End If
          End If
          If Tipo = "D" Then
            If TypeOf Controle Is Label Then
              Controle.Caption = Format(Nnull(Rs.Fields(Controle.Tag), 1), "DD/MM/YYYY")
            ElseIf TypeOf Controle Is MaskEdBox Then
              oldMask = Controle.Mask
              Controle.Mask = Empty
              Controle.text = Empty
              Controle.Mask = oldMask
              Controle.text = Format(Nnull(Rs.Fields(Controle.Tag), 1), "DD/MM/YYYY")
            Else
              Controle.text = Format(Nnull(Rs.Fields(Controle.Tag), 1), "DD/MM/YYYY")
            End If
          End If
          If Tipo = "V" Then
            If TypeOf Controle Is DataCombo Then
              Controle.BoundText = Nnull(Rs.Fields(Controle.Tag), 1)
            Else
              If Rs.Fields(Controle.Tag).NumericScale > 0 Then
                Zeros = Mid$("0000000000", 1, IIf(Rs.Fields(Controle.Tag).NumericScale <> 255, Rs.Fields(Controle.Tag).NumericScale, 2))
                If TypeOf Controle Is Label Then
                  Controle.Caption = Nnull(Rs.Fields(Controle.Tag), 1)
                Else
                  If TypeOf Controle Is ComboBox Then
                    Controle.text = Nnull(Rs.Fields(Controle.Tag), 1)
                  Else
                    Controle.text = Nnull(Rs.Fields(Controle.Tag), 1)
                  End If
                End If
              End If
              If Rs.Fields(Controle.Tag).NumericScale = 0 Or Rs.Fields(Controle.Tag).NumericScale <> 255 Then
                If TypeOf Controle Is Label Then
                  Controle.Caption = Nnull(Rs.Fields(Controle.Tag), 1)
                ElseIf TypeOf Controle Is CheckBox Then
                  Controle.Value = CBool(Nnull(Rs.Fields(Controle.Tag), 1))
                Else
                  Controle.text = Nnull(Rs.Fields(Controle.Tag), 1)
                End If
              End If
              If Rs.Fields(Controle.Tag).NumericScale = 0 Or Rs.Fields(Controle.Tag).NumericScale = 255 Then
                If TypeOf Controle Is CheckBox Then
                  Controle.Value = IIf(Nnull(Rs.Fields(Controle.Tag), 0) = 0, vbUnchecked, vbChecked)
                End If
              End If
            End If
          End If
          Err.Number = 0
        End If
      End If
    End If
  Next
Else
  For i = 0 To Tela.Controls.Count - 1
    Set Controle = Tela.Controls(i)
    Err.Number = 0
    If Controle.Tag <> "" And Controle.HelpContextID = Num Then
      If Err.Number = 0 Or (TypeOf Controle Is Label) Then
        If TypeOf Controle Is DataCombo Then
          Controle.BoundText = ""
        Else
          Controle.text = ""
        End If
        Err.Number = 0
      End If
    End If
  Next
End If

On Error GoTo 0
End Function

Public Sub Filtrar_Dados02(DC As Object, Optional blnDCVazio As Boolean = False)
  If Not blnDCVazio Then
    If TypeOf DC Is Adodc Then
      If Not DC.Recordset.Filter <> Empty Then
        If DC.Recordset.RecordCount = 0 Then Exit Sub
      End If
    ElseIf TypeOf DC Is ADODB.Recordset Then
      If Not DC.Filter <> Empty Then
        If DC.RecordCount = 0 Then Exit Sub
      End If
    End If
  End If
  
  Load frmFiltro02
  Set frmFiltro02.oRSFILTRO = DC.Recordset
  frmFiltro02.Show 1
  Set frmFiltro02.oRSFILTRO = Nothing
End Sub

Public Sub Formata_Insert_Container(Tela As Form, _
                                    Container As Object)
  Dim Controle As Object
  Dim i As Integer
  Dim Num As Integer
  
  On Error Resume Next
  For i = 0 To Tela.Controls.Count - 1
    Set Controle = Tela.Controls(i)
    
    If Controle.Tag <> "" Then
      If Not TypeOf Controle Is CommandButton Then
        If Not TypeOf Controle Is Frame Then
          If Controle.Container.Name = Container.Name Then
            If Controle.Container.Index = Container.Index Then
              SQl = SQl & Controle.Tag & ","
            End If
          End If
        End If
      End If
    End If
  Next i
End Sub

Public Sub Formata_SQL_Container(Tela As Form, _
                                 Container As Object, _
                                 Operacao As String)
  Dim Tipo As String
  Dim LocalCampo As Variant
  Dim Controle As Object
  Dim i As Integer
  Dim Num As Integer
  Dim Rs As ADODB.Recordset
  Dim varTabela As Variant
  Dim strSQL As String
  
  varTabela = Split(Container.Tag, ";")
  
  strSQL = "SELECT * FROM " & varTabela(0) & " WHERE 0 <> 0"
  PRSet Rs, strSQL
  
  On Error Resume Next
  For i = 0 To Tela.Controls.Count - 1
    Set Controle = Tela.Controls(i)
    
    Err.Number = 0
    
    If Not TypeOf Controle Is CommandButton And Controle.Tag <> "" And Controle.HelpContextID = Num Then
      If Controle.Container.Name = Container.Name And Controle.Container.Index = Container.Index Then
        If Err.Number = 0 Or (TypeOf Controle Is Label) Then
          If Controle.Tag <> "" Then
            If Rs.Fields(Controle.Tag).Type = 202 _
              Or Rs.Fields(Controle.Tag).Type = 203 _
              Or Rs.Fields(Controle.Tag).Type = 200 _
              Or Rs.Fields(Controle.Tag).Type = 201 Then
              Tipo = "T"
            Else
              If Rs.Fields(Controle.Tag).Type = 7 _
                Or Rs.Fields(Controle.Tag).Type = 133 _
                Or Rs.Fields(Controle.Tag).Type = 135 Then
                Tipo = "D"
              Else
                Tipo = "V"
              End If
            End If
            
            If Operacao = "INSERT" Then
              Select Case Tipo
                Case "V"
                  If TypeOf Controle Is CheckBox Then
                    SQl = SQl & IIf(Nnull(PPonto(Controle.Value), 0) = 0, 0, 1) & ","
                  Else
                    If TypeOf Controle Is DataCombo Then
                      SQl = SQl & Nnull(PPonto(Controle.BoundText), 0) & ","
                    Else
                      If TypeOf Controle Is Label Then
                        SQl = SQl & Nnull(PPonto(Controle.Caption), 0) & ","
                      Else
                        SQl = SQl & Nnull(PPonto(Controle.text), 0) & ","
                      End If
                    End If
                  End If
                 
                Case "T"
                  LocalCampo = Tira_AspaSimples(Nnull(Controle.text, 1))
                  If TypeOf Controle Is Label Then LocalCampo = Tira_AspaSimples(Nnull(Controle.Caption, 1))
                  If TypeOf Controle Is DataCombo Then LocalCampo = Controle.BoundText
                  SQl = SQl & "'" & Nnull(LocalCampo, 1) & "',"
                
                Case "D"
                  Controle.text = Nnull(Controle.text, 1)
                  If Controle.text <> "" Then
                    SQl = SQl & " TO_DATE('" & Format$(Controle.text, "DD/MM/YYYY") & "', 'DD/MM/YYYY')" & ","
                  Else
                    SQl = SQl & " NULL, "
                  End If
              End Select
            End If
              
            If Operacao = "UPDATE" Then
              Select Case Tipo
                Case "T"
                  LocalCampo = Tira_AspaSimples(Nnull(Controle.text, 1))
                  If TypeOf Controle Is Label Then LocalCampo = Tira_AspaSimples(Nnull(Controle.Caption, 1))
                  If TypeOf Controle Is DataCombo Then LocalCampo = Controle.BoundText
                  SQl = SQl & Controle.Tag & " = " & "'" & Nnull(LocalCampo, 1) & "',"
                Case "V"
                  If TypeOf Controle Is CheckBox Then
                    SQl = SQl & Controle.Tag & " = " & IIf(Nnull(Controle.Value, 0) = 0, 0, 1) & ","
                  Else
                    If TypeOf Controle Is DataCombo Then
                      SQl = SQl & Controle.Tag & " = " & Nnull(PPonto(Controle.BoundText), 0) & ","
                    Else
                      If TypeOf Controle Is ComboBox Then
                        SQl = SQl & Controle.Tag & " = " & Nnull(PPonto(Controle.text), 0) & ","
                      Else
                        If TypeOf Controle Is Label Then
                          SQl = SQl & Controle.Tag & " = " & Nnull(PPonto(Controle.Caption), 0) & ","
                        Else
                          SQl = SQl & Controle.Tag & " = " & Nnull(PPonto(Controle.text), 0) & ","
                        End If
                      End If
                    End If
                  End If
                  
                Case "D"
                  If Controle.text <> "" Then
                    SQl = SQl & Controle.Tag & " = " & " TO_DATE('" & Format(Controle.text, "DD/MM/YYYY") & "', 'DD/MM/YYYY')" & ","
                  Else
                    SQl = SQl & Controle.Tag & " = NULL, "
                  End If
              End Select
            End If
          End If
        End If
      End If
    End If
  Next i
  
  SQl = Trim$(SQl)
  If Right$(SQl, 1) = "," Then
    SQl = Left$(SQl, (Len(SQl) - 1))
  End If
End Sub

Public Function CarregaTela(Tela As Form, Rs As Recordset, Num As Integer)
Dim Controle As Object
Dim i As Integer
Dim Tipo As String
Dim Zeros As String
Dim oldMask As String

On Error Resume Next
If Not Rs.EOF And Not Rs.BOF Then
  For i = 0 To Tela.Controls.Count - 1
    Set Controle = Tela.Controls(i)
    Err.Number = 0
    If Controle.Tag <> "" Then
      If Err.Number = 0 Or (TypeOf Controle Is Label) Then
        If Rs.Fields(Controle.Tag).Type >= 200 Or Rs.Fields(Controle.Tag).Type = 201 Then
          Tipo = "T"
        Else
          If Rs(Controle.Tag).Type = 7 _
             Or Rs(Controle.Tag).Type = 133 _
             Or Rs(Controle.Tag).Type = 134 _
             Or Rs(Controle.Tag).Type = 135 Then
            Tipo = "D"
          Else
            Tipo = "V"
          End If
        End If
        
        If Tipo = "T" Then
          If TypeOf Controle Is DataCombo Then
            Controle.BoundText = Nnull(Rs.Fields(Controle.Tag), 1)
          Else
            If TypeOf Controle Is Label Then
              Controle.Caption = Nnull(Rs.Fields(Controle.Tag), 1)
            Else
              Controle.text = Nnull(Rs.Fields(Controle.Tag), 1)
            End If
          End If
        End If
        If Tipo = "D" Then
          If TypeOf Controle Is Label Then
            Controle.Caption = Format(Nnull(Rs.Fields(Controle.Tag), 1), "DD/MM/YYYY")
          ElseIf TypeOf Controle Is MaskEdBox Then
            oldMask = Controle.Mask
            Controle.Mask = Empty
            Controle.text = Empty
            Controle.Mask = oldMask
            Controle.text = Format(Nnull(Rs.Fields(Controle.Tag), 1), "DD/MM/YYYY")
          Else
            Controle.text = Format(Nnull(Rs.Fields(Controle.Tag), 1), "DD/MM/YYYY")
          End If
        End If
        If Tipo = "V" Then
          If TypeOf Controle Is DataCombo Then
            Controle.BoundText = Nnull(Rs.Fields(Controle.Tag), 1)
          Else
            If Rs.Fields(Controle.Tag).NumericScale > 0 Then
              Zeros = Mid$("0000000000", 1, IIf(Rs.Fields(Controle.Tag).NumericScale <> 255, Rs.Fields(Controle.Tag).NumericScale, 2))
              If TypeOf Controle Is Label Then
                Controle.Caption = Format(Nnull(Rs.Fields(Controle.Tag), 1), "#,##0." & Zeros)
              Else
                If TypeOf Controle Is ComboBox Then
                  Controle.text = Nnull(Rs.Fields(Controle.Tag), 1)
                Else
                  Controle.text = Format(Nnull(Rs.Fields(Controle.Tag), 1), "#,##0." & Zeros)
                End If
              End If
            End If
            If Rs.Fields(Controle.Tag).NumericScale = 0 Or Rs.Fields(Controle.Tag).NumericScale <> 255 Then
              If TypeOf Controle Is Label Then
                Controle.Caption = Format(Nnull(Rs.Fields(Controle.Tag), 1), "#,##0." & Zeros)
              ElseIf TypeOf Controle Is CheckBox Then
                Controle.Value = CBool(Nnull(Rs.Fields(Controle.Tag), 1))
              Else
                'Controle.Text = Format(Nnull(Rs.Fields(Controle.Tag), 1), "#,##0." & Zeros)
                Controle.text = Nnull(Rs.Fields(Controle.Tag), 1)
              End If
            End If
            If Rs.Fields(Controle.Tag).NumericScale = 0 Or Rs.Fields(Controle.Tag).NumericScale = 255 Then
              If TypeOf Controle Is CheckBox Then
'MsgBox Controle.Tag
                
                Controle.Value = IIf(Nnull(Rs.Fields(Controle.Tag), 0) = 0, vbUnchecked, vbChecked)
              End If
            End If
          End If
        End If
        Err.Number = 0
      End If
    End If
  Next
Else
  For i = 0 To Tela.Controls.Count - 1
    Set Controle = Tela.Controls(i)
    Err.Number = 0
    If Controle.Tag <> "" And Controle.HelpContextID = Num Then
      If Err.Number = 0 Or (TypeOf Controle Is Label) Then
        If TypeOf Controle Is DataCombo Then
          Controle.BoundText = ""
        Else
          Controle.text = ""
        End If
        Err.Number = 0
      End If
    End If
  Next
End If

On Error GoTo 0
End Function

Public Sub Formata_Insert(Tela As Form)
  Dim Controle As Object
  Dim i As Integer
  Dim Num As Integer
  
  On Error Resume Next
  For i = 0 To Tela.Controls.Count - 1
    Set Controle = Tela.Controls(i)
    Err.Number = 0
    If Not TypeOf Controle Is CommandButton And Controle.Tag <> "" And Controle.HelpContextID = Num Then
      If Err.Number = 0 Or (TypeOf Controle Is Label) Then
        If Trim$(Controle.Tag) <> Empty Then
          SQl = SQl & Controle.Tag & ","
        End If
      End If
    End If
  Next i
End Sub
Public Sub Formata_SQL(Tela As Form, Rs As Recordset, Operacao As String)
  Dim Tipo As String
  Dim LocalCampo As Variant
  Dim Controle As Object
  Dim i As Integer
  Dim Num As Integer
  
  On Error Resume Next
  For i = 0 To Tela.Controls.Count - 1
    Set Controle = Tela.Controls(i)
    
    Err.Number = 0
    
    If Not TypeOf Controle Is CommandButton And Controle.Tag <> "" And Controle.HelpContextID = Num Then
      If Err.Number = 0 Or (TypeOf Controle Is Label) Then
        If Controle.Tag <> "" Then
          If Rs.Fields(Controle.Tag).Type = 202 _
            Or Rs.Fields(Controle.Tag).Type = 203 _
            Or Rs.Fields(Controle.Tag).Type = 200 _
            Or Rs.Fields(Controle.Tag).Type = 201 Then
            Tipo = "T"
          Else
            If Rs.Fields(Controle.Tag).Type = 7 _
              Or Rs.Fields(Controle.Tag).Type = 133 _
              Or Rs.Fields(Controle.Tag).Type = 135 Then
              Tipo = "D"
            Else
              Tipo = "V"
            End If
          End If
          
          If Operacao = "INSERT" Then
            Select Case Tipo
              Case "V"
                If TypeOf Controle Is CheckBox Then
                  SQl = SQl & IIf(Nnull(PPonto(Controle.Value), 0) = 0, 0, 1) & ","
                Else
                  If TypeOf Controle Is DataCombo Then
                    SQl = SQl & Nnull(PPonto(Controle.BoundText), 0) & ","
                  Else
                    If TypeOf Controle Is Label Then
                      SQl = SQl & Nnull(PPonto(Controle.Caption), 0) & ","
                    Else
                      SQl = SQl & Nnull(PPonto(Controle.text), 0) & ","
                    End If
                  End If
                End If
               
              Case "T"
                LocalCampo = Tira_AspaSimples(Nnull(Controle.text, 1))
                If TypeOf Controle Is Label Then LocalCampo = Tira_AspaSimples(Nnull(Controle.Caption, 1))
                If TypeOf Controle Is DataCombo Then LocalCampo = Controle.BoundText
                SQl = SQl & "'" & Nnull(LocalCampo, 1) & "',"
              
              Case "D"
                Controle.text = Nnull(Controle.text, 1)
                If Controle.text <> "" Then
                  SQl = SQl & "TO_DATE('" & Format$(Controle.text, "DD/MM/YYYY") & "', 'DD/MM/YYYY')" & ","
                Else
                  SQl = SQl & " NULL, "
                End If
            End Select
          End If
            
          If Operacao = "UPDATE" Then
            Select Case Tipo
              Case "T"
                LocalCampo = Tira_AspaSimples(Nnull(Controle.text, 1))
                If TypeOf Controle Is Label Then LocalCampo = Tira_AspaSimples(Nnull(Controle.Caption, 1))
                If TypeOf Controle Is DataCombo Then LocalCampo = Controle.BoundText
                SQl = SQl & Controle.Tag & " = " & "'" & Nnull(LocalCampo, 1) & "',"
              Case "V"
                If TypeOf Controle Is CheckBox Then
                  SQl = SQl & Controle.Tag & " = " & IIf(Nnull(Controle.Value, 0) = 0, 0, 1) & ","
                Else
                  If TypeOf Controle Is DataCombo Then
                    SQl = SQl & Controle.Tag & " = " & Nnull(PPonto(Controle.BoundText), 0) & ","
                  Else
                    If TypeOf Controle Is ComboBox Then
                      SQl = SQl & Controle.Tag & " = " & Nnull(PPonto(Controle.text), 0) & ","
                    Else
                      If TypeOf Controle Is Label Then
                        SQl = SQl & Controle.Tag & " = " & Nnull(PPonto(Controle.Caption), 0) & ","
                      Else
                        SQl = SQl & Controle.Tag & " = " & Nnull(PPonto(Controle.text), 0) & ","
                      End If
                    End If
                  End If
                End If
                
              Case "D"
                If Controle.text <> "" Then
                  SQl = SQl & Controle.Tag & " = " & "TO_DATE('" & Format(Controle.text, "DD/MM/YYYY") & "', 'DD/MM/YYYY')" & ","
                Else
                  SQl = SQl & Controle.Tag & " = NULL, "
                End If
            End Select
          End If
        End If
      End If
    End If
  Next i
  
  SQl = Trim$(SQl)
  If Right$(SQl, 1) = "," Then
    SQl = Left$(SQl, (Len(SQl) - 1))
  End If
End Sub

Public Sub AjustaImpostoTabela()
On Error GoTo trataErro:
Dim rsServicos As New ADODB.Recordset
Dim rstabelas As New ADODB.Recordset
Dim rsConfere As New ADODB.Recordset
Dim Ind As Integer
Dim Conta As Long

sSql = Empty
sSql = "SELECT AUTONUM FROM " & Banco_Sgipa & "TB_SERVICOS_IPA ORDER BY AUTONUM "
PRSet rsServicos, sSql

sSql = Empty
sSql = "SELECT AUTONUM FROM " & Banco_Sgipa & "TB_LISTAS_PRECOS WHERE AUTONUM > 708 ORDER BY AUTONUM "
PRSet rstabelas, sSql

While Not rstabelas.EOF
    rsServicos.MoveFirst
    While Not rsServicos.EOF
        Ind = 1
        For Ind = 1 To 3
            DoEvents
            sSql = Empty
            sSql = "SELECT * FROM TB_LP_SERVICOS_IMPOSTOS "
            sSql = sSql & " WHERE ID_IMPOSTO = " & Ind
            sSql = sSql & " AND ID_TABELA = " & Nnull(rstabelas!autonum, 0)
            sSql = sSql & " AND ID_SERVICO = " & Nnull(rsServicos!autonum, 0)
            PRSet rsConfere, sSql
            If rsConfere.EOF Then
                sSql = "INSERT INTO " & Banco_Sgipa & "TB_LP_SERVICOS_IMPOSTOS (ID_TABELA, "
                sSql = sSql & "ID_SERVICO, ID_IMPOSTO, FLAG_PADRAO, CALCULAR) VALUES ("
                sSql = sSql & Nnull(rstabelas!autonum, 0) & "," & Nnull(rsServicos!autonum, 0)
                sSql = sSql & "," & Ind & ",1,1) "
                Executa sSql
                Conta = Conta + 1
            Else
                If Nnull(rsConfere!Calcular, 0) = 0 Or Nnull(rsConfere!FLAG_PADRAO, 0) = 0 Then
                    sSql = Empty
                    sSql = "UPDATE " & Banco_Sgipa & "TB_LP_SERVICOS_IMPOSTOS SET CALCULAR = 1, FLAG_PADRAO = 1 "
                    sSql = sSql & " WHERE AUTONUM = " & Nnull(rsConfere!autonum, 0)
                    Executa sSql
                    Conta = Conta + 1
                End If
            End If
            
        Next Ind
        rsServicos.MoveNext
    Wend
    rstabelas.MoveNext
Wend

Exit Sub

trataErro:
    Err.Description = Empty

End Sub

Public Function obtemAliqImposto(AutonumImposto As Long, Tipo As Byte) As Currency
On Error GoTo trataErro:
'1 - ISS
'2 - COFINS
'3 - PIS
Dim rsAuxImp As New ADODB.Recordset

 obtemAliqImposto = 0

 sSql = Empty
 sSql = "SELECT TAXA "
 sSql = sSql & " FROM SGIPA.DBO.TB_CAD_IMPOSTOS "
 sSql = sSql & " WHERE AUTONUM = " & Nnull(AutonumImposto, 0)
 sSql = sSql & " AND ISNULL(TAXA ,0) > 0 "
 PRSet rsAuxImp, sSql
 If Not rsAuxImp.EOF Then
     obtemAliqImposto = Nnull(rsAuxImp!taxa, 0) / 100
 End If
 If Tipo = 2 Then obtemAliqImposto = 5 / 100
Exit Function

trataErro:
    Err.Description = Empty
    

End Function

