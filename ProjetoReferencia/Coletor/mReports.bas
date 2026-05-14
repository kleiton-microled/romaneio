Attribute VB_Name = "mReports"

Public Sub RunRelCntrsBloqueados()
End Sub

Public Sub AcertaArmadores()
Dim nCODE As Long
Dim oRS   As ADODB.Recordset
Dim oRT   As ADODB.Recordset
PRSet oRS, "SELECT * FROM " & BD_I & ".DTE_TB_ARMADORES ORDER BY CODE", 0
While Not oRS.EOF = True
    nCODE = Empty
    nCODE = Trim("9") & Format$(Trim(str(oRS.Fields("CODE"))), "00000000")
'    Gerencial.EnviaMensagem "Atualizando o Armador para os B/Ls com Armador = " & Trim(str(oRS.Fields("CODE"))) & "..."
    Executa "UPDATE " & BD_I & ".TB_BL SET ARMADOR = " & nCODE & " WHERE ARMADOR = " & Trim(str(oRS.Fields("CODE"))) & " AND DT_CADASTRO < TO_DATE('13/02/2005 00:00:00','DD/MM/YYYY HH24:MI:SS')"
    'ARMADOR DESCONSOLIDADOR
    PRSet oRT, "SELECT B.*, A.AUTONUM FROM TB_BL A, TB_AUDITORIA_PACOTE B WHERE A.AUDIT_DESCONSOLIDADO = B.AUTONUM AND B.FLAG_RECEITA = 2 AND B.EVENTO = 17 AND A.ARMADOR_CONSOLIDADOR = " & Trim(str(oRS.Fields("CODE"))) & " AND B.DT_FLAG < TO_DATE('13/02/2005 00:00:00','DD/MM/YYYY HH24:MI:SS')", 0
    While Not oRT.EOF = True
'        Gerencial.EnviaMensagem "Atualizando B/Ls..."
        Executa "UPDATE " & BD_I & ".TB_BL SET ARMADOR_CONSOLIDADOR = " & nCODE & " WHERE AUTONUM = " & oRT.Fields("AUTONUM") & " "
        oRT.MoveNext
    Wend
    oRT.Close
'    Gerencial.EnviaMensagem "Atualizando Armador..."
    Executa "UPDATE " & BD_I & ".DTE_TB_ARMADORES SET CODE = " & nCODE & " WHERE CODE = " & Trim(str(oRS.Fields("CODE"))) & " "
    oRS.MoveNext
Wend
oRS.Close
'Gerencial.EnviaMensagem
MsgBox "Processamento terminado."
nCODE = Empty
Set oRS = Nothing
Set oRT = Nothing
End Sub

'=================

Public Sub RunRelAtracacao()
Dim Qual        As String
Dim DataInicial As String
Dim DataFinal   As String
Dim Sql         As String
Dim sPatio      As String

Do
    Qual = InputBox("Selecione Tipo de Relatório:" & vbCrLf & "[ 0 ] - Por Aviso de Chegada" & vbCrLf & "[ 1 ] - Por Atracação Prevista" & vbCrLf & "[ 2 ] - Por Atracação Efetiva", , "0")
Loop Until Qual = "0" Or Qual = "1" Or Qual = "2" Or Qual = ""
If Qual = "" Then Exit Sub


While Not IsDate(DataInicial)
    DataInicial = InputBox("Informe a data inicial:", , Format(Now, "dd/mm/yyyy"))
    If DataInicial = "" Then
        Exit Sub
    ElseIf Not IsDate(DataInicial) = True Then
        MsgBox "Data inválida.", vbOKOnly + vbInformation
        Exit Sub
    End If
Wend
While Not IsDate(DataFinal)
    DataFinal = InputBox("Informe a data final:", , Format(DateAdd("d", 7, Now), "dd/mm/yyyy"))
    If DataFinal = "" Then
        Exit Sub
    ElseIf Not IsDate(DataFinal) = True Then
        MsgBox "Data inválida.", vbOKOnly + vbInformation
        Exit Sub
    End If
Wend
If CVDate(DataFinal) < CVDate(DataInicial) Then
    MsgBox "Período de datas inválidas", vbCritical
    Exit Sub
End If
MousePointer = 11
Select Case Qual
    Case Is = "0"
        RptName = App.Path & "\" & LOCAL_REPORT & "\avisochegada.RPT"
        Sql = "  SELECT DTE_TB_VIAGENS.HR_CHEGADA, DTE_TB_VIAGENS.VIAGEM, DTE_TB_VIAGENS.TIPO_ESCALA, DTE_TB_VIAGENS.DT_ATRACACAOP, DTE_TB_VIAGENS.HR_ATRACACAOP, DTE_TB_ARMAZENS.DESCR, DTE_TB_VIAGENS.DT_ATRACACAO, DTE_TB_VIAGENS.HR_ATRACACAO, DTE_TB_ARMAZENS_1.DESCR, DTE_TB_VIAGENS.DT_CHEGADA, DTE_TB_VIAGENS.NOMENAVIO"
        Sql = Sql & "  FROM   " & BD_I & ".DTE_TB_VIAGENS DTE_TB_VIAGENS left join " & BD_I & ".DTE_TB_NAVIOS DTE_TB_NAVIOS"
        Sql = Sql & " ON DTE_TB_VIAGENS.NAVIO = DTE_TB_NAVIOS.CODE "
        Sql = Sql & " INNER JOIN " & BD_I & ".DTE_TB_ARMAZENS DTE_TB_ARMAZENS  "
        Sql = Sql & " ON  DTE_TB_VIAGENS.ARM_ATRACACAOP = DTE_TB_ARMAZENS.CODE"
        Sql = Sql & " INNER JOIN " & BD_I & ".DTE_TB_ARMAZENS DTE_TB_ARMAZENS_1"
        Sql = Sql & " ON DTE_TB_VIAGENS.ARM_ATRACACAO = DTE_TB_ARMAZENS_1.CODE"
        Sql = Sql & " WHERE dte_tb_viagens.dt_atracacao >= to_date('" & Format(DataInicial, "dd/mm/yyyy") & "','DD/MM/YYYY') "
        Sql = Sql & " AND dte_tb_viagens.dt_atracacao <= to_date('" & Format(DataFinal, "dd/mm/yyyy") & "','DD/MM/YYYY')"
    Case Is = "1"
        RptName = App.Path & "\" & LOCAL_REPORT & "\aprevista.RPT"
        Sql = "     SELECT DTE_TB_VIAGENS.VIAGEM, DTE_TB_VIAGENS.TIPO_ESCALA, DTE_TB_VIAGENS.DT_ATRACACAOP, DTE_TB_VIAGENS.HR_ATRACACAOP, DTE_TB_ARMAZENS.DESCR, DTE_TB_VIAGENS.DT_ATRACACAO, DTE_TB_VIAGENS.HR_ATRACACAO, DTE_TB_ARMAZENS_1.DESCR, DTE_TB_VIAGENS.DT_CHEGADA, DTE_TB_VIAGENS.NOMENAVIO"
        Sql = Sql & "  FROM   " & BD_I & ".DTE_TB_VIAGENS DTE_TB_VIAGENS left join " & BD_I & ".DTE_TB_NAVIOS DTE_TB_NAVIOS"
        Sql = Sql & " ON DTE_TB_VIAGENS.NAVIO = DTE_TB_NAVIOS.CODE "
        Sql = Sql & " INNER JOIN " & BD_I & ".DTE_TB_ARMAZENS DTE_TB_ARMAZENS  "
        Sql = Sql & " ON  DTE_TB_VIAGENS.ARM_ATRACACAOP = DTE_TB_ARMAZENS.CODE"
        Sql = Sql & " INNER JOIN " & BD_I & ".DTE_TB_ARMAZENS DTE_TB_ARMAZENS_1"
        Sql = Sql & " ON DTE_TB_VIAGENS.ARM_ATRACACAO = DTE_TB_ARMAZENS_1.CODE"
        Sql = Sql & " and dte_tb_viagens.dt_chegada >= to_date('" & Format(DataInicial, "dd/mm/yyyy") & "','dd/mm/yyyy') and dte_tb_viagens.dt_chegada <= to_date('" & Format(DataFinal, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    Case Is = "2"
        RptName = App.Path & "\" & LOCAL_REPORT & "\aefetiva.RPT"
        RptPaperOrientation = crLandscape
        Sql = "         SELECT TB_CNTR_BL.DT_ENTRADA, TB_CNTR_BL.ID_CONTEINER, TB_CNTR_BL.TAMANHO, DTE_TB_TIPOS_CONTEINER.CODIGO, TB_BL.AUTONUM, DTE_TB_VIAGENS.NOMENAVIO, DTE_TB_PORTOS.DESCR, TB_AMR_CNTR_BL.BRUTO, TB_TIPOS_DOCUMENTOS.DESCR, TB_BL.NUM_DOCUMENTO, TB_CNTR_BL.PATIO, TB_BL.SEQ_FMA, TB_BL.ANO_FMA, TB_BL.TG, TB_BL.CANAL_SISCOMEX, TB_BL.FLAG_BLOQUEIO, TB_CNTR_BL.FLAG_BLOQUEIO, TB_CNTR_BL.YARD"
        Sql = Sql & " FROM   " & BD_I & ".TB_CNTR_BL TB_CNTR_BL INNER JOIN  " & BD_I & ".TB_AMR_CNTR_BL TB_AMR_CNTR_BL "
        Sql = Sql & " ON TB_CNTR_BL.AUTONUM=TB_AMR_CNTR_BL.CNTR "
        Sql = Sql & "INNER JOIN " & BD_I & ".DTE_TB_TIPOS_CONTEINER DTE_TB_TIPOS_CONTEINER"
        Sql = Sql & " ON TB_CNTR_BL.TIPO=DTE_TB_TIPOS_CONTEINER.CODE "
        Sql = Sql & " INNER JOIN " & BD_I & ".DTE_TB_VIAGENS DTE_TB_VIAGENS "
        Sql = Sql & " ON TB_CNTR_BL.VIAGEM=DTE_TB_VIAGENS.VIAGEM "
        Sql = Sql & " INNER JOIN " & BD_I & ".TB_BL TB_BL ON TB_AMR_CNTR_BL.BL=TB_BL.AUTONUM "
        Sql = Sql & " INNER JOIN " & BD_I & ".DTE_TB_PORTOS DTE_TB_PORTOS ON TB_BL.PORTO_ORIGEM=DTE_TB_PORTOS.CODE "
        Sql = Sql & " INNER JOIN " & BD_I & ".TB_TIPOS_DOCUMENTOS TB_TIPOS_DOCUMENTOS"
        Sql = Sql & " ON TB_BL.TIPO_DOCUMENTO=TB_TIPOS_DOCUMENTOS.CODE "
        Sql = Sql & "  WHERE  TB_BL." & Cod_Empresa_Patio
        Sql = Sql & " and tb_cntr_bl.dt_entrada >= to_date('" & Format(DataInicial, "dd/mm/yyyy") & "','dd/mm/yyyy') and tb_cntr_bl.dt_entrada <= to_date('" & Format(DataFinal, "dd/mm/yyyy") & "','dd/mm/yyyy')"
        Sql = Sql & "  ORDER BY TB_CNTR_BL.ID_CONTEINER"
End Select

Sql = CONVERTE_SQL(Sql)

RptFormulasName$(0) = "{@usuario}"
RptFormulas$(0) = "'" & Usuario_Sistema & "'"
RptFormulasName$(1) = "{@periodoini}"
RptFormulas$(1) = "'" & DataInicial & "'"
RptFormulasName$(2) = "{@periodofim}"
RptFormulas$(2) = "'" & DataFinal & "'"
SqlQuery = CONVERTE_SQL(Sql)
FrmReportViewer.Show 1: Set FrmReportViewer = Nothing
Screen.MousePointer = 0

Qual = Empty
DataInicial = Empty
DataFinal = Empty
Sql = Empty
sPatio = Empty
MousePointer = 0
End Sub
Public Sub RunRelCSTransp()
Dim Viagem As String
'Dim Rpt As Crystal.CrystalReport
Viagem = InputBox("Informe o número da viagem", , Viagens.Text1(0))
Viagem = Trim(UCase(Viagem))
If Viagem = Empty Then Exit Sub
MousePointer = 11
Sql = " SELECT TB_CARGA_SOLTA.GMVI, TB_CARGA_SOLTA.QUANTIDADE, DTE_TB_EMBALAGENS.DESCR, TB_CAD_PARCEIROS.RAZAO, TB_CARGA_SOLTA.IMO, TB_CARGA_SOLTA.MARCA, TB_CARGA_SOLTA.PESO_BRUTO, DTE_TB_VIAGENS.DT_ATRACACAOP, DTE_TB_VIAGENS.HR_ATRACACAOP, TB_CARGA_SOLTA.VIAGEM, DTE_TB_VIAGENS.NOMENAVIO, TB_CARGA_SOLTA.PATIO"
Sql = Sql & "  FROM   " & BD_I & ".TB_CARGA_SOLTA TB_CARGA_SOLTA, " & BD_I & ".TB_BL TB_BL, " & BD_I & ".DTE_TB_EMBALAGENS DTE_TB_EMBALAGENS, " & BD_I & ".DTE_TB_VIAGENS DTE_TB_VIAGENS, " & BD_I & ".TB_CAD_PARCEIROS TB_CAD_PARCEIROS"
Sql = Sql & "  Where (TB_CARGA_SOLTA.BL = TB_BL.Autonum) And (TB_CARGA_SOLTA.EMBALAGEM = DTE_TB_EMBALAGENS.Code) And (TB_BL.Viagem = DTE_TB_VIAGENS.Viagem) And (TB_BL.IMPORTADOR = TB_CAD_PARCEIROS.Autonum)"
Sql = Sql & "  ORDER BY TB_CARGA_SOLTA.VIAGEM"
RptName$ = App.Path & "\" & LOCAL_REPORT & "\CostadoCS.RPT"
RptPaperOrientation = crLandscape
RptFormulasName$(0) = "{@usuario}"
RptFormulas$(0) = "'" & Usuario_Sistema & "'"
RptFormulasName$(1) = "{@viagem}"
RptFormulas$(1) = "'" & Viagem & "'"
SqlQuery = CONVERTE_SQL(Sql)
FrmReportViewer.Show 1: Set FrmReportViewer = Nothing
Screen.MousePointer = 0

'Set Rpt = PreparaReport(Drv + "\CostadoCS.rpt")
'Rpt.Formulas(0) = "usuario = '" & Usuario_Sistema & "'"
'Rpt.Formulas(1) = "viagem = '" & Viagem & "'"
'Rpt.Action = 0
''limpareport Rpt
Viagem = Empty
MousePointer = 0
End Sub
Public Sub RunRelCntrSolicitado()
Dim Viagem    As String
Dim Cntr20()  As Long
Dim CNTR40()  As Long
'Dim Rpt       As Crystal.CrystalReport
Dim Tbh       As ADODB.Recordset
ReDim Cntr20(8)
ReDim CNTR40(8)
Viagem = InputBox("Informe o número da viagem", , Viagens.Text1(0))
Viagem = Trim(UCase(Viagem))
If Viagem = Empty Then Exit Sub
MousePointer = 11
For i = 1 To 8
    Cntr20(i) = 0: CNTR40(i) = 0
Next i
PRSet Tbh, "SELECT COUNT(A.AUTONUM) AS QTDE,B.CODIGO FROM " & BD_I & ".TB_CNTR_BL A," & BD_I & ".DTE_TB_TIPOS_CONTEINER B WHERE A.TIPO=B.CODE AND A.VIAGEM='" & Viagem & "' AND A.TAMANHO=20 GROUP BY B.CODIGO", 0
If Not Tbh.EOF Then
    While Not Tbh.EOF
        Select Case Tbh!Codigo
            Case "DC":                Cntr20(1) = Nnull(Tbh!qtde, 0)
            Case "FR":                Cntr20(2) = Nnull(Tbh!qtde, 0)
            Case "TK":                Cntr20(3) = Nnull(Tbh!qtde, 0)
            Case "PL":                Cntr20(4) = Nnull(Tbh!qtde, 0)
            Case "OT":                Cntr20(5) = Nnull(Tbh!qtde, 0)
            Case "RE":                Cntr20(6) = Nnull(Tbh!qtde, 0)
            Case "HC":                Cntr20(7) = Nnull(Tbh!qtde, 0)
            Case "HR":                Cntr20(8) = Nnull(Tbh!qtde, 0)
        End Select
        Tbh.MoveNext
    Wend
End If
Tbh.Close
PRSet Tbh, "SELECT COUNT(A.AUTONUM) AS QTDE,B.CODIGO FROM " & BD_I & ".TB_CNTR_BL A," & BD_I & ".DTE_TB_TIPOS_CONTEINER B WHERE A.TIPO=B.CODE AND A.VIAGEM='" & Viagem & "'  AND A.TAMANHO=40 GROUP BY B.CODIGO"
If Not Tbh.EOF Then
    While Not Tbh.EOF
        Select Case Tbh!Codigo
            Case "DC":                CNTR40(1) = Nnull(Tbh!qtde, 0)
            Case "FR":                CNTR40(2) = Nnull(Tbh!qtde, 0)
            Case "TK":                CNTR40(3) = Nnull(Tbh!qtde, 0)
            Case "PL":                CNTR40(4) = Nnull(Tbh!qtde, 0)
            Case "OT":                CNTR40(5) = Nnull(Tbh!qtde, 0)
            Case "RE":                CNTR40(6) = Nnull(Tbh!qtde, 0)
            Case "HC":                CNTR40(7) = Nnull(Tbh!qtde, 0)
            Case "HR":                CNTR40(8) = Nnull(Tbh!qtde, 0)
        End Select
        Tbh.MoveNext
    Wend
End If
Tbh.Close
Executa "DELETE from " & BD_I & ".TB_TEMP_CNTR_SOL"
Executa "INSERT INTO " & BD_I & ".TB_TEMP_CNTR_SOL SELECT AUTONUM FROM " & BD_I & ".TB_CNTR_BL WHERE VIAGEM='" & Viagem & "'"
'Set Rpt = PreparaReport(App.Path & "\" & "CntrSolicitadosNew.rpt", , "{TB_BL.FLAG_ATIVO}=1")
'Rpt.Formulas(0) = "usuario = '" & Usuario_Sistema & "'"
'Rpt.Formulas(1) = "DC20 = " & Cntr20(1)
'Rpt.Formulas(2) = "FR20 = " & Cntr20(2)
'Rpt.Formulas(3) = "TK20 = " & Cntr20(3)
'Rpt.Formulas(4) = "PL20 = " & Cntr20(4)
'Rpt.Formulas(5) = "OT20 = " & Cntr20(5)
'Rpt.Formulas(6) = "RE20 = " & Cntr20(6)
'Rpt.Formulas(7) = "HC20 = " & Cntr20(7)
'Rpt.Formulas(8) = "HR20 = " & Cntr20(8)
'Rpt.Formulas(9) = "DC40 = " & CNTR40(1)
'Rpt.Formulas(10) = "FR40 = " & CNTR40(2)
'Rpt.Formulas(11) = "TK40 = " & CNTR40(3)
'Rpt.Formulas(12) = "PL40 = " & CNTR40(4)
'Rpt.Formulas(13) = "OT40 = " & CNTR40(5)
'Rpt.Formulas(14) = "RE40 = " & CNTR40(6)
'Rpt.Formulas(15) = "HC40 = " & CNTR40(7)
'Rpt.Formulas(16) = "HR40 = " & CNTR40(8)
'Rpt.Destination = 0
'Rpt.Action = 1
''limpareport Rpt
Viagem = Empty
Erase Cntr20()
Erase CNTR40()
'set rpt = nothing
Set Tbh = Nothing
MousePointer = 0
End Sub
Public Sub RunRelProtGMCINew()
Dim Viagem As String
Viagem = InputBox("Informe o número da viagem", , Viagens.Text1(0))
If Viagem = "" Then Exit Sub
Screen.MousePointer = 11
Executa "DELETE FROM " & BD_I & ".TB_TEMP_CNTR_PROT"
Executa "INSERT INTO " & BD_I & ".TB_TEMP_CNTR_PROT SELECT AUTONUM FROM " & BD_I & ".TB_CNTR_BL WHERE GMCI IS NOT NULL AND  VIAGEM='" & Viagem & "'"
'Set Rpt = Principal.CrystalReport1
'Rpt.ReportFileName = App.Path & "\" & "PROTOCOLO_GMCINew.rpt"
'Rpt.Formulas(0) = "usuario = '" & Usuario_Sistema & "'"
'Rpt.Formulas(1) = "viagem = '" & Viagem & "'"
'Rpt.Action = 0
'Rpt.Formulas(0) = ""
'Rpt.Formulas(1) = ""
'Screen.MousePointer = 0
End Sub
Public Sub RunRelTRAS()
Dim Viagem As String
Viagem = InputBox("Informe o número da viagem", , Viagens.Text1(0))
If Viagem = "" Then Exit Sub
Set tb_h = New ADODB.Recordset
tb_h.Open "SELECT MIN(A.autonum) AS AUTONUM,SUM(B.BRUTO) AS BRUTO,MIN(A.TARA) AS TARA FROM " & BD_I & ".TB_CNTR_BL A," & BD_I & ".TB_AMR_CNTR_BL B WHERE B.CNTR=A.AUTONUM AND A.VIAGEM='" & Viagem & "' GROUP BY A.AUTONUM", db, 3, 3
If Not tb_h.EOF Then
    While Not tb_h.EOF
        Executa "UPDATE " & BD_I & ".TB_CNTR_BL SET TEMP_BRUTO=" & str(tb_h!Bruto + tb_h!Tara) & " WHERE AUTONUM=" & tb_h!AUTONUM
        tb_h.MoveNext
    Wend
End If
tb_h.Close
Screen.MousePointer = 11
Executa "DELETE FROM " & BD_I & ".TB_TEMP_CNTR_ENTREGA"
Executa "INSERT INTO " & BD_I & ".TB_TEMP_CNTR_ENTREGA SELECT AUTONUM FROM " & BD_I & ".TB_CNTR_BL WHERE DT IS NOT NULL AND VIAGEM='" & Viagem & "'"
'Set Rpt = Principal.CrystalReport1
'Rpt.ReportFileName = App.Path & "\" & "Entrega_GMCINEW.rpt"
'Rpt.Formulas(0) = "usuario = '" & Usuario_Sistema & "'"
'Rpt.Formulas(1) = "viagem = '" & Viagem & "'"
'Rpt.Action = 0
'Rpt.Formulas(0) = ""
'Rpt.Formulas(1) = ""
'Rpt.SelectionFormula = ""
'Screen.MousePointer = 0
End Sub
Public Sub RunRelP4()
Dim sSql As String
Dim dDATAS() As Date
ReDim dDATAS(2)
Load frmGetPeriodo
frmGetPeriodo.DataFim = 0
frmGetPeriodo.DataInicio = 0
frmGetPeriodo.Show 1
dDATAS(0) = frmGetPeriodo.DataInicio
dDATAS(1) = frmGetPeriodo.DataFim

frmGetPeriodo.DataFim = 0
frmGetPeriodo.DataInicio = 0
Unload frmGetPeriodo
Set frmGetPeriodo = Nothing
Sql = "insert into tb_temp_maritimo ("
Sql = Sql & " SELECT "
Sql = Sql & "     VW_MARITIMO_INDICADOR.AUTONUM, VW_MARITIMO_INDICADOR.NUM_DOCUMENTO, VW_MARITIMO_INDICADOR.SEQ_GR, VW_MARITIMO_INDICADOR.DT_IMPRESSAO, VW_MARITIMO_INDICADOR.TABELA_GR, VW_MARITIMO_INDICADOR.IMPORTADOR, VW_MARITIMO_INDICADOR.INDICADOR, VW_MARITIMO_INDICADOR.VIAGEM, VW_MARITIMO_INDICADOR.VALOR, VW_MARITIMO_INDICADOR.NAVIO, VW_MARITIMO_INDICADOR.ARMAZEM, VW_MARITIMO_INDICADOR.SERVICO, TB_PATIOS.DESCR_RESUMIDO, VW_MARITIMO_INDICADOR.REGIME,'" & Usuario_Sistema & "',VW_MARITIMO_INDICADOR.AUTONUM_IND,VW_MARITIMO_INDICADOR.datafaturado"
Sql = Sql & "  FROM "
Sql = Sql & "     " & BD_O & ".TB_PATIOS TB_PATIOS, " & BD_I & ".VW_MARITIMO_INDICADOR VW_MARITIMO_INDICADOR "
Sql = Sql & " WHERE VW_MARITIMO_INDICADOR.PATIO = TB_PATIOS.DESCR and VW_MARITIMO_INDICADOR." & Cod_Empresa_Patio & " And "
Sql = Sql & "   VW_MARITIMO_INDICADOR.DT_IMPRESSAO >= TO_DATE('" & Format$(dDATAS(0), "DD/MM/YYYY") & "','DD/MM/YYYY') "
Sql = Sql & "   AND VW_MARITIMO_INDICADOR.DT_IMPRESSAO <= TO_DATE('" & Format$(dDATAS(1), "DD/MM/YYYY") & "','DD/MM/YYYY') ) "
'Executa Sql
Sql = ""
Sql = Sql & "SELECT VW_MARITIMO_INDICADOR.INDICADOR, VW_MARITIMO_INDICADOR.IMPORTADOR, VW_MARITIMO_INDICADOR.AUTONUM, VW_MARITIMO_INDICADOR.VIAGEM, VW_MARITIMO_INDICADOR.NAVIO, VW_MARITIMO_INDICADOR.ARMAZEM, VW_MARITIMO_INDICADOR.NUM_DOCUMENTO, VW_MARITIMO_INDICADOR.SERVICO, VW_MARITIMO_INDICADOR.VALOR, VW_MARITIMO_INDICADOR.DT_IMPRESSAO, VW_MARITIMO_INDICADOR.TABELA_GR, VW_MARITIMO_INDICADOR.SEQ_GR, VW_MARITIMO_INDICADOR.REGIME, VW_MARITIMO_INDICADOR.AUTONUM_IND, VW_MARITIMO_INDICADOR.PATIO"
Sql = Sql & " FROM   SGIPA.dbo.VW_MARITIMO_INDICADOR VW_MARITIMO_INDICADOR"
Sql = Sql & " WHERE  "
Sql = Sql & "   VW_MARITIMO_INDICADOR.DT_IMPRESSAO >= TO_DATE('" & Format$(dDATAS(0), "DD/MM/YYYY") & "','DD/MM/YYYY') "
Sql = Sql & "   AND VW_MARITIMO_INDICADOR.DT_IMPRESSAO <= TO_DATE('" & Format$(dDATAS(1), "DD/MM/YYYY") & "','DD/MM/YYYY') ) "
Sql = Sql & " ORDER BY VW_MARITIMO_INDICADOR.PATIO, VW_MARITIMO_INDICADOR.INDICADOR, VW_MARITIMO_INDICADOR.AUTONUM"
RptFormulasName(0) = "{@INI}"
RptFormulasName(1) = "{@FIM}"
RptFormulasName(2) = "{@USUARIO}"
RptFormulasName(3) = "{@cab}"
RptFormulas(0) = "'" & Format(dDATAS(0), "DD/MM/YYYY") & "'"
RptFormulas(1) = "'" & Format(dDATAS(1), "DD/MM/YYYY") & "'"
RptFormulas(2) = "'" & Usuario_Sistema & "'"
RptFormulas(3) = "'" & Razao_Empresa & "'"
RptPaperOrientation = crLandscape
RptName = App.Path & "\" & LOCAL_REPORT & "\MARITIND.RPT"
FrmReportViewer.Show 1: Set FrmReportViewer = Nothing
Sql = "DELETE FROM TB_TEMP_MARITIMO WHERE   usuario='" & Usuario_Sistema & "'"
Executa Sql
Erase dDATAS

End Sub
Public Sub RunRelGRTipoServico()
Dim DataIni As String
Dim DataFim As String
Do
    DataIni = InputBox("Qual a data inicial de pagamento", , Format(Now, "dd/mm/yyyy"))
    If DataIni = "" Then Exit Sub
Loop While Not IsDate(DataIni)
DataIni = Format(CVDate(DataIni), "dd/mm/yyyy")
Do
    DataFim = InputBox("Qual a data final de pagamento", , Format(Now, "dd/mm/yyyy"))
    If DataFim = "" Then Exit Sub
Loop While Not IsDate(DataFim)
DataFim = Format(CVDate(DataFim), "dd/mm/yyyy")
If CVDate(DataFim) < CVDate(DataIni) Then
    MsgBox "Data inicial maior que a data final", vbCritical: Exit Sub
End If
Screen.MousePointer = 11
'CrystalReport1.ReportFileName = App.Path & "\" & "grservicos.rpt"
'CrystalReport1.Formulas(0) = "usuario = '" & Usuario_Sistema & "'"
'CrystalReport1.Formulas(1) = "periodoini = '" & DataIni & "'"
'CrystalReport1.Formulas(2) = "periodofim = '" & DataFim & "'"
'CrystalReport1.Action = 0
'For I = 0 To 2
'    CrystalReport1.Formulas(I) = ""
'Next I
'CrystalReport1.SelectionFormula = ""
'Screen.MousePointer = 0
End Sub

Public Sub RunRelGRCI()
Dim Viagem As String
Dim GMCI As String
Dim DT As String
Viagem = InputBox("Informe o número da viagem", , Viagens.Text1(0))
If Viagem = "" Then Exit Sub
Screen.MousePointer = 11
GMCI = InputBox("Informe o número da GMCI,(0) Para todas", , 0)
If GMCI = "" Then Exit Sub
If GMCI = "0" Then GMCI = ""
DT = InputBox("Informe o número da DT,(0) Para todas", , 0)
If DT = "" Then Exit Sub
If DT = "0" Then DT = ""
Call Rel_GRCI_GMCI(DT, GMCI, Viagem, 2)
Screen.MousePointer = 0
End Sub
Public Sub RunRelGMCI()
Dim Viagem As String
Dim GMCI As String
Dim DT As String
Viagem = InputBox("Informe o número da viagem", , Viagens.Text1(0))
If Viagem = "" Then Exit Sub
MousePointer = 11
GMCI = InputBox("Informe o número da GMCI,(0) Para todas", , 0)
If GMCI = "" Then Exit Sub
If GMCI = "0" Then GMCI = ""
DT = InputBox("Informe o número da DT,(0) Para todas", , 0)
If DT = "" Then Exit Sub
If DT = "0" Then DT = ""
Call Rel_GRCI_GMCI(DT, GMCI, Viagem, 1)
Screen.MousePointer = 0
End Sub
Public Sub RunRelEntregaCS()
Dim Viagem As String
Viagem = InputBox("Informe o número da viagem", , Viagens.Text1(0))
If Viagem = "" Then Exit Sub
Screen.MousePointer = 11
'Set Rpt = Principal.CrystalReport1
'Rpt.ReportFileName = Drv + "\entrega_GMVI.rpt"
'Rpt.Formulas(0) = "usuario = '" & Usuario_Sistema & "'"
'Rpt.Formulas(1) = "viagem = '" & Viagem & "'"
'Rpt.Action = 0
'Rpt.Formulas(0) = ""
'Rpt.Formulas(1) = ""
'Rpt.SelectionFormula = ""
Screen.MousePointer = 0
End Sub
Public Sub RunRelCSOL()
Dim Viagem As String
Viagem = InputBox("Informe o número da viagem", , Viagens.Text1(0))
If Viagem = "" Then Exit Sub
Screen.MousePointer = 11
'Set Rpt = Principal.CrystalReport1
'Rpt.ReportFileName = Drv + "\CSSolicitadA.rpt"
'Rpt.Formulas(0) = "usuario = '" & Usuario_Sistema & "'"
'Rpt.SelectionFormula = "{TB_CARGA_SOLTA.VIAGEM} = '" & Viagem & "'"
'Rpt.Action = 0
'For I = 0 To 0
'    Rpt.Formulas(I) = ""
'Next I
'Rpt.SelectionFormula = ""
Screen.MousePointer = 0
End Sub
Public Sub RunRelCs()
Dim Viagem As String
Viagem = InputBox("Informe o número da viagem", , Viagens.Text1(0))
If Viagem = "" Then Exit Sub
'Set Rpt = Principal.CrystalReport1
'Rpt.ReportFileName = Drv + "\CSSolicitada.rpt"
'Rpt.Formulas(0) = "usuario = '" & Usuario_Sistema & "'"
'Rpt.SelectionFormula = "{TB_CARGA_SOLTA.VIAGEM} = '" & Viagem & "'"
'Rpt.Action = 0
'For I = 0 To 0
'    Rpt.Formulas(I) = ""
'Next I
End Sub
Public Sub RunRelCntrTransp()
Dim Viagem As String
Viagem = InputBox("Informe o número da viagem", , Viagens.Text1(0))
If Viagem = "" Then Exit Sub
Set tb_h = New ADODB.Recordset
tb_h.Open "SELECT MIN(A.autonum) AS AUTONUM,SUM(B.BRUTO) AS BRUTO FROM " & BD_I & ".TB_CNTR_BL A," & BD_I & ".TB_AMR_CNTR_BL B WHERE B.CNTR=A.AUTONUM AND A.VIAGEM='" & Viagem & "' GROUP BY A.AUTONUM", db, 3, 3
If Not tb_h.EOF Then
    While Not tb_h.EOF
        Set TBTARA = New ADODB.Recordset
        TBTARA.Open "Select tara from " & BD_I & ".TB_CNTR_BL where autonum=" & tb_h!AUTONUM, db, 3, 3
        If Not TBTARA.EOF Then
            Executa "UPDATE " & BD_I & ".TB_CNTR_BL SET TEMP_BRUTO=" & str(tb_h!Bruto + TBTARA!Tara) & " WHERE AUTONUM=" & tb_h!AUTONUM
        End If
        TBTARA.Close
        tb_h.MoveNext
    Wend
End If
tb_h.Close
Screen.MousePointer = 11
Dim Cntr20(8) As Long
Dim CNTR40(8) As Long
For i = 1 To 8
    Cntr20(i) = 0: CNTR40(i) = 0
Next i
Set Tbh = New ADODB.Recordset
Tbh.Open "SELECT COUNT(A.AUTONUM) AS QTDE,B.CODIGO FROM " & BD_I & ".TB_CNTR_BL A," & BD_I & ".DTE_TB_TIPOS_CONTEINER B WHERE A.TIPO=B.CODE AND A.VIAGEM='" & Viagem & "' AND A.DT IS NOT NULL AND A.TAMANHO=20 GROUP BY B.CODIGO", db, 3, 3
If Not Tbh.EOF Then
    While Not Tbh.EOF
        Select Case Tbh!Codigo
            Case "DC"
                Cntr20(1) = Nnull(Tbh!qtde, 0)
            Case "FR"
                Cntr20(2) = Nnull(Tbh!qtde, 0)
            Case "TK"
                Cntr20(3) = Nnull(Tbh!qtde, 0)
            Case "PL"
                Cntr20(4) = Nnull(Tbh!qtde, 0)
            Case "OT"
                Cntr20(5) = Nnull(Tbh!qtde, 0)
            Case "RE"
                Cntr20(6) = Nnull(Tbh!qtde, 0)
            Case "HC"
                Cntr20(7) = Nnull(Tbh!qtde, 0)
            Case "HR"
                Cntr20(8) = Nnull(Tbh!qtde, 0)
        End Select
        Tbh.MoveNext
    Wend
End If
Tbh.Close
Set Tbh = New ADODB.Recordset
Tbh.Open "SELECT COUNT(A.AUTONUM) AS QTDE,B.CODIGO FROM " & BD_I & ".TB_CNTR_BL A," & BD_I & ".DTE_TB_TIPOS_CONTEINER B WHERE A.TIPO=B.CODE AND A.VIAGEM='" & Viagem & "' AND A.DT IS NOT NULL AND A.TAMANHO=40 GROUP BY B.CODIGO", db, 3, 3
If Not Tbh.EOF Then
    While Not Tbh.EOF
        Select Case Tbh!Codigo
            Case "DC"
                CNTR40(1) = Nnull(Tbh!qtde, 0)
            Case "FR"
                CNTR40(2) = Nnull(Tbh!qtde, 0)
            Case "TK"
                CNTR40(3) = Nnull(Tbh!qtde, 0)
            Case "PL"
                CNTR40(4) = Nnull(Tbh!qtde, 0)
            Case "OT"
                CNTR40(5) = Nnull(Tbh!qtde, 0)
            Case "RE"
                CNTR40(6) = Nnull(Tbh!qtde, 0)
            Case "HC"
                CNTR40(7) = Nnull(Tbh!qtde, 0)
            Case "HR"
                CNTR40(8) = Nnull(Tbh!qtde, 0)
        End Select
        Tbh.MoveNext
    Wend
End If
Tbh.Close
'Set Rpt = Principal.CrystalReport1
'Rpt.ReportFileName = App.Path & "\" & "Costado.rpt"
'Rpt.Formulas(0) = "usuario = '" & Usuario_Sistema & "'"
'Rpt.Formulas(1) = "DC20 = " & Cntr20(1)
'Rpt.Formulas(2) = "FR20 = " & Cntr20(2)
'Rpt.Formulas(3) = "TK20 = " & Cntr20(3)
'Rpt.Formulas(4) = "PL20 = " & Cntr20(4)
'Rpt.Formulas(5) = "OT20 = " & Cntr20(5)
'Rpt.Formulas(6) = "RE20 = " & Cntr20(6)
'Rpt.Formulas(7) = "HC20 = " & Cntr20(7)
'Rpt.Formulas(8) = "HR20 = " & Cntr20(8)
'Rpt.Formulas(9) = "DC40 = " & CNTR40(1)
'Rpt.Formulas(10) = "FR40 = " & CNTR40(2)
'Rpt.Formulas(11) = "TK40 = " & CNTR40(3)
'Rpt.Formulas(12) = "PL40 = " & CNTR40(4)
'Rpt.Formulas(13) = "OT40 = " & CNTR40(5)
'Rpt.Formulas(14) = "RE40 = " & CNTR40(6)
'Rpt.Formulas(15) = "HC40 = " & CNTR40(7)
'Rpt.Formulas(16) = "HR40 = " & CNTR40(8)
'Rpt.ReplaceSelectionFormula "{TB_CNTR_BL.VIAGEM} = '" & Viagem & "' and not isnull({TB_CNTR_BL.DT}) and {TB_BL.FLAG_ATIVO}=1"
'Rpt.Action = 0
'For I = 0 To 16
'    Rpt.Formulas(I) = ""
'Next I
'Rpt.ReplaceSelectionFormula Empty
'Screen.MousePointer = 0
End Sub
Public Sub RunRelCntrImp()
Dim Viagem As String
Viagem = InputBox("Informe o número da viagem", , Viagens.Text1(0))
If Viagem = "" Then Exit Sub
Set tb_h = New ADODB.Recordset
tb_h.Open "SELECT MIN(A.autonum) AS AUTONUM,MIN(A.TARA) AS TARA,SUM(B.BRUTO) AS BRUTO FROM " & BD_I & ".TB_CNTR_BL A," & BD_I & ".TB_AMR_CNTR_BL B WHERE B.CNTR=A.AUTONUM AND A.VIAGEM='" & Viagem & "' GROUP BY A.AUTONUM", db, 3, 3
If Not tb_h.EOF Then
    While Not tb_h.EOF
        'Set tbtara = New ADODB.Recordset
        'tbtara.Open "Select tara from " & bd_i & ".TB_CNTR_BL where autonum=" & tb_h!autonum, Db, 3, 3
        'If Not tbtara.Eof    Then
        Executa "UPDATE " & BD_I & ".TB_CNTR_BL SET TEMP_BRUTO=" & str(tb_h!Bruto + tb_h!Tara) & " WHERE AUTONUM=" & tb_h!AUTONUM
        'End If
        'tbtara.Close
        tb_h.MoveNext
    Wend
End If
tb_h.Close
Screen.MousePointer = 11
Dim Cntr20(8) As Long
Dim CNTR40(8) As Long
For i = 1 To 8
    Cntr20(i) = 0: CNTR40(i) = 0
Next i
Set Tbh = New ADODB.Recordset
Tbh.Open "SELECT COUNT(A.AUTONUM) AS QTDE,B.CODIGO FROM " & BD_I & ".TB_CNTR_BL A," & BD_I & ".DTE_TB_TIPOS_CONTEINER B WHERE A.TIPO=B.CODE AND A.VIAGEM='" & Viagem & "' AND A.DT IS NOT NULL AND A.TAMANHO=20 GROUP BY B.CODIGO", db, 3, 3
If Not Tbh.EOF Then
    While Not Tbh.EOF
        Select Case Tbh!Codigo
            Case "DC"
                Cntr20(1) = Nnull(Tbh!qtde, 0)
            Case "FR"
                Cntr20(2) = Nnull(Tbh!qtde, 0)
            Case "TK"
                Cntr20(3) = Nnull(Tbh!qtde, 0)
            Case "PL"
                Cntr20(4) = Nnull(Tbh!qtde, 0)
            Case "OT"
                Cntr20(5) = Nnull(Tbh!qtde, 0)
            Case "RE"
                Cntr20(6) = Nnull(Tbh!qtde, 0)
            Case "HC"
                Cntr20(7) = Nnull(Tbh!qtde, 0)
            Case "HR"
                Cntr20(8) = Nnull(Tbh!qtde, 0)
        End Select
        Tbh.MoveNext
    Wend
End If
Tbh.Close
Set Tbh = New ADODB.Recordset
Tbh.Open "SELECT COUNT(A.AUTONUM) AS QTDE,B.CODIGO FROM " & BD_I & ".TB_CNTR_BL A," & BD_I & ".DTE_TB_TIPOS_CONTEINER B WHERE A.TIPO=B.CODE AND A.VIAGEM='" & Viagem & "' AND A.DT IS NOT NULL AND A.TAMANHO=40 GROUP BY B.CODIGO", db, 3, 3
If Not Tbh.EOF Then
    While Not Tbh.EOF
        Select Case Tbh!Codigo
            Case "DC"
                CNTR40(1) = Nnull(Tbh!qtde, 0)
            Case "FR"
                CNTR40(2) = Nnull(Tbh!qtde, 0)
            Case "TK"
                CNTR40(3) = Nnull(Tbh!qtde, 0)
            Case "PL"
                CNTR40(4) = Nnull(Tbh!qtde, 0)
            Case "OT"
                CNTR40(5) = Nnull(Tbh!qtde, 0)
            Case "RE"
                CNTR40(6) = Nnull(Tbh!qtde, 0)
            Case "HC"
                CNTR40(7) = Nnull(Tbh!qtde, 0)
            Case "HR"
                CNTR40(8) = Nnull(Tbh!qtde, 0)
        End Select
        Tbh.MoveNext
    Wend
End If
Tbh.Close
'Set Rpt = Principal.CrystalReport1
'Rpt.ReportFileName = App.Path & "\" & "Costado_Importador.rpt"
'Rpt.Formulas(0) = "usuario = '" & Usuario_Sistema & "'"
'Rpt.Formulas(1) = "DC20 = " & Cntr20(1)
'Rpt.Formulas(2) = "FR20 = " & Cntr20(2)
'Rpt.Formulas(3) = "TK20 = " & Cntr20(3)
'Rpt.Formulas(4) = "PL20 = " & Cntr20(4)
'Rpt.Formulas(5) = "OT20 = " & Cntr20(5)
'Rpt.Formulas(6) = "RE20 = " & Cntr20(6)
'Rpt.Formulas(7) = "HC20 = " & Cntr20(7)
'Rpt.Formulas(8) = "HR20 = " & Cntr20(8)
'Rpt.Formulas(9) = "DC40 = " & CNTR40(1)
'Rpt.Formulas(10) = "FR40 = " & CNTR40(2)
'Rpt.Formulas(11) = "TK40 = " & CNTR40(3)
'Rpt.Formulas(12) = "PL40 = " & CNTR40(4)
'Rpt.Formulas(13) = "OT40 = " & CNTR40(5)
'Rpt.Formulas(14) = "RE40 = " & CNTR40(6)
'Rpt.Formulas(15) = "HC40 = " & CNTR40(7)
'Rpt.Formulas(16) = "HR40 = " & CNTR40(8)
'Rpt.SelectionFormula = "{TB_CNTR_BL.VIAGEM} = '" & Viagem & "' and not isnull({TB_CNTR_BL.DT}) and {TB_BL.FLAG_ATIVO}=1"
'Rpt.Action = 0
'For I = 0 To 16
'    Rpt.Formulas(I) = ""
'Next I
'Rpt.SelectionFormula = ""
'Screen.MousePointer = 0
End Sub
Public Sub RunRelAgendaDesova()
Dim DataInicial As String
Dim DataFinal As String
Do
    DataInicial = InputBox("Informe a data inicial", , Format(Now, "dd/mm/yyyy"))
    If DataInicial = "" Then Exit Sub
Loop While Not IsDate(DataInicial)
Do
    DataFinal = InputBox("Informe a data final", , Format(DateAdd("d", 7, Now), "dd/mm/yyyy"))
    If DataFinal = "" Then Exit Sub
Loop While Not IsDate(DataFinal)
If CVDate(DataFinal) < CVDate(DataInicial) Then
    MsgBox "Período de datas inválidas", vbCritical: Exit Sub
End If
Screen.MousePointer = 11
'Set Rpt = Principal.CrystalReport1
'Rpt.ReportFileName = Drv + "\Abertura.rpt"
'Rpt.Formulas(0) = "usuario = '" & Usuario_Sistema & "'"
'Rpt.Formulas(1) = "DTI = '" & DataInicial & "'"
'Rpt.Formulas(2) = "DTF = '" & DataFinal & "'"
'Rpt.SelectionFormula = "{TB_AGENDAMENTO_DESOVA.DT_PREVISTA} >= Date (" & Year(DataInicial) & ", " & Month(DataInicial) & " , " & Day(DataInicial) & ") and {TB_AGENDAMENTO_DESOVA.DT_PREVISTA} <= Date (" & Year(DataFinal) & ", " & Month(DataFinal) & " , " & Day(DataFinal) & ")"
'Rpt.Action = 0
'For I = 0 To 2
'    Rpt.Formulas(I) = ""
'Next I
'Rpt.SelectionFormula = ""
'Screen.MousePointer = 0
End Sub
Public Sub RunRelRegistroSaida()
Dim DataInicial  As String
Dim DataFinal    As String
Dim Sql          As String
Dim CT20         As Integer
Dim CT40         As Integer
Dim VOL          As Integer
Dim SAIDA20      As Integer
Dim SAIDA40      As Integer
Dim SAIDAVOL     As Integer
Dim SAIDAReal20  As Integer
Dim SAIDAReal40  As Integer
Dim SAIDARealVOL As Integer
Dim Transporte   As String
Dim DtSaida      As String
Dim DtSaidaCs    As String
Dim sPatio       As String
Dim sDATATEMP    As String
Dim Tbh          As ADODB.Recordset
Dim oRS()        As ADODB.Recordset
ReDim oRS(5)
While Not IsDate(DataInicial) = True
    DataInicial = InputBox("Informe a data inicial", , Format(Now, "dd/mm/yyyy"))
    If DataInicial = "" Then Exit Sub
Wend
While Not IsDate(DataFinal) = True
    DataFinal = InputBox("Informe a data final", , Format(Now, "dd/mm/yyyy"))
    If DataFinal = "" Then Exit Sub
Wend
'While sPatio <> "0" And sPatio <> "1" And sPatio <> "2"
'    sPatio = InputBox("Informe o Pátio", , "0")
'    If sPatio = "" Then Exit Sub
'Wend
If Cod_Empresa = 2 Then
  While sPatio <> "5"
   sPatio = InputBox("Informe o Pátio", , "5")
   If sPatio = "" Then Exit Sub
  Wend
Else
  While sPatio <> "0" And sPatio <> "1" And sPatio <> "8" And sPatio <> "9"
      sPatio = InputBox("Informe o Pátio", , "1")
      If sPatio = "" Then Exit Sub
  Wend
End If
If CVDate(DataFinal) < CVDate(DataInicial) Then
    sDATATEMP = DataFinal
    DataFinal = DataInicial
    DataInicial = sDATATEMP
End If
Screen.MousePointer = vbHourglass
Executa "DELETE FROM " & BD_I & ".TB_TEMP_RELREGISTRO"
Sql = Empty
Sql = Sql & "SELECT "
Sql = Sql & "E.AUTONUM, "
Sql = Sql & "TO_CHAR(E.DT_REGISTRO,'dd/mm/yyyy hh24:mi:ss') as Data_Reg, "
Sql = Sql & "E.NUM_DOCUMENTO, "
Sql = Sql & "G.RAZAO AS IMPORTADOR, "
Sql = Sql & "TO_CHAR(H.Dt_Envio,'dd/mm/yyyy hh24:mi:ss') as dt_envio "

'Sql = Sql & " from "
'Sql = Sql & "" & BD_I & ".TB_BL E, "
'Sql = Sql & "" & BD_I & ".TB_CAD_PARCEIROS G, "
'Sql = Sql & "" & BD_I & ".TB_AUDITORIA_PACOTE H "
'Sql = Sql & " where "
'Sql = Sql & "E.IMPORTADOR = G.AUTONUM(+) "
'Sql = Sql & "AND E.AUDIT_DESEMBARACO = H.AUTONUM(+) "
'Sql = Sql & "AND E.DT_REGISTRO >= TO_DATE('" & DataInicial & "','dd/mm/yyyy hh24:mi:ss') "
'Sql = Sql & "AND E.DT_REGISTRO <= TO_DATE('" & DataFinal & " 23:59:59','dd/mm/yyyy hh24:mi:ss') "

Sql = Sql & " from "
Sql = Sql & " " & BD_I & ".TB_BL E LEFT JOIN "
Sql = Sql & " " & BD_I & ".TB_CAD_PARCEIROS G ON E.IMPORTADOR = G.AUTONUM LEFT JOIN  "
Sql = Sql & " " & BD_I & ".TB_AUDITORIA_PACOTE H ON E.AUDIT_DESEMBARACO = H.AUTONUM "
Sql = Sql & " WHERE E.DT_REGISTRO >= TO_DATE('" & DataInicial & "','dd/mm/yyyy hh24:mi:ss') "
Sql = Sql & "AND E.DT_REGISTRO <= TO_DATE('" & DataFinal & " 23:59:59','dd/mm/yyyy hh24:mi:ss') "


PRSet Tbh, Sql, 0
If Not Tbh.EOF Then
    While Not Tbh.EOF = True
        
        
        CT20 = Empty
        CT40 = Empty
        VOL = Empty
        SAIDA20 = Empty
        SAIDA40 = Empty
        SAIDAVOL = Empty
        SAIDAReal20 = Empty
        SAIDAReal40 = Empty
        SAIDARealVOL = Empty
        DtSaida = "01/01/1980"
        DtSaidaCs = "01/01/1980"
        Transporte = Empty
        Sql = Empty
        Sql = Sql & "SELECT "
        Sql = Sql & "C.TAMANHO, "
        Sql = Sql & "C.AUTONUM "
        Sql = Sql & " from "
        Sql = Sql & "" & BD_I & ".TB_CNTR_BL C, "
        Sql = Sql & "" & BD_I & ".TB_AMR_CNTR_BL D, "
        Sql = Sql & "" & BD_I & ".TB_BL E "
        Sql = Sql & " where "
        Sql = Sql & "C.AUTONUM = D.CNTR "
        If sPatio <> "0" Then Sql = Sql & "AND C.PATIO = '" & sPatio & "' "
        Sql = Sql & "AND D.BL = E.AUTONUM "
        Sql = Sql & "AND (FLAG_DESOVADO = 0 OR FLAG_DESOVADO IS NULL) "
        Sql = Sql & "AND (FLAG_TERMINAL=1) "
        Sql = Sql & "AND E.AUTONUM = " & Tbh!AUTONUM & " "
        PRSet oRS(0), Sql, 0
        While Not oRS(0).EOF = True
            If oRS(0)!Tamanho = 20 Then CT20 = CT20 + 1
            If oRS(0)!Tamanho = 40 Then CT40 = CT40 + 1
            Sql = Empty
            Sql = Sql & "SELECT "
            Sql = Sql & "F.RAZAO AS TRANSPORTADOR, "
            Sql = Sql & "TO_CHAR(B.DT_SAIDA,'dd/mm/yyyy hh24:mi:ss') AS DT_SAIDA "
            
            
          '  Sql = Sql & " from "
          '  Sql = Sql & "" & BD_I & ".TB_REGISTRO_SAIDA_CNTR A, "
          '  Sql = Sql & "" & BD_I & ".TB_ORDEM_CARREGAMENTO B, "
          '  Sql = Sql & "" & BD_O & ".TB_CAD_TRANSPORTADORAS F "
          '  Sql = Sql & " where "
          '  Sql = Sql & "A.ORDEM_CARREG = B.AUTONUM  "
          '  Sql = Sql & "AND B.TRANSPORTADORA = F.AUTONUM(+) "
           ' Sql = Sql & "AND NOT B.DT_SAIDA IS NULL "
           ' Sql = Sql & "AND A.CNTR = " & oRS(0)!Autonum & " "
          '  Sql = Sql & "ORDER BY "
          '  Sql = Sql & "B.DT_SAIDA "
            
            
            Sql = Sql & " from "
            Sql = Sql & " " & BD_I & ".TB_REGISTRO_SAIDA_CNTR A  "
            Sql = Sql & " INNER JOIN " & BD_I & ".TB_ORDEM_CARREGAMENTO B ON  A.ORDEM_CARREG = B.AUTONUM "
            Sql = Sql & " LEFT JOIN " & BD_O & ".TB_CAD_TRANSPORTADORAS F  ON B.TRANSPORTADORA = F.AUTONUM "
            Sql = Sql & " where   NOT B.DT_SAIDA IS NULL "
            Sql = Sql & " AND A.CNTR = " & oRS(0)!AUTONUM & " "
            Sql = Sql & " ORDER BY "
            Sql = Sql & " B.DT_SAIDA "
            
        
            
            PRSet oRS(1), Sql, 0
            While Not oRS(1).EOF = True
                If oRS(0)!Tamanho = 20 Then SAIDA20 = SAIDA20 + 1
                    If oRS(0)!Tamanho = 40 Then SAIDA40 = SAIDA40 + 1
                    Transporte = Nnull(oRS(1)!Transportador, 0)
                    If Not IsNull(oRS(1)!dt_saida) Then
                    If CVDate(Format(oRS(1)!dt_saida, "DD/MM/YYYY")) >= CVDate(DataInicial) Then
                        If oRS(0)!Tamanho = 20 Then SAIDAReal20 = SAIDAReal20 + 1
                        If oRS(0)!Tamanho = 40 Then SAIDAReal40 = SAIDAReal40 + 1
                        If DtSaida < oRS(1)!dt_saida Then DtSaida = oRS(1)!dt_saida
                    End If
                End If
                oRS(1).MoveNext
            Wend
            oRS(0).MoveNext
            oRS(1).Close
        Wend
        oRS(0).Close
        Sql = Empty
        Sql = Sql & "SELECT "
        Sql = Sql & "C.QUANTIDADE_REAL, "
        Sql = Sql & "C.AUTONUM "
        Sql = Sql & " from "
        Sql = Sql & "" & BD_I & ".TB_CARGA_SOLTA C, "
        Sql = Sql & "" & BD_I & ".TB_BL E "
        Sql = Sql & " where "
        Sql = Sql & "C.BL = E.AUTONUM "
        If sPatio <> "0" Then Sql = Sql & "AND C.PATIO = '" & sPatio & "' "
        Sql = Sql & "AND (FLAG_TERMINAL=1) "
        Sql = Sql & "AND E.AUTONUM = " & Tbh!AUTONUM & " "
        PRSet oRS(0), Sql, 0
        While Not oRS(0).EOF = True
            VOL = VOL + Nnull(oRS(0)!QUANTIDADE_REAL, 0)
            Sql = Empty
            Sql = Sql & "SELECT "
            Sql = Sql & "A.QUANTIDADE, "
            Sql = Sql & "F.RAZAO AS TRANSPORTADOR, "
            Sql = Sql & "TO_CHAR(B.DT_SAIDA,'dd/mm/yyyy hh24:mi:ss') AS DT_SAIDA "
            
         '   Sql = Sql & " from "
         '   Sql = Sql & "" & BD_I & ".TB_REGISTRO_SAIDA_CS A, "
         '   Sql = Sql & "" & BD_I & ".TB_ORDEM_CARREGAMENTO B, "
         '   Sql = Sql & "" & BD_O & ".TB_CAD_TRANSPORTADORAS F "
         '   Sql = Sql & " where "
         '   Sql = Sql & "A.ORDEM_CARREG = B.AUTONUM "
         '   Sql = Sql & "AND B.TRANSPORTADORA = F.AUTONUM(+) "
         '   Sql = Sql & "AND A.CS=" & oRS(0)!Autonum & " "
         '   Sql = Sql & " AND NOT B.DT_SAIDA IS NULL"
            
            Sql = Sql & " from "
            Sql = Sql & " " & BD_I & ".TB_REGISTRO_SAIDA_CS A "
            Sql = Sql & " INNER JOIN " & BD_I & ".TB_ORDEM_CARREGAMENTO B ON A.ORDEM_CARREG = B.AUTONUM "
            Sql = Sql & " LEFT JOIN " & BD_O & ".TB_CAD_TRANSPORTADORAS F ON B.TRANSPORTADORA = F.AUTONUM "
            Sql = Sql & " where "
            Sql = Sql & " A.CS=" & oRS(0)!AUTONUM & " "
            Sql = Sql & " AND NOT B.DT_SAIDA IS NULL"
   
            
            PRSet oRS(1), Sql, 0
            While Not oRS(1).EOF
                SAIDAVOL = SAIDAVOL + oRS(1)!Quantidade
                Transporte = Nnull(oRS(1)!Transportador, 0)
                If CVDate(Format(oRS(1)!dt_saida, "DD/MM/YYYY")) >= CVDate(DataInicial) Then
                    SAIDARealVOL = SAIDARealVOL + oRS(1)!Quantidade
                    DtSaidaCs = oRS(1)!dt_saida
                End If
                oRS(1).MoveNext
            Wend
            oRS(0).MoveNext
            oRS(1).Close
        Wend
        oRS(0).Close
        Sql = Empty
        Sql = Sql & "INSERT INTO " & BD_I & ".TB_TEMP_RELREGISTRO (AUTONUM ,NUM_DOCUMENTO,TAM20,TAM40,VOLUME,IMPORTADOR,TRANSPORTADORA,SALDO20,SALDO40,SALDOVOL,SALDOREAL20,SALDOREAL40,SALDOREALVOL,RECEBIMENTO,DESEMBARACO,SAIDA,desembaraco_cs,saida_cs,PATIO,lote) VALUES "
        Sql = Sql & "(" & BD_I & ".SEQ_TEMP_RELREGISTRO.NEXTVAL,"
        Sql = Sql & "'" & Tbh!Num_Documento & "',"
        Sql = Sql & CT20 & ","
        Sql = Sql & CT40 & ","
        Sql = Sql & VOL & ","
        Sql = Sql & "'" & Tbh!Importador & "',"
        Sql = Sql & "'" & Transporte & "',"
        Sql = Sql & SAIDA20 & ","
        Sql = Sql & SAIDA40 & ","
        Sql = Sql & SAIDAVOL & ","
        Sql = Sql & SAIDAReal20 & ","
        Sql = Sql & SAIDAReal40 & ","
        Sql = Sql & SAIDARealVOL & ","
        Sql = Sql & "TO_DATE('" & Nnull(Tbh!Data_Reg, 1) & "','DD/MM/YYYY hh24:mi:ss'),"
        'DATA DO DESEMBARACO=DATA DO CARREGAMENTO
        DATAD = ""
        DATAD_cs = ""
        Set tbC = New ADODB.Recordset
        tbC.Open "SELECT MAX(A.DATA_ORDEM) AS DATA_ORDEM FROM " & BD_I & ".TB_ORDEM_CARREGAMENTO A," & BD_I & ".TB_REGISTRO_SAIDA_CNTR B," & BD_I & ".TB_CNTR_BL C," & BD_I & ".TB_AMR_CNTR_BL D," & BD_I & ".TB_BL E WHERE B.ORDEM_CARREG=A.AUTONUM AND B.CNTR=C.AUTONUM AND D.CNTR=C.AUTONUM AND D.BL=E.AUTONUM AND E.NUM_DOCUMENTO='" & Tbh!Num_Documento & "'", db, 3, 3
        If Not tbC.EOF Then
            DATAD = Format(tbC!data_ordem, "DD/MM/YYYY HH:MM:SS")
        End If
        tbC.Close
        Set tbC = New ADODB.Recordset
        tbC.Open "SELECT MAX(A.DATA_ORDEM) AS DATA_ORDEM FROM " & BD_I & ".TB_ORDEM_CARREGAMENTO A," & BD_I & ".TB_REGISTRO_SAIDA_CS B," & BD_I & ".TB_CARGA_SOLTA C," & BD_I & ".TB_BL E WHERE B.ORDEM_CARREG=A.AUTONUM AND B.CS=C.AUTONUM AND C.BL=E.AUTONUM AND E.NUM_DOCUMENTO='" & Tbh!Num_Documento & "'", db, 3, 3
        If Not tbC.EOF Then    ' And DATAD = "" Then
            DATAD_cs = Format(tbC!data_ordem, "DD/MM/YYYY HH:MM:SS")
        End If
        ' SQL = SQL & "TO_DATE('" & Nnull(TBh!Dt_Envio, 1) & "','DD/MM/YYYY hh24:mi:ss'),"
        If DATAD <> "" Then
            Sql = Sql & "TO_DATE('" & Nnull(DATAD, 1) & "','DD/MM/YYYY hh24:mi:ss'),"
        Else
            Sql = Sql & "NULL,"
        End If
        If Format(DtSaida, "DD/MM/YYYY") <> "01/01/1980" Then
            Sql = Sql & "TO_DATE('" & Nnull(DtSaida, 1) & "','DD/MM/YYYY hh24:mi:ss'),"
        Else
            Sql = Sql & "NULL,"
        End If
        If DATAD_cs <> "" Then
            Sql = Sql & "TO_DATE('" & Nnull(DATAD_cs, 1) & "','DD/MM/YYYY hh24:mi:ss'),"
        Else
            Sql = Sql & "NULL,"
        End If
        If Format(DtSaidaCs, "DD/MM/YYYY") <> "01/01/1980" Then
            Sql = Sql & "TO_DATE('" & Nnull(DtSaidaCs, 1) & "','DD/MM/YYYY hh24:mi:ss'), "
        Else
            Sql = Sql & "NULL, "
        End If
        Sql = Sql & " '" & sPatio & "'"
        Sql = Sql & ", " & Nnull(Tbh!AUTONUM, 0)
        Sql = Sql & ")"
        Executa (Sql)
        Tbh.MoveNext
    Wend
End If
Tbh.Close
Screen.MousePointer = 11
Sql = "  SELECT , TB_TEMP_RELREGISTRO.lote,TB_TEMP_RELREGISTRO.NUM_DOCUMENTO, TB_TEMP_RELREGISTRO.IMPORTADOR, TB_TEMP_RELREGISTRO.TRANSPORTADORA, TB_TEMP_RELREGISTRO.RECEBIMENTO, TB_TEMP_RELREGISTRO.DESEMBARACO, TB_TEMP_RELREGISTRO.SAIDA, TB_TEMP_RELREGISTRO.TAM40, TB_TEMP_RELREGISTRO.SALDO20, TB_TEMP_RELREGISTRO.SALDO40, TB_TEMP_RELREGISTRO.TAM20, TB_TEMP_RELREGISTRO.SALDOREAL20, TB_TEMP_RELREGISTRO.SALDOREAL40, TB_PATIOS.DESCR_RESUMIDO"
Sql = Sql & "  FROM   " & BD_O & ".TB_PATIOS TB_PATIOS, " & BD_I & ".TB_TEMP_RELREGISTRO TB_TEMP_RELREGISTRO"
Sql = Sql & "  WHERE TB_TEMP_RELREGISTRO.PATIO = TB_PATIOS.DESCR "
Sql = Sql & " and isnull(TB_TEMP_RELREGISTRO.tam20,0) + isnull(TB_TEMP_RELREGISTRO.tam40,0) >0"
Sql = Sql & "  ORDER BY TB_TEMP_RELREGISTRO.RECEBIMENTO, TB_TEMP_RELREGISTRO.NUM_DOCUMENTO"
RptName$ = App.Path & "\" & LOCAL_REPORT & "\RegistroSaidaCntr.RPT"
RptPaperOrientation = crLandscape
RptFormulasName$(0) = "{@usuario}"
RptFormulas$(0) = "'" & Usuario_Sistema & "'"
RptFormulasName$(1) = "{@periodoini}"
RptFormulas$(1) = "'" & DataInicial & "'"
RptFormulasName$(2) = "{@periodofim}"
RptFormulas$(2) = "'" & DataFinal & "'"
SqlQuery = CONVERTE_SQL(Sql)
FrmReportViewer.Show 1: Set FrmReportViewer = Nothing
Screen.MousePointer = 0

'Principal.CrystalReport1.ReportFileName = App.Path & "\" & "RegistroSaidaCntr.rpt"
'Principal.CrystalReport1.Formulas(0) = "usuario = '" & Usuario_Sistema & "'"
'Principal.CrystalReport1.Formulas(1) = "periodoini = '" & DataInicial & "'"
'Principal.CrystalReport1.Formulas(2) = "periodofim = '" & DataFinal & "'"
'Principal.CrystalReport1.Action = 1
'For I = 0 To 2
'    Principal.CrystalReport1.Formulas(I) = ""
'Next I
'Screen.MousePointer = 0

Sql = "  SELECT TB_TEMP_RELREGISTRO.NUM_DOCUMENTO, TB_TEMP_RELREGISTRO.IMPORTADOR, TB_TEMP_RELREGISTRO.TRANSPORTADORA, TB_TEMP_RELREGISTRO.RECEBIMENTO, TB_TEMP_RELREGISTRO.SALDO20, TB_TEMP_RELREGISTRO.SALDOVOL, TB_TEMP_RELREGISTRO.VOLUME, TB_TEMP_RELREGISTRO.SALDOREALVOL, TB_TEMP_RELREGISTRO.DESEMBARACO_CS, TB_TEMP_RELREGISTRO.SAIDA_CS, TB_PATIOS.DESCR_RESUMIDO"
Sql = Sql & "  FROM   " & BD_O & ".TB_PATIOS TB_PATIOS, " & BD_I & ".TB_TEMP_RELREGISTRO TB_TEMP_RELREGISTRO"
Sql = Sql & "  Where TB_TEMP_RELREGISTRO.Volume <> 0 AND TB_TEMP_RELREGISTRO.PATIO = TB_PATIOS.DESCR"
Sql = Sql & "  ORDER BY TB_TEMP_RELREGISTRO.RECEBIMENTO, TB_TEMP_RELREGISTRO.NUM_DOCUMENTO"
RptName$ = App.Path & "\" & LOCAL_REPORT & "\RegistroSaidaCs.RPT"
RptPaperOrientation = crLandscape
RptFormulasName$(0) = "{@usuario}"
RptFormulas$(0) = "'" & Usuario_Sistema & "'"
RptFormulasName$(1) = "{@periodoini}"
RptFormulas$(1) = "'" & DataInicial & "'"
RptFormulasName$(2) = "{@periodofim}"
RptFormulas$(2) = "'" & DataFinal & "'"
SqlQuery = CONVERTE_SQL(Sql)
FrmReportViewer.Show 1: Set FrmReportViewer = Nothing
Screen.MousePointer = 0


'Screen.MousePointer = 11
'Principal.CrystalReport1.ReportFileName = App.Path & "\" & "RegistroSaidaCs.rpt"
'Principal.CrystalReport1.Formulas(0) = "usuario = '" & Usuario_Sistema & "'"
'Principal.CrystalReport1.Formulas(1) = "periodoini = '" & DataInicial & "'"
'Principal.CrystalReport1.Formulas(2) = "periodofim = '" & DataFinal & "'"
'Principal.CrystalReport1.Action = 1
'For I = 0 To 2
'    Principal.CrystalReport1.Formulas(I) = ""
'Next I
'Screen.MousePointer = 0

End Sub
Public Sub RunRelProtGMVI()
Dim Viagem As String
Viagem = InputBox("Informe o número da viagem", , Viagens.Text1(0))
If Viagem = "" Then Exit Sub
Screen.MousePointer = 11
'Set Rpt = Principal.CrystalReport1
'Rpt.ReportFileName = Drv + "\PROTOCOLO_GMVI.rpt"
'Rpt.Formulas(0) = "usuario = '" & Usuario_Sistema & "'"
'Rpt.Formulas(1) = "viagem = '" & Viagem & "'"
'Rpt.Action = 0
'Rpt.Formulas(0) = ""
'Rpt.Formulas(1) = ""
'Rpt.SelectionFormula = ""
'Screen.MousePointer = 0
End Sub
Public Sub RunRelGMVI()
Dim tbh1 As ADODB.Recordset
Dim Viagem As String
Viagem = InputBox("Informe o número da viagem", , Viagens.Text1(0))
If Viagem = "" Then Exit Sub
Screen.MousePointer = 11
Dim GMVI As String
GMVI = InputBox("Informe o número da GMVI,(0) Para todas", , 0)
If GMVI = "" Then Exit Sub
If GMVI = "0" Then GMVI = ""
Dim DT As String
DT = InputBox("Informe o número da DT,(0) Para todas", , 0)
If DT = "" Then Exit Sub
If DT = "0" Then DT = ""
Sql = Empty
Sql = Sql & " SELECT "
Sql = Sql & " A.GMVI, "
Sql = Sql & " A.ITEM, "
Sql = Sql & " A.DTA, "
Sql = Sql & " A.BL, "
Sql = Sql & " A.QUANTIDADE, "
Sql = Sql & " A.PESO_BRUTO, "
Sql = Sql & " A.MERCADORIA, "
Sql = Sql & " B.DESCR AS DESCR_EMBALAGEM, "
Sql = Sql & " C.DESCR AS DESCR_GENERO, "
Sql = Sql & " A.GMVI, "
Sql = Sql & " A.DT, "
Sql = Sql & " A.VIAGEM, "
Sql = Sql & " A.PATIO "

'Sql = Sql & " FROM "
'Sql = Sql & " " & BD_I & ".TB_CARGA_SOLTA A, "
'Sql = Sql & " " & BD_I & ".DTE_TB_EMBALAGENS B, "
'Sql = Sql & " " & BD_I & ".TB_CAD_GRUPO_PRODUTOS C "
'Sql = Sql & " WHERE "
'Sql = Sql & " A.EMBALAGEM = B.CODE(+) "
'Sql = Sql & " AND A.GENERO = C.AUTONUM(+) "
'If Not GMVI = Empty Then Sql = Sql & " AND A.gmvi = '" & GMVI & "' "
'If Not DT = Empty Then Sql = Sql & " AND A.DT = '" & DT & "' "
'If Not Viagem = Empty Then Sql = Sql & " AND A.VIAGEM = '" & Viagem & "' "
'Sql = Sql & " ORDER BY A.ITEM "


Sql = Sql & " FROM "
Sql = Sql & " " & BD_I & ".TB_CARGA_SOLTA A  "
Sql = Sql & " LEFT JOIN " & BD_I & ".DTE_TB_EMBALAGENS B  ON A.EMBALAGEM = B.CODE"
Sql = Sql & " LEFT JOIN " & BD_I & ".TB_CAD_GRUPO_PRODUTOS C ON A.GENERO = C.AUTONUM "
Sql = Sql & " WHERE 1=1 "


If Not GMVI = Empty Then Sql = Sql & " AND A.gmvi = '" & GMVI & "' "
If Not DT = Empty Then Sql = Sql & " AND A.DT = '" & DT & "' "
If Not Viagem = Empty Then Sql = Sql & " AND A.VIAGEM = '" & Viagem & "' "
Sql = Sql & " ORDER BY A.ITEM "







PRSet tbh1, Sql
While Not tbh1.EOF = True
    'ADE02 - FASE 2
    If IsNull(tbh1!Dta) = True Then
        REL_GMVI tbh1!DT, tbh1!GMVI, tbh1!Viagem, 1, (Format$(Now, "YY") & Format$(tbh1!BL, "000000") & "0")
    Else
        REL_GMVI tbh1!DT, tbh1!GMVI, tbh1!Viagem, 2, (Format$(Now, "YY") & Format$(tbh1!BL, "000000") & "0")
    End If
    tbh1.MoveNext
Wend
tbh1.Close
Screen.MousePointer = 0
End Sub
Public Sub RunExportaLogix()
Dim DataIni As String
Dim DataFim As String
Dim Arquivo As String
Arquivo = InputBox("Qual o nome do arquivo", , "c:\" & Format(Now, "DDMM") & Format(Now, "HHMM") & ".TXT")
If Arquivo = "" Then Exit Sub
DataIni = InputBox("Qual a data Inicial", , Format(Now, "dd/mm/? yyyy"))
If DataIni = "" Then Exit Sub
If Not IsDate(DataIni) Then
    MsgBox "Data invalida", vbCritical: Exit Sub
End If
DataFim = InputBox("Qual a data final", , Format(Now, "dd/mm/yyyy"))
If DataFim = "" Then Exit Sub
If Not IsDate(DataFim) Then
    MsgBox "Data invalida", vbCritical: Exit Sub
End If
If CVDate(DataFim) < CVDate(DataIni) Then
    MsgBox "Data final anterior a inicial", vbCritical: Exit Sub
End If
Dim Xfatura As String
Xfatura = "S"
Set TbGr = New ADODB.Recordset
TbGr.Open "Select forma_pagamento,bl,SEQ_GR,dt_gr from " & BD_I & ".TB_GR_BL where dt_gr>=to_date('" & Format(DataIni, "dd/mm/yyyy") & " 00:00:00','dd/mm/yyyy hh24:mi:ss') and dt_gr<to_date('" & Format(DataFim, "dd/mm/yyyy") & " 23:59:59','dd/mm/yyyy hh24:mi:ss') and (flag_logix is null or flag_logix=0) and flag_gr_paga=1", db, 1, 3
If Not TbGr.EOF Then
    Screen.MousePointer = 11
    Open Arquivo For Output As #1
    While Not TbGr.EOF
        Set Tbbl = New ADODB.Recordset
       ' Tbbl.Open "Select a.viagem,a.numero,a.importador,b.cgc,b.razao,b.logradouro,b.cidade,b.estado from " & BD_I & ".TB_BL a," & BD_I & ".TB_CAD_PARCEIROS b where a.importador=b.autonum(+) and a.autonum=" & TbGr!BL, db, 1, 3
        Sql = "Select a.viagem,a.numero,a.importador,b.cgc,b.razao,b.logradouro,b.cidade,b.estado "
        'Sql = Sql & " from " & BD_I & ".TB_BL a," & BD_I & ".TB_CAD_PARCEIROS b "
        'Sql = Sql & " where a.importador=b.autonum(+) and a.autonum=" & TbGr!BL
        
        Sql = Sql & " from " & BD_I & ".TB_BL a left join " & BD_I & ".TB_CAD_PARCEIROS b "
        Sql = Sql & " on  a.importador=b.autonum where a.autonum=" & TbGr!BL
        
        Linha = "47"
        Linha = Linha & "|"
        Linha = Linha & "C"
        Linha = Linha & "|"
        Linha = Linha & Left$(Nnull(Tbbl!Viagem, 1) & s(20), 20)
        Linha = Linha & "|"
        Linha = Linha & Left$(Nnull(Tbbl!Numero, 1) & s(20), 20)
        Linha = Linha & "|0"
        Linha = Linha & Left$(Nnull(Tbbl!CGC, 1) & s(18), 18)
        Linha = Linha & "|"
        Linha = Linha & Left$(Nnull(Tbbl!Razao, 1) & s(36), 36)
        Linha = Linha & "|"
        If Nnull(Tbbl!logradouro, 1) <> "" Then
           Linha = Linha & Left$(Nnull(Tbbl!logradouro, 1) & s(36), 36)
        Else
            Linha = Linha & Left$("NAO CADASTRADA" & s(36), 36)
        End If
        Linha = Linha & "|"
        If Nnull(Tbbl!Cidade, 1) <> "" Then
            Linha = Linha & Left$(Nnull(Tbbl!Cidade, 1) & s(30), 30)
        Else
            Linha = Linha & Left$("NAO CADASTRADA" & s(30), 30)
        End If
        Linha = Linha & "|"
        If Nnull(Tbbl!Estado, 1) <> "" Then
             Linha = Linha & Left$(Nnull(Tbbl!Estado, 1) & s(2), 2)
        Else
            Linha = Linha & Left$("NC" & s(2), 2)
        End If
        Linha = Linha & "|"
        Linha = Linha & Left$("SANTOS" & s(10), 10)
        Linha = Linha & "|"
        Linha = Linha & "0050"
        Linha = Linha & "|"
        Linha = Linha & "22"
        Linha = Linha & "|"
        Linha = Linha & "GR"
        Linha = Linha & "|"
        Linha = Linha & Format(TbGr!SEQ_GR, "0000000000")
        Linha = Linha & "|"
        Linha = Linha & "N"
        Linha = Linha & "|"
        If Nnull(TbGr!Forma_Pagamento, 0) = 3 Then
            Xfatura = "N"
        Else
            Xfatura = "S"
        End If
        Linha = Linha & Xfatura
        Linha = Linha & "|"
        Linha = Linha & "00050"
        Linha = Linha & "|"
        Set TBV = New ADODB.Recordset
        TBV.Open "Select sum(valor+adicional+desconto)as valor from " & BD_I & ".TB_SERVICOS_FATURADOS where seq_gr=" & TbGr!SEQ_GR, db, 1, 3
        'linha = linha & Format(100 * Nnull(TBV!Valor, 0),"00000000000000000")
        Linha = Linha & Left$(Format(100 * Nnull(TBV!Valor, 0), "0000000000000000"), 14) & "." & Right$(Format(100 * Nnull(TBV!Valor, 0), "0000000000000000"), 2)
        TBV.Close
        Linha = Linha & "|"
        Linha = Linha & String(17, "0")
        Linha = Linha & "|"
        Linha = Linha & String(17, "0")
        Linha = Linha & "|"
        Linha = Linha & String(17, "0")
        Linha = Linha & "|"
        Linha = Linha & String(17, "0")
        Linha = Linha & "|"
        Linha = Linha & String(17, "0")
        Linha = Linha & "|"
        Linha = Linha & String(17, "0")
        Linha = Linha & "|"
        Linha = Linha & Format(TbGr!DT_GR, "dd/mm/yyyy")
        Linha = Linha & "|"
        Linha = Linha & "|"
        Print #1, Linha
        Executa "UPDATE " & BD_I & ".TB_GR_BL SET FLAG_LOGIX=1 WHERE SEQ_GR=" & TbGr!SEQ_GR
        TbGr.MoveNext
    Wend
    Close (1)
    Screen.MousePointer = 0
Else
    MsgBox "Nenhum registro encontrado", vbInformation
End If
TbGr.Close
End Sub
Public Sub RunDesovaSemDevolução()
 
Dim Sql As String
Screen.MousePointer = vbHourglass
Sql = Empty
Sql = "DELETE FROM " & BD_I & ".TMP_DESOVA_SEM_DEV WHERE USUARIO = '" & Usuario_Sistema & "'"
Executa Sql

intbutton = MsgBox("Incluir DDC?", 4, "")

Sql = Empty
Sql = Sql & "INSERT INTO "
Sql = Sql & "" & BD_I & ".TMP_DESOVA_SEM_DEV "
Sql = Sql & "(LOTE, VIAGEM, SIGLA, INICIO_DESOVA, FIM_DESOVA, USUARIO,TIPO,ARMADOR,TAMANHO,DEMURRAGE,DESTINO,MARGEM,dias) "

Sql = Sql & "SELECT "
Sql = Sql & "MAX(B.AUTONUM) AS LOTE, "
Sql = Sql & "MAX(B.VIAGEM) AS VIAGEM, "
Sql = Sql & "A.ID_CONTEINER AS SIGLA, "
Sql = Sql & "A.DT_INICIO_DESOVA AS INICIO_DESOVA, "
Sql = Sql & "A.DT_FIM_DESOVA AS FIM_DESOVA, "
Sql = Sql & "'" & Usuario_Sistema & "' AS USUARIO "
Sql = Sql & ",MAX(A.TIPO) AS TIPO, MAX(B.ARMADOR) ARMADOR, A.TAMANHO, A.DATA_DEMURRAGE,D.DESTINO"
Sql = Sql & " ,DECODE(MARGEM,0,'Direita',1,'Entre Margens',2,'Esquerda','') AS MARGEM "
Sql = Sql & " ,datediff(day,a.dt_entrada,getdate()) as dias"
Sql = Sql & " from "
Sql = Sql & BD_I & ".TB_CNTR_BL A left OUTER join ( "
Sql = Sql & " SELECT CNTR_CARRETA_1 AS CNTR, DT_DEV_CNTR_CARRETA_1 AS DEVOLUCAO "
Sql = Sql & " , DESTINO_CNTR_CARRETA_1 AS DESTINO ,MARGEM   FROM " & BD_I & ".TB_DEVOLUCAO_VAZIO "

Sql = Sql & " Union All "

Sql = Sql & " SELECT CNTR_CARRETA_2 AS CNTR, DT_DEV_CNTR_CARRETA_2 AS DEVOLUCAO "
Sql = Sql & " , DESTINO_CNTR_CARRETA_2 AS DESTINO, MARGEM "
Sql = Sql & " FROM " & BD_I & ".TB_DEVOLUCAO_VAZIO) D "
Sql = Sql & " ON A.AUTONUM=D.CNTR," & BD_I & ".TB_BL B, " & BD_I & ". TB_AMR_CNTR_BL c  Where a.Autonum = c.cntr"
Sql = Sql & " AND    B.FLAG_ATIVO  = 1 AND NOT A.DT_INICIO_DESOVA  IS NULL AND NOT A.DT_FIM_DESOVA  IS NULL"
Sql = Sql & " AND    D.DEVOLUCAO  IS NULL  AND    B.AUTONUM  = C.BL"
Sql = Sql & " AND B." & Cod_Empresa_Patio
Sql = Sql & " Group By "
Sql = Sql & " A.ID_CONTEINER, A.TAMANHO,"
Sql = Sql & " A.DT_INICIO_DESOVA, "
Sql = Sql & " A.DT_FIM_DESOVA, "
Sql = Sql & " D.DEVOLUCAO, A.DATA_DEMURRAGE,D.DESTINO,D.MARGEM,datediff(day,a.dt_entrada,getdate())  "

If intbutton = 6 Then
    Sql = Sql & " "
    Sql = Sql & "         Union All"
    Sql = Sql & " "
    Sql = Sql & "         SELECT MAX (B.AUTONUM) AS LOTE,"
    Sql = Sql & "           MAX (B.VIAGEM) AS VIAGEM,"
    Sql = Sql & "           A.ID_CONTEINER AS SIGLA,"
    Sql = Sql & "           A.DT_INICIO_DESOVA AS INICIO_DESOVA,"
    Sql = Sql & "           A.DT_FIM_DESOVA AS FIM_DESOVA,"
    Sql = Sql & "'" & Usuario_Sistema & "' AS USUARIO, "
    Sql = Sql & "           MAX (A.TIPO) AS TIPO,"
    Sql = Sql & "           MAX (B.ARMADOR) ARMADOR,"
    Sql = Sql & "           A.TAMANHO,"
    Sql = Sql & "           A.DATA_DEMURRAGE,"
    Sql = Sql & "           D.DESTINO,"
    Sql = Sql & "           Case Margem"
    Sql = Sql & "              WHEN 0 THEN 'DIREITA'"
    Sql = Sql & "              WHEN 1 THEN 'ENTRE MARGENS'"
    Sql = Sql & "              WHEN 2 THEN 'ESQUERDA'"
    Sql = Sql & "              Else ''"
    Sql = Sql & "           End"
    Sql = Sql & "              AS MARGEM"
    Sql = Sql & " ,datediff(day,a.dt_entrada,getdate()) as dias"
    Sql = Sql & "      FROM    SGIPA.DBO.TB_CNTR_BL A"
    Sql = Sql & "           Inner Join"
    Sql = Sql & "              (SELECT CNTR_CARRETA_1 AS CNTR,"
    Sql = Sql & "                      DT_DEV_CNTR_CARRETA_1 AS DEVOLUCAO,"
    Sql = Sql & "                      DESTINO_CNTR_CARRETA_1 AS DESTINO,"
    Sql = Sql & "                      Margem"
    Sql = Sql & "                 From SGIPA.DBO.TB_DEVOLUCAO_VAZIO"
    Sql = Sql & "                 Where IsNull(DT_DEV_CNTR_CARRETA_1, 0) = 0 And IsNull(DT_DEV_CNTR_CARRETA_2, 0) = 0"
    Sql = Sql & "               Union All"
    Sql = Sql & "               SELECT CNTR_CARRETA_2 AS CNTR,"
    Sql = Sql & "                      DT_DEV_CNTR_CARRETA_2 AS DEVOLUCAO,"
    Sql = Sql & "                      DESTINO_CNTR_CARRETA_2 AS DESTINO,"
    Sql = Sql & "                      Margem"
    Sql = Sql & "                 From SGIPA.DBO.TB_DEVOLUCAO_VAZIO"
    Sql = Sql & "                 Where IsNull(DT_DEV_CNTR_CARRETA_1, 0) = 0 And IsNull(DT_DEV_CNTR_CARRETA_2, 0) = 0"
    Sql = Sql & "               ) D ON A.AUTONUM = D.CNTR"
    Sql = Sql & " inner join SGIPA.DBO.TB_AMR_CNTR_BL C on a.autonum=c.cntr"
    Sql = Sql & " inner join SGIPA.DBO.TB_BL B on c.bl = b.autonum"
    Sql = Sql & " inner join sgipa.dbo.TB_DESCARGA_ARMAZEM da on c.cntr=da.cntr"
    Sql = Sql & "     Where B.FLAG_ATIVO = 1"
    Sql = Sql & "           AND ISNULL (A.FLAG_DDC, 0) <> 0"
    Sql = Sql & "           AND A.FLAG_TERMINAL = 1"
    Sql = Sql & "           AND B.PATIO IN (1, 7, 8, 9)"
    Sql = Sql & "    GROUP BY A.ID_CONTEINER,"
    Sql = Sql & "             A.TAMANHO,"
    Sql = Sql & "             A.DT_INICIO_DESOVA,"
    Sql = Sql & "             A.DT_FIM_DESOVA,"
    Sql = Sql & "             D.DEVOLUCAO,"
    Sql = Sql & "             A.DATA_DEMURRAGE,"
    Sql = Sql & "             D.DESTINO,"
    Sql = Sql & "             D.Margem"
    Sql = Sql & " ,datediff(day,a.dt_entrada,getdate())"

End If
 

db.Execute CONVERTE_SQL(Sql)
Sql = "SELECT TMP_DESOVA_SEM_DEV.LOTE, TMP_DESOVA_SEM_DEV.VIAGEM, TMP_DESOVA_SEM_DEV.SIGLA, TMP_DESOVA_SEM_DEV.INICIO_DESOVA"
Sql = Sql & " , TMP_DESOVA_SEM_DEV.FIM_DESOVA"
Sql = Sql & ", TMP_DESOVA_SEM_DEV.USUARIO, DTE_TB_TIPOS_CONTEINER.CODIGO, DTE_TB_ARMADORES.DESCR, TMP_DESOVA_SEM_DEV.TAMANHO"
Sql = Sql & " , TMP_DESOVA_SEM_DEV.DEMURRAGE"
Sql = Sql & ", TMP_DESOVA_SEM_DEV.DESTINO,TMP_DESOVA_SEM_DEV.MARGEM,TMP_DESOVA_SEM_DEV.dias "
Sql = Sql & " FROM SGIPA.dbo.TMP_DESOVA_SEM_DEV"
Sql = Sql & " LEFT OUTER JOIN SGIPA.dbo.DTE_TB_TIPOS_CONTEINER ON TMP_DESOVA_SEM_DEV.TIPO=DTE_TB_TIPOS_CONTEINER.CODE"
Sql = Sql & " LEFT OUTER JOIN SGIPA.dbo.DTE_TB_ARMADORES ON TMP_DESOVA_SEM_DEV.ARMADOR=DTE_TB_ARMADORES.CODE"
Sql = Sql & " WHERE TMP_DESOVA_SEM_DEV.USUARIO = '" & Usuario_Sistema & "'"
Sql = Sql & " ORDER BY TMP_DESOVA_SEM_DEV.SIGLA"

RptName$ = App.Path & "\" & LOCAL_REPORT & "\dev0101.RPT"

RptFormulasName$(0) = "{@operador}"
RptFormulas$(0) = "'" & wNomeUsuarioTicket & "'"
SqlQuery = CONVERTE_SQL(Sql)
 RptPaperOrientation = crLandscape
FrmReportViewer.Show 1: Set FrmReportViewer = Nothing
Screen.MousePointer = 0



Sql = Empty
End Sub

Public Sub RunDemoFat()
Dim wGr As String
wGr = InputBox$("Nr. GR :", "Demonstrativo faturamento")
If wGr = "" Then
    Exit Sub
End If
Dim wTb As New ADODB.Recordset
Dim wLINHA As String
Dim wAutonumviagem As Long
Dim wLista As Long

Sql = "select a.line,a.autonumviagem, b.lista "
Sql = Sql & " FROM " & BD_O & ".tb_servicos_faturados a, " & BD_O & ".tb_lista_preco_servicos_fixos b "
Sql = Sql & " where a.servico = b.servico and seq_gr=" & wGr

wTb.Open Sql, db

If wTb.EOF Then
    MsgBox "GR não Encontrada - Verifique"
    wTb.Close
    Exit Sub
End If

wLINHA = Nnull(wTb!Line, 1)
wAutonumviagem = Nnull(wTb!Autonumviagem, 0)
wLista = Nnull(wTb!Lista, 0)
wTb.Close

WSQL = "Select a.viagem, a.data_op_inicio, a.data_op_fim, b.nome"
WSQL = WSQL & " from " & BD_O & ".tb_viagens a, " & BD_O & ".tb_cad_navios b"
WSQL = WSQL & " where a.navio = b.autonum"
WSQL = WSQL & " and a.autonum=" & wAutonumviagem
wTb.Open WSQL, db, adOpenForwardOnly, adLockReadOnly

Dim wViagem As String
Dim wNavio As String
Dim wData1 As String
Dim wData2 As String

If Not wTb.EOF Then
    wViagem = Nnull(wTb!Viagem, 1)
    wNavio = Nnull(wTb!Nome, 1)
    wData1 = Format$(Nnull(wTb!DATA_OP_INICIO, 1), "dd/mm/yyyy")
    wData2 = Format$(Nnull(wTb!DATA_OP_FIM, 1), "dd/mm/yyyy")
End If

wTb.Close

'Principal.CrystalReport1.ReportFileName = App.Path & "\servicosfaturados.rpt"

WSQL = "SELECT"
WSQL = WSQL & " TB_SERVICOS_IPA.DESCR, TB_SERVICOS_IPA.CODGRUPO,"
WSQL = WSQL & " TB_GRUPO_SERVICO.DESCRICAO,"
WSQL = WSQL & " TB_LISTA_PRECO_SERVICOS_FIXOS.LISTA, TB_LISTA_PRECO_SERVICOS_FIXOS.PRECO_UNITARIO,"
WSQL = WSQL & " TB_SERVICOS_FATURADOS.VALOR, TB_SERVICOS_FATURADOS.SEQ_GR, TB_SERVICOS_FATURADOS.QUANTIDADE"

'wSQL = wSQL & " From"
'wSQL = wSQL & " " & BD_O & ".TB_SERVICOS_IPA TB_SERVICOS_IPA,"
'wSQL = wSQL & " " & BD_O & ".TB_GRUPO_SERVICO TB_GRUPO_SERVICO,"
'wSQL = wSQL & " " & BD_O & ".TB_LISTA_PRECO_SERVICOS_FIXOS TB_LISTA_PRECO_SERVICOS_FIXOS,"
'wSQL = wSQL & " " & BD_O & ".TB_SERVICOS_FATURADOS TB_SERVICOS_FATURADOS"
'wSQL = wSQL & " Where"
'wSQL = wSQL & " TB_SERVICOS_IPA.CODGRUPO = TB_GRUPO_SERVICO.CODGRUPO"
'wSQL = wSQL & " AND TB_SERVICOS_IPA.AUTONUM = TB_LISTA_PRECO_SERVICOS_FIXOS.SERVICO"
'wSQL = wSQL & " AND TB_LISTA_PRECO_SERVICOS_FIXOS.SERVICO = TB_SERVICOS_FATURADOS.SERVICO (+)"
'wSQL = wSQL & " AND TB_LISTA_PRECO_SERVICOS_FIXOS.LISTA = 1"
'wSQL = wSQL & " AND (TB_SERVICOS_FATURADOS.SEQ_GR = 1 OR TB_SERVICOS_FATURADOS.SEQ_GR IS NULL)"
'wSQL = wSQL & " Order By"
'wSQL = wSQL & " TB_SERVICOS_IPA.CODGRUPO ASC,"
'wSQL = wSQL & " TB_SERVICOS_IPA.DESCR ASC"


WSQL = WSQL & " From "
WSQL = WSQL & " " & BD_O & ".TB_SERVICOS_IPA TB_SERVICOS_IPA "
WSQL = WSQL & " INNER JOIN " & BD_O & ".TB_GRUPO_SERVICO TB_GRUPO_SERVICO "
WSQL = WSQL & " ON TB_SERVICOS_IPA.CODGRUPO = TB_GRUPO_SERVICO.CODGRUPO "
WSQL = WSQL & " INNER JOIN " & BD_O & ".TB_LISTA_PRECO_SERVICOS_FIXOS TB_LISTA_PRECO_SERVICOS_FIXOS"
WSQL = WSQL & " TB_SERVICOS_IPA.AUTONUM = TB_LISTA_PRECO_SERVICOS_FIXOS.SERVICO"
WSQL = WSQL & " LEFT JOIN " & BD_O & ".TB_SERVICOS_FATURADOS TB_SERVICOS_FATURADOS"
WSQL = WSQL & " ON TB_LISTA_PRECO_SERVICOS_FIXOS.SERVICO = TB_SERVICOS_FATURADOS.SERVICO "

WSQL = WSQL & " WHERE TB_LISTA_PRECO_SERVICOS_FIXOS.LISTA = 1"
WSQL = WSQL & " AND (TB_SERVICOS_FATURADOS.SEQ_GR = 1 OR TB_SERVICOS_FATURADOS.SEQ_GR IS NULL)"
WSQL = WSQL & " Order By"
WSQL = WSQL & " TB_SERVICOS_IPA.CODGRUPO ASC,"
WSQL = WSQL & " TB_SERVICOS_IPA.DESCR ASC"

'Z3 ?????????????/

'Principal.CrystalReport1.SqlQuery = wSQL
'Principal.CrystalReport1.Formulas(0) = "viagem='" & wViagem & "'"
'Principal.CrystalReport1.Formulas(1) = "navio='" & wNavio & "'"
'Principal.CrystalReport1.Formulas(2) = "entrada='" & wData1 & "'"
'Principal.CrystalReport1.Formulas(3) = "saida='" & wData2 & "'"
'Principal.CrystalReport1.Action = 1
'Principal.CrystalReport1.Formulas(0) = ""
'Principal.CrystalReport1.Formulas(1) = ""
'Principal.CrystalReport1.Formulas(2) = ""
'Principal.CrystalReport1.Formulas(3) = ""
'Principal.CrystalReport1.ReplaceSelectionFormula ""
End Sub




Public Sub CorrigeTBPrecos()
'''''CORREÇÃO DE REGISTROS (TABELAS DE COBRANÇA)
Dim OTBBBBBBB As ADODB.Recordset
Dim Sql As String
Dim t As Long
For t = 0 To 1000
    Sql = Empty
    Sql = Sql & "SELECT "
    Sql = Sql & "MAX(A.AUTONUM) AS AUTONUM, "
    Sql = Sql & "A.LISTA, "
    Sql = Sql & "A.SERVICO, "
    Sql = Sql & "A.TIPO_CARGA, "
    Sql = Sql & "A.BASE_CALCULO, "
    Sql = Sql & "A.VARIANTE_LOCAL, "
    Sql = Sql & "A.PRECO_UNITARIO, "
    Sql = Sql & "A.MOEDA, "
    Sql = Sql & "A.PRECO_MINIMO, "
    Sql = Sql & "A.VALOR_ACRESCIMO, "
    Sql = Sql & "A.LOCAL_ATRACACAO, "
    Sql = Sql & "A.AUTONUM_VINCULADO, "
    Sql = Sql & "MAX(A.USUARIO_REDE) AS USUARIO_REDE,MAX(A.USUARIO_SIS) AS USUARIO_SIS,MAX(A.MAQUINA_REDE) AS MAQUINA_REDE "
    Sql = Sql & " from "
    Sql = Sql & "TB_LISTA_PRECO_SERVICOS_FIXOS a "
    Sql = Sql & " where "
    Sql = Sql & "LISTA = " & t & " "
    Sql = Sql & "GROUP BY "
    Sql = Sql & "A.LISTA, "
    Sql = Sql & "A.SERVICO, "
    Sql = Sql & "A.TIPO_CARGA, "
    Sql = Sql & "A.BASE_CALCULO, "
    Sql = Sql & "A.VARIANTE_LOCAL, "
    Sql = Sql & "A.PRECO_UNITARIO, "
    Sql = Sql & "A.MOEDA, "
    Sql = Sql & "A.PRECO_MINIMO, "
    Sql = Sql & "A.VALOR_ACRESCIMO, "
    Sql = Sql & "A.LOCAL_ATRACACAO, "
    Sql = Sql & "A.AUTONUM_VINCULADO "
    Sql = Sql & "ORDER BY "
    Sql = Sql & "LISTA "
    PRSet OTBBBBBBB, Sql, 0
    Executa "DELETE FROM " & BD_I & ".TB_LISTA_PRECO_SERVICOS_FIXOS WHERE LISTA = " & t & " "
    While Not OTBBBBBBB.EOF = True
        Sql = Empty
        Sql = Sql & "INSERT INTO " & BD_I & ".TB_LISTA_PRECO_SERVICOS_FIXOS (AUTONUM,LISTA,SERVICO,TIPO_CARGA,BASE_CALCULO,VARIANTE_LOCAL,PRECO_UNITARIO,MOEDA,PRECO_MINIMO,VALOR_ACRESCIMO,LOCAL_ATRACACAO,AUTONUM_VINCULADO, USUARIO_REDE, USUARIO_SIS, MAQUINA_REDE ) VALUES ("
        Sql = Sql & "'" & OTBBBBBBB.Fields("AUTONUM") & "', "
        Sql = Sql & "'" & OTBBBBBBB.Fields("LISTA") & "', "
        Sql = Sql & "'" & OTBBBBBBB.Fields("SERVICO") & "', "
        Sql = Sql & "'" & OTBBBBBBB.Fields("TIPO_CARGA") & "', "
        Sql = Sql & "'" & OTBBBBBBB.Fields("BASE_CALCULO") & "', "
        Sql = Sql & "'" & OTBBBBBBB.Fields("VARIANTE_LOCAL") & "', "
        Sql = Sql & "'" & OTBBBBBBB.Fields("PRECO_UNITARIO") & "', "
        Sql = Sql & "'" & OTBBBBBBB.Fields("MOEDA") & "', "
        Sql = Sql & "'" & OTBBBBBBB.Fields("PRECO_MINIMO") & "', "
        Sql = Sql & "'" & OTBBBBBBB.Fields("VALOR_ACRESCIMO") & "', "
        Sql = Sql & "'" & OTBBBBBBB.Fields("LOCAL_ATRACACAO") & "', "
        Sql = Sql & "'" & OTBBBBBBB.Fields("AUTONUM_VINCULADO") & "', "
        Sql = Sql & "'" & OTBBBBBBB.Fields("USUARIO_REDE") & "', "
        Sql = Sql & "'" & OTBBBBBBB.Fields("USUARIO_SIS") & "', "
        Sql = Sql & "'" & OTBBBBBBB.Fields("MAQUINA_REDE") & "' "
        Sql = Sql & ") "
        Executa Sql
        OTBBBBBBB.MoveNext
    Wend
    OTBBBBBBB.Close
Next t
End Sub
Public Sub AcertaCntrsDasGRs()
Set TbGr = New ADODB.Recordset
TbGr.Open "SELECT AUTONUM,BL FROM TB_GR_BL  ORDER BY BL", db, 1, 1
If Not TbGr.EOF Then
    While Not TbGr.EOF
        Set tbC = New ADODB.Recordset
        
        Sql = "Select count(a.id_conteiner) as qtos "
        Sql = Sql & " From TB_CNTR_BL a INNER JOIN TB_AMR_CNTR_BL b ON b.Cntr = a.autonum "
        Sql = Sql & " LEFT JOIN TB_REGISTRO_SAIDA_CNTR c ON C.CNTR = A.AUTONUM "
        Sql = Sql & " LEFT JOIN TB_ORDEM_CARREGAMENTO D ON D.AUTONUM = C.ORDEM_CARREG "
        Sql = Sql & " Where  a.flag_terminal = 1 and b.bl = " & TbGr!BL
                
        tbC.Open Sql
        
        Executa "UPDATE TB_GR_BL SET F_QTD_CNTR=" & Nnull(tbC!qtos, 0) & " WHERE AUTONUM=" & TbGr!AUTONUM
        
        tbC.Close
        Set tbcs = New ADODB.Recordset
        Sql = "Select SUM(a.quantidade) AS QTDE ,SUM(A.PESO_BRUTO) AS PB "
        'Sql = Sql & " from TB_CARGA_SOLTA a,DTE_TB_EMBALAGENS b "
        'Sql = Sql & " where a.embalagem=b.code(+) and a.bl= " & TbGr!BL
        'Sql = Sql & " and a.quantidade<>0"
        
        Sql = Sql & " from TB_CARGA_SOLTA a LEFT JOIN DTE_TB_EMBALAGENS b "
        Sql = Sql & " ON a.embalagem=b.code"
        Sql = Sql & " WHERE a.bl= " & TbGr!BL
        Sql = Sql & " and a.quantidade<>0"
        
        
        
        
        tbcs.Open Sql
        Executa "UPDATE TB_GR_BL SET F_QTD_VOLUMES=" & Nnull(tbcs!qtde, 0) & ",F_PESO_CS=" & str(Nnull(tbcs!PB, 0)) & " WHERE AUTONUM=" & TbGr!AUTONUM
        tbcs.Close
        TbGr.MoveNext
    Wend
End If
MsgBox "FIM"
End Sub
Private Sub AcertaMercasCS()
Dim tb As ADODB.Recordset
Sql = Empty
Sql = Sql & "select "
Sql = Sql & "a.*, "
Sql = Sql & "b.importador, "
Sql = Sql & "C.Razao "
Sql = Sql & " from "
Sql = Sql & "tb_carga_cntr a, "
Sql = Sql & "tb_bl b, "
Sql = Sql & "tb_cad_parceiros C "
Sql = Sql & " where "
Sql = Sql & "A.BL = b.autonum "
Sql = Sql & "and b.importador = c.autonum "
Sql = Sql & "and a.marca is null "
PRSet tb, Sql, 0
While Not tb.EOF = True
    Executa "UPDATE TB_carga_cntr SET MARCA = '" & UCase(tb.Fields("RAZAO")) & "' WHERE AUTONUM = " & tb.Fields("AUTONUM") & ""
    tb.MoveNext
Wend
tb.Close
End Sub









'ADE02 - FASE 2
'==============
Public Sub Gera_Rel_Captacao(BL As Long, Apaga As Boolean)
Dim Tbh             As ADODB.Recordset
Dim tbh1            As ADODB.Recordset
Dim TBH2            As ADODB.Recordset
Dim TEXTO           As String
Dim Sql             As String
Dim Tot_CNTR20      As Long
Dim Tot_CNTR40      As Long
Dim Tot_Peso_Bruto  As Long
Dim Linha           As Integer
Dim Cod_Usuario     As Long
Cod_Usuario = tbSenha!AUTONUM
Linha = Empty
Tot_CNTR20 = Empty
Tot_CNTR40 = Empty
Tot_Peso_Bruto = Empty
Sql = Empty
TEXTO = Empty
Sql = "DELETE FROM " & BD_I & ".TEMP_REL_CAP WHERE " & BD_I & ".TEMP_REL_CAP.uSUARIO =" & Val(Cod_Usuario) & ""
If Apaga Then Executa (Sql)

'-------------------------------------- Container  -------------------------------------
'IF BANCO_EM_USO = "ORACLE" Then
'    Sql = "SELECT A.ID_CONTEINER, A.TARA, A.TAMANHO, A.REGIME, C.BRUTO, A.LACRE_ORIGEM, A.TEMP, A.ESCALA, A.IMO1, A.UNDG, B.DESCR, A.PATIO, A.DT, A.DTA, A.GMCI, A.VIAGEM FROM " & BD_I & ".TB_CNTR_BL A, " & BD_I & ".DTE_TB_TIPOS_CONTEINER B, " & BD_I & ".TB_AMR_CNTR_BL C WHERE A.AUTONUM = C.CNTR AND A.TIPO = B.CODE(+) AND C.BL = '" & BL & "' ORDER BY A.ID_CONTEINER"
'Else
    Sql = "SELECT A.ID_CONTEINER, A.TARA, A.TAMANHO, A.REGIME, C.BRUTO, A.LACRE_ORIGEM, A.TEMP, A.ESCALA, A.IMO1, A.UNDG, B.DESCR, A.PATIO, A.DT"
    Sql = Sql & ", A.DTA, A.GMCI, A.VIAGEM FROM " & BD_I & ".TB_CNTR_BL A"
    Sql = Sql & " inner join " & BD_I & ".TB_AMR_CNTR_BL C on A.AUTONUM = C.CNTR"
    Sql = Sql & " left join " & BD_I & ".DTE_TB_TIPOS_CONTEINER B on A.TIPO = B.CODE"
    Sql = Sql & " WHERE C.BL = '" & BL & "' ORDER BY A.ID_CONTEINER "
'End If
PRSet Tbh, Sql
If Not Tbh.EOF = True Then
    'ADE02
    If IsNull(Tbh!Dta) = True Then
        REL_GMCI_GRCI Nnull(Tbh!DT, 1), Nnull(Tbh!GMCI, 1), Tbh!Viagem, 1 'GERA COMO GMCI
    Else
        REL_GMCI_GRCI Nnull(Tbh!Dta, 1), Nnull(Tbh!GMCI, 1), Tbh!Viagem, 2 'GERA COMO GRCI
    End If
    '--
    Linha = Linha + 1
    TEXTO = "CONTEINERES"
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = " "
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = "-------------------------------------------------------------------------------------------"
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = Left$("SIGLA" & s(12), 12) & "|" & Left$("TIPO" & s(14), 14) & "|" & Left$("TARA" & s(5), 5) & "|"
    TEXTO = TEXTO & Left$("TAM." & s(4), 4) & "|" & Left$("REG." & s(3), 3) & "|"
    TEXTO = TEXTO & Left$("P.BRUTO." & s(7), 7) & "|" & Left$("LACRE" & s(11), 11) & "|"
    TEXTO = TEXTO & Left$("TEMP." & s(7), 7) & "|" & Left$("C.PERIG." & s(8), 8)
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = "------------+--------------------+----+---+-------+-----------+-------+--------------------"
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    While Not Tbh.EOF = True
        Linha = Linha + 1
        TEXTO = Left$(Nnull(Tbh!Id_Conteiner, 1) & s(12), 12) & "|" & Left$(Nnull(Tbh!Descr, 1) & s(14), 14) & "|" & Left$(Nnull(Tbh!Tara, 1) & s(5), 5) & "|"
        TEXTO = TEXTO & Left$(Nnull(Tbh!Tamanho, 0) & s(4), 4) & "|" & Left$(Nnull(Tbh!Regime, 1) & s(3), 3) & "|"
        TEXTO = TEXTO & Right$(s(7) & Nnull(Tbh!Bruto, 0), 7) & "|" & Left$(Nnull(Tbh!lacre_origem, 1) & s(11), 11) & "|"
        TEXTO = TEXTO & Left$(Nnull(Tbh!TEMP, 1) & s(5), 5) & " " & Left$(Nnull(Tbh!Escala, 1) & s(1), 1) & "|"
        TEXTO = TEXTO & Left$(Nnull(Tbh!IMO1, 1) & s(3), 3) & " " & Left$(Nnull(Tbh!UNDG, 1) & s(4), 4)
        TEXTO = TEXTO & Left$(Nnull(Tbh!Patio, 1) & s(3), 3) & " "
        If Nnull(Tbh!Tamanho, 0) = 20 Then Tot_CNTR20 = Tot_CNTR20 + 1
        If Nnull(Tbh!Tamanho, 0) = 40 Then Tot_CNTR40 = Tot_CNTR40 + 1
        Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
        Tbh.MoveNext
    Wend
    Linha = Linha + 1
    TEXTO = " "
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = "-------------------------------------------------------------------------------------------"
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = s(20) & "Total CNTR 20...: " & Tot_CNTR20 & s(20) & "Total CNTR 40...: " & Tot_CNTR40
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = "-------------------------------------------------------------------------------------------"
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = " "
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
End If
Tbh.Close

'-------------------------------------- Carga Conteinerizada -------------------------------------

Sql = "Select  A.Item,  A.Quantidade,  A.Peso_Bruto,  A.Mercadoria,   B.Descr as Descr_Embalagem, C.Descr as Descr_Genero"
Sql = Sql & " From " & BD_I & ".TB_CARGA_CNTR A"
Sql = Sql & " left join " & BD_I & ".DTE_TB_EMBALAGENS B on A.Embalagem=B.code"
Sql = Sql & " left join " & BD_I & ".TB_CAD_GRUPO_PRODUTOS C on A.genero=C.autonum"
Sql = Sql & " Where"
Sql = Sql & " A.bl  = " & BL & " Order by a.ITEM"
PRSet tbh1, Sql
If Not tbh1.EOF = True Then
    Linha = Linha + 1
    TEXTO = "CARGA CONTEINERIZADA"
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = " "
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = "-------------------------------------------------------------------------------------------"
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = Left$("ITEM" & s(4), 4) & "|" & Left$("DESCR_EMBALAGEM" & s(20), 20) & "|"
    TEXTO = TEXTO & Left$("QUANT." & s(6), 6) & "|" & Left$("P.BRUTO " & s(8), 8) & "|"
    TEXTO = TEXTO & Left$("MERCADORIA." & s(25), 25) & "|" & Left$("DESCR_GENERO" & s(20), 20)
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = "----+--------------------+------+--------+-------------------------+-----------------------------"
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    While Not tbh1.EOF = True
        Linha = Linha + 1
        TEXTO = Left$(Nnull(tbh1!Item, 0) & s(4), 4) & "|" & Left$(Nnull(tbh1!DESCR_EMBALAGEM, 1) & s(20), 20) & "|"
        TEXTO = TEXTO & Right$(s(6) & Nnull(tbh1!Quantidade, 0), 6) & "|" & Right$(s(8) & Nnull(tbh1!PESO_BRUTO, 0), 8) & "|"
        TEXTO = TEXTO & Left$(Nnull(tbh1!Mercadoria, 1) & s(25), 25) & "|" & Left$(Nnull(tbh1!Descr_Genero, 1) & s(30), 30)
        Tot_Peso_Bruto = Tot_Peso_Bruto + Nnull(tbh1!PESO_BRUTO, 0)
        Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
        tbh1.MoveNext
    Wend
    Linha = Linha + 1
    TEXTO = " "
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = "-------------------------------------------------------------------------------------------"
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = s(20) & "Total peso Bruto ...: " & Tot_Peso_Bruto
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = "-------------------------------------------------------------------------------------------"
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = " "
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
End If
tbh1.Close

'-------------------------------------- Carga SOLTA -------------------------------------
Tot_Peso_Bruto = 0
Sql = "SELECT A.GMVI, A.ITEM, A.DTA, A.BL, A.QUANTIDADE, A.PESO_BRUTO, A.MERCADORIA, B.DESCR AS DESCR_EMBALAGEM, C.DESCR AS DESCR_GENERO"
Sql = Sql & ", A.GMVI, A.DT, A.VIAGEM, A.PATIO"
Sql = Sql & " FROM " & BD_I & ".TB_CARGA_SOLTA A"
Sql = Sql & " left join " & BD_I & ".DTE_TB_EMBALAGENS B on A.EMBALAGEM = B.CODE"
Sql = Sql & " left join " & BD_I & ".TB_CAD_GRUPO_PRODUTOS C on A.GENERO = C.AUTONUM"
Sql = Sql & " WHERE A.BL = '" & BL & "' ORDER BY A.ITEM"
PRSet tbh1, Sql
If Not tbh1.EOF = True Then
    'ADE02 - FASE 2
    If IsNull(tbh1!Dta) = True Then
        REL_GMVI tbh1!DT, tbh1!GMVI, tbh1!Viagem, 1, (Format$(Now, "YY") & Format$(tbh1!BL, "000000") & "0")
    Else
        REL_GMVI tbh1!DT, tbh1!GMVI, tbh1!Viagem, 2, (Format$(Now, "YY") & Format$(tbh1!BL, "000000") & "0")
    End If
    '--
    Linha = Linha + 1
    TEXTO = "CARGA SOLTA"
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = " "
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = "-------------------------------------------------------------------------------------------"
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = Left$("ITEM" & s(4), 4) & "|" & Left$("DESCR_EMBALAGEM" & s(20), 20) & "|"
    TEXTO = TEXTO & Left$("QUANT." & s(6), 6) & "|" & Left$("P.BRUTO " & s(8), 8) & "|"
    TEXTO = TEXTO & Left$("MERCADORIA." & s(25), 25) & "|" & Left$("DESCR_GENERO" & s(20), 20)
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = "----+--------------------+------+--------+-------------------------+-----------------------------"
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    While Not tbh1.EOF = True
        Linha = Linha + 1
        TEXTO = Left$(Nnull(tbh1!Item, 0) & s(4), 4) & "|" & Left$(Nnull(tbh1!DESCR_EMBALAGEM, 1) & s(20), 20) & "|"
        TEXTO = TEXTO & Right$(s(6) & Nnull(tbh1!Quantidade, 0), 6) & "|" & Right$(s(8) & Nnull(tbh1!PESO_BRUTO, 0), 8) & "|"
        TEXTO = TEXTO & Left$(Nnull(tbh1!Mercadoria, 1) & s(25), 25) & "|" & Left$(Nnull(tbh1!Descr_Genero, 1) & s(30), 30)
        TEXTO = TEXTO & Nnull(tbh1!Patio, 1) & " "
        Tot_Peso_Bruto = Tot_Peso_Bruto + Nnull(tbh1!PESO_BRUTO, 0)
        Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
        tbh1.MoveNext
    Wend
    Linha = Linha + 1
    TEXTO = " "
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = "-------------------------------------------------------------------------------------------"
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = s(20) & "Total peso Bruto ...: " & Tot_Peso_Bruto
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = "-------------------------------------------------------------------------------------------"
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = " "
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
End If
tbh1.Close

'-------------------------------------- Veículos -------------------------------------
PRSet TBH2, "Select A.Chassi from " & BD_I & ".TB_VEICULOS_BL A Where a.BL = " & BL & " Order by a.Chassi"
If Not TBH2.EOF = True Then
    Linha = Linha + 1
    TEXTO = "VEÍCULOS"
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = " "
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = "-------------------------------------------------------------------------------------------"
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = Left$("CHASSIS" & s(20), 20)
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = "-------------------------------------------------------------------------------------------"
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    While Not TBH2.EOF = True
        Linha = Linha + 1
        TEXTO = Left$(Nnull(TBH2!CHASSI, 1) & s(20), 20)
        Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
        TBH2.MoveNext
    Wend
    Linha = Linha + 1
    TEXTO = "-------------------------------------------------------------------------------------------"
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
    Linha = Linha + 1
    TEXTO = " "
    Call Acrescenta_Linha_Relatorio_Captacao(BL, Linha, TEXTO, Cod_Usuario)
End If
TBH2.Close
Set Tbh = Nothing
Set tbh1 = Nothing
Set TBH2 = Nothing
TEXTO = Empty
Sql = Empty
Tot_CNTR20 = Empty
Tot_CNTR40 = Empty
Tot_Peso_Bruto = Empty
Linha = Empty
End Sub

Public Sub Acrescenta_Linha_Relatorio_Captacao(ByVal Campo1 As Long, ByVal Campo2 As Integer, ByVal Campo3 As String, ByVal Campo4 As Long)
Dim Sql As String
Sql = Empty
Sql = Sql & " INSERT INTO "
Sql = Sql & " " & BD_I & ".TEMP_REL_CAP ("
Sql = Sql & " bl, "
Sql = Sql & " linha, "
Sql = Sql & " texto, "
Sql = Sql & " usuario "
Sql = Sql & " ) values ( "
Sql = Sql & " '" & Nnull(Campo1, 0) & "',"
Sql = Sql & " '" & Nnull(Campo2, 0) & "',"
Sql = Sql & " '" & Nnull(Campo3, 1) & "',"
Sql = Sql & " '" & Nnull(Campo4, 1) & "' "
Sql = Sql & " ) "
Executa Sql
Sql = Empty
End Sub

Public Function Rel_GRCI_GMCI(ByVal DT As String, ByVal GMCI As String, ByVal Viagem As String, ByVal Tipo As String)
Dim Sql As String
Sql = Empty
Sql = Sql & " INSERT INTO "
Sql = Sql & " " & BD_I & ".TB_EMISSAO_GUIAS ( "
Sql = Sql & " GUIA, "
Sql = Sql & " TIPO, "
Sql = Sql & " DT_EMISSAO, "
Sql = Sql & " USUARIO "
Sql = Sql & " ) VALUES ( "
Sql = Sql & " '" & Left$(Nnull(GMCI, 1) & s(13), 13) & "', "
Sql = Sql & " 'G" & IIf(Tipo = 1, "M", "R") & "CI', "
Sql = Sql & " SYSDATE, "
Sql = Sql & " '" & tbSenha!AUTONUM & "' "
Sql = Sql & " ) "
Executa Sql
Sql = Empty
End Function

Public Sub Acrescenta_Linha_Relatorio_Dt(ByVal Campo1 As String, ByVal Campo2 As Integer, ByVal Campo3 As String, ByVal Campo4 As Long)
Dim Sql As String
Sql = "INSERT INTO " & BD_I & ".TEMP_REL_DT (dt,linha,texto,usuario) values "
Sql = Sql & "('" & Nnull(Campo1, 1) & "'"
Sql = Sql & "," & Nnull(Campo2, 0) & ""
Sql = Sql & ",'" & Nnull(Campo3, 1) & "'"
Sql = Sql & "," & Nnull(Campo4, 1) & ")"
Executa (Sql)
Sql = Empty
End Sub

Public Sub Acrescenta_Linha_Relatorio_Servicos(ByVal Campo1 As Long, ByVal Campo2 As Integer, ByVal Campo3 As String, ByVal Campo4 As Long)
Sql = "INSERT INTO " & BD_I & ".TEMP_REL_SERVICOS (Lista_Preco, linha, texto,usuario) values "
Sql = Sql & "(" & Nnull(Campo1, 0) & ""
Sql = Sql & "," & Nnull(Campo2, 0) & ""
Sql = Sql & ",'" & Nnull(Campo3, 1) & "'"
Sql = Sql & "," & Nnull(Campo4, 1) & ")"
Executa (Sql)
End Sub
