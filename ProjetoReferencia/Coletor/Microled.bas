Attribute VB_Name = "Microled"

Public Declare Function CreateFieldDefinitionFile Lib "P2smon.dll" Alias "CreateFieldDefFile" (lpUnk As Object, ByVal sFile As String, ByVal bOverWrite As Long) As Long
Public Declare Function CreateReportOnRuntime Lib "P2smon.dll" Alias "CreateReportOnRuntimeDS" (lpUnk As Object, ByVal sFile As String, ByVal sDefFile As String, ByVal bOverWrite As Long, ByVal bLaunchDesigner As Long) As Long
Public ModoRepesagem As Byte

Public ProFileUsuario As String
Public Banco_Session As String
Public INSERE_INDEX As String
Public CodEmpSis As String
Public ErroLeitura As Boolean
Public InicioFila As Integer
Public WSQL As String
Public DiretoImpressora As Boolean
Public ExeLeitura As Boolean
Public Flag_Pacote_Local As Byte
Public LOTE_PRECALCULO_ATENDIMENTO As Long
Public BaseDeTestes As Boolean
Public Ultimo_Robo As Date
Public Homologando As Boolean
Public Trimestre As Byte
Public NFSe As Boolean
Public Processo_Leitura As String
Public EmailDesignar As String
Public NumeroNota As String
Public Tem_Sid As String
Public NotaNFE As String
Public Rs As New ADODB.Recordset
Public Banco_Cliente As String
Public Banco_Sgipa As String
Public Banco_Operador As String
Public Banco_Redex As String
Public Banco_Fatura As String
Public Banco_TECTER As String
Public EnvioPacote17 As Boolean
Public DataServidor As Date
Public msgIcms As String
Public msgdoc As String
Public CntrVazio  As Long
Public LoteDTA As Long
Public db               As ADODB.Connection
Public DbTeste          As ADODB.Connection
Public ADOTableRpt      As New ADODB.Recordset
Public Rst              As New ADODB.Recordset
Public tbSenha          As New ADODB.Recordset
Public iLogo            As StdPicture
Public TipoCarga        As Byte
Public Destino_Relatorio As Byte
Public Nome_Destino_Relatorio As String
Public BD_O As String
Public BD_I As String
Public Banco_Em_Uso As String
Public Erro_SID As String
Public Tem_ferrovia As Boolean
Public N_Process(6) As Long
Public Tipo_Yard As String
Public Autonum_Yard As Long
Public Msg_Bdcc As String
Public Bloqueio_BDCC As Boolean
Public Viagem_XML As String

'declaração de variáveis do tipo string
Public ConnectionString As String
Public ConnectionStringAdo As String

Public Numero_Envio_46  As String
Public MatrizCores()    As String
Public wModelo          As String    'Carrega a configuração das portas seriais das balanças desde o arquivo de inicialização (linha 2).
Public sEsquema         As String    'Indica em qual esquema o usuário deve ser buscado (SGIPA ou OPERADOR, por enquanto).
Public Assunto          As String
Public Connect          As String
Public Connect2         As String
Public CorpoMensagem    As String
Public Digito1          As String
Public Digito2          As String
Public Drv              As String
Public DSN              As String
Public FiltroTela       As String
Public Maquina_Rede     As String
Public Motivo           As String
Public Para             As String
Public SQl              As String
Public sSql              As String
Public strSQL As String
Public IdVend_Comissao  As Double
Public CepCons          As String
Public Base_Sap         As String
Public LoteObs          As String
Public BlDocCanc        As Double
Public Num_Nota         As String
Public Retorno          As String
Public RPS              As Boolean
Public StrInstruction   As String
Public Resp_Exec_XML    As String
Public RsParametro      As New ADODB.Recordset
Public SqlT             As String
Public Usuario_Rede     As String
Public Usuario_Sistema  As String
Public wBLInformado     As String
Public wCampo           As String
Public wDataInformada   As String
Public wDtFim           As String
Public wDtInicio        As String
Public wNomeEDI         As String
Public wPODtri          As String
Public wPOLtri          As String
Public wTipoEmail       As String
Public wTipoEnvio       As String
Public wUsuarioBase     As String
Public wVersao          As String
Public sLicenciado      As String
Public P_XML_Evento     As String
Public P_XML_Bl         As String
Public P_XML_Id_Cntr    As String
Public P_XML_Viagem     As String
Public P_XML_Autonum_cntr As String
Public RsAudit As New ADODB.Recordset
'declaração de variáveis do tipo boolean
Public GateManual       As Boolean   'Determina se será possivel realizar pesagens manualmente - tbsenha!flag_balanca_manual
Public AtualizaCif      As Boolean
Public BaseTeste        As Boolean
Public Desconsolidando  As Boolean
Public Inicio           As Boolean
Public wAmbos           As Boolean
Public wAtendimento     As Boolean
Public wCancel          As Boolean
Public wCancela_Envio    As Boolean
Public wSoEntregue      As Boolean
Public wSoNoTerminal    As Boolean
'declaração de variáveis do tipo integer
Public P_Cntr           As Long 'Variavel que irá guardar o autonum da TB_CNTR_BL (XML)
Public P_BL             As Long 'Variavel que irá guardar o autonum da TB_BL (XML)
Public wVerificaVersao  As Integer
Public wTipoRel         As Integer
Public TipodeBanco      As Integer
Public nPATIO           As Integer   'Carrega o pátio desde o arquivo de inicialização (linha 1).
'declaração de variáveis do tipo long
Public UltimoAutonum    As Long
Public Cod_Usuario      As Long      'Código do usuário no cadastro.
Public nUser            As Long
Public Cod_Empresa      As Long      'Código da empresa do usuario.
Public Cod_Empresa_Patio As String ' faixa de patios validos para empresa
Public Cod_Terminal As String
Public p_in             As String  ' instrucao in sem o campo para patios validos
Public Tabela_Padrao As Long
Public Razao_Empresa As String
Public wCpf_Usuario As String
Public Wnome_Usuario As String

'declaração de variáveis do tipo byte
Public nSplash          As Byte

Public TipoChamada As Integer

Public SAPUSU As String
Public SAPSEN As String
Public AlfaWeb As String
Public FAT_PAR As New ADODB.Recordset
Public AlfaUsu As String

Public MaXi As Integer

Public TrabDsn As String
Public UsuBanco As String
Public PasBanco As String

 
Public IdProt As Integer

Public SqlQuery$
Public TipoRel As Integer
Private Const MAXLEN_PHYSADDR As Long = 8

Public Const Dominios As String = "AERO BIZ COM COOP EDU GOV INFO INT MIL MUSEUM NAME NET ORG PRO " & _
                                    "AC AD AE AF AG AI AL AM AN AO AQ AR AS AT AU AW AZ BA BB BD " & _
                                    "BE BF BG BH BI BJ BM BN BO BR BS BT BV BW BY BZ CA CC CD CF " & _
                                    "CG CH CI CK CL CM CN CO CR CU CV CX CY CZ DE DJ DK DM DO DZ " & _
                                    "EC EE EG EH ER ES ET FI FJ FK FM FO FR GA GD GE GF GG GH GI " & _
                                    "GL GM GN GP GQ GR GS GT GU GW GY HK HM HN HR HT HU ID IE IL " & _
                                    "IM IN IO IQ IR IS IT JE JM JO JP KE KG KH KI KM KN KP KR KW " & _
                                    "KY KZ LA LB LC LI LK LR LS LT LU LV LY MA MC MD MG MH MK ML " & _
                                    "MM MN MO MP MQ MR MS MT MU MV MW MX MY MZ NA NC NE NF NG NI " & _
                                    "NL NO NP NR NU NZ OM PA PE PF PG PH PK PL PM PN PR PS PT PW " & _
                                    "PY QA RE RO RU RW SA SB SC SD SE SG SH SI SJ SK SL SM SN SO " & _
                                    "SR ST SV SY SZ TC TD TF TG TH TJ TK TM TN TO TP TR TT TV TW " & _
                                    "TZ UA UG UK UM US UY UZ VA VC VE VG VI VN VU WF WS YE YT YU " & _
                                    "ZA ZM ZW"

Private Type MIB_IPNETROW
    dwIndex                             As Long
    dwPhysAddrLen                       As Long
    bPhysAddr(0 To MAXLEN_PHYSADDR - 1) As Byte
    dwAddr                              As Long
    dwType                              As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Enum TpDados
    [0] ' numero
    [1] ' texto
    [2] ' data
End Enum

Public Enum eBCORES
    [AMARELO] = &HE3F9FF
    [VERDE] = &HC8FFD9
    [VERMELHO] = &H8EA6D2
    [BRANCO] = vbWhite
End Enum

Public Enum eDIREITOS
    eACESSAR = 0
    eINCLUIR = 1
    eALTERAR = 2
    eEXCLUIR = 3
    eSALVAR = 4
    eLOCALIZAR = 5
    eFILTRAR = 6
    eLISTAR = 7
    eIMPRIMIR = 8
    eENVIAR_PACOTES = 9
    eCANCELAR_PACOTES = 10
    eAUDITAR = 11
End Enum

'declaração de funções da API
Private Declare Function GetIpNetTable Lib "Iphlpapi" (pIpNetTable As Byte, pdwSize As Long, ByVal bOrder As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function FindWindowEx& Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetParent& Lib "user32" (ByVal hWnd As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

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


Function Consiste_Temp(Temperatura) As String
Consiste_Temp = Replace(Temperatura, ",", ".")
Consiste_Temp = Replace(Temperatura, "+", "")
If Nnull(Temperatura, 1) <> "" Then
    If Not IsNumeric(Temperatura) Then
        Consiste_Temp = ""
    Else
        'Consiste_Temp = Val(Temperatura)
        Consiste_Temp = Replace(Consiste_Temp, ",", ".")
    End If
Else
    Consiste_Temp = ""
End If

End Function

Function Consiste_Umidade(Umidade) As String

If Nnull(Umidade, 1) <> "" Then
    If Not IsNumeric(Umidade) Then
        Consiste_Umidade = "OFF"
    Else
        If Nnull(Umidade, 1) <> "0" Then
            Consiste_Umidade = Val(Umidade)
        Else
            Consiste_Umidade = "OFF"
        End If
    End If
Else
    Consiste_Umidade = "OFF"
End If
Consiste_Umidade = Replace(Consiste_Umidade, ",", ".")
End Function

Function Consiste_Ventilacao(Ventilacao) As String

If Nnull(Ventilacao, 1) <> "" Then
    If Not IsNumeric(Ventilacao) Then
        Consiste_Ventilacao = "CLOSED"
    Else
        If Nnull(Ventilacao, 1) <> "0" Then
            Consiste_Ventilacao = Val(Ventilacao)
        Else
            Consiste_Ventilacao = "CLOSED"
        End If
    End If
Else
    Consiste_Ventilacao = "CLOSED"
End If
End Function

Public Sub UnloadAllForms()
Dim frm As Form
For Each frm In Forms
Unload frm
Next
End Sub

Public Function Ver_Bloqueio_Hubport(Lote As String) As Boolean

Dim Bloq As Byte
Bloq = 1
Ver_Bloqueio_Hubport = False

SQl = "select bl.captador,"
SQl = SQl & " case when bl2.grau<>'M' then bl2.importador else 0 end as coloader1,"
SQl = SQl & " case when bl3.grau<>'M' then bl3.importador else 0 end as coloader2,"
SQl = SQl & " case when bl4.grau<>'M' then bl4.importador else 0 end as coloader3"
SQl = SQl & " ,bl.importador"
SQl = SQl & " from sgipa.dbo.tb_bl bl"
SQl = SQl & " left join sgipa.dbo.tb_bl bl2 on bl.bl_master=bl2.autonum"
SQl = SQl & " left join sgipa.dbo.tb_bl bl3 on bl2.bl_master=bl3.autonum"
SQl = SQl & " left join sgipa.dbo.tb_bl bl4 on bl3.bl_master=bl4.autonum"
SQl = SQl & " left join sgipa.dbo.tb_bl bl5 on bl4.bl_master=bl5.autonum"
SQl = SQl & " where bl.autonum=" & Lote
Dim RsB As New ADODB.Recordset
PRSet RsB, SQl
SQl = "select count(*) from sgipa.dbo.tb_bloqueio_hubport where indicador=" & RsB!Captador & " and isnull(coloader,0)=0 and isnull(importador,0)=0 and ativo=1"
If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then Bloq = 0
If Bloq <> 0 Then
    SQl = "select count(*) from sgipa.dbo.tb_bloqueio_hubport where indicador=" & RsB!Captador & " and isnull(coloader,0)=0 and importador=" & RsB!Importador & " and ativo=1"
    If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then Bloq = 0
End If

If Bloq <> 0 And Nnull(RsB!coloader1, 0) <> 0 Then
    SQl = "select count(*) from sgipa.dbo.tb_bloqueio_hubport where indicador=" & RsB!Captador & " and isnull(importador,0)=0 and coloader = " & RsB!coloader1 & " and ativo=1"
    If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then Bloq = 0
End If
If Bloq <> 0 And Nnull(RsB!coloader1, 0) <> 0 Then
    SQl = "select count(*) from sgipa.dbo.tb_bloqueio_hubport where indicador=" & RsB!Captador & " and importador=" & RsB!Importador & " and coloader = " & RsB!coloader1 & " and ativo=1"
    If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then Bloq = 0
End If

If Bloq <> 0 And Nnull(RsB!coloader2, 0) <> 0 Then
    SQl = "select count(*) from sgipa.dbo.tb_bloqueio_hubport where indicador=" & RsB!Captador & " and isnull(importador,0)=0 and coloader = " & RsB!coloader2 & " and ativo=1"
    If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then Bloq = 0
End If
If Bloq <> 0 And Nnull(RsB!coloader2, 0) <> 0 Then
    SQl = "select count(*) from sgipa.dbo.tb_bloqueio_hubport where indicador=" & RsB!Captador & " and importador=" & RsB!Importador & " and coloader = " & RsB!coloader2 & " and ativo=1"
    If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then Bloq = 0
End If

If Bloq <> 0 And Nnull(RsB!coloader3, 0) <> 0 Then
    SQl = "select count(*) from sgipa.dbo.tb_bloqueio_hubport where indicador=" & RsB!Captador & " and isnull(importador,0)=0 and coloader = " & RsB!coloader3 & " and ativo=1"
    If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then Bloq = 0
End If
If Bloq <> 0 And Nnull(RsB!coloader3, 0) <> 0 Then
    SQl = "select count(*) from sgipa.dbo.tb_bloqueio_hubport where indicador=" & RsB!Captador & " and importador=" & RsB!Importador & " and coloader = " & RsB!coloader3 & " and ativo=1"
    If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then Bloq = 0
End If

If Bloq = 0 Then Ver_Bloqueio_Hubport = True


End Function

Public Function Tolerancia_Peso(CNTR As Long) As Double
    
    Dim Tolera As Double
    Set tb_t = New ADODB.Recordset
    tb_t.Open "SELECT TOLERANCIA_BAL FROM TB_PARAMETROS", db, 3, 3
    Tolera = Nnull(tb_t!tolerancia_bal, 0)
    
    Dim Pesobl As Long
    Dim PB As Long
    Dim PTara As Long
    Dim Ptc As Long
    Dim Percent As Double
    Percent = 0
    
    Pesobl = 0
    PB = 0
    PTara = 0
    Ptc = 0
    Set Tb2 = New ADODB.Recordset
    Tb2.Open "Select Sum(a.Bruto) as Quanto from tb_amr_cntr_bl a,tb_bl b where a.bl=b.autonum and b.flag_ativo=1 and a.cntr=" & CNTR & "", db, 3, 3
    If Not Tb2.EOF Then
        Pesobl = CLng(Nnull(Tb2!Quanto, 0))
    End If
    
    SQl = "select tara from sgipa.dbo.tb_cntr_bl where autonum=" & CNTR
    Ptc = Nnull(db.Execute(SQl).Fields(0), 0)
    
    SQl = "select * from operador.dbo.tb_amr_gate where cntr_ipa=" & CNTR & " and funcao_gate=11"
    Dim RsG As New ADODB.Recordset
    PRSet RsG, SQl
    If Not RsG.EOF Then
        PB = Nnull(RsG!Peso_Entrada, 0)
        PTara = Nnull(RsG!Peso_Saida, 0)
        
        If (100 * (Pesobl / (PB - PTara - Ptc))) > (100 + Tolera) Or (100 * (Pesobl / (PB - PTara - Ptc))) < (100 - Tolera) Then
            Percent = Round(Abs((100 * (Pesobl / (PB - PTara - Ptc))) - (100 + Tolera)), 4)
        End If
    End If
    RsG.Close

    Tolerancia_Peso = Percent
    
End Function

Sub Atualiza_Robo()
    If DateDiff("n", Ultimo_Robo, Now) > 1 Then
        Ultimo_Robo = Now
        Executa "Update operador.dbo.tb_robos set data_prog=GETDATE() where programa='LEITURA_PACOTES.EXE' and parametro_exe=" & Tipo_leitura
    End If
End Sub



Public Function Formata_CGCCPF(Doc As String, Tipo As Integer) As String
'TIPO = 1 - CGC / 2 - CPF
Dim Base As String
Dim Strs As String

Formata_CGCCPF = Nnull(Doc, 1)
Tipo = Nnull(Tipo, 0)

If Tipo = 1 Then
    Formata_CGCCPF = String(14, "0") & Nnull(Doc, 0)
    Formata_CGCCPF = Right(Formata_CGCCPF, 14)
    Formata_CGCCPF = Mid(Formata_CGCCPF, 1, 2) & "." & Mid(Formata_CGCCPF, 3, 3) & "." & Mid(Formata_CGCCPF, 6, 3) & "/" & Mid(Formata_CGCCPF, 9, 4) & "-" & Mid(Formata_CGCCPF, 13, 2)
    Exit Function
ElseIf Tipo = 2 Then
    Formata_CGCCPF = String(11, "0") & Nnull(Doc, 0)
    Formata_CGCCPF = Right(Formata_CGCCPF, 11)
    Formata_CGCCPF = Mid(Formata_CGCCPF, 1, 3) & "." & Mid(Formata_CGCCPF, 4, 3) & "." & Mid(Formata_CGCCPF, 7, 3) & "-" & Mid(Formata_CGCCPF, 10, 2)
    Exit Function
Else
    Exit Function
End If

End Function

Public Function Tira_Salto(Valor As String) As String

Dim RET As String
RET = ""
For i = 1 To Len(Valor)
    If Asc(Mid(Valor, i, 1)) <> 10 And Asc(Mid(Valor, i, 1)) <> 13 Then
        RET = RET & Mid(Valor, i, 1)
    Else
        If Asc(Mid(Valor, i, 1)) = 13 Then
            RET = RET & " "
        End If
    End If
Next
    
Tira_Salto = RET

End Function
Public Function Dir_Backup_Local(Optional Esquema As String = "IPA") As String
    Dim oRS As New ADODB.Recordset
    If Trim$(UCase$(Esquema)) = "OPERADOR" Then
        PRSet oRS, "SELECT NVL(DIR_SAIDA_local,'C:\') AS SAIDA, NVL(DIR_BACKUP_local,'C:\BKP\') AS BKP FROM " & Banco_Operador & "DTE_TB_PARAMETROS_SISTEMA"
    Else
        PRSet oRS, "SELECT NVL(DIR_SAIDA_local,'C:\') AS SAIDA, NVL(DIR_BACKUP_local,'C:\BKP\') AS BKP FROM " & Banco_Sgipa & "DTE_TB_PARAMETROS"
    End If
    Dir_Backup_Local = Nnull(oRS!bkp, 1)
    Dir_Backup_Local = Dir_Backup_Local & IIf(Right$(Dir_Backup_Local, 1) = "\", Empty, "\")
End Function

Public Function Dir_Saida_Local(Optional Esquema As String = "IPA") As String
    Dim oRS As New ADODB.Recordset
    If UCase(Esquema) = "OPERADOR" Then
        PRSet oRS, "SELECT NVL(DIR_SAIDA,'C:\') AS SAIDA, NVL(DIR_BACKUP,'C:\BKP\') AS BKP FROM " & Banco_Operador & "DTE_TB_PARAMETROS_SISTEMA"
    Else
        PRSet oRS, "SELECT NVL(DIR_SAIDA_local,'C:\') AS SAIDA, NVL(DIR_BACKUP_local,'C:\BKP\') AS BKP FROM " & Banco_Sgipa & "DTE_TB_PARAMETROS"
    End If
    Dir_Saida_Local = Nnull(oRS!Saida, 1)
    Dir_Saida_Local = Dir_Saida_Local & IIf(Right$(Dir_Saida_Local, 1) = "\", Empty, "\")
End Function

Public Function FDataServidor() As Date

    If Banco_Em_Uso = "ORACLE" Then
        FDataServidor = db.Execute("select GETDATE() from dual").Fields(0)
    Else
        FDataServidor = db.Execute("select getdate()").Fields(0)
    End If

End Function

Public Function Proximo_id(BaseDeDados As String, Nome_Sequencia As String, Optional Proxima_Sequencias As Boolean = False) As Long

        Dim pSQL As String
        If Proxima_Sequencias Then
            pSQL = "SELECT count(*) from " & BaseDeDados & "sysobjects where name='" & Nome_Sequencia & "' and type='U'"
            If Nnull(db.Execute(pSQL).Fields(0), 0) = 0 Then
                pSQL = "create table " & BaseDeDados & Nome_Sequencia & " (NEXTVAL  NUMERIC(6) IDENTITY(1,1),Data datetime)"
                db.Execute pSQL
            End If
            pSQL = "INSERT INTO " & BaseDeDados & Nome_Sequencia & " (Data) values (getdate())"
            db.Execute pSQL
            pSQL = "SELECT IDENT_CURRENT('" & BaseDeDados & Nome_Sequencia & "') AS ID"
            Proximo_id = Nnull(db.Execute(pSQL).Fields(0), 0)
        Else
            pSQL = "SELECT IDENT_CURRENT('" & BaseDeDados & Nome_Sequencia & "') AS ID"
            Proximo_id = Nnull(db.Execute(pSQL).Fields(0), 0)
        End If
        
        
End Function

Public Function CONVERTE_NEXTVAL(TEXTO As String, Optional TEM_INDEX As Integer) As String

If Banco_Em_Uso = "SQLSERVER" Then

Dim rstemp As New ADODB.Recordset
Dim RSNEXT As New ADODB.Recordset
Dim RSMAX As New ADODB.Recordset
Dim banco(4) As String
banco(1) = "" & Banco_Sgipa '& "DBO."
banco(2) = "" & Banco_Operador '& "DBO."
banco(3) = "REDEX.DBO."
banco(4) = "FATURA.DBO."

'TEXTO = ""
' Sql = "INSERT INTO " & Banco_Sgipa & "tb_empresas (autonum,razao_social,nome_fantasia,LIXO) VALUES "
'      Sql = Sql & "(" & Banco_Sgipa & "SEQ_EMPRESAS.NEXTVAL ,'TESTE','(00)0000-0000','A')"
'TEXTO = UCase(Sql)

NOVASTRING = UCase(TEXTO)
'MsgBox NOVASTRING
For i = 1 To 4
   If InStr(UCase(TEXTO), banco(i)) Then
    BaseDeDados = banco(i)
    'MsgBox banco(I)
    Exit For
   End If
Next i

'If InStr(UCase(TEXTO), "CREATE TABLE") <> 0 Then
'    CONVERTE_NEXTVAL = TEXTO
 '   Exit Function
'End If

If InStr(UCase(TEXTO), "NEXTVAL") <> 0 Then
    
        NOVASTRING = Replace(UCase(TEXTO), ".NEXTVAL", "")
        FIM_CAMPO_NEXT = InStr(1, UCase(TEXTO), ".NEXTVAL")
            
        
'MsgBox TEXTO
        COMECO_CAMPO_NEXT = Left(TEXTO, FIM_CAMPO_NEXT - 1)
           
        POS_NEXT = InStrRev(COMECO_CAMPO_NEXT, "(") + 1
        CAMPO_NEXT = LTrim(Mid(COMECO_CAMPO_NEXT, POS_NEXT))
        
        INI_SEQUENCIA = InStrRev(CAMPO_NEXT, ".")
        Nome_Sequencia = Right(CAMPO_NEXT, Len(CAMPO_NEXT) - INI_SEQUENCIA)
            
            If InStr(UCase(TEXTO), "UPDATE") <> 0 Then
                POS2 = InStr(1, FRASE2, ".NEXTVAL") + 7
            Else
                POS2 = InStr(1, FRASE2, ",")
            End If
        
        pos1 = InStrRev(NOVASTRING, "(", FIM_CAMPO_NEXT)
        FRASE2 = Mid(NOVASTRING, pos1 + 1)
            
            If InStr(UCase(TEXTO), "UPDATE") <> 0 Then
                POS2 = InStr(1, FRASE2, ".NEXTVAL") + 7
            Else
                POS2 = InStr(1, FRASE2, ",")
            End If
        
        FRASE2 = Mid(FRASE2, POS2)
        FRASE1 = Left(NOVASTRING, pos1)

        If InStr(UCase(TEXTO), "UPDATE") <> 0 Then
            pos1 = InStr(1, NOVASTRING, "(")
            POS2 = InStr(pos1, NOVASTRING, ",")
        Else
            pos1 = InStr(1, NOVASTRING, "(")
            POS2 = InStr(pos1, NOVASTRING, ",")
        End If
        
        
        
        Campoid = Mid(NOVASTRING, pos1 + 1, POS2 - pos1 - 1)
        
        POS3 = InStr(1, NOVASTRING, "(")
        FRASE3 = Mid(NOVASTRING, 1, POS3)
        pos4 = InStr(POS3, NOVASTRING, ",")
        FRASE4 = Mid(NOVASTRING, pos4 - 1)
                
        X_TABELA = Mid(NOVASTRING, 1, pos1 - 1)
        If InStrRev(X_TABELA, ".") <> 0 Then
            POS_TAB1 = InStrRev(X_TABELA, ".")
        Else
           X_TABELA = RTrim(X_TABELA)
         POS_TAB1 = InStrRev(X_TABELA, " ")
        End If

        X_TABELA = Mid(X_TABELA, POS_TAB1 + 1)
        
        If RSNEXT.State = adStateOpen Then RSNEXT.Close
            
        
        X_TABELA = LTrim(RTrim(X_TABELA))
       
       SQl = " SELECT  name, is_identity From  sys.ALL_Columns WHERE   [object_id] = object_id('" & BaseDeDados & X_TABELA & "')"
       SQl = SQl & " AND NAME='" & RTrim(LTrim(Campoid)) & "'"
         
          
        If rstemp.State = adStateOpen Then rstemp.Close
        rstemp.Open SQl, db
        
        If rstemp!IS_IDENTITY = True Then
            INSERE_INDEX = "NAO"
        Else
            INSERE_INDEX = "SIM"
            
            'NOVA ROTINA DE INSERÇAO
                
                SQl = " SELECT 1 from " & BaseDeDados & "sysobjects where name = '" & Nome_Sequencia & "'  and xtype = 'U'"
             
                If RSNEXT.State = adStateOpen Then RSNEXT.Close
                 RSNEXT.Open SQl, db
            
                If RSNEXT.EOF Then
                     SQl = " SELECT MAX(CAST(" & Campoid & " AS Int)) AS ID FROM " & BaseDeDados & X_TABELA
                     RSMAX.Open SQl, db
                     
                     If IsNull(RSMAX!id) Then
                         inicio_sequencia = 1
                     Else
                         inicio_sequencia = RSMAX!id + 1
                     End If

                    SQl = " create table " & BaseDeDados & Nome_Sequencia & " (NEXTVAL  NUMERIC(38) IDENTITY(" & inicio_sequencia & ",1))"
                    Executa SQl
                End If

                INSERTE_SQL = "INSERT INTO " & BaseDeDados & Nome_Sequencia & " DEFAULT VALUES"
                Executa INSERTE_SQL
                
                SQL_FINAL = " SELECT IDENT_CURRENT('" & BaseDeDados & Nome_Sequencia & "') AS ID"
            If RSNEXT.State = adStateOpen Then RSNEXT.Close
            RSNEXT.Open SQL_FINAL
            
            NOVO_ID = RSNEXT!id
            ' FIM ROTINA DE INSERÇAO
         
        End If
   
        'FIM NOVA VERIFICAÇÃO"
        
        If INSERE_INDEX = "SIM" Then
            NOVASTRING = FRASE1 & NOVO_ID & FRASE2
        ElseIf INSERE_INDEX = "NAO" Then
            NOVASTRING = Replace(FRASE1, Campoid & ",", "") & Right(FRASE2, Len(FRASE2) - 1)
            NOVASTRING = Replace(FRASE1, Campoid & ",", "") & Right(FRASE2, Len(FRASE2) - 1)
            NOVASTRING = NOVASTRING
        End If

Else
GoTo TEMPORARIO
    If InStr(UCase(TEXTO), "INSERT") <> 0 Then        'VERIFICA SE NEXTVAL É DE INSERÇAO
        
        TEXTO = UCase(TEXTO)
        NOVASTRING = Replace(UCase(TEXTO), ".NEXTVAL", "")
        
        comeco_campo_id = InStr(1, UCase(TEXTO), "(") + 1
        fim_campo_id = InStr(comeco_campo_id, TEXTO, ",")
        campo_id = LTrim(Mid(TEXTO, comeco_campo_id, fim_campo_id - comeco_campo_id))
        campo_id = RTrim(LTrim(campo_id))
        
        COMECO_CAMPO_NEXT = InStr(comeco_campo_id, TEXTO, "VALUES")
        COMECO_CAMPO_NEXT = InStr(COMECO_CAMPO_NEXT, TEXTO, "(") + 1
        FIM_CAMPO_NEXT = InStr(COMECO_CAMPO_NEXT, TEXTO, ",")
        CAMPO_NEXT = LTrim(Mid(TEXTO, COMECO_CAMPO_NEXT, FIM_CAMPO_NEXT - COMECO_CAMPO_NEXT))
         
        instrucao_sem_indeX = Replace(TEXTO, campo_id & " ,", "")
        instrucao_sem_indeX = Replace(instrucao_sem_indeX, campo_id & ",", "")
      
        instrucao_sem_indeX = Replace(instrucao_sem_indeX, CAMPO_NEXT & " ,", "")
        instrucao_sem_indeX = Replace(instrucao_sem_indeX, CAMPO_NEXT & ",", "")
        
        pos1 = InStrRev(NOVASTRING, "(")
        FRASE2 = Mid(NOVASTRING, pos1 + 1)
        POS2 = InStr(1, FRASE2, ",")
        FRASE2 = Mid(FRASE2, POS2)

        FRASE1 = Left(NOVASTRING, pos1)

        pos1 = InStr(1, NOVASTRING, "(")
        POS2 = InStr(pos1, NOVASTRING, ",")
        
        Campoid = Mid(NOVASTRING, pos1 + 1, POS2 - pos1 - 1)
        
        POS3 = InStr(1, NOVASTRING, "(")
        FRASE3 = Mid(NOVASTRING, 1, POS3)
        pos4 = InStr(POS3, NOVASTRING, ",")
        FRASE4 = Mid(NOVASTRING, pos4 - 1)
                
        X_TABELA = Mid(NOVASTRING, 1, pos1 - 1)
        
        If InStrRev(X_TABELA, ".") <> 0 Then
            POS_TAB1 = InStrRev(X_TABELA, ".")
        Else
            X_TABELA = RTrim(X_TABELA)
            POS_TAB1 = InStrRev(X_TABELA, " ")
        End If
        
        X_TABELA = Mid(X_TABELA, POS_TAB1 + 1)

        If RSNEXT.State = adStateOpen Then RSNEXT.Close
        SQl = " SELECT MAX(CAST(" & Campoid & " AS Int)) AS ID FROM " & BaseDeDados & X_TABELA
        RSNEXT.Open SQl, db
        
        If IsNull(RSNEXT!id) Then
            NOVO_ID = 1
        Else
            NOVO_ID = RSNEXT!id + 1
        End If
     
                    
        'VERIFICA SE TABELA JÁ POSSUI TEM CHAVE PRIMARIA
            SQl = " select COLUMN_NAME, TABLE_NAME  from INFORMATION_SCHEMA.COLUMNS"
            SQl = SQl & " where TABLE_SCHEMA = 'dbo'  AND TABLE_NAME='" & RTrim(X_TABELA) & "'"
            SQl = SQl & " AND COLUMN_NAME ='" & Campoid & "'"
            SQl = SQl & " and COLUMNPROPERTY(object_id(TABLE_NAME), COLUMN_NAME, 'IsIdentity') = 1  order by TABLE_NAME"
        
            If rstemp.State = adStateOpen Then rstemp.Close
            rstemp.Open SQl, db

            If IS_AUTO = True Then INSERE_INDEX = "NAO"
        
            If rstemp.EOF Then
                INSERE_INDEX = "SIM"
            Else
                INSERE_INDEX = "NAO"
            End If
            'FIM NOVA VERIFICAÇÃO"
            
        SQl = " SELECT  name, is_identity From sys.ALL_Columns WHERE [object_id] = object_id('" & LTrim(RTrim(X_TABELA)) & "')"
        SQl = SQl & " AND NAME='" & RTrim(LTrim(Campoid)) & "'"
           
        If rstemp.State = adStateOpen Then rstemp.Close

        rstemp.Open SQl, db
        
        If rstemp!IS_IDENTITY = True Then
            INSERE_INDEX = "NAO"
        Else
            INSERE_INDEX = "SIM"
        End If

            If INSERE_INDEX = "SIM" Then
                NOVASTRING = FRASE1 & NOVO_ID & FRASE2
            ElseIf INSERE_INDEX = "NAO" Then
                NOVASTRING = instrucao_sem_indeX
            NOVASTRING = NOVASTRING & " SELECT @@IDENTITY AS '" & Campoid & "'"

            End If
   End If
TEMPORARIO:
'MsgBox TEXTO
End If


'MsgBox NOVASTRING
    
NOVASTRING = Replace(NOVASTRING, "GETDATE()", "GETDATE()")

'MsgBox NOVASTRING
CONVERTE_NEXTVAL = NOVASTRING
Else
CONVERTE_NEXTVAL = TEXTO
End If
End Function


Public Function executa_to_date(TEXTO As String) As String



If Banco_Em_Uso = "SQLSERVER" Then

TEXTO2 = UCase(TEXTO)
nova_string = ""

Do While InStr(1, TEXTO2, "TO_DATE") <> 0 'SÓ INICIA SE POSSUIR TO_DATE
  
    Posicao = 1
    X_INICIO = InStr(Posicao, UCase(TEXTO2), "TO_DATE") 'INICIO DA INSTRUCAO TO_DATE"

'// ISOLA AS VARIÁVEIS DO CAMPO E DA FORMATAÇAO PARA TROCA
    
    INICIO_CAMPO = InStr(X_INICIO, TEXTO2, "(") 'POSICAO INICIO DO CAMPO
    FIM_CAMPO = InStr(X_INICIO, TEXTO2, ",") ' POSICAO FIM DO CAMPO

    X_INI_FORMATO = InStr(X_INICIO, TEXTO2, ",")
    X_FIM_FORMATO = InStr(X_INI_FORMATO, TEXTO2, ")")

    X_FORMATO = TEXTO2

    X_FORMATO = (Right(TEXTO2, Len(TEXTO2) - X_INI_FORMATO))
    X_FORMATO = Left(X_FORMATO, X_FIM_FORMATO - X_INI_FORMATO - 1)

    X_CAMPO = (Right(TEXTO2, Len(TEXTO2) - INICIO_CAMPO))
    X_CAMPO = Left(X_CAMPO, FIM_CAMPO - INICIO_CAMPO - 1)

'//

    FIND_STRING = Mid(TEXTO2, X_INICIO, X_FIM_FORMATO - X_INICIO + 1) 'PROCURA A INSTRUÇÃO ORACLE ANTERIOR PARA SUBSTITUIÇÃO
    
    Select Case UCase(LTrim(X_FORMATO))
                
                Case "'DD/MM/YYYY'"
                     X_CONVERTE = "CONVERT(DATETIME, " & X_CAMPO & " , 103)"
                Case "'DD/MM/YY HH24:MI:SS'"
                     X_CONVERTE = "CONVERT(DATETIME, " & X_CAMPO & ", 103)"
                Case "'DD/MM/YYYY HH24:MI'"
                    X_CONVERTE = "CONVERT(DateTime, " & X_CAMPO & ", 103)"
                Case "'DD/MM/YYYY HH24:MI:SS'"
                     X_CONVERTE = "CONVERT(DATETIME, " & X_CAMPO & ", 103)"
                Case Else
                    MsgBox "FORMATO DE DATA NÃO PREVISTO"
                              
    End Select
 
nova_string = nova_string & Replace(Left(TEXTO2, X_FIM_FORMATO), FIND_STRING, X_CONVERTE)
TEXTO2 = Right(TEXTO2, Len(TEXTO2) - X_FIM_FORMATO) 'RESTO DA STRING PARA ANALISAR
  
Loop

nova_string = nova_string & TEXTO2
executa_to_date = nova_string

End If


End Function
Public Function SELECIONA_FORMATO(X_FORMATO, X_CAMPO As String) As String

Select Case LTrim(UCase(X_FORMATO))
                
                
                Case "'MM/YYYY'"
                    X_CONVERTE = "RIGHT(CONVERT(VarChar(10), " & X_CAMPO & ", 103),7)"
                           'RIGHT(CONVERT(VARCHAR(10), GETDATE(), 103), 7)
                Case "'DD/MM/YY HH:MM'"
                     X_CONVERTE = "CONVERT(VARCHAR(8), " & X_CAMPO & ", 3) + ' ' + CONVERT(VarChar(5), " & X_CAMPO & ", 108)"
                Case "'999999.999'"
                     X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ")"
                Case "'999'"
                     X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ")"
                Case "'YYYYMMDD HH24MISS'"
                     X_CONVERTE = "CONVERT(VARCHAR(8), " & X_CAMPO & ", 112) + ' ' + REPLACE(CONVERT(VarChar(8), " & X_CAMPO & ", 114),':','')"
                Case "'MM'"
                     X_CONVERTE = "DATEPART(MM," & X_CAMPO & ")"
                Case "'99'"
                    X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ")"
                Case "'9999999990D00'"
                    X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ")"
                Case "'9999999.999'"
                    X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ")"
                
                Case "'HH24:MI:SS'"
                   ' X_CONVERTE = "CONVERT(VarChar(23), " & X_CAMPO & ", 108)"
                   X_CONVERTE = "CONVERT(VarChar(8), " & X_CAMPO & ", 114)"
                Case "'HH24MISS'"
                    'X_CONVERTE = "CONVERT(VarChar(23), " & X_CAMPO & ", 108)"
                     X_CONVERTE = "REPLACE(CONVERT(VarChar(8), " & X_CAMPO & ", 114),':','')"
                Case "'YYYY'"
                    X_CONVERTE = "DATEPART(YYYY," & X_CAMPO & ")"
                Case "'YYYYMMDD'"
                    X_CONVERTE = "CONVERT(VarChar(23), " & X_CAMPO & ", 112)"
                Case "'DDMMYYYY'"
                    'X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ")"
                    X_CONVERTE = "REPLACE(CONVERT(VARCHAR(10), " & X_CAMPO & ", 103), '/', '')"



                Case "'DD/MM/YYYY'"
                    X_CONVERTE = "CONVERT(VarChar(10), " & X_CAMPO & ", 103)"
                Case "'DD/MM/YY'"
                    X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ", 3)"
               
               Case "'DD/MM/YYYY HH24:MI:SS'"
                    X_CONVERTE = "CONVERT(VarChar(10), " & X_CAMPO & ", 103) + ' ' + CONVERT(VarChar(8), " & X_CAMPO & ", 108)"
                
               Case "'DD/MM/YYYYHH24:MI:SS'"
                    X_CONVERTE = "CONVERT(VarChar(10), " & X_CAMPO & ", 103) + CONVERT(VarChar(8), " & X_CAMPO & ", 108)"
                
                Case "'DD/MM/YY HH24:MI:SS'"
                    X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ", 3) + ' ' + CONVERT(VarChar(8), " & X_CAMPO & ", 108)"
                
                Case "'DD/MM/YY HH24:MM'"
                    'X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ", 3)"
                    X_CONVERTE = "CONVERT(VARCHAR(8), " & X_CAMPO & ", 3) + ' ' + CONVERT(VarChar(5), " & X_CAMPO & ", 108)"
                    
                Case "'DD/MM/YY HH24:MI'"
                    'X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ", 3)"
                    X_CONVERTE = "CONVERT(VARCHAR(8), " & X_CAMPO & ", 3) + ' ' + CONVERT(VarChar(5), " & X_CAMPO & ", 108)"
                    
                Case "'YYYYMM'"
                    'X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ", 3)"
                    X_CONVERTE = "CONVERT(VARCHAR(6), " & X_CAMPO & ",112)"
                    
                 Case "'DD/MM/YYYY HH24:MI'", "'DD/MM/YYYY  HH24:MI'"
                    'X_CONVERTE = "CONVERT(VarChar(23), " & X_CAMPO & ", 103)"
                     X_CONVERTE = "CONVERT(VARCHAR(10), " & X_CAMPO & ", 103) + ' ' + CONVERT(VarChar(5), " & X_CAMPO & ", 108)"
                
                Case "'DD/MM/RR HH24:MI'"
                   'X_CONVERTE = "CONVERT(VarChar(23), " & X_CAMPO & ", 103)"
                    X_CONVERTE = "CONVERT(VARCHAR(8), " & X_CAMPO & ", 3) + ' ' + CONVERT(VarChar(5), " & X_CAMPO & ", 108)"
                Case "'DD/MM/RR'"
                    'X_CONVERTE = "CONVERT(VarChar(23), " & X_CAMPO & ", 103)"
                     X_CONVERTE = "CONVERT(VARCHAR(8), " & X_CAMPO & ", 3)"
                
                Case "'DD/MM/RR HH:24:MI'"
                    'X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ", 3) + ' ' + CONVERT(VarChar(23), " & X_CAMPO & ", 108)"
                    X_CONVERTE = "CONVERT(VARCHAR(8), " & X_CAMPO & ", 3) + ' ' + CONVERT(VarChar(5), " & X_CAMPO & ", 108)"
                   
                Case "'DDMMYYYYHH24MISS'"
                    X_CONVERTE = "REPLACE(CONVERT(VARCHAR(10)," & X_CAMPO & ",103), '/', '') + "
                    X_CONVERTE = X_CONVERTE & "REPLACE(CONVERT(VARCHAR(8)," & X_CAMPO & ", 114),':','')"
                Case "0"
                    X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ")"
                Case "'0'"
                    'X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ")"
                    X_CONVERTE = "RIGHT('0' + CONVERT(VARCHAR," & X_CAMPO & "),1)"
                Case "'00'"
                    'X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ")"
                        X_CONVERTE = "RIGHT('00' + CONVERT(VARCHAR," & X_CAMPO & "),2)"
                Case "'000'"
                    'X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ")"
                        X_CONVERTE = "RIGHT('000' + CONVERT(VARCHAR," & X_CAMPO & "),3)"
                Case "'0000'"
                    'X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ")"
                        X_CONVERTE = "RIGHT('0000' + CONVERT(VARCHAR," & X_CAMPO & "),4)"
                Case "'00000'"
                    'X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ")"
                        X_CONVERTE = "RIGHT('00000' + CONVERT(VARCHAR," & X_CAMPO & "),5)"
                Case "'000000'"
                    'X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ")"
                        X_CONVERTE = "RIGHT('000000' + CONVERT(VARCHAR," & X_CAMPO & "),6)"
                Case "'0000000'"
                    'X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ")"
                        X_CONVERTE = "RIGHT('0000000' + CONVERT(VARCHAR," & X_CAMPO & "),7)"
                Case "'00000000'"
                    'X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ")"
                        X_CONVERTE = "RIGHT('00000000' + CONVERT(VARCHAR," & X_CAMPO & "),8)"
                Case "'000000000'"
                    'X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ")"
                        X_CONVERTE = "RIGHT('000000000' + CONVERT(VARCHAR," & X_CAMPO & "),9)"
                Case "'0000000000'"
                    'X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ")"
                        X_CONVERTE = "RIGHT('0000000000' + CONVERT(VARCHAR," & X_CAMPO & "),10)"
                Case "'00000000000'"
                    'X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ")"
                        X_CONVERTE = "RIGHT('00000000000' + CONVERT(VARCHAR," & X_CAMPO & "),11)"
                Case "OPCIONAL"
                    X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ")"
                Case "'9999999990.000'"
                    X_CONVERTE = "CONVERT(VARCHAR(23), " & X_CAMPO & ")"
                
                Case Else
                    MsgBox "FORMATAÇÃO NÃO PREVISTA"
        End Select
        
    SELECIONA_FORMATO = X_CONVERTE
End Function
Public Function CONVERTE_SQL(TEXTO As String) As String
Dim X_FORMATO As String
Dim X_CAMPO As String

If Banco_Em_Uso = "SQLSERVER" Then
Banco_Sgipa = UCase(Banco_Sgipa)
Banco_Operador = UCase(Banco_Operador)
Dim Posicao As Integer


'0 CRIA SEQUENCIA SQL VERIFICA SE INSTRUÇAO É DO TIPO SELECT NEXTVAL FROM DUAL E CRIA NOVA SEQUENCIA SE FOR NECESSARIO

 Dim RSNEXT As New ADODB.Recordset
 Dim banco(4) As String
BaseDeDados = ""
 sql_inicial = UCase(TEXTO)

    TEM_SEQUENCIA = 0
      
        If InStr(1, sql_inicial, "SELECT") Then
            If InStr(1, sql_inicial, "DUAL") Then
                    If InStr(1, sql_inicial, "NEXTVAL") Then
                       TEM_SEQUENCIA = 1
                    Else   'NOVO
                       sql_inicial = Replace(sql_inicial, "FROM DUAL", "") 'NOVO
                       CONVERTE_SQL = Replace(sql_inicial, "SYSDATE", "GETDATE()")
                       Exit Function   'NOVO TESTAR
                    End If
            End If
        End If
 
    If TEM_SEQUENCIA = 1 Then

    'SELECT NEXTVAL

        If Right(Banco_Sgipa, 4) = "DBO." Then banco(1) = Banco_Sgipa Else banco(1) = Banco_Sgipa & "DBO."
        If Right(Banco_Operador, 4) = "DBO." Then banco(2) = Banco_Operador Else banco(2) = Banco_Operador & "DBO."
        banco(3) = "REDEX.DBO."
        banco(4) = "FATURA.DBO."

        For i = 1 To 4
            If InStr(sql_inicial, banco(i)) Then
              BaseDeDados = banco(i)
              Exit For
            End If
        Next i
        
        NAO_TEM_BASE = 0
        If BaseDeDados = "" Then
            BaseDeDados = Banco_Sgipa
            NAO_TEM_BASE = 1
        End If
        
        'REFORMULA A STRING E SEPARA O NOME DA SEQUENCIA E DO ALIAS
        
            FIM_SEQUENCIA = InStr(1, sql_inicial, ".NEXTVAL")
            SEMI_sQL = Left(sql_inicial, FIM_SEQUENCIA - 1)
            
            If NAO_TEM_BASE = 0 Then
                inicio_sequencia = InStrRev(SEMI_sQL, ".") + 1
            Else
                inicio_sequencia = InStrRev(SEMI_sQL, " ") + 1
            End If
            
            Nome_Sequencia = Mid(SEMI_sQL, inicio_sequencia)
            Nome_Sequencia = LTrim(RTrim(Nome_Sequencia))

            NOVO_SQL = Replace(sql_inicial, "FROM DUAL", "")
            INICIO_ALIAS = InStr(FIM_SEQUENCIA, NOVO_SQL, " AS ")
            If INICIO_ALIAS <> 0 Then
                NOME_ALIAS = Mid(NOVO_SQL, INICIO_ALIAS)
            Else
            NOME_ALIAS = ""
            End If
        
        'VERIFICA SE TABELA SEQUENCIA EXISTE, CASO NEGATIVO CRIA UMA NOVA COM VALOR INICIAL 0
                
                SQl = " SELECT 1 from " & BaseDeDados & "sysobjects where name = '" & Nome_Sequencia & "'  and xtype = 'U'"
             
                If RSNEXT.State = adStateOpen Then RSNEXT.Close
                 RSNEXT.Open SQl, db
            
                If RSNEXT.EOF Then
                    'cria tabela
                    SQl = " create table " & BaseDeDados & Nome_Sequencia & " (NEXTVAL  NUMERIC(8) IDENTITY(1,1))"
                    Executa SQl
                End If
    
        'INSERE NOVO VALOR NA SEQUENCIA E SELECIONA ULTIMO VALOR INSERIDO
  
                INSERTE_SQL = "INSERT INTO " & BaseDeDados & Nome_Sequencia & " DEFAULT VALUES"
                Executa INSERTE_SQL
                
                SQL_FINAL = " SELECT IDENT_CURRENT('" & BaseDeDados & Nome_Sequencia & "') " & NOME_ALIAS
                CONVERTE_SQL = SQL_FINAL
                
                Exit Function

    End If
'
'FIM CRIAÇÃO SEQUENCIA


'INICIO

If Banco_Em_Uso = "SQLSERVER" Then
   ' TEXTO = UCase(TEXTO)
   
    If InStr(1, UCase(TEXTO), "SYSDATE") Then
        If InStr(1, UCase(TEXTO), "DUAL") Then
             TEXTO = Replace(TEXTO, "SYSDATE", "GETDATE()")
              TEXTO = Replace(TEXTO, "FROM DUAL", "")
        End If
    End If
'End If

'TEXTO = UCase(TEXTO)
'/*1 ---------------------------------- CONVERTER NVL , SUBSTR, LENGHT,(trim(------------------------------------------------
TEXTO = Replace$(UCase(TEXTO), "NVL", "ISNULL")
   'MsgBox TEXTO
TEXTO = Replace$(UCase(TEXTO), "SUBSTR(", "SUBSTRING(")
TEXTO = Replace$(UCase(TEXTO), "SUBSTR (", "SUBSTRING(")

TEXTO = Replace$(UCase(TEXTO), "LENGHT", "LEN")

TEXTO = Replace$(UCase(TEXTO), "(TRIM(", "(RTRIM(")
TEXTO = Replace$(UCase(TEXTO), " TRIM(", " RTRIM(")
TEXTO = Replace$(UCase(TEXTO), ",TRIM(", ",RTRIM(")
TEXTO = Replace$(UCase(TEXTO), "|TRIM(", "|RTRIM(")
'

'SQL =" SELECT 'SCG'||TRIM(TO_CHAR(NVL(A.PORTO_ORIGEM,0),'00000'))||TRIM(TO_CHAR(NVL(A.ARMADOR,0),'0000000'))||TRIM(TO
'-----------------TO_CHAR(NVL
TEXTO2 = TEXTO
nova_string = ""

Do While InStr(1, TEXTO2, "TO_CHAR(ISNULL(") <> 0 'SÓ INICIA SE POSSUIR TO_CHAR(NVL
  
    Posicao = 1
    
    X_INICIO = InStr(Posicao, UCase(TEXTO2), "TO_CHAR(ISNULL(") 'INICIO DA INSTRUCAO TO_CHAR"

'// ISOLA AS VARIÁVEIS DO CAMPO E DA FORMATAÇAO PARA TROCA
    
    INICIO_CAMPO = InStr(X_INICIO, TEXTO2, "(") 'POSICAO INICIO DO CAMPO
     INICIO_CAMPO = InStr(INICIO_CAMPO, TEXTO2, "(") 'POSICAO INICIO DO CAMPO
    
    FIM_CAMPO_Z = InStr(INICIO_CAMPO, TEXTO2, ")") 'POSICAO FIM  DO CAMPO
    
     
    X_CAMPO = ""
    z = Len(TEXTO2)
    contador = 0
    
        For i = INICIO_CAMPO To FIM_CAMPO_Z
            If Mid(TEXTO2, i, 1) = "," Then
                 contador = contador + 1
            End If
        Next i
   
    If contador > 0 Then
        FIM_CAMPO = InStr(X_INICIO, TEXTO2, ")") ' POSICAO FIM DO CAMPO
        X_INI_FORMATO = InStr(FIM_CAMPO, TEXTO2, ",")
        X_FIM_FORMATO = InStr(X_INI_FORMATO, TEXTO2, ")")
        X_FORMATO = TEXTO2

        X_FORMATO = (Right(TEXTO2, Len(TEXTO2) - X_INI_FORMATO))
        X_FORMATO = Left(X_FORMATO, X_FIM_FORMATO - X_INI_FORMATO - 1)

        X_CAMPO = (Right(TEXTO2, Len(TEXTO2) - INICIO_CAMPO))
        X_CAMPO = Left(X_CAMPO, FIM_CAMPO - INICIO_CAMPO)

    Else
        FIM_CAMPO = InStr(X_INICIO, TEXTO2, ")") ' POSICAO FIM DO CAMPO
        X_INI_FORMATO = InStr(X_INICIO, TEXTO2, "(")
        X_FIM_FORMATO = InStr(X_INI_FORMATO, TEXTO2, ")")

        X_FORMATO = "OPCIONAL"
        X_CAMPO = (Right(TEXTO2, Len(TEXTO2) - INICIO_CAMPO))
        X_CAMPO = Left(X_CAMPO, FIM_CAMPO - INICIO_CAMPO - 1)
    End If
    
        FIND_STRING = Mid(TEXTO2, X_INICIO, X_FIM_FORMATO - X_INICIO + 1) 'PROCURA A INSTRUÇÃO ORACLE ANTERIOR PARA SUBSTITUIÇÃO
    
        X_CONVERTE = SELECIONA_FORMATO(X_FORMATO, X_CAMPO)
  
nova_string = nova_string & Replace(Left(TEXTO2, X_FIM_FORMATO), FIND_STRING, X_CONVERTE)

TEXTO2 = Right(TEXTO2, Len(TEXTO2) - X_FIM_FORMATO) 'RESTO DA STRING PARA ANALISAR
Loop

nova_string = nova_string & TEXTO2
'MsgBox NOVA_sTRING



'FIM TO_CHAR(NVL)



'*2---------------------------------- TRANSFORMAR TO_CHAR EM CONVERT----------------------------------

'SELECT RIGHT('00000'+ CONVERT(VARCHAR,Num),6) AS NUM FROM Numbers
'wSQL =" SELECT 'SCG'||TRIM(TO_CHAR(NVL(A.PORTO_ORIGEM,0),'00000'))||TRIM(TO_CHAR(NVL(A.ARMADOR,0),'0000000'))||TRIM(TO_CHAR(NVL(B.EMBALAGEM,0),'00'))||'" & Modo & "'||TRIM(TO_CHAR(NVL(C.RECINTO_DESTINO,0),'000')) AS SCG "



TEXTO2 = nova_string
nova_string = ""

Do While InStr(1, TEXTO2, "TO_CHAR") <> 0 'SÓ INICIA SE POSSUIR TO_CHAR
  
    Posicao = 1
    
    X_INICIO = InStr(Posicao, UCase(TEXTO2), "TO_CHAR") 'INICIO DA INSTRUCAO TO_CHAR"

'// ISOLA AS VARIÁVEIS DO CAMPO E DA FORMATAÇAO PARA TROCA
    
    INICIO_CAMPO = InStr(X_INICIO, TEXTO2, "(") 'POSICAO INICIO DO CAMPO
    
    FIM_CAMPO_Z = InStr(INICIO_CAMPO, TEXTO2, ")") 'POSICAO INICIO DO CAMPO
    
     
    X_CAMPO = ""
    z = Len(TEXTO2)
    contador = 0
    
        For i = INICIO_CAMPO To FIM_CAMPO_Z
            If Mid(TEXTO2, i, 1) = "," Then
                 contador = contador + 1
            End If
        Next i
   
    If contador > 0 Then
        FIM_CAMPO = InStr(X_INICIO, TEXTO2, ",") ' POSICAO FIM DO CAMPO
        X_INI_FORMATO = InStr(X_INICIO, TEXTO2, ",")
        X_FIM_FORMATO = InStr(X_INI_FORMATO, TEXTO2, ")")
        X_FORMATO = TEXTO2

        X_FORMATO = (Right(TEXTO2, Len(TEXTO2) - X_INI_FORMATO))
        X_FORMATO = Left(X_FORMATO, X_FIM_FORMATO - X_INI_FORMATO - 1)

        X_CAMPO = (Right(TEXTO2, Len(TEXTO2) - INICIO_CAMPO))
        X_CAMPO = Left(X_CAMPO, FIM_CAMPO - INICIO_CAMPO - 1)

    Else
        FIM_CAMPO = InStr(X_INICIO, TEXTO2, ")") ' POSICAO FIM DO CAMPO
        X_INI_FORMATO = InStr(X_INICIO, TEXTO2, "(")
        X_FIM_FORMATO = InStr(X_INI_FORMATO, TEXTO2, ")")

        X_FORMATO = "OPCIONAL"
        X_CAMPO = (Right(TEXTO2, Len(TEXTO2) - INICIO_CAMPO))
        X_CAMPO = Left(X_CAMPO, FIM_CAMPO - INICIO_CAMPO - 1)
    End If
    
        FIND_STRING = Mid(TEXTO2, X_INICIO, X_FIM_FORMATO - X_INICIO + 1) 'PROCURA A INSTRUÇÃO ORACLE ANTERIOR PARA SUBSTITUIÇÃO
    
'***************************************************
  X_CONVERTE = SELECIONA_FORMATO(X_FORMATO, X_CAMPO)
'***************************************************

nova_string = nova_string & Replace(Left(TEXTO2, X_FIM_FORMATO), FIND_STRING, X_CONVERTE)

TEXTO2 = Right(TEXTO2, Len(TEXTO2) - X_FIM_FORMATO) 'RESTO DA STRING PARA ANALISAR
Loop

nova_string = nova_string & TEXTO2




'MsgBox NOVA_STRING
'-------------------------- '*3 PROCURA CONVERTER TO_DATE ----------------------------


TEXTO2 = UCase(nova_string)
'MsgBox TEXTO2
nova_string = ""

Do While InStr(1, TEXTO2, "TO_DATE") <> 0 'SÓ INICIA SE POSSUIR TO_DATE
  
    Posicao = 1
    X_INICIO = InStr(Posicao, UCase(TEXTO2), "TO_DATE") 'INICIO DA INSTRUCAO TO_DATE"

'// ISOLA AS VARIÁVEIS DO CAMPO E DA FORMATAÇAO PARA TROCA
    
    INICIO_CAMPO = InStr(X_INICIO, TEXTO2, "(") 'POSICAO INICIO DO CAMPO
    FIM_CAMPO = InStr(X_INICIO, TEXTO2, ",") ' POSICAO FIM DO CAMPO

    X_INI_FORMATO = InStr(X_INICIO, TEXTO2, ",")
    X_FIM_FORMATO = InStr(X_INI_FORMATO, TEXTO2, ")")

    X_FORMATO = TEXTO2

    X_FORMATO = (Right(TEXTO2, Len(TEXTO2) - X_INI_FORMATO))
    X_FORMATO = Left(X_FORMATO, X_FIM_FORMATO - X_INI_FORMATO - 1)

    X_CAMPO = (Right(TEXTO2, Len(TEXTO2) - INICIO_CAMPO))
    X_CAMPO = Left(X_CAMPO, FIM_CAMPO - INICIO_CAMPO - 1)

'//
    FIND_STRING = Mid(TEXTO2, X_INICIO, X_FIM_FORMATO - X_INICIO + 1) 'PROCURA A INSTRUÇÃO ORACLE ANTERIOR PARA SUBSTITUIÇÃO
    
    X_FORMATO = LTrim(X_FORMATO)
    X_FORMATO = RTrim(X_FORMATO)
    Select Case UCase(LTrim(X_FORMATO))
                
                Case "'YYYYMMDD'"
                     X_CONVERTE = "CONVERT(DATETIME, " & X_CAMPO & " , 112)"
                Case "'HH24:MI:SS'"
                     X_CONVERTE = "CONVERT(DATETIME, " & X_CAMPO & " , 103)"
                Case "'DD/MM/YYYY'"
                     X_CONVERTE = "CONVERT(DATETIME, " & X_CAMPO & " , 103)"
                Case "'DD/MM/YY HH24:MI:SS'"
                     X_CONVERTE = "CONVERT(DATETIME, " & X_CAMPO & ", 103)"
                Case "'DD/MM/YY HH24:MI:SS'"
                     X_CONVERTE = "CONVERT(DATETIME, " & X_CAMPO & ", 103)"
                Case "'DD/MM/YY HH24:MI'"
                     X_CONVERTE = "CONVERT(DATETIME, " & X_CAMPO & ", 103)"
                Case "'DD/MM/YYYY HH24:MI'"
                     X_CONVERTE = "CONVERT(DATETIME, " & X_CAMPO & ", 103)"
                Case "'DD/MM/YY'"
                     X_CONVERTE = "CONVERT(DATETIME, " & X_CAMPO & ", 103)"
                Case "'DD/MM/YYYY HH24:MI:SS'"
                     X_CONVERTE = "CONVERT(DATETIME, " & X_CAMPO & " , 103)"
                Case "'YYYY/MM/DD'"
                     X_CONVERTE = "CONVERT(DATETIME, " & X_CAMPO & ", 111)"
                Case "'YYYY/MM/DD HH24:MI:SS'"
                     X_CONVERTE = "CONVERT(DATETIME, " & X_CAMPO & ", 111)"
                Case "'DD/MM/RR'"
                     X_CONVERTE = "CONVERT(DATETIME, " & X_CAMPO & ", 3)"
                
                Case Else
                     MsgBox "FORMATO DE DATA NÃO PREVISTO"
    End Select
 
nova_string = nova_string & Replace(Left(TEXTO2, X_FIM_FORMATO), FIND_STRING, X_CONVERTE)
TEXTO2 = Right(TEXTO2, Len(TEXTO2) - X_FIM_FORMATO) 'RESTO DA STRING PARA ANALISAR
  
Loop

nova_string = nova_string & TEXTO2

'*4 ------------------------------------- DECODE -------------------------------------------------

Dim str_ARRAY(0 To 12) As String
Dim XCAMPO As String
Dim ESQUERDO As Integer

TEXTO2 = UCase(nova_string)
nova_string = ""

Do While InStr(1, TEXTO2, "DECODE") <> 0 'SÓ INICIA SE POSSUIR DECODE
'  MsgBox TEXTO2
    pos1 = InStr(1, TEXTO2, "DECODE")
    FRASE_INICIO = Left(TEXTO2, pos1 - 1)
     
    POS2 = InStr(pos1, TEXTO2, ")")
    
      ESQUERDO = 0
      Direito = 0
      
      TEXTO_PARENTESES = Mid(TEXTO2, pos1, POS2 - pos1 + 1)
      
     For z = 1 To Len(TEXTO_PARENTESES) 'VERIFICA SE HA INSTRUÇOES ANINHADAS NO COMANDO
           If Mid(TEXTO_PARENTESES, z, 1) = "(" Then
               ESQUERDO = ESQUERDO + 1
           End If
           If Mid(TEXTO_PARENTESES, z, 1) = ")" Then
               Direito = Direito + 1
           End If
     Next z
     
          
     
        pos4 = pos1
        Do Until ESQUERDO = Direito
            POS3 = InStr(pos4, TEXTO2, ")")
            pos4 = InStr(POS3 + 1, TEXTO2, ")")
            POS2 = pos4
            Direito = Direito + 1
        Loop
    
    FRASE_FIM = Right(TEXTO2, Len(TEXTO2) - POS2)

    X_INSTRUCAO = (Mid(TEXTO2, pos1, POS2 + 1 - pos1))
    pos1 = InStr(1, X_INSTRUCAO, "(")
    X_INSTRUCAO = Mid(X_INSTRUCAO, pos1)
  
    'X_INSTRUCAO = Replace(X_INSTRUCAO, ")", ",")
    Ultimo = Len(X_INSTRUCAO) - 1
    'X_INSTRUCAO = Replace(X_INSTRUCAO, ")", ",")
    X_INSTRUCAO = Left(X_INSTRUCAO, Len(X_INSTRUCAO) - 1)
    
    'X_INSTRUCAO = Replace(X_INSTRUCAO, "(", "")
    X_INSTRUCAO = Replace(X_INSTRUCAO, "(", "", 1, 1)
    X_CAMPO = ""
    z = Len(X_INSTRUCAO)

    contador = 0
    
        For i = 1 To z
            If Mid(X_INSTRUCAO, i, 1) = "," Then
                str_ARRAY(contador) = X_CAMPO
                X_CAMPO = ""
                contador = contador + 1
            Else
                X_CAMPO = X_CAMPO & Mid(X_INSTRUCAO, i, 1)
            End If
        Next i
str_ARRAY(contador) = X_CAMPO

    contador = contador
    
Select Case contador
  Case 10
     instrucao_final = " CASE " & str_ARRAY(0) & " WHEN " & str_ARRAY(1) & " THEN " & str_ARRAY(2) & " WHEN " & str_ARRAY(3) & " THEN " & str_ARRAY(4) & " WHEN " & str_ARRAY(5) & " THEN " & str_ARRAY(6) & " WHEN " & str_ARRAY(7) & " THEN " & str_ARRAY(8) & " WHEN " & str_ARRAY(9) & " THEN " & str_ARRAY(10) & " END"
  Case 9
     instrucao_final = " CASE " & str_ARRAY(0) & " WHEN " & str_ARRAY(1) & " THEN " & str_ARRAY(2) & " WHEN " & str_ARRAY(3) & " THEN " & str_ARRAY(4) & " WHEN " & str_ARRAY(5) & " THEN " & str_ARRAY(6) & " WHEN " & str_ARRAY(7) & " THEN " & str_ARRAY(8) & " ELSE " & str_ARRAY(9) & " END "
  Case 8
     instrucao_final = " CASE " & str_ARRAY(0) & " WHEN " & str_ARRAY(1) & " THEN " & str_ARRAY(2) & " WHEN " & str_ARRAY(3) & " THEN " & str_ARRAY(4) & " WHEN " & str_ARRAY(5) & " THEN " & str_ARRAY(6) & " WHEN " & str_ARRAY(7) & " THEN " & str_ARRAY(8) & " END"
  Case 7
     instrucao_final = " CASE " & str_ARRAY(0) & " WHEN " & str_ARRAY(1) & " THEN " & str_ARRAY(2) & " WHEN " & str_ARRAY(3) & " THEN " & str_ARRAY(4) & " WHEN " & str_ARRAY(5) & " THEN " & str_ARRAY(6) & " ELSE " & str_ARRAY(7) & " END "
  Case 6
     instrucao_final = " CASE " & str_ARRAY(0) & " WHEN " & str_ARRAY(1) & " THEN " & str_ARRAY(2) & " WHEN " & str_ARRAY(3) & " THEN " & str_ARRAY(4) & " WHEN " & str_ARRAY(5) & " THEN " & str_ARRAY(6) & " END"
  Case 5
     instrucao_final = " CASE " & str_ARRAY(0) & " WHEN " & str_ARRAY(1) & " THEN " & str_ARRAY(2) & " WHEN " & str_ARRAY(3) & " THEN " & str_ARRAY(4) & " ELSE " & str_ARRAY(5) & " END "
  Case 4
     instrucao_final = " CASE " & str_ARRAY(0) & " WHEN " & str_ARRAY(1) & " THEN " & str_ARRAY(2) & " WHEN " & str_ARRAY(3) & " THEN " & str_ARRAY(4) & " END "
  Case 3
     instrucao_final = " CASE " & str_ARRAY(0) & " WHEN " & str_ARRAY(1) & " THEN " & str_ARRAY(2) & " ELSE " & str_ARRAY(3) & " END "
End Select
 
TEXTO2 = FRASE_INICIO & instrucao_final & FRASE_FIM

Loop

nova_string = nova_string & TEXTO2


'* 5 --------------------------------------CONVERTE TO_NUMBER-------------------------------

TEXTO2 = UCase(nova_string)
nova_string = ""

    Do While InStr(1, TEXTO2, "TO_NUMBER") <> 0 'SÓ INICIA SE POSSUIR TO_NUMBER
        Posicao = 1
        X_INICIO = InStr(Posicao, UCase(TEXTO2), "TO_NUMBER") 'INICIO DA INSTRUCAO TO_NUMBER"
    
        INICIO_CAMPO = InStr(X_INICIO, TEXTO2, "(")
        FIM_CAMPO = InStr(INICIO_CAMPO, TEXTO2, ")")
    
        X_CAMPO = Mid(TEXTO2, INICIO_CAMPO + 1, FIM_CAMPO - INICIO_CAMPO - 1)

        STR_SQL = "CONVERT(NUMERIC(8,2), " & X_CAMPO & ")"
        STR_ORACLE = Mid(TEXTO2, X_INICIO, FIM_CAMPO - X_INICIO + 1)
     
        nova_string = nova_string & Replace(Left(TEXTO2, FIM_CAMPO), STR_ORACLE, " " & STR_SQL)
        TEXTO2 = Right(TEXTO2, Len(TEXTO2) - FIM_CAMPO) 'RESTO DA STRING PARA ANALISAR
    Loop

nova_string = nova_string & TEXTO2



'-----------------'------------------------------CONVERTE RPAD2 -----------------------


TEXTO2 = UCase(nova_string)
nova_string = ""
'MsgBox TEXTO2
'MsgBox TEXTO2
TEM_RPAD = 0

Do While InStr(1, TEXTO2, "RPAD(CONVERT") <> 0 'SÓ INICIA SE POSSUIR RPAD(CONVERT

                    X_INICIO = InStr(Posicao, UCase(TEXTO2), "CONVERT")
                    INICIO_CAMPO = InStr(X_INICIO, TEXTO2, ",")
                    FIM_CAMPO = InStr(INICIO_CAMPO, TEXTO2, ",")
                    FIM_INSTRUCAO = InStr(FIM_CAMPO + 1, TEXTO2, ")")
                    X_CAMPO = LTrim(RTrim(Mid(TEXTO2, X_INICIO, FIM_INSTRUCAO - X_INICIO + 1)))
                    
                    PAD_INICIO = InStr(X_INICIO, TEXTO2, X_CAMPO)
                    PAD_INICIO = PAD_INICIO + Len(X_CAMPO) - 1
                    PAD_INICIO = InStr(PAD_INICIO, TEXTO2, ",")
                    PAD_FIM = InStr(PAD_INICIO, TEXTO2, ")")
                    X_PAD = Mid(TEXTO2, PAD_INICIO + 1, PAD_FIM - PAD_INICIO - 1)

                    FIM_INSTRUCAO = PAD_FIM
                    X_INICIO = InStr(X_INICIO - 6, UCase(TEXTO2), "RPAD(")
                    X_FIM = InStr(X_INICIO, UCase(TEXTO2), X_PAD & ")")
                    X_FIM = InStr(X_FIM, UCase(TEXTO2), ")")
                    STR_ORACLE = Mid(TEXTO2, X_INICIO, X_FIM - X_INICIO + 1)
                    FIM_CAMPO = FIM_INSTRUCAO

        STR_SQL = " ( " & X_CAMPO & " + REPLICATE(' ', abs((" & X_PAD & " - LEN(" & X_CAMPO & ")))  )) "
        nova_string = nova_string & Replace(Left(TEXTO2, FIM_CAMPO), STR_ORACLE, " " & STR_SQL)
        TEXTO2 = Right(TEXTO2, Len(TEXTO2) - FIM_CAMPO) 'RESTO DA STRING PARA ANALISAR
   Loop

nova_string = nova_string & TEXTO2




' 6 --------------------------------------- CONVERTE RPAD ----------------------------------


TEXTO2 = UCase(nova_string)
nova_string = ""
'MsgBox TEXTO2


    Do While InStr(1, TEXTO2, "RPAD(") <> 0 'SÓ INICIA SE POSSUIR RPAD
        

       
                Posicao = 1
                 If InStr(1, TEXTO2, "MIN(") Then 'NO CASO DE VIR AGREGADO A INSTRUÇAO (MIN)
'                    MsgBox TEXTO2
                    X_INICIO = InStr(Posicao, UCase(TEXTO2), "MIN")
                    INICIO_CAMPO = InStr(X_INICIO, TEXTO2, "(")
                    FIM_CAMPO = InStr(INICIO_CAMPO, TEXTO2, ")")
                    FIM_INSTRUCAO = InStr(FIM_CAMPO + 1, TEXTO2, ")")
                    X_CAMPO = Mid(TEXTO2, X_INICIO, FIM_CAMPO - X_INICIO + 1)
                    
                    PAD_INICIO = InStr(1, TEXTO2, X_CAMPO)
                    PAD_INICIO = InStr(PAD_INICIO, TEXTO2, ",")
                    PAD_FIM = InStr(PAD_INICIO, TEXTO2, ")")
                    X_PAD = Mid(TEXTO2, PAD_INICIO + 1, PAD_FIM - PAD_INICIO - 1)

                    X_INICIO = InStr(Posicao, UCase(TEXTO2), "RPAD(")
                    X_FIM = InStr(X_INICIO, UCase(TEXTO2), ",")
                    X_FIM = InStr(X_FIM, UCase(TEXTO2), ")")
                    STR_ORACLE = Mid(TEXTO2, X_INICIO, X_FIM - X_INICIO + 1)
                    FIM_CAMPO = FIM_INSTRUCAO
                Else
                    X_INICIO = InStr(Posicao, UCase(TEXTO2), "RPAD(") 'INICIO DA INSTRUCAO RPAD"
                    INICIO_CAMPO = InStr(X_INICIO, TEXTO2, "(")
                    FIM_CAMPO = InStr(INICIO_CAMPO, TEXTO2, ")")
                    FIM_INSTRUCAO = InStr(INICIO_CAMPO, TEXTO2, ",")
                    STR_ORACLE = Mid(TEXTO2, X_INICIO, FIM_CAMPO - X_INICIO + 1)
                    X_CAMPO = Mid(TEXTO2, INICIO_CAMPO + 1, FIM_INSTRUCAO - INICIO_CAMPO - 1)
            
                    PAD_INICIO = InStr(1, STR_ORACLE, ",")
                    PAD_FIM = InStr(PAD_INICIO, STR_ORACLE, ")")
        
                    X_PAD = Mid(STR_ORACLE, PAD_INICIO + 1, PAD_FIM - PAD_INICIO - 1)
                End If
                
     
        'MsgBox Len(X_CAMPO)
        If X_CAMPO = "' '" Then
            STR_SQL = " ( " & " + REPLICATE(' ', abs((" & X_PAD & " - LEN(" & X_CAMPO & ")))  )) "
        Else
            STR_SQL = " ( " & X_CAMPO & " + REPLICATE(' ', abs((" & X_PAD & " - LEN(" & X_CAMPO & ")))  )) "
        End If
        nova_string = nova_string & Replace(Left(TEXTO2, FIM_CAMPO), STR_ORACLE, " " & STR_SQL)
        TEXTO2 = Right(TEXTO2, Len(TEXTO2) - FIM_CAMPO) 'RESTO DA STRING PARA ANALISAR
   Loop

nova_string = nova_string & TEXTO2






' 7 --------------------------- GETDATE() / ||  / SUBSTR / lenght -----------------------------------------

nova_string = Replace$(nova_string, "SYSDATE", "GETDATE()")
nova_string = Replace$(nova_string, "||", "+")
nova_string = Replace$(nova_string, "LENGTH(", "LEN(")
nova_string = Replace$(nova_string, "LENGTH (", "LEN(")

'8 --------------------------- ALL_TAB_COLUMNS --------------------------------------------------------------

If InStr(1, nova_string, "ALL_TAB_COLUMNS") Then
    nova_string = Replace$(nova_string, "ALL_TAB_COLUMNS", "INFORMATION_SCHEMA.COLUMNS")
    nova_string = Replace$(nova_string, "OWNER=", "DOMAIN_CATALOG=")
End If


'9 ----------------------------- INSTR -----------------------------------------------


'INSTR
TEXTO2 = UCase(nova_string)
nova_string = ""

Do While InStr(1, TEXTO2, "INSTR(") <> 0 'SÓ INICIA SE POSSUIR INSTR
  
    Posicao = 1
    X_INICIO = InStr(Posicao, UCase(TEXTO2), "INSTR") 'INICIO DA INSTRUCAO INSTR"

'// ISOLA AS VARIÁVEIS DO CAMPO E DA FORMATAÇAO PARA TROCA
    
    INICIO_CAMPO = InStr(X_INICIO, TEXTO2, "(") 'POSICAO INICIO DO CAMPO
    FIM_CAMPO = InStr(X_INICIO, TEXTO2, ",") ' POSICAO FIM DO CAMPO

    X_INI_FORMATO = InStr(X_INICIO, TEXTO2, ",")
    X_FIM_FORMATO = InStr(X_INI_FORMATO, TEXTO2, ")")

    X_FORMATO = TEXTO2

    X_FORMATO = (Right(TEXTO2, Len(TEXTO2) - X_INI_FORMATO))
    X_FORMATO = Left(X_FORMATO, X_FIM_FORMATO - X_INI_FORMATO - 1)

    X_CAMPO = (Right(TEXTO2, Len(TEXTO2) - INICIO_CAMPO))
    X_CAMPO = Left(X_CAMPO, FIM_CAMPO - INICIO_CAMPO - 1)

'//

    FIND_STRING = Mid(TEXTO2, X_INICIO, X_FIM_FORMATO - X_INICIO + 1) 'PROCURA A INSTRUÇÃO ORACLE ANTERIOR PARA SUBSTITUIÇÃO
    
X_CONVERTE = "CHARINDEX(" & X_FORMATO & "," & X_CAMPO & ")"
 
 
 
nova_string = nova_string & Replace(Left(TEXTO2, X_FIM_FORMATO), FIND_STRING, X_CONVERTE)
TEXTO2 = Right(TEXTO2, Len(TEXTO2) - X_FIM_FORMATO) 'RESTO DA STRING PARA ANALISAR
  
Loop

nova_string = nova_string & TEXTO2
'MsgBox NOVA_STRING

'FIM INSTR

CONVERTE_SQL = Trim$(UCase$(nova_string))
Else
CONVERTE_SQL = TEXTO

End If

Else
    CONVERTE_SQL = TEXTO
End If

End Function
Public Function Valida_BDCC(Tipo As String, Valor As String, Valida As Byte, Sistema As Byte) As Boolean
'Retorno
'0-NAO MOSTRA MENSAGEM NAO IMPEDITIVO
'1-MOSTRA MENSAGEM NAO IMPEDITIVO
'2-MOSTRA MENSAGEM IMPEDITIVO

Dim oSOAP As New MSSOAPLib30.SoapClient30
Dim objDOMDoc
Dim Rs As ADODB.Recordset
Dim Ip_Bdcc As String
Dim ListaDeNohsDoXML
Dim CodigoRetorno As String
Dim MsgRetorno As String
Dim Tipo_Bdcc As String

Valida_BDCC = True
Ip_Bdcc = ""
Msg_Bdcc = ""
SQl = "select Ip_Bdcc from " & Banco_Sgipa & "tb_empresas where autonum=1"
If Sistema = 1 Then SQl = "select Ip_Bdcc_Redex as Ip_Bdcc from " & Banco_Sgipa & "tb_empresas where autonum=1"
PRSet Rs, SQl
If Not Rs.EOF Then
    Ip_Bdcc = Nnull(Rs!Ip_Bdcc, 1)
End If
Rs.Close
If Ip_Bdcc <> "" Then
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    oSOAP.MSSoapInit (Ip_Bdcc)
    Set objDOMDoc = CreateObject("Msxml2.DOMDocument.4.0")
    Tipo = UCase(Tipo)
    Select Case UCase(Tipo)
    Case "RENAVAM"
        VM = "V"
        Set ListaDeNohsDoXML = oSOAP.ConsultaRenavam(Valor, Valida)
        CodigoRetorno = ListaDeNohsDoXML(0).Text
        MsgRetorno = ListaDeNohsDoXML(1).Text
    Case "CPF"
        VM = "M"
        Set ListaDeNohsDoXML = oSOAP.Consultacpf(Valor, Valida)
        CodigoRetorno = ListaDeNohsDoXML(0).Text
        MsgRetorno = ListaDeNohsDoXML(1).Text
    End Select
    Tipo_Bdcc = "0"
    Msg_Bdcc = ""
    SQl = "select Tipo_Retorno from " & Banco_Sgipa & "tb_Bdcc_Retorno where cod_retorno='" & CodigoRetorno & "'"
    PRSet Rs, SQl
    If Not Rs.EOF Then
        Tipo_Bdcc = Nnull(Rs!Tipo_Retorno, 0)
    End If
    Rs.Close
    If Tipo_Bdcc = "1" Or Tipo_Bdcc = "2" Then MsgBox MsgRetorno
    If Tipo_Bdcc = "2" Then Valida_BDCC = False
Else
    Valida_BDCC = True
End If

End Function


Public Function PPonto(Valor As Variant) As Variant
Dim WValor As Variant
If Valor = "" Then Valor = "0"
WValor = Valor
 
If CodEmpSis = "BAN" Then
    WValor = Replace(WValor, ",", ".")
Else
    Do
      If InStr(WValor, ".") > 0 Then
        WValor = Mid$(WValor, 1, InStr(WValor, ".") - 1) + Mid$(WValor, InStr(WValor, ".") + 1)
      Else
        Exit Do
      End If
    Loop
    If InStr(WValor, ",") > 0 Then
        WValor = Mid$(WValor, 1, InStr(WValor, ",") - 1) + "." + Mid$(WValor, InStr(WValor, ",") + 1)
    End If
End If
 
 PPonto = WValor
End Function

Public Sub Envia(Endereco As String, Rotulo As String, Mensagem As String, C_Empresa As String, Optional Anexo As String, Optional Mens_Html As String, Optional Anexo1 As String, Optional Profile As String = "CIABAND", Optional MudaTitulo As Boolean = False, Optional Lote As Long = 0, Optional CNTR As Long = 0, Optional CS As Long = 0, Optional EmCopia As String)


'


Dim t As Long
t = Len(Mensagem) + 256


     
Dim Recip() As String
Dim cmd As New ADODB.Command
Dim Aux As Integer
Dim To_Ende As String
Dim ErroEmail As Boolean

Screen.MousePointer = 11
If Profile = "" Then Profile = "CIABAND"
Endereco = LCase(Endereco)
If Right(Endereco, 1) = ";" Then Endereco = Left(Endereco, Len(Endereco) - 1)

Recip() = Split(Endereco, ";")

    For Aux = 0 To UBound(Recip)
        ErroEmail = False
        
        To_Ende = Recip(Aux)
        
        cmd.ActiveConnection = ConnectionString
        cmd.CommandType = adCmdStoredProc
    
        cmd.CommandText = "" & Banco_Sgipa & "PROC_CHRONOS_EMAIL"
        
        cmd.Parameters.Append cmd.CreateParameter("From_name", adVarChar, adParamInput, 500, Profile)
        cmd.Parameters.Append cmd.CreateParameter("to_names", adVarChar, adParamInput, 500, "<" & To_Ende & ">")
        cmd.Parameters.Append cmd.CreateParameter("subject", adVarChar, adParamInput, 500, Rotulo)
        
        If EmCopia <> "" And UCase(To_Ende) <> UCase(EmCopia) Then
            cmd.Parameters.Append cmd.CreateParameter("cc_names", adVarChar, adParamInput, 500, UCase(EmCopia))
        Else
            cmd.Parameters.Append cmd.CreateParameter("cc_names", adVarChar, adParamInput, 500, Null)
        End If
        
        cmd.Parameters.Append cmd.CreateParameter("bcc_names", adVarChar, adParamInput, 500, Null)
        cmd.Parameters.Append cmd.CreateParameter("message", adVarChar, adParamInput, t, Mensagem)
        
        If Mens_Html = "" Then
          cmd.Parameters.Append cmd.CreateParameter("html_message", adLongVarChar, adParamInput, 32000, Null)
        Else
          cmd.Parameters.Append cmd.CreateParameter("html_message", adLongVarChar, adParamInput, 32000, Mens_Html)
        End If
        If Anexo > "" Then
           cmd.Parameters.Append cmd.CreateParameter("file_attachments", adVarChar, adParamInput, 500, Anexo)
        Else
           cmd.Parameters.Append cmd.CreateParameter("file_attachments", adVarChar, adParamInput, 500, Null)
        End If
        cmd.Parameters.Append cmd.CreateParameter("ErroCode", adVarChar, adParamOutput, 100)
        cmd.Execute
        If cmd.Parameters("ErroCode").Value <> "" Then
            If Not CalcAuto Then MsgBox "Error Code: " & cmd.Parameters("ErroCode").Value
            ErroEmail = True
        End If
        Set cmd = Nothing
    
        
        'registro envio do email
        If Lote <> 0 Then
            SQl = "INSERT INTO " & Banco_Sgipa & "tb_envia_email"
            SQl = SQl & "("
            SQl = SQl & "autonumcs,autonumcntr, usuario, data_solicita, data_envio, email,para, autonumbl, flag_tipo_envio"
            SQl = SQl & ",flag_automatico,flag_enviado,profile,anexo,assunto)"
            SQl = SQl & "VALUES"
            SQl = SQl & "("
            SQl = SQl & CS
            SQl = SQl & "," & CNTR
            SQl = SQl & "," & Cod_Usuario
            SQl = SQl & ",GETDATE()"
            SQl = SQl & ",GETDATE()"
            SQl = SQl & ",'" & To_Ende & "'"
            SQl = SQl & ",''"
            SQl = SQl & "," & Lote
            SQl = SQl & ",0"
            SQl = SQl & ",0"
            SQl = SQl & ",1"
            SQl = SQl & ",'" & Profile & "'"
            SQl = SQl & ",'" & Anexo & "'"
            SQl = SQl & ",'" & Rotulo & "'"
            SQl = SQl & ")"
            Executa SQl
        End If
        
    Next



Screen.MousePointer = 0

End Sub
Public Function Calcula_Valor_Periodo(Lote As Long, Servico As Long, Periodo As Integer, Lista As Long, Dias As Long) As Double
Dim rsaux As ADODB.Recordset
Dim WSQL As String
Dim Valor As Double
Dim Fator As Double
Dim Total20 As Double
Dim Total40 As Double
Dim Minimo As Double

Total20 = 0
Total40 = 0
PRSet rsaux, "SELECT count(a.autonum)  as cntr FROM " & Banco_Sgipa & "TB_CNTR_BL A," & Banco_Sgipa & "TB_AMR_CNTR_BL B WHERE B.CNTR=A.AUTONUM AND a.tamanho=20 and a.flag_terminal=1 and b.bl=" & Lote
Total20 = Nnull(rsaux!CNTR, 0)
PRSet rsaux, "SELECT count(a.autonum)  as cntr FROM " & Banco_Sgipa & "TB_CNTR_BL A," & Banco_Sgipa & "TB_AMR_CNTR_BL B WHERE B.CNTR=A.AUTONUM AND  a.tamanho=40 and a.flag_terminal=1 and b.bl=" & Lote
Total40 = Nnull(rsaux!CNTR, 0)
WSQL = "Select PRECO_UNITARIO,PRECO_MINIMO  from " & Banco_Sgipa & "TB_LISTA_P_S_PERIODO"
WSQL = WSQL & " where lista=" & Lista & " and N_PERIODO=" & Periodo
WSQL = WSQL & " and servico=" & Servico
If Total40 = 0 And Total20 <> 0 Then
   WSQL = WSQL & " and (tipo_carga='SVAR20' OR tipo_carga='SVAR' OR tipo_carga='CHR20' OR tipo_carga='CHF20'  )"
Else
  If Total40 <> 0 And Total20 = 0 Then
     WSQL = WSQL & " and (tipo_carga='SVAR40' OR tipo_carga='SVAR' or tipo_carga='CHR40' or tipo_carga='CHF40'  )"
   Else
    WSQL = WSQL & " and (tipo_carga='SVAR20' OR tipo_carga='SVAR40' OR tipo_carga='SVAR' OR tipo_carga='CHR20' OR tipo_carga='CHF20' or tipo_carga='CHF40'  )"
   End If
End If
Fator = 1
PRSet rsaux, WSQL
If Not rsaux.EOF Then
   Fator = rsaux!preco_unitario
End If
WSQL = "Select PRECO_UNITARIO,PRECO_MINIMO  from " & Banco_Sgipa & "TB_LISTA_P_S_PERIODO"
WSQL = WSQL & " where lista=" & Lista & " and N_PERIODO=" & Periodo + 1
WSQL = WSQL & " and servico=" & Servico
If Total40 = 0 And Total20 <> 0 Then
   WSQL = WSQL & " and (tipo_carga='SVAR20' OR tipo_carga='SVAR' OR tipo_carga='CHR20' OR tipo_carga='CHF20'  )"
Else
  If Total40 <> 0 And Total20 = 0 Then
     WSQL = WSQL & " and (tipo_carga='SVAR40' OR tipo_carga='SVAR' or tipo_carga='CHR40' or tipo_carga='CHF40'  )"
   Else
    WSQL = WSQL & " and (tipo_carga='SVAR20' OR tipo_carga='SVAR40' OR tipo_carga='SVAR' OR tipo_carga='CHR20' OR tipo_carga='CHF20' or tipo_carga='CHF40'  )"
   End If
End If
Minimo = 0
PRSet rsaux, WSQL
If Not rsaux.EOF Then
  If Fator <> 0 Then
    Fator = rsaux!preco_unitario / Fator
  Else
    Fator = 1
  End If
  Minimo = rsaux!Preco_Minimo
Else
  Fator = 1
End If

WSQL = "Select Parcela+Parcela_Ad as ValorPer,  datediff(day,convert(datetime,DATAINIPERIODO,103),convert(datetime,DATAFIMPERIODO,103)) as dias "
WSQL = WSQL & "From Tb_Descricao_Calculo Where "
WSQL = WSQL & " Bl=" & Lote
WSQL = WSQL & " and Servico=" & Servico
WSQL = WSQL & " and Periodos=" & Periodo
PRSet rsaux, WSQL


Valor = 0
If Not rsaux.EOF Then
   If Nnull(rsaux!Dias, 0) = 0 Then
      Valor = Nnull(rsaux!ValorPer, 0)
   Else
     Valor = Nnull(rsaux!ValorPer, 0) * (Int(Dias / Nnull(rsaux!Dias, 0)) + 1)
   End If
End If
If (Valor * Fator) > Minimo Then
   Calcula_Valor_Periodo = Valor * Fator
Else
   Calcula_Valor_Periodo = Minimo

End If
End Function

Public Sub Envia_Email_XML(Endereco As String, Rotulo As String, Mensagem As String, C_Empresa As String, Optional Anexo As String, Optional Mens_Html As String, Optional Anexo1 As String, Optional Email_Envio As String)
Dim Recip() As String
Dim cmd As New ADODB.Command
Dim Aux As Integer
Dim To_Ende As String

Screen.MousePointer = 11
Endereco = LCase(Endereco)
If Right(Endereco, 1) = ";" Then Endereco = Left(Endereco, Len(Endereco) - 1)

Recip() = Split(Endereco, ";")
For Aux = 0 To UBound(Recip)
    To_Ende = Recip(Aux)
    cmd.ActiveConnection = ConnectionString
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "" & Banco_Sgipa & "PROC_CHRONOS_EMAIL"
    cmd.Parameters.Append cmd.CreateParameter("From_name", adVarChar, adParamInput, 500, "<" & Nnull(Email_Envio, 1) & ">")
    cmd.Parameters.Append cmd.CreateParameter("to_names", adVarChar, adParamInput, 500, "<" & To_Ende & ">")
    cmd.Parameters.Append cmd.CreateParameter("subject", adVarChar, adParamInput, 500, Rotulo)
    cmd.Parameters.Append cmd.CreateParameter("cc_names", adVarChar, adParamInput, 500, Null)
    cmd.Parameters.Append cmd.CreateParameter("bcc_names", adVarChar, adParamInput, 500, Null)
    cmd.Parameters.Append cmd.CreateParameter("message", adVarChar, adParamInput, 32000, Mensagem)
    If Mens_Html = "" Then
      cmd.Parameters.Append cmd.CreateParameter("html_message", adLongVarChar, adParamInput, 32000, Null)
    Else
      cmd.Parameters.Append cmd.CreateParameter("html_message", adLongVarChar, adParamInput, 32000, Mens_Html)
    End If
    If Anexo > "" Then
       cmd.Parameters.Append cmd.CreateParameter("filename1", adVarChar, adParamInput, 500, Anexo)
       If InStr(UCase(Anexo), "PDF") = 0 Then
          cmd.Parameters.Append cmd.CreateParameter("filetype1", adVarChar, adParamInput, 500, "text/plain")
       Else
          cmd.Parameters.Append cmd.CreateParameter("filetype1", adVarChar, adParamInput, 500, "application/pdf")
       End If
    Else
       cmd.Parameters.Append cmd.CreateParameter("filename1", adVarChar, adParamInput, 500, Null)
       cmd.Parameters.Append cmd.CreateParameter("filetype1", adVarChar, adParamInput, 500, Null)
    End If
    If Anexo1 > "" Then
       cmd.Parameters.Append cmd.CreateParameter("filename2", adVarChar, adParamInput, 500, Anexo1)
       If InStr(UCase(Anexo1), "PDF") = 0 Then
          cmd.Parameters.Append cmd.CreateParameter("filetype2", adVarChar, adParamInput, 500, "text/plain")
       Else
          cmd.Parameters.Append cmd.CreateParameter("filetype2", adVarChar, adParamInput, 500, "application/pdf")
       End If
    Else
       cmd.Parameters.Append cmd.CreateParameter("filename2", adVarChar, adParamInput, 500, Null)
       cmd.Parameters.Append cmd.CreateParameter("filetype2", adVarChar, adParamInput, 500, Null)
    End If
    cmd.Parameters.Append cmd.CreateParameter("filename3", adVarChar, adParamInput, 500, Null)
    cmd.Parameters.Append cmd.CreateParameter("filetype3", adVarChar, adParamInput, 500, Null)
    cmd.Parameters.Append cmd.CreateParameter("ErroCode", adVarChar, adParamOutput, 100)
    cmd.Execute
    If cmd.Parameters("ErroCode").Value <> "" Then
       MsgBox "Error Code: " & cmd.Parameters("ErroCode").Value
    End If
    Set cmd = Nothing
Next

Screen.MousePointer = 0
End Sub


Public Sub EnviaFtpXML(Tela As Form, FtpSite$, FTPUser$, FtpPass$, FtpFile$, FtpShortFile$, Optional FtpPasta$)

'''Dim Tot As Integer
'''Dim Linha As String
'''Dim posini As Integer
'''Dim Erroftp$
'''Erroftp$ = ""
'''
'''Tela.Inet1.AccessType = icDirect
'''Tela.Inet1.Protocol = icFTP
'''Tela.Inet1.url = FtpSite$
'''Tela.Inet1.UserName = FtpUser$
'''Tela.Inet1.Password = FtpPass$
'''
'''
'''If FtpPasta <> "" Then
'''    Tela.Inet1.Execute , "PUT " & FtpFile$ & " " & FtpPasta & "/" & FtpShortFile$ & vbCrLf
'''Else
'''    Tela.Inet1.Execute , "PUT " & FtpFile$ & " " & FtpShortFile$
'''End If


Dim tot As Integer
Dim Linha As String
Dim posini As Integer
Dim Erroftp$
Erroftp$ = ""

'
On Error GoTo trataErro:
'

Tela.Inet1.AccessType = icDirect
Tela.Inet1.Protocol = icFTP
Tela.Inet1.url = FtpSite$
Tela.Inet1.UserName = FTPUser$
Tela.Inet1.Password = FtpPass$



'If Nnull(FtpPasta$, 1) <> "" Then
'
'    Tela.Inet1.Execute , "CD /" & FtpPasta$ & "/" & vbCrLf
'
'    Do
'        X = DoEvents
'    Loop Until Tela.Inet1.StillExecuting = False
'
'End If

If FtpPasta <> "" Then
    Tela.Inet1.Execute , "PUT " & FtpFile$ & " " & FtpPasta & "/" & FtpShortFile$ & vbCrLf
Else
    Tela.Inet1.Execute , "PUT " & FtpFile$ & " " & FtpShortFile$
End If



Do
    X = DoEvents
Loop Until Tela.Inet1.StillExecuting = False


If Nnull(FtpPasta$, 1) <> "" Then

    Tela.Inet1.Execute , "CD /" & FtpPasta$ & "/"

     Do
        X = DoEvents
    Loop Until Tela.Inet1.StillExecuting = False

End If

If FtpPasta <> "" Then
    Tela.Inet1.Execute , "Size " & FtpPasta & "/" & FtpShortFile$ & vbCrLf
Else
    Tela.Inet1.Execute , "Size " & FtpShortFile$
End If

Do
    X = DoEvents
Loop Until Tela.Inet1.StillExecuting = False

Dim Result, strData As String
strData = CStr(Tela.Inet1.GetChunk(1024))

Do While Len(strData) > 0
    Result = Result & strData
    strData = Tela.Inet1.GetChunk(1024)
Loop

Tela.Inet1.Execute , "CLOSE "

Exit Sub



trataErro:
Select Case Err
    Case 35754
        Screen.MousePointer = 0
        'MsgBox "A Conexão não foi estabelecida,tente novamente", vbCritical
        Erroftp$ = "A Conexão não foi estabelecida,tente novamente"

    Case 35761
        Screen.MousePointer = 0
        'MsgBox "Tempo esgotado para conexão,tente novamente", vbCritical
        Erroftp$ = "Tempo esgotado para conexão,tente novamente"

    Case 35764
        Screen.MousePointer = 0
        'Call HangUp
        'MsgBox "Perda de Conexão,tente novamente", vbCritical
        Erroftp$ = "Perda de Conexão"

    Case 35795
        Screen.MousePointer = 0
        Erroftp$ = "A Conexão não foi estabelecida"
     
    Case 35796
        Screen.MousePointer = 0
        'Call HangUp
        Erroftp$ = "Perda de Conexão"
        
    Case Else
        'Call MsgBox(UCase(Err.Description), vbCritical)
        Erroftp$ = UCase(Err.Description)
        
End Select



End Sub

Public Sub EnviaFtp(Tela As Form, FtpSite$, FTPUser$, FtpPass$, FtpFile$, FtpShortFile$, Optional FtpPasta$)

Dim tot As Integer
Dim Linha As String
Dim posini As Integer
Dim Erroftp$
Erroftp$ = ""

'
On Error GoTo trataErro:
'

Tela.Inet1.AccessType = icDirect
Tela.Inet1.Protocol = icFTP
Tela.Inet1.url = FtpSite$
Tela.Inet1.UserName = FTPUser$
Tela.Inet1.Password = FtpPass$



If Nnull(FtpPasta$, 1) <> "" Then

    Tela.Inet1.Execute , "CD /" & FtpPasta$ & "/" & vbCrLf

    Do
        X = DoEvents
    Loop Until Tela.Inet1.StillExecuting = False

End If

Enviou_Ok = False

'If FtpPasta <> "" Then
'    Tela.Inet1.Execute , "SEND " & FtpFile$ & " " & FtpPasta & "/" & FtpShortFile$
'Else
'    Tela.Inet1.Execute , "SEND " & FtpFile$ & " " & FtpShortFile$
'End If


If FtpPasta <> "" Then
    Tela.Inet1.Execute , "PUT " & FtpFile$ & " " & FtpPasta & "/" & FtpShortFile$ & vbCrLf
Else
    Tela.Inet1.Execute , "PUT " & FtpFile$ & " " & FtpShortFile$
End If



Do
    X = DoEvents
Loop Until Tela.Inet1.StillExecuting = False


Compara_Arquivo = True

Tamanho_Arquivo_ftp = 0

If Nnull(FtpPasta$, 1) <> "" Then

    Tela.Inet1.Execute , "CD /" & FtpPasta$ & "/"

     Do
        X = DoEvents
    Loop Until Tela.Inet1.StillExecuting = False

End If

If FtpPasta <> "" Then
    Tela.Inet1.Execute , "Size " & FtpPasta & "/" & FtpShortFile$ & vbCrLf
Else
    Tela.Inet1.Execute , "Size " & FtpShortFile$
End If

Do
    X = DoEvents
Loop Until Tela.Inet1.StillExecuting = False

Dim Result, strData As String
strData = CStr(Tela.Inet1.GetChunk(1024))

Do While Len(strData) > 0
    'txtOutput.Text = txtOutput.Text & strMyFile & " file size part = " & strData & vbCrLf
    Result = Result & strData
    strData = Tela.Inet1.GetChunk(1024)
Loop

Compara_Arquivo = False
        



If FileLen(FtpFile$) = Val(Result) Then
        Executa "UPDATE " & Banco_Operador & "TB_LINE_FTP_ARQUIVOS SET DATA_POST_FTP=GETDATE(),ERRO_POST='' WHERE ARQUIVO='" & FtpFile$ & "'"
Else
    If Erroftp$ <> "" Then
        Executa "UPDATE " & Banco_Operador & "TB_LINE_FTP_ARQUIVOS SET ERRO_POST='" & Erroftp$ & "' WHERE ARQUIVO='" & FtpFile$ & "'"
    Else
        Executa "UPDATE " & Banco_Operador & "TB_LINE_FTP_ARQUIVOS SET ERRO_POST='NAO FOI POSSIVEL CONFERIR A TRANSMISSAO' WHERE ARQUIVO='" & FtpFile$ & "'"
    End If
    
End If

Tela.Inet1.Execute , "CLOSE "



Exit Sub



trataErro:
Select Case Err
    Case 35754
        Screen.MousePointer = 0
        'MsgBox "A Conexão não foi estabelecida,tente novamente", vbCritical
        Erroftp$ = "A Conexão não foi estabelecida,tente novamente"

    Case 35761
        Screen.MousePointer = 0
        'MsgBox "Tempo esgotado para conexão,tente novamente", vbCritical
        Erroftp$ = "Tempo esgotado para conexão,tente novamente"

    Case 35764
        Screen.MousePointer = 0
        'Call HangUp
        'MsgBox "Perda de Conexão,tente novamente", vbCritical
        Erroftp$ = "Perda de Conexão"

    Case 35795
        Screen.MousePointer = 0
        Erroftp$ = "A Conexão não foi estabelecida"
     
    Case 35796
        Screen.MousePointer = 0
        'Call HangUp
        Erroftp$ = "Perda de Conexão"
        
    Case Else
        'Call MsgBox(UCase(Err.Description), vbCritical)
        Erroftp$ = UCase(Err.Description)
        
End Select

End Sub


Public Sub Envia_HTML(Endereco As String, Rotulo As String, Mensagem As String, C_Empresa As String, Optional Anexo As String)
Dim Recip() As String
Dim cmd As New ADODB.Command
Dim Aux As Integer
Dim To_Ende As String

Screen.MousePointer = 11
Endereco = LCase(Endereco)
Recip() = Split(Endereco, ";")
For Aux = 0 To UBound(Recip)
    To_Ende = Recip(Aux)
    cmd.ActiveConnection = ConnectionString
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "" & Banco_Sgipa & "html_email"
    cmd.Parameters.Append cmd.CreateParameter("p_to", adVarChar, adParamInput, 500, "<" & To_Ende & ">")
    If CodEmpSis = "BAN" Then
        cmd.Parameters.Append cmd.CreateParameter("p_from", adVarChar, adParamInput, 500, "<sac@ciaband.com.br>")
    Else
        If C_Empresa = 1 Then
           cmd.Parameters.Append cmd.CreateParameter("p_from", adVarChar, adParamInput, 500, "<sac@ciaband.com.br>")
        Else
          If C_Empresa = 4 Then
             cmd.Parameters.Append cmd.CreateParameter("p_from", adVarChar, adParamInput, 500, "<sac@ciaband.com.br>")
           Else
             cmd.Parameters.Append cmd.CreateParameter("p_from", adVarChar, adParamInput, 500, "<sac@ciaband.com.br>")
           End If
        End If
    End If
    cmd.Parameters.Append cmd.CreateParameter("p_subject", adVarChar, adParamInput, 500, Rotulo)
    cmd.Parameters.Append cmd.CreateParameter("p_text", adVarChar, adParamInput, 20000, Null)
    cmd.Parameters.Append cmd.CreateParameter("p_html", adVarChar, adParamInput, 20000, Mensagem)
    cmd.Execute
    Set cmd = Nothing
    
Next

Screen.MousePointer = 0
End Sub

Public Function ObtemTermLog(Cod_Empresa As Long) As Long
Dim Rs As New ADODB.Recordset
SQl = "select TRANSP_TERMLOG from " & Banco_Sgipa & "tb_empresas where autonum=" & Cod_Empresa
PRSet Rs, SQl
If Not Rs.EOF Then
    ObtemTermLog = Nnull(Rs!TRANSP_TERMLOG, 0)
Else
    ObtemTermLog = 0
End If
Rs.Close

End Function
Public Function ObtemRazaoTermlog() As String
Dim Rs As New ADODB.Recordset

SQl = "select TRANSP_TERMLOG from " & Banco_Sgipa & "tb_empresas "
PRSet Rs, SQl
If Not Rs.EOF Then
   SQl = "select razao from " & Banco_Operador & "tb_cad_transportadoras where autonum=" & Nnull(Rs!TRANSP_TERMLOG, 0)
   PRSet Rs, SQl
   If Not Rs.EOF Then
      ObtemRazaoTermlog = Nnull(Rs!Razao, 1)
   Else
      ObtemRazaoTermlog = ""
   End If
End If
Rs.Close

End Function


Public Function BuscaIndicador(Optional Importador As Long = 0, Optional Despachante As Long = 0, Optional Lote As Long = 0, Optional Viagem As String = "") As Long
Dim Data_Validade As Date
Dim Tbh As New ADODB.Recordset
Dim wImportador As Long
Dim wDespachante As Long
Dim wCaptador As Long
Dim wNvocc As Long
Dim wVinculo As Long
Dim wBL As Long
Dim wViagem As String
Dim DtasIndeferida As Byte
Dim wHubPort As Byte
Dim Rs As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
Dim Rs3 As New ADODB.Recordset
Dim tbh1 As New ADODB.Recordset
Dim wLista As Long
Dim wCargaPatio As Byte
Dim wLista_Anterior As Long
Dim wTipo_Vinculo As String
Dim BlMaster As Long
Dim Empresa_BL As Long
Dim i As Integer
Dim wSobra As Long
Dim wAcordo As Long
Dim sql_LD As String
Dim LoteMaster As Long
Dim Importadores As Integer
Dim CGCs(50) As String
Dim wRegime As String
Dim wTipoCC As Byte
Dim Regra As String

wLista_Anterior = 0
wTipo_Vinculo = ""
wImportador = Importador
wDespachante = Despachante


If Lote <> 0 Then
    SQl = "select nvl(a.flag_carga_patio,0) flag_carga_patio, nvl(a.autonum_lista,0) lista, nvl(a.flag_hubport,0) flag_hubport"
    SQl = SQl & ", nvl(a.flag_dtas_indeferida,0) flag_dtas_indeferida, b.razao, nvl(a.bl_master,0) bl_master, a.viagem"
    SQl = SQl & ", nvl(a.importador,0) importador,nvl(a.despachante,0) despachante ,nvl(a.captador,0) captador"
    SQl = SQl & ", nvl(a.audit_av,0) audit_av, a.primeira_entrada, a.patio, a.tipo_documento, a.bl_master, a.viagem"
 '   Sql = Sql & " from " & Banco_Sgipa & "tb_bl a, " & Banco_Sgipa & "tb_cad_parceiros b where"
 '   Sql = Sql & " a.autonum=" & Lote
 '   Sql = Sql & " and a.captador = b.autonum(+)"
    
    SQl = SQl & " from " & Banco_Sgipa & "tb_bl a LEFT JOIN " & Banco_Sgipa & "tb_cad_parceiros b "
    SQl = SQl & " ON  a.captador = b.autonum "
    SQl = SQl & " WHERE a.autonum=" & Lote
    
    PRSet Rs, SQl
    If Rs.EOF Then
        MsgBox "Indicador nao cadastrado"
        Exit Function
    Else
        wImportador = Rs!Importador
        wDespachante = Rs!Despachante
        wCaptador = Rs!Captador
        BlMaster = Nnull(Rs!BL_Master, 0)
        wViagem = Rs!Viagem
    End If
    SQl = "SELECT nvl(cod_empresa,0) cod_empresa FROM " & Banco_Operador & "TB_PATIOS WHERE autonum = " & Nnull(Rs!Patio, 0)
    PRSet Rs3, SQl
    Empresa_BL = Rs3!Cod_Empresa
    Rs3.Close
Else
    Empresa_BL = Cod_Empresa
    wViagem = Viagem
End If


sql_LD = ""
SQl = "select b.grupo from " & Banco_Sgipa & "dte_tb_viagens a, " & Banco_Sgipa & "dte_tb_armazens b where a.arm_atracacaop = b.code and a.viagem='" & wViagem & "'"
PRSet Rs, SQl
If Not Rs.EOF Then
    If UCase(Nnull(Rs!Grupo, 1)) = "TECONDI" Then sql_LD = " and a.tc=1"
    If UCase(Nnull(Rs!Grupo, 1)) = "RODRIMAR" Then sql_LD = " and a.ro=1"
    If UCase(Nnull(Rs!Grupo, 1)) = "PUBLICO" Then sql_LD = " and a.cp=1"
    If UCase(Nnull(Rs!Grupo, 1)) = "LIBRA" Then sql_LD = " and a.lb=1"
    If UCase(Nnull(Rs!Grupo, 1)) = "SANTOS BRASIL" Then sql_LD = " and a.sb=1"
End If
Rs.Close

SQl = "select indicador_sobra, indicador_acordo from " & Banco_Sgipa & "tb_empresas where autonum=" & Empresa_BL
PRSet Rs, SQl
If Not Rs.EOF Then
    wSobra = Nnull(Rs!indicador_sobra, 0)
    wAcordo = Nnull(Rs!indicador_acordo, 0)
End If
Rs.Close


LoteMaster = 0
Do While BlMaster <> 0
    SQl = "select a.autonum,a.bl_master,a.importador,b.cgc from " & Banco_Sgipa & "tb_bl a, " & Banco_Sgipa & "tb_cad_parceiros b where a.importador = b.autonum and a.autonum=" & BlMaster
    PRSet Rs, SQl
    wNvocc = Nnull(Rs!Importador, 0)
    BlMaster = Nnull(Rs!BL_Master, 0)
    LoteMaster = Nnull(Rs!autonum, 0)
    Rs.Close
Loop


i = 0
For i = 1 To 50
    CGCs(i) = ""
Next
Importadores = 0

If LoteMaster <> 0 Then
    SQl = "select a.importador,b.cgc from " & Banco_Sgipa & "tb_bl a, " & Banco_Sgipa & "tb_cad_parceiros b where a.importador = b.autonum and a.bl_master=" & LoteMaster
    PRSet Rs, SQl
    Do While Not Rs.EOF
        If Nnull(Rs!CGC, 1) <> "" Then
            For i = 1 To 50
                If Len(Rs!CGC) = 18 Then
                    If CGCs(i) = "" Then CGCs(i) = Mid(Rs!CGC, 1, 18): Importadores = Importadores + 1: Exit For
                    If CGCs(i) = Mid(Rs!CGC, 1, 18) Then Exit For
                Else
                    If CGCs(i) = "" Then CGCs(i) = Mid(Rs!CGC, 1, 14): Importadores = Importadores + 1: Exit For
                    If CGCs(i) = Mid(Rs!CGC, 1, 14) Then Exit For
                End If
            Next
        End If
        Rs.MoveNext
    Loop
    Rs.Close
Else
    Importadores = 1
End If
wVinculo = 0

'verifica regime dos cntr's
wRegime = "HH"
wTipoCC = 0
If Lote <> 0 Then
    SQl = "select regime,tipo from " & Banco_Sgipa & "tb_cntr_bl a, tb_amr_cntr_bl b"
    SQl = SQl & " where a.autonum = b.cntr"
    SQl = SQl & " and b.bl=" & Lote
    PRSet Rs, SQl
    If Rs.EOF Then
        SQl = "select * from " & Banco_Sgipa & "tb_carga_solta a"
        SQl = SQl & " where"
        SQl = SQl & " a.bl=" & Lote
        PRSet Rs, SQl
        If Rs.EOF Then
            Rs.Close
            Rem MsgBox "Nao foi possivel realizar o vinculo automatico" & vbCr & "pois o sistema nao encontrou conteiner ou carga solta vinculados a este BL" & vbCr & "Verifique e comunique o setor de Captação"
            Exit Function
        Else
            wRegime = "CS"
        End If
    Else
        wRegime = Rs!Regime
        Do While Not Rs.EOF
            If Nnull(Rs!Tipo, 0) = 3 Or Nnull(Rs!Tipo, 0) = 4 Then wTipoCC = Rs!Tipo
            Rs.MoveNext
        Loop
    End If
    If wTipoCC = 3 Or wTipoCC = 4 Then wRegime = "CS"
End If
'------------------------------------------------


If wRegime = "CS" Then
    wVinculo = wAcordo
    GoTo Sair
End If

Regra = ""
If (wRegime = "HH" Or wRegime = "PH") And wNvocc = 0 Then
    GoSub Procura_Importador
    If wVinculo <> 0 Then BuscaIndicador = wVinculo: GoTo Sair
    If wVinculo = 0 Then GoSub Procura_Importador_Despachante
    If wVinculo <> 0 Then BuscaIndicador = wVinculo: GoTo Sair
    If wVinculo = 0 Then BuscaIndicador = wSobra: GoTo Sair
End If

If (wRegime = "HH" Or wRegime = "PH") And wNvocc <> 0 Then
    GoSub Procura_Importador
    If wVinculo <> 0 Then BuscaIndicador = wVinculo: GoTo Sair
    If wVinculo = 0 Then GoSub Procura_Importador_Despachante
    If wVinculo <> 0 Then BuscaIndicador = wVinculo: GoTo Sair
    If wVinculo = 0 Then GoSub Procura_Importador_Nvocc
    If wVinculo <> 0 Then BuscaIndicador = wVinculo: GoTo Sair
    If wVinculo = 0 Then GoSub Procura_Nvocc
    If wVinculo <> 0 Then BuscaIndicador = wVinculo: GoTo Sair
    If wVinculo = 0 Then BuscaIndicador = wSobra: GoTo Sair
End If

If (wRegime = "PP" Or wRegime = "HP") Then
    If Importadores > 1 Then
        If wVinculo = 0 Then GoSub Procura_Nvocc
        If wVinculo <> 0 Then BuscaIndicador = wVinculo: GoTo Sair
        If wVinculo = 0 Then BuscaIndicador = wSobra: GoTo Sair
    Else
        GoSub Procura_Importador
        If wVinculo <> 0 Then BuscaIndicador = wVinculo: GoTo Sair
        GoSub Procura_Importador_Despachante
        If wVinculo <> 0 Then BuscaIndicador = wVinculo: GoTo Sair
        If wVinculo = 0 Then GoSub Procura_Nvocc
        If wVinculo <> 0 Then BuscaIndicador = wVinculo: GoTo Sair
        If wVinculo = 0 Then BuscaIndicador = wSobra: GoTo Sair
    End If
End If

If wVinculo = 0 Then BuscaIndicador = wSobra


Sair:
SQl = "INSERT INTO " & Banco_Sgipa & "TB_HIST_VINCULO_INDICADOR ("
SQl = SQl & " LOTE, IMP, DESP,"
SQl = SQl & " IND, IND_VINCULADO, DATA,"
SQl = SQl & " USUARIO, NVOCC, REGRA, regime"
SQl = SQl & " ) VALUES ("
SQl = SQl & Lote
SQl = SQl & "," & wImportador
SQl = SQl & "," & wDespachante
SQl = SQl & "," & wCaptador
SQl = SQl & "," & wVinculo
SQl = SQl & ", GETDATE()"
SQl = SQl & "," & Cod_Usuario
SQl = SQl & "," & wNvocc
SQl = SQl & ",'" & Regra & "'"
SQl = SQl & ",'" & wRegime & "'"
SQl = SQl & ")"
Executa SQl
Exit Function

Procura_Importador:
Regra = "I"
If wImportador <> 0 Then
    SQl = "select cgc from " & Banco_Sgipa & "tb_cad_parceiros where autonum=" & wImportador
    PRSet Rs2, SQl
    If Nnull(Rs2!CGC, 1) <> "" Then
        SQl = "select a.autonum_indicador"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_cad_proposta a, tb_cad_parceiros b, tb_cad_parceiros c"
        SQl = SQl & " where (a.autonum_terminal = " & Empresa_BL & " or a.autonum_terminal = 3) and a.inativo=0 and"
        SQl = SQl & " a.autonum_indicador = c.autonum and nvl(c.flag_captador,0)=1"
        SQl = SQl & " and a.autonum_importador = b.autonum"
        SQl = SQl & " and nvl(a.autonum_despachante,0)=0"
        If Len(Rs2!CGC) = 18 Then
            SQl = SQl & " and substr(b.cgc,1,18)='" & Left(Rs2!CGC, 18) & "'"
        Else
            SQl = SQl & " and substr(b.cgc,1,14)='" & Left(Rs2!CGC, 14) & "'"
        End If
        SQl = SQl & " " & sql_LD
        SQl = SQl & " group by a.autonum_indicador"
        PRSet Rs, SQl
        If Rs.RecordCount = 1 Then
            wVinculo = Rs!AUTONUM_INDICADOR
        End If
        Rs.Close
    Else
        SQl = "select a.autonum_indicador"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_cad_proposta a, " & Banco_Sgipa & "tb_cad_parceiros c"
        SQl = SQl & " where (a.autonum_terminal = " & Empresa_BL & " or a.autonum_terminal = 3) and a.inativo=0 and"
        SQl = SQl & " a.autonum_indicador = c.autonum and nvl(c.flag_captador,0)=1"
        SQl = SQl & " and a.autonum_importador=" & wImportador
        SQl = SQl & " " & sql_LD
        PRSet Rs, SQl
        If Rs.RecordCount = 1 Then
            wVinculo = Rs!AUTONUM_INDICADOR
        End If
        Rs.Close
    End If
    Rs2.Close
End If
Return

Procura_Importador_Despachante:
Regra = "ID"
If wImportador <> 0 And wDespachante <> 0 Then
    SQl = "select cgc from " & Banco_Sgipa & "tb_cad_parceiros where autonum=" & wImportador
    PRSet Rs2, SQl
    SQl = "select cgc from " & Banco_Sgipa & "tb_cad_parceiros where autonum=" & wDespachante
    PRSet Rs3, SQl
    If Nnull(Rs2!CGC, 1) <> "" And Nnull(Rs3!CGC, 1) <> "" Then
        SQl = "select a.autonum_indicador"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_cad_proposta a, tb_cad_parceiros b, tb_cad_parceiros c, tb_cad_parceiros d"
        SQl = SQl & " where (a.autonum_terminal = " & Empresa_BL & " or a.autonum_terminal = 3) and a.inativo=0 and"
        SQl = SQl & " a.autonum_indicador = d.autonum and nvl(d.flag_captador,0)=1 and"
        SQl = SQl & " a.autonum_importador = b.autonum"
        SQl = SQl & " and a.autonum_despachante = c.autonum"
        If Len(Rs2!CGC) = 18 Then
            SQl = SQl & " and substr(b.cgc,1,18)='" & Left(Rs2!CGC, 18) & "'"
        Else
            SQl = SQl & " and substr(b.cgc,1,14)='" & Left(Rs2!CGC, 14) & "'"
        End If
        If Len(Rs3!CGC) = 18 Then
            SQl = SQl & " and substr(c.cgc,1,10)='" & Left(Rs3!CGC, 10) & "'"
        Else
            SQl = SQl & " and substr(c.cgc,1,14)='" & Left(Rs3!CGC, 14) & "'"
        End If
        SQl = SQl & " " & sql_LD
        SQl = SQl & " group by a.autonum_indicador"
        PRSet Rs, SQl
        If Rs.RecordCount = 1 Then
            wVinculo = Rs!AUTONUM_INDICADOR
        End If
        Rs.Close
    Else
        SQl = "select a.*"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_cad_proposta a, " & Banco_Sgipa & "tb_cad_parceiros d"
        SQl = SQl & " where (a.autonum_terminal = " & Empresa_BL & " or a.autonum_terminal = 3) and a.inativo=0 and"
        SQl = SQl & " a.autonum_indicador = d.autonum and nvl(d.flag_captador,0)=1 and"
        SQl = SQl & " a.autonum_importador=" & wImportador
        SQl = SQl & " and a.autonum_despachante=" & wDespachante
        SQl = SQl & " " & sql_LD
        PRSet Rs, SQl
        If Not Rs.EOF Then
            wVinculo = Rs!AUTONUM_INDICADOR
        End If
        Rs.Close
    End If
    Rs2.Close
    Rs3.Close
End If
Return


Procura_Importador_Nvocc:
Regra = "IN"
'procura tabela especifica
'nvocc + importador
If wNvocc <> 0 And wImportador <> 0 Then
    SQl = "select cgc from " & Banco_Sgipa & "tb_cad_parceiros where autonum=" & wImportador
    PRSet Rs2, SQl
    SQl = "select cgc from " & Banco_Sgipa & "tb_cad_parceiros where autonum=" & wNvocc
    PRSet Rs3, SQl
    If Nnull(Rs2!CGC, 1) <> "" And Nnull(Rs3!CGC, 1) <> "" Then
        SQl = "select a.autonum_indicador"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_cad_proposta a, tb_cad_parceiros b, tb_cad_parceiros c"
        SQl = SQl & " where (a.autonum_terminal = " & Empresa_BL & " or a.autonum_terminal = 3) and a.inativo=0 and"
        SQl = SQl & " a.autonum_importador = b.autonum"
        SQl = SQl & " and a.autonum_indicador = c.autonum and nvl(c.flag_captador,0)=1"
        If Len(Rs2!CGC) = 18 Then
            SQl = SQl & " and substr(b.cgc,1,18)='" & Left(Rs2!CGC, 18) & "'"
        Else
            SQl = SQl & " and substr(b.cgc,1,14)='" & Left(Rs2!CGC, 14) & "'"
        End If
        If Len(Rs3!CGC) = 18 Then
            SQl = SQl & " and substr(c.cgc,1,18)='" & Left(Rs3!CGC, 18) & "'"
        Else
            SQl = SQl & " and substr(c.cgc,1,14)='" & Left(Rs3!CGC, 14) & "'"
        End If
        SQl = SQl & " " & sql_LD
        SQl = SQl & " group by a.autonum_indicador"
        PRSet Rs, SQl
        If Rs.RecordCount = 1 Then
            wVinculo = Rs!AUTONUM_INDICADOR
        End If
        Rs.Close
    Else
        SQl = "select a.*"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_cad_proposta a, " & Banco_Sgipa & "tb_cad_parceiros d"
        SQl = SQl & " where (a.autonum_terminal = " & Empresa_BL & " or a.autonum_terminal = 3) and a.inativo=0 and"
        SQl = SQl & " a.autonum_indicador = d.autonum and nvl(d.flag_captador,0)=1 and"
        SQl = SQl & " a.autonum_importador=" & wImportador
        SQl = SQl & " and a.autonum_indicador=" & wNvocc
        SQl = SQl & " and nvl(a.autonum_despachante,0)=" & wDespachante
        SQl = SQl & " " & sql_LD
        PRSet Rs, SQl
        If Not Rs.EOF Then
            wVinculo = Rs!AUTONUM_INDICADOR
        End If
        Rs.Close
    End If
    Rs2.Close
    Rs3.Close
End If
Return


Procura_Nvocc:
Regra = "N"
'procura tabela especifica  (nao olha o grupo)
'nvocc
If wNvocc <> 0 Then
    SQl = "select cgc from " & Banco_Sgipa & "tb_cad_parceiros where autonum=" & wNvocc
    PRSet Rs2, SQl
    If Nnull(Rs2!CGC, 1) <> "" Then
        SQl = "select a.autonum_indicador"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_cad_proposta a, tb_cad_parceiros b"
        SQl = SQl & " where (a.autonum_terminal = " & Empresa_BL & " or a.autonum_terminal = 3) and a.inativo=0 and"
        SQl = SQl & " a.autonum_importador = b.autonum and nvl(b.flag_importador,0)=1"
        If Len(Rs2!CGC) = 18 Then
            SQl = SQl & " and substr(b.cgc,1,18)='" & Left(Rs2!CGC, 18) & "'"
        Else
            SQl = SQl & " and substr(b.cgc,1,14)='" & Left(Rs2!CGC, 14) & "'"
        End If
        SQl = SQl & " " & sql_LD
        SQl = SQl & " group by a.autonum_indicador"
        PRSet Rs, SQl
        If Not Rs.EOF Then
            wVinculo = Rs!AUTONUM_INDICADOR
        End If
        Rs.Close
    Else
        SQl = "select a.autonum_indicador"
        SQl = SQl & " from"
        SQl = SQl & " " & Banco_Sgipa & "tb_cad_proposta a, " & Banco_Sgipa & "tb_cad_parceiros d"
        SQl = SQl & " where (a.autonum_terminal = " & Empresa_BL & " or a.autonum_terminal = 3) and a.inativo=0 and"
        SQl = SQl & " a.autonum_importador = d.autonum and nvl(d.flag_importador,0)=1 and"
        SQl = SQl & " a.autonum_importador=" & wNvocc
        SQl = SQl & " " & sql_LD
        PRSet Rs, SQl
        If Not Rs.EOF Then
            wVinculo = Rs!AUTONUM_INDICADOR
        End If
        Rs.Close
    End If
    Rs2.Close
End If
Return

End Function
Public Function UsuarioTemSenha(ByVal Conexao As ADODB.Connection, _
                                ByVal Tabela As String, _
                                ByVal CampoUsuario As String, _
                                ByVal CampoSenha As String, _
                                ByVal usuario As String) As Boolean
  Dim strSQL As String
  Dim rsUsuario As New ADODB.Recordset
  
  strSQL = "SELECT"
  strSQL = strSQL & " " & CampoSenha
  strSQL = strSQL & " FROM "
  strSQL = strSQL & Banco_Sgipa & Tabela
  strSQL = strSQL & " WHERE"
  strSQL = strSQL & " " & CampoUsuario & " = '" & UCase(usuario) & "'"
  PRSet rsUsuario, strSQL
  
  If Not rsUsuario.EOF Then
    If Nnull(rsUsuario(CampoSenha), 1) <> "CIA.BANDEIRANTES" And Nnull(rsUsuario(CampoSenha), 1) <> "CIA.BANDEIRANTES" Then
       UsuarioTemSenha = True
    Else
       UsuarioTemSenha = False
    End If
  Else
    Err.Raise vbError + 1000, _
              "UsuarioTemSenha", _
              "Usuário não cadastrado na base!"
  End If
End Function
Public Function Verifica_Destino(Arquivo As String) As Boolean
'On Error GoTo Trata:
'Dim VL As Integer
'Dim Linha As String
'Verifica_Destino = True
'If Dir(Arquivo) = "" Then Verifica_Destino = False
'If Verifica_Destino Then
'    Open Arquivo For Input As #50
'    Close #50
'End If
'
'Exit Function
'
'Trata:
'    Verifica_Destino = False
    
    
Dim VL As Integer
Dim Linha As String

Verifica_Destino = True


If Dir(Arquivo) = "" Then Verifica_Destino = False
If Verifica_Destino Then

    
    Open Arquivo For Input As #50
    Do While Not EOF(50)
        Line Input #50, Linha
        If Linha <> "" Then VL = VL + 1
        If VL = 2 Then Exit Do
    Loop
    Close #50
    If VL = 0 Then Verifica_Destino = False
End If
    
End Function

Sub Verifica_Lote_Sigvig(Lote As Long)

    'atualiza o lote na sigvig_pacote para casos de nova desconsolidacao apos um cancelamento de desconsolidacao
    'verifica CE do Master
    SQl = "select bl.audit_sigvig_pc, bl.ce, bl.autonum, pct.lote, pct.autonum as id_pct from"
    SQl = SQl & " sgipa.dbo.tb_bl bl"
    SQl = SQl & " inner join sgipa.dbo.tb_desconsolidacao_ce dce on bl.ce = dce.ce"
    SQl = SQl & " inner join sgipa.dbo.tb_sigvig_pacote pct on dce.idce = pct.idce"
    SQl = SQl & " where bl.autonum = " & Lote
    PRSet Rs, SQl
    Do While Not Rs.EOF
        If Nnull(Rs!autonum, 0) <> Nnull(Rs!Lote, 0) Then
            SQl = "update sgipa.dbo.tb_sigvig_pacote set lote=" & Rs!autonum & " where AUTONUM=" & Rs!ID_PCT
            db.Execute SQl
        End If
        Rs.MoveNext
    Loop
    Rs.Close
    
    'verifica CE dos filhotes
    SQl = "select bl.audit_sigvig_pc, bl.ce, bl.autonum, pct.lote, pct.autonum as id_pct from"
    SQl = SQl & " sgipa.dbo.tb_bl bl"
    SQl = SQl & " inner join sgipa.dbo.tb_desconsolidacao_ce dce on bl.ce = dce.ce"
    SQl = SQl & " inner join sgipa.dbo.tb_sigvig_pacote pct on dce.idce = pct.idce"
    SQl = SQl & " where bl.bl_master = " & Lote
    PRSet Rs, SQl
    Do While Not Rs.EOF
        If Nnull(Rs!autonum, 0) <> Nnull(Rs!Lote, 0) Then
            SQl = "update sgipa.dbo.tb_sigvig_pacote set lote=" & Rs!autonum & " where AUTONUM=" & Rs!ID_PCT
            db.Execute SQl
        End If
        Rs.MoveNext
    Loop
    Rs.Close
End Sub

Public Function VerificaPendenteVista(BL As Long)
Dim TbTotal As New ADODB.Recordset
VerificaPendenteVista = 0
'Set TbTotal = New ADODB.Recordset
'OS 030338913
'Sql = "SELECT count(1) as qtde  FROM TB_SERVICOS_FATURADOS WHERE BL=" & BL & " AND NVL(SEQ_GR,0)=0 "

SQl = "select flag_brascenter from sgipa.dbo.tb_bl where autonum=" & BL
If Nnull(db.Execute(SQl).Fields(0), 0) = 1 Then Exit Function


SQl = "SELECT count(1) as qtde  "
SQl = SQl & " FROM " & Banco_Sgipa & "TB_SERVICOS_FATURADOS A "
SQl = SQl & " INNER JOIN ("
SQl = SQl & "   SELECT BL FROM " & Banco_Sgipa & "TB_GR_PRE_CALCULO WHERE FORMAPAGAMENTO=2 GROUP BY BL ) b"
SQl = SQl & " ON A.BL = B.BL "
SQl = SQl & " WHERE A.BL=" & BL & " AND NVL(A.SEQ_GR,0)=0 "
'OS 030338913

PRSet TbTotal, SQl
'.Open "SELECT count(1) as qtde  FROM TB_SERVICOS_FATURADOS WHERE BL=" & BL & " AND NVL(SEQ_GR,0)=0 ", db, 1, 1
If Not TbTotal.EOF Then
   
   VerificaPendenteVista = Nnull(TbTotal!QTDE, 0)
End If

TbTotal.Close
End Function


Public Function VerificaFormaPagamento(BL As Long) As Integer
Dim irs As New ADODB.Recordset
Dim tbLISTA1 As New ADODB.Recordset
Dim tbcondi As New ADODB.Recordset

VerificaFormaPagamento = 0
SQl = "select nvl(forma_pagamento,0) as fp from tb_listas_precos where autonum=(select nvl(autonum_lista,0) from " & Banco_Sgipa & "tb_bl where autonum=" & BL & ")"
PRSet irs, SQl, 0
If Not irs.EOF Then
    VerificaFormaPagamento = Nnull(irs!FP, 0)
    SQl = Empty
    SQl = SQl & "SELECT "
    SQl = SQl & "B.FLAG_COND_IMPORTADOR, "
    SQl = SQl & "C.FLAG_COND_DESPACHANTE, "
    SQl = SQl & "D.FLAG_COND_INDICADOR, "
    SQl = SQl & "D.FLAG_COND_NVOCC, "
    SQl = SQl & "B.FLAG_IMPORTADOR, "
    SQl = SQl & "C.FLAG_DESPACHANTE, "
    SQl = SQl & "D.FLAG_CAPTADOR, "
    SQl = SQl & "D.FLAG_NVOCC "
    
    'Sql = Sql & "FROM "
    'Sql = Sql & "" & Banco_Sgipa & "TB_BL A, "
    'Sql = Sql & "" & Banco_Sgipa & "TB_CAD_PARCEIROS B, "
    'Sql = Sql & "" & Banco_Sgipa & "TB_CAD_PARCEIROS C, "
    'Sql = Sql & "" & Banco_Sgipa & "TB_CAD_PARCEIROS D "
    'Sql = Sql & "WHERE "
    'Sql = Sql & "A.AUTONUM = " & BL & " "
    'Sql = Sql & "AND A.IMPORTADOR = B.AUTONUM(+) "
    'Sql = Sql & "AND A.DESPACHANTE = C.AUTONUM(+) "
    'Sql = Sql & "AND A.CAPTADOR = D.AUTONUM(+) "
        
    SQl = SQl & " FROM "
    SQl = SQl & " " & Banco_Sgipa & "TB_BL A  "
    SQl = SQl & " left join " & Banco_Sgipa & "TB_CAD_PARCEIROS B on A.IMPORTADOR = B.AUTONUM "
    SQl = SQl & " left join " & Banco_Sgipa & "TB_CAD_PARCEIROS C on A.DESPACHANTE = C.AUTONUM "
    SQl = SQl & " left join " & Banco_Sgipa & "TB_CAD_PARCEIROS D on A.CAPTADOR = D.AUTONUM "
    SQl = SQl & " WHERE "
    SQl = SQl & " A.AUTONUM = " & BL & " "
    
    PRSet tbLISTA1, SQl, 0
    If Not tbLISTA1.EOF Then
        'Se Importador
        If tbLISTA1!flag_importador = 1 And Nnull(tbLISTA1!flag_cond_importador, 0) <> 0 Then
           VerificaFormaPagamento = Nnull(tbLISTA1!flag_cond_importador, 0)
        End If
        'Se Despachante
        If tbLISTA1!flag_despachante = 1 And Nnull(tbLISTA1!flag_Cond_despachante, 0) <> 0 Then
           VerificaFormaPagamento = Nnull(tbLISTA1!flag_Cond_despachante, 0)
        End If
        'Se NVOCC
        If tbLISTA1!Flag_Nvocc = 1 And Nnull(tbLISTA1!flag_Cond_nvocc, 0) <> 0 Then
           VerificaFormaPagamento = Nnull(tbLISTA1!flag_Cond_nvocc, 0)
        End If
        'Se Indicador
        If tbLISTA1!flag_captador = 1 And Nnull(tbLISTA1!flag_Cond_indicador, 0) <> 0 Then
           VerificaFormaPagamento = Nnull(tbLISTA1!flag_Cond_indicador, 0)
        End If
    End If
End If
irs.Close
SQl = "Select forma_Pagamento from " & Banco_Sgipa & "tb_gr_bl where autonum in (SELECT Max(Autonum) FROM " & Banco_Sgipa & "tb_gr_bl  WHERE bl = " & BL & " and forma_Pagamento=2 and  (tb_gr_bl.status_gr = 'IM' OR tb_gr_bl.status_gr = 'GE')) "
PRSet irs, SQl, 0
If Not irs.EOF Then
   VerificaFormaPagamento = Nnull(irs!Forma_Pagamento, 0)
End If

Set irs = Nothing


'VerificaFormaPagamento = 2


End Function


Public Function Clip(TEXTO) As String
    If TEXTO <> "" Then
    Clip = Replace(Replace(Replace(Replace(TEXTO, "-", ""), "_", ""), ".", ""), "/", "")
    Else
    Clip = ""
    End If
End Function

Sub Grava_GR_Pendente(Lote, CNTR, CS, Ordem, Optional Origem As String, Optional EnviaEmail As Boolean = False)
    Dim Rs As New ADODB.Recordset
    Dim Mandar_Email As Boolean
    Dim wData_Ordem As String
    
    Mandar_Email = EnviaEmail
    
    'se cliente faturado enviar email de aviso para saida de conteiner sem GR
    'Obs: mandar e-mail apenas no primeiro registro
    SQl = "select autonum from " & Banco_Sgipa & "tb_saidas_grs_pendentes where bl=" & Lote
    PRSet Rs, SQl
    If Not Rs.EOF Then Mandar_Email = False
    Rs.Close
    
    wData_Ordem = ""
    SQl = "select data_ordem from " & Banco_Sgipa & "tb_ordem_carregamento where autonum=" & Ordem
    PRSet Rs, SQl
    If Not Rs.EOF Then
        If Not IsNull(Rs!data_ordem) Then wData_Ordem = Format(Rs!data_ordem, "dd/mm/yyyy hh:mm")
    End If
    Rs.Close
    
    If CNTR <> 0 Then
        SQl = "select autonum from " & Banco_Sgipa & "tb_saidas_grs_pendentes where bl=" & Lote & " and cntr=" & CNTR
    Else
        SQl = "select autonum from " & Banco_Sgipa & "tb_saidas_grs_pendentes where bl=" & Lote & " and cs=" & CS
    End If
    PRSet Rs, SQl
    If Rs.EOF Then
        SQl = "insert into " & Banco_Sgipa & "tb_saidas_grs_pendentes (autonum,bl,cntr,cs,ordem)"
        SQl = SQl & " values"
        SQl = SQl & " (" & Banco_Sgipa & "seq_grs_pendentes.nextval,"
        SQl = SQl & Lote
        SQl = SQl & "," & CNTR
        SQl = SQl & "," & CS
        SQl = SQl & "," & Ordem
        SQl = SQl & ")"
        Executa SQl  'z4
    Else
        SQl = "update " & Banco_Sgipa & "tb_saidas_grs_pendentes set ordem=" & Ordem & " where autonum=" & Rs!autonum
        Executa SQl  'z4
    End If
    Rs.Close
    
    If Origem = "REGISTRO" Then
        SQl = Empty
        SQl = SQl & "SELECT "
        SQl = SQl & "MAX(VALIDADE_GR) AS FREE_TIME "
        SQl = SQl & "FROM "
        SQl = SQl & "" & Banco_Sgipa & "TB_GR_BL "
        SQl = SQl & "WHERE BL = '" & Lote & "'"
        PRSet Rs, SQl
        If Not Rs.EOF Then
            If Not IsNull(Rs!free_time) Then
                If Now < CDate(Format(Rs!free_time, "dd/mm/yyyy hh:mm:ss")) Then
                    Mandar_Email = False
                End If
            Else
                Mandar_Email = False
            End If
        End If
        Rs.Close
    End If
    
    'gera email
    If Mandar_Email Then
        SQl = "select * from " & Banco_Sgipa & "tb_parametros where not email is null"
        PRSet Rs, SQl
        If Not Rs.EOF Then
            '1o endereco de email
            SQl = "insert into " & Banco_Sgipa & "tb_email_gr_faturada (autonum,para,endereco,assunto,mensagem,data_envio,enviado)"
            SQl = SQl & " values"
            SQl = SQl & "("
            SQl = SQl & "" & Banco_Sgipa & "seq_tb_email_gr_faturada.nextval"
            SQl = SQl & ",'" & Rs!Email & "'"
            SQl = SQl & ",'" & Rs!Email & "'"
            SQl = SQl & ",'Liberacao de Carga sem GR e GRs vencidas - Cliente Faturado'"
            
            Mensagem = "Lote = " & Lote & " " & vbCrLf & " " & vbCrLf
            
            SQL2 = "select numero,b.razao nm_importador,c.razao nm_despachante "
            
            'SQL2 = SQL2 & " from " & Banco_Sgipa & "tb_bl a, " & Banco_Sgipa & "tb_cad_parceiros b, " & Banco_Sgipa & "tb_cad_parceiros c"
            'SQL2 = SQL2 & " where"
            'SQL2 = SQL2 & " a.importador = b.autonum(+)"
            'SQL2 = SQL2 & " and a.despachante = c.autonum(+)"
            'SQL2 = SQL2 & " and a.autonum=" & Lote
            
            SQL2 = SQL2 & " from " & Banco_Sgipa & "tb_bl a "
            SQL2 = SQL2 & " left join  " & Banco_Sgipa & "tb_cad_parceiros b on a.importador = b.autonum "
            SQL2 = SQL2 & " left join " & Banco_Sgipa & "tb_cad_parceiros c on a.despachante = c.autonum "
            SQL2 = SQL2 & " where "
            SQL2 = SQL2 & " a.autonum=" & Lote

            Dim Rs2 As New ADODB.Recordset
            PRSet Rs2, SQL2
            If Not Rs2.EOF Then
                Mensagem = Mensagem & "Registro    = " & wData_Ordem & " " & vbCrLf
                Mensagem = Mensagem & "BL          = " & Rs2!Numero & " " & vbCrLf
                Mensagem = Mensagem & "Importador  = " & Rs2!Nm_Importador & " " & vbCrLf
                Mensagem = Mensagem & "Despachante = " & Rs2!Nm_Despachante & " " & vbCrLf & " " & vbCrLf
            End If
            Rs2.Close
            
            'Conteiner
            If CNTR <> 0 Then
                Dim Rs3 As New ADODB.Recordset
                SQL2 = "select a.id_conteiner, b.autonum from " & Banco_Sgipa & "tb_cntr_bl a, tb_amr_cntr_bl b where a.autonum = b.cntr and b.cntr=" & CNTR & " and b.bl=" & Lote
                PRSet Rs2, SQL2
                If Not Rs2.EOF Then
                    Mensagem = Mensagem & "Conteiner = " & Rs2!Id_Conteiner
                End If
                Rs2.Close
            End If
            'procura Carga Solta
            If CS <> 0 Then
                Mensagem = Mensagem & "Para retirada de Carga Solta"
            End If
                
            SQl = SQl & ",'" & Mensagem & "'"
            SQl = SQl & ",GETDATE()"
            SQl = SQl & ",0"
            SQl = SQl & ")"
            
            Executa SQl  'z4
            
            '2o endereco de email
            If Not IsNull(Rs!Email2) Then
                SQl = "insert into " & Banco_Sgipa & "tb_email_gr_faturada (autonum,para,endereco,assunto,mensagem,data_envio,enviado)"
                SQl = SQl & " values"
                SQl = SQl & "("
                SQl = SQl & "" & Banco_Sgipa & "seq_tb_email_gr_faturada.nextval"
                SQl = SQl & ",'" & Rs!Email2 & "'"
                SQl = SQl & ",'" & Rs!Email2 & "'"
                SQl = SQl & ",'Liberacao de Carga sem GR e GRs vencidas - Cliente Faturado'"
                SQl = SQl & ",'" & Mensagem & "'"
                SQl = SQl & ",GETDATE()"
                SQl = SQl & ",0"
                SQl = SQl & ")"
                
                Executa SQl  'z4
            End If
            
            '3o endereco de email
            If Not IsNull(Rs!Email3) Then
                SQl = "insert into " & Banco_Sgipa & "tb_email_gr_faturada (autonum,para,endereco,assunto,mensagem,data_envio,enviado)"
                SQl = SQl & " values"
                SQl = SQl & "("
                SQl = SQl & "" & Banco_Sgipa & "seq_tb_email_gr_faturada.nextval"
                SQl = SQl & ",'" & Rs!Email3 & "'"
                SQl = SQl & ",'" & Rs!Email3 & "'"
                SQl = SQl & ",'Liberacao de Carga sem GR e GRs vencidas - Cliente Faturado'"
                SQl = SQl & ",'" & Mensagem & "'"
                SQl = SQl & ",GETDATE()"
                SQl = SQl & ",0"
                SQl = SQl & ")"
                
                Executa SQl  'z4
            End If
            
            Rs.MoveNext
        End If
        Rs.Close
    End If
'--------------------------------------------------------------


End Sub
'=============================================================================

Public Function GetIP() As String
Dim Listing()       As MIB_IPNETROW
Dim RET             As Long
Dim cnt             As Long
Dim bBytes()        As Byte
Dim sRET            As String
Dim bTemp()         As Byte
ReDim bTemp(0 To 3)
GetIpNetTable ByVal 0&, RET, False
If RET <= 0 Then Exit Function
ReDim bBytes(0 To RET - 1) As Byte
GetIpNetTable bBytes(0), RET, False
CopyMemory RET, bBytes(0), 4
If RET > 0 Then ReDim Listing(0 To RET - 1) As MIB_IPNETROW
For cnt = 0 To RET - 1
    CopyMemory Listing(cnt), bBytes(4 + 24 * cnt), 24
    CopyMemory bTemp(0), Listing(cnt).dwAddr, 4
    sRET = Empty
    'sRet = sRet & " Item " + CStr(Listing(Cnt).dwIndex)
    sRET = sRET & ConvertAddressToString(bTemp(), 4)
    'sRet = sRet & " Physical address " + ConvertAddressToString(Listing(Cnt).bPhysAddr, Listing(Cnt).dwPhysAddrLen)
    Select Case Listing(cnt).dwType
        Case 4 'Static
            'sRet = sRet & " type: Static"
        Case 3 'Dynamic
            'sRet = sRet & " type: Dynamic"
        Case 2 'Invalid
            'sRet = sRet & " type: Invalid"
        Case 1 'Other
            'sRet = sRet & " type: Other"
    End Select
Next cnt
GetIP = sRET
Erase Listing
Erase bBytes
Erase bTemp
RET = Empty
cnt = Empty
sRET = Empty
End Function
Private Function ConvertAddressToString(bArray() As Byte, lLength As Long) As String
Dim cnt As Long
For cnt = 0 To lLength - 1
    ConvertAddressToString = ConvertAddressToString + CStr(bArray(cnt)) + "."
Next cnt
ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
cnt = Empty
End Function

Public Function DBConnect(ByRef Cn As ADODB.Connection, Optional ByVal DBPass As String, Optional ByVal DBUserName As String, Optional NomeIni As String = "SGIPA", Optional BaseTeste As Boolean = False, Optional Mostra As Boolean) As Boolean


Dim Tbh As ADODB.Recordset
Dim lRET As Boolean
Dim Caminho As String

On Error GoTo DBCErr:

Caminho = StrReverse(Mid$(StrReverse(App.Path), InStr(StrReverse(App.Path), "\")))

If BaseTeste Then
   Open Caminho & NomeIni & ".TES" For Input As #1
Else
   Open Caminho & NomeIni & ".INI" For Input As #1
End If

Line Input #1, TrabDsn
Line Input #1, Base_Sap
Line Input #1, Banco_Em_Uso '
Line Input #1, Banco_Operador '
Line Input #1, Banco_Sgipa '
Line Input #1, Banco_Redex '
Line Input #1, Banco_Fatura '
Line Input #1, Banco_Cliente '
Line Input #1, Tem_Sid
Line Input #1, NotaNFE
Close #1

NotaNFE = UCase(NotaNFE)


Select Case UCase(Banco_Cliente)
Case "MICROLED"
    UsuBanco = "SGIPA"
    PasBanco = "SGIPA"
    CodEmpSis = "TEC"
Case "BANDEIRANTES"
    UsuBanco = "SGIPA"
    PasBanco = "CIABAND"
    CodEmpSis = "BAN"
Case "TESTEBAND"
        UsuBanco = "SA"
        PasBanco = "Tranc#M2017"
        CodEmpSis = "BAN"
Case "HOMOLOGBAND"
        UsuBanco = "SGIPA"
        PasBanco = "BuYG14@$z!@$!"
        CodEmpSis = "BAN"
Case "BAND_DEIC"
        UsuBanco = "SGIPA"
        PasBanco = "BuYG14@$z!@$!"
        CodEmpSis = "BAN"
End Select
'Close #1
 
 If Banco_Em_Uso = "SQLSERVER" Then
    sEsquema = "" & Banco_Sgipa
    LOCAL_REPORT = "RPTS_SQL"
 Else
  sEsquema = "" & Banco_Sgipa
    LOCAL_REPORT = "RPTS"
 End If

BD_O = Left(Banco_Operador, Len(Banco_Operador) - 1)
BD_I = Left(Banco_Sgipa, Len(Banco_Sgipa) - 1)


  If Banco_Em_Uso = "ORACLE" Then
     Banco_TECTER = "TECTER."
  Else
      Banco_TECTER = "TECTER.DBO."
  End If

If Mostra Then
  Form_Senha_Calc.Label4.Caption = Caminho & NomeIni & ".INI"
  Form_Senha_Calc.Label4.Refresh
  DoEvents
End If

If Not Cn Is Nothing Then
    If Cn.State = adStateOpen Then Cn.Close
    Set Cn = Nothing
End If

Set Cn = New ADODB.Connection

With Cn
    .IsolationLevel = adXactReadCommitted
    .Mode = adModeShareDenyNone
        
     ConnectionString = "Provider=sqloledb;Data Source=" & TrabDsn & ";Initial Catalog=" & UsuBanco & ";User Id=" & UsuBanco & ";Password=" & PasBanco
    ConnectionStringAdo = ConnectionString
    
    db.CommandTimeout = 1800

    .ConnectionString = ConnectionString
    
    .CursorLocation = adUseClient
    
    .Open

End With
ConnectionStringAdo = ConnectionString
Banco_Session = ""

lRET = True
GoTo DBCExit:
DBCErr:
 While Not FreeFile = 1
     Close (FreeFile - 1)
 Wend
 MsgBox Err.Description + " " + ConnectionString
 
 
 
 
 Err.Clear
 If Not Cn Is Nothing Then Cn.Errors.Clear
 End
DBCExit:
DBConnect = lRET
lRET = Empty
End Function

Public Function DBDisconnect(ByRef Cn As ADODB.Connection)
On Error GoTo DBDErr:
If Not Cn Is Nothing Then
    If Cn.State = adStateOpen Then Cn.Close
    Set Cn = Nothing
End If
GoTo DBDExit:
DBDErr:
MsgBox Err.Description
Err.Clear
Cn.Errors.Clear
DBDExit:
End Function


Sub Grava_Historico_Pacote(Num_Pacote As String, Dt_Envio As Date, Usuario_Envio As Long, Viagem As String, Evento As Integer, BL As Long, Id_Conteiner As String, Obs_Receita As String)
SQl = Empty
SQl = SQl & " INSERT "
SQl = SQl & " INTO "
SQl = SQl & " " & Banco_Sgipa & "TB_AUDITORIA_PACOTE ( "
SQl = SQl & " AUTONUM, "
SQl = SQl & " NUM_PACOTE, "
SQl = SQl & " DT_ENVIO, "
SQl = SQl & " USUARIO_ENVIO, "
SQl = SQl & " VIAGEM, "
SQl = SQl & " EVENTO, "
SQl = SQl & " BL, "
SQl = SQl & " ID_CONTEINER, "
SQl = SQl & " OBS_RECEITA, "
SQl = SQl & " MAQUINA_REDE, "
SQl = SQl & " USUARIO_REDE "
SQl = SQl & " ) VALUES ( " & Banco_Sgipa & "SEQ_AUDITORIA_PACOTE.NEXTVAL, "
SQl = SQl & " '" & Num_Pacote & "', "
SQl = SQl & " TO_DATE ('" & Dt_Envio & "','dd/mm/yy hh24:mi:ss'), "
SQl = SQl & " " & Usuario_Envio & ", "
SQl = SQl & " '" & Viagem & "', "
SQl = SQl & " " & Evento & ", "
SQl = SQl & " " & BL & ", "
SQl = SQl & " '" & Id_Conteiner & "', "
SQl = SQl & " '" & Obs_Receita & "', "
SQl = SQl & " '" & Trim(Left$(Limpa32(Maquina_Rede), 15)) & "', "
SQl = SQl & " '" & Trim(Left$(Limpa32(Usuario_Rede), 15)) & "' "
SQl = SQl & " )"
Executa SQl
End Sub


Public Sub MovePacoteOperador(Arquivo As String)
On Error GoTo MPOErr:
Dim oRS      As New ADODB.Recordset
Dim sARQDEST As String
Dim xEXT     As String
If Not Dir(Arquivo) = Empty Then
    If Banco_Em_Uso = "ORACLE" Then
        PRSet oRS, "SELECT DIR_ENTRADA AS PASTA, '.'||TRIM(TO_CHAR(NUM_RECINTO,'000')) AS RECINTO FROM " & Banco_Operador & "DTE_TB_PARAMETROS_SISTEMA", 0
    Else
        SQl = "SELECT DIR_ENTRADA AS PASTA, '.'||"
        SQl = SQl & " RTRIM(LTRIM(TO_CHAR(NUM_RECINTO,'000'))) AS RECINTO "
        SQl = SQl & " FROM " & Banco_Operador & "DTE_TB_PARAMETROS_SISTEMA"
        PRSet oRS, SQl
    End If
    
    If Not Dir(oRS!Pasta, vbDirectory) = Empty Then
        xEXT = Right$(Arquivo, 4)
        sARQDEST = oRS!Pasta & IIf(Right$(oRS!Pasta, 1) = "\", Empty, "\") & Mid$(Arquivo, InStrRev(Arquivo, "\") + 1)
        sARQDEST = Replace(sARQDEST, xEXT, oRS!Recinto)
        If Not Dir(sARQDEST) <> Empty Then
            FileCopy Arquivo, sARQDEST
        End If
    End If
    oRS.Close
End If
GoTo MPOExit:
MPOErr:
MsgBox Err.Description
Err.Clear

'Resume

MPOExit:
sARQDEST = Empty
Set oRS = Nothing
End Sub

Sub Grava_Historico_Pacote_Resp(Num_Pacote As String, Dt_Envio As Date, Usuario_Envio As Long, Viagem As String, Evento As Integer, BL As Long, Id_Conteiner As String, Obs_Receita As String)
Dim oTB As New ADODB.Recordset
Dim SQl As String
PRSet oTB, "SELECT AUTONUM FROM " & Banco_Sgipa & "TB_AUDITORIA_PACOTE WHERE NUM_PACOTE_RESP = '" & Num_Pacote & "' AND USUARIO_ENVIO = " & Usuario_Envio & " AND EVENTO = " & Evento & " AND BL = " & BL & " AND OBS_RECEITA = '" & Obs_Receita & "'", 0
If oTB.EOF Then
    SQl = Empty
    SQl = SQl & "INSERT INTO "
    SQl = SQl & "" & Banco_Sgipa & "TB_AUDITORIA_PACOTE ( "
    SQl = SQl & "AUTONUM, "
    SQl = SQl & "NUM_PACOTE_RESP, "
    SQl = SQl & "DT_FLAG,USUARIO_ENVIO, "
    SQl = SQl & "VIAGEM, "
    SQl = SQl & "EVENTO, "
    SQl = SQl & "BL, "
    SQl = SQl & "ID_CONTEINER, "
    SQl = SQl & "OBS_RECEITA, "
    SQl = SQl & "MAQUINA_REDE, "
    SQl = SQl & "USUARIO_REDE "
    SQl = SQl & ") VALUES ("
    SQl = SQl & "" & Banco_Sgipa & "SEQ_AUDITORIA_PACOTE.NEXTVAL, "
    SQl = SQl & "'" & Num_Pacote & "', "
    SQl = SQl & "TO_DATE ('" & Dt_Envio & "','dd/mm/yy hh24:mi:ss'), "
    SQl = SQl & "" & Usuario_Envio & ", "
    SQl = SQl & "'" & Viagem & "', "
    SQl = SQl & "" & Evento & ", "
    SQl = SQl & "" & BL & ", "
    SQl = SQl & "'" & Id_Conteiner & "', "
    SQl = SQl & "'" & Obs_Receita & "', "
    SQl = SQl & "'" & Trim(Left$(Limpa32(Maquina_Rede), 15)) & "', "
    SQl = SQl & "'" & Trim(Left$(Limpa32(Usuario_Rede), 15)) & "' )"
    Executa SQl
End If
oTB.Close
Set oTB = Nothing
SQl = Empty
End Sub



Sub Verifica_Saida_Total_Bl(AutonumCntr As Long, AutonumCs As Long)

On Error GoTo Trata_Erro:

Dim SaidaTotal As Boolean
Dim tb1 As New ADODB.Recordset, tbC As New ADODB.Recordset, tbcs As New ADODB.Recordset

If AutonumCntr <> 0 Then
    'Set tb1 = New ADODB.Recordset
    tb1.Open "Select b.autonum as Lote from " & Banco_Sgipa & "TB_CNTR_BL a," & Banco_Sgipa & "TB_BL b," & Banco_Sgipa & "TB_AMR_CNTR_BL c where c.cntr=a.autonum and c.bl=b.autonum and b.flag_ativo=1 and a.autonum=" & AutonumCntr, db, 3, 3
    If Not tb1.EOF Then
        While Not tb1.EOF
            'Procura amarracao de Algum Cntr no terminal
            SaidaTotal = True
            Set tbC = New ADODB.Recordset
            tbC.Open "Select a.id_conteiner from " & Banco_Sgipa & "TB_CNTR_BL a," & Banco_Sgipa & "TB_BL b," & Banco_Sgipa & "TB_AMR_CNTR_BL c where c.cntr=a.autonum and c.bl=b.autonum and a.flag_terminal=1 and (a.flag_historico=0) and b.autonum=" & tb1!Lote, db, 3, 3
            If Not tbC.EOF Then
                SaidaTotal = False
            End If
            'Procura carga Solta no Terminal
            Set tbcs = New ADODB.Recordset
            tbcs.Open "Select autonum from " & Banco_Sgipa & "TB_CARGA_SOLTA where bl=" & tb1!Lote & " and flag_terminal=1 and (flag_historico=0)", db, 3, 3
            If Not tbcs.EOF Then
                SaidaTotal = False
            End If
            If SaidaTotal Then
                SQl = "UPDATE " & Banco_Sgipa & "TB_BL SET FLAG_HISTORICO=1,ULTIMA_SAIDA=GETDATE() WHERE AUTONUM=" & tb1!Lote
                Executa SQl
                tb1.MoveLast
            End If
            tb1.MoveNext
        Wend
    End If
End If
If AutonumCs <> 0 Then
    SaidaTotal = True
    Set tb1 = New ADODB.Recordset
    tb1.Open "Select bl as lote from " & Banco_Sgipa & "TB_CARGA_SOLTA where autonum=" & AutonumCs, db, 3, 3
    If Not tb1.EOF Then
        'Verifica se existe Algum Cntr
        Set tbC = New ADODB.Recordset
        tbC.Open "Select b.autonum from " & Banco_Sgipa & "TB_CNTR_BL a," & Banco_Sgipa & "TB_BL b," & Banco_Sgipa & "TB_AMR_CNTR_BL c where c.cntr=a.autonum and c.bl=b.autonum and a.flag_terminal=1 and (a.flag_historico=0 or a.flag_historico is null) and b.autonum=" & tb1!Lote, db, 3, 3
        If Not tbC.EOF Then
            SaidaTotal = False
        End If
        'Procura carga Solta no Terminal
        Set tbcs = New ADODB.Recordset
        tbcs.Open "Select autonum from " & Banco_Sgipa & "TB_CARGA_SOLTA where bl=" & tb1!Lote & " and flag_terminal=1 and (flag_historico=0 or flag_historico is null)", db, 3, 3
        If Not tbcs.EOF Then
            SaidaTotal = False
        End If
        If SaidaTotal Then
            SQl = "UPDATE " & Banco_Sgipa & "TB_BL SET FLAG_HISTORICO=1,ULTIMA_SAIDA=GETDATE() WHERE AUTONUM=" & tb1!Lote
            Executa SQl
        End If
    End If
End If

GoTo Sair:

Trata_Erro:
MsgBox Err.Description
Err.Clear
SubErro = True

Sair:
On Error GoTo 0


End Sub
Public Function RemoveAcentos(TEXTO As String) As String
Dim a As Long
Dim sSTR As String
Dim sTXT As String
sSTR = "ÁÀÃÂÄáàãâäÉÈÊËéèêëÍÌÎÏíìîïÓÒÕÔÖóòõôöÚÙÛÜúùûüçÇ'"
sTXT = "AAAAAaaaaaEEEEeeeeIIIIiiiiOOOOOoooooUUUUuuuucC'"
For a = 1 To Len(sSTR)
    TEXTO = Replace(TEXTO, Mid$(sSTR, a, 1), Mid$(sTXT, a, 1))
Next a
RemoveAcentos = TEXTO
a = Empty
sSTR = Empty
sTXT = Empty
End Function

Sub Verifica_Chegada_Total(CS As Long)
Dim tbGateCs As New ADODB.Recordset
Dim tbcs As New ADODB.Recordset
Dim Chegou As Long, Captou As Long
tbGateCs.Open "Select Sum(quantidade) as Qtde from " & Banco_Sgipa & "TB_CARGA_GATE_CS where autonumcs=" & CS, db, 3, 3
Chegou = Nnull(tbGateCs!QTDE, 0)
tbcs.Open "Select Quantidade from " & Banco_Sgipa & "TB_CARGA_SOLTA where autonum=" & CS, db, 3, 3
Captou = Nnull(tbcs!quantidade, 0)
If Chegou >= Captou Then
    SQl = "UPDATE " & Banco_Sgipa & "TB_CARGA_SOLTA SET FLAG_CHEGADA_TOTAL=1,STATUS_ST='Z',FLAG_TERMINAL=1 WHERE AUTONUM=" & CS
    Executa SQl
End If
End Sub

Public Function IsPartLoad(BL As String, Optional Acresce As String) As Boolean
Dim SQl As String
Dim oRS As New ADODB.Recordset
On Error GoTo IPLErr:
'If Len(BL) < 5 Then GoTo IPLExit:
If BL = Empty Then GoTo IPLExit:
If Not IsNumeric(BL) = True Then GoTo IPLExit:


SQl = " SELECT COUNT(LOTE) AS QTD FROM "
SQl = SQl & " (Select Distinct Bl As Lote  From  "
SQl = SQl & "           (Select * From " & Banco_Sgipa & "Tb_Cntr_Bl Where Regime = 'HH') A,  "
SQl = SQl & " " & Banco_Sgipa & "Tb_Amr_Cntr_Bl B,  "
SQl = SQl & " (Select * From " & Banco_Sgipa & "Tb_Bl Where Flag_Ativo = 1) C, "

SQl = SQl & " (Select B.Cntr From (Select * From " & Banco_Sgipa & "Tb_Bl Where Flag_Ativo = 1) A,"
SQl = SQl & " " & Banco_Sgipa & "Tb_Amr_Cntr_Bl B Where"

SQl = SQl & " A.Autonum = B.Bl And B.Bl = " & BL & ") D  "

SQl = SQl & " Where  C.Autonum = B.Bl  "
SQl = SQl & " AND A.Autonum = B.Cntr  "
SQl = SQl & " AND B.Cntr = D.Cntr  )"

If Banco_Em_Uso = "SQLSERVER" Then
    SQl = SQl & "  ALIAS1 "
End If


PRSet oRS, SQl, 0
IsPartLoad = (Nnull(oRS!Qtd, 0) > 1)
oRS.Close
GoTo IPLExit:
IPLErr:
MsgBox Err.Description
Err.Clear

IPLExit:
Set oRS = Nothing
sSql = Empty
End Function


Sub Verifica_Bl_Concluido(BL As Long)


On Error GoTo Trata_Erro:

Dim Concluido        As Boolean
Dim Liberado         As Boolean
Dim Desiste_Avaria   As Boolean
Dim Vistoria_Aduaneira As Boolean
Dim Avariado         As Boolean
Dim ContCaptados     As Integer
Dim ContEntraram     As Integer
Dim ContCaptadosPier As Integer
Dim ContDesovados    As Integer
Dim CScaptados       As Integer
Dim CSentraram       As Integer
Dim Tb3              As New ADODB.Recordset
Dim Tbh              As New ADODB.Recordset
Dim DTBL             As String
Dim lMBM             As Boolean

Concluido = False

Liberado = True
Avariado = False
Motivo = Empty
'DESISTÊNCIA DA VISTORIA
PRSet Tb3, "Select nvl(FLAG_VISTORIA_ADUANEIRA,0) FLAG_VISTORIA_ADUANEIRA, NVL(flag_desiste_vistoria,0) as flag_desiste_vistoria,DT_BL,ID_SGTAP from " & Banco_Sgipa & "TB_BL where autonum=" & BL & " ", 0
Desiste_Avaria = (Nnull(Tb3!FLAG_DESISTE_VISTORIA, 0) = 1)
Desiste_Avaria = True
Vistoria_Aduaneira = (Nnull(Tb3!FLAG_VISTORIA_ADUANEIRA, 0) = 1)
DTBL = Trim(UCase(Nnull(Tb3!dt_bl, 1)))
lMBM = Not (Tb3.Fields("ID_SGTAP") = 0 Or IsNull(Tb3.Fields("ID_SGTAP")))
Tb3.Close
If lMBM Then
    Executa "UPDATE " & Banco_Sgipa & "TB_BL SET FLAG_LIBERADO_PRESENCA = 1 WHERE AUTONUM = " & BL & " "
    VERIFICABLCONCLUIDO = True
    Exit Sub
End If
'NVOCC
PRSet Tb3, "Select importador, flag_nvocc, flag_carga_patio, flag_desconsolidado from " & Banco_Sgipa & "TB_BL where autonum = " & BL & " ", 0
If Not Tb3.EOF Then
    If Nnull(Tb3!Flag_Nvocc, 0) = 1 Then
        PRSet Tbh, "Select flag_nvocc from " & Banco_Sgipa & "TB_CAD_PARCEIROS where autonum=" & Tb3!Importador & " ", 0
        If Not Tbh.EOF Then
            If Nnull(Tbh!Flag_Nvocc, 0) = 1 Then
                Liberado = False
                Motivo = "B/L está informado como sendo  NVOCC."
                Executa "UPDATE " & Banco_Sgipa & "tb_bl set motivo_presenca_negada='" & Left(Motivo, 200) & "' where autonum = " & BL
            Else
                If Nnull(Tb3!flag_desconsolidado, 0) = 0 Then
                    Executa "UPDATE " & Banco_Sgipa & "TB_BL SET FLAG_NVOCC = 0 WHERE AUTONUM = " & BL & " "
                End If
            End If
        End If
        Tbh.Close
    End If
Else
    Liberado = False
    Motivo = "Bl Não Encontrado"
End If
Tb3.Close
If Not Motivo = Empty Then GoTo VBCExit:


'Se Existe Cntr com avaria fora do Padrao 'Amassado.Arranhado.Enferrujado.Remendado
SQl = Empty
SQl = SQl & "Select "
SQl = SQl & "a.id_conteiner, "
SQl = SQl & "a.flag_avariado, "
SQl = SQl & "c.local, "
SQl = SQl & "c.tipo, "
SQl = SQl & "c.complemento, "
SQl = SQl & "a.FLAG_DESOVADO, "
SQl = SQl & "D.AUDIT_DESOVA "
SQl = SQl & "From "
SQl = SQl & "" & Banco_Sgipa & "TB_CNTR_BL a, "
SQl = SQl & "" & Banco_Sgipa & "TB_AMR_CNTR_BL b, "
SQl = SQl & "" & Banco_Sgipa & "TB_AVARIAS_CONTEINER C, "
SQl = SQl & "TB_BL D "
SQl = SQl & "Where "
SQl = SQl & "a.autonum = b.Cntr "
SQl = SQl & "and a.autonum=c.cntr "
SQl = SQl & "AND B.BL = D.AUTONUM "
SQl = SQl & "and b.bl = " & BL
SQl = SQl & " and a.regime<>'HP' and a.regime<>'PP'"
PRSet Tb3, SQl, 0
While Not Tb3.EOF
    'marcado como avariado
    If Nnull(Tb3!flag_avariado, 0) = 1 Then Avariado = True
    If Not (Nnull(Tb3!local, 0) = 32 And Nnull(Tb3!Tipo, 0) = 1 And Nnull(Tb3!COMPLEMENTO, 0) = 3) And _
    Not (Nnull(Tb3!local, 0) = 32 And Nnull(Tb3!Tipo, 0) = 2 And Nnull(Tb3!COMPLEMENTO, 0) = 3) And _
    Not (Nnull(Tb3!local, 0) = 32 And Nnull(Tb3!Tipo, 0) = 4 And Nnull(Tb3!COMPLEMENTO, 0) = 3) Then
        Avariado = True
    End If
    
'conforme orientação Marcia em 28/09/2009, havendo avaria e usuario indicou uma vistoria aduaneira, não deverá ser verificado
'o status de desistencia de vistoria.
    
    If Avariado = True And Desiste_Avaria = False And Vistoria_Aduaneira = False Then
        If Nnull(Tb3!audit_desova, 0) = 0 Or Nnull(Tb3!Flag_Desovado, 0) = 0 Then
            Liberado = False
            Motivo = "O Contêiner " & Tb3!Id_Conteiner & " contém avarias e não há nenhuma desistência de vistoria."
            Executa "UPDATE " & Banco_Sgipa & "tb_bl set motivo_presenca_negada='" & Left(Motivo, 200) & "' where autonum = " & BL
        End If
    End If
    Tb3.MoveNext
Wend
Tb3.Close
If Not Motivo = Empty Then GoTo VBCExit:

'Se Existe Cntr sem 29 enviado
'If Not Desiste_Avaria Then
    SQl = Empty
    SQl = SQl & "Select "
    SQl = SQl & "a.id_conteiner, "
    SQl = SQl & "a.flag_avariado, "
    SQl = SQl & "a.FLAG_DESOVADO, "
    SQl = SQl & "a.audit_avaria, "
    SQl = SQl & "D.AUDIT_DESOVA "
    SQl = SQl & "From "
    SQl = SQl & "" & Banco_Sgipa & "TB_CNTR_BL a, "
    SQl = SQl & "" & Banco_Sgipa & "TB_AMR_CNTR_BL b, "
    SQl = SQl & "TB_BL D "
    SQl = SQl & "Where "
    SQl = SQl & "a.autonum = b.Cntr "
    SQl = SQl & "AND B.BL = D.AUTONUM "
    SQl = SQl & "and b.bl = " & BL
    SQl = SQl & " and a.regime<>'HP' and a.regime<>'PP'"
    PRSet Tb3, SQl, 0
    While Not Tb3.EOF
        If Nnull(Tb3!AUDIT_AVARIA, 0) = 0 Then
            Liberado = False
            Motivo = "Nao consta envio de 29 para o contêiner " & Tb3!Id_Conteiner
            Executa "UPDATE " & Banco_Sgipa & "tb_bl set motivo_presenca_negada='" & Left(Motivo, 200) & "' where autonum = " & BL
        End If
        Tb3.MoveNext
    Wend
    Tb3.Close
    If Not Motivo = Empty Then GoTo VBCExit:
'End If




'Carga Solta
'If Not Desiste_Avaria Then
    SQl = Empty
    SQl = SQl & "Select "
    SQl = SQl & "item, "
    SQl = SQl & "flag_avariado, nvl(audit_avaria,0) audit_avaria"
    SQl = SQl & " from "
    SQl = SQl & "" & Banco_Sgipa & "TB_CARGA_SOLTA "
    SQl = SQl & "where "
    SQl = SQl & "bl = " & BL & " "
    SQl = SQl & "and flag_avariado = 1 "
    PRSet Tb3, SQl, 0
    If Not Tb3.EOF Then
        Avariado = True
        If Tb3!AUDIT_AVARIA <> 0 Then
            Avariado = False
        End If
        If Avariado Then
            Liberado = False
            Motivo = " O item " & Tb3!Item & " está indicado como tendo avaria e a mesma não foi enviada à DTE"
            Executa "UPDATE " & Banco_Sgipa & "tb_bl set motivo_presenca_negada='" & Left(Motivo, 200) & "' where autonum = " & BL
        End If
    End If
    Tb3.Close
    If Not Motivo = Empty Then GoTo VBCExit:
'End If


'Captados com DT
ContCaptados = 0
ContEntraram = 0
SQl = Empty
SQl = SQl & "Select "
SQl = SQl & "Count(a.autonum) as Qtdos "
SQl = SQl & "from "
SQl = SQl & "" & Banco_Sgipa & "TB_CNTR_BL a, "
SQl = SQl & "" & Banco_Sgipa & "TB_AMR_CNTR_BL b "
SQl = SQl & "where "
SQl = SQl & "a.autonum = b.cntr "
SQl = SQl & "and b.bl = " & BL & " "
SQl = SQl & "and not a.gmci is null"
PRSet Tb3, SQl, 0
ContCaptados = Nnull(Tb3!qtdos, 0)
Tb3.Close

SQl = Empty
SQl = SQl & " Select "
SQl = SQl & " Count(a.autonum) as Qtdos "
SQl = SQl & " from "
SQl = SQl & " " & Banco_Sgipa & "TB_CNTR_BL a, "
SQl = SQl & " " & Banco_Sgipa & "TB_AMR_CNTR_BL b "
SQl = SQl & " where "
SQl = SQl & " a.autonum = b.cntr "
SQl = SQl & " and b.bl = " & BL & " "
SQl = SQl & " and not a.gmci is null"
SQl = SQl & " And a.dt_entrada is not null"
PRSet Tb3, SQl, 0
ContEntraram = Nnull(Tb3!qtdos, 0)
Tb3.Close
If Nnull(ContCaptados, 0) <> Nnull(ContEntraram, 0) Then
    Concluido = False
    Liberado = False
    Motivo = "O número de Conteineres que entraram (" & ContEntraram & ") está diferente do número de Conteineres que foram captados (" & ContCaptados & ")."
    Executa "UPDATE " & Banco_Sgipa & "tb_bl set motivo_presenca_negada='" & Left(Motivo, 200) & "' where autonum = " & BL
End If

If Not Motivo = Empty Then GoTo VBCExit:

'Captados com DT e Regime Pier
ContCaptadosPier = 0
ContDesovados = 0
SQl = Empty
SQl = SQl & "Select "
SQl = SQl & "Count(a.autonum) as Qtdos "
SQl = SQl & "from "
SQl = SQl & "" & Banco_Sgipa & "TB_CNTR_BL a, "
SQl = SQl & "" & Banco_Sgipa & "TB_AMR_CNTR_BL b "
SQl = SQl & "where "
SQl = SQl & "a.autonum = b.cntr "
SQl = SQl & "and b.bl = " & BL & " "
SQl = SQl & "and not a.gmci is null "
SQl = SQl & "and (a.regime = 'HP' or a.regime = 'PP')"
PRSet Tb3, SQl, 0
ContCaptadosPier = Nnull(Tb3!qtdos, 0)
Tb3.Close

'Cntr Pier Desovado
SQl = Empty
SQl = SQl & "Select Count(a.autonum) as Qtdos"
SQl = SQl & " from " & Banco_Sgipa & "TB_CNTR_BL a"
SQl = SQl & "," & Banco_Sgipa & "TB_AMR_CNTR_BL b"
SQl = SQl & "," & Banco_Sgipa & "TB_BL c"
SQl = SQl & "," & Banco_Sgipa & "TB_AUDITORIA_PACOTE d"
SQl = SQl & " where a.autonum=b.cntr"
SQl = SQl & " and b.bl=c.autonum"
SQl = SQl & " and d.autonum=c.audit_desova"
SQl = SQl & " and b.bl=" & BL
SQl = SQl & " and a.gmci is not null"
SQl = SQl & " and (a.regime='HP' or a.regime='PP')"
SQl = SQl & " and d.dt_envio is not null"
PRSet Tb3, SQl, 0
ContDesovados = Nnull(Tb3!qtdos, 0)
Tb3.Close
If Nnull(ContDesovados, 0) <> Nnull(ContCaptadosPier, 0) Then
    Concluido = False
    Liberado = False
    Motivo = "O nº de Conteineres Pier  (" & ContCaptadosPier & ") está diferente do nº de Conteineres que foram desovados (" & ContDesovados & ")"
    Executa "UPDATE " & Banco_Sgipa & "tb_bl set motivo_presenca_negada='" & Left(Motivo, 200) & "' where autonum = " & BL
End If
If Not Motivo = Empty Then GoTo VBCExit:

'Captados com DT
SQl = Empty
SQl = SQl & "Select "
SQl = SQl & "Count(a.item) as Qtdos "
SQl = SQl & "from "
SQl = SQl & "" & Banco_Sgipa & "TB_CARGA_SOLTA a "
SQl = SQl & "where "
SQl = SQl & "a.bl = " & BL & " "
SQl = SQl & "and not a.gmvi is null"
PRSet Tb3, SQl, 0
CScaptados = Nnull(Tb3!qtdos, 0)
Tb3.Close

'Carga solta
SQl = Empty
SQl = SQl & "Select "
SQl = SQl & "Count(a.item) as Qtdos "
SQl = SQl & "from "
SQl = SQl & "" & Banco_Sgipa & "TB_CARGA_SOLTA a "
SQl = SQl & "where a.bl=" & BL & " "
SQl = SQl & "and not a.gmvi is null "
SQl = SQl & "and flag_chegada_total = 1 "
PRSet Tb3, SQl, 0
CSentraram = Nnull(Tb3!qtdos, 0)
Tb3.Close
If Nnull(CSentraram, 0) <> Nnull(CScaptados, 0) Then
    Concluido = False
    Liberado = False
    Motivo = "O nº de Cargas Soltas que entraram  (" & CSentraram & ") está diferente do nº de Cargas Soltas que foram captadas (" & CScaptados & ")"
    Executa "UPDATE " & Banco_Sgipa & "tb_bl set motivo_presenca_negada='" & Left(Motivo, 200) & "' where autonum = " & BL
End If
If Not Motivo = Empty Then GoTo VBCExit:



If CSentraram = 0 And ContEntraram = 0 Then
    Concluido = False
    Liberado = False
    Motivo = "Nao consta entrada de carga para este lote"
    Executa "UPDATE " & Banco_Sgipa & "tb_bl set motivo_presenca_negada='" & Left(Motivo, 200) & "' where autonum = " & BL
End If
If Not Motivo = Empty Then GoTo VBCExit:


Concluido = True

VBCExit:

If Concluido Then
    SQl = "UPDATE " & Banco_Sgipa & "TB_BL SET FLAG_CONCLUIDO=1, DT_CONCLUIDO=GETDATE() WHERE autonum=" & BL
    Executa SQl
Else
    SQl = "UPDATE " & Banco_Sgipa & "TB_BL SET FLAG_CONCLUIDO=0, DT_CONCLUIDO=NULL WHERE autonum=" & BL
    Executa SQl
End If

If Liberado = True Then
    SQl = "UPDATE " & Banco_Sgipa & "TB_BL SET FLAG_LIBERADO_PRESENCA = 1 WHERE autonum = " & BL & " "
    Executa SQl
Else
    SQl = "UPDATE " & Banco_Sgipa & "TB_BL SET FLAG_LIBERADO_PRESENCA = 0 WHERE autonum = " & BL & " "
    Executa SQl

    SQl = "insert into " & Banco_Sgipa & "tb_presenca_negada (data,lote,motivo) values (GETDATE()," & BL & ",'" & Left(Motivo, 100) & "')"
    Executa SQl  'z4
    
End If

On Error GoTo 0
Exit Sub

Trata_Erro:
On Error GoTo 0
SubErro = True
End Sub


Public Function Abre_Arquivo_Ini(Arq As String, campo As String) As String
On Error GoTo AAIErr:
Dim sRET As String
Dim Line As String
Dim nArq As Byte



If Not Dir(Arq) = Empty Then
    nArq = FreeFile

    Open Arq For Input As #nArq
    Do While Not EOF(nArq) = True
        Line Input #nArq, Line
        If UCase$(Trim$(Line)) = UCase$(campo) Then
            If Not EOF(nArq) = True Then
                Line Input #nArq, Line
                sRET = Trim(Right(Line, 3))
            End If
            Exit Do
        End If
    Loop
    Close #nArq
End If
Abre_Arquivo_Ini = sRET
GoTo AAIExit:
AAIErr:
MsgBox Err.Description
Err.Clear
AAIExit:
While Not FreeFile = 1
    Close (FreeFile - 1)
Wend
sRET = Empty
Line = Empty
nArq = Empty
Screen.MousePointer = vbDefault
End Function

Public Function BloqueioPorOrdemCarregamento(Ordem As Long) As Boolean
On Error GoTo BPOCErr:
Dim oRS     As New ADODB.Recordset
Dim SQl     As String
Dim lRET    As Boolean
Dim sMS     As String
sMS = Empty
If Not Ordem = Empty Then
    SQl = Empty
    SQl = SQl & " SELECT "
    SQl = SQl & " MAX(NVL(C.FLAG_BLOQUEIO,0)) AS BDCN, "
    SQl = SQl & " MAX(NVL(C.FLAG_BLOQUEIO_MANUAL,0)) AS BMCN, "
    SQl = SQl & " MAX(NVL(E.FLAG_BLOQUEIO,0)) AS BDBL, "
    SQl = SQl & " MAX(NVL(E.FLAG_BLOQUEIO_MANUAL,0)) AS BMBL "
    SQl = SQl & " FROM "
    SQl = SQl & " " & Banco_Sgipa & "TB_ORDEM_CARREGAMENTO A, "
    SQl = SQl & " " & Banco_Sgipa & "TB_REGISTRO_SAIDA_CNTR B, "
    SQl = SQl & " " & Banco_Sgipa & "TB_CNTR_BL C, "
    SQl = SQl & " " & Banco_Sgipa & "TB_AMR_CNTR_BL D, "
    SQl = SQl & " " & Banco_Sgipa & "TB_BL E "
    SQl = SQl & " WHERE "
    SQl = SQl & " A.AUTONUM = B.ORDEM_CARREG "
    SQl = SQl & " AND B.CNTR = C.AUTONUM "
    SQl = SQl & " AND B.CNTR = D.AUTONUM "
    SQl = SQl & " AND D.BL = E.AUTONUM "
    SQl = SQl & " AND E.FLAG_ATIVO = 1 "
    SQl = SQl & " AND A.AUTONUM = '" & Ordem & "' "
    PRSet oRS, SQl, 0
    If Not oRS.EOF Then
        If oRS.Fields("BDCN") > 0 Then sMS = sMS & "BLOQUEIO POR CONTEINER, VIA DTE" & vbCrLf
        If oRS.Fields("BMCN") > 0 Then
            If CodEmpSis = "BAN" Then
                sMS = sMS & "BLOQUEIO POR CONTEINER, VIA BANDEIRANTES" & vbCrLf
            Else
              If Cod_Empresa = 1 Then sMS = sMS & "BLOQUEIO POR CONTEINER, VIA CIA.BANDEIRANTES" & vbCrLf
              If Cod_Empresa = 2 Then sMS = sMS & "BLOQUEIO POR CONTEINER, VIA CIA.BANDEIRANTES" & vbCrLf
            End If
        End If
        If oRS.Fields("BDBL") > 0 Then sMS = sMS & "BLOQUEIO POR B/L, VIA DTE" & vbCrLf
        If oRS.Fields("BMBL") > 0 Then
            If CodEmpSis = "BAN" Then
                sMS = sMS & "BLOQUEIO POR B/L, VIA BANDEIRANTES" & vbCrLf
            Else
                If Cod_Empresa = 1 Then sMS = sMS & "BLOQUEIO POR B/L, VIA CIA.BANDEIRANTES" & vbCrLf
                If Cod_Empresa = 2 Then sMS = sMS & "BLOQUEIO POR B/L, VIA CIA.BANDEIRANTES" & vbCrLf
            End If
        End If
    End If
    oRS.Close
    SQl = Empty
    SQl = SQl & " SELECT "
    SQl = SQl & " MAX(NVL(E.FLAG_BLOQUEIO,0)) AS BDBL, "
    SQl = SQl & " MAX(NVL(E.FLAG_BLOQUEIO_MANUAL,0)) AS BMBL "
    SQl = SQl & " FROM "
    SQl = SQl & " " & Banco_Sgipa & "TB_ORDEM_CARREGAMENTO A, "
    SQl = SQl & " " & Banco_Sgipa & "TB_REGISTRO_SAIDA_CS B, "
    SQl = SQl & " " & Banco_Sgipa & "TB_CARGA_SOLTA C, "
    SQl = SQl & " " & Banco_Sgipa & "TB_BL E "
    SQl = SQl & " WHERE "
    SQl = SQl & " A.AUTONUM = B.ORDEM_CARREG "
    SQl = SQl & " AND B.CS = C.AUTONUM "
    SQl = SQl & " AND C.BL = E.AUTONUM "
    SQl = SQl & " AND E.FLAG_ATIVO = 1 "
    SQl = SQl & " AND A.AUTONUM = '" & Ordem & "' "
    PRSet oRS, SQl, 0
    If Not oRS.EOF Then
        If oRS.Fields("BDBL") > 0 Then sMS = sMS & "BLOQUEIO POR B/L, VIA DTE" & vbCrLf
        If oRS.Fields("BMBL") > 0 Then
            If CodEmpSis = "BAN" Then
                sMS = sMS & "BLOQUEIO POR B/L, VIA BANDEIRANTES" & vbCrLf
            Else
                If Cod_Empresa = 1 Then sMS = sMS & "BLOQUEIO POR B/L, VIA CIA.BANDEIRANTES" & vbCrLf
                If Cod_Empresa = 2 Then sMS = sMS & "BLOQUEIO POR B/L, VIA CIA.BANDEIRANTES" & vbCrLf
            End If
        End If
    End If
    oRS.Close
End If
lRET = (Len(Trim(sMS)) > 0)
If (Len(Trim(sMS)) > 0) Then MsgBox sMS, vbOKOnly + vbInformation, "ALERTA DE SISTEMA"
BloqueioPorOrdemCarregamento = lRET
GoTo BPOCExit:
BPOCErr:
MsgBox Err.Description
Err.Clear
BPOCExit:
Set oRS = Nothing
SQl = Empty
sMS = Empty
lRET = Empty
End Function

Public Function Calculo_Digito_DsiManual(Numero As String) As Byte
'formato inicial do numero deve ser 99/999999

'N -AASSSSSSS - d
'
'N - diferenciador para cálculo,  não compõe o número, será sempre:  5
'A - ano corrente;
'S - seqüencial no ano;
'D - dígito verificador.
'
'Cálculo do dígito = módulo 11.
'
'ex: número 3 / 5978 - d
'
'Cálculo para 503005978 (multiplicar cada algarismo por 1 a 9, soma-se os resultados, divide o total por 11,
'e o resto da divisão, será o D)
'
'D = (5x1) + (0x2) + (3x3) + (0x4) + (0x5) + (5x6) + (9x7) + (7x8) + (8x9) = 235
'235/11 = 21,         Resto = 4
'
'logo d = 4
'
'Portanto a  composição do número da DSI = 03/005978-4
' 
'Obs. caso o resto seja 10, o digito devera ser 0 (zero).
' 
'ex. 03/005975-d
'composição do numero p/ calculo:
'503005975
'd=(5x1) + (0x2) + (3x3) + (0x4) + (0x5) + (5x6) + (9x7) + (7x8) + (5x9) = 208
'208/11 = 18   --- resto = 10
'logo Digito = 0
' 
'Numero DSI = 3 / 5975 - 0#

Dim Base As String
Dim D As Long
Base = "5" & Mid(Numero, 1, 2) & Mid(Numero, 4, 6)
D = 0
D = D + Val(Mid(Base, 1, 1)) * 1
D = D + Val(Mid(Base, 2, 1)) * 2
D = D + Val(Mid(Base, 3, 1)) * 3
D = D + Val(Mid(Base, 4, 1)) * 4
D = D + Val(Mid(Base, 5, 1)) * 5
D = D + Val(Mid(Base, 6, 1)) * 6
D = D + Val(Mid(Base, 7, 1)) * 7
D = D + Val(Mid(Base, 8, 1)) * 8
D = D + Val(Mid(Base, 9, 1)) * 9
D = D Mod 11
If D = 10 Then D = 0

Calculo_Digito_DsiManual = D

End Function

Public Function Calcula_Digito_Codesp(DT As String, CNTR As String, GMCI As String, Tam As Integer) As Byte
On Error GoTo CDCErr:
Dim a()     As String
Dim L1      As String
Dim L2      As String
Dim S1      As Long
Dim S2      As Long
Dim P1()    As Long
Dim P2()    As Long
Dim i       As Integer
Dim j       As Integer
Dim VA()    As Integer
Dim N       As Byte
ReDim a(1 To 26) As String
ReDim VA(1 To 26) As Integer
ReDim P1(1 To 12) As Long
ReDim P2(1 To 12) As Long
a(1) = "A"
a(2) = "B"
a(3) = "C"
a(4) = "D"
a(5) = "E"
a(6) = "F"
a(7) = "G"
a(8) = "H"
a(9) = "I"
a(10) = "J"
a(11) = "K"
a(12) = "L"
a(13) = "M"
a(14) = "N"
a(15) = "O"
a(16) = "P"
a(17) = "Q"
a(18) = "R"
a(19) = "S"
a(20) = "T"
a(21) = "U"
a(22) = "V"
a(23) = "W"
a(24) = "X"
a(25) = "Y"
a(26) = "Z"
VA(1) = 10
VA(2) = 12
VA(3) = 13
VA(4) = 14
VA(5) = 15
VA(6) = 16
VA(7) = 17
VA(8) = 18
VA(9) = 19
VA(10) = 20
VA(11) = 21
VA(12) = 23
VA(13) = 24
VA(14) = 25
VA(15) = 26
VA(16) = 27
VA(17) = 28
VA(18) = 29
VA(19) = 30
VA(20) = 31
VA(21) = 32
VA(22) = 34
VA(23) = 35
VA(24) = 36
VA(25) = 37
VA(26) = 38
L1 = Left$(DT, 6) & Left$(GMCI, 6)
L2 = Left$(CNTR, 10) & Trim(Str(Tam))
For i = 1 To 12
    S1 = S1 + (Val(Mid$(L1, i, 1)) * 2 ^ (i - 1))
    If i <= 4 Then
        For j = 1 To 26
            If a(j) = Mid$(L2, i, 1) Then
                Exit For
            End If
        Next j
        S2 = S2 + (VA(j) * 2 ^ (i - 1))
    Else
        S2 = S2 + (Val(Mid$(L2, i, 1)) * 2 ^ (i - 1))
    End If
Next i
N = CInt(Right$(CStr((S1 + S2) Mod 11), 1))
GoTo CDCExit:
CDCErr:
N = Empty
MsgBox Err.Description
Err.Clear
CDCExit:
Calcula_Digito_Codesp = N
Erase a
Erase P1
Erase P2
Erase VA
L1 = Empty
L2 = Empty
S1 = Empty
S2 = Empty
i = Empty
j = Empty
N = Empty
End Function
Public Function Extenso(vn As Double) As String

Dim j As Integer, k As Integer, n1 As Double, n2 As Double, n3 As Double, m1 As Double, M2 As Double, N As Double, M3 As Double
Dim px As String * 1, ve As String

    n1 = Fix(vn / 1000000#)
    vn = vn - n1 * 1000000#
    n2 = Fix(vn / 1000#)
    n3 = Fix(vn - n2 * 1000#)
  
    
    ve = ""
    k = 1
    Do While k < 4
        Select Case k
            Case 1: N = n1
            Case 2: N = n2
            Case 3: N = n3
        End Select
        If (k = 3) And (n1 * n3 <> 0) And (n2 = 0) Then ve = ve + " E "
        m1 = Fix(N / 100)
        j = N - m1 * 100
        M2 = Fix(j / 10)
        M3 = j - M2 * 10
        Select Case m1
            Case 9: ve = ve + " NOVECENTOS "
            Case 8: ve = ve + " OITOCENTOS "
            Case 7: ve = ve + " SETECENTOS "
            Case 6: ve = ve + " SEISCENTOS "
            Case 5: ve = ve + " QUINHENTOS "
            Case 4: ve = ve + " QUATROCENTROS "
            Case 3: ve = ve + " TREZENTOS "
            Case 2: ve = ve + " DUZENTOS "
            Case 1:
                If N = 100 Then
                    ve = ve + " CEM "
                Else
                    ve = ve + " CENTO "
                End If
        End Select
        If m1 * M2 <> 0 Then ve = ve + " E "
        Select Case M2
            Case 9: ve = ve + " NOVENTA "
            Case 8: ve = ve + " OITENTA "
            Case 7: ve = ve + " SETENTA "
            Case 6: ve = ve + " SESSENTA "
            Case 5: ve = ve + " CINQUENTA "
            Case 4: ve = ve + " QUARENTA "
            Case 3: ve = ve + " TRINTA "
            Case 2: ve = ve + " VINTE "
            Case 1: Select Case M3
                        Case 9: ve = ve + " DEZENOVE "
                        Case 8: ve = ve + " DEZOITO "
                        Case 7: ve = ve + " DEZESSETE "
                        Case 6: ve = ve + " DEZESSEIS "
                        Case 5: ve = ve + " QUINZE "
                        Case 4: ve = ve + " QUATORZE "
                        Case 3: ve = ve + " TREZE "
                        Case 2: ve = ve + " DOZE "
                        Case 1: ve = ve + " ONZE "
                        Case 0: ve = ve + " DEZ "
                    End Select
        End Select
        If (M2 <> 1) And ((m1 * M3 <> 0) Or (M2 * M3 <> 0)) Then ve = ve + " E "
        If M2 <> 1 Then
            Select Case M3
                Case 9: ve = ve + " NOVE "
                Case 8: ve = ve + " OITO "
                Case 7: ve = ve + " SETE "
                Case 6: ve = ve + " SEIS "
                Case 5: ve = ve + " CINCO "
                Case 4: ve = ve + " QUATRO "
                Case 3: ve = ve + " TRES "
                Case 2: ve = ve + " DOIS "
                Case 1: ve = ve + " UM "
            End Select
        End If
        If N <> 0 Then
            If (k = 1) And (N <> 1) Then ve = ve + "MILHOES "
            If (k = 1) And (N = 1) Then ve = ve + " MILHAO "
            If k = 2 Then ve = ve + " MIL "
        End If
        If (n1 * n2 <> 0) And (k = 1) Then ve = ve + " E "
        If (n2 * n3 <> 0) And (k = 2) Then ve = ve + " E "
        k = k + 1
    Loop
    If (n2 + n3 = 0) And (n1 <> 0) Then ve = ve + " DE "
    Extenso = ve
End Function

Public Function TiraMaskCntr(CNTR As String)
TiraMaskCntr = CNTR
End Function

Public Function MontaMaskCntr(CNTR As String)
Dim CntrBase As String
Dim IC, CLetra, CNumero, LIni, NIni As Integer

CLetra = 0: CNumero = 0: IC = 1: LIni = 0: NIni = 0

CntrBase = TiraMaskCntr(CNTR)

For IC = 1 To Len(CntrBase)
    If IsNumeric(Mid(CntrBase, IC, 1)) = False Then
        If LIni = 0 Then LIni = IC
        CLetra = CLetra + 1
    Else
        If NIni = 0 Then NIni = IC
        CNumero = CNumero + 1
    End If
Next IC
If CLetra = 0 Then MontaMaskCntr = "____"
If CLetra = 1 Then MontaMaskCntr = "___" & Mid(CntrBase, 1, CLetra)
If CLetra = 2 Then MontaMaskCntr = "__" & Mid(CntrBase, 1, CLetra)
If CLetra = 3 Then MontaMaskCntr = "_" & Mid(CntrBase, 1, CLetra)
If CLetra = 4 Then MontaMaskCntr = Mid(CntrBase, 1, CLetra)
If CNumero = 0 Then MontaMaskCntr = MontaMaskCntr & "______-_"
If CNumero = 1 Then MontaMaskCntr = MontaMaskCntr & Mid(CntrBase, NIni, CNumero) & "_____-_"
If CNumero = 2 Then MontaMaskCntr = MontaMaskCntr & Mid(CntrBase, NIni, CNumero) & "____-_"
If CNumero = 3 Then MontaMaskCntr = MontaMaskCntr & Mid(CntrBase, NIni, CNumero) & "___-_"
If CNumero = 4 Then MontaMaskCntr = MontaMaskCntr & Mid(CntrBase, NIni, CNumero) & "__-_"
If CNumero = 5 Then MontaMaskCntr = MontaMaskCntr & Mid(CntrBase, NIni, CNumero) & "_-_"
If CNumero = 6 Then MontaMaskCntr = MontaMaskCntr & Mid(CntrBase, NIni, CNumero) & "-_"
If CNumero = 7 Then MontaMaskCntr = MontaMaskCntr & Mid(CntrBase, NIni, CNumero - 1) & "-" & Right(CntrBase, 1)

End Function

Public Function Nnull(Par As Variant, Tipo As Integer, Optional RemoveEspaços As Boolean = True) As Variant
On Error Resume Next
Dim sVAR As Variant
If TypeOf Par Is ADODB.Field Then
    If Par.Type = adDBTimeStamp Then
        If IsNull(Par) Then
            Nnull = ""
        Else
            sVAR = ConverteTimeStamp(Par)
            Par = sVAR
        End If
    ElseIf IsNull(Par) = True Then
        Par = Empty
    End If
End If
If IsNull(Par) = True Then
    Nnull = IIf(Tipo = 0, 0, Empty)
Else
    Nnull = IIf(Trim(Par) <> Empty, IIf(Tipo = 0, Par, IIf(RemoveEspaços = True, Trim(Par), Par)), IIf(Tipo = 0, 0, Empty))
End If
If Nnull = "00:00:00" Then Nnull = Empty
sVAR = Empty
End Function

Public Sub Consulta_Cruzada(ByRef Matriz As MSHFlexGrid, Origem As String, Campo_Linha As String, Campo_Coluna As String, Campo_Valor As String, Criterio As String)

Matriz.Clear
Matriz.FixedCols = 0
Matriz.FixedRows = 0

Dim i As Integer
Dim X As Integer
Dim Y As Integer

Dim Rx As Integer
SQl = "select count(*) from (select distinct " & Campo_Coluna & " as fColuna from " & Origem & " where " & Criterio & ") T"
Rx = Nnull(db.Execute(SQl).Fields(0), 0)

Dim Ry As Integer
SQl = "select count(*) from (select distinct " & Campo_Linha & " as fLinha from " & Origem & " where " & Criterio & ") T"
Ry = Nnull(db.Execute(SQl).Fields(0), 0)

ReDim mx(Rx) As String
ReDim tx(Rx) As Integer
ReDim my(Ry) As String
ReDim tMatriz(Ry, Rx) As String
For R = 1 To Ry
    For c = 1 To Rx
        tMatriz(R, c) = ""
    Next
Next

SQl = "select " & Campo_Linha & " as fLinha," & Campo_Coluna & " as fColuna," & Campo_Valor & " as fValor from " & Origem & " where " & Criterio & " order by fLinha,fColuna"
Dim Rs As New ADODB.Recordset
PRSet Rs, SQl
Do While Not Rs.EOF
    'Busca Eixo X
    For X = 1 To Rx
        If Rs!fColuna = mx(X) Then Exit For
        If mx(X) = "" Then Exit For
    Next

    'Busca Eixo y
    For Y = 1 To Ry
        If Rs!fLinha = my(Y) Then Exit For
        If my(Y) = "" Then Exit For
    Next
    If mx(X) = "" Then mx(X) = Rs!fColuna: tMatriz(0, X) = mx(X)
    If Len(mx(X)) > Len(tx(X)) Then tx(X) = Len(mx(X))
    
    If my(Y) = "" Then my(Y) = Rs!fLinha:  tMatriz(Y, 0) = my(Y)

    tMatriz(Y, X) = Rs!fValor

    Rs.MoveNext
Loop
Rs.Close




Matriz.Cols = Rx + 1
Matriz.Rows = Ry + 1
For R = 0 To Ry
    Matriz.Row = R
    For c = 0 To Rx
        Matriz.Col = c
        Matriz = tMatriz(R, c)
    Next
Next

Matriz.Font.Name = "Courier New"
Matriz.Font.Size = 8
For c = 1 To Rx
    Matriz.ColWidth(c) = tx(c) * 140
Next
maior = 0
For R = 0 To Ry
    If Len(tMatriz(R, 0)) > maior Then maior = Len(tMatriz(R, 0))
Next
Matriz.ColWidth(0) = maior * 140


Matriz.FixedCols = 1
Matriz.FixedRows = 1
Matriz.AllowUserResizing = flexResizeBoth


End Sub


Public Sub Apaga_Session(Esquema As String, Maquina As String, Programa As String)
Dim Tbh As ADODB.Recordset
If Banco_Session <> "" Then
   SQl = "Select s.sid || ',' || s.serial#  from   v$session s,v$process p  Where s.paddr = P.addr  and s.username='" & Esquema & "'  and s.terminal='" & Maquina & "'  and (s.program='VB6.EXE' OR SUBSTR(s.program,1,9)='" & Mid$(Programa, 1, 9) & "') order by sql_exec_start desc"
   PRSet Tbh, SQl
   If Not Tbh.EOF Then
      If Banco_Session <> Tbh.Fields(0).Value Then Executa "ALTER SYSTEM KILL SESSION '" & Tbh.Fields(0).Value & "'"
    End If
End If
End Sub



Public Sub P_Select(ByRef RSTabela As ADODB.Recordset, P_Sql As String, Optional Tipo1 As TpDados, Optional P_Var1 As String, Optional Tipo2 As TpDados, Optional P_Var2 As String, Optional Tipo3 As TpDados, Optional P_Var3 As String, Optional Tipo4 As TpDados, Optional P_Var4 As String, Optional Tipo5 As TpDados, Optional P_Var5 As String, Optional Tipo6 As TpDados, Optional P_Var6 As String, Optional Tipo7 As TpDados, Optional P_Var7 As String, Optional Tipo8 As TpDados, Optional P_Var8 As String, Optional Tipo9 As TpDados, Optional P_Var9 As String, Optional Tipo10 As TpDados, Optional P_Var10 As String)

Dim cmd As New ADODB.Command
Dim t As Byte
Dim Tp_Sel(10) As TpDados
Dim Variavel(10) As String


 cmd.ActiveConnection = ConnectionString
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "PROC_CHRONOS_SELECT"

Tp_Sel(1) = Tipo1
Variavel(1) = P_Var1
Tp_Sel(2) = Tipo2
Variavel(2) = P_Var2
Tp_Sel(3) = Tipo3
Variavel(3) = P_Var3
Tp_Sel(4) = Tipo4
Variavel(4) = P_Var4
Tp_Sel(5) = Tipo5
Variavel(5) = P_Var5
Tp_Sel(6) = Tipo6
Variavel(6) = P_Var6
Tp_Sel(7) = Tipo7
Variavel(7) = P_Var7
Tp_Sel(8) = Tipo8
Variavel(8) = P_Var8
Tp_Sel(Sel9) = Tipo9
Variavel(9) = P_Var9

For t = 1 To 9
   If Nnull(Variavel(t), 1) <> "" Then
      cmd.Parameters.Append cmd.CreateParameter("Var_Par" & t, adVarChar, adParamInput, 500, Variavel(t))
   Else
      cmd.Parameters.Append cmd.CreateParameter("Var_Par" & t, adVarChar, adParamInput, 500, Null)
   End If
Next t

cmd.Parameters.Append cmd.CreateParameter("SSQL", adVarChar, adParamInput, 9000, P_Sql)

'IF BANCO_EM_USO = "SQLSERVER" Then
'    Tp_Sel(Sel10) = Tipo10
'    Variavel(10) = P_Var10

    'If Nnull(Variavel(10), 1) <> "" Then
     '     cmd.Parameters.Append cmd.CreateParameter("Var_Par" & 10, adVarChar, adParamInput, 500, Variavel(10))
   ' Else
    '  cmd.Parameters.Append cmd.CreateParameter("Var_Par" & 10, adVarChar, adParamInput, 500, Null)
   ' End If
'End If
cmd.Parameters.Append cmd.CreateParameter(" & RSTabela & ", adVarChar, adParamOutput, 100)
cmd.Parameters.Append cmd.CreateParameter("ErroCode", adVarChar, adParamOutput, 100)

'MsgBox P_Sql
Set RSTabela = cmd.Execute
If cmd.Parameters("ErroCode").Value <> "" Then
   MsgBox P_Sql
   MsgBox "Error Code: " & cmd.Parameters("ErroCode").Value
End If
Set cmd = Nothing
End Sub

Public Sub P_Bloqueio(Lote As Long, Motivo As String, Acao As String, usuario As String, CPF As String)

Exit Sub

Dim cmd As New ADODB.Command
Dim Variavel(10) As String

cmd.ActiveConnection = ConnectionString
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "" & Banco_Sgipa & "PROC_CHRONOS_BLOQUEIO"
If Banco_Em_Uso = "ORACLE" Then
    cmd.Parameters.Append cmd.CreateParameter("ID_LOTE", adNumeric, adParamInput, 20, Lote)
Else
   cmd.Parameters.Append cmd.CreateParameter("ID_LOTE", adInteger, adParamInput, 20, Lote)
End If

cmd.Parameters.Append cmd.CreateParameter("V_Motivo", adVarChar, adParamInput, 10, Motivo)
cmd.Parameters.Append cmd.CreateParameter("Acao", adVarChar, adParamInput, 5, Acao)
If Banco_Em_Uso = "ORACLE" Then
cmd.Parameters.Append cmd.CreateParameter("Usuario", adVarChar, adParamInput, 20, usuario)
cmd.Parameters.Append cmd.CreateParameter("Cpf", adVarChar, adParamInput, 20, CPF)
End If
cmd.Parameters.Append cmd.CreateParameter("ErroCode", adVarChar, adParamOutput, 100)

cmd.Execute
If cmd.Parameters("ErroCode").Value <> "" Then
   MsgBox cmd.Parameters("ErroCode").Value
End If
Set cmd = Nothing
End Sub


Public Sub Valor_Carga_Doc_Canc(BL As Double, tp_doc As Double)

Dim sSTR As String
sSTR = Empty
sSTR = "Insert into " & Banco_Sgipa & "TB_VALOR_CARGA_DOC_CANC ("
sSTR = sSTR & "BL, " '1
sSTR = sSTR & "MOEDA_ORIGEM_CNTR, " '2
sSTR = sSTR & "MOEDA_ORIGEM_CS, " '3
sSTR = sSTR & "VALOR_CIF_CNTR, " '4
sSTR = sSTR & "VALOR_CIF_CS, " '5
sSTR = sSTR & "DATA_BASE_CNTR, " '6
sSTR = sSTR & "DATA_BASE_CS, " '7
sSTR = sSTR & "IPI_CNTR, " '8
sSTR = sSTR & "IPI_CS, "  '9
sSTR = sSTR & "II_CNTR, " '10
sSTR = sSTR & "II_CS, " '11
sSTR = sSTR & "TAXA_CNTR, " '12
sSTR = sSTR & "TAXA_CS, " '13
sSTR = sSTR & "CIF_REAIS_CNTR, " '14
sSTR = sSTR & "CIF_REAIS_CS, " '15
sSTR = sSTR & "TIPO_DOCUMENTO )" '16
sSTR = sSTR & "SELECT BL, " '1
sSTR = sSTR & "MOEDA_ORIGEM_CNTR, " '2
sSTR = sSTR & "MOEDA_ORIGEM_CS, " '3
sSTR = sSTR & "VALOR_CIF_CNTR, " '4
sSTR = sSTR & "VALOR_CIF_CS, " '5
sSTR = sSTR & "DATA_BASE_CNTR, " '6
sSTR = sSTR & "DATA_BASE_CS, " '7
sSTR = sSTR & "IPI_CNTR, " '8
sSTR = sSTR & "IPI_CS, " '9
sSTR = sSTR & "II_CNTR, " '10
sSTR = sSTR & "II_CS, " '11
sSTR = sSTR & "TAXA_CNTR, " '12
sSTR = sSTR & "TAXA_CS, " '13
sSTR = sSTR & "CIF_REAIS_CNTR, " '14
sSTR = sSTR & "CIF_REAIS_CS, " '15
sSTR = sSTR & tp_doc & " FROM " & Banco_Sgipa & "TB_VALOR_CARGA WHERE BL = " & BL
Call Executa(sSTR)

End Sub

Public Function Executa(ByVal Instrução As String, Optional ByRef Retorno As Long, Optional MostraResultado As Boolean = False, Optional MostraErros As Boolean = True, Optional IntegraTerm As Boolean = False) As Boolean
'On Error GoTo ExecErr:
Dim msge  As String
Dim nRecs As Long
Dim oERR  As ADODB.Error
Dim RSNEXT As New ADODB.Recordset

Rem 01 - VERIFICA GETDATE(), SE CAMPO ID É AUTONUM OU NÃO

If Banco_Em_Uso = "SQLSERVER" Then
    Instrução = CONVERTE_NEXTVAL(Instrução)
End If

If Not db Is Nothing Then
    If Not db.State <> 0 Then db.Open
End If
'executa (Instrução)
With db

'MsgBox Instrução
'MsgBox Right(Instrução, 800)
If Banco_Em_Uso = "SQLSERVER" Then
    'Instrução = executa_to_date(Instrução)
    Instrução = CONVERTE_SQL(Instrução)
End If

'MsgBox Instrução
'MsgBox Right(Instrução, 900)
.Execute Instrução, nRecs
    

    Executa = True
End With
If MostraResultado = True Then
    If Not automatico And nRecs > 0 Then
        If nRecs = 1 Then
            MsgBox "Foi atualizado 1 registro.", vbOKOnly + vbInformation, "Atualização de registros"
        Else
            MsgBox "Foram atualizados " & nRecs & " registros.", vbOKOnly + vbInformation, "Atualização de registros"
        End If
    End If
End If

     


If Not IsMissing(Retorno) = True Then Retorno = nRecs
GoTo ExecExit:
ExecErr:
msge = Empty
For Each oERR In db.Errors
    msge = msge & oERR.Description & vbCrLf & vbTab & oERR.Source & vbCrLf
Next
If MostraErros = True Then MsgBox msge, vbOKOnly + vbInformation, "Tratamento de Erros"

Resume Next

db.Errors.Clear
Err.Clear
ExecExit:
nRecs = Empty
Set oERR = Nothing
End Function


Public Function PRSet(ByRef ObjetoTabela As ADODB.Recordset, ByVal Instrução As String, Optional Modo As Integer = 0, Optional CARREGAR As Boolean = False, Optional LadoCliente As Boolean = True) As Boolean
'DB É UM OBJETO ADODB CONNECTION PÚBLICO
'0=Consulta / 1=Atualização.
'Carregar = True move o ponteiro para o último registro, carregando o recordset integral.

'On Error GoTo PRErr:

Instrução = CONVERTE_SQL(Instrução)

Dim sSTR As String
Screen.MousePointer = vbHourglass
If Not ObjetoTabela Is Nothing Then  'Encerra-se o recordset em caso de erros
    sSTR = ObjetoTabela.Source
    If Not ObjetoTabela.State = 0 Then ObjetoTabela.Close
    Set ObjetoTabela = Nothing
End If
Set ObjetoTabela = New ADODB.Recordset
Set ObjetoTabela.ActiveConnection = db

If LadoCliente = True Then
    ObjetoTabela.CursorLocation = adUseClient
Else
    ObjetoTabela.CursorLocation = adUseServer
End If
If Modo = 0 Then
    ObjetoTabela.CursorType = adOpenForwardOnly
    ObjetoTabela.LockType = adLockReadOnly
Else
    ObjetoTabela.CursorType = adOpenDynamic
    ObjetoTabela.LockType = adLockBatchOptimistic
End If

If Banco_Em_Uso = "SQLSERVER" Then
    Instrução = CONVERTE_SQL(Instrução)
End If

If Not Instrução = Empty Then
    ObjetoTabela.Source = Instrução
Else
    ObjetoTabela.Source = sSTR
End If

'MsgBox Instrução
'MsgBox Right(Instrução, 900)

ObjetoTabela.Open

'If Carregar = True Then ObjetoTabela.MoveLast
                If Modo = 0 Then
    If Not ObjetoTabela.CursorLocation = adUseServer Then Set ObjetoTabela.ActiveConnection = Nothing
End If
PRSet = True
GoTo PRExit:

PRExit:
Screen.MousePointer = vbDefault
End Function


Sub Atualiza_Calculo(Lote As Long, Adicional As Long)
Dim Imposto_Total As Currency
Dim TotalGr As Currency
Dim Fator As Currency
Dim WValor As Currency
Dim Tbs  As ADODB.Recordset
Dim tb  As New ADODB.Recordset
Dim Tbbl  As ADODB.Recordset
Dim TbGr  As ADODB.Recordset
Dim TbIM  As ADODB.Recordset
Dim RsImposto As ADODB.Recordset
Dim TbTotal As ADODB.Recordset
Dim TbPosi As ADODB.Recordset

PRSet Tbbl, "Select Count(1) AS Contar From " & Banco_Sgipa & "TB_SERVICOS_Faturados where bl=" & Lote & " and seq_gr is null"
PRSet TbGr, "Select Max(Seq_Gr) as MaxGr From " & Banco_Sgipa & "TB_GR_BL where bl=" & Lote & " and STATUS_GR='GE'"
PRSet TbIM, "Select Max(Seq_Gr) as MaxGr From " & Banco_Sgipa & "TB_GR_BL where bl=" & Lote & " and STATUS_GR='IM'"
If Nnull(Tbbl!Contar, 0) > 0 Or Nnull(TbGr!MaxGr, 0) > 0 Then
   PRSet Tbs, "Select * from " & Banco_Sgipa & "TB_SERVICOS_ADICIONAIS_BL where Autonum=" & Adicional
   If Not Tbs.EOF Then
        
        PRSet TbPosi, "Select " & Banco_Sgipa & "SEQ_SERVICOS_FATURADOS.NEXTVAL as Posicao from dual"
        
        WSQL = "INSERT INTO " & Banco_Sgipa & "TB_SERVICOS_FATURADOS"
        WSQL = WSQL & " (AUTONUM,BL,SERVICO,VALOR,DESCONTO,adicional,inicio,final"
        WSQL = WSQL & ",PERIODOS,USUARIO"
        
        If Nnull(TbGr!MaxGr, 0) > 0 Then
           WSQL = WSQL & ",SEQ_GR"
        End If
        
        WSQL = WSQL & ") VALUES ("
        WSQL = WSQL & TbPosi!Posicao
        WSQL = WSQL & "," & Lote
        WSQL = WSQL & "," & Tbs!Servico
        WSQL = WSQL & "," & Str(Nnull(Tbs!Valor, 0))
        WSQL = WSQL & ",0"
        WSQL = WSQL & ",0"
        WSQL = WSQL & ",to_date('" & Format(Date, "dd/mm/yyyy") & "','dd/mm/yyyy')"
        WSQL = WSQL & ",to_date('" & Format(Date, "dd/mm/yyyy") & "','dd/mm/yyyy')"
        WSQL = WSQL & ",0"
        WSQL = WSQL & "," & Cod_Usuario
        If Nnull(TbGr!MaxGr, 0) > 0 Then
           WSQL = WSQL & "," & Nnull(TbGr!MaxGr, 0)
        End If
        WSQL = WSQL & ")"
        Executa (WSQL)
        SQl = Empty
        SQl = SQl & "INSERT INTO " & Banco_Sgipa & "TB_DESCRICAO_CALCULO "
        SQl = SQl & " (AUTONUM,Bl,SERVICO,SEQ_GR,BASECALCULO,RATEIOCALCULO,DESCRBASECALCULO,"
        SQl = SQl & " PRECOUNITARIO,BASEMINIMO,ACRESCIMO,DATAINIPERIODO,DATAFIMPERIODO,PARCELA,PARCELA_AD,PRECOCALCULO) VALUES ("
        SQl = SQl & " " & Banco_Sgipa & "SEQ_DESCRICAO_CALCULO.NEXTVAL," & Lote & "," & Tbs!Servico & ","
        If Nnull(TbGr!MaxGr, 0) > 0 Then
           SQl = SQl & Nnull(TbGr!MaxGr, 0)
        Else
           SQl = SQl & "null"
        End If
        SQl = SQl & ",0"
        SQl = SQl & ",0"
        SQl = SQl & ",''"
        SQl = SQl & ",0"
        SQl = SQl & ",0"
        SQl = SQl & ",0"
        SQl = SQl & ",NULL"
        SQl = SQl & ",NULL"
        SQl = SQl & "," & PPonto(Nnull(Tbs!Valor, 0)) & ""
        SQl = SQl & ",0"
        SQl = SQl & "," & PPonto(Nnull(Tbs!Valor, 0)) & ") "
        Executa SQl
        SQl = "Update Tb_Servicos_Adicionais_Bl Set Autonum_Servico=" & TbPosi!Posicao & " where autonum=" & Adicional
        Executa SQl
        PRSet Tbbl, "Select importador , Autonum_Lista from " & Banco_Sgipa & "TB_BL where Autonum=" & Lote
        If Not Tbbl.EOF Then
           SQl = "Select Flag_Isento_imposto from tb_cad_parceiros where autonum =" & Nnull(Tbbl!Importador, 0) & " and Flag_Isento_imposto = 0"
           PRSet RsImposto, SQl
           If Not RsImposto.EOF Then
              SQl = " SELECT a.autonum, a.valor+a.adicional+a.desconto AS valor_servico, a.servico, c.autonum autonum_imposto, c.taxa"
              SQl = SQl & " ,(SELECT SUM(taxa) FROM " & Banco_Sgipa & "tb_lp_servicos_impostos d, " & Banco_Sgipa & "tb_cad_impostos e WHERE"
              SQl = SQl & " d.id_imposto = e.autonum AND d.id_servico = a.servico and d.id_tabela=" & Nnull(Tbbl!AUTONUM_LISTA, 0) & " and d.calcular=1) Imposto_Total"
              SQl = SQl & " FROM"
              SQl = SQl & " " & Banco_Sgipa & "tb_servicos_faturados a"
              SQl = SQl & " ," & Banco_Sgipa & "tb_lp_servicos_impostos b"
              SQl = SQl & " ," & Banco_Sgipa & "tb_cad_impostos c"
              SQl = SQl & " WHERE "
              SQl = SQl & " b.id_tabela=" & Nnull(Tbbl!AUTONUM_LISTA, 0)
              SQl = SQl & " and b.calcular=1"
              SQl = SQl & " and a.servico = b.id_servico"
              SQl = SQl & " AND b.id_imposto = c.autonum"
              SQl = SQl & " AND a.bl=" & Lote
              SQl = SQl & " AND nvl(a.seq_gr,0)=" & Nnull(TbGr!MaxGr, 0)
              SQl = SQl & " AND a.servico=" & Nnull(Tbs!Servico, 0)
              
           '   Set tb = New ADODB.Recordset
              
              PRSet tb, SQl
              Do While Not tb.EOF
                   If Nnull(tb!Imposto_Total, 0) <> 0 Then
                       Fator = (100 - Nnull(tb!Imposto_Total, 0)) / 100
                       Imposto_Total = Round(Nnull(tb!valor_servico, 0) / Fator, 2)
                       WValor = Round((Imposto_Total * tb!taxa) / 100, 2)
                   Else
                       WValor = 0
                   End If
                   Set tb1 = New ADODB.Recordset
                   tb1.Open "Select count(1) as contar from " & Banco_Sgipa & "tb_servicos_faturados_impostos where autonum_servico_faturado=" & tb!autonum & " and autonum_imposto=" & tb!autonum_imposto, db, 3, 3
                   If Nnull(tb1!Contar, 0) = 0 Then
                        SQl = "insert into " & Banco_Sgipa & "tb_servicos_faturados_impostos"
                        SQl = SQl & " (autonum,autonum_servico_faturado,autonum_imposto,valor_imposto) values ("
                        SQl = SQl & " " & Banco_Sgipa & "seq_faturados_impostos.nextval"
                        SQl = SQl & " ," & tb!autonum
                        SQl = SQl & " ," & tb!autonum_imposto
                        SQl = SQl & " ,'" & WValor & "'"
                        SQl = SQl & ")"
                   Else
                     SQl = "update " & Banco_Sgipa & "tb_servicos_faturados_impostos set Valor_imposto='" & WValor & "'  Where autonum_servico_faturado = " & tb!autonum & " And autonum_imposto = " & tb!autonum_imposto
                   End If
                   Executa SQl  'z4
                   tb.MoveNext
              Loop
              tb.Close
           End If
           If Nnull(TbGr!MaxGr, 0) > 0 Then
              TotalGr = 0
              PRSet TbTotal, "SELECT SUM (valor+ adicional+desconto) as Liquido FROM TB_SERVICOS_FATURADOS WHERE BL=" & Lote & " AND nvl(SEQ_GR,0)=" & Nnull(TbGr!MaxGr, 0)
              TotalGr = Format(Nnull(TbTotal!Liquido, 0), "###,###,##0.00")
              TbTotal.Close
              PRSet TbTotal, "SELECT SUM (valor_imposto) as Liquido FROM tb_servicos_faturados a, TB_SERVICOS_FATURADOS_impostos b WHERE a.autonum = b.autonum_servico_faturado and a.BL=" & Lote & " AND nvl(a.SEQ_GR,0)=" & Nnull(TbGr!MaxGr, 0)
              TotalGr = Format(CDbl(TotalGr) + Nnull(TbTotal!Liquido, 0), "###,###,##0.00")
              SQl = "Update Tb_Gr_Bl set Valor_Gr=" & PPonto(TotalGr) & " Where Seq_Gr=" & Nnull(TbGr!MaxGr, 0)
              Executa SQl
           End If
        End If
   End If
Else
  If Nnull(TbIM!MaxGr, 0) > 0 Then
     SQl = "Update Tb_Servicos_Adicionais_Bl Set Calc_Posic=1 where autonum=" & Adicional
     Executa SQl
  End If
End If
End Sub
 Sub Apaga_Calculo(Adicional As Long)

Dim TotalGr As Currency
Dim Tbs  As ADODB.Recordset
Dim Tbbl  As ADODB.Recordset
Dim TbGr  As ADODB.Recordset
Dim TbTotal As ADODB.Recordset

PRSet Tbs, "Select * from " & Banco_Sgipa & "TB_SERVICOS_ADICIONAIS_BL where Autonum=" & Adicional
If Not Tbs.EOF Then
   PRSet Tbbl, "Select Count(1) AS Contar From " & Banco_Sgipa & "TB_SERVICOS_Faturados where servico=" & Nnull(Tbs!Servico, 0) & " and bl=" & Nnull(Tbs!BL, 0) & " and seq_gr is null"
   PRSet TbGr, "Select Max(Seq_Gr) as MaxGr From " & Banco_Sgipa & "TB_GR_BL where bl=" & Nnull(Tbs!BL, 0) & " and STATUS_GR='GE'"
   If Nnull(Tbbl!Contar, 0) > 0 Or Nnull(TbGr!MaxGr, 0) > 0 Then
        SQl = "Delete FROM " & Banco_Sgipa & "TB_SERVICOS_FATURADOS_IMPOSTOS WHERE "
        SQl = SQl & "Autonum_Servico_Faturado=" & Nnull(Tbs!Autonum_Servico, 0)
        Executa SQl
        SQl = "Delete  from " & Banco_Sgipa & "TB_SERVICOS_FATURADOS Where Autonum=" & Nnull(Tbs!Autonum_Servico, 0)
        Executa SQl
        SQl = "Delete  from " & Banco_Sgipa & "TB_DESCRICAO_CALCULO Where Seq_gr=" & Nnull(TbGr!MaxGr, 0)
        SQl = SQl & " And Servico=" & Nnull(Tbs!Servico, 0)
        Executa SQl
        TotalGr = 0
        If Nnull(TbGr!MaxGr, 0) > 0 Then
           PRSet TbTotal, "SELECT SUM (valor+ adicional+desconto) as Liquido FROM TB_SERVICOS_FATURADOS WHERE BL=" & Tbs!BL & " AND nvl(SEQ_GR,0)=" & Nnull(TbGr!MaxGr, 0)
           TotalGr = Format(Nnull(TbTotal!Liquido, 0), "###,###,##0.00")
           TbTotal.Close
           PRSet TbTotal, "SELECT SUM (valor_imposto) as Liquido FROM tb_servicos_faturados a, TB_SERVICOS_FATURADOS_impostos b WHERE a.autonum = b.autonum_servico_faturado and a.BL=" & Tbs!BL & " AND nvl(a.SEQ_GR,0)=" & Nnull(TbGr!MaxGr, 0)
           TotalGr = Format(CDbl(TotalGr) + Nnull(TbTotal!Liquido, 0), "###,###,##0.00")
           SQl = "Update Tb_Gr_Bl set Valor_Gr=" & PPonto(TotalGr) & " Where Seq_Gr=" & Nnull(TbGr!MaxGr, 0)
           Executa SQl
        End If
   End If
End If
End Sub



Function ValidaPlaca(Placa As String, Optional Estrangeiro As Byte = 0) As Boolean

Dim ERRO As Boolean
ERRO = False

If Estrangeiro = 1 Then
    Placa = Left(Placa & Space(8), 8)
        
    If (Mid(Placa, 1, 1) < "A" Or Mid(Placa, 1, 1) > "Z") Then ERRO = True
    If (Mid(Placa, 2, 1) < "A" Or Mid(Placa, 2, 1) > "Z") Then ERRO = True
    If (Mid(Placa, 3, 1) < "A" Or Mid(Placa, 3, 1) > "Z") And (Mid(Placa, 3, 1) <> " ") Then ERRO = True
    
    If (Mid(Placa, 5, 1) < "0" Or Mid(Placa, 5, 1) > "9") Then ERRO = True
    If (Mid(Placa, 6, 1) < "0" Or Mid(Placa, 6, 1) > "9") Then ERRO = True
    If (Mid(Placa, 7, 1) < "0" Or Mid(Placa, 7, 1) > "9") And (Mid(Placa, 7, 1) <> " ") Then ERRO = True
    If (Mid(Placa, 8, 1) < "0" Or Mid(Placa, 8, 1) > "9") And (Mid(Placa, 8, 1) <> " ") Then ERRO = True
Else
    If Len(Placa) <> 8 Then ERRO = True
    If Not ERRO Then
        For i = 1 To 3
            If Mid(Placa, i, 1) < "A" Or Mid(Placa, i, 1) > "Z" Then ERRO = True
        Next
        For i = 5 To 8
            If Mid(Placa, i, 1) < "0" Or Mid(Placa, i, 1) > "9" Then ERRO = True
        Next
    End If
End If

If ERRO Then
    ValidaPlaca = False
Else
    ValidaPlaca = True
End If


End Function

Function Digito_DT(pNumero As Long) As String
Dim wNumero As String
wNumero = Format$(pNumero, "0000000") & "03"
Dim wSomaProduto As Long
Dim i As Integer
wSomaProduto = 0
For i = 2 To 9
    wSomaProduto = wSomaProduto + (i * Val(Mid(wNumero, i, 1)))
Next
wSomaProduto = wSomaProduto Mod 11
If wSomaProduto = 10 Then wSomaProduto = 0

Digito_DT = Format$(pNumero, "000000") & "-" & Format$(wSomaProduto, "0") & "/" & Format$(Year(Now), "0000")

End Function
Public Function VerificaBloqueios(BL As Long, Optional ByRef listaCNTR As String = Empty) As String
Dim oRS     As New ADODB.Recordset
Dim sOBS    As String

SQl = Empty


SQl = " SELECT 'BL '||TO_CHAR(A.AUTONUM) ||DECODE(B.FLAG_BLOQUEIA,1,' BLOQUEADO EM ',0,' LIBERADO EM ')||"
SQl = SQl & "TO_CHAR(B.DATA,'dd/mm/yyyy hh24:mi:ss')||' NA MODALIDADE '"
SQl = SQl & " ||B.TIPO||'   '||NVL(C.DESCR, B.MOTIVO) ||'  -  USUÁRIO '||B .USUARIO_SIS AS CAMPO2  , B.DATA"
SQl = SQl & " FROM       " & Banco_Sgipa & "TB_BL A"
SQl = SQl & " INNER JOIN " & Banco_Sgipa & "TB_HIST_BLOQUEIO B ON A.AUTONUM = B.BL"
SQl = SQl & " LEFT OUTER JOIN " & Banco_Sgipa & "TB_CAD_MOTIVO_BLOQUEIO C ON B.COD_MOTIVO_BLOQUEIO = C.AUTONUM"
SQl = SQl & " where A.AUTONUM =" & BL
SQl = SQl & " and (isnull(a.flag_brascenter,0)=0 or (a.flag_brascenter=1 and a.flag_bloqueio=1))"

SQl = SQl & "   UNION"

SQl = SQl & " SELECT    a.id_conteiner"
SQl = SQl & "        || ' lote '"
SQl = SQl & "        || TO_CHAR(d.bl)"
SQl = SQl & "        || DECODE (b.flag_bloqueia, 1, ' BLOQUEADO EM ', 0, 'LIBERADO EM')"
SQl = SQl & "        || TO_CHAR (b.DATA, 'dd/mm/yyyy hh24:mi:ss')"
SQl = SQl & "        || ' NA MODALIDADE '"
SQl = SQl & "        || b.tipo"
SQl = SQl & "        || '   '"
SQl = SQl & "        || NVL (c.descr, b.motivo)"
SQl = SQl & "        || '  -  USUÁRIO '"
SQl = SQl & "        || b.usuario_sis AS campo2,"
SQl = SQl & "        b.DATA"
SQl = SQl & "   FROM " & Banco_Sgipa & "tb_cntr_bl a INNER JOIN " & Banco_Sgipa & "tb_amr_cntr_bl d ON a.autonum = d.cntr "
SQl = SQl & "   INNER JOIN " & Banco_Sgipa & "tb_hist_bloqueio b ON d.cntr = b.cntr "
SQl = SQl & "   LEFT OUTER JOIN " & Banco_Sgipa & "tb_cad_motivo_bloqueio c ON b.cod_motivo_bloqueio = c.autonum"
SQl = SQl & "   inner join " & Banco_Sgipa & "tb_bl bl on d.bl=bl.autonum"
SQl = SQl & "  WHERE d.bl = " & BL
SQl = SQl & " and (isnull(bl.flag_brascenter,0)=0 or (bl.flag_brascenter=1 and a.flag_bloqueio=1))"

SQl = SQl & " group by a.id_conteiner,d.bl,b.flag_bloqueia,b.DATA,b.tipo,c.descr, b.motivo,b.usuario_sis,b.DATA"

SQl = SQl & " ORDER BY 2"

PRSet oRS, SQl, 0
sOBS = Empty
While Not oRS.EOF = True
    sOBS = sOBS & oRS.Fields("CAMPO2") & vbCrLf & vbCrLf
    oRS.MoveNext
Wend
oRS.Close
If Trim$(Replace$(sOBS, vbCrLf, Empty)) = Empty Then
    SQl = Empty
    SQl = SQl & "    SELECT autonum AS lote, '' AS cntr"
    SQl = SQl & " FROM " & Banco_Sgipa & "TB_BL WHERE autonum = '" & BL & "' And isnull(flag_bloqueio,0) + isnull(flag_bloqueio_manual,0) > 0"
    SQl = SQl & " and (isnull(flag_brascenter,0)=0)" ' or (flag_brascenter=1 and isnull(flag_bloqueio_manual,0)=1))
    SQl = SQl & "    UNION "
    SQl = SQl & "  SELECT 0 AS lote, id_conteiner AS cntr "
    SQl = SQl & " FROM " & Banco_Sgipa & "TB_CNTR_BL cc"
    SQl = SQl & " inner join " & Banco_Sgipa & "TB_AMR_CNTR_BL aa on cc.autonum = aa.cntr"
    SQl = SQl & " inner join " & Banco_Sgipa & "tb_bl bb on aa.bl = bb.autonum"
    SQl = SQl & " WHERE bl = '" & BL & "'"
    SQl = SQl & " AND cc.flag_bloqueio + cc.flag_bloqueio_manual > 0"
    SQl = SQl & " and (isnull(bb.flag_brascenter,0)=0)" ' or (isnull(bb.flag_brascenter,0)=1 and isnull(bb.flag_bloqueio_manual,0)=1))
    SQl = SQl & ""
    PRSet oRS, SQl
    While Not oRS.EOF = True
        sOBS = sOBS & oRS.Fields("LOTE") & " BLOQUEADO E SEM HISTÓRICO" & vbCrLf & vbCrLf
        sOBS = sOBS & oRS.Fields("CNTR") & " BLOQUEADO E SEM HISTÓRICO" & vbCrLf & vbCrLf
        oRS.MoveNext
    Wend
    oRS.Close
End If
listaCNTR = sOBS

SQl = Empty
SQl = " SELECT "
'Z3 Sql = Sql & "   DECODE(FLAG_BLOQUEIO,2,1,1,1,0,0) AS FLAG_BLOQUEIO,"
SQl = SQl & " CASE FLAG_BLOQUEIO WHEN 2 THEN 1 WHEN 1 THEN 1 WHEN 0 THEN 0 END As FLAG_BLOQUEIO,"

SQl = SQl & "   CAMPO2"
SQl = SQl & " FROM"
SQl = SQl & " ( "
SQl = SQl & "  ("
SQl = SQl & "    SELECT"
SQl = SQl & "       A.FLAG_BLOQUEIO + A.FLAG_BLOQUEIO_MANUAL AS FLAG_BLOQUEIO, "
SQl = SQl & "              'BL '||TO_CHAR(B.BL)||' BLOQUEADO EM '||TO_CHAR(B.DATA,'dd/mm/yyyy hh24:mi:ss')||' NA MODALIDADE '||B.TIPO||'   '||NVL(C.DESCR, B.MOTIVO)||'  -  USUÁRIO '||NVL(D.NOME_AGENTE||' ' ||D.CPF_AGENTE,' ')||B .USUARIO_SIS AS CAMPO2  "
SQl = SQl & "     FROM"
SQl = SQl & "       " & Banco_Sgipa & "TB_BL A"
SQl = SQl & "         INNER JOIN (SELECT * FROM " & Banco_Sgipa & "TB_HIST_BLOQUEIO WHERE FLAG_BLOQUEIA = 1) B ON"
SQl = SQl & "         A.AUTONUM = B.BL "
SQl = SQl & "         AND A.AUTONUM = '" & BL & "' "
SQl = SQl & "           LEFT OUTER JOIN " & Banco_Sgipa & "TB_CAD_MOTIVO_BLOQUEIO C ON"
SQl = SQl & "           B.COD_MOTIVO_BLOQUEIO = C.AUTONUM"
SQl = SQl & "           LEFT OUTER JOIN  " & Banco_Sgipa & "INTERNET_BLOQUEIO_NVO D ON"
SQl = SQl & "           A.AUTONUM= D.AUTONUM_BL "
SQl = SQl & "     WHERE(A.FLAG_BLOQUEIO + A.FLAG_BLOQUEIO_MANUAL) > 0"

SQl = SQl & " and (isnull(a.flag_brascenter,0)=0 or (a.flag_brascenter=1 and (a.flag_bloqueio+a.flag_bloqueio_manual)>0))"

SQl = SQl & "   UNION "

SQl = SQl & "     SELECT"
SQl = SQl & "       A.FLAG_BLOQUEIO + A.FLAG_BLOQUEIO_MANUAL AS FLAG_BLOQUEIO , "
SQl = SQl & "       'CONTEINER '||A.ID_CONTEINER||' BLOQUEADO EM '||TO_CHAR(B.DATA)||' NA MODALIDADE '||B.TIPO||' POR '||NVL(C.DESCR, B.MOTIVO)||' PELO USUÁRIO '||B.USUARIO_SIS AS CAMPO2  "
SQl = SQl & "     FROM "
SQl = SQl & "       " & Banco_Sgipa & "TB_CNTR_BL A"
SQl = SQl & "       INNER JOIN (SELECT * FROM " & Banco_Sgipa & "TB_HIST_BLOQUEIO WHERE FLAG_BLOQUEIA = 1) B ON"
SQl = SQl & "       A.AUTONUM = B.CNTR "
SQl = SQl & "         LEFT OUTER JOIN " & Banco_Sgipa & "TB_CAD_MOTIVO_BLOQUEIO C ON"
SQl = SQl & "         B.COD_MOTIVO_BLOQUEIO = C.AUTONUM"

SQl = SQl & " inner join " & Banco_Sgipa & "tb_amr_cntr_bl aa on a.autonum = aa.cntr"
SQl = SQl & " inner join " & Banco_Sgipa & "tb_bl bb on aa.bl=bb.autonum"

SQl = SQl & "     WHERE "
SQl = SQl & "       A.AUTONUM IN(SELECT CNTR FROM " & Banco_Sgipa & "TB_AMR_CNTR_BL WHERE BL = '" & BL & "') "
'Sql = Sql & "       AND (A.FLAG_BLOQUEIO + A.FLAG_BLOQUEIO_MANUAL) > 0"

SQl = SQl & " AND a.flag_bloqueio + a.flag_bloqueio_manual > 0"
SQl = SQl & " and (isnull(bb.flag_brascenter,0)=0 or (bb.flag_brascenter=1 and (bb.flag_bloqueio+bb.flag_bloqueio_manual>0)))"

SQl = SQl & "   )"
SQl = SQl & " ) "

SQl = SQl & "  ALIAS1 "

SQl = SQl & " WHERE "
SQl = SQl & "   FLAG_BLOQUEIO > 0"

SQl = SQl & " ORDER BY 2"
pulo:
SQl = CONVERTE_SQL(SQl)

Dim rstemp As New ADODB.Recordset
'If RSTEMP.State = adStateOpen Then RSTEMP.Close
'MsgBox Sql
'MsgBox Right(Sql, 700)
'RSTEMP.Open Sql, Db

PRSet oRS, SQl, 0
sOBS = Empty


If Not oRS.RecordCount = 0 Then VerificaBloqueios = "BLOQUEIO EXISTENTE"

Set oRS = Nothing
sOBS = ""

End Function




Public Function QualMaquina(Tamanho As String)
Dim sRETORNO As String * 255
Dim nChar    As Integer
Dim lRET     As Boolean
If Tamanho > 255 Then Err.Raise 80, , "Parâmetro inválido (Tamanho é superior á 255 caracteres)."
nChar = 255
lRET = GetComputerName(sRETORNO, (nChar))
QualMaquina = Left$(Replace(Replace(sRETORNO, Chr(1), Empty), Chr(0), Empty), Tamanho)
lRET = Empty
nChar = Empty
sRETORNO = Empty
End Function

Public Function QualUsuario(Tamanho As String)
Dim sRETORNO As String * 255
Dim nChar    As Integer
Dim lRET     As Boolean
If Tamanho > 255 Then Err.Raise 80, , "Parâmetro inválido (Tamanho é superior á 255 caracteres)."
nChar = 255
lRET = GetUserName(sRETORNO, (nChar))
QualUsuario = Left$(Replace(Replace(sRETORNO, Chr(1), Empty), Chr(0), Empty), Tamanho)
lRET = Empty
nChar = Empty
sRETORNO = Empty
End Function

Public Function ValidaDsi(Valor As String)

    Dim Multiplos(11) As Integer
    Dim SomaProdutos As Long
    Dim DigitoLido As Integer
    Dim Digito As Integer
    Dim i As Integer
    Multiplos(1) = 3
    Multiplos(2) = 2
    Multiplos(3) = 9
    Multiplos(4) = 8
    Multiplos(5) = 7
    Multiplos(6) = 6
    Multiplos(7) = 5
    Multiplos(8) = 4
    Multiplos(9) = 3
    Multiplos(10) = 2
    
    DigitoLido = Val(Right$(Valor, 1))
    Valor = "4" & Mid$(Valor, 1, 2) & Mid$(Valor, 4, 7)
    
    For i = 10 To 1 Step -1
        If Mid$(Valor, i, 1) < "0" Or Mid$(Valor, i, 1) > "9" Then
            ValidaDsi = False
            Exit Function
        End If
        SomaProdutos = SomaProdutos + Multiplos(i) * Val(Mid$(Valor, i, 1))
    Next i
    
    'SomaProdutos = SomaProdutos * 10
    'Verifica se digito bate
    Digito = SomaProdutos Mod 11
    Digito = 11 - Digito
    
    If Digito > 9 Then Digito = 0
    
    If Digito = DigitoLido Then
        ValidaDsi = True
    Else
        ValidaDsi = False
    End If

End Function


Function Tira_AspaSimples(campo As String) As String
Tira_AspaSimples = Replace(campo, "'", Empty)
End Function
Function Tira_Espaco(campo As String)
If campo = "" Then Exit Function
Dim i As Integer
For i = 1 To Len(campo)
    If Mid$(campo, i, 1) <> " " Then
        Tira_Espaco = Tira_Espaco + Mid$(campo, i, 1)
    End If
Next i
End Function

Function Tira_Ponto(campo As String)
If campo = "" Then Exit Function
Dim i As Integer
For i = 1 To Len(campo)
    If Mid$(campo, i, 1) <> "." Then
        Tira_Ponto = Tira_Ponto + Mid$(campo, i, 1)
    End If
Next i
End Function

Function Calcula_Digito_Conteiner(Conteiner As String)
ReDim Alpha(1 To 26) As String
ReDim Valores_Alpha(1 To 26) As Integer
ReDim Parcelas(1 To 10) As String
Dim Somatorio As Long
Dim CleanConteiner As String, i As Integer, j As Integer, Ok As Boolean

'Inicializacao Variávies
CleanConteiner = ""
Somatorio = 0
Ok = True
Conteiner = UCase(Conteiner)
'Valores dos Caracteres Alpha
Alpha(1) = "A": Alpha(2) = "B": Alpha(3) = "C": Alpha(4) = "D": Alpha(5) = "E": Alpha(6) = "F": Alpha(7) = "G"
Alpha(8) = "H": Alpha(9) = "I": Alpha(10) = "J": Alpha(11) = "K": Alpha(12) = "L": Alpha(13) = "M"
Alpha(14) = "N": Alpha(15) = "O": Alpha(16) = "P": Alpha(17) = "Q": Alpha(18) = "R": Alpha(19) = "S"
Alpha(20) = "T": Alpha(21) = "U": Alpha(22) = "V": Alpha(23) = "W": Alpha(24) = "X": Alpha(25) = "Y": Alpha(26) = "Z"
Valores_Alpha(1) = 10: Valores_Alpha(2) = 12: Valores_Alpha(3) = 13: Valores_Alpha(4) = 14
Valores_Alpha(5) = 15: Valores_Alpha(6) = 16: Valores_Alpha(7) = 17: Valores_Alpha(8) = 18
Valores_Alpha(9) = 19: Valores_Alpha(10) = 20: Valores_Alpha(11) = 21: Valores_Alpha(12) = 23
Valores_Alpha(13) = 24: Valores_Alpha(14) = 25: Valores_Alpha(15) = 26: Valores_Alpha(16) = 27
Valores_Alpha(17) = 28: Valores_Alpha(18) = 29: Valores_Alpha(19) = 30: Valores_Alpha(20) = 31
Valores_Alpha(21) = 32: Valores_Alpha(22) = 34: Valores_Alpha(23) = 35: Valores_Alpha(24) = 36
Valores_Alpha(25) = 37: Valores_Alpha(26) = 38
'Elimina Espaços e Caracteres Indesejáveis
For i = 1 To Len(Conteiner)
    If Trim(Mid$(Conteiner, i, 1)) <> "" Then
        If (Asc(Mid$(Conteiner, i, 1)) >= 48 And Asc(Mid$(Conteiner, i, 1)) <= 57) Or (Asc(Mid$(Conteiner, i, 1)) >= 65 And Asc(Mid$(Conteiner, i, 1)) <= 90) Then
            CleanConteiner = CleanConteiner & Mid$(Conteiner, i, 1)
        End If
    End If
Next i
If CleanConteiner = "" Then Exit Function
'Consistencia da Digitacao
On Error Resume Next
For i = 1 To 4

    If (Asc(Mid$(CleanConteiner, i, 1)) < 65 Or Asc(Mid$(CleanConteiner, i, 1)) > 90) Then
        Ok = False
    End If
Next i
For i = 5 To 10
    If (Asc(Mid$(CleanConteiner, i, 1)) < 48 Or Asc(Mid$(CleanConteiner, i, 1)) > 57) Then
        Ok = False
    End If
Next i
If Len(CleanConteiner) <> 10 And Len(CleanConteiner) <> 11 Then
    Ok = False
End If

'If Not Ok Then MsgBox "Formato incorreto da Identificacao do Conteiner "
'Calculo
For i = 1 To 10
    'Quatro Primeiros Digitos são Alpha
    If i < 5 Then
        For j = 1 To 26
            If Alpha(j) = Mid$(CleanConteiner, i, 1) Then
                Exit For
            End If
        Next j
        Somatorio = Somatorio + (Valores_Alpha(j) * 2 ^ (i - 1))
    Else
        Somatorio = Somatorio + (Val(Mid$(CleanConteiner, i, 1)) * 2 ^ (i - 1))
    End If
Next i
Calcula_Digito_Conteiner = Right$(CStr(Somatorio Mod 11), 1)
On Error GoTo 0
End Function
Function So09Letras(Palavra As String)
Dim i As Integer
For i = 1 To Len(Palavra)
    If Trim(Mid$(Palavra, i, 1)) <> "" Then
        If (Asc(Mid$(Palavra, i, 1)) >= 48 And Asc(Mid$(Palavra, i, 1)) <= 57) Or (Asc(Mid$(Palavra, i, 1)) >= 65 And Asc(Mid$(Palavra, i, 1)) <= 90) Then
           So09Letras = So09Letras & Mid$(Palavra, i, 1)
        End If
    Else
           So09Letras = So09Letras & " "
    End If
Next i
End Function
Function Formata_Placa(Placa As String)
Dim Entrada As String
Dim i As Integer

For i = 1 To Len(Placa)
    If Mid$(Placa, i, 1) <> "-" Then
        Entrada = Entrada & Mid$(Placa, i, 1)
    End If
Next i
Select Case Len(Trim(Entrada))
    Case 7
        Formata_Placa = UCase(Mid$(Entrada, 1, 3) & "-" & Right(Entrada, 4))
    Case 6
        Formata_Placa = UCase(Mid$(Entrada, 1, 2) & "-" & Right(Entrada, 4))
    Case 0
    Case Else
        MsgBox "Placa em formato Inválido", vbCritical
End Select
End Function
Function FormatAN(campo As String) As String
Dim X As Integer
For X = 1 To Len(campo)
   Select Case UCase(Mid$(campo, X, 1))
   Case "A" To "Z", "0" To "9", ".", "/", "-"
   Case "À" To "Ã"
   Mid$(campo, X, 1) = "A"
   Case "È" To "Ê"
   Mid$(campo, X, 1) = "E"
   Case "Ì" To "Î"
   Mid$(campo, X, 1) = "I"
   Case "Ò" To "Õ"
   Mid$(campo, X, 1) = "O"
   Case "Ù" To "Û"
   Mid$(campo, X, 1) = "U"
   Case "Ç"
   Mid$(campo, X, 1) = "C"
   Case Else
   Mid$(campo, X, 1) = " "
   End Select
Next X
FormatAN = UCase(campo)
End Function
Function FormatA(campo As String) As String
Dim X As Integer

For X = 1 To Len(campo)
   Select Case UCase(Mid$(campo, X, 1))
   Case "A" To "Z"
   Case "À" To "Ã"
   Mid$(campo, X, 1) = "A"
   Case "È" To "Ê"
   Mid$(campo, X, 1) = "E"
   Case "Ì" To "Î"
   Mid$(campo, X, 1) = "I"
   Case "Ò" To "Õ"
   Mid$(campo, X, 1) = "O"
   Case "Ù" To "Û"
   Mid$(campo, X, 1) = "U"
   Case "Ç"
   Mid$(campo, X, 1) = "C"
   Case Else
   Mid$(campo, X, 1) = " "
   End Select
Next X
FormatA = UCase(campo)
End Function


Function FormatN(campo As String, Tam As Integer, Carac As String) As String
Dim X As Integer
For X = 1 To Len(campo)
   If UCase(Mid$(campo, X, 1)) < "0" Or UCase(Mid$(campo, X, 1)) > "9" Then
      Mid$(campo, X, 1) = " "
   End If
Next X
campo = Val(campo)
If Tam > Len(campo) Then
   FormatN = String(Tam - Len(campo), Carac) & campo
Else
   FormatN = campo
End If
End Function
Function Limpa_(Palavra As Variant)
Dim i As Integer, Letr As String
Dim Aux As String
Aux = ""
For i = 1 To Len(Palavra)
    Letr = Mid$(Palavra, i, 1)
    If Asc(Letr) <> 95 Then Aux = Aux & Letr
Next i
Limpa_ = Aux
End Function
Function Limpa32(Palavra As Variant)
Dim i As Integer, Letr As String
Dim Aux As String
Aux = ""
For i = 1 To Len(Palavra)
    Letr = Mid$(Palavra, i, 1)
    If Asc(Letr) <> 0 Then Aux = Aux & Letr
Next i
Limpa32 = Aux
End Function
Function LimpaEspacosConsecutivos(campo) As String
Dim Aux As String
Dim i As Integer
    For i = 1 To Len(campo)
        If Mid$(campo, i, 1) = " " Then
            If i > 1 And Mid$(campo, i - 1, 1) = " " Then
                '2 Brancos Consecutivos
            Else
                Aux = Aux & Mid$(campo, i, 1)
            End If
        Else
            Aux = Aux & Mid$(campo, i, 1)
        End If
    Next i
Aux = Left$(Aux + Space(100), Len(campo))
LimpaEspacosConsecutivos = Aux
End Function
Public Function Consiste_Primeira_Entrada(Lote As Long) As Boolean
Dim lote_entrada As String
Dim data_anvisa As Date
Dim Data_Entrada As String
Dim tb   As New ADODB.Recordset
Dim tbC  As New ADODB.Recordset
Dim tbcs As New ADODB.Recordset
PRSet tb, "SELECT NVL(ID_SGTAP,0) AS MBM FROM " & Banco_Sgipa & "TB_BL WHERE AUTONUM = " & Lote & " ", 0
If tb.EOF Then
    Data_Entrada = ""
    'Data_Entrada = "27/07/2004 08:00:00"
    Executa ("UPDATE " & Banco_Sgipa & "TB_BL SET PRIMEIRA_ENTRADA = NULL WHERE (AUTONUM = " & Lote & ") AND (ID_SGTAP IS NULL OR ID_SGTAP = 0)")
    Set tbC = New ADODB.Recordset
    tbC.Open "SELECT MIN(a.DT_ENTRADA) AS DT_ENTRADA FROM " & Banco_Sgipa & "TB_CNTR_BL A," & Banco_Sgipa & "TB_AMR_CNTR_BL B WHERE B.CNTR=A.AUTONUM AND B.BL=" & Lote & " and a.dt_entrada is not null", db, 3, 3
    If tbC.EOF = False Then
        If Not IsNull(tbC!Dt_Entrada) Then
            Data_Entrada = CVDate(tbC!Dt_Entrada)
        End If
    End If
    tbC.Close
    '--------
    Set tbcs = New ADODB.Recordset
    tbcs.Open "select min(a.dt_ENTRADA) as dt_entrada from " & Banco_Sgipa & "TB_CARGA_GATE_CS b, " & Banco_Sgipa & "TB_CARGA_SOLTA c where c.autonum=b.autonumcs AND C.DT_ENTRADA IS NOT NULL  and c.bl= " & Lote, db, 3, 3
    If Not tbcs.EOF Then
        If Not IsNull(tbcs!Dt_Entrada) Then
            If Data_Entrada = "" Then
                Data_Entrada = CVDate(tbcs!Dt_Entrada)
            Else
                 If CVDate(tbcs!Dt_Entrada) < CVDate(Data_Entrada) Then
                    Data_Entrada = CVDate(tbcs!Dt_Entrada)
                 End If
            End If
        End If
    End If
    tbcs.Close
    If Not IsNull(Data_Entrada) Then
        If IsDate(Data_Entrada) Then Executa "UPDATE " & Banco_Sgipa & "TB_BL SET PRIMEIRA_ENTRADA=TO_DATE('" & Data_Entrada & "','DD/MM/YYYY HH24:MI:SS') WHERE autonum=" & Lote
        data_anvisa = Data_Entrada
'        Call HISTORICO_ANVISA(lote_entrada, data_anvisa)
'        Call HISTORICO_IMO(lote_entrada, data_anvisa)
    
    End If
    Consiste_Primeira_Entrada = Not (Trim(Data_Entrada) = Empty)
Else
    Consiste_Primeira_Entrada = True
End If
Data_Entrada = Empty
Set tb = Nothing
Set tbC = Nothing
Set tbcs = Nothing
End Function
Sub PintaGrid(ByRef Grid As Object, Row As Integer, Tipo As String)
Dim k As Integer
Dim Cor As Variant
Grid.Row = Row
Select Case Tipo
    Case "VERDE"
        For k = 1 To Grid.Cols - 1
            Grid.Col = k
            Grid.CellBackColor = &HC0FFC0
            '12648384
            Grid.CellForeColor = vbBlack
        Next k
    Case "AMARELO"
        For k = 1 To Grid.Cols - 1
            Grid.Col = k
            Grid.CellBackColor = &HC0FFFF
            '12648447
            Grid.CellForeColor = vbBlack
        Next k
    Case "VERMELHO"
        For k = 1 To Grid.Cols - 1
            Grid.Col = k
            Grid.CellBackColor = &H8080FF
            '8421631
            Grid.CellForeColor = vbWhite
        Next k
    Case "AZUL"
        For k = 1 To Grid.Cols - 1
            Grid.Col = k
            Grid.CellBackColor = &HFFC0C0
            '16711680
            Grid.CellForeColor = vbWhite
        Next k
    Case "BRANCO"
        For k = 1 To Grid.Cols - 1
            Grid.Col = k
            Grid.CellBackColor = vbWhite
            '16711680
            Grid.CellForeColor = vbBlack
        Next k
End Select
End Sub





'
Public Sub GridLineColor(ByRef Grid As Object, Row As Integer, Optional BackColor As OLE_COLOR = vbWhite, Optional ForeColor As OLE_COLOR = vbBlack)
Dim t As Long
Grid.Row = Row
For t = 1 To Grid.Cols - 1
    Grid.Col = t
    Grid.CellBackColor = BackColor
    Grid.CellForeColor = ForeColor
Next t
t = Empty
End Sub
Function Hora_Nula_Mask(Par As Variant)
If Not IsDate(Par) Then
    Hora_Nula_Mask = Null
Else
    Hora_Nula_Mask = Par
End If
End Function
Function Data_Nula_Mask(Par As Variant)
If Not IsDate(Par) Then
    Data_Nula_Mask = Null
Else
    Data_Nula_Mask = Par
End If
End Function

Function s(Tam As Integer)
    s = Left$("                                                                            ", Tam)
End Function
Function Formata_Numero(Valor As Double)
   Formata_Numero = Format(Valor, "###,###,###,##0.00")
End Function
Function ValidaCPF(WCgc As String)
Static NumA(1 To 11) As Double
Static NumB(1 To 11) As Double
Dim SomaB As Double
Dim Trabalho As String
Dim Num As Double
Dim i As Integer
  For i = 1 To 11
   NumA(i) = 0
   NumB(i) = 0
  Next i
  For i = 1 To 9
   NumA(i) = Val(Mid$(WCgc, i, 1))
  Next i
  NumB(1) = NumA(1) * 10
  NumB(2) = NumA(2) * 9
  NumB(3) = NumA(3) * 8
  NumB(4) = NumA(4) * 7
  NumB(5) = NumA(5) * 6
  NumB(6) = NumA(6) * 5
  NumB(7) = NumA(7) * 4
  NumB(8) = NumA(8) * 3
  NumB(9) = NumA(9) * 2
  SomaB = 0
  For i = 1 To 9
   SomaB = SomaB + NumB(i)
  Next i
  Num = Fix(SomaB * 10 / 11) * 11
  Num = (SomaB * 10) - Num
  If Num > 9 Then Num = 0
  NumA(10) = Num
  NumB(1) = NumA(1) * 11
  NumB(2) = NumA(2) * 10
  NumB(3) = NumA(3) * 9
  NumB(4) = NumA(4) * 8
  NumB(5) = NumA(5) * 7
  NumB(6) = NumA(6) * 6
  NumB(7) = NumA(7) * 5
  NumB(8) = NumA(8) * 4
  NumB(9) = NumA(9) * 3
  NumB(10) = NumA(10) * 2
  SomaB = 0
  For i = 1 To 10
   SomaB = SomaB + NumB(i)
  Next i
  Num = Fix(SomaB * 10 / 11) * 11
  Num = (SomaB * 10) - Num
  If Num > 9 Then Num = 0
  NumA(11) = Num
  Digito1 = Mid$(Trim(Str(Fix(NumA(10)))), 1, 1)
  Digito2 = Mid$(Trim(Str(Fix(NumA(11)))), 1, 1)
  ValidaCPF = (Digito1 = Mid$(WCgc, 10, 1)) And (Digito2 = Mid$(WCgc, 11, 1))
End Function
Function FormatarCgcCpf(txtcgc) As String
Dim wPri As String
Dim wSeg As String
Dim wTri As String
Dim wQua As String
Dim wMuda As String
Dim wGuardaNum As String

If txtcgc = "" Then
   FormatarCgcCpf = ""
   Exit Function
End If
 If Len(txtcgc) > 11 Then
    wGuardaNum = txtcgc
    wMuda = wGuardaNum
    wPri = Left(wMuda, 2) & "."
    wMuda = wGuardaNum
    wSeg = wPri & Mid(wMuda, 3, 3) & "."
    wMuda = wGuardaNum
    wTri = wSeg & Mid(wMuda, 6, 3) & "/"
    wMuda = wGuardaNum
    wQua = wTri & Mid(wMuda, 9, 4) & "-"
    wMuda = wGuardaNum
    wMuda = wQua & Mid(wMuda, 13, 2)
    FormatarCgcCpf = wMuda
Else
    wGuardaNum = txtcgc
    wMuda = wGuardaNum
    wPri = Left(wMuda, 3) & "."
    wMuda = wGuardaNum
    wSeg = wPri & Mid(wMuda, 4, 3) & "."
    wMuda = wGuardaNum
    wTri = wSeg & Mid(wMuda, 7, 3) & "-"
    wMuda = wGuardaNum
    wMuda = wTri & Mid(wMuda, 10, 2)
    FormatarCgcCpf = wMuda
    End If
End Function
Function ValidaDI(Valor As String)

    Dim Multiplos(11) As Integer
    Dim SomaProdutos As Long
    Dim DigitoLido As Integer
    Dim Digito As Integer
    Dim i As Integer
    Multiplos(1) = 3
    Multiplos(2) = 2
    Multiplos(3) = 9
    Multiplos(4) = 8
    Multiplos(5) = 7
    Multiplos(6) = 6
    Multiplos(7) = 5
    Multiplos(8) = 4
    Multiplos(9) = 3
    Multiplos(10) = 2
    
    DigitoLido = Val(Right$(Valor, 1))
    Valor = "2" & Mid$(Valor, 3, 2) & Mid$(Valor, 6, 7)
    
    For i = 10 To 1 Step -1
        If Mid$(Valor, i, 1) < "0" Or Mid$(Valor, i, 1) > "9" Then
            ValidaDI = False
            Exit Function
        End If
        SomaProdutos = SomaProdutos + Multiplos(i) * Val(Mid$(Valor, i, 1))
    Next i
    
    'SomaProdutos = SomaProdutos * 10
    'Verifica se digito bate
    Digito = SomaProdutos Mod 11
    Digito = 11 - Digito
    
    If Digito > 9 Then Digito = 0
    
    If Digito = DigitoLido Then
        ValidaDI = True
    Else
        ValidaDI = False
    End If

End Function

Function ValidaDTA(Valor As String)
    
    Dim Multiplos(11) As Integer
    Dim SomaProdutos As Long
    Dim DigitoLido As Integer
    Dim Digito As Integer
    Dim i As Integer
    
    Multiplos(1) = 3
    Multiplos(2) = 2
    Multiplos(3) = 9
    Multiplos(4) = 8
    Multiplos(5) = 7
    Multiplos(6) = 6
    Multiplos(7) = 5
    Multiplos(8) = 4
    Multiplos(9) = 3
    Multiplos(10) = 2
    
    DigitoLido = Val(Mid(Valor, 8, 1))
    Valor = "0" & Mid$(Valor, 12, 2) & Mid$(Valor, 1, 7)
    
    For i = 10 To 1 Step -1
        If Mid$(Valor, i, 1) < "0" Or Mid$(Valor, i, 1) > "9" Then
            ValidaDTA = False
            Exit Function
        End If
        SomaProdutos = SomaProdutos + Multiplos(i) * Val(Mid$(Valor, i, 1))
    Next i
    
    'SomaProdutos = SomaProdutos * 10
    'Verifica se digito bate
    Digito = SomaProdutos Mod 11
    Digito = 11 - Digito
    
    If Digito > 9 Then Digito = 0
    
    If Digito = DigitoLido Then
        ValidaDTA = True
    Else
        ValidaDTA = False
    End If
    
End Function


Function ValidaCgc(WCgc As String)
Static NumA(1 To 14) As Double
Static NumB(1 To 14) As Double
Dim SomaB As Double
Dim Trabalho As String
Dim Num As Double
Dim i As Integer

For i = 1 To 14
 NumA(i) = 0
 NumB(i) = 0
Next i
For i = 1 To 12
 NumA(i) = Val(Mid$(WCgc, i, 1))
Next i
NumB(1) = NumA(1) * 5
NumB(2) = NumA(2) * 4
NumB(3) = NumA(3) * 3
NumB(4) = NumA(4) * 2
NumB(5) = NumA(5) * 9
NumB(6) = NumA(6) * 8
NumB(7) = NumA(7) * 7
NumB(8) = NumA(8) * 6
NumB(9) = NumA(9) * 5
NumB(10) = NumA(10) * 4
NumB(11) = NumA(11) * 3
NumB(12) = NumA(12) * 2
SomaB = 0
For i = 1 To 12
 SomaB = SomaB + NumB(i)
Next i
Num = Fix(SomaB * 10 / 11) * 11
Num = (SomaB * 10) - Num
If Num > 9 Then Num = 0
NumA(13) = Num
NumB(1) = NumA(1) * 6
NumB(2) = NumA(2) * 5
NumB(3) = NumA(3) * 4
NumB(4) = NumA(4) * 3
NumB(5) = NumA(5) * 2
NumB(6) = NumA(6) * 9
NumB(7) = NumA(7) * 8
NumB(8) = NumA(8) * 7
NumB(9) = NumA(9) * 6
NumB(10) = NumA(10) * 5
NumB(11) = NumA(11) * 4
NumB(12) = NumA(12) * 3
NumB(13) = NumA(13) * 2
SomaB = 0
For i = 1 To 13
 SomaB = SomaB + NumB(i)
Next i
Num = Fix(SomaB * 10 / 11) * 11
Num = (SomaB * 10) - Num
If Num > 9 Then Num = 0
NumA(14) = Num
Digito1 = Mid$(Trim(Str(Fix(NumA(13)))), 1, 1)
Digito2 = Mid$(Trim(Str(Fix(NumA(14)))), 1, 1)
ValidaCgc = (Digito1 = Mid$(WCgc, 13, 1)) And (Digito2 = Mid$(WCgc, 14, 1) And SomaB <> 0)
End Function

Function Marca(Tela As Form)
On Error GoTo ErroMarca
    Tela.ActiveControl.SelStart = 0
    Tela.ActiveControl.SelLength = Len(Tela.ActiveControl)
Exit Function
ErroMarca:
If Err = 438 Then
  Resume Next
End If
End Function

Public Function SoNumeros(ks As Integer) As String
Dim X As Integer
Dim N As String
Dim s As String
Dim a As String

If ks = 8 Or ks = 9 Then SoNumeros = ks: Exit Function

a = Chr$(ks)
If Banco_Em_Uso = "ORACLE" Then
    s = QualSeparadorDecimalOracle
Else
    s = ","
End If
N = IIf(s = ".", ",", ".")
If Not InStr(1, "0123456789-" & s, a) > 0 Then
    If a = N Then
        ks = Asc(s)
        SoNumeros = ks
    Else
        ks = Empty
        SoNumeros = 0
    End If
Else
    SoNumeros = ks
End If
a = Empty
N = Empty
s = Empty
End Function
Public Function SoData(ks As Integer) As Integer
   If ks >= 65 Or Chr(ks) = " " Then
      MsgBox "Este campo só ira aceitar Numeros ou datas!", , "Atenção"
      SoData = 0
   Else
        If Chr(ks) = "." Or Chr(ks) = "," Or Chr(ks) = ";" Or Chr(ks) = "-" Then
            SoData = Asc("/")
        Else
            SoData = ks
        End If

   End If

End Function
Public Function SoLetras(ks As Integer) As String
If UCase(Chr$(ks)) >= "A" And UCase(Chr$(ks)) <= "Z" Then
   SoLetras = ks
Else
   SoLetras = 0
End If
End Function

Public Function SoNumeros09(ks As Integer) As String
If (ks >= 58 Or ks <= 46) And ks <> 8 And ks <> 13 Then
   'X = MsgBox("Este campo não aceita este caracter!", 48, "Atenção")
   SoNumeros09 = 0
Else
   SoNumeros09 = ks
End If
End Function

Public Function SoNumerosVirgulaDecimal(ks As Integer) As String
If (ks >= 58 Or ks <= 46) And ks <> 8 And ks <> 13 Then
   'X = MsgBox("Este campo não aceita este caracter!", 48, "Atenção")
   SoNumerosVirgulaDecimal = 0
Else
   SoNumerosVirgulaDecimal = ks
End If
End Function
Public Function QuantosCaracteres(TEXTO As String, Caractere As String) As Integer
If Len(TEXTO) = 0 Or Len(Caractere) = 0 Then Exit Function
Dim t As Long
Dim N As Integer
For t = 1 To Len(TEXTO)
    N = N + IIf(Mid$(TEXTO, t, 1) = Caractere, 1, 0)
Next
QuantosCaracteres = N
t = Empty
N = Empty
End Function
Public Function QualSeparadorDecimalOracle() As String
On Error GoTo TErr:
Dim oRS As New ADODB.Recordset
PRSet oRS, "SELECT TO_CHAR(10/3,'999G999D99') AS VALOR FROM DUAL", 0
If Not oRS.RecordCount = 0 Then
    QualSeparadorDecimalOracle = Mid$(Trim$(oRS.Fields("VALOR")), 2, 1)
End If
oRS.Close
GoTo TExit:
TErr:
Err.Clear
'IF BANCO_EM_USO = "ORACLE" Then
    PRSet oRS, "SELECT TO_NUMBER('123456,9853') AS VALOR FROM DUAL", 0
        If Not oRS.EOF Then
            QualSeparadorDecimalOracle = Mid$(oRS.Fields("VALOR"), 7, 1)
        End If
    oRS.Close
'End If

TExit:
Set oRS = Nothing
End Function
Public Function TransfData(strData As String) As String
Dim Dia As String
Dim mes As String
Dim Ano As String
Dim dataFormt As String
Dim i As Integer

dataFormt = ""
For i = 1 To Len(strData)
    If IsNumeric(Mid(strData, i, 1)) Then
        dataFormt = dataFormt + Mid(strData, i, 1)
    End If
Next

Select Case Len(dataFormt)
    Case 2
        Dia = Val(Format(dataFormt, "00"))
        mes = Val(Format(Month(Date), "00"))
        Ano = Val(Format(Year(Date), "0000"))
    Case 4
        Dia = Val(Format(Left(dataFormt, 2), "00"))
        mes = Val(Format(Right(dataFormt, 2), "00"))
        Ano = Val(Format(Year(Date), "0000"))
    Case 6
        Dia = Val(Format(Left(dataFormt, 2), "00"))
        mes = Val(Format(Mid(dataFormt, 3, 2), "00"))
        If Format(Right(dataFormt, 2), "00") < "80" Then
            Ano = Val(Format("20" + Right(dataFormt, 2), "00"))
        Else
            Ano = Val(Format("19" + Right(dataFormt, 2), "00"))
        End If
    Case 8
        Dia = Val(Format(Left(dataFormt, 2), "00"))
        mes = Val(Format(Mid(dataFormt, 3, 2), "00"))
        Ano = Val(Format(Right(dataFormt, 4), "0000"))
    Case Else
        TransfData = ""
        Exit Function
End Select
If IsDate(DateSerial(Ano, mes, Dia)) Then
    TransfData = Format(DateSerial(Ano, mes, Dia), "dd/mm/yyyy")
Else
    TransfData = ""
End If
End Function
Function Form(VL As Variant, ant As Integer, dep As Integer)
Dim X As String, fo As String
If VL <> 0 Then
    fo = Right$("#################0", ant)
    dep = dep + 1
    If dep > 1 Then fo = fo + Left$(".0000000000000000000", dep)
    X = Format$(VL, fo)
    X = Left$(Space(30), ant + dep - Len(X)) + X
Else
    X = Space(ant + dep + 1)
End If
    Form = X
End Function
Function formt(tx As String, Tam As Integer)
Dim X As String
    If IsNull(tx) Then
        X = " "
    Else
        X = tx
    End If
    X = Format$(X, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
    formt = Left$(X, Tam) & String(Tam - Len(Left$(X, Tam)), Space(1))
End Function
Function criptografa(SENHA As String)
Dim i As Integer

criptografa = ""
For i = 1 To Len(SENHA)
    criptografa = criptografa & Chr$(Asc(UCase(Mid$(SENHA, i, 1))) - 1)
Next i
End Function



Function Descriptografa(SENHA As String)
Dim i As Integer
Descriptografa = ""
For i = 1 To Len(SENHA)
    Descriptografa = Descriptografa & Chr$(Asc(UCase(Mid$(SENHA, i, 1))) + 1)
Next i
End Function
Function TransformaVirgulaEmPonto(CodigoAscii As Integer)
If CodigoAscii = 44 Then
    TransformaVirgulaEmPonto = 46
Else
    TransformaVirgulaEmPonto = CodigoAscii
End If
End Function

Function LimpaForm(Tela As Form)
Dim i As Integer
For i = 1 To Tela.Controls.Count - 1
    On Error Resume Next
    If UCase(Mid$(Tela.Controls(i).Name, 1, 5)) = "COMBO" Then
        Tela.Controls(i).ListIndex = -1
    End If
    On Error GoTo 0
    If UCase(Mid$(Tela.Controls(i).Name, 1, 7)) = "DBCOMBO" Then
        Tela.Controls(i).Text = ""
    End If
    
    If UCase(Mid$(Tela.Controls(i).Name, 1, 9)) = "DATACOMBO" Then
        Tela.Controls(i).Text = ""
    End If
    
    If UCase(Mid$(Tela.Controls(i).Name, 1, 9)) = "TXTFIELDS" Then
        Tela.Controls(i).Text = ""
    End If
    
    If UCase(Mid$(Tela.Controls(i).Name, 1, 5)) = "CHECK" Then
        Tela.Controls(i).Value = 0
    End If
Next i



End Function


Public Function VirgulaPonto(pValor)
    Dim X As String
    Dim Y As String
    Dim i As Integer
    X = Trim(Str$(pValor))
    Y = ""
    For i = 1 To Len(X)
        If Mid$(X, i, 1) = "," Then
            Y = Y & "."
        Else
            Y = Y & Mid$(X, i, 1)
        End If
    Next
    VirgulaPonto = X
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function ConverteTimeStamp(ByVal sBinRep As ADODB.Field) As Date
If Not IsNull(sBinRep) = True Then
    ConverteTimeStamp = CDate(Format$(sBinRep, "DD/MM/YYYY HH:MM:SS"))
End If
End Function
'--Tratamento de teclas pressionadas em um DataCombo vinculado
'Campos padrão ID (para código retornado pelo DC) e Display (para valores listados)
'Argumentos: Teclas, datacombo, recordset e nome do item (mostrado ao usuário, caso não encontrado).
Public Sub CBKeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer, ByRef CB As DataCombo, ByRef Rs As ADODB.Recordset, ByVal Qual As String)
Select Case KeyCode
    Case Is = vbKeyF1
        With Rs
            .MoveFirst
            .Find "DISPLAY LIKE '" & CB & "%'", 0, adSearchForward, 0
            If Not .EOF = True Then
                CB.BoundText = .Fields("ID")
            Else
                CB.BoundText = Empty
                MsgBox Qual & CB.Text & " não foi encontrado.", vbOKOnly + vbInformation, "Tabelas de Cobrança"
                .MoveFirst
            End If
        End With
    Case Is = vbKeyReturn: KeyCode = vbKeyTab
    Case Is = vbKeyDelete
        CB.Text = Empty
        CB.BoundText = Empty
    Case Is = vbKeyBack: KeyCode = vbKeyBack
    Case Else: KeyCode = KeyCode
End Select
End Sub
Public Function IsEmail(ByVal Email As String) As Boolean

Dim w        As Integer
Dim sLetra   As String
Dim sSplit() As String
     
    On Error Resume Next
    
    If Len(Email) > 0 Then
        
        If UBound(Split(Email, "@")) = 0 Or InStr(Email, ".") = 0 Then
            Exit Function
        End If
        
        If Left$(Email, 1) = "@" Or Mid$(Email, Len(Email), 1) = "@" Or InStr(Email, "@.") Or InStr(Email, ".@") Then
            Exit Function
        End If

        If Left$(Email, 1) = "." Or Mid$(Email, Len(Email), 1) = "." Or InStr(Email, "..") Then
            Exit Function
        End If
        
        For w = 1 To Len(Email)
            sLetra = Mid$(Email, w, 1)
            If Not (LCase$(sLetra) Like "[a-z]" Or sLetra = ";" Or sLetra = "@" Or sLetra = "." Or sLetra = "-" Or sLetra = "_" Or IsNumeric(sLetra)) Then
                Exit Function
            End If
        Next w
        
        sSplit = Split(UCase$(Trim$(Email)), ".")

        If InStr(Dominios, sSplit(UBound(sSplit))) = 0 Then
            Exit Function
        End If
        
        IsEmail = True
    
    End If
   
   On Error GoTo 0
   
End Function
Public Function ValidaNumero(Par As String, Optional Dígitos As Integer = 2) As Double
Dim sRET As String

If CodEmpSis = "BAN" Then
    sRET = Replace(Par, ".", ",")
    ValidaNumero = CDbl(sRET)
Else
    If IsNumeric(Par) = True Then
        sRET = FormatNumber(Par, Dígitos, vbFalse, vbUseDefault, vbUseDefault)
        ValidaNumero = CDbl(sRET)
    Else
        ValidaNumero = CDbl(0)
    End If
End If
sRET = Empty
End Function

Public Sub PintaFaixa(ByRef Grid As Object, Row As Integer, Tipo As String, Coluna_Inicio, Coluna_Fim)
Dim k As Integer
Dim Cor As Variant
Grid.Row = Row
Select Case Tipo
    Case "VERDE"
        For k = Coluna_Inicio To Grid.Cols - 1
            Grid.Col = k
            Grid.CellBackColor = &HC0FFC0
            Grid.CellForeColor = vbBlack
        Next k
    Case "AMARELO"
        For k = Coluna_Inicio To Grid.Cols - 1
            Grid.Col = k
            Grid.CellBackColor = &HC0FFFF
            Grid.CellForeColor = vbBlack
        Next k
    Case "VERMELHO"
        For k = Coluna_Inicio To Grid.Cols - 1
            Grid.Col = k
            Grid.CellBackColor = &H8080FF
            Grid.CellForeColor = vbWhite
        Next k
    Case "AZUL"
        For k = Coluna_Inicio To Grid.Cols - 1
            Grid.Col = k
            Grid.CellBackColor = vbBlue
            Grid.CellForeColor = vbWhite
        Next k
    Case "ROSA"
        For k = Coluna_Inicio To Grid.Cols - 1
            Grid.Col = k
            Grid.CellBackColor = &HFEEADE
            Grid.CellForeColor = vbBlack
        Next k
End Select
End Sub

Public Sub ValidaPath(Caminho As String, Optional Criar As Boolean = True)


Dim sPar() As String
Dim N      As Integer
Dim P      As Integer
Dim sPasta As String
N = StrCount(Caminho, "\")
sPar = Split(Caminho, "\")
If Not Dir(sPar(0), vbDirectory) = Empty Then
    While Not P = (N + 1)
        sPasta = sPasta & sPar(P) & "\"
        P = P + 1
        If Not Dir(sPasta, vbDirectory) <> Empty Then MkDir (sPasta)
    Wend
End If
N = Empty
P = Empty
sPasta = Empty
Erase sPar
End Sub

Public Function StrCount(Valor As String, Caractere As String) As Long
Dim nRET As Long
Dim nCNT As Long
For nCNT = (Len(Valor) - Len(Valor) + 1) To Len(Valor)
    If Mid$(Valor, nCNT, 1) = Caractere Then nRET = nRET + 1
Next
StrCount = nRET
nCNT = Empty
nRET = Empty
End Function

Public Sub SetGrid(ByRef Grid As MSHFlexGrid, ByRef Rs As ADODB.Recordset, Optional ByVal HeaderString As String, Optional NumericLen As Integer = 10, Optional AutoNumber As Boolean = True, Optional ByVal SelByRow As Boolean = True)
On Error GoTo SGErr:
'"Liga" um MSHFlexGrid á um recordset ADO, montando Headers das colunas,
'alinhamentos por tipo de dados, tamanhos de colunas e
'ocultando colunas "indesejáveis" ao usuário.
'HeaderString é apenas uma string com os nomes desejados para as
'colunas, separados por barra vertical, sendo que duas barras seguidas,
'inibem a visualização da coluna.
Dim oFLD    As ADODB.Field
Dim sHSTemp As String
Dim HCols() As String
Dim nPos    As Long
If HeaderString = Empty Then
    HeaderString = HeaderString & "     |"
    For Each oFLD In Rs.Fields
        Select Case ReduzTiposDB(oFLD)
            Case Is = 0: HeaderString = HeaderString & ">" & Replace(oFLD.Name, "_", " ") & String$(NumericLen, " ") & "|"
            Case Is = 1: HeaderString = HeaderString & "<" & Replace(oFLD.Name, "_", " ") & String$(oFLD.DefinedSize, " ") & "|"
            Case Is = 2: HeaderString = HeaderString & ">" & Replace(oFLD.Name, "_", " ") & String$(oFLD.DefinedSize, " ") & "|"
        End Select
    Next
    HeaderString = HeaderString & "|"
Else
    HeaderString = Replace$(HeaderString, "^", Empty)
    HeaderString = Replace$(HeaderString, "<", Empty)
    HeaderString = Replace$(HeaderString, ">", Empty)
    HCols = Split(HeaderString, "|")
    If Len(HCols(0)) > 3 And Trim$(HCols(0)) = Empty Then
        sHSTemp = sHSTemp & String$(Len(HCols(0)), " ") & "|"
    Else
        sHSTemp = sHSTemp & "     |"
    End If
    For Each oFLD In Rs.Fields
        nPos = nPos + 1
        If Not nPos > UBound(HCols) Then
        If Not Trim$(HCols(nPos)) <> Empty Then
            sHSTemp = sHSTemp & "|"
        Else
            Select Case ReduzTiposDB(oFLD)
                Case Is = 0:
                    sHSTemp = sHSTemp & ">" & HCols(nPos) & String$(NumericLen, " ") & "|"
                Case Is = 1:
                    If Not Rs.EOF Then
                        If IsDate(oFLD) = True Then
                            sHSTemp = sHSTemp & "^" & HCols(nPos) & String$(Abs(oFLD.DefinedSize), " ") & "|"
                        Else
                            sHSTemp = sHSTemp & "<" & HCols(nPos) & String$(Abs(oFLD.DefinedSize), " ") & "|"
                        End If
                    Else
                        sHSTemp = sHSTemp & "<" & HCols(nPos) & String$(Abs(oFLD.DefinedSize), " ") & "|"
                    End If
                Case Is = 2:
                    sHSTemp = sHSTemp & "^" & HCols(nPos) & String$(Abs(oFLD.DefinedSize), " ") & "|"
            End Select
        End If
        End If
    Next
    sHSTemp = Mid$(sHSTemp, 1, Len(sHSTemp) - 1)
    HeaderString = sHSTemp
    sHSTemp = Empty
    nPos = Empty
End If
HeaderString = UCase$(HeaderString)
HCols = Split(HeaderString, "|")
'If Grid.Rows > Grid.FixedRows Then UnsetGrid Grid
With Grid
    .Visible = True
    Set .Recordset = Nothing
    Set .DataSource = Nothing
    .Appearance = flex3D
    .Appearance = flexFlat
    '.AddItem "A"
    .Clear
    .ClearStructure
    .FormatString = Empty
    .Cols = 1
    .Rows = 1
    .Refresh
    .Visible = False
    Set .Recordset = Nothing
    Set .DataSource = Nothing
    .Cols = 2
    .Rows = 2
    .FixedCols = 1
    .FixedRows = 1
    .Refresh
    .AllowBigSelection = False
    .AllowUserResizing = flexResizeBoth
    .PictureType = flexPictureMonochrome
    .ScrollBars = flexScrollBarBoth
    .ScrollTrack = True
    .Appearance = flexFlat
    If SelByRow = True Then
        .RowSizingMode = flexRowSizeAll
        .SelectionMode = flexSelectionByRow
    Else
        .RowSizingMode = flexRowSizeIndividual
        .SelectionMode = flexSelectionFree
    End If
    .FocusRect = flexFocusHeavy
    '.FontHeader(0).Bold = True
    '.Font.Bold = True
    .WordWrap = True
    'somente "liga" se houverem registros.
    If Not Rs Is Nothing Then
        If Not Rs.State = 0 Then
            Set .DataSource = Rs
        End If
    End If
    .FormatString = HeaderString
    .Cols = (UBound(Split(.FormatString, "|")) + .FixedCols)
    .Refresh
    For nPos = .FixedCols To (.Cols - .FixedCols)
        If Len(HCols(nPos)) <= 1 Then
            .ColHeaderCaption(0, nPos) = Empty
            .ColAlignment(nPos) = flexAlignLeftTop
            .ColWidth(nPos) = Empty
        End If
    Next
    .Visible = True
    If AutoNumber = True Then
        For nPos = 0 To Rs.RecordCount - 1
            .TextMatrix(nPos + 1, 0) = (nPos + 1)
        Next
    End If
End With
GoTo SGExit:
SGErr:
MsgBox Err.Description
Err.Clear
SGExit:
Erase HCols
Set oFLD = Nothing
sHSTemp = Empty
nPos = Empty
Erase HCols
End Sub
Private Sub SetLV(lv As ListView, Rs As ADODB.Recordset, Optional Headers As String)
Dim oLI     As ListItem
Dim oFL     As ADODB.Field
Dim N       As Long
Dim c       As Long
Dim sKey    As String
If Not Rs Is Nothing Then
    If Not Rs.State = 0 Then
        If Not Rs.EOF Then
            With lv
                .Visible = False
                .ListItems.Clear
                .ColumnHeaders.Clear
                For Each oFL In Rs.Fields
                    .ColumnHeaders.Add , oFL.Name, oFL.Name
                Next
                .View = lvwReport
                While Not Rs.EOF = True
                    For Each oFL In Rs.Fields
                        N = N + 1
                        sKey = (oFL.Name & "|" & Trim$(Str$(N)))
                        If oFL.Name = Rs.Fields(0).Name Then
                            Set oLI = .ListItems.Add(, sKey, Nnull(oFL.Value, 1), IIf(Not Rs.Fields("ALERTA") = 0, 2, 1), IIf(Not Rs.Fields("ALERTA") = 0, 2, 1))
                        Else
                            oLI.ListSubItems.Add = Nnull(oFL.Value, 1)
                        End If
                        c = c + 1
                    Next
                    c = Empty
                    Rs.MoveNext
                Wend
                .AllowColumnReorder = True
                .Appearance = ccFlat
                .Arrange = lvwAutoLeft
                .Checkboxes = False
                .FullRowSelect = True
                .GridLines = False
                .HotTracking = True
                .HoverSelection = True
                .LabelEdit = lvwManual
                .LabelWrap = True
                .MultiSelect = False
                .PictureAlignment = lvwTile
                .Sorted = True
                .TextBackground = lvwTransparent
                .View = lvwIcon
                .Visible = True
                .FlatScrollBar = False
            End With
        End If
    End If
End If
End Sub

Public Function InvaliDate(Valor As String) As Boolean
On Error GoTo IDErr:
Dim lRET As Boolean
lRET = (Trim$(Valor) = Empty)
If Not lRET = True Then
    lRET = (Not IsDate(Valor) = True)
    If Not lRET = True Then
        lRET = (DateDiff("D", CDate(Format$(Now, "DD/MM/YYYY")), Format$(Valor, "DD/MM/YYYY")) > 0)
    End If
End If
InvaliDate = lRET
GoTo IDExit:
IDErr:
MsgBox Err.Description
Err.Clear
IDExit:
lRET = Empty
End Function
'============================================================================================================================================
'           FUNÇÃO DE SUPORTE PARA SETLV E SETGRID.
'============================================================================================================================================
Private Function ReduzTiposDB(ByRef campo As ADODB.Field) As Integer
'0 = NÚMEROS
'1 = TEXTO
'2 = DATAS
If campo.Type = adArray Then ReduzTiposDB = 1
If campo.Type = adBigInt Then ReduzTiposDB = 0
If campo.Type = adBinary Then ReduzTiposDB = 0
If campo.Type = adBoolean Then ReduzTiposDB = 0
If campo.Type = adBSTR Then ReduzTiposDB = 1
If campo.Type = adChapter Then ReduzTiposDB = 1
If campo.Type = adChar Then ReduzTiposDB = 1
If campo.Type = adCurrency Then ReduzTiposDB = 0
If campo.Type = adDate Then ReduzTiposDB = 2
If campo.Type = adDBDate Then ReduzTiposDB = 2
If campo.Type = adDBTime Then ReduzTiposDB = 2
If campo.Type = adDBTimeStamp Then ReduzTiposDB = 2
If campo.Type = adDecimal Then ReduzTiposDB = 0
If campo.Type = adDouble Then ReduzTiposDB = 0
If campo.Type = adEmpty Then ReduzTiposDB = 1
If campo.Type = adError Then ReduzTiposDB = 0
If campo.Type = adFileTime Then ReduzTiposDB = 2
If campo.Type = adGUID Then ReduzTiposDB = 1
If campo.Type = adIDispatch Then ReduzTiposDB = 1
If campo.Type = adInteger Then ReduzTiposDB = 0
If campo.Type = adIUnknown Then ReduzTiposDB = 1
If campo.Type = adLongVarBinary Then ReduzTiposDB = 1
If campo.Type = adLongVarChar Then ReduzTiposDB = 1
If campo.Type = adLongVarWChar Then ReduzTiposDB = 1
If campo.Type = adNumeric Then ReduzTiposDB = 0
If campo.Type = adPropVariant Then ReduzTiposDB = 1
If campo.Type = adSingle Then ReduzTiposDB = 0
If campo.Type = adSmallInt Then ReduzTiposDB = 0
If campo.Type = adTinyInt Then ReduzTiposDB = 0
If campo.Type = adUnsignedBigInt Then ReduzTiposDB = 0
If campo.Type = adUnsignedInt Then ReduzTiposDB = 0
If campo.Type = adUnsignedSmallInt Then ReduzTiposDB = 0
If campo.Type = adUnsignedTinyInt Then ReduzTiposDB = 0
If campo.Type = adUserDefined Then ReduzTiposDB = 1
If campo.Type = adVarBinary Then ReduzTiposDB = 0
If campo.Type = adVarChar Then ReduzTiposDB = 1
If campo.Type = adVariant Then ReduzTiposDB = 1
If campo.Type = adVarNumeric Then ReduzTiposDB = 0
If campo.Type = adVarWChar Then ReduzTiposDB = 1
If campo.Type = adWChar Then ReduzTiposDB = 1
End Function





'rotinas já reformuládas - para aproveitamento na DLL


Public Function IsLoaded(Formulário As String) As Boolean
On Error GoTo ILErr:
Dim o       As VB.Form
Dim lRET    As Boolean
For Each o In Forms
    If Not lRET = True Then lRET = (Trim$(UCase$(o.Name)) = Trim$(UCase$(Formulário)))
    If lRET = True Then Exit For
Next
IsLoaded = lRET
GoTo ILExit:
ILErr:
MsgBox Err.Description
Err.Clear
ILExit:
Set o = Nothing
lRET = Empty
End Function

Public Sub PaintGrid(ByRef Grade As MSHFlexGrid, ByRef MatrizCores() As String, ByVal ColAval As Integer, Optional ByVal ColIni As Integer = 0, Optional ByVal ColFim As Integer = 0)
'Cada elemento da matriz deve ser um trio: BCor/FCor/Valor.
'BCor é a cor de fundo, FCor é a cor da letra e Valor é o
'valor que determina o uso da cor.
'Cada elemento deve vir como texto, separados os itens por
'uma barra vertical. (ex: "&H0BFFFEF|&H0BCFAEF|1").
'Base 0 para a matriz.
On Error GoTo PGErr:
Dim Elemento As Byte    'Elemento da matriz - Só pode receber até 254.
Dim wItem()  As String  'Cada ítem dos elementos é lido nessa matriz.
Dim c        As Integer 'Loop de colunas.
Dim R        As Integer 'Loop de linhas.
Dim nSel     As Integer 'Mantém a linha atualmente selecionada, se houver uma.
If Grade.Rows = Grade.FixedRows Then Exit Sub
With Grade
    .Visible = False
    nSel = .RowSel
    If (IsMissing(ColIni) = True Or ColIni = 0) Then ColIni = .LeftCol
    If (IsMissing(ColFim) = True Or ColFim = 0) Then ColFim = (.Cols - .FixedCols)
    For R = .TopRow To .Rows - 1
        If .RowIsVisible(R) = True Then
            For Elemento = 0 To (UBound(MatrizCores) - 1)
                wItem = Split(MatrizCores(Elemento), "|")
                If .TextMatrix(R, ColAval) = wItem(2) Then
                    For c = ColIni To ColFim
                        If .ColIsVisible(c) = True Then
                            .Row = R
                            .Col = c
                            .CellBackColor = wItem(0)
                            .CellForeColor = wItem(1)
                        Else
                            Exit For
                        End If
                    Next
                End If
            Next
        Else
            Exit For
        End If
    Next
    .Visible = True
    If .Visible = True Then .SetFocus
    .Row = nSel
    .RowSel = nSel
    .ColSel = .LeftCol
End With
GoTo PGExit:
PGErr:
MsgBox Err.Description
Err.Clear
PGExit:
c = Empty
R = Empty
Elemento = Empty
Erase wItem
End Sub


Public Sub GridPaint(ByRef Grid As MSHFlexGrid, ByVal EvalCol As Integer, ByRef iBackColor() As OLE_COLOR, ByRef iForeColor() As OLE_COLOR, Optional ByVal StartCol As Integer = 1, Optional ByVal EndCol As String = 1, Optional StartRow As Long = 1, Optional OnlyVisible As Boolean = True)
Dim nRow As Long
Dim nCol As Integer
If Not UBound(iForeColor) <> UBound(iBackColor) Then
    With Grid
        .Visible = False
        For nRow = StartRow To .Rows - 1
            If OnlyVisible = True Then
                If .RowIsVisible(nRow) = True Then
                    If Not (Nnull(.TextMatrix(nRow, EvalCol), 0) < LBound(iBackColor) Or Nnull(.TextMatrix(nRow, EvalCol), 0) > UBound(iBackColor)) Then
                        .Row = nRow
                        For nCol = StartCol To EndCol
                            If .ColIsVisible(nCol) = True Then
                                .Col = nCol
                                .CellBackColor = iBackColor(Nnull(.TextMatrix(nRow, EvalCol), 0))
                                .CellForeColor = iForeColor(Nnull(.TextMatrix(nRow, EvalCol), 0))
                            Else
                                Exit For
                            End If
                        Next
                    End If
                Else
                    Exit For
                End If
            Else
                If Not (Nnull(.TextMatrix(nRow, EvalCol), 0) < LBound(iBackColor) Or Nnull(.TextMatrix(nRow, EvalCol), 0) > UBound(iBackColor)) Then
                    .Row = nRow
                    For nCol = StartCol To EndCol
                        .Col = nCol
                        .CellBackColor = iBackColor(Nnull(.TextMatrix(nRow, EvalCol), 0))
                        .CellForeColor = iForeColor(Nnull(.TextMatrix(nRow, EvalCol), 0))
                    Next
                End If
            End If
        Next
        .Visible = True
    End With
End If
nRow = Empty
End Sub

Public Function Máquina() As String
On Error GoTo MQErr:
Dim sRET As String
sRET = String$(255, " ")
GetComputerName sRET, 255
sRET = Trim$(UCase(sRET))
sRET = Replace$(sRET, Chr$(0), Empty)
sRET = Replace$(sRET, Chr$(1), Empty)
sRET = Replace$(sRET, Chr$(34), Empty)
sRET = Replace$(sRET, "'", Empty)
Máquina = sRET
GoTo MQExit:
MQErr:
MsgBox Err.Description
Err.Clear
MQExit:
sRET = Empty
End Function

Public Function Usuário() As String
On Error GoTo USErr:
Dim sRET As String
sRET = String$(255, " ")
GetUserName sRET, 255
sRET = Trim$(UCase(sRET))
sRET = Replace$(sRET, Chr$(0), Empty)
sRET = Replace$(sRET, Chr$(1), Empty)
sRET = Replace$(sRET, Chr$(34), Empty)
sRET = Replace$(sRET, "'", Empty)
Usuário = sRET
GoTo USExit:
USErr:
MsgBox Err.Description
Err.Clear
USExit:
sRET = Empty
End Function

Public Sub MontaMSHFlexGrid(ByRef CB As MSHFlexGrid, ByRef Rs As ADODB.Recordset)
    Set CB.DataSource = Rs
End Sub

'Public Sub MontaDataGrid(ByRef CB As DataGrid, ByRef Rs As ADODB.Recordset)
'    Set CB.DataSource = Rs
'End Sub

Public Sub MontaCombo(ByRef CB As DataCombo, ByRef Rs As ADODB.Recordset, Optional LF As String = "DISPLAY", Optional BC As String = "ID", Optional Mostrar As Boolean = False)
With CB
    Set .RowSource = Rs
    .ListField = LF
    .BoundColumn = BC
    If Mostrar = True Then
        If Not Rs.EOF Then
            If Not Rs.BOF = True Then Rs.MoveFirst
            .BoundText = Rs.Fields(BC)
        End If
    End If
End With
End Sub

Public Sub MontaList(ByRef CB As DataList, ByRef Rs As ADODB.Recordset, Optional LF As String = "DISPLAY", Optional BC As String = "ID", Optional Mostrar As Boolean = False)
With CB
    Set .RowSource = Rs
    .ListField = LF
    .BoundColumn = BC
    If Mostrar = True Then
        If Not Rs.EOF Then
            If Not Rs.BOF = True Then Rs.MoveFirst
            .BoundColumn = Rs.Fields(BC)
        End If
    End If
End With
End Sub


Public Sub UnsetGrid(ByRef Grade As MSHFlexGrid)
Dim sHEAD As String
sHEAD = Grade.FormatString
With Grade
    .Visible = False
    Set .DataSource = Nothing
    .Clear
    .ClearStructure
    .Cols = 0
    .Cols = 2
    .Rows = 0
    .Rows = 2
    .FixedCols = 1
    .FixedRows = 1
    .FormatString = SHEADE
    '.Enabled = False
    .Visible = True
End With
End Sub

Public Function RemoveCaracteres(TEXTO As String) As String
'Monta o texto de retorno removendo separadores e substituindo
'os Underlines por espaços (ID_CONTEINER).
TEXTO = Replace(TEXTO, "-", Empty)
TEXTO = Replace(TEXTO, ".", Empty)
TEXTO = Replace(TEXTO, ",", Empty)
TEXTO = Replace(TEXTO, "/", Empty)
TEXTO = Replace(TEXTO, "\", Empty)
TEXTO = Replace(TEXTO, "_", " ")
RemoveCaracteres = TEXTO
End Function

Public Function CentraTexto(TEXTO As String, Tamanho As Integer, Optional Caractere As String) As String
Dim sRET As String
If IsMissing(Caractere) = True Then Caractere = Space(1)
If Trim(Caractere) = Empty Then Caractere = Space(1)
If Len(TEXTO) > Tamanho Then TEXTO = Mid$(TEXTO, 1, Tamanho)
If Len(TEXTO) >= Tamanho Then
    If Len(TEXTO) > Tamanho Then Tamanho = Len(TEXTO)
    sRET = String$(CInt((Tamanho - Len(TEXTO)) / 2), Caractere)
Else
    sRET = String$(CInt((Tamanho - Len(TEXTO)) / 2) - 1, Caractere)
End If
If Len(sRET & TEXTO & sRET) <= Tamanho Then
    sRET = sRET & TEXTO & sRET
Else
    sRET = TEXTO
End If
If Len(sRET) > Tamanho Then sRET = Mid$(sRET, 1, Tamanho)
If Len(sRET) < Tamanho Then sRET = sRET & String(Tamanho - Len(sRET), Caractere)
CentraTexto = sRET
End Function

Public Function AlinhaTexto(TEXTO As String, Tamanho As Integer, Optional Caractere As String, Optional Esquerda As Boolean = True) As String
Dim sRET As String
If IsMissing(Caractere) = True Then Caractere = Space(1)
If Trim(Caractere) = Empty Then Caractere = Space(1)
If Len(TEXTO) > Tamanho Then TEXTO = Mid$(TEXTO, 1, Tamanho)
sRET = String$(CInt((Tamanho - Len(TEXTO))), Caractere)
If Len(sRET) <= Tamanho Then
    If Esquerda = True Then
        sRET = TEXTO & sRET
    Else
        sRET = TEXTO & sRET
    End If
Else
    sRET = TEXTO
End If
If Len(sRET) > Tamanho Then sRET = Mid$(sRET, 1, Tamanho)
If Len(sRET) < Tamanho Then sRET = sRET & String(Tamanho - Len(sRET), Caractere)
AlinhaTexto = sRET
End Function

Public Sub GridImage(ByRef Grade As MSHFlexGrid, _
                     ByVal ImagemSim As StdPicture, _
                     ByVal ImagemNao As StdPicture, _
                     Colunas() As Integer, _
                     Optional ApagaValor As Boolean = False)
On Error GoTo GIErr:
'Insere uma das imagens enviadas como parâmetro
'nas colunas definidas, mediante a avaliação do
'valor da própria coluna.
'o valor da célula á ser testada deve ser sempre
'1 (verdadeiro/sim) ou 0 (falso/não).
Dim nCols As Integer    '"varre" a matriz de colunas á "desenhar"
Dim R     As Long       'loop de linhas
Dim c     As Long       'loop de colunas
Dim nSel  As Integer    'Mantém a linha atualmente selecionada, se houver.
If Grade.Rows = Grade.FixedRows Then Exit Sub
With Grade
    nSel = .RowSel
    .Visible = False
    For R = .TopRow To (.Rows - .FixedRows)
        If .RowIsVisible(R) = True Then
            For nCols = LBound(Colunas) To UBound(Colunas) - 1  'assumir que a matriz se inicia com zero
                If .ColIsVisible(Colunas(nCols)) = True Then
                    .Row = R
                    .Col = Colunas(nCols)
                    If Val(.TextMatrix(R, Colunas(nCols))) = 1 Then
                        Set .CellPicture = ImagemSim
                        .CellPictureAlignment = flexAlignCenterCenter
                    Else
                        Set .CellPicture = ImagemNao
                        .CellPictureAlignment = flexAlignCenterCenter
                    End If
                    .CellForeColor = .CellBackColor
                    .ColAlignment(.Col) = flexAlignCenterCenter
                    If ApagaValor = True Then
                        .TextMatrix(R, Colunas(nCols)) = Empty
                    Else
                        .CellFontSize = 1
                    End If
                Else
                    Exit For
                End If
            Next
        Else
            Exit For
        End If
    Next
    .Visible = True
    If .Visible = True Then .SetFocus
    .Row = nSel
    .RowSel = nSel
    .Col = .LeftCol
End With
GIErr:
GIExit:
Grade.Visible = True
nCols = Empty
R = Empty
c = Empty
nSel = Empty
End Sub


Public Function LimpaEntradaTexto(ByVal TEXTO As String, Optional ByVal RemoveAcentos As Boolean = False) As String
'REMOVE CARACTERES QUE POSSAM SER INTERPRETADOS
'POR ENGINES DE DADOS COMO CLÁUSULAS, PARÂMETROS.
'CORINGAS OU DELIMITADORES.
TEXTO = Replace$(TEXTO, Chr$(34), Empty) '(")
TEXTO = Replace$(TEXTO, Chr$(35), Empty) '(#)
TEXTO = Replace$(TEXTO, Chr$(36), Empty) '($)
TEXTO = Replace$(TEXTO, Chr$(37), Empty) '(%)
TEXTO = Replace$(TEXTO, Chr$(38), Empty) '(&)
TEXTO = Replace$(TEXTO, Chr$(39), Empty) '(')
TEXTO = Replace$(TEXTO, Chr$(42), Empty) '(*)
TEXTO = Replace$(TEXTO, Chr$(64), Empty) '(@)
TEXTO = Replace$(TEXTO, Chr$(63), Empty) '(?)
If RemoveAcentos = True Then
    TEXTO = Replace$(TEXTO, "á", "a")
    TEXTO = Replace$(TEXTO, "à", "a")
    TEXTO = Replace$(TEXTO, "ã", "a")
    TEXTO = Replace$(TEXTO, "ä", "a")
    TEXTO = Replace$(TEXTO, "â", "a")
    TEXTO = Replace$(TEXTO, "é", "e")
    TEXTO = Replace$(TEXTO, "è", "e")
    TEXTO = Replace$(TEXTO, "ê", "e")
    TEXTO = Replace$(TEXTO, "ë", "e")
    TEXTO = Replace$(TEXTO, "í", "i")
    TEXTO = Replace$(TEXTO, "ì", "i")
    TEXTO = Replace$(TEXTO, "î", "i")
    TEXTO = Replace$(TEXTO, "ï", "i")
    TEXTO = Replace$(TEXTO, "ó", "o")
    TEXTO = Replace$(TEXTO, "ò", "o")
    TEXTO = Replace$(TEXTO, "õ", "o")
    TEXTO = Replace$(TEXTO, "ö", "o")
    TEXTO = Replace$(TEXTO, "ô", "o")
    TEXTO = Replace$(TEXTO, "ú", "u")
    TEXTO = Replace$(TEXTO, "ù", "u")
    TEXTO = Replace$(TEXTO, "û", "u")
    TEXTO = Replace$(TEXTO, "ü", "u")
    TEXTO = Replace$(TEXTO, "Á", "A")
    TEXTO = Replace$(TEXTO, "À", "A")
    TEXTO = Replace$(TEXTO, "Ã", "A")
    TEXTO = Replace$(TEXTO, "Â", "A")
    TEXTO = Replace$(TEXTO, "Ä", "A")
    TEXTO = Replace$(TEXTO, "É", "E")
    TEXTO = Replace$(TEXTO, "È", "E")
    TEXTO = Replace$(TEXTO, "Ê", "E")
    TEXTO = Replace$(TEXTO, "Ë", "E")
    TEXTO = Replace$(TEXTO, "Í", "I")
    TEXTO = Replace$(TEXTO, "Ì", "I")
    TEXTO = Replace$(TEXTO, "Î", "I")
    TEXTO = Replace$(TEXTO, "Ï", "I")
    TEXTO = Replace$(TEXTO, "Ó", "O")
    TEXTO = Replace$(TEXTO, "Ò", "O")
    TEXTO = Replace$(TEXTO, "Õ", "O")
    TEXTO = Replace$(TEXTO, "Ô", "O")
    TEXTO = Replace$(TEXTO, "Ö", "O")
    TEXTO = Replace$(TEXTO, "Ú", "U")
    TEXTO = Replace$(TEXTO, "Ù", "U")
    TEXTO = Replace$(TEXTO, "Û", "U")
    TEXTO = Replace$(TEXTO, "Ü", "U")
    TEXTO = Replace$(TEXTO, "Ç", "C")
    TEXTO = Replace$(TEXTO, "ç", "c")
End If
LimpaEntradaTexto = TEXTO
End Function
Public Function CGCValido(Optional CGC As String = vbEmpty) As Boolean
Dim a As Integer
Dim j As Integer
Dim i As Integer
Dim D1 As Integer
Dim D2 As Integer
If IsEmpty(CGC) = True Then Exit Function
CGC = Trim$(Replace$(Replace$(Replace$(CGC, ".", Empty), "/", Empty), "-", Empty))
If Not IsNumeric(CGC) Then Exit Function
If Len(CGC) > 0 Then
    If CGC = String$(Len(CGC), Mid$(CGC, 1, 1)) Then
        CGCValido = False
        Exit Function
    End If
Else
    CGCValido = False
    Exit Function
End If
If Len(CGC) = 8 And Val(CGC) > 0 Then
    a = 0
    j = 0
    D1 = 0
    For i = 1 To 7
        a = Val(Mid(CGC, i, 1))
        If (i Mod 2) <> 0 Then a = a * 2
        If a > 9 Then
            j = j + Int(a / 10) + (a Mod 10)
        Else
            j = j + a
        End If
    Next i
    D1 = IIf((j Mod 10) <> 0, 10 - (j Mod 10), 0)
    If D1 = Val(Mid(CGC, 8, 1)) Then
        CGCValido = True
    Else
        CGCValido = False
    End If
Else
    If Len(CGC) = 14 And Val(CGC) > 0 Then
        a = 0
        i = 0
        D1 = 0
        D2 = 0
        j = 5
        For i = 1 To 12 Step 1
            a = a + (Val(Mid(CGC, i, 1)) * j)
            j = IIf(j > 2, j - 1, 9)
        Next i
        a = a Mod 11
        D1 = IIf(a > 1, 11 - a, 0)
        a = 0
        i = 0
        j = 6
        For i = 1 To 13 Step 1
            a = a + (Val(Mid(CGC, i, 1)) * j)
            j = IIf(j > 2, j - 1, 9)
        Next i
        a = a Mod 11
        D2 = IIf(a > 1, 11 - a, 0)
        If (D1 = Val(Mid(CGC, 13, 1)) And D2 = Val(Mid(CGC, 14, 1))) Then
           CGCValido = True
        Else
           CGCValido = False
        End If
    Else
        CGCValido = False
    End If
End If
End Function

Public Function CPFValido(CpfEnt As String) As Boolean
Dim Soma As Integer
Dim Resto As Integer
Dim i As Integer
Dim CPF As String
On Error Resume Next
CPF = CpfEnt
If IsEmpty(CPF) = True Then Exit Function
CPF = Trim$(Replace$(Replace$(Replace$(CPF, ".", Empty), "/", Empty), "-", Empty))
If Val(CPF) = 0 Then Exit Function
If Len(CPF) > 0 Then
    If CPF = String$(Len(CPF), Mid$(CPF, 1, 1)) Then
        CPFValido = False
        Exit Function
    End If
Else
    CPFValido = False
    Exit Function
End If
If Val(CPF) = 0 Then
    CPFValido = False
    Exit Function
End If
If Len(CPF) < 11 Then
    CPF = Right$("00000000000" & CPF, 11)
End If
Soma = 0
For i = 1 To 9
    Soma = Soma + Val(Mid$(CPF, i, 1)) * (11 - i)
Next i
Resto = 11 - (Soma - (Int(Soma / 11) * 11))
If Resto = 10 Or Resto = 11 Then Resto = 0
If Resto <> Mid$(CPF, 10, 1) Then
    CPFValido = False
    Exit Function
End If
Soma = 0
For i = 1 To 10
    Soma = Soma + Val(Mid$(CPF, i, 1)) * (12 - i)
Next i
Resto = 11 - (Soma - (Int(Soma / 11) * 11))
If Resto = 10 Or Resto = 11 Then Resto = 0
If Resto <> Val(Mid(CPF, 11, 1)) Then
    CPFValido = False
    Exit Function
End If
CPFValido = True
End Function

Public Sub Cadastrar_Menus()

Dim Mnu As Control, Aux As Integer
Dim NumFunc As Long
If Nnull(db.Execute("Select count(*) as Regs from " & Banco_Sgipa & "TB_SYS_FUNCOES where UPPER(Sistema) = '" & UCase(App.ProductName) & "'")(0).Value, 0) = 0 Then
  For Each Mnu In Principal
      If TypeOf Mnu Is Menu Then
          If Mnu.Caption <> "-" And Mnu.Caption <> "&-" And Mnu.Visible = True Then _
              Executa "INSERT into " & Banco_Sgipa & "TB_SYS_FUNCOES (codfunc, nomeobj, captionobj,sistema) values ((select count(codfunc)+ 1 as CodFunc from " & Banco_Sgipa & "TB_SYS_FUNCOES), '" & Mnu.Name & "', '" & Replace(Mnu.Caption, "&", "") & "', '" & App.ProductName & "')"
      End If
  Next
  
  Executa "DELETE from " & Banco_Sgipa & "tb_sys_grp_permissoes where codgrupo=1"
  For Aux = 1 To 5
    Executa "INSERT INTO " & Banco_Sgipa & "tb_sys_grp_permissoes (codfunc, codgrupo, codtipoperm) " & _
    "(SELECT tb_sys_funcoes.codfunc, 1 as CodGrupo, " & Aux & " as CodTipoPerm  FROM " & Banco_Sgipa & "tb_sys_funcoes where tb_sys_funcoes.sistema= '" & App.ProductName & "')"
  Next
  
Else
  For Each Mnu In Principal
      If TypeOf Mnu Is Menu And tbSenha!CodGrupo = 1 Then
         If Mnu.Caption <> "-" And Mnu.Caption <> "&-" And Mnu.Visible = True Then
            If Nnull(db.Execute("Select count(*) as Regs from " & Banco_Sgipa & "TB_SYS_FUNCOES where UPPER(Sistema) = '" & UCase(App.ProductName) & "' and UPPER(NomeObj) = '" & UCase(Mnu.Name) & "'")(0).Value, 0) = 0 Then
                NumFunc = db.Execute("select count(codfunc)+ 1 as CodFunc from " & Banco_Sgipa & "TB_SYS_FUNCOES")(0).Value
                Executa "INSERT into " & Banco_Sgipa & "TB_SYS_FUNCOES (codfunc, nomeobj, captionobj,sistema) values (" & NumFunc & ", '" & Mnu.Name & "', '" & Replace(Mnu.Caption, "&", "") & "', '" & App.ProductName & "')"
                For Aux = 1 To 5
                  Executa "INSERT INTO " & Banco_Sgipa & "tb_sys_grp_permissoes (codfunc, codgrupo, codtipoperm) " & _
                  "Values (" & NumFunc & ", 1," & Aux & ")"
                Next
            End If
          End If
      End If
  Next
  
  'DB.Execute "alter trigger " & Banco_Sgipa & "tb_sys_grp_permissoes_log disable"
  'executa  "DELETE from " & Banco_Sgipa & "tb_sys_grp_permissoes where codgrupo=1"
  'For Aux = 1 To 5
   ' executa  "INSERT INTO " & Banco_Sgipa & "tb_sys_grp_permissoes (codfunc, codgrupo, codtipoperm) " & _
    '"(SELECT tb_sys_funcoes.codfunc, 1 as CodGrupo, " & Aux & " as CodTipoPerm  FROM " & Banco_Sgipa & "tb_sys_funcoes where tb_sys_funcoes.sistema= '" & App.ProductName & "')"
  'Next
 ' Db.Execute "alter trigger " & Banco_Sgipa & "tb_sys_grp_permissoes_log enable"
  
End If


End Sub

Public Sub Controlar_Acesso_Mnus()
Dim Mnu As Control
Dim Rs As New ADODB.Recordset
For Each Mnu In Principal
    If TypeOf Mnu Is Menu Then
       If Mnu.Caption <> "-" And Mnu.Caption <> "&-" And Mnu.Visible = True Then
            'sql = "SELECT COUNT(1) CONTAR "
            'sql = sql & " FROM"
            'sql = sql & " " & Banco_Sgipa & "TB_SYS_FUNCOES"
            'sql = sql & " ," & Banco_Sgipa & "TB_SYS_GRP_PERMISSOES"
            'sql = sql & " WHERE"
            'sql = sql & " " & Banco_Sgipa & "TB_SYS_FUNCOES.codfunc = " & Banco_Sgipa & "TB_SYS_GRP_PERMISSOES.codfunc"
            'sql = sql & " AND " & Banco_Sgipa & "TB_SYS_GRP_PERMISSOES.codgrupo=" & tbSenha!CodGrupo
            'sql = sql & " AND UPPER(" & Banco_Sgipa & "TB_SYS_FUNCOES.nomeobj) = '" & UCase(Mnu.Name) & "'"
            'sql = sql & " AND UPPER(" & Banco_Sgipa & "TB_SYS_FUNCOES.sistema) = '" & UCase(App.ProductName) & "'"
            
            SQl = "SELECT COUNT(1) CONTAR "
            SQl = SQl & " FROM"
            SQl = SQl & " " & Banco_Sgipa & "TB_SYS_FUNCOES"
            SQl = SQl & " ," & Banco_Sgipa & "TB_SYS_GRP_PERMISSOES"
            SQl = SQl & " ," & Banco_Sgipa & "TB_SYS_USER_GRUPOS "
            SQl = SQl & " ," & Banco_Sgipa & "TB_CAD_USUARIOS "
            SQl = SQl & " WHERE"
            SQl = SQl & "((tb_cad_usuarios.autonum = tb_sys_user_grupos.autonumuser) "
            SQl = SQl & " AND (tb_sys_funcoes.codfunc = tb_sys_grp_permissoes.codfunc) "
            SQl = SQl & " AND (tb_sys_grp_permissoes.codgrupo = tb_sys_user_grupos.codgrupo) "
            SQl = SQl & " AND (tb_cad_usuarios.usuario = '" & Usuario_Sistema & "') "
            SQl = SQl & " AND (tb_sys_funcoes.sistema = '" & App.ProductName & "') "
            SQl = SQl & " AND (UPPER(tb_sys_funcoes.Nomeobj) = '" & UCase(Mnu.Name) & "')) "
            
            PRSet Rs, SQl
            If Rs.EOF Then
                Mnu.Enabled = False
            Else
              If Val(Nnull(Rs!Contar, 0)) = 0 Then Mnu.Enabled = False
            End If
        End If
                
                
                   
                              
    End If
Next
End Sub





Public Sub Controlar_Acesso_Forms(frm As Form, MnuOwnerForm As String)
Dim Control As Control
For Each Control In frm

'MsgBox Control.Name

If TypeOf Control Is CommandButton Or TypeOf Control Is CheckBox Or TypeOf Control Is OptionButton Then


    If Control.Tag <> "" And IsNumeric(Control.Tag) Then
                     
        If Usuario_Sistema = "MICROLED" Then
            Control.Visible = True
        Else
           If db.Execute("SELECT COUNT(1) " & _
                     "  FROM " & Banco_Sgipa & "tb_cad_usuarios," & Banco_Sgipa & "tb_sys_grp_permissoes, " & _
                     "  " & Banco_Sgipa & "tb_sys_user_grupos , " & Banco_Sgipa & "tb_sys_funcoes " & _
                     " WHERE ((tb_sys_user_grupos.codgrupo = tb_sys_grp_permissoes.codgrupo) " & _
                     " AND (tb_cad_usuarios.autonum = tb_sys_user_grupos.autonumuser) " & _
                     " AND (tb_sys_funcoes.codfunc = tb_sys_grp_permissoes.codfunc) " & _
                     " AND (tb_cad_usuarios.usuario = '" & Usuario_Sistema & "') " & _
                     " AND (UPPER(tb_sys_funcoes.nomeobj) = '" & UCase(MnuOwnerForm) & "') " & _
                     " AND (tb_sys_grp_permissoes.codtipoperm = " & Control.Tag & "))")(0).Value = 0 Then
                     Control.Visible = False
                     End If
        
        End If
    End If
End If
Next
End Sub

Public Function CentralizaTexto(TEXTO As String, _
                                Tamanho As Long, _
                                Optional Caractere As String = " ")
Dim strAux As String
Dim lngEsquerda As Long
Dim lngDireita As Long

If Len(TEXTO) >= Tamanho Then
  CentralizaTexto = Left$(TEXTO, Tamanho)
  Exit Function
End If

lngEsquerda = (Tamanho - Len(TEXTO)) \ 2
lngDireita = lngEsquerda + (Tamanho - Len(TEXTO)) Mod 2

strAux = String(lngEsquerda, Caractere) & TEXTO
strAux = strAux & String(lngDireita, Caractere)
CentralizaTexto = strAux
End Function

Public Sub P_Execute(P_Sql As String, Optional Tipo1 As TpDados, Optional P_Var1 As String, Optional Tipo2 As TpDados, Optional P_Var2 As String, Optional Tipo3 As TpDados, Optional P_Var3 As String, Optional Tipo4 As TpDados, Optional P_Var4 As String, Optional Tipo5 As TpDados, Optional P_Var5 As String, Optional Tipo6 As TpDados, Optional P_Var6 As String, Optional Tipo7 As TpDados, Optional P_Var7 As String, Optional Tipo8 As TpDados, Optional P_Var8 As String, Optional Tipo9 As TpDados, Optional P_Var9 As String)
Dim cmd As New ADODB.Command
Dim t As Byte
Dim Tp_Sel(10) As TpDados
Dim Variavel(10) As String

cmd.ActiveConnection = ConnectionString
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "PROC_CHRONOS_EXECUTE"

Tp_Sel(1) = Tipo1
Variavel(1) = P_Var1
Tp_Sel(2) = Tipo2
Variavel(2) = P_Var2
Tp_Sel(3) = Tipo3
Variavel(3) = P_Var3
Tp_Sel(4) = Tipo4
Variavel(4) = P_Var4
Tp_Sel(5) = Tipo5
Variavel(5) = P_Var5
Tp_Sel(6) = Tipo6
Variavel(6) = P_Var6
Tp_Sel(7) = Tipo7
Variavel(7) = P_Var7
Tp_Sel(8) = Tipo8
Variavel(8) = P_Var8
Tp_Sel(9) = Tipo9
Variavel(9) = P_Var9

For t = 1 To 9
   If Nnull(Variavel(t), 1) <> "" Then
      cmd.Parameters.Append cmd.CreateParameter("Var_Par" & t, adVarChar, adParamInput, 500, Variavel(t))
   Else
      cmd.Parameters.Append cmd.CreateParameter("Var_Par" & t, adVarChar, adParamInput, 500, Null)
   End If
Next t
cmd.Parameters.Append cmd.CreateParameter("SSQL", adVarChar, adParamInput, 5000, P_Sql)
cmd.Parameters.Append cmd.CreateParameter("ErroCode", adVarChar, adParamOutput, 100)
cmd.Execute
If cmd.Parameters("ErroCode").Value <> "" Then
   MsgBox P_Sql
   MsgBox "Error Code: " & cmd.Parameters("ErroCode").Value
End If
Set cmd = Nothing
End Sub
Public Sub CriarDDF(ByVal Recordset As ADODB.Recordset, ByVal locDDF As String)
Dim retval As Long
    
     'Criação do ficheiro DDF
     retval = CreateFieldDefinitionFile(Recordset, locDDF, False)
     If (retval <> 1) Then
          MsgBox "Erro na criação do DDF!", vbOKOnly + vbCritical, "Erro"
     End If
 
    
    Set Recordset = Nothing
End Sub
Public Sub CriarRPT(ByVal Recordset As ADODB.Recordset, ByVal locRel As String, ByVal locDDF As String)
Dim retval As Long

     'Chamar a API para a criação do relatório
     retval = CreateReportOnRuntime(Recordset, locRel, locDDF, True, True)
     If (retval <> 1) Then
          MsgBox "Erro na criação do relatório!", vbOKOnly + vbCritical, "Erro"
     End If
     
     '
     Set Recordset = Nothing

End Sub


Public Function XFILE_COPY(Origem As String, Destino As String) As Boolean
    
    FileCopy Origem, Destino
    If Not Verifica_Destino(Destino) Then
        XFILE_COPY = False
        MsgBox "ERRO NA GRAVAÇÃO DA PASTA PACOTES - FAVOR REENVIAR O PACOTE"
    Else
        XFILE_COPY = True
    End If

End Function


Public Function eBandeirantes(Parc As Long, Lote As Long, Optional PermiteBandB2 As Boolean = False) As Boolean
On Error GoTo trataErro:
eBandeirantes = False
Dim Emp As Long
Dim sSql As String
Dim pCNPJ As String

Emp = 0
If PermiteBandB2 Then
    sSql = Empty
    sSql = "SELECT NVL(CODIGO_EMPRESA,0) FROM " & Banco_Sgipa & "TB_BL WHERE AUTONUM =  " & Nnull(Lote, 0)
    Emp = Nnull(db.Execute(sSql).Fields(0), 0)
    Emp = IIf(Emp = 0, AutonumCIABand, AutonumBracenter)
End If

sSql = Empty
sSql = "SELECT CGC FROM " & Banco_Sgipa & "TB_CAD_PARCEIROS WHERE AUTONUM = " & Nnull(Parc, 0)
pCNPJ = Nnull(db.Execute(sSql).Fields(0), 1)


sSql = Empty
sSql = "SELECT COUNT(1) AS QTD "
sSql = sSql & " FROM " & Banco_Sgipa & "TB_EMPRESAS "
sSql = sSql & " WHERE CNPJ ='" & pCNPJ & "' "
If Emp > 0 Then sSql = sSql & " AND AUTONUM = " & Emp
If Nnull(db.Execute(sSql).Fields(0), 0) > 0 Then
    eBandeirantes = True
End If

Exit Function
trataErro:
    Err.Description = Empty
    eBandeirantes = False


End Function

Sub Bloqueio_BL_Parceiro(Id_BL As Long)

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
SQl = SQl & "IMPORTADOR,DESPACHANTE,PATIO FROM " & Banco_Sgipa & "TB_BL WHERE AUTONUM=" & Id_BL

PRSet Tbbl, SQl





If Not Tbbl.EOF Then

    bloqueio_anterior = Tbbl!FLAG_BLOQUEIO_MANUAL
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
       bloqueio_Atual = 0
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
    Call P_Bloqueio(Id_BL, Trim(Str(MotivoBloq)), "B", Usuario_Sistema, wCpf_Usuario)
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
    SQl = SQl & " GETDATE(), "
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
    bloqueio_Atual = 1
    If bloqueio_Atual <> bloqueio_anterior Then
        
        'atualizar na integraçao
        If Tbbl!Patio = "7" Then
            
        End If
        
    End If
End If
Return

End Sub


Public Function geraChaveDeposito(formaPagto As Integer, Lote As Long) As String
Dim sSql As String
Dim idChave As Long
Dim rsaux As New ADODB.Recordset
On Error GoTo trataErro:
geraChaveDeposito = ""
'Exit Function
'NÃO ENTROU NO AR

If formaPagto = 2 Then
    sSql = "SELECT ISNULL(FLAG_CHDEP_AV,0) FROM " & Banco_Sgipa & "TB_EMPRESAS WHERE AUTONUM = " & Cod_Empresa
    If Nnull(db.Execute(sSql).Fields(0), 0) = 0 Then Exit Function
ElseIf formaPagto = 3 Then
    sSql = "SELECT ISNULL(FLAG_CHDEP_FAT,0) FROM " & Banco_Sgipa & "TB_EMPRESAS WHERE AUTONUM = " & Cod_Empresa
    If Nnull(db.Execute(sSql).Fields(0), 0) = 0 Then Exit Function
Else
    Exit Function
End If

If Nnull(Lote, 0) = 0 Then
    Exit Function
Else
    sSql = "SELECT ISNULL(FLAG_SEM_DEP_IDENT,0) FROM " & Banco_Sgipa & "TB_CAD_PARCEIROS WHERE AUTONUM IN("
    sSql = sSql & " SELECT IMPORTADOR FROM " & Banco_Sgipa & "TB_BL WHERE AUTONUM = " & Lote & ") "
    If Nnull(db.Execute(sSql).Fields(0), 0) = 1 Then Exit Function

    sSql = "SELECT CH.AUTONUM, CH.CHAVE, GR.SEQ_GR FROM " & Banco_Sgipa & "TB_CHAVE_DEPOSITO CH "
    sSql = sSql & " LEFT JOIN " & Banco_Sgipa & "TB_GR_BL GR ON CH.CHAVE = GR.CHAVE_DEPOSITO "
    sSql = sSql & " WHERE ISNULL(CH.SEQ_GR,0) = 0 AND CH.LOTE =  " & Lote
    PRSet rsaux, sSql
    While Not rsaux.EOF
        If Nnull(rsaux!SEQ_GR, 0) = 0 Then
            geraChaveDeposito = Nnull(rsaux!Chave, 1)
            Exit Function
        Else
            sSql = " UPDATE " & Banco_Sgipa & "TB_CHAVE_DEPOSITO SET SEQ_GR = " & Nnull(rsaux!SEQ_GR, 0)
            sSql = sSql & " WHERE AUTONUM = " & Nnull(rsaux!autonum, 0)
            db.Execute sSql
        End If
        
        rsaux.MoveNext
    Wend

    If Nnull(geraChaveDeposito, 1) = "" Then
        sSql = "INSERT INTO " & Banco_Sgipa & "TB_CHAVE_DEPOSITO (LOTE) VALUES (" & Lote & ") "
        db.Execute (sSql)
        
        idChave = Nnull(db.Execute("SELECT MAX(AUTONUM) FROM " & Banco_Sgipa & "TB_CHAVE_DEPOSITO WHERE LOTE = " & Lote).Fields(0), 1)
        
        If Nnull(idChave, 0) > 0 Then
            geraChaveDeposito = calculaChaveDeposito(idChave)
            
            sSql = "UPDATE " & Banco_Sgipa & "TB_CHAVE_DEPOSITO SET CHAVE = '" & geraChaveDeposito & "' "
            sSql = sSql & " WHERE AUTONUM =" & Nnull(idChave, 0)
            db.Execute sSql
        End If
        
        
    
    End If

End If


Exit Function

trataErro:
    Err.Clear



End Function


Public Function calculaChaveDeposito(Chave) As String
Dim mult As Long
Dim Soma As Long
Dim dig As Long
Dim Base As String
On Error GoTo trataErro:

calculaChaveDeposito = ""
'Exit Function
'NÃO ENTROU NO AR

calculaChaveDeposito = Format(Chave, "0##############")

For X = 15 To 1 Step -1
    If X Mod 2 = 0 Then
        mult = mult + Mid(calculaChaveDeposito, X, 1)
    Else
        Base = (Mid(calculaChaveDeposito, X, 1) * 2)
        If Len(Base) = 2 Then
            Soma = Nnull(Mid(Base, 1, 1), 0)
            Soma = Soma + Nnull(Mid(Base, 2, 1), 0)
        Else
            Soma = Base
        End If
        mult = mult + Soma
    End If
Next X

dig = mult Mod 10
dig = 10 - dig

calculaChaveDeposito = Trim(calculaChaveDeposito & dig)
If Len(calculaChaveDeposito) <> 16 Then
    calculaChaveDeposito = ""
End If

Exit Function

trataErro:
    Err.Clear
End Function

Public Function obtemChaveDeposito(Lote As Long) As String
On Error GoTo trataErro:
Dim SQl As String
Dim rsaux As New ADODB.Recordset

obtemChaveDeposito = ""
'Exit Function
'NÃO ENTROU NO AR

sSql = "SELECT CH.AUTONUM, CH.CHAVE, GR.SEQ_GR FROM " & Banco_Sgipa & "TB_CHAVE_DEPOSITO CH "
sSql = sSql & " LEFT JOIN " & Banco_Sgipa & "TB_GR_BL GR ON CH.CHAVE = GR.CHAVE_DEPOSITO "
sSql = sSql & " WHERE ISNULL(CH.SEQ_GR,0) = 0 AND CH.LOTE =  " & Lote
PRSet rsaux, sSql
While Not rsaux.EOF
    If Nnull(rsaux!SEQ_GR, 0) = 0 Then
        obtemChaveDeposito = Nnull(rsaux!Chave, 1)
        Exit Function
    Else
        sSql = " UPDATE " & Banco_Sgipa & "TB_CHAVE_DEPOSITO SET SEQ_GR = " & Nnull(rsaux!SEQ_GR, 0)
        sSql = sSql & " WHERE AUTONUM = " & Nnull(rsaux!autonum, 0)
        db.Execute sSql
    End If
    rsaux.MoveNext
Wend


Exit Function

trataErro:
Err.Clear

End Function

Public Function obtemChaveDepositoGR(seqGR As Long) As String
On Error GoTo trataErro:
Dim SQl As String
Dim rsaux As New ADODB.Recordset

obtemChaveDepositoGR = ""
'Exit Function
'NÃO ENTROU NO AR

sSql = "SELECT GR.CHAVE_DEPOSITO FROM " & Banco_Sgipa & "TB_GR_BL GR  "
sSql = sSql & " WHERE GR.SEQ_GR = " & seqGR
PRSet rsaux, sSql
If Not rsaux.EOF Then
    obtemChaveDepositoGR = Nnull(rsaux!CHAVE_DEPOSITO, 1)
End If



Exit Function

trataErro:
Err.Clear

End Function

