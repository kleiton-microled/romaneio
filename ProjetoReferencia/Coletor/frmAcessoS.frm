VERSION 5.00
Begin VB.Form Senha_E 
   BackColor       =   &H00A26842&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4956
   ClientLeft      =   3780
   ClientTop       =   3420
   ClientWidth     =   3276
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAcessoS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4956
   ScaleWidth      =   3276
   Begin VB.PictureBox pBARRA 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00BA9F8D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   492
      ScaleWidth      =   3276
      TabIndex        =   16
      Top             =   0
      Width           =   3270
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "TESTE DE TITULO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E3F7FF&
         Height          =   300
         Index           =   5
         Left            =   60
         TabIndex        =   17
         Top             =   90
         Width           =   6300
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         Picture         =   "frmAcessoS.frx":0442
         Stretch         =   -1  'True
         Top             =   90
         Width           =   300
      End
   End
   Begin VB.PictureBox pQUADRO 
      Appearance      =   0  'Flat
      BackColor       =   &H00BA9F8D&
      ForeColor       =   &H80000008&
      Height          =   8010
      Left            =   0
      ScaleHeight     =   7992
      ScaleWidth      =   3204
      TabIndex        =   8
      Top             =   510
      Width           =   3225
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "BASE DE TESTES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   180
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   456
         Width           =   1860
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         ItemData        =   "frmAcessoS.frx":210C
         Left            =   1380
         List            =   "frmAcessoS.frx":210E
         TabIndex        =   3
         Top             =   1035
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.PictureBox pACESSA 
         Appearance      =   0  'Flat
         BackColor       =   &H00BA9F8D&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1605
         Left            =   0
         ScaleHeight     =   1608
         ScaleWidth      =   3132
         TabIndex        =   13
         Top             =   1590
         Width           =   3135
         Begin VB.TextBox txtFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   30
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   1
            Top             =   1200
            Width           =   3120
         End
         Begin VB.TextBox txtFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   30
            MaxLength       =   50
            TabIndex        =   0
            Top             =   420
            Width           =   3120
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "SENHA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   390
            Index           =   1
            Left            =   30
            TabIndex        =   15
            Top             =   810
            Width           =   3120
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "USUÁRIO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   45
            TabIndex        =   14
            Top             =   60
            Width           =   3090
         End
      End
      Begin VB.CommandButton cmd 
         Appearance      =   0  'Flat
         BackColor       =   &H00BA9F8D&
         Caption         =   "ACESSAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "0"
         Top             =   3570
         Width           =   1575
      End
      Begin VB.CommandButton cmd 
         Appearance      =   0  'Flat
         BackColor       =   &H00BA9F8D&
         Cancel          =   -1  'True
         Caption         =   "CANCELAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   2
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3570
         Width           =   1515
      End
      Begin VB.CommandButton cmd 
         Appearance      =   0  'Flat
         BackColor       =   &H00BA9F8D&
         Caption         =   "ALTERAR SENHA"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "2"
         Top             =   7830
         Visible         =   0   'False
         Width           =   3285
      End
      Begin VB.PictureBox pTROCA 
         Appearance      =   0  'Flat
         BackColor       =   &H00BA9F8D&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1050
         Left            =   -660
         ScaleHeight     =   1056
         ScaleWidth      =   3396
         TabIndex        =   10
         Top             =   7860
         Width           =   3390
         Begin VB.TextBox txtFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   45
            MaxLength       =   8
            PasswordChar    =   "@"
            TabIndex        =   6
            Top             =   1725
            Width           =   6600
         End
         Begin VB.TextBox txtFields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   30
            MaxLength       =   8
            PasswordChar    =   "@"
            TabIndex        =   7
            Top             =   1920
            Width           =   6615
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CONFIRME A SENHA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   360
            Index           =   2
            Left            =   45
            TabIndex        =   12
            Top             =   1545
            Width           =   6600
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "NOVA SENHA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   435
            Index           =   3
            Left            =   45
            TabIndex        =   11
            Top             =   1275
            Width           =   6600
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "microLED"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A26842&
         Height          =   240
         Index           =   3
         Left            =   2025
         TabIndex        =   19
         Top             =   210
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PÁTIO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   150
         TabIndex        =   18
         Top             =   1020
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   1515
         Left            =   30
         Top             =   60
         Width           =   3090
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   -15
      X2              =   4965
      Y1              =   375
      Y2              =   375
   End
End
Attribute VB_Name = "Senha_E"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public lAcessou         As Boolean   'Identifica se o log do usuário foi bem sucedido.
Public lShowPatio       As Boolean   'Determina se deve permitir ao usuário escolher o pátio ou não.
Public lShowServer      As Boolean   'Determina se deve permitir ao usuário indicar o serviço de dados.

Private Sub Command1_Click()
  
End Sub

Private Sub Form_Activate()
Dim Teste As Integer
Dim Param As String

Me.Left = 0
Me.Top = 0


'TEMPORARIO
'á despachantes
If App.Title = "Atendimento á Despachantes" Then
    txtFields(0).text = "DESPACHANTE"
    txtFields(1).text = "DESP"
     Call cmd_Click(0)
Else
    txtFields(0).text = ""
    txtFields(1).text = ""
End If
'FIM

If Trim(Command$) <> "" Then
  Param = Trim(Command$)
  txtFields(0) = Trim(Mid$(Param, 1, InStr(Param, ",") - 1))
  Param = Replace(Param, txtFields(0) & ",", "")
  txtFields(1) = Trim(Mid$(Param, 1, InStr(Param, ",") - 1))
  Teste = Trim(Mid$(Param, InStr(Param, ",") + 1))
  If Teste = 1 Then Check1.Value = 1
  Me.Visible = False
  Call cmd_Click(0)
End If


End Sub

Private Sub Form_Load()


On Error Resume Next
If Dir("c:\microledtemp", vbDirectory) = "" Then
    MkDir "c:\microledtemp"
End If
On Error GoTo 0

If Dir("c:\microledtemp", vbDirectory) = "" Then
    MsgBox "Atenção, não foi encontrada a pasta temporária na máquina, contate o suporte de TI", vbCritical: Exit Sub
End If


BD_O = "operador"
BD_I = "sgipa"

DoEvents
'Valores Default
sEsquema = "SGIPA"
nPATIO = 1
wModelo = "9600,n,8,1"
PreparaAcesso
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Sendkeys "{TAB}"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Image1.BorderStyle = 0 Then Image1.BorderStyle = 0
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 1
End Sub

Private Sub pBARRA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Image1.BorderStyle = 0 Then Image1.BorderStyle = 0
End Sub

Private Sub txtFields_Change(Index As Integer)
If Not cmd(1).Caption = "SALVAR" Then
    cmd(1).Enabled = (Len(Trim$(txtFields(0).text)) > 0 And Len(Trim$(txtFields(1).text)) > 0)
Else
    cmd(1).Enabled = ((Len(Trim$(txtFields(2).text)) > 0 And Len(Trim$(txtFields(3).text)) > 0) And (txtFields(2).text = txtFields(3).text))
End If
End Sub

Private Sub txtfields_GotFocus(Index As Integer)
txtFields(Index).BackColor = lbl(5).ForeColor
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
If Not (KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete) Then KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtfields_LostFocus(Index As Integer)
  txtFields(Index).BackColor = vbWhite
  

End Sub

Private Sub cmd_Click(Index As Integer)



Dim Rs2 As New ADODB.Recordset

Col_Patio = 1
 ServidorOperador = ""

'If Check1.Value = 1 Then
 '   ServidorOperador = "PARADISE."
'Else
 '   ServidorOperador = "MATIRA."
'End If



If Index = 0 Then
  ControlarPrimeiroAcesso txtFields(0).text, _
                          "" & Banco_Sgipa & "TB_CAD_USUARIOS", _
                          "USUARIO", _
                          "SENHA", _
                          txtFields(1)
    db.Close
    Set db = Nothing
    lAcessou = Acessa
ElseIf Index = 1 Then
    TrocaSenha
ElseIf Index = 2 Then
    Unload Me
End If

If lAcessou = True Then
        
        Sql = " INSERT INTO " & Banco_Sgipa & "TB_SYS_USER_GRUPOS (AUTONUMUSER,CODGRUPO)   "
        Sql = Sql & "SELECT "
        Sql = Sql & "U.AUTONUM,U.CODGRUPO "
        Sql = Sql & " FROM "
        Sql = Sql & " " & Banco_Sgipa & "TB_CAD_USUARIOS U "
        Sql = Sql & " left join  " & Banco_Sgipa & "TB_SYS_USER_GRUPOS G on U.AUTONUM = G.AUTONUMUSER"
        Sql = Sql & " WHERE "
        Sql = Sql & " G.AUTONUMUSER IS NULL "
          
        db.Execute Sql  'z4

    If CompletaAcesso = True Then
        If Val(Cod_Empresa) = 0 Then FrmEscolhaEmpresa.Show 1
        
   
        Sql = "INSERT INTO " & Banco_Sgipa & "LOG_ACESSO ("
        Sql = Sql & "id, USUARIO,DATA_INICIO, MAQUINA, MODULO) VALUES ("
        Sql = Sql & "seq_log_acesso_ipa.nextval,'" & Usuario_Sistema & "', SYSDATE ,'" & Maquina_Rede & "','" & App.ProductName & "')"
     '   executa sql  'z4
 
        Executa Sql
   
        Unload Me
    End If
End If
End Sub

Public Sub PreparaAcesso()
Dim t As Single
lbl(5).Caption = UCase$(App.Title & " v." & App.Major & "." & Format$(App.Minor, "00") & "." & Format$(App.Revision, "0000") & App.Comments)
If lShowPatio = True Then
    With List1
        For t = 1 To 10
            .AddItem t
        Next
        .ListIndex = 0
        lbl(4).Visible = True
        .Visible = True
    End With
End If
'Check1.Visible = (lShowServer = True)
t = Empty
End Sub

Private Function Acessa() As Boolean
Dim lRET As Boolean
Dim Tbpatio As ADODB.Recordset
Dim wFinal  As String
Dim R As ADODB.Recordset

If Not db Is Nothing Then
    If Not db.State = 0 Then db.Close
    Set db = Nothing
    Set db = New ADODB.Connection
End If
DBConnect db, , , "SGIPA", (Check1.Value = 1)


If Not db Is Nothing Then

lRET = False
   Sql = "SELECT * FROM " & Banco_Sgipa & "TB_CAD_USUARIOS WHERE USUARIO = '" & txtFields(0).text & "' and flag_ativo=1"
  ' MsgBox Sql
   PRSet tbSenha, "SELECT * FROM " & Banco_Sgipa & "TB_CAD_USUARIOS WHERE USUARIO = '" & txtFields(0).text & "' and flag_ativo=1"
   If Not tbSenha.EOF Then
        If Not UCase(tbSenha.Fields("SENHA")) = UCase(txtFields(1).text) Then
            MsgBox "Senha incorreta.", vbOKOnly + vbInformation, App.Title
            txtFields(1).SetFocus
        Else
            Patio_coletor = tbSenha.Fields("patio")
            Col_Patio = Patio_coletor
            If Patio_coletor = 0 Then Patio_coletor = 1
            
            If Patio_coletor = 1 Then
                Col_Tam_Pos_Yard = 6
            Else
                Col_Tam_Pos_Yard = 7
            End If
            
            Cod_Usuario = tbSenha.Fields("autonum")
            Cod_Empresa = tbSenha.Fields("Cod_Empresa")
            wCpf_Usuario = Nnull(tbSenha!CPF, 1)
            Wnome_Usuario = Nnull(tbSenha!Nome, 1)
            PRSet Tbpatio, "SELECT  Tabela_Padrao ,razao_social ,Flag_ferrovia FROM " & Banco_Sgipa & "TB_Empresas  WHERE Autonum=" & Cod_Empresa & " OR 0=" & Cod_Empresa
            If Not Tbpatio.EOF Then
               Tabela_Padrao = Nnull(Tbpatio!Tabela_Padrao, 0)
               Razao_Empresa = Nnull(Tbpatio!razao_social, 1)
               Tem_ferrovia = (Nnull(Tbpatio!Flag_ferrovia, 0) = 1)
            End If
            PRSet Tbpatio, "SELECT  * FROM " & Banco_Operador & "TB_PATIOS  WHERE recinto<>'000' and (COD_EMPRESA=" & Cod_Empresa & " OR 0=" & Cod_Empresa & ")"
            If Not Tbpatio.EOF Then
               Cod_Empresa_Patio = "PATIO IN("
               Do
                 p_in = p_in & Tbpatio!autonum
                 Cod_Empresa_Patio = Cod_Empresa_Patio & Tbpatio!autonum
                 Tbpatio.MoveNext
                 If Not Tbpatio.EOF Then
                   p_in = p_in & ","
                   Cod_Empresa_Patio = Cod_Empresa_Patio & ","
                 Else
                   Cod_Empresa_Patio = Cod_Empresa_Patio & ")"
                 End If
               Loop Until Tbpatio.EOF
            End If
            lRET = True
        End If
    Else
        MsgBox "Usuário não encontrado.", vbOKOnly + vbInformation, App.Title
    End If
End If
If lRET = False Then tbSenha.Close

Acessa = lRET
lRET = Empty
End Function

Private Sub TrocaSenha()
Static lTroca As Boolean
Dim nRecs     As Long
If MsgBox("Deseja alterar a senha de acesso agora?", vbYesNo + vbQuestion, App.Title) = vbNo Then Exit Sub
lTroca = Not lTroca
If lTroca = True Then
  DBConnect db, , , "SGIPA", (Check1.Value = 1)
  If Not db Is Nothing Then
     If PRSet(tbSenha, "SELECT * FROM " & Banco_Sgipa & "TB_CAD_USUARIOS WHERE USUARIO = '" & txtFields(0).text & "' and flag_ativo=1 ") = True Then
        If Not tbSenha.RecordCount = 0 Then
            If Not UCase(tbSenha.Fields("SENHA")) = UCase(txtFields(1).text) Then
                MsgBox "Senha incorreta.", vbOKOnly + vbInformation, App.Title
                txtFields(1).SetFocus
                Exit Sub
            End If
        End If
     End If
  Else
   Exit Sub
  End If
  Me.pACESSA.Visible = False
  Me.pTROCA.Visible = True
  cmd(1).Caption = "SALVAR"
  cmd(1).Enabled = False
  txtFields(2).SetFocus
Else
    If Not db Is Nothing Then
        If Not db.State = 0 Then db.Close
        Set db = Nothing
    End If
    Set db = New ADODB.Connection
    DBConnect db, , , "SGIPA", (Check1.Value = 1)
    'Executa "UPDATE " & banco_sgipa & "TB_CAD_USUARIOS SET SENHA = '" & UCase(txtfields(2).Text) & "' WHERE USUARIO = '" & UCase(txtfields(0).Text) & "' AND SENHA = '" & UCase(txtfields(1).Text) & "' ", nRecs, True
    db.Execute "UPDATE " & Banco_Sgipa & "TB_CAD_USUARIOS SET SENHA = '" & UCase(txtFields(2).text) & "' WHERE USUARIO = '" & UCase(txtFields(0).text) & "' AND SENHA = '" & UCase(txtFields(1).text) & "' ", nRecs
    If nRecs > 0 Then
        txtFields(1).text = txtFields(2).text
    Else
        MsgBox "Verifique o nome de usuário e a senha de acesso.", vbOKOnly + vbInformation, App.Title
    End If
    Me.pACESSA.Visible = True
    Me.pTROCA.Visible = False
    cmd(1).Caption = "ALTERAR SENHA"
    cmd(1).Enabled = (Len(Trim$(txtFields(0).text)) > 0 And Len(Trim$(txtFields(1).text)) > 0)
    txtFields(0).SetFocus
End If
nRecs = Empty
End Sub

Private Function CompletaAcesso() As Boolean
On Error GoTo CAErr:
Dim nArq As Byte
Dim sArq As String
Dim sLin As String
Dim nLin As Integer
If lShowPatio = True Then
    nLin = 0
    sArq = "C:\gate.ini"
    If Dir(sArq) <> Empty Then
        nArq = FreeFile
        'Conta as linhas
        Open sArq For Input As #nArq
        While Not EOF(nArq) = True
            Line Input #nArq, sLin: nLin = nLin + 1
        Wend
        Close #nArq
        If nLin = 2 Then
            nArq = FreeFile
            Open sArq For Input As #nArq
            If Not EOF(nArq) = True Then
                Line Input #nArq, sLin
                Line Input #nArq, wModelo
            End If
            Close #nArq
            If Not sLin = Empty Then
                If Not sLin = List1.text Then
                    lAcessou = False
                    MsgBox "Pátio solicitado difere do permitido.", vbOKOnly + vbInformation, App.Title
                End If
            Else
                lAcessou = False
                MsgBox "Arquivo de configuração está corrompido.", vbOKOnly + vbInformation, App.Title
            End If
        Else
            lAcessou = False
            MsgBox "Arquivo de configuração está corrompido.", vbOKOnly + vbInformation, App.Title
        End If
    Else
        lAcessou = False
        MsgBox "Arquivo de configuração não encontrado.", vbOKOnly + vbInformation, App.Title
    End If
End If
If lAcessou = True Then
    If Not List1.ListIndex < 0 Then
        nPATIO = List1.List(List1.ListIndex)
    End If
    Maquina_Rede = QualMaquina(255)
    Usuario_Rede = QualUsuario(255)
    Usuario_Sistema = UCase(tbSenha.Fields("USUARIO"))
    Cod_Usuario = tbSenha.Fields("AUTONUM")
    Cod_Empresa = tbSenha.Fields("Cod_EMPRESA")
    GateManual = Nnull(tbSenha.Fields("FLAG_BALANCA_MANUAL"), 0)
    CompletaAcesso = True
End If
GoTo CAExit:
CAErr:
MsgBox Err.Description, vbOKOnly + vbCritical, App.Title
Err.Clear
CAExit:
nArq = Empty
sArq = Empty
nLin = Empty
sLin = Empty
End Function

Private Sub ControlarPrimeiroAcesso(usuario As String, _
                                    Tabela As String, _
                                    CampoUsuario As String, _
                                    CampoSenha As String, _
                                    ControleSenha As TextBox)
  Dim intResposta As Integer
  
  On Error GoTo ErroSenha
  
  
  
  
  'If UCase(App.EXEName) = "LEITURADTE" Or UCase(App.EXEName) = "LEITURA" Or UCase(App.EXEName) = "LEITURANEW" Then
  '    DBConnect DbTeste, , , , True
  'End If
  
  DBConnect db, , , , CBool(Check1.Value)
  
  
  If usuario <> Empty Then
    If Not UsuarioTemSenha(db, _
                           Tabela, _
                           CampoUsuario, _
                           CampoSenha, _
                           usuario) Then
      Load frmPrimeiraSenha
      
      With frmPrimeiraSenha
        Set .mcnConexao = db
        .mstrTabela = Tabela
        .mstrCampoUsu = CampoUsuario
        .mstrUsuario = usuario
        .mstrCampoSenha = CampoSenha
        
        intResposta = vbRetry
        While intResposta = vbRetry
          frmPrimeiraSenha.Show 1
          
          If Not .mblnAtualizouSenha Then
            If .mblnCancelou Then
              intResposta = MsgBox("É necessário cadastrar uma senha para continuar.", _
                                    vbRetryCancel, "Cadastro de Senha")
              If intResposta = vbCancel Then End
            End If
          Else
            intResposta = vbOK
            ControleSenha.text = .mstrSenha
          End If
        Wend
      End With
    End If
  End If
  
  Exit Sub
  
ErroSenha:
  If Err.Number = 1010 Then
    Exit Sub
  End If
End Sub
