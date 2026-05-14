VERSION 5.00
Begin VB.Form frmPrimeiraSenha 
   BackColor       =   &H00C1EAF9&
   Caption         =   "CADASTRO DE SENHA"
   ClientHeight    =   2400
   ClientLeft      =   5805
   ClientTop       =   3930
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   3630
   Begin VB.TextBox txtSenha02 
      Appearance      =   0  'Flat
      DataSource      =   "txtSenha01"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   480
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1350
      Width           =   2625
   End
   Begin VB.TextBox txtSenha01 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   480
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   735
      Width           =   2625
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   450
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1800
      Picture         =   "PrimeiraSenha.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancelar"
      Top             =   1680
      Width           =   450
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C1EAF9&
      Caption         =   "Digite novamente a senha:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   6
      Top             =   1125
      Width           =   2820
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00C1EAF9&
      Caption         =   "Por favor, informe abaixo a nova senha:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   510
      Width           =   3060
   End
   Begin VB.Label lblTITULO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cadastro de Senha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C1EAF9&
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4005
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C1EAF9&
      BackStyle       =   1  'Opaque
      Height          =   1935
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   390
      Width           =   3315
   End
End
Attribute VB_Name = "frmPrimeiraSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mblnCancelou As Boolean
Public mblnAtualizouSenha As Boolean
Public mcnConexao As ADODB.Connection
Public mstrTabela As String
Public mstrCampoUsu As String
Public mstrUsuario As String
Public mstrCampoSenha As String
Public mstrSenha As String
Public mstrStringProibida As String

Private Sub CmdCancelar_Click()
  mblnCancelou = True
  Unload Me
End Sub

Private Sub CmdOK_Click()
  Dim strSQL As String
  Dim intRecAff As Integer
  Dim i As Integer
  
  If mcnConexao.State = 0 Then
    mcnConexao.Open
  Else
    mcnConexao.Close
    mcnConexao.Open
  End If
  
  If Not ValidarCampos Then Exit Sub
  
  mstrSenha = txtSenha01.text
  
  strSQL = "UPDATE"
  strSQL = strSQL & " " & mstrTabela
  strSQL = strSQL & " SET"
  strSQL = strSQL & " " & mstrCampoSenha
  strSQL = strSQL & " = '" & UCase(mstrSenha) & "'"
  strSQL = strSQL & " WHERE"
  strSQL = strSQL & " " & mstrCampoUsu
  strSQL = strSQL & " = '" & UCase(mstrUsuario) & "'"
  i = 0
  For i = 0 To 10
    mcnConexao.Execute strSQL, intRecAff
    If intRecAff > 0 Then Exit For
  Next i
  
  If intRecAff > 0 Then
    mblnAtualizouSenha = True
    mstrSenha = txtSenha01.text
    Me.Hide
  Else
    mblnAtualizouSenha = False
    mstrSenha = Empty
    Me.Hide
  End If
End Sub

Private Sub Form_Load()
  mblnCancelou = False
  mblnAtualizouSenha = False
  mstrTabela = Empty
  mstrCampoUsu = Empty
  mstrUsuario = Empty
  mstrCampoSenha = Empty
  mstrSenha = Empty
  mstrStringProibida = "TECOND"
End Sub

Private Function ValidarCampos() As Boolean
  Dim strSenha01 As String
  Dim strSenha02 As String
  Dim strMsg As String
  
  ValidarCampos = False
  
  strSenha01 = txtSenha01.text
  strSenha02 = txtSenha02.text
  
  If strSenha01 = Empty Then
    MsgBox "Por favor, digite a senha."
    txtSenha01.SetFocus
    Marca Me
    Exit Function
  ElseIf strSenha02 = Empty Then
    MsgBox "Por favor, confirme a senha."
    txtSenha02.SetFocus
    Marca Me
    Exit Function
  ElseIf strSenha01 <> strSenha02 Then
    MsgBox "As senha digitadas não coincidem. Por favor, digite novamente."
    txtSenha01.SetFocus
    Marca Me
    Exit Function
  ElseIf InStr(1, strSenha01, mstrStringProibida, vbTextCompare) Then
    strMsg = "A senha não pode conter a seguinte seqüência de caracteres:"
    strMsg = strMsg & vbCr & mstrStringProibida & vbCr
    strMsg = strMsg & "Por favor, digite novamente."
    MsgBox strMsg
    txtSenha01.SetFocus
    Marca Me
    Exit Function
  End If
  
  ValidarCampos = True
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

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub txtSenha01_GotFocus()
  Marca Me
End Sub

Private Sub txtSenha01_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub txtSenha02_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub txtSenha02_GotFocus()
  Marca Me
End Sub

