VERSION 5.00
Begin VB.Form MovChassis 
   Appearance      =   0  'Flat
   BackColor       =   &H00BA9F8D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimentação"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   3345
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPerfil 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   810
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2670
      Width           =   2535
   End
   Begin VB.TextBox txtAutonum 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3420
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txtYardPos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2070
      TabIndex        =   24
      Top             =   2940
      Width           =   1275
   End
   Begin VB.CommandButton cmdLimpa 
      Caption         =   "[F2]LIMPA"
      Height          =   375
      Left            =   90
      TabIndex        =   23
      Top             =   3990
      Width           =   1065
   End
   Begin VB.CommandButton cmdSalva 
      Caption         =   "[F5]SALVA"
      Height          =   375
      Left            =   1200
      TabIndex        =   22
      Top             =   3990
      Width           =   1065
   End
   Begin VB.CommandButton cmdEscape 
      Caption         =   "[ESC]SAIR"
      Height          =   375
      Left            =   2340
      TabIndex        =   21
      Top             =   3990
      Width           =   975
   End
   Begin VB.TextBox txtEntrada 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   810
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox txtCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   810
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2130
      Width           =   2535
   End
   Begin VB.TextBox txtYard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   810
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2940
      Width           =   1275
   End
   Begin VB.TextBox txtNavio 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   810
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1860
      Width           =   2535
   End
   Begin VB.TextBox txtMarca 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   810
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1590
      Width           =   2535
   End
   Begin VB.TextBox txtModelo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   810
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txtIDCarga 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   810
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1050
      Width           =   2535
   End
   Begin VB.TextBox txtLote 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   810
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   780
      Width           =   2535
   End
   Begin VB.TextBox txtBL 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   810
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   510
      Width           =   2535
   End
   Begin VB.TextBox txtChassis 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   810
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PERFIL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   27
      Top             =   2700
      Width           =   570
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ENTRADA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   19
      Top             =   2400
      Width           =   780
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "YARD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   300
      TabIndex        =   16
      Top             =   3000
      Width           =   465
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NAVIO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   1860
      Width           =   525
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MARCA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   11
      Top             =   1620
      Width           =   615
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MODELO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   10
      Top             =   1350
      Width           =   690
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2130
      Width           =   660
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ID CARGA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   825
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LOTE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   810
      Width           =   405
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   570
      TabIndex        =   3
      Top             =   510
      Width           =   195
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CHASSIS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   30
      TabIndex        =   1
      Top             =   270
      Width           =   825
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MOV. CHASSIS"
      ForeColor       =   &H00BA9F8D&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3345
   End
End
Attribute VB_Name = "MovChassis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub cmdEscape_Click()
    Unload Me
End Sub

Sub cmdLimpa_Click()
Limpa
Me.txtChassis.Text = ""
Me.txtChassis.SetFocus
End Sub

Sub cmdSalva_Click()
If Valida Then

    Dim SQl As String
    If Left(Me.txtPerfil.Text, 3) = "IPA" Then
        SQl = "UPDATE TB_CARGA_SOLTA_CHASSIS "
        SQl = SQl & " SET YARD='" & Me.txtYardPos & "'"
        SQl = SQl & " WHERE CS=" & Me.txtAutonum.Text
        SQl = SQl & " AND CHASSIS='" & Me.txtChassis.Text & "'"
        db.Execute SQl
        Col_Mensagem = "CHASSIS POS. COM SUCESSO"
        Col_Imagem = 2
        Load frmMensagem2
        frmMensagem2.Command1.Visible = False
        frmMensagem2.Show 1
    Else
        SQl = "UPDATE " & ServidorOperador
        SQl = SQl & "OPERADOR.DBO.TB_CS_BOOKING_ITEM_ID "
        SQl = SQl & " SET YARD='" & Me.txtYardPos.Text & "'"
        SQl = SQl & " ,DT_YARD=GETDATE() "
        SQl = SQl & " WHERE AUTONUM_PATIO=" & Me.txtAutonum.Text
        SQl = SQl & " AND CHASSIS='" & Me.txtChassis.Text & "'"
        db.Execute SQl
        Col_Mensagem = "CHASSIS POS. COM SUCESSO"
        Col_Imagem = 2
        Load frmMensagem2
        frmMensagem2.Command1.Visible = False
        frmMensagem2.Show 1
    End If

    Limpa
    Me.txtChassis.Text = ""
    Me.txtChassis.SetFocus

End If

End Sub

Private Function Valida() As Boolean
Valida = True
If Me.txtAutonum.Text = "" Or Me.txtAutonum.Text = "0" Then
    Col_Mensagem = "Chassis não encontrado"
    frmMensagem2.Show vbModal
    Valida = False
    Exit Function
End If
 
If Me.txtYard.Text = Me.txtYardPos.Text Then
    Col_Mensagem = "Chassis já se encontra nesta posicao"
    frmMensagem2.Show vbModal
    Valida = False
    Exit Function
End If
    
Dim SQl As String
SQl = "Select autonum from operador..tb_yard "
'Sql = Sql & " where patio=" & IIf(Col_Patio = 7, 1, Col_Patio)
SQl = SQl & " where PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio) & ""
SQl = SQl & " and yard='" & Me.txtYardPos.Text & "'"
Set tby = New ADODB.Recordset
tby.Open SQl, db, 1, 1
If tby.EOF Then
    tby.Close
    Col_Mensagem = "Posição inválida"
    frmMensagem2.Show vbModal
    Valida = False
    Exit Function
End If
tby.Close


End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Call Me.cmdEscape_Click
End If

If KeyCode = vbKeyF2 Then
    Call Me.cmdLimpa_Click
End If

End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0

End Sub

Private Sub txtChassis_LostFocus()
Limpa
Dim SQl As String

SQl = " SELECT [ID_CARGA]"
SQl = SQl & " ,[CHASSIS]"
SQl = SQl & " ,[BL]"
SQl = SQl & " ,[LOTE]"
SQl = SQl & " ,[MODELO]"
SQl = SQl & " ,[MARCA]"
SQl = SQl & " ,[NAVIO]"
SQl = SQl & " ,[VIAGEM]"
SQl = SQl & " ,[PATIO]"
SQl = SQl & " ,[YARD]"
SQl = SQl & " ,[DT_GATE_IN]"
SQl = SQl & " ,[DT_SAIDA]"
SQl = SQl & " ,[FANTASIA]"
SQl = SQl & " ,[SISTEMA]"
SQl = SQl & " ,[FINALITY]"
SQl = SQl & " ,[LOTE_STR]"
SQl = SQl & " ,[AUTONUMCS]"
SQl = SQl & " ,[DIAS_ARM_FAIXA]"
SQl = SQl & " From [dbo].[VW_INVENT_CHASSIS]"
SQl = SQl & " where PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio) & ""
SQl = SQl & " and chassis='" & Me.txtChassis.Text & "'"
Set tb1 = New ADODB.Recordset
tb1.Open SQl, db, 1, 1
If Not tb1.EOF Then
    
    Me.txtBL.Text = Nnull(tb1!BL, 1)
    Me.txtCliente.Text = Nnull(tb1!Fantasia, 1)
    Me.txtEntrada.Text = Format(tb1!Dt_Gate_In, "dd/mm/yyyy hh:mm")
    Me.txtIDCarga.Text = Nnull(tb1!id_carga, 0)
    Me.txtLote.Text = Nnull(tb1!lote_str, 1)
    Me.txtMarca.Text = Nnull(tb1!Marca, 1)
    Me.txtModelo.Text = Nnull(tb1!Modelo, 1)
    Me.txtYard.Text = Nnull(tb1!Yard, 1)
    Me.txtNavio.Text = Nnull(tb1!Navio, 1) & " " & Nnull(tb1!Viagem, 1)
    Me.txtPerfil.Text = Nnull(tb1!finality, 1)
    Me.txtAutonum.Text = Nnull(tb1!AutonumCs, 0)
End If
tb1.Close

End Sub


Private Sub Limpa()
Me.txtBL.Text = ""
Me.txtCliente.Text = ""
Me.txtEntrada.Text = ""
Me.txtIDCarga.Text = ""
Me.txtLote.Text = ""
Me.txtMarca.Text = ""
Me.txtModelo.Text = ""
Me.txtNavio.Text = ""
Me.txtYard.Text = ""
Me.txtYardPos.Text = ""
Me.txtAutonum.Text = ""
Me.txtPerfil.Text = ""

End Sub

Private Sub txtYardPos_LostFocus()
txtYardPos = UCase(txtYardPos)
End Sub
