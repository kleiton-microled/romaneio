VERSION 5.00
Begin VB.Form Principal 
   BackColor       =   &H00BA9F8D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema Coletor - Opções"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   3480
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Carregamento LTL"
      Height          =   420
      Index           =   11
      Left            =   180
      TabIndex        =   11
      Top             =   3600
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Manutenção -  Marcantes"
      Height          =   420
      Index           =   10
      Left            =   180
      TabIndex        =   10
      Top             =   1800
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Inventário Carga Solta"
      Height          =   420
      Index           =   9
      Left            =   180
      TabIndex        =   9
      Top             =   30
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Marcantes Carga Solta"
      Height          =   420
      Index           =   8
      Left            =   180
      TabIndex        =   8
      Top             =   1350
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Carregamento Carga Solta"
      Height          =   420
      Index           =   7
      Left            =   180
      TabIndex        =   7
      Top             =   3150
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Desova de Contêiner"
      Height          =   420
      Index           =   6
      Left            =   180
      TabIndex        =   6
      Top             =   900
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Movimentação Carga Solta"
      Height          =   420
      Index           =   5
      Left            =   180
      TabIndex        =   5
      Top             =   2700
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Movimentação Conteiner"
      Height          =   420
      Index           =   4
      Left            =   180
      TabIndex        =   4
      Top             =   2250
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Descarga CD"
      Enabled         =   0   'False
      Height          =   420
      Index           =   3
      Left            =   180
      TabIndex        =   3
      Top             =   450
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Sair"
      Height          =   450
      Index           =   2
      Left            =   165
      TabIndex        =   2
      Top             =   4080
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Estufagem"
      Enabled         =   0   'False
      Height          =   420
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   1350
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Descarga Armazem"
      Enabled         =   0   'False
      Height          =   420
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   30
      Width           =   3000
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

    Principal.Hide
    Select Case Index
    Case 0
        DD = False
        CD = True
        FrmTalieColetor.Show 1
        Principal.Show
    Case 1
        frm_carga_ova.Show 1
        Principal.Show
    Case 2
        Unload Me
    Case 3
        DD = True
        CD = False
        FrmTalieColetor.Show 1
        Principal.Show
    Case 4
        FrmMovimentacao.Show 1
        Principal.Show
        
    Case 5
        Inventario_CS.Show 1
        Principal.Show
    
    Case 6
        DesovaPier.Show 1
        Principal.Show
    Case 7
        CarregaCS.Show 1
        Principal.Show
        
    Case 8
        Marcantes.Show 1
        Principal.Show
        
    Case 9
        frmInventarioCego.Show 1
        Principal.Show
    
    Case 10
        AdminMarcante.Show 1
        Principal.Show
    
    Case 11
        CarregaLTL.Show 1
        Principal.Show
    
    End Select
    
    
End Sub

Private Sub Form_Activate()

Me.Top = 0
Me.Left = 0

If Sistema_Coletor = "REDEX" Then
    If Nnull(tbSenha!flag_rdx_DA, 0) = 0 Then Command1(0).Enabled = False
    If Nnull(tbSenha!flag_rdx_cd, 0) = 0 Then Command1(3).Enabled = False
    If Nnull(tbSenha!flag_rdx_estuf, 0) = 0 Then Command1(1).Enabled = False
    Command1(4).Visible = False
    Command1(5).Visible = False
    Command1(6).Visible = False
    Command1(7).Visible = False
    Command1(8).Visible = False
    Command1(9).Visible = False
    Command1(10).Visible = False
    
Else
    Sistema_Coletor = "IPA"
    Command1(0).Visible = False
    Command1(3).Visible = False
    Command1(1).Visible = False
    Command1(9).Visible = True
    Command1(10).Visible = True
    If CODEMSIS = "TEC" Then
        If Nnull(tbSenha!flag_ipa_mov, 0) = 0 Then Command1(4).Enabled = False: Command1(7).Enabled = False
    End If
End If



If Valida_Acesso_Botao(Col_Login, "COMMAND1(9)") = False Then Command1(9).Enabled = False
If Valida_Acesso_Botao(Col_Login, "COMMAND1(6)") = False Then Command1(6).Enabled = False
If Valida_Acesso_Botao(Col_Login, "COMMAND1(8)") = False Then Command1(8).Enabled = False
If Valida_Acesso_Botao(Col_Login, "COMMAND1(10)") = False Then Command1(10).Enabled = False
If Valida_Acesso_Botao(Col_Login, "COMMAND1(4)") = False Then Command1(4).Enabled = False
If Valida_Acesso_Botao(Col_Login, "COMMAND1(5)") = False Then Command1(5).Enabled = False
If Valida_Acesso_Botao(Col_Login, "COMMAND1(7)") = False Then Command1(7).Enabled = False



End Sub

Private Sub Form_Load()

Set frmLogin = Nothing

End Sub


Private Sub Form_Unload(Cancel As Integer)
UnloadAllForms
Unload Me
End Sub
