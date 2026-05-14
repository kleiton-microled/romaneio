VERSION 5.00
Begin VB.Form Principal 
   BackColor       =   &H00BA9F8D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema Coletor - Opções"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3240
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
   ScaleWidth      =   3240
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Descarga DDC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   15
      Left            =   135
      TabIndex        =   15
      Top             =   3795
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Carregamento Carga Solta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   14
      Left            =   135
      TabIndex        =   14
      Top             =   1020
      Width           =   3030
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Marcantes Carga Solta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   13
      Left            =   135
      TabIndex        =   13
      Top             =   555
      Width           =   3030
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Movimentação Veículos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   12
      Left            =   135
      TabIndex        =   12
      Top             =   1515
      Width           =   3030
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Movimentação Carga Solta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   11
      Left            =   1665
      TabIndex        =   11
      Top             =   45
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Veiculos - Carregamento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   5
      Left            =   135
      TabIndex        =   10
      Top             =   3375
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Contêiner - REEFER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   7
      Left            =   135
      TabIndex        =   9
      Top             =   1995
      Width           =   3030
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Manutenção -  Marcantes"
      Height          =   420
      Index           =   10
      Left            =   3510
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Consulta - Lib. Carregamento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   9
      Left            =   135
      TabIndex        =   7
      Top             =   2910
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Localizacao Carregamentos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   8
      Left            =   135
      TabIndex        =   6
      Top             =   2445
      Width           =   3030
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Desova de Contêiner"
      Height          =   420
      Index           =   6
      Left            =   3540
      TabIndex        =   5
      Top             =   30
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Movimentação Conteiner"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   4
      Left            =   135
      TabIndex        =   0
      Top             =   45
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Descarga CD"
      Enabled         =   0   'False
      Height          =   420
      Index           =   3
      Left            =   3570
      TabIndex        =   4
      Top             =   30
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   135
      TabIndex        =   3
      Top             =   4215
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Estufagem"
      Enabled         =   0   'False
      Height          =   420
      Index           =   1
      Left            =   3540
      TabIndex        =   2
      Top             =   1350
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Descarga Armazem"
      Enabled         =   0   'False
      Height          =   420
      Index           =   0
      Left            =   3630
      TabIndex        =   1
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
        End
    Case 3
        DD = True
        CD = False
        FrmTalieColetor.Show 1
        Principal.Show
    Case 4
        FrmMovimentacao.Show 1
        Principal.Show
        
    Case 5
        FrmCarregaVeiculo.Show 1
        Principal.Show
    
    Case 6
        DesovaPier.Show 1
        Principal.Show
       
   Case 7
        Reefer.Show 1
        Principal.Show
        
    Case 8
        Localizacao.Show 1
         Principal.Show
    Case 9
        frmLibRegistro.Show 1
        Principal.Show
        
    Case 11
        Inventario_CSPatio.Show 1
        Principal.Show
        
    Case 12
        MovChassis.Show 1
        Principal.Show
        
    Case 13
        MovimentaCS.Show 1
        Principal.Show
 
    Case 14
        CarregaCSPatio.Show 1
        Principal.Show
    Case 15
        FrmDescargaArmazem.Show 1
        Principal.Show
    End Select
    
    
End Sub

Private Sub Command2_Click()

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
Else
    Sistema_Coletor = "IPA"
    Command1(0).Visible = False
    Command1(3).Visible = False
    Command1(1).Visible = False
    If CODEMSIS = "TEC" Then
        If Nnull(tbSenha!flag_ipa_mov, 0) = 0 Then Command1(4).Enabled = False
    End If
End If



End Sub

Private Sub Form_Load()

Set frmLogin = Nothing

End Sub


