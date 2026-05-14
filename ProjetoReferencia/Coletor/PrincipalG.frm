VERSION 5.00
Begin VB.Form Principal 
   BackColor       =   &H00BA9F8D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema Coletor - Opções"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7005
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
   ScaleHeight     =   8190
   ScaleWidth      =   7005
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Consulta - Lib. Carregamento"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   9
      Left            =   570
      TabIndex        =   10
      Top             =   6400
      Width           =   6000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Localização - Carregamentos"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   8
      Left            =   570
      TabIndex        =   8
      Top             =   5600
      Width           =   6000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Contêiner - REEFER"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   7
      Left            =   570
      TabIndex        =   7
      Top             =   4800
      Width           =   6000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Desova de Contêiner"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   6
      Left            =   570
      TabIndex        =   6
      Top             =   2420
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Movimentação Carga Solta"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   5
      Left            =   570
      TabIndex        =   5
      Top             =   4000
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Movimentação Contêiner"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   4
      Left            =   570
      TabIndex        =   4
      Top             =   3210
      Width           =   6000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Descarga CD"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   3
      Left            =   570
      TabIndex        =   3
      Top             =   830
      Width           =   6000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   2
      Left            =   570
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7320
      Width           =   6000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Estufagem"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   1
      Left            =   570
      TabIndex        =   1
      Top             =   1620
      Width           =   6000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00BA9F8D&
      Caption         =   "Descarga Armazem"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   0
      Left            =   570
      TabIndex        =   0
      Top             =   30
      Width           =   6000
   End
   Begin VB.Label MNUEMPRESAS 
      Height          =   105
      Left            =   120
      TabIndex        =   9
      Top             =   8760
      Width           =   285
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
        Inventario_CS.Show 1
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
    Command1(7).Visible = True
    If Nnull(tbSenha!flag_reefer, 0) = 0 Then Command1(7).Enabled = False
    
    Command1(9).Visible = False
Else
    Sistema_Coletor = "IPA"
    Command1(0).Visible = False
    Command1(3).Visible = False
    Command1(1).Visible = False
    If CODEMSIS = "TEC" Then
        If Nnull(tbSenha!flag_ipa_mov, 0) = 0 Then Command1(4).Enabled = False
    End If
    Command1(7).Visible = True
    If Nnull(tbSenha!flag_reefer, 0) = 0 Then Command1(7).Enabled = False
    
    Command1(9).Visible = True
End If

Me.Caption = "SISTEMA COL. PATIO :" & Col_Patio



End Sub

Private Sub Form_Load()

Set frmLogin = Nothing

End Sub


