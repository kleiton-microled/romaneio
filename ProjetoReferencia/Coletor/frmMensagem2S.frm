VERSION 5.00
Begin VB.Form frmMensagem2 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3570
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "PROSSEGUIR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   2
      Top             =   3900
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CANCELAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1920
      TabIndex        =   1
      Top             =   3900
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   390
      Picture         =   "frmMensagem2S.frx":0000
      Top             =   150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageInformation 
      Height          =   480
      Left            =   330
      Picture         =   "frmMensagem2S.frx":030A
      Top             =   180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImageExclamation 
      Height          =   480
      Left            =   450
      Picture         =   "frmMensagem2S.frx":0614
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   2640
      Top             =   150
      Width           =   765
   End
   Begin VB.Label lblMensagem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   270
      TabIndex        =   0
      Top             =   1260
      Width           =   3045
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00BA9F8D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   4755
      Left            =   0
      Top             =   0
      Width           =   3525
   End
End
Attribute VB_Name = "frmMensagem2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Col_Cancelado = True
Unload Me
End Sub

Private Sub Command2_Click()
Col_Cancelado = False
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = vbKeyEscape And Me.Command1.Visible = False Then
    Col_Mensagem = ""
    Unload Me
End If


End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.lblMensagem = Col_Mensagem
If Col_Imagem = 1 Then Me.Image1 = Me.ImageExclamation
If Col_Imagem = 2 Then Me.Image1 = Me.ImageInformation
If Col_Imagem = 3 Then Me.Image1 = Me.Image2: Me.Command1.Visible = True: Me.Command2.Visible = True
Col_Cancelado = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

