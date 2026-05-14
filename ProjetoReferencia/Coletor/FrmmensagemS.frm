VERSION 5.00
Begin VB.Form Frmmensagem 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   825
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   3435
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   780
      TabIndex        =   1
      Top             =   4020
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Atenção!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   390
      TabIndex        =   2
      Top             =   30
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Left            =   0
      TabIndex        =   0
      Top             =   330
      Width           =   3345
   End
End
Attribute VB_Name = "Frmmensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Me.Top = 0
Me.Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

