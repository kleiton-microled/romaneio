VERSION 5.00
Begin VB.Form Frmmensagem 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   825
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   7005
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   1
      Top             =   7050
      Width           =   2505
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Atenção!!!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1710
      TabIndex        =   2
      Top             =   60
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6315
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   6165
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

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

