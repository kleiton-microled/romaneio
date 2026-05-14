VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SELEÇÃO DE REGISTROS"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6270
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSDataListLib.DataList DataList1 
      Height          =   2700
      Left            =   30
      TabIndex        =   0
      Top             =   225
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   4763
      _Version        =   393216
      Appearance      =   0
      BackColor       =   12708601
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SELECIONE (DUPLO CLIQUE) O ARMADOR CORRETO:"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C1EAF9&
      Height          =   210
      Left            =   23
      TabIndex        =   1
      Top             =   30
      Width           =   6225
   End
End
Attribute VB_Name = "frmSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public s As String
Private Sub DataList1_DblClick()
s = DataList1.BoundText
Me.Hide
End Sub


