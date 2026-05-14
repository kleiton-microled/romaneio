VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmSelecaoPeriodo 
   BackColor       =   &H00C1EAF9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SELEÇÃO DE PERÍODO"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5250
   ControlBox      =   0   'False
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
   ScaleHeight     =   1080
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   285
      Left            =   2730
      TabIndex        =   2
      Top             =   255
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
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
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C1EAF9&
      Caption         =   "&AMBOS"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   3780
      TabIndex        =   5
      Top             =   690
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C1EAF9&
      Caption         =   "&NO TERMINAL"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   2017
      TabIndex        =   4
      Top             =   690
      Width           =   1275
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C1EAF9&
      Caption         =   "&ENTREGUES"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   255
      TabIndex        =   3
      Top             =   705
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   450
      Index           =   1
      Left            =   4530
      Picture         =   "FRMSELECAOPERIODO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "SAIR"
      Top             =   90
      Width           =   450
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   450
      Index           =   0
      Left            =   3960
      Picture         =   "FRMSELECAOPERIODO.frx":076A
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "4"
      ToolTipText     =   "IMPRIMIR"
      Top             =   90
      Width           =   450
   End
   Begin MSMask.MaskEdBox DataIni 
      Height          =   255
      Left            =   285
      TabIndex        =   0
      Top             =   255
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox DataFim 
      Height          =   255
      Left            =   1530
      TabIndex        =   1
      Top             =   255
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TERMINAL"
      Height          =   165
      Index           =   1
      Left            =   2730
      TabIndex        =   10
      Top             =   90
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA FINAL"
      Height          =   165
      Left            =   1530
      TabIndex        =   9
      Top             =   90
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA INICIAL"
      Height          =   165
      Index           =   0
      Left            =   285
      TabIndex        =   8
      Top             =   90
      Width           =   915
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C1EAF9&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      Height          =   945
      Index           =   0
      Left            =   30
      Shape           =   4  'Rounded Rectangle
      Top             =   15
      Width           =   5115
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00EAFBFF&
      Height          =   960
      Index           =   1
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   15
      Width           =   5115
   End
End
Attribute VB_Name = "FrmSelecaoPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nOpRel As Long
Private oRS As ADODB.recordset
Private Sub Form_Load()
Screen.MousePointer = vbHourglass

Sql = "Select descr_resumido as descr, autonum as id from " & Banco_Operador & "tb_patios where "
Sql = Sql & " (cod_empresa =" & Cod_Empresa & " or 0=" & Cod_Empresa & ")"
Sql = Sql & " and recinto <> '000' "
Sql = Sql & " order by descr_resumido "
PRSet oRS, Sql, 0
Set Me.DataCombo1.RowSource = oRS
Me.DataCombo1.ListField = "descr"
Me.DataCombo1.BoundColumn = "id"
Me.Caption = "Seleção de Período/Pátio"
nOpRel = 2
Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Not oRS Is Nothing Then
    If Not oRS.State = 0 Then oRS.Close
    Set oRS = Nothing
End If
nOpRel = Empty
End Sub
Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
    Me.Caption = "-"
    If DataIni.ClipText = Empty Then
        MsgBox "Data inicial não informada.", vbOKOnly + vbInformation, "Contêineres para Desova"
        Exit Sub
    End If
    If DataFim.ClipText = "" Then
        MsgBox "Data final não informada.", vbOKOnly + vbInformation, "Contêineres para Desova"
        Exit Sub
    End If
    If Not IsDate(DataIni) Then
        MsgBox "Data inicial inválida.", vbOKOnly + vbInformation, "Contêineres para Desova"
        Exit Sub
    End If
    If Not IsDate(DataFim) Then
        MsgBox "Data final inválida.", vbOKOnly + vbInformation, "Contêineres para Desova"
        Exit Sub
    End If
Else
    Me.Caption = "!"
End If
Me.Hide
End Sub
Private Sub DataCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub
Private Sub DataFim_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub
Private Sub DataIni_GotFocus()
DataIni.BackColor = vbYellow
End Sub
Private Sub DataIni_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub
Private Sub DataIni_LostFocus()
DataIni.BackColor = vbWhite
End Sub
Private Sub DataFim_GotFocus()
DataFim.BackColor = vbYellow
End Sub
Private Sub DataFim_LostFocus()
DataFim.BackColor = vbWhite
End Sub
Private Sub Option1_Click(Index As Integer)
nOpRel = Index
End Sub
Private Sub Option1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub
