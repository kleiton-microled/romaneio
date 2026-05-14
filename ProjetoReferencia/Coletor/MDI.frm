VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.MDIForm MDI 
   BackColor       =   &H8000000C&
   Caption         =   "REDEX"
   ClientHeight    =   4410
   ClientLeft      =   4665
   ClientTop       =   2265
   ClientWidth     =   3960
   Icon            =   "MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   4020
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   9702
            MinWidth        =   9702
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1429
            MinWidth        =   1411
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1429
            MinWidth        =   1411
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1429
            MinWidth        =   1411
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1429
            MinWidth        =   1411
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1402
            MinWidth        =   1411
            TextSave        =   "13:54"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   2223
            MinWidth        =   1941
            TextSave        =   "11/06/2018"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   3960
      TabIndex        =   1
      Top             =   0
      Width           =   3960
   End
   Begin VB.Menu MnuDescarga 
      Caption         =   "Descarga"
   End
   Begin VB.Menu MnuEstufagem 
      Caption         =   "Estufagem"
   End
   Begin VB.Menu MnuSair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
'&H80000001&
'Unload frmLogin

Set frmLogin = Nothing



'Dim Rs1 As New ADODB.Recordset
'Dim Rs2 As New ADODB.Recordset
'sql = "SELECT a.autonum_pcs, b.autonum_boo FROM " & Esquema_Op & ".TB_PATIO_cs a, tb_booking_carga b WHERE"
'sql = sql & " a.autonum_bcg = b.autonum_bcg"
'PRSet Rs1, sql
'Do While Not Rs1.EOF
'    sql = "update " & Esquema_Op & ".tb_amr_gate set id_booking = " & Rs1!autonum_boo
'    sql = sql & " where cs_rdx = " & Rs1!autonum_pcs
'    DB.Execute sql
'
'    Rs1.MoveNext
'Loop
'Rs1.Close
'MsgBox "FIM"
'End



'Controlar_Acesso_Mnus

End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'If UnloadMode = vbFormControlMenu Then Cancel = True
'If UnloadMode = vbFormCode Then Cancel = False
'If UnloadMode = vbAppWindows Then Cancel = True
'If UnloadMode = vbAppTaskManager Then Cancel = True
'If UnloadMode = vbFormMDIForm Then Cancel = False
'If UnloadMode = vbFormOwner Then Cancel = False
'If Cancel = True Then
'    MsgBox "Por favor, encerre este aplicativo pelo menu.", vbOKOnly + vbInformation, App.FileDescription
'End If
End Sub



Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub



