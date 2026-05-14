VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00BA9F8D&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4965
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3465
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   1860
      Top             =   3705
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1650
      Top             =   4560
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COLETOR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   1
      Left            =   570
      TabIndex        =   1
      Top             =   2220
      Width           =   2400
   End
   Begin VB.Image Image2 
      Height          =   930
      Left            =   1575
      Picture         =   "frmSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   165
      Width           =   1875
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DataField       =   "imagem"
      DataSource      =   "Adodc1"
      Height          =   1080
      Left            =   120
      Stretch         =   -1  'True
      Top             =   195
      Width           =   1395
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Versão 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   270
      Left            =   150
      TabIndex        =   0
      Tag             =   "Version"
      Top             =   3210
      Width           =   3180
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Horizontal Me, RGB(255, 255, 255), RGB(156, 198, 239)
'Me.Show
'Me.Refresh

Me.Top = 0
Me.Left = 0

Cadastrar_Menus

lblVersion = "VERSÃO " & App.Major & "." & App.Minor & "." & App.Revision & "." & App.Comments & "       Base : " & Esquema
lbllicense = "Licenciado para"

'Adodc1.ConnectionString = ConnectionString
'Adodc1.RecordSource = "SELECT imagem FROM " & Esquema_Op & ".TB_EMPRESAS WHERE AUTONUM = " & Cod_Empresa
'Adodc1.Refresh
'Set Adodc1.Recordset.ActiveConnection = Nothing

End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Unload Me
Principal.Show
End Sub

