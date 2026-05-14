VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form frmLibRegistro 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Consulta \ Liberação de Registro"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   3345
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Documentos"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   39
      Top             =   3720
      Width           =   975
   End
   Begin VB.ListBox lstSimultaneos 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2220
      TabIndex        =   35
      Top             =   2670
      Width           =   1065
   End
   Begin MSDataListLib.DataCombo cbTipoVeiculo 
      Bindings        =   "frmLibRegistroS.frx":0000
      Height          =   255
      Left            =   1050
      TabIndex        =   32
      Top             =   3400
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   503
      _Version        =   393216
      MatchEntry      =   -1  'True
      Appearance      =   0
      Style           =   2
      ListField       =   "Display"
      BoundColumn     =   "Autonum"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtLocal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   1300
      Width           =   1665
   End
   Begin VB.TextBox TxtFiltroBL 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   870
      TabIndex        =   28
      Top             =   510
      Width           =   2370
   End
   Begin VB.TextBox txtCpfMot 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1680
      TabIndex        =   24
      Top             =   2145
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.ListBox listaCarga 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   23
      Top             =   3195
      Width           =   3285
   End
   Begin VB.TextBox txtPlacas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2490
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   2120
      Width           =   795
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdRegistra 
      Caption         =   "Registrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtSefaz 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   20
      Top             =   2910
      Width           =   2205
   End
   Begin VB.TextBox txtSiscarga 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   18
      Top             =   2550
      Width           =   2205
   End
   Begin VB.TextBox txtMot 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   16
      Top             =   2120
      Width           =   2445
   End
   Begin VB.TextBox txtTransp 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   14
      Top             =   1730
      Width           =   3285
   End
   Begin VB.TextBox txtBL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   12
      Top             =   1300
      Width           =   1605
   End
   Begin VB.TextBox txtLote 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2520
      TabIndex        =   10
      Top             =   900
      Width           =   795
   End
   Begin VB.TextBox txtPeriodo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   8
      Top             =   900
      Width           =   2475
   End
   Begin VB.CommandButton cmdFiltra 
      Caption         =   "Filtrar"
      Height          =   240
      Left            =   2565
      TabIndex        =   3
      Top             =   5
      Width           =   675
   End
   Begin MSMask.MaskEdBox mskProtocolo 
      Height          =   240
      Left            =   870
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "######/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskPlaca 
      Height          =   240
      Left            =   870
      TabIndex        =   1
      Top             =   240
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">???-#A##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox CNTR 
      Height          =   240
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "######-#"
      PromptChar      =   "_"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1320
      Top             =   4350
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblGLME 
      AutoSize        =   -1  'True
      Caption         =   "GLME,favor entregar documentação (guias) no Registro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   150
      Left            =   0
      TabIndex        =   38
      Top             =   3960
      Visible         =   0   'False
      Width           =   3330
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prot. Simultâneos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   150
      Index           =   2
      Left            =   2220
      TabIndex        =   37
      Top             =   2520
      Width           =   1050
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Placas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   150
      Index           =   1
      Left            =   30
      TabIndex        =   36
      Top             =   540
      Width           =   375
   End
   Begin VB.Label lblbloqueioCntr 
      AutoSize        =   -1  'True
      Caption         =   "Bloq. Conteiner"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   150
      Left            =   1500
      TabIndex        =   34
      Top             =   2370
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblbloqueioBL 
      AutoSize        =   -1  'True
      Caption         =   "Bloq. BL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   150
      Left            =   2220
      TabIndex        =   33
      Top             =   750
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Veículo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   150
      Left            =   60
      TabIndex        =   31
      Top             =   3400
      Width           =   750
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Localização Carga"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   150
      Left            =   1785
      TabIndex        =   29
      Top             =   1160
      Width           =   1065
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   150
      Left            =   660
      TabIndex        =   27
      Top             =   555
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CNTR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   150
      Index           =   0
      Left            =   1860
      TabIndex        =   26
      Top             =   255
      Width           =   330
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Placa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   150
      Left            =   510
      TabIndex        =   25
      Top             =   255
      Width           =   315
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Placas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   150
      Index           =   0
      Left            =   2580
      TabIndex        =   21
      Top             =   1970
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Status Sefaz"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   60
      TabIndex        =   19
      Top             =   2780
      Width           =   1845
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Status Siscarga"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   60
      TabIndex        =   17
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Motorista"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   150
      Left            =   60
      TabIndex        =   15
      Top             =   1970
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transportadora"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   150
      Left            =   30
      TabIndex        =   13
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Documento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   150
      Left            =   60
      TabIndex        =   11
      Top             =   1160
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lote"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   150
      Left            =   2730
      TabIndex        =   9
      Top             =   750
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   150
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   750
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Protocolo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   150
      Left            =   300
      TabIndex        =   4
      Top             =   30
      Width           =   555
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   8055
      Left            =   -180
      Top             =   0
      Width           =   6885
   End
End
Attribute VB_Name = "frmLibRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False











Public autonumOC As Long
Public Ano As String
Public AutonumCntr
Public AutonumCs
Public autonumag
Public autonumlote

Private Sub cmdFiltra_Click()
Dim Rst As New ADODB.Recordset
Dim SQl As String
Dim rsaux As New ADODB.Recordset
Dim rsPre  As New ADODB.Recordset
Dim prot As String
Dim UF As String


AutonumCntr = 0
AutonumCs = 0
autonumlote = 0


If mskProtocolo.ClipText = "" And mskPlaca.ClipText = "" And TxtFiltroBL = "" Then
      MsgBox "Informe o BL, placa ou protocolo"
      cmdRegistra.Enabled = False
      Exit Sub
End If

If mskProtocolo.ClipText = "" Then
    If mskPlaca.ClipText = "" Then
          MsgBox "Informe a placa"
          cmdRegistra.Enabled = False
          Exit Sub
    End If
    If CNTR.ClipText = "" And TxtFiltroBL = "" Then
          MsgBox "Informe o conteiner ou o BL"
          cmdRegistra.Enabled = False
          Exit Sub
    End If
End If

cmdRegistra.Enabled = True
prot = ObtemNumero(mskProtocolo.Text)
prot = Right("0000000000" & prot, 10)
prot = Val(Mid(prot, 1, 6)) & "/" & Mid(prot, 7, 4)
SQl = ""
SQl = "Select AG.PROTOCOLO, AG.PERIODO, AG.LOTE, AG.NUM_DOCUMENTO,  c.razao As TRANSPORTADORA, AG.NOME_MOTORISTA, AG.PLACA_CAVALO, AG.PLACA_CARRETA, AG.AUTONUM_CS, P.FLAG_CONSULTA_LIBERACAO_COLETOR, AG.CPFMOT "
SQl = SQl & "  ,isnull(ag.id_tipo_caminhao,0) as id_tipo_caminhao,BL.PATIO "
SQl = SQl & " From " & Banco_Sgipa & "VW_AGENDA_CS ag "
SQl = SQl & " inner Join " & Banco_Operador & "tb_cad_transportadoras c ON Ag.COD_TRANSPORTADORA=C.AUTONUM "
SQl = SQl & " INNER JOIN " & Banco_Sgipa & "TB_BL BL ON AG.LOTE = BL.AUTONUM "
SQl = SQl & " INNER JOIN " & Banco_Operador & "TB_PATIOS P ON BL.PATIO = P.AUTONUM "
SQl = SQl & " WHERE 0=0 "
If Me.mskProtocolo.ClipText <> "" Then SQl = SQl & " and ag.PROTOCOLO ='" & prot & "' "
If mskPlaca.ClipText <> "" Then SQl = SQl & "  AND Ag.placa_cavalo='" & Trim(mskPlaca.Text) & "' "
If TxtFiltroBL <> "" Then SQl = SQl & " and bl.numero like '%" & TxtFiltroBL & "%' "

If Me.mskProtocolo.ClipText = "" Then

    SQl = SQl & " union all "
    SQl = SQl & "  Select distinct "
    SQl = SQl & " agcs.PROTOCOLO , agcs.PERIODO"
    SQl = SQl & " ,bl.autonum as lote  , bl.NUM_DOCUMENTO,"
    SQl = SQl & "  c.razao As TRANSPORTADORA, m.nome as NOME_MOTORISTA, a.placa_c as PLACA_CAVALO, A.PLACA_CARRETA, co.cs as AUTONUM_CS , P.FLAG_CONSULTA_LIBERACAO_COLETOR,  m.cpf as CPFMOT"
    SQl = SQl & "  ,isnull(agcs.id_tipo_caminhao,0) as id_tipo_caminhao,BL.PATIO , "
    SQl = SQl & "  From tb_ordem_carregamento a"
    SQl = SQl & "   inner Join operador..tb_cad_transportadoras c ON A.tRANSPORTADORA=C.AUTONUM"
    SQl = SQl & "   inner Join operador..tb_motoristas m ON A.id_motorista=m.AUTONUM"
    SQl = SQl & "   inner Join tb_registro_saida_cs  co ON co.ordem_carreg=a.autonum"
    SQl = SQl & "   inner Join tb_carga_solta  ct ON ct.autonum=co.cs"
    SQl = SQl & "   INNER JOIN TB_BL BL ON ct.bl=bl.autonum"
    SQl = SQl & "   INNER JOIN operador..TB_PATIOS P ON BL.PATIO = P.AUTONUM"
    SQl = SQl & " left join sgipa.dbo.vw_agenda_cs agcs on (a.placa_c=agcs.placa_cavalo and co.cs=agcs.autonum_cs)" 'select protocolo,placa_cavalo,periodo,autonum_cs from VW_AGENDA_CS
    If mskPlaca.ClipText <> "" Then SQl = SQl & " AND A.placa_c='" & Trim(mskPlaca.Text) & "' "
    If TxtFiltroBL <> "" Then SQl = SQl & " and bl.numero like '%" & TxtFiltroBL & "%' "
End If


PRSet Rst, SQl
If Not Rst.EOF Then


AutonumCntr = 0
AutonumCs = Nnull(Rst!AUTONUM_CS, 0)
autonumlote = Nnull(Rst!Lote, 0)
Command1.Enabled = True

    If Col_Patio = 1 Or Col_Patio = 7 Then
        If Rst!Patio <> 1 And Rst!Patio <> 7 Then
            MsgBox "Patio divergente do usuário.", vbExclamation
            Exit Sub
        End If
    Else
        If Col_Patio <> Rst!Patio Then
            MsgBox "Patio divergente do usuário.", vbExclamation
            Exit Sub
        End If
    End If

    If Nnull(Rst!Protocolo, 1) <> "" And mskProtocolo.ClipText = "" Then
        prot = ObtemNumero(Rst!Protocolo)
        prot = Right("0000000000" & prot, 10)
        prot = Mid(prot, 1, 6) & "/" & Mid(prot, 7, 4)
        mskProtocolo = prot
    End If
    txtPeriodo.Text = Nnull(Rst!Periodo, 1)
    TxtLote.Text = Nnull(Rst!Lote, 1)
    txtBL.Text = Nnull(Rst!Num_Documento, 1)
    txtTransp.Text = Nnull(Rst!Transportadora, 1)
    txtMot.Text = Nnull(Rst!NOME_MOTORISTA, 1)
    txtPlacas.Text = Nnull(Rst!PLACA_CAVALO, 1) & vbCrLf & Nnull(Rst!PLACA_CARRETA, 1)
    txtCpfMot.Text = Nnull(Rst!cpfMot, 1)
    
    Me.cbTipoVeiculo.BoundText = Nnull(Rst!id_tipo_caminhao, 0)
    
    
    Dim SQL2 As String
    SQL2 = "select autonum as ordem from tb_ordem_carregamento where protocolo='" & mskProtocolo & "'"
    PRSet rsaux, SQL2
    If rsaux.EOF Then
        rsaux.Close
        SQl = "select isnull(min(ordem_carreg),0) ordem from tb_registro_saida_cs where cs=" & Nnull(Rst!AUTONUM_CS, 9)
        SQl = SQl & " and ordem_carreg in (select autonum from tb_ordem_carregamento where data_ordem is null)"
        If Me.mskPlaca.ClipText <> "" Then
            SQl = SQl & " and ordem_carreg in (select autonum from tb_ordem_carregamento where placa_c='" & Me.mskPlaca & "')"
        End If
        PRSet rsaux, SQl
    End If
    
    
    autonumOC = Nnull(rsaux!Ordem, 0)
    If autonumOC = 0 Then
      MsgBox "Registro de Saida não localizado favor entrar no sistema Gerencial e efetuar o registor de saida da carga"
      cmdRegistra.Enabled = False
      Exit Sub
    End If
    Call Carrega_GridCS(Nnull(autonumOC, 0))
    
    Dim tbLocal As New Recordset
    tbLocal.Open "select (descr_armazem + ' ' + posicao) as display from vw_invent_armazem_item where autonumcs=" & Nnull(Rst!AUTONUM_CS, 0) & " order by qtde desc", db, 1, 1
    If Not tbLocal.EOF Then
        Me.TXTLOCAL.Text = Nnull(tbLocal!display, 1)
    Else
        Me.TXTLOCAL.Text = ""
    End If
    tbLocal.Close
    
    
    
    msgdoc = ""
    msgIcms = ""
    
    If validaEntregaDoc(TxtLote.Text, 2) = 2 Then
        If Nnull(Trim(msgdoc), 1) <> "" Then
            txtSiscarga.Text = msgdoc
        Else
            txtSiscarga.Text = "NÃO ANALISADO"
        End If

        If Nnull(Trim(msgIcms), 1) <> "" Then
            txtSefaz.Text = msgIcms
        Else
            txtSefaz.Text = "NÃO ANALISADO"
        End If
    Else
        txtSefaz.Text = "LIBERADO"
        txtSiscarga.Text = "LIBERADO"
    End If

    If Nnull(Rst!FLAG_CONSULTA_LIBERACAO_COLETOR, 0) > 0 Then
        cmdRegistra.Visible = True
    Else
        cmdRegistra.Visible = False
    End If
    
    
    SQl = "Select isnull(max(isnull(flag_bloqueio,0) + isnull(flag_bloqueio_manual,0)),0) as bloq "
    SQl = SQl & " from sgipa..tb_bl where autonum=" & Val(Me.TxtLote.Text)
    If db.Execute(SQl)(0) > 0 Then
        Me.lblbloqueioBL.Visible = True
    Else
        Me.lblbloqueioBL.Visible = False
    End If
    
    
    
Else
    SQl = "SELECT DISTINCT c.autonum,isnull(c.flag_bloqueio,0) as flag_bloqueio, "
    SQl = SQl & " CONVERT(VARCHAR, Gd.PERIODO_INICIAL, 103) + ' ' + CONVERT(VARCHAR(5),Gd.PERIODO_INICIAL, 108) + ' - ' + CONVERT(VARCHAR, Gd.PERIODO_FINAL, 103) + ' ' + CONVERT(VARCHAR(5), Gd.PERIODO_FINAL, 108) AS periodo, "
    SQl = SQl & " BL.AUTONUM AS LOTE, BL.NUM_DOCUMENTO, t.razao AS TRANSPORTADORA,  B.NOME AS NOME_MOTORISTA, "
    SQl = SQl & " oc.placa_c as placa_cavalo, oc.placa_carreta, P.FLAG_CONSULTA_LIBERACAO_COLETOR, b.cpf as cpfMot, OC.AUTONUM ordem "
    SQl = SQl & ",c.NUM_PROTOCOLO,c.ano_PROTOCOLO, ag.id_tipo_caminhao as id_tipo_caminhao,bl.patio "
    SQl = SQl & " FROM " & Banco_Sgipa & "tb_cntr_bl c INNER JOIN  "
    SQl = SQl & " " & Banco_Sgipa & "TB_AMR_CNTR_BL Amr on c.autonum = amr.cntr inner join "
    SQl = SQl & " " & Banco_Sgipa & "TB_BL BL on Amr.bl= BL.autonum left join "
    SQl = SQl & " " & Banco_Operador & "TB_GD_RESERVA GD on c.autonum_gd_reserva = gd.autonum_gd_reserva inner JOIN "
    SQl = SQl & " " & Banco_Sgipa & "TB_registro_saida_cntr scc on c.autonum=scc.cntr inner join"
    SQl = SQl & " " & Banco_Sgipa & "TB_ordem_carregamento oc on scc.ordem_carreg=oc.autonum"
    SQl = SQl & " inner join " & Banco_Operador & "tb_motoristas b ON oc.id_motorista=b.autonum"
    SQl = SQl & " inner join " & Banco_Operador & "tb_cad_transportadoras t ON oc.transportadora = t.autonum INNER JOIN "
    SQl = SQl & " " & Banco_Operador & "TB_PATIOS P ON C.PATIO = P.AUTONUM "
    SQl = SQl & " left join tb_ag_veiculos ag on c.autonum_veiculo=ag.autonum "
    SQl = SQl & " WHERE 0=0 "
    SQl = SQl & " and bl.flag_ativo=1  "
    If mskPlaca.ClipText = "" Then
        SQl = SQl & "   AND C.NUM_PROTOCOLO=" & Trim(Str(Val(Mid$(prot, 1, 6))))
        SQl = SQl & "   AND C.ANO_PROTOCOLO=" & Trim(Str(Val(Right$(prot, 4))))
    Else
        SQl = SQl & "   and oc.placa_c='" & Trim(mskPlaca.Text) & "'"
        If CNTR.ClipText <> "" Then SQl = SQl & " and SUBSTRING(c.ID_CONTEINER,5,8)='" & CNTR & "'"
    End If
    If TxtFiltroBL <> "" Then SQl = SQl & " and bl.numero like '%" & TxtFiltroBL & "%' "
    
    PRSet Rst, SQl
    
    If Not Rst.EOF Then
        
    AutonumCntr = Nnull(Rst!autonum, 0)
    AutonumCs = 0
    autonumlote = Nnull(Rst!Lote, 0)
    Command1.Enabled = True
        
    If Col_Patio = 1 Or Col_Patio = 7 Then
        If Rst!Patio <> 1 And Rst!Patio <> 7 Then
            MsgBox "Patio divergente do usuário.", vbExclamation
            Exit Sub
        End If
    Else
        If Col_Patio <> Rst!Patio Then
            MsgBox "Patio divergente do usuário.", vbExclamation
            Exit Sub
        End If
    End If
        
        If Nnull(Rst!num_protocolo, 1) <> "" And Nnull(Rst!ano_protocolo, 1) <> "" And mskProtocolo.ClipText = "" Then
            prot = Rst!num_protocolo & Rst!ano_protocolo
            prot = Right("0000000000" & prot, 10)
            prot = Mid(prot, 1, 6) & "/" & Mid(prot, 7, 4)
            mskProtocolo = prot
        End If
        
        txtPeriodo.Text = Nnull(Rst!Periodo, 1)
        TxtLote.Text = Nnull(Rst!Lote, 0)
        txtBL.Text = Nnull(Rst!Num_Documento, 1)
        txtTransp.Text = Nnull(Rst!Transportadora, 1)
        txtMot.Text = Nnull(Rst!NOME_MOTORISTA, 1)
        txtPlacas.Text = Nnull(Rst!PLACA_CAVALO, 1) & vbCrLf & Nnull(Rst!PLACA_CARRETA, 1)
        txtCpfMot.Text = Nnull(Rst!cpfMot, 1)
        Me.cbTipoVeiculo.BoundText = Nnull(Rst!id_tipo_caminhao, 0)
    
        autonumOC = Nnull(Rst!Ordem, 0)
        If autonumOC = 0 Then
          MsgBox "Registro de Saida não localizado favor entrar no sistema Gerencial e efetuar o registro de saida do Conteiner"
          cmdRegistra.Enabled = False
          Exit Sub
        End If
        Call Carrega_GridCNTR(autonumOC)

        msgdoc = ""
        msgIcms = ""
        If validaEntregaDoc(TxtLote.Text, 2) = 2 Then
            If msgdoc <> "" Then
                txtSiscarga.Text = msgdoc
            Else
                txtSiscarga.Text = "NÃO ANALISADO"
            End If

            If msgIcms <> "" Then
                txtSefaz.Text = msgIcms
            Else
                txtSefaz.Text = "NÃO ANALISADO"
            End If
        Else
            txtSefaz.Text = "LIBERADO"
            txtSiscarga.Text = "LIBERADO"
        End If

        If Nnull(Rst!FLAG_CONSULTA_LIBERACAO_COLETOR, 0) > 0 Then
            cmdRegistra.Visible = True
        Else
            cmdRegistra.Visible = False
        End If
        
        
        If Rst!FLAG_BLOQUEIO = 1 Then
            Me.lblbloqueioCntr.Visible = True
        Else
            Me.lblbloqueioCntr.Visible = False
        End If
        
    Else
       MsgBox "Registro de Saida não localizado favor entrar no sistema Gerencial e efetuar o registro de saida da carga", vbExclamation
        Limpa
        mskPlaca.SetFocus

    End If
End If
Rst.Close


If Me.mskProtocolo.ClipText <> "" Then

    Me.lstSimultaneos.Clear
    SQl = "Select DISTINCT B.PROTOCOLO FROM "
    SQl = SQl & " VW_AGENDA_CS A INNER JOIN VW_AGENDA_CS B"
    SQl = SQl & " ON A.PLACA_CAVALO=B.PLACA_CAVALO"
    SQl = SQl & " AND A.PERIODO=B.PERIODO"
    SQl = SQl & " WHERE A.PROTOCOLO='" & Me.mskProtocolo.Text & "'"
    SQl = SQl & " AND B.PROTOCOLO<>'" & Me.mskProtocolo.Text & "' ORDER BY B.PROTOCOLO"
    Dim Tbs As New ADODB.Recordset
    Tbs.Open SQl, db, 1, 1
    While Not Tbs.EOF
        Me.lstSimultaneos.AddItem (Nnull(Tbs!Protocolo, 1))
        Tbs.MoveNext
    Wend

End If
    lblGLME.Visible = False

                SQl = "select count(1) contar  from  SGIPA.DBO.TB_registro_saida_cs  where isnull(flag_glme,0)=0  and ordem_carreg=" & Val(autonumOC)
                SQl = SQl & " and cs in (select autonum from tb_carga_solta where bl in (select bl   from  SGIPA.DBO.TB_etapas_carga  where isnull(flag_glme,0)=1))"

                PRSet rsPre, SQl
                If rsPre!Contar <> 0 Then
                   lblGLME.Visible = True
                End If
 




End Sub

Private Sub Limpa()

mskProtocolo.Text = "______/____"


txtPeriodo.Text = ""
txtSiscarga.Text = ""
txtSefaz.Text = ""
TxtLote.Text = ""
txtBL.Text = ""
txtMot.Text = ""
txtPlacas.Text = ""
txtTransp.Text = ""
listaCarga.Clear
lstSimultaneos.Clear
txtCpfMot.Text = ""
Me.TxtFiltroBL.Text = ""
Me.TXTLOCAL.Text = ""
Me.cbTipoVeiculo.Text = ""

mskPlaca = "___-____"
CNTR = "______-_"

Me.lblbloqueioBL.Visible = False
Me.lblbloqueioCntr.Visible = False
AutonumCntr = 0
AutonumCs = 0
autonumlote = 0

Me.Command1.Enabled = False
End Sub

Private Sub cmdLimpar_Click()
Limpa

End Sub

Private Sub cmdRegistra_Click()
On Error GoTo trataErro:
Dim sSql As String
Dim rsaux As New ADODB.Recordset
Dim rsPre As New ADODB.Recordset
Dim UF As String
Dim Glme As Integer





If Me.lblbloqueioBL.Visible = True Then
    MsgBox "Existe Bloqueio de BL", vbExclamation
    Exit Sub
End If
If Me.lblbloqueioCntr.Visible = True Then
    MsgBox "Existe Bloqueio de Contêiner", vbExclamation
    Exit Sub
End If

 Glme = 0
 
                SQl = "select count(1) contar  from  SGIPA.DBO.TB_registro_saida_cs  where isnull(flag_glme,0)=0  and ordem_carreg=" & Val(autonumOC)
                SQl = SQl & " and cs in (select autonum from tb_carga_solta where bl in (select bl   from  SGIPA.DBO.TB_etapas_carga  where isnull(flag_glme,0)=1))"

                PRSet rsPre, SQl
                If rsPre!Contar <> 0 Then
                   MsgBox "Atenção, Existe Carga sem Liberação GLME ."
                   Exit Sub
                End If
 
                SQl = "select count(1) contar  from  SGIPA.DBO.TB_registro_saida_cntr  where isnull(flag_glme,0)=0  and ordem_carreg=" & Val(autonumOC)
                SQl = SQl & " and cntr in (select cntr from tb_amr_cntr_bl where bl in (select bl   from  SGIPA.DBO.TB_etapas_carga  where isnull(flag_glme,0)=1))"
                PRSet rsPre, SQl
                If rsPre!Contar <> 0 Then
                   MsgBox "Atenção, Existe Carga sem Liberação GLME ."
                   Exit Sub
                End If
    
 
 
 
 
 SQl = "select * from " & Banco_Sgipa & "tb_ordem_carregamento where autonum=" & Val(autonumOC)
 PRSet rsaux, SQl
 If Not rsaux.EOF Then
    UF = Nnull(rsaux!UF, 1)
    If UF = "" Then UF = "SP"
    If Not IsNull(rsaux!data_ordem) Then Exit Sub
 End If
 
 
 
 
 

 
 If VerificaFormaPagamento(TxtLote.Text) <> 3 Then
    If VerificaPendenteVista(TxtLote.Text) > 0 Then
        MsgBox "Existe Pré  Calculo Iniciado SEM PAGAMENTO - Pagamento a Vista"
        Exit Sub
    End If
     'FREE TIME
        SQl = Empty
        SQl = SQl & "SELECT "
        SQl = SQl & "MAX(NVL(Dt_Base_Calculo_Reefer,VALIDADE_GR)) AS FREE_TIME "
        SQl = SQl & "FROM "
        SQl = SQl & "" & Banco_Sgipa & "TB_GR_BL "
        SQl = SQl & "WHERE BL = " & TxtLote.Text & " AND STATUS_GR IN('IM','GE')"
        PRSet rsaux, SQl, 0
        If Not rsaux.EOF Then
            If Nnull(rsaux!free_time, 1) = Empty Then
                    MsgBox "Atenção:" & vbCrLf & "Não Consta Data de Free-Time." & vbCrLf & "Por Favor, Verifique Antes De Prosseguir.", vbOKOnly + vbInformation, "Registro"
                    Exit Sub
            End If
            If Not IsNull(rsaux!free_time) = True Then
                        If CVDate(Format$(Now, "dd/mm/yyyy")) > CVDate(Format$(Nnull(rsaux!free_time, 1), "dd/mm/yyyy")) Then
                            MsgBox "Data de Registro é Maior do Que a Data de Free-Time." & vbCrLf & "Por Favor, Verifique Antes De Prosseguir.", vbOKOnly + vbInformation, "Registro"
                            Exit Sub
                        End If
            End If
        End If
End If


sSql = "SELECT P.FLAG_CONSULTA_LIBERACAO "
sSql = sSql & " FROM " & Banco_Sgipa & "TB_BL BL "
sSql = sSql & " LEFT JOIN " & Banco_Operador & "TB_PATIOS P ON BL.PATIO = P.AUTONUM "
sSql = sSql & " WHERE BL.AUTONUM =" & TxtLote.Text
PRSet rsaux, sSql
If Not rsaux.EOF Then
    If Nnull(rsaux!FLAG_CONSULTA_LIBERACAO, 0) = 0 Then
        MsgBox "Pátio parametrizado para não realizar consultas!", vbInformation
        Exit Sub
    End If
End If


If VerificaFormaPagamento(Val(TxtLote.Text)) <> 3 Then
     If VerificaPendenteVista(Val(TxtLote.Text)) > 0 Then
         MsgBox "Existe Pré  Calculo Iniciado - Pagamento a Vista", vbExclamation
         Exit Sub
     End If
 End If


msgIcms = ""
msgdoc = ""

sSql = "UPDATE " & Banco_Sgipa & "TB_AG_VEICULOS SET ID_TIPO_CAMINHAO=" & Val(Nnull(Me.cbTipoVeiculo.BoundText, 0))
sSql = sSql & " WHERE "
sSql = sSql & " PLACA_CAVALO='" & Left(Me.txtPlacas.Text, 8) & "'"
sSql = sSql & " AND PLACA_CARRETA='" & Right(Me.txtPlacas.Text, 8) & "'"
db.Execute (sSql)




' teste de janela
Entrada = True
SQl = "select c.autonum_gd_reserva ,  c.Periodo_inicial , c.Periodo_Final , c.Janela "
SQl = SQl & " from  SGIPA..TB_registro_saida_cntr  a  inner join sgipa..tb_cntr_bl b on b.autonum=a.cntr  "
SQl = SQl & " inner join operador..tb_gd_reserva c  on b.autonum_gd_reserva=c.autonum_gd_reserva "
SQl = SQl & "  Where a.ordem_carreg = " & Val(autonumOC)
PRSet rsPre, SQl
If rsPre.EOF Then
    SQl = "select c.autonum_gd_reserva ,  c.Periodo_inicial , c.Periodo_Final , c.Janela "
    SQl = SQl & " from  SGIPA..TB_ordem_carregamento a  inner join sgipa..tb_ag_cs b on b.autonum=a.id_agendamento "
    SQl = SQl & " inner join operador..tb_gd_reserva c  on b.autonum_gd_reserva=c.autonum_gd_reserva "
    SQl = SQl & "  Where a.autonum = " & Val(autonumOC)
    PRSet rsPre, SQl
End If

If Not rsPre.EOF Then
   SQl = "Update SGIPA..TB_ordem_carregamento Set USUARIO_REGISTRO_CAM=" & Cod_Usuario & ", AUTONUM_GD_PERIODO=" & Nnull(rsPre!autonum_gd_reserva, 0) & "  where autonum=" & Val(autonumOC)
   Executa SQl
   Entrada = True
   If DateDiff("H", Now, rsPre!periodo_inicial) > Nnull(rsPre!JANELA, 0) Then
      Entrada = False
   End If
   If DateDiff("H", rsPre!Periodo_FINAL, Now) > Nnull(rsPre!JANELA, 0) Then
      Entrada = False
   End If
   If Not Entrada Then
      If MsgBox("Atenção !!Período de Agendamento fora da janela estabelecida  => Janela " & Nnull(rsPre!JANELA, 0) & " hora(s) , Período Agendamento " & Format(rsPre!periodo_inicial, "DD/MM/YYYY HH:MM") & "-" & Format(rsPre!Periodo_FINAL, "DD/MM/YYYY HH:MM") & "! Deseja Continuar ? ", vbYesNo, "Registro de Saida") = vbYes Then
         SQl = "select nome  from "
         SQl = SQl & " sgipa..tb_cad_usuarios where ISNULL(flag_libera_janela,0)=0 and autonum =" & Cod_Usuario
         PRSet Rs, SQl
         If Not Rs.EOF Then
            Entrada = True
            SQl = "Update SGIPA..TB_ordem_carregamento Set FLAG_FORA_PERIODO=1 where autonum=" & Val(autonumOC)
            Executa SQl
         Else
            MsgBox ("  Atenção !! Usuário não tem permissão para Liberar o registro de saída ")
            Exit Sub
         End If
         Rs.Close
      Else
         Exit Sub
      End If
   End If
   If Entrada = False Then Exit Sub
End If



Dim Liberar_Bloqueado As Boolean
Liberar_Bloqueado = False
Dim RsL As New ADODB.Recordset
SQl = "select isnull(flag_libera_sefaz,0) FL, Tipo_Documento from sgipa..tb_bl where autonum=" & TxtLote.Text
PRSet RsL, SQl
If RsL!FL = "1" Then Liberar_Bloqueado = True
Dim EhDTA As Boolean
EhDTA = False
If Nnull(RsL!Tipo_Documento, 0) = 8 Then EhDTA = True
    
    
If validaSaidaComex(Nnull(TxtLote.Text, 0), 2, UF) = False Then
    If msgdoc <> "" Or msgIcms <> "" Then
        MSG = "Liberação Não autorizada (Sefaz\SISCARGA)" & vbCrLf
        If msgdoc <> "" Then MSG = MSG & "Retorno SISCARGA:" & Mid$(msgdoc, 1, 50) & vbCrLf
        If msgIcms <> "" Then MSG = MSG & "Retorno SEFAZ:" & Mid$(msgIcms, 1, 50) & vbCrLf
        If Liberar_Bloqueado And Not EhDTA Then
            MSG = MSG & "Deseja Continuar?"
            If MsgBox(MSG, vbYesNo, Me.Caption) = vbNo Then Liberar_Bloqueado = False
        Else
            X = MsgBox(MSG, vbExclamation)
        End If
        If Not Liberar_Bloqueado And Not EhDTA Then
            txtSefaz.Text = msgIcms
            txtSiscarga.Text = msgdoc
            Exit Sub
       Else
            Call atualizaLiberacao(Val(autonumOC), Nnull(TxtLote.Text, 0), IIf(mskProtocolo.ClipText = "", "", mskProtocolo.Text))
            If mskProtocolo.ClipText <> "" Then
                SQl = "select count(*) from sgipa.dbo.tb_bl where autonum=" & TxtLote.Text & " and flag_ltl=1"
                If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then
                    MsgBox "Carregamento LTL" & vbCr & "Todos os lotes do protocolo foram liberados"
                End If
            End If
            MsgBox ("Registrado com sucesso ")
            Limpa
            mskPlaca.SetFocus
       End If
    Else
        MSG = "Não foi possivel consultar a liberação (Sefaz\SISCARGA)" & vbCrLf
        If Liberar_Bloqueado And Not EhDTA Then
            MSG = MSG & "Deseja Continuar?"
            If MsgBox(MSG, vbYesNo, Me.Caption) = vbNo Then Liberar_Bloqueado = False
        Else
            MsgBox (MSG)
        End If
        If Not Liberar_Bloqueado And Not EhDTA Then
            txtSefaz.Text = Mid$(msgIcms, 1, 10)
            txtSiscarga.Text = Mid$(msgdoc, 1, 10)
            Exit Sub
        Else
            Call atualizaLiberacao(Val(autonumOC), Nnull(TxtLote.Text, 0), IIf(mskProtocolo.ClipText = "", "", mskProtocolo.Text))
            If mskProtocolo.ClipText <> "" Then
                SQl = "select count(*) from sgipa.dbo.tb_bl where autonum=" & TxtLote.Text & " and flag_ltl=1"
                If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then
                    MsgBox "Carregamento LTL" & vbCr & "Todos os lotes do protocolo foram liberados"
                End If
            End If
            MsgBox ("Registrado com sucesso ")
            Limpa
            mskPlaca.SetFocus
        End If
    End If
Else
    'ENTREGA SISCARGA\SEFAZ
    If registraSaidaComex(Nnull(TxtLote.Text, 0), 2, UF, Val(autonumOC), txtCpfMot.Text) = False Then
        If msgdoc <> "" Or msgIcms <> "" Then
            MSG = "Não foi possivel registrar a Saída (Sefaz\SISCARGA)" & vbCrLf
            If msgdoc <> "" Then MSG = MSG & "Retorno SISCARGA:" & Mid$(msgdoc, 1, 50) & vbCrLf
            If msgIcms <> "" Then MSG = MSG & "Retorno SEFAZ:" & Mid$(msgIcms, 1, 50) & vbCrLf
            If Liberar_Bloqueado And Not EhDTA Then
                MSG = MSG & "Deseja Continuar?"
                If MsgBox(MSG, vbYesNo, Me.Caption) = vbNo Then Liberar_Bloqueado = False
            Else
                MsgBox (MSG)
            End If
            If Not Liberar_Bloqueado And Not EhDTA Then
                txtSefaz.Text = Mid$(msgIcms, 1, 10)
                txtSiscarga.Text = Mid$(msgdoc, 1, 10)
                Exit Sub
            Else
                Call atualizaLiberacao(Val(autonumOC), Nnull(TxtLote.Text, 0), IIf(mskProtocolo.ClipText = "", "", mskProtocolo.Text))
                If mskProtocolo.ClipText <> "" Then
                    SQl = "select count(*) from sgipa.dbo.tb_bl where autonum=" & TxtLote.Text & " and flag_ltl=1"
                    If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then
                        MsgBox "Carregamento LTL" & vbCr & "Todos os lotes do protocolo foram liberados"
                    End If
                End If
                X = MsgBox("Registrado com sucesso ", vbInformation)
                Limpa
                mskPlaca.SetFocus
            End If
        Else
            MSG = "Não foi possivel registrar a Saída (Sefaz\SISCARGA)" & vbCrLf
            If Liberar_Bloqueado And Not EhDTA Then
                MSG = MSG & "Deseja Continuar?"
                If MsgBox(MSG, vbYesNo, Me.Caption) = vbNo Then Liberar_Bloqueado = False
            Else
                MsgBox (MSG)
            End If
            If Not Liberar_Bloqueado And Not EhDTA Then
                txtSefaz.Text = Mid$(msgIcms, 1, 10)
                txtSiscarga.Text = Mid$(msgdoc, 1, 10)
                Exit Sub
            Else
               Call atualizaLiberacao(Val(autonumOC), Nnull(TxtLote.Text, 0), IIf(mskProtocolo.ClipText = "", "", mskProtocolo.Text))
                If mskProtocolo.ClipText <> "" Then
                    SQl = "select count(*) from sgipa.dbo.tb_bl where autonum=" & TxtLote.Text & " and flag_ltl=1"
                    If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then
                        MsgBox "Carregamento LTL" & vbCr & "Todos os lotes do protocolo foram liberados"
                    End If
                End If
               X = MsgBox("Registrado com sucesso ", vbInformation)
               Limpa
               mskPlaca.SetFocus
            End If
        End If
    Else
        Call atualizaLiberacao(Val(autonumOC), Nnull(TxtLote.Text, 0), IIf(mskProtocolo.ClipText = "", "", mskProtocolo.Text))
        If mskProtocolo.ClipText <> "" Then
            SQl = "select count(*) from sgipa.dbo.tb_bl where autonum=" & TxtLote.Text & " and flag_ltl=1"
            If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then
                MsgBox "Carregamento LTL" & vbCr & "Todos os lotes do protocolo foram liberados"
            End If
        End If
        X = MsgBox("Saída Registrada com sucesso no SISCARGA\SEFAZ!", vbInformation)
        Limpa
        mskPlaca.SetFocus
    End If
End If
Exit Sub
trataErro:
    MsgBox "Não foi possivel registrar a saida da carga!"
    Err.Clear
End Sub
Private Sub Carrega_GridCS(autonumOC As Long)

listaCarga.Clear

Dim rsaux As New ADODB.Recordset
Dim SQl As String
SQl = "Select 'Carga Solta - ' + CONVERT(VARCHAR,A.ITEM) + ' - ' + C.DESCR AS CARGA "
SQl = SQl & " FROM SGIPA..TB_REGISTRO_SAIDA_CS R "
SQl = SQl & " INNER JOIN SGIPA..TB_CARGA_SOLTA A ON R.CS=A.AUTONUM "
SQl = SQl & " LEFT JOIN SGIPA..DTE_TB_EMBALAGENS C ON A.EMBALAGEM = C.CODE "
SQl = SQl & " WHERE R.ORDEM_CARREG= " & Val(autonumOC)

PRSet rsaux, SQl
While Not rsaux.EOF
    listaCarga.AddItem (Nnull(rsaux!CARGA, 1))
    rsaux.MoveNext
Wend

        
End Sub
    
Private Sub Carrega_GridCNTR(Ordem As Long)

listaCarga.Clear

Dim rsaux As New ADODB.Recordset
Dim SQl As String
        
SQl = " SELECT ID_CONTEINER AS CARGA FROM " & Banco_Sgipa & "TB_CNTR_BL  a inner join tb_registro_saida_cntr b on a.autonum=b.cntr "
SQl = SQl & " WHERE b.ordem_carreg=" & Ordem
PRSet rsaux, SQl
While Not rsaux.EOF
    listaCarga.AddItem (Nnull(rsaux!CARGA, 1))
    rsaux.MoveNext
Wend
    
End Sub


Private Sub Command1_Click()
  Dim Tb_Cv As ADODB.Recordset
  Dim TestaEntrada As Boolean
  Dim Resp As Integer
  Dim nIDBL As Long
  Dim X As String
  Dim TbNovo As ADODB.Recordset
  Dim Dir_Sistema As String
  Dim VLTL As Boolean


If AutonumCntr = 0 And AutonumCs = 0 Then Exit Sub
If autonumlote = 0 Then Exit Sub
  
    
   
VLTL = False
   Dir_Sistema = App.Path & "\ConsultaDocumentos\ConsultaDocColetor.exe "
     '187.60.18.89,1435 SGIPA Tranc#M2017 1 49809 false
  If AutonumCntr > 0 Then
     SQl = "SELECT autonum  FROM TB_CNTR_BL WHERE autonum=" & AutonumCntr
     PRSet TbNovo, SQl
     If Not TbNovo.EOF Then
        TestaEntrada = False
        Call Shell(Dir_Sistema & TrabDsn & " " & UsuBanco & " " & PasBanco & " " & Cod_Usuario & " " & Nnull(TbNovo!autonum, 0) & " " & TestaEntrada & " " & VLTL, vbNormalFocus)
     End If
   End If
  If AutonumCs > 0 Then
    SQl = "SELECT isnull(max(flag_ltl),0) FROM TB_SOLICITACAO_LTL A INNER JOIN TB_BL B ON A.LOTE = B.AUTONUM  WHERE b.autonum=" & autonumlote
    If Nnull(db.Execute(SQl).Fields(0), 0) = 1 Then
         VLTL = True
         TestaEntrada = True
         Call Shell(Dir_Sistema & TrabDsn & " " & UsuBanco & " " & PasBanco & " " & Cod_Usuario & " " & Nnull(autonumlote, 0) & " " & TestaEntrada & " " & VLTL, vbNormalFocus)
    Else
        VLTL = False
        SQl = "SELECT max(autonum) AUTONUM  FROM TB_AG_CS  WHERE LOTE=" & autonumlote
        PRSet TbNovo, SQl
        If Nnull(TbNovo!autonum, 0) > 0 Then
            TestaEntrada = True
            Call Shell(Dir_Sistema & TrabDsn & " " & UsuBanco & " " & PasBanco & " " & Cod_Usuario & " " & Nnull(TbNovo!autonum, 0) & " " & TestaEntrada & " " & VLTL, vbNormalFocus)
        End If
    End If

  
 
  End If

End Sub

Private Sub Form_Activate()
Me.Top = 0
Me.Left = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Dim SQl As String
SQl = "SELECT AUTONUM , DESCR as display FROM " & Banco_Sgipa & "TB_TIPOS_CAMINHAO ORDER BY DESCR"
Me.Adodc1.ConnectionString = ConnectionString
Me.Adodc1.RecordSource = SQl
Me.Adodc1.Refresh
Set Me.Adodc1.Recordset.ActiveConnection = Nothing

End Sub

