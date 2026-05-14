VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form Localizacao 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7110
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
   ScaleHeight     =   8505
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8535
      Left            =   90
      TabIndex        =   0
      Top             =   -120
      Width           =   7065
      Begin VB.TextBox txtResponsavel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2880
         Width           =   4785
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   4050
         Top             =   7470
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
      Begin VB.TextBox txtTransportadora 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2430
         Width           =   5535
      End
      Begin VB.TextBox txtRua 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   5550
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1215
      End
      Begin VB.TextBox txtQuadra 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   465
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1980
         Width           =   2835
      End
      Begin VB.TextBox txtCarga 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   465
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1530
         Width           =   5535
      End
      Begin VB.TextBox TxtPatio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5580
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "@"
         Top             =   180
         Width           =   1185
      End
      Begin VB.TextBox txtGateIN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1965
      End
      Begin VB.TextBox txtLibera 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1995
      End
      Begin VB.TextBox txtCarreta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4770
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   630
         Width           =   1995
      End
      Begin VB.TextBox txtCavalo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   630
         Width           =   1995
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "FILTRO"
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   90
         TabIndex        =   3
         Top             =   3330
         Width           =   6705
         Begin VB.CommandButton Command1 
            Caption         =   "Filtrar"
            Height          =   495
            Left            =   5340
            TabIndex        =   23
            Top             =   150
            Width           =   1245
         End
         Begin MSMask.MaskEdBox mskPlaca 
            Height          =   375
            Left            =   1140
            TabIndex        =   22
            Top             =   210
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">AAA-9999"
            PromptChar      =   "_"
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PLACA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   180
            TabIndex        =   21
            Top             =   210
            Width           =   825
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
         Bindings        =   "Localizacao.frx":0000
         Height          =   4005
         Left            =   90
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   4050
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   7064
         _Version        =   393216
         Cols            =   11
         RowHeightMin    =   420
         WordWrap        =   -1  'True
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   $"Localizacao.frx":0015
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   11
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RESPONSAVEL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   31
         Top             =   2940
         Width           =   1785
      End
      Begin VB.Label LBLTOTCS 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   4770
         TabIndex        =   30
         Top             =   8070
         Width           =   705
      End
      Begin VB.Label LBLTOTCNTR 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   3510
         TabIndex        =   29
         Top             =   8070
         Width           =   705
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5640
         TabIndex        =   28
         Top             =   8100
         Width           =   510
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4320
         TabIndex        =   27
         Top             =   8100
         Width           =   315
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CNTR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2730
         TabIndex        =   26
         Top             =   8100
         Width           =   690
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CARREGAMENTOS : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   25
         Top             =   8100
         Width           =   2490
      End
      Begin VB.Label lblTotEnt 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   6240
         TabIndex        =   24
         Top             =   8070
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRANSP."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   2490
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RUA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   4920
         TabIndex        =   17
         Top             =   2070
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QUADRA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   90
         TabIndex        =   15
         Top             =   2070
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CARGA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   300
         TabIndex        =   12
         Top             =   1590
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GATE IN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3660
         TabIndex        =   10
         Top             =   1170
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LIB. ENT."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   90
         TabIndex        =   8
         Top             =   1110
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CARRETA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   3540
         TabIndex        =   6
         Top             =   750
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CAVALO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   180
         TabIndex        =   4
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CARREGAMENTOS - LOCALIZAÇÃO - PATIO "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   210
         Width           =   5385
      End
   End
End
Attribute VB_Name = "Localizacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Carrega_Grid1
End Sub

Private Sub Form_Activate()
Carrega_Grid1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Me.TxtPatio.text = Nnull(db.Execute("Select descr_resumido from operador..tb_patios where autonum=" & Col_Patio)(0), 1)
'Carrega_Grid1
End Sub

Private Sub Carrega_Grid1()

Screen.MousePointer = 11
Dim Sql As String
Me.Grid1.Visible = False

Sql = " SELECT gn.placa,"
Sql = Sql & "  gn.carreta ,"
Sql = Sql & "  convert(varchar(5),o.dt_lib_ent_cam,103) + ' ' + convert(varchar(5),o.dt_lib_ent_cam,108) dt_lib_ent_cam,"
Sql = Sql & "  convert(varchar(5),gn.dt_gate_in,103) + ' ' + convert(varchar(5),gn.dt_gate_in,108) dt_gate_in,"
Sql = Sql & "  cc.id_conteiner,"
Sql = Sql & "   cc.yard,"
Sql = Sql & "   isnull(Rem.temp_rua,' ') as rua,"
Sql = Sql & "   T.fantasia as Nometransportadora, u.usuario, 1 AS FLAG_CNTR "
Sql = Sql & "   FROM operador..tb_gate_new gn"
Sql = Sql & "   inner join operador..tb_amr_gate amr on gn.autonum = amr.gate"
Sql = Sql & "   inner join sgipa..TB_CNTR_BL cc on amr.CNTR_ipa=cc.AUTONUM"
Sql = Sql & "   inner join operador..TB_CAD_TRANSPORTADORAS T on gn.id_TRANSPORTADORA=T.AUTONUM"
Sql = Sql & "   inner join sgipa..TB_ORDEM_CARREGAMENTO O on Amr.ID_OC=O.AUTONUM"
Sql = Sql & "   left join operador..tb_remocao rem on cc.autonum=rem.AUTONUM and rem.sistema='I'"
Sql = Sql & "   left join sgipa..tb_cad_usuarios u on o.usu_lib_ent_cam=u.autonum "
Sql = Sql & "   Where"
Sql = Sql & "   gn.dt_gate_out is null and gn.dt_gate_in>getdate()-1"
Sql = Sql & "   and cc.patio=" & Col_Patio
If Me.mskPlaca.ClipText <> "" Then Sql = Sql & " AND (GN.PLACA='" & Me.mskPlaca & "' OR GN.CARRETA='" & Me.mskPlaca & "')"
Sql = Sql & "   Union All"
Sql = Sql & "   SELECT max(gn.placa),"
Sql = Sql & "   max(gn.carreta) ,"
Sql = Sql & "   max(convert(varchar(5),o.dt_lib_ent_cam,103) + ' ' + convert(varchar(5),o.dt_lib_ent_cam,108)) dt_lib_ent_cam,"
Sql = Sql & "   max(convert(varchar(5),gn.dt_gate_in,103) + ' ' + convert(varchar(5),gn.dt_gate_in,108)) dt_gate_in,"
Sql = Sql & "   convert(varchar,sum(scs.quantidade)) + ' vols ' as carga  ,"
Sql = Sql & "   max(ISNULL(LOC.ARMAZEM,' ') + '-' + ISNULL(LOC.YARD,' ')) AS YARD,"
Sql = Sql & "   ' ' as rua,"
Sql = Sql & "   max(T.fantasia) as Nometransportadora, max(u.usuario) as usuario , 0 AS FLAG_CNTR"
Sql = Sql & "   FROM operador..tb_gate_new gn"
Sql = Sql & "   inner join operador..tb_amr_gate amr on gn.autonum = amr.gate"
Sql = Sql & "   inner join sgipa..TB_CARGA_SOLTA CS on amr.CS_ipa=cS.AUTONUM"
Sql = Sql & "   left join sgipa..DTE_TB_EMBALAGENS E ON CS.embalagem = E.CODE"
Sql = Sql & "   inner join operador..TB_CAD_TRANSPORTADORAS T on gn.id_TRANSPORTADORA=T.AUTONUM"
Sql = Sql & "   inner join sgipa..TB_ORDEM_CARREGAMENTO O on Amr.ID_OC=O.AUTONUM"
Sql = Sql & "   left join sgipa..tb_registro_saida_cs scs on o.autonum=scs.ordem_carreg"
Sql = Sql & "   LEFT JOIN sgipa..VW_LOCAL_OC LOC ON O.AUTONUM=LOC.AUTONUMORDEM "
Sql = Sql & "   left join sgipa..tb_cad_usuarios u on o.usu_lib_ent_cam=u.autonum "
Sql = Sql & "   Where"
Sql = Sql & "   gn.dt_gate_out is null and gn.dt_gate_in>getdate()-1"
Sql = Sql & "   and cs.patio=" & Col_Patio
If Me.mskPlaca.ClipText <> "" Then Sql = Sql & " AND (GN.PLACA='" & Me.mskPlaca & "' OR GN.CARRETA='" & Me.mskPlaca & "')"
Sql = Sql & "   group by gn.autonum"
Sql = Sql & " order by 1"

Adodc1.ConnectionString = ConnectionStringAdo
Adodc1.RecordSource = Sql
Adodc1.Refresh

Me.lblTotEnt.Caption = Me.Adodc1.Recordset.RecordCount

Set Adodc1.Recordset.ActiveConnection = Nothing

Me.Grid1.FormatString = "   |^Cavalo      |^Carreta     |^Lib. Entrada|^Gate IN       |^Carga               |^Quadra        |^Rua   |<Transportadora                        |<Responsavel                      |"

Me.Grid1.Visible = True

Dim TOTCNTR As Integer
Dim TOTCS As Integer

TOTCNTR = 0
TOTCS = 0

For i = 1 To Grid1.Rows - 1
    If Grid1.TextMatrix(i, Grid1.Cols - 1) = 1 Then
        TOTCNTR = TOTCNTR + 1
    Else
        TOTCS = TOTCS + 1
    End If
Next i

Me.LBLTOTCNTR = TOTCNTR
Me.LBLTOTCS = TOTCS
Screen.MousePointer = 0


End Sub

Private Sub Grid1_Click()
popula
End Sub


Private Sub popula()
Me.txtCavalo.text = Me.Grid1.TextMatrix(Me.Grid1.Row, 1)
Me.txtCarreta.text = Me.Grid1.TextMatrix(Me.Grid1.Row, 2)
Me.txtLibera.text = Me.Grid1.TextMatrix(Me.Grid1.Row, 3)
Me.txtGateIN.text = Me.Grid1.TextMatrix(Me.Grid1.Row, 4)
Me.txtCarga.text = Me.Grid1.TextMatrix(Me.Grid1.Row, 5)
Me.txtQuadra.text = Me.Grid1.TextMatrix(Me.Grid1.Row, 6)
Me.txtRua.text = Me.Grid1.TextMatrix(Me.Grid1.Row, 7)
Me.txtTransportadora.text = Me.Grid1.TextMatrix(Me.Grid1.Row, 8)
Me.txtResponsavel = Me.Grid1.TextMatrix(Me.Grid1.Row, 9)

End Sub

Private Sub Grid1_RowColChange()
popula
End Sub

