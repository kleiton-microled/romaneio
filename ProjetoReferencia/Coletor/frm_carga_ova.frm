VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_carga_ova 
   BackColor       =   &H00C1EAF9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Carga Conteinerizada (OVA)"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   3165
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C1EAF9&
      Caption         =   "CrossDocking"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   75
      TabIndex        =   53
      Top             =   3780
      Visible         =   0   'False
      Width           =   1290
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3675
      Left            =   0
      TabIndex        =   45
      Top             =   0
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   6482
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Conteiner"
      TabPicture(0)   =   "frm_carga_ova.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DataCombo1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DataCombo2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtreserva"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmd"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Carga Estufada"
      TabPicture(1)   =   "frm_carga_ova.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label8(1)"
      Tab(1).Control(1)=   "Label8(2)"
      Tab(1).Control(2)=   "Label8(3)"
      Tab(1).Control(3)=   "Label8(4)"
      Tab(1).Control(4)=   "Label8(5)"
      Tab(1).Control(5)=   "Label8(6)"
      Tab(1).Control(6)=   "Label8(7)"
      Tab(1).Control(7)=   "Label8(8)"
      Tab(1).Control(8)=   "Label8(9)"
      Tab(1).Control(9)=   "Label8(10)"
      Tab(1).Control(10)=   "DataGrid2"
      Tab(1).Control(11)=   "TxtProduto"
      Tab(1).Control(12)=   "Command1"
      Tab(1).Control(13)=   "Command2"
      Tab(1).Control(14)=   "Txt_qtde"
      Tab(1).ControlCount=   15
      Begin VB.TextBox Txt_qtde 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74190
         TabIndex        =   70
         Tag             =   "qtde_saida"
         Top             =   825
         Width           =   795
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<"
         Height          =   270
         Left            =   -72435
         TabIndex        =   78
         Top             =   1605
         Width           =   420
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">"
         Height          =   300
         Left            =   -72435
         TabIndex        =   77
         Top             =   1095
         Width           =   435
      End
      Begin VB.TextBox TxtProduto 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74385
         TabIndex        =   67
         Top             =   480
         Width           =   2460
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C1EAF9&
         Caption         =   "Talie"
         Enabled         =   0   'False
         Height          =   1965
         Left            =   60
         TabIndex        =   54
         Top             =   1635
         Width           =   3030
         Begin VB.CommandButton CmdGravaTalie 
            Appearance      =   0  'Flat
            Height          =   585
            Left            =   2205
            MaskColor       =   &H00E3F7FF&
            Picture         =   "frm_carga_ova.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "SALVA DADOS NO REGISTRO ATUAL"
            Top             =   315
            Width           =   705
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            ItemData        =   "frm_carga_ova.frx":07A2
            Left            =   675
            List            =   "frm_carga_ova.frx":07AC
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   1545
            Width           =   2295
         End
         Begin MSMask.MaskEdBox MskData 
            Height          =   270
            Index           =   0
            Left            =   690
            TabIndex        =   57
            Top             =   345
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   16
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/#### ##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskData 
            Height          =   270
            Index           =   1
            Left            =   690
            TabIndex        =   58
            Top             =   645
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   16
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/#### ##:##"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo DataCombo5 
            Bindings        =   "frm_carga_ova.frx":07C6
            Height          =   285
            Left            =   690
            TabIndex        =   59
            Top             =   945
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   503
            _Version        =   393216
            MatchEntry      =   -1  'True
            Appearance      =   0
            Style           =   2
            ListField       =   "nome_eqp"
            BoundColumn     =   "autonum_eqp"
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
         Begin MSDataListLib.DataCombo DataCombo4 
            Bindings        =   "frm_carga_ova.frx":07DB
            Height          =   285
            Left            =   690
            TabIndex        =   60
            Top             =   1245
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   503
            _Version        =   393216
            MatchEntry      =   -1  'True
            Appearance      =   0
            Style           =   2
            ListField       =   "nome_eqp"
            BoundColumn     =   "autonum_eqp"
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Confer."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   8
            Left            =   60
            TabIndex        =   66
            Top             =   1005
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Equipe"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   7
            Left            =   60
            TabIndex        =   65
            Top             =   1275
            Width           =   420
         End
         Begin VB.Label LbTalie 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   615
            TabIndex        =   64
            Top             =   120
            Width           =   45
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modo :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   60
            TabIndex        =   63
            Top             =   1560
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "De"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   6
            Left            =   75
            TabIndex        =   62
            Top             =   405
            Width           =   180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Até"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   5
            Left            =   60
            TabIndex        =   61
            Top             =   660
            Width           =   225
         End
      End
      Begin VB.CommandButton cmd 
         Appearance      =   0  'Flat
         Caption         =   "Lista de Conteineres"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         MaskColor       =   &H00E3F7FF&
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "LOCALIZA REGISTRO ESPECÍFICO"
         Top             =   1005
         Width           =   2340
      End
      Begin VB.TextBox txtreserva 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   735
         TabIndex        =   46
         Top             =   375
         Width           =   1545
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frm_carga_ova.frx":07F0
         DataSource      =   "Adodc_cliente"
         Height          =   315
         Left            =   735
         TabIndex        =   48
         Top             =   675
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         ListField       =   "fantasia"
         BoundColumn     =   "autonum"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frm_carga_ova.frx":080C
         DataSource      =   "Adodc_conteiner"
         Height          =   315
         Index           =   0
         Left            =   720
         TabIndex        =   50
         Top             =   1290
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "id_conteiner"
         BoundColumn     =   "id"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frm_carga_ova.frx":082A
         Height          =   1335
         Left            =   -74865
         TabIndex        =   79
         Top             =   1950
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   2355
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   12
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "id_conteiner"
            Caption         =   "Conteiner"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "reference"
            Caption         =   "Reserva"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "descricao_emb"
            Caption         =   "Embalagem"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "mercadoria"
            Caption         =   "Mercadoria"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "num_nf"
            Caption         =   "NF"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "qtde_saida"
            Caption         =   "Quantidade"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "peso_bruto"
            Caption         =   "Peso"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "volume"
            Caption         =   "Volume"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#.##0,000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "razao"
            Caption         =   "Cliente"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "Navio_viagem"
            Caption         =   "Navio \ Viagem"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "cod_produto"
            Caption         =   "Cod. Barra"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1154,835
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   2294,929
            EndProperty
         EndProperty
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   10
         Left            =   -73515
         TabIndex        =   82
         Top             =   3360
         Width           =   105
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estufado no Conteiner :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   9
         Left            =   -74865
         TabIndex        =   81
         Top             =   3390
         Width           =   1245
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   8
         Left            =   -74910
         TabIndex        =   80
         Top             =   900
         Width           =   675
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   7
         Left            =   -74325
         TabIndex        =   75
         Top             =   1725
         Width           =   330
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   6
         Left            =   -74310
         TabIndex        =   74
         Top             =   1440
         Width           =   195
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reserva"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   5
         Left            =   -74310
         TabIndex        =   73
         Top             =   1155
         Width           =   615
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   4
         Left            =   -74895
         TabIndex        =   72
         Top             =   1725
         Width           =   255
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   3
         Left            =   -74895
         TabIndex        =   71
         Top             =   1440
         Width           =   195
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reserva"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   2
         Left            =   -74910
         TabIndex        =   69
         Top             =   1155
         Width           =   510
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   1
         Left            =   -74910
         TabIndex        =   68
         Top             =   525
         Width           =   480
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Contêiner"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   0
         Left            =   45
         TabIndex        =   51
         Top             =   1350
         Width           =   600
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   120
         TabIndex        =   49
         Top             =   750
         Width           =   480
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Reserva:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   105
         TabIndex        =   47
         Top             =   435
         Width           =   555
      End
   End
   Begin VB.TextBox Boo 
      Height          =   285
      Left            =   6015
      TabIndex        =   44
      Text            =   "0"
      Top             =   735
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton Cmd_descarga 
      Appearance      =   0  'Flat
      Caption         =   "<"
      Enabled         =   0   'False
      Height          =   450
      Left            =   6510
      MaskColor       =   &H00E3F7FF&
      Style           =   1  'Graphical
      TabIndex        =   43
      Tag             =   "3"
      Top             =   60
      Width           =   450
   End
   Begin VB.CommandButton Cmd_carga 
      Appearance      =   0  'Flat
      Caption         =   ">"
      Enabled         =   0   'False
      Height          =   450
      Left            =   6075
      MaskColor       =   &H00E3F7FF&
      Style           =   1  'Graphical
      TabIndex        =   42
      Tag             =   "1"
      Top             =   135
      Width           =   450
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6285
      TabIndex        =   40
      Top             =   4020
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6300
      TabIndex        =   38
      Top             =   3555
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtLote 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6015
      TabIndex        =   35
      Top             =   2865
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6045
      TabIndex        =   34
      Top             =   3180
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Frame Frame11 
      Appearance      =   0  'Flat
      BackColor       =   &H00C1EAF9&
      Caption         =   "Dimensões (cm)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1650
      TabIndex        =   23
      Top             =   3840
      Visible         =   0   'False
      Width           =   1080
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   2400
         TabIndex        =   27
         Top             =   400
         Width           =   675
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   72
         MaxLength       =   5
         TabIndex        =   26
         Tag             =   "Altura"
         Top             =   400
         Width           =   660
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   795
         MaxLength       =   5
         TabIndex        =   25
         Tag             =   "Comprimento"
         Top             =   400
         Width           =   690
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1600
         MaxLength       =   5
         TabIndex        =   24
         Tag             =   "Largura"
         Top             =   400
         Width           =   675
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Volume (m3)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   2415
         TabIndex        =   31
         Top             =   195
         Width           =   885
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Altura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   51
         Left            =   120
         TabIndex        =   30
         Top             =   195
         Width           =   435
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Comp."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   50
         Left            =   945
         TabIndex        =   29
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Largura"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   49
         Left            =   1620
         TabIndex        =   28
         Top             =   195
         Width           =   555
      End
   End
   Begin VB.CommandButton Cmd_descarga_tudo 
      Appearance      =   0  'Flat
      Caption         =   "<<"
      Enabled         =   0   'False
      Height          =   450
      Left            =   5550
      MaskColor       =   &H00E3F7FF&
      Style           =   1  'Graphical
      TabIndex        =   22
      Tag             =   "3"
      ToolTipText     =   "LOCALIZA REGISTRO ESPECÍFICO"
      Top             =   855
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CommandButton Cmd_carga_tudo 
      Appearance      =   0  'Flat
      Caption         =   ">>"
      Enabled         =   0   'False
      Height          =   450
      Left            =   4830
      MaskColor       =   &H00E3F7FF&
      Style           =   1  'Graphical
      TabIndex        =   21
      Tag             =   "1"
      ToolTipText     =   "LOCALIZA REGISTRO ESPECÍFICO"
      Top             =   825
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6765
      TabIndex        =   18
      Top             =   1950
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6765
      TabIndex        =   17
      Top             =   1650
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtInstrucao 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5910
      TabIndex        =   13
      Top             =   1170
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.CommandButton cmdImp 
      Appearance      =   0  'Flat
      Height          =   450
      Left            =   7185
      MaskColor       =   &H00E3F7FF&
      Picture         =   "frm_carga_ova.frx":0844
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "4"
      ToolTipText     =   "LISTA REGISTROS COM OPÇÃO DE IMPRESSÃO"
      Top             =   120
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CommandButton cmdSair 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Height          =   450
      Left            =   6765
      MaskColor       =   &H00E3F7FF&
      Picture         =   "frm_carga_ova.frx":1C66
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "SAIR DO MÓDULO"
      Top             =   240
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C1EAF9&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   3285
      TabIndex        =   4
      Top             =   2445
      Width           =   7845
      Begin VB.Frame Frame1 
         BackColor       =   &H00C1EAF9&
         Caption         =   "Dados da Saida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1635
         Left            =   75
         TabIndex        =   5
         Top             =   300
         Width           =   4260
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   1
            Top             =   435
            Width           =   705
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   2895
            TabIndex        =   2
            Tag             =   "volume"
            Top             =   1050
            Width           =   870
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   285
            TabIndex        =   3
            Tag             =   "peso_bruto"
            Top             =   1080
            Width           =   1005
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "frm_carga_ova.frx":23D0
            DataSource      =   "Adodc_emb"
            Height          =   315
            Index           =   1
            Left            =   1095
            TabIndex        =   0
            Tag             =   "AUTONUM_EMB"
            Top             =   420
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ListField       =   "descr"
            BoundColumn     =   "id"
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C1EAF9&
            Caption         =   "Disponível"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   9
            Top             =   210
            Width           =   720
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00C1EAF9&
            Caption         =   "Embalagem"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   1110
            TabIndex        =   8
            Top             =   210
            Width           =   810
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Volume Total"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   2865
            TabIndex        =   7
            Top             =   840
            Width           =   915
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C1EAF9&
            Caption         =   "Peso Total"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   405
            TabIndex        =   6
            Top             =   870
            Width           =   750
         End
      End
   End
   Begin VB.TextBox tst 
      Height          =   285
      Left            =   6210
      TabIndex        =   10
      Top             =   330
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "frm_carga_ova.frx":23E8
      DataSource      =   "Adodc_navio"
      Height          =   315
      Left            =   6075
      TabIndex        =   14
      Top             =   660
      Visible         =   0   'False
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ListField       =   "descr"
      BoundColumn     =   "id"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc_cliente 
      Height          =   330
      Left            =   3570
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   ""
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
   Begin MSAdodcLib.Adodc Adodc_navio 
      Height          =   330
      Left            =   5955
      Top             =   660
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   ""
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   7875
      Top             =   1305
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
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
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   7890
      Top             =   855
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc_emb 
      Height          =   330
      Left            =   4770
      Top             =   390
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc Adodc_SAIDA 
      Height          =   330
      Left            =   3420
      Top             =   795
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   ""
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
   Begin MSAdodcLib.Adodc Adodc_conteiner 
      Height          =   330
      Left            =   3465
      Top             =   465
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   ""
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
   Begin MSAdodcLib.Adodc Adodc_carga 
      Height          =   330
      Left            =   3255
      Top             =   1110
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   ""
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_carga_ova.frx":2402
      Height          =   975
      Left            =   3420
      TabIndex        =   76
      Top             =   1425
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   1720
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   20
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "reference"
         Caption         =   "Reserva"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "instrucao"
         Caption         =   "Instrução"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "num_nf"
         Caption         =   "NF"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "lote"
         Caption         =   "Lote"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "descricao_emv"
         Caption         =   "Embalagem"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "descricao_mer"
         Caption         =   "Mercadoria"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "qtde_entrada"
         Caption         =   "Qtde Entrada"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "qtde_saida"
         Caption         =   "Qtde Saida"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "saldo"
         Caption         =   "Qtde Disp."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "cliente"
         Caption         =   "Cliente"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "navio_viagem"
         Caption         =   "Navio \ Viagem"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "bruto"
         Caption         =   "Peso Unit."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "volume_declarado"
         Caption         =   "Volume Unit."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1560,189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1590,236
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C1EAF9&
      Caption         =   "Mercadoria:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5310
      TabIndex        =   41
      Top             =   4095
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C1EAF9&
      Caption         =   "Reserva :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5475
      TabIndex        =   39
      Top             =   3555
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C1EAF9&
      BackStyle       =   0  'Transparent
      Caption         =   "Lote :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5565
      TabIndex        =   37
      Top             =   2910
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C1EAF9&
      BackStyle       =   0  'Transparent
      Caption         =   "NF :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5715
      TabIndex        =   36
      Top             =   3240
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label LBLID 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   33
      Top             =   4170
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label LBLIDSC 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2325
      TabIndex        =   32
      Top             =   4215
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C1EAF9&
      Caption         =   "Maxgross:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   5910
      TabIndex        =   20
      Top             =   1950
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C1EAF9&
      Caption         =   "Volume:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   6090
      TabIndex        =   19
      Top             =   1665
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C1EAF9&
      Caption         =   "Navio \ Viagem:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   5790
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C1EAF9&
      Caption         =   "Instrução:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5910
      TabIndex        =   15
      Top             =   975
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "frm_carga_ova"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filtro As String
Dim sql_grid_carga As String
Dim sql_grid_saida As String
Dim qtde_disp As Integer
Dim rs_dados_cntr As New ADODB.Recordset
Dim rs_vol As New ADODB.Recordset
Dim vol As Currency
Dim vol_total As Currency
Dim msg_vol As Currency

Sub Busca_Carga()

    If TxtProduto = "" Then Exit Sub

    Adodc_carga.ConnectionString = ConnectionString


    sql_grid_carga = "Select a.instrucao,a.reference, b.autonum_bcg, a.autonum_boo,b.lote,i.autonum_regcs"
    sql_grid_carga = sql_grid_carga & " , case when (select sum(qtde_saida) from " & Esquema_Op & ".tb_saida_carga  where autonum_pcs = c.autonum_pcs) is null then 0 Else"
    sql_grid_carga = sql_grid_carga & " (select sum(qtde_saida) from " & Esquema_Op & ".tb_saida_carga where autonum_pcs = c.autonum_pcs) end as qtde_saida"
    sql_grid_carga = sql_grid_carga & " , c.qtde_entrada,  c.qtde_entrada - case when (select sum(qtde_saida) from " & Esquema_Op & ".tb_saida_carga  where autonum_pcs = c.autonum_pcs) is null  then 0 Else"
    sql_grid_carga = sql_grid_carga & " (select sum(qtde_saida) from " & Esquema_Op & ".tb_saida_carga where autonum_pcs = c.autonum_pcs) end as saldo"
    sql_grid_carga = sql_grid_carga & " , d.razao as cliente, e.descricao_nav || '--' || f.num_viagem as navio_viagem, g.desc_produto as descricao_mer , h.descricao_emb , b.marca, c.*"
    sql_grid_carga = sql_grid_carga & " , j.num_nf,I.AUTONUM_NFI"
    sql_grid_carga = sql_grid_carga & " From"
    sql_grid_carga = sql_grid_carga & " " & Esquema_Op & ".tb_booking a inner join " & Esquema_Op & ".tb_booking_carga b on a.autonum_boo = b.autonum_boo"
    sql_grid_carga = sql_grid_carga & " inner join " & Esquema_Op & ".tb_patio_cs c on b.autonum_bcg = c.autonum_bcg"
    sql_grid_carga = sql_grid_carga & " inner join " & Esquema_Op & ".tb_cad_parceiros d on a.autonum_parceiro = d.autonum"
    sql_grid_carga = sql_grid_carga & " inner join " & Esquema_Op & ".tb_viagens f on a.autonum_via  = f.autonum_via"
    sql_grid_carga = sql_grid_carga & " inner join " & Esquema_Op & ".tb_cad_navios e on f.autonum_nav = e.autonum_nav"
    sql_grid_carga = sql_grid_carga & " inner join " & Esquema_Op & ".tb_cad_produtos g on c.AUTONUM_PRO = g.AUTONUM_PRO"
    sql_grid_carga = sql_grid_carga & " inner join " & Esquema_Op & ".tb_cad_embalagens h on c.autonum_emb = h.autonum_emb"
    sql_grid_carga = sql_grid_carga & " inner join " & Esquema_Op & ".tb_notas_itens i on i.autonum_nfi = c.autonum_nf"
    sql_grid_carga = sql_grid_carga & " inner join " & Esquema_Op & ".tb_notas_fiscais j on j.autonum_nf = i.autonum_nf"
    sql_grid_carga = sql_grid_carga & " WHERE"
    sql_grid_carga = sql_grid_carga & " c.codproduto='" & TxtProduto & "'"

    Adodc_carga.RecordSource = UCase(sql_grid_carga)
    Adodc_carga.Refresh

    If Adodc_carga.Recordset.EOF Then
        TxtProduto = ""
        MsgBox "Produto não encontrado"
        Exit Sub
    End If
    Txt_qtde = Nnull(Adodc_carga.Recordset!saldo, 0)
    Label8(5) = Adodc_carga.Recordset!reFERENCE
    Label8(6) = Nnull(Adodc_carga.Recordset!num_nf, 1)
    Label8(7) = Nnull(Adodc_carga.Recordset!lote, 1)


    If Nnull(Adodc_carga.Recordset!saldo, 0) < 0 Then
        MsgBox "Saldo Disponivel Negativo"
        Txt_qtde = ""
    End If

    If CargaSuzano Then
        sql = "select estufado, pedido,item from " & Esquema_Op & ".tb_integra_carga"
        sql = sql & " where codbarra='" & TxtProduto & "'"
        sql = UCase(sql)
        Dim Rss As New ADODB.Recordset
        Rss.Open sql, DB, adOpenForwardOnly, adLockReadOnly
        If Rss.EOF Then
            TxtProduto = ""
            Rss.Close
            MsgBox "RESERVA DE CARGA SUZANO" & vbCr & "Produto não encontrado no arquivo de integração" & vbCr & "operação cancelada"
            Exit Sub
        End If
        If Nnull(Rss!estufado, 0) = 1 Then
            TxtProduto = ""
            Rss.Close
            MsgBox "RESERVA DE CARGA SUZANO" & vbCr & "Volume já consta como estufado no arquivo de integração" & vbCr & "operação cancelada"
            Exit Sub
        End If
        
        Dim QtdeProg As Integer
        sql = "select quantidade from " & Esquema_Op & ".tb_packing where"
        sql = sql & " pedido = '" & Rss!pedido & "'"
        sql = sql & " and item = '" & Rss!Item & "'"
        sql = sql & " and autonum_patio = " & Nnull(DataCombo1(0).BoundText, 0)
        QtdeProg = Nnull(DB.Execute(UCase(sql)).Fields(0), 0)
        If QtdeProg = 0 Then
            TxtProduto = ""
            Rss.Close
            MsgBox "NÃO CONSTA PROGRAMAÇÃO DE ESTUFAGEM PARA ESTE PEDIDO/ITEM" & vbCr & "operação cancelada"
            Exit Sub
        End If
        
        
        Dim QtdeEstuf As Integer
        sql = "select sum(qtde_saida) from " & Esquema_Op & ".tb_integra_carga cg"
        sql = sql & " inner join " & Esquema_Op & ".tb_saida_carga sc on cg.codbarra = sc.cod_produto"
        sql = sql & " where"
        sql = sql & " cg.pedido = '" & Rss!pedido & "'"
        sql = sql & " and cg.item = '" & Rss!Item & "'"
        sql = sql & " and sc.autonum_patio = " & Nnull(DataCombo1(0).BoundText, 0)
        QtdeEstuf = Nnull(DB.Execute(UCase(sql)).Fields(0), 0)
                
        If QtdeEstuf >= QtdeProg Then
            MsgBox "Todos os volumes programados para este conteiner no pedido selecionado já foram coletados"
            TxtProduto = ""
            TxtProduto.SetFocus
            Exit Sub
        End If
        
        Rss.Close
    End If

End Sub


Sub Limpa_Dados()

    Text4.Text = Empty
    Text2.Text = Empty
    Text3(3).Text = Empty
    Text3(5).Text = Empty
    Text1.Text = Empty
    DataCombo1(1).BoundText = Empty
    txtLote = ""
    Text7 = ""
    Txt_qtde.Text = Empty
    
    Text3(0).Text = Empty
    Text3(1).Text = Empty
    Text3(2).Text = Empty
    Text3(4).Text = Empty
    LBLID.Caption = Empty
    LBLIDSC.Caption = Empty
    
'    Cmd_carga.Enabled = False

End Sub

Sub Mostra_Estufagem()

    sql = "select sum(qtde_saida) from " & Esquema_Op & ".tb_saida_carga"
    sql = sql & " where"
    sql = sql & " autonum_patio = " & Nnull(DataCombo1(0).BoundText, 0)
    Label8(10) = Nnull(DB.Execute(UCase(sql)).Fields(0), 0)

End Sub

Private Sub Adodc_carga_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

If Adodc_carga.Recordset.RecordCount > 0 Then
    
    LBLID.Caption = Nnull(Adodc_carga.Recordset.Fields("AUTONUM_PCS"), 0)
    qtde_disp = Nnull(Adodc_carga.Recordset.Fields("saldo"), 0)
    Text1.Text = Nnull(Adodc_carga.Recordset.Fields("Reference"), 1)
    Text2.Text = Nnull(Adodc_carga.Recordset.Fields("Descricao_mer"), 1)
    Text4.Text = Nnull(Adodc_carga.Recordset.Fields("saldo"), 0)
    
    Text7.Text = Nnull(Adodc_carga.Recordset.Fields("num_nf"), 1)
    Text7.Tag = Nnull(Adodc_carga.Recordset.Fields("autonum_nfi"), 0)
    
    DataCombo1(1).BoundText = Nnull(Adodc_carga.Recordset.Fields("autonum_emb"), 0)
    
    Text3(5).Text = Nnull(Adodc_carga.Recordset.Fields("saldo"), 0) * Nnull(Adodc_carga.Recordset.Fields("bruto"), 0)
    
    
    Text3(0).Text = Nnull(Adodc_carga.Recordset.Fields("Largura"), 0)
    Text3(1).Text = Nnull(Adodc_carga.Recordset.Fields("Altura"), 0)
    Text3(2).Text = Nnull(Adodc_carga.Recordset.Fields("Comprimento"), 0)
    txtLote.Text = Nnull(Adodc_carga.Recordset.Fields("Lote"), 1)
    CALC_VOL
    
'    Cmd_carga_tudo.Enabled = True
'    Cmd_carga.Enabled = True
    
'    List1.Clear
'    sql = "select a.num_nf, b.qtde, b.qtde_estufada,c.desc_produto,d.descricao_emb, b.autonum_nfi from"
'    sql = sql & " " & Esquema_Op & ".tb_notas_fiscais a"
'    sql = sql & " ," & Esquema_Op & ".tb_notas_itens b"
'    sql = sql & " ," & Esquema_Op & ".tb_cad_produtos c"
'    sql = sql & " ," & Esquema_Op & ".tb_cad_embalagens d"
'    sql = sql & " where"
'    sql = sql & " a.autonum_nf = b.autonum_nf"
'    sql = sql & " and b.autonum_pro = c.autonum_pro"
'    sql = sql & " and b.autonum_emb = d.autonum_emb"
'    sql = sql & " and b.autonum_regcs = " & Nnull(Adodc_carga.Recordset!autonum_regcs, 0)
'    Dim Rsnf As New ADODB.Recordset
'    PRSet Rsnf, sql
'    Do While Not Rsnf.EOF
'        List1.AddItem Rsnf!NUM_NF & " / " & Nnull(Rsnf!qtde, 0) - Nnull(Rsnf!qtde_estufada, 0) & " " & Rsnf!DESC_PRODUTO & " " & Rsnf!DESCRICAO_EMB
'        List1.ItemData(NewIndex) = Rsnf!autonum_nfi
'        Rsnf.MoveNext
'    Loop
'    Rsnf.Close
Else
'    Cmd_carga.Enabled = False
'    Cmd_carga_tudo.Enabled = False
End If


End Sub

Private Sub Adodc_SAIDA_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

If Adodc_SAIDA.Recordset.RecordCount > 0 Then
    LBLIDSC.Caption = Nnull(Adodc_SAIDA.Recordset.Fields("AUTONUM_SC"), 0)
    Cmd_descarga.Enabled = True
    Cmd_descarga_tudo.Enabled = True
Else
    Cmd_descarga.Enabled = False
    Cmd_descarga_tudo.Enabled = False
End If

End Sub

Private Sub Cmd_carga_Click()

'If Nnull(Txt_qtde.Text, 0) = 0 Then
'    MsgBox "Favor informar a quantidade/Peso!", vbInformation
'    Exit Sub
'End If
If Txt_qtde = "" Then Txt_qtde = 0

If Val(Txt_qtde.Text) > Val(Text4.Text) Then
    MsgBox "Quantidade Insuficiente em estoque!", vbInformation
    Exit Sub
End If


If DataCombo1(0).BoundText = "" Then
  MsgBox "  Selecione o contêiner.", vbInformation
  DataCombo1(0).SetFocus
  Exit Sub
End If

If Text7 = "" Then
    MsgBox "Carga selecionada não possui NF informada - Operação cancelada"
    Exit Sub
End If


If LbTalie = "" Then
    MsgBox "Não consta talie para este conteiner"
    Exit Sub
End If


'If MsgBox("Deseja vincular a nota fiscal " & Text7 & " ao contêiner '" & Nnull(DataCombo1(0), 1) & "'", vbYesNo + vbQuestion) = vbNo Then
'  Exit Sub
'End If

sql = Empty
sql = sql & "INSERT INTO " & Esquema_Op & ".TB_AMR_NF_SAIDA (autonum,AUTONUM_PATIO, AUTONUM_NFI, QTDE_ESTUFADA) "
sql = sql & " VALUES (redex.seq_amr_nf_saida," & Nnull(DataCombo1(0).BoundText, 0) & "," & Nnull(Text7.Tag, 0) & ", " & Txt_qtde & ")  "
DB.Execute TROCA_DBO(UCase(sql))

'atualiza o campo qtde estufada

sql = Empty
sql = sql & " UPDATE " & Esquema_Op & ".TB_NOTAS_ITENS SET QTDE_ESTUFADA = " & Txt_qtde & " WHERE AUTONUM_NFI = " & Nnull(Text7.Tag, 0) & ""
DB.Execute TROCA_DBO(UCase(sql))

MsgBox "Vínculo efetuado.", vbInformation


'sql = "Select sum(volume) total from " & Esquema_Op & ".tb_saida_carga where autonum_patio = " & Nnull(DataCombo1(0).BoundText, 0)
'PRSet rs_vol, sql
'If rs_vol.RecordCount > 0 Then
'    vol_total = Nnull(rs_vol!Total, 0)
'End If
'
'msg_vol = Nnull(Text5, 0) - Nnull(vol_total + vol, 0)
'If Val(Nnull(vol_total + vol, 0)) > Val(Nnull(Text5, 0)) Then
'    MsgBox "Volume da Carga excede o volume suportado pelo contêiner!" & vbCrLf & vbCrLf & "Volume Suportado: " & Nnull(Text5.Text, 0) & " ." & vbCrLf & "Volume Total da Carga: " & Nnull(vol_total + vol, 0) & " ." & vbCrLf & "Diferença: " & msg_vol & " .", vbInformation
'Else
'    MsgBox "Volume Suportado: " & Nnull(Text5.Text, 0) & " ." & vbCrLf & "Volume Total da Carga: " & Nnull(vol_total + vol, 0) & " ." & vbCrLf & "Diferença: " & msg_vol, vbInformation
'End If


sql = "INSERT INTO " & Esquema_Op & ".TB_SAIDA_CARGA (autonum_sc, AUTONUM_PCS,"
sql = sql & "QTDE_SAIDA,AUTONUM_EMB,PESO_BRUTO,ALTURA,COMPRIMENTO,LARGURA,VOLUME,autonum_patio,ID_CONTEINER,MERCADORIA,DATA_ESTUFAGEM,autonum_nfi"
sql = sql & ",autonum_talie"
sql = sql & ") VALUES (redex.seq_saida_carga.nextval,"
sql = sql & Nnull(LBLID.Caption, 0) & ","
sql = sql & Nnull(Txt_qtde.Text, 0) & "," & Nnull(DataCombo1(1).BoundText, 0) & ","
sql = sql & PPonto(Nnull(Text3(5).Text, 0)) & "," & PPonto(Nnull(Text3(1).Text, 0)) & "," & PPonto(Nnull(Text3(2).Text, 0)) & ","
sql = sql & PPonto(Nnull(Text3(0).Text, 0)) & "," & PPonto(Nnull(Text3(3).Text, 0)) & "," & Nnull(DataCombo1(0).BoundText, 0) & ",'"
sql = sql & Nnull(DataCombo1(0).Text, 1) & "','" & Nnull(Text2.Text, 1) & "',to_date('" & Format(MskData(0), "dd/mm/yyyy") & "','dd/mm/yyyy')"
sql = sql & "," & Nnull(Text7.Tag, 0)
sql = sql & "," & LbTalie
sql = sql & ")"
DB.Execute TROCA_DBO(UCase(sql))

'sql = "update " & Esquema_Op & ".tb_patio set ef='F' where autonum_patio = " & Nnull(DataCombo1(0).BoundText, 0)
'DB.Execute UCase(sql)


If Val(Txt_qtde.Text) - Val(Text4.Text) = 0 Then
    sql = "update " & Esquema_Op & ".tb_patio_cs set flag_historico=1 where autonum_pcs=" & Nnull(LBLID.Caption, 0)
    DB.Execute UCase(sql)
End If


Adodc_carga.Refresh
Adodc_SAIDA.Refresh
limpa_tela
        
End Sub

Private Sub Cmd_carga_tudo_Click()
Dim volu, alt, lar, com As Currency


'Dim IdTalie As Long
'sql = "select autonum_talie from " & Esquema_Op & ".tb_talie where autonum_patio=" & DataCombo1(0).BoundText
'PRSet Rs, sql
'If Rs.EOF Then
'    MsgBox "Não consta talie para este conteiner"
'    Rs.Close
'    Exit Sub
'End If
'IdTalie = Rs!autonum_talie
'Rs.Close


If LbTalie = "" Then
    MsgBox "Não consta talie para este conteiner"
    Exit Sub
End If


Adodc_carga.Recordset.MoveFirst
While Not Adodc_carga.Recordset.EOF
    If Nnull(Adodc_carga.Recordset.Fields("saldo"), 0) <> 0 Then
        sql = "INSERT INTO " & Esquema_Op & ".TB_SAIDA_CARGA (autonum_sc,AUTONUM_PCS,"
        sql = sql & "QTDE_SAIDA,MARCA,AUTONUM_EMB,PESO_BRUTO,ALTURA,COMPRIMENTO,LARGURA,VOLUME,autonum_patio, ID_CONTEINER"
        sql = sql & ",MERCADORIA,DATA_ESTUFAGEM,autonum_talie"
        sql = sql & ") VALUES (redex.saida_carga.nextval,"
        sql = sql & Nnull(Adodc_carga.Recordset.Fields("autonum_pcs"), 0) & ","
        sql = sql & Nnull(Adodc_carga.Recordset.Fields("saldo"), 0) & ",'"
        sql = sql & Nnull(Adodc_carga.Recordset.Fields("marca"), 1) & "',"
        sql = sql & Nnull(Adodc_carga.Recordset.Fields("autonum_emb"), 0) & ","
        sql = sql & PPonto(Nnull(Adodc_carga.Recordset.Fields("bruto"), 0)) & ","
        sql = sql & PPonto(Nnull(Adodc_carga.Recordset.Fields("altura"), 0)) & ","
        sql = sql & PPonto(Nnull(Adodc_carga.Recordset.Fields("comprimento"), 0)) & ","
        sql = sql & PPonto(Nnull(Adodc_carga.Recordset.Fields("largura"), 0)) & ","
        'calcula o volume da carga
        Text3(1).Text = Nnull(Adodc_carga.Recordset.Fields("altura"), 0)
        Text3(0).Text = Nnull(Adodc_carga.Recordset.Fields("largura"), 0)
        Text3(2).Text = Nnull(Adodc_carga.Recordset.Fields("comprimento"), 0)
        Txt_qtde.Text = Nnull(Adodc_carga.Recordset.Fields("saldo"), 0)
        CALC_VOL
        '
        sql = sql & PPonto(Nnull(vol, 0)) & ","
        sql = sql & Nnull(DataCombo1(0).BoundText, 0) & ",'"
        sql = sql & Nnull(DataCombo1(0).Text, 1) & "','"
        sql = sql & Nnull(Adodc_carga!descricao_mer, 1) & "',"
        sql = sql & " sysdate"
        sql = sql & "," & LbTalie
        sql = sql & ")"
        PRSet rs_vol, "Select sum(volume) total from " & Esquema_Op & ".tb_saida_carga where autonum_patio = " & Nnull(DataCombo1(0).BoundText, 0)
        If rs_vol.RecordCount > 0 Then
            vol_total = Nnull(rs_vol!Total, 0)
        End If
        
        msg_vol = Nnull(Text5, 0) - Nnull(vol_total + vol, 0)
        If Val(Nnull(vol_total + vol, 0)) > Val(Nnull(Text5, 0)) Then
            MsgBox "Volume da Carga excede o volume suportado pelo contêiner!" & vbCrLf & vbCrLf & "Volume Suportado: " & Nnull(Text5.Text, 0) & " ." & vbCrLf & "Volume Total da Carga: " & Nnull(vol_total + vol, 0) & " ." & vbCrLf & "Diferença: " & msg_vol & " .", vbInformation
            'Exit Sub
        Else
            MsgBox "Volume Suportado: " & Nnull(Text5.Text, 0) & " ." & vbCrLf & "Volume Total da Carga: " & Nnull(vol_total + vol, 0) & " ." & vbCrLf & "Diferença: " & msg_vol, vbInformation
        End If
        
'        If Nnull(vol_total + vol, 0) > Nnull(Text5, 0) Then
'            MsgBox "Volume da Carga excede o volume suportado pelo contêiner!", vbInformation
'            Exit Sub
'        End If
    
        DB.Execute TROCA_DBO(UCase(sql))
    End If
    Adodc_carga.Recordset.MoveNext
Wend
Adodc_carga.Refresh
Adodc_SAIDA.Refresh
limpa_tela
End Sub


Private Sub limpa_tela()

Text1.Text = Empty
Text2.Text = Empty
DataCombo1(1).BoundText = Empty
Text4.Text = Empty
Txt_qtde.Text = Empty
Text3(0).Text = Empty
Text3(1).Text = Empty
Text3(2).Text = Empty
Text3(3).Text = Empty
Text3(4).Text = Empty
Text3(5).Text = Empty
LBLID.Caption = Empty
LBLIDSC.Caption = Empty

Label8(5) = ""
Label8(6) = ""
Label8(7) = ""

TxtProduto = ""

'Cmd_carga.Enabled = False



End Sub
Private Sub Cmd_Click()
    

If txtreserva = "" Then MsgBox "Informe a reserva": Exit Sub


'Filtra conteineres da reserva
sql = "SELECT a.autonum_patio as id,a.id_conteiner as id_conteiner, c.autonum_boo, c.flag_suzano FROM"
sql = sql & " " & Esquema_Op & ".TB_PATIO A inner join  " & Esquema_Op & ".TB_BOOKING_CARGA B on A.AUTONUM_BCG = B.AUTONUM_BCG"
sql = sql & " inner join " & Esquema_Op & ".TB_BOOKING C on B.AUTONUM_BOO = C.AUTONUM_BOO"
sql = sql & " WHERE "   'A.FLAG_HISTORICO = 0 and"
If txtreserva <> "" Then sql = sql & " C.REFERENCE = '" & txtreserva & "'"
If txtInstrucao <> "" Then sql = sql & " AND C.instrucao = '" & txtInstrucao & "'"
If DataCombo2.BoundText <> "" Then sql = sql & " and c.autonum_parceiro = " & DataCombo2.BoundText
If DataCombo3.BoundText <> "" Then sql = sql & " and c.autonum_via = " & DataCombo3.BoundText
sql = sql & " order by a.id_conteiner"
Adodc_conteiner.ConnectionString = ConnectionString
Adodc_conteiner.RecordSource = UCase(sql)
Adodc_conteiner.Refresh

CargaSuzano = False
If Not Adodc_conteiner.Recordset.EOF Then
    Boo = Adodc_conteiner.Recordset!autonum_boo
    If Nnull(Adodc_conteiner.Recordset!flag_suzano, 0) = 1 Then CargaSuzano = True
Else
    Boo = 0
End If

'------------------------------

'filtro = ""
'If txtreserva <> "" Then filtro = filtro & " AND a.REFERENCE = '" & txtreserva & "'"
'If Nnull(DataCombo2.BoundText, 0) <> 0 Then filtro = filtro & " and a.autonum_parceiro =" & Nnull(DataCombo2.BoundText, 0)
'If Nnull(txtInstrucao.Text, 1) <> "" Then filtro = filtro & " and a.instrucao ='" & Nnull(txtInstrucao.Text, 1) & "' "
'If Nnull(DataCombo3.BoundText, 0) <> 0 Then filtro = filtro & " and a.autonum_via = " & Nnull(DataCombo3.BoundText, 0)
'Adodc_carga.RecordSource = TROCA_DBO(UCase(sql_grid_carga & filtro))
'Adodc_carga.Refresh


filtro = ""
If txtreserva <> "" Then filtro = filtro & " AND d.REFERENCE = '" & txtreserva & "'"
If Nnull(DataCombo2.BoundText, 0) <> 0 Then filtro = filtro & " and d.autonum_parceiro =" & Nnull(DataCombo2.BoundText, 0)
If Nnull(txtInstrucao.Text, 1) <> "" Then filtro = filtro & " and d.instrucao ='" & Nnull(txtInstrucao.Text, 1) & "' "
If Nnull(DataCombo3.BoundText, 0) <> 0 Then filtro = filtro & " and d.autonum_via = " & Nnull(DataCombo3.BoundText, 0)

'Adodc_SAIDA.RecordSource = TROCA_DBO(UCase(sql_grid_saida & filtro & " ORDER BY CC.ID_CONTEINER"))
'Adodc_SAIDA.Refresh



Frame3.Enabled = True
Frame4.Enabled = False

DataCombo1(0).SetFocus


End Sub

Private Sub Cmd_descarga_Click()
    
    qtde_saida = Nnull(Adodc_SAIDA.Recordset.Fields("QTDE_SAIDA"), 0)
    nfi = Adodc_SAIDA.Recordset.Fields("autonum_nfi")
    sc = Adodc_SAIDA.Recordset.Fields("autonum_sc")
    p = Adodc_SAIDA.Recordset.Fields("autonum_patio")

    sql = "delete from " & Esquema_Op & ".TB_AMR_NF_SAIDA"
    sql = sql & " where"
    sql = sql & " autonum_nfi=" & nfi
    sql = sql & " and autonum_patio=" & p
    DB.Execute TROCA_DBO(UCase(sql))
    
    sql = "UPDATE " & Esquema_Op & ".TB_NOTAS_ITENS SET QTDE_ESTUFADA = qtde_estufada - " & qtde_saida & " WHERE AUTONUM_NFI = " & nfi
    DB.Execute TROCA_DBO(UCase(sql))
    
    sql = "delete from tb_saida_carga where autonum_sc=" & sc
    DB.Execute sql
    
'    sql = "select autonum_sc from tb_saida_carga where autonum_patio=" & p
'    PRSet Rs, sql
'    If Rs.EOF Then
        sql = "update " & Esquema_Op & ".tb_patio set ef='E' where autonum_patio = " & p
        DB.Execute UCase(sql)
'    End If
'    Rs.Close
    
    Adodc_carga.Refresh
    Adodc_SAIDA.Refresh
    Call Limpa_Dados

End Sub

Private Sub Cmd_descarga_tudo_Click()

If MsgBox("Deseja Retornar todas as Cargas ao Estoque?", vbYesNo) = vbYes Then
    Adodc_SAIDA.Recordset.MoveFirst
    While Not Adodc_SAIDA.Recordset.EOF
'        sql = "DELETE " & Esquema_Op & ".TB_SAIDA_CARGA WHERE AUTONUM_SC = " & Nnull(Adodc_SAIDA.Recordset.Fields("autonum_sc"), 0)
'        DB.Execute TROCA_DBO(UCase(sql))
        
        
        qtde_saida = Nnull(Adodc_SAIDA.Recordset.Fields("QTDE_SAIDA"), 0)
        nfi = Adodc_SAIDA.Recordset.Fields("autonum_nfi")
        sc = Adodc_SAIDA.Recordset.Fields("autonum_sc")
        p = Adodc_SAIDA.Recordset.Fields("autonum_patio")
    
        sql = "delete from " & Esquema_Op & ".TB_AMR_NF_SAIDA"
        sql = sql & " where"
        sql = sql & " autonum_nfi=" & nfi
        sql = sql & " and autonum_patio=" & p
        DB.Execute TROCA_DBO(UCase(sql))
        
        sql = "UPDATE " & Esquema_Op & ".TB_NOTAS_ITENS SET QTDE_ESTUFADA = qtde_estufada - " & qtde_saida & " WHERE AUTONUM_NFI = " & nfi
        DB.Execute TROCA_DBO(UCase(sql))
        
        sql = "delete from tb_saida_carga where autonum_sc=" & sc
        DB.Execute sql
        
    '    sql = "select autonum_sc from tb_saida_carga where autonum_patio=" & p
    '    PRSet Rs, sql
    '    If Rs.EOF Then
            sql = "update " & Esquema_Op & ".tb_patio set ef='E' where autonum_patio = " & p
            DB.Execute UCase(sql)
    '    End If
    '    Rs.Close
        
        Adodc_SAIDA.Recordset.MoveNext
    Wend
End If
Adodc_carga.Refresh
Adodc_SAIDA.Refresh
limpa_tela

End Sub



Private Sub CmdGravaTalie_Click()

        If DataCombo1(0).BoundText = "" Then
            MsgBox "Conteiner não selecionado"
            Exit Sub
        End If
        If MskData(0).ClipText = "" Then
            MsgBox "Favor informar a data de inicio"
            MskData(0).SetFocus
            Exit Sub
        End If
        If Not IsDate(MskData(0)) Then
            MsgBox "Data Inválida"
            MskData(0).SetFocus
            Exit Sub
        End If
        If MskData(1).ClipText <> "" Then
            If Not IsDate(MskData(1)) Then
                MsgBox "Data Inválida"
                MskData(0).SetFocus
                Exit Sub
            End If
            If CDate(MskData(1)) < CDate(MskData(0)) Then
                MsgBox "Período Inválido"
                MskData(0).SetFocus
                Exit Sub
            End If
        End If
        
        If DataCombo5.BoundText = "" Then
            MsgBox "Conferente não informado"
            DataCombo5.SetFocus
            Exit Sub
        End If
        If DataCombo4.BoundText = "" Then
            MsgBox "Equipe não informada"
            DataCombo4.SetFocus
            Exit Sub
        End If
        
        
        sql = "select dt_entrada from " & Esquema_Op & ".tb_patio where autonum_patio = " & DataCombo1(0).BoundText
        PRSet Rs, sql
        
'        MsgBox CDate(MskData(0)) & " ..... " & Rs!dt_entrada
        
        If CDate(MskData(0)) < Rs!dt_entrada Then
            Rs.Close
            MsgBox "Inicio de estufagem não pode ser anterior a data de entrada do conteiner"
            Exit Sub
        End If
        Rs.Close
        
        
        
        'consiste quantidade programada com a quantidade estufada
        If MskData(1).ClipText <> "" Then
            Dim qe As Long
            sql = "SELECT sum(qtde_saida) FROM " & Esquema_Op & ".TB_SAIDA_CARGA WHERE  autonum_patio=" & DataCombo1(0).BoundText
            qe = Nnull(DB.Execute(sql).Fields(0), 0)
            
            Dim qp As Long
            sql = "select sum(quantidade) from " & Esquema_Op & ".tb_packing where autonum_patio=" & DataCombo1(0).BoundText
            qp = Nnull(DB.Execute(sql).Fields(0), 0)
            
            If (qp <> qe) Or (qe = 0) Then
                MsgBox "Quantidade Programada diverge da quantidade estufada para este conteiner" & vbCr & "Fechamento de Talie não permitido" & vbCr & "Operação Cancelada"
                Exit Sub
            End If
        End If
        
        
        
        
        If LbTalie = "0" Then
            
            sql = "select autonum_talie from " & Esquema_Op & ".tb_talie where autonum_patio=" & DataCombo1(0).BoundText & " and nvl(crossdocking,0)=0"
            PRSet Rs, sql
            If Not Rs.EOF Then
                MsgBox "Já existe talie para este conteiner"
                Rs.Close
                Exit Sub
            End If
            
            sql = UCase("select redex.seq_talie.nextval from dual")
            PRSet Rs, sql
            TxtRegistro = Rs.Fields(0)
            LbTalie.Caption = TxtRegistro
            LbTalie.Refresh
            Rs.Close
            
            sql = "Insert into " & Esquema_Op & ".tb_talie (autonum_talie,"
            sql = sql & "autonum_patio,inicio,termino,flag_estufagem"
            sql = sql & ",crossdocking,autonum_boo,forma_operacao,conferente,equipe"
            sql = sql & ") values (" & TxtRegistro & ","
            sql = sql & DataCombo1(0).BoundText
            sql = sql & ",to_date('" & Format(MskData(0), "dd/mm/yyyy hh:mm") & "','dd/mm/yyyy hh24:mi')"
            If MskData(1).ClipText = "" Then
                sql = sql & ",NULL"
            Else
                sql = sql & ",to_date('" & Format(MskData(1), "dd/mm/yyyy hh:mm") & "','dd/mm/yyyy hh24:mi')"
            End If
            sql = sql & ",1"
            sql = sql & "," & Check1.Value
            sql = sql & "," & Boo
            sql = sql & ",'A'"
            sql = sql & "," & DataCombo5.BoundText
            sql = sql & "," & DataCombo4.BoundText
            sql = sql & ")"
            DB.Execute UCase(sql)
            

        Else
            Rec = Nnull(LbTalie, 0)
            
            sql = "update " & Esquema_Op & ".tb_talie set"
            sql = sql & " autonum_patio=" & DataCombo1(0).BoundText
            sql = sql & ",inicio=to_date('" & Format(MskData(0), "dd/mm/yyyy hh:mm") & "','dd/mm/yyyy hh24:mi')"
            If MskData(1).ClipText <> "" Then sql = sql & ",termino=to_date('" & Format(MskData(1), "dd/mm/yyyy hh:mm") & "','dd/mm/yyyy hh24:mi')"
            sql = sql & ",crossdocking=" & Check1.Value
            sql = sql & ",autonum_boo = " & Boo
            sql = sql & ",forma_operacao='" & UCase(Left(Combo1, 1)) & "'"
            sql = sql & ",conferente=" & DataCombo5.BoundText
            sql = sql & ",equipe=" & DataCombo4.BoundText
            sql = sql & " where autonum_talie = " & LbTalie
            DB.Execute UCase(sql)
            
        End If
        
        If MskData(1).ClipText <> "" Then
            sql = "update " & Esquema_Op & ".tb_patio set ef='F' where autonum_patio = " & DataCombo1(0).BoundText
        Else
            sql = "update " & Esquema_Op & ".tb_patio set ef='E' where autonum_patio = " & DataCombo1(0).BoundText
        End If
        DB.Execute UCase(sql)
        
'        If MskData(1).ClipText <> "" Then
'            Call Evento_06(CLng(LbTalie), "O", "", "")
'            Call Evento_07(CLng(LbTalie), "O")
'        End If

End Sub

Private Sub cmdImp_Click()
frmImpressaoCargaConteinerizada.Show 1
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Command1_Click()

If Val(Txt_qtde) = 0 Then
    MsgBox "Quantidade não informada"
    Txt_qtde.SetFocus
    Exit Sub
End If

If Val(Txt_qtde.Text) > Val(Text4.Text) Then
    MsgBox "Quantidade Insuficiente em estoque!", vbInformation
    Exit Sub
End If


If DataCombo1(0).BoundText = "" Then
  MsgBox "  Selecione o contêiner.", vbInformation
  DataCombo1(0).SetFocus
  Exit Sub
End If

If Text7 = "" Then
    MsgBox "Carga selecionada não possui NF informada - Operação cancelada"
    Exit Sub
End If


If LbTalie = "0" Then
    MsgBox "Não consta talie para este conteiner"
    Exit Sub
End If


'If MsgBox("Deseja vincular a nota fiscal " & Text7 & " ao contêiner '" & Nnull(DataCombo1(0), 1) & "'", vbYesNo + vbQuestion) = vbNo Then
'  Exit Sub
'End If

sql = Empty
sql = sql & "INSERT INTO " & Esquema_Op & ".TB_AMR_NF_SAIDA (autonum,AUTONUM_PATIO, AUTONUM_NFI, QTDE_ESTUFADA) "
sql = sql & "VALUES(redex.seq_amr_nf_saida.nextval," & Nnull(DataCombo1(0).BoundText, 0) & "," & Nnull(Text7.Tag, 0) & ", " & Txt_qtde & ")  "
DB.Execute TROCA_DBO(UCase(sql))

'atualiza o campo qtde estufada

sql = Empty
sql = sql & " UPDATE " & Esquema_Op & ".TB_NOTAS_ITENS SET QTDE_ESTUFADA = " & Txt_qtde & " WHERE AUTONUM_NFI = " & Nnull(Text7.Tag, 0) & ""
DB.Execute TROCA_DBO(UCase(sql))

MsgBox "Vínculo efetuado.", vbInformation


'sql = "Select sum(volume) total from " & Esquema_Op & ".tb_saida_carga where autonum_patio = " & Nnull(DataCombo1(0).BoundText, 0)
'PRSet rs_vol, sql
'If rs_vol.RecordCount > 0 Then
'    vol_total = Nnull(rs_vol!Total, 0)
'End If
'
'msg_vol = Nnull(Text5, 0) - Nnull(vol_total + vol, 0)
'If Val(Nnull(vol_total + vol, 0)) > Val(Nnull(Text5, 0)) Then
'    MsgBox "Volume da Carga excede o volume suportado pelo contêiner!" & vbCrLf & vbCrLf & "Volume Suportado: " & Nnull(Text5.Text, 0) & " ." & vbCrLf & "Volume Total da Carga: " & Nnull(vol_total + vol, 0) & " ." & vbCrLf & "Diferença: " & msg_vol & " .", vbInformation
'Else
'    MsgBox "Volume Suportado: " & Nnull(Text5.Text, 0) & " ." & vbCrLf & "Volume Total da Carga: " & Nnull(vol_total + vol, 0) & " ." & vbCrLf & "Diferença: " & msg_vol, vbInformation
'End If


sql = "INSERT INTO " & Esquema_Op & ".TB_SAIDA_CARGA (autonum_sc,AUTONUM_PCS,"
sql = sql & "QTDE_SAIDA,AUTONUM_EMB,PESO_BRUTO,ALTURA,COMPRIMENTO,LARGURA,VOLUME,autonum_patio,ID_CONTEINER,MERCADORIA,DATA_ESTUFAGEM,autonum_nfi"
sql = sql & ",autonum_talie,COD_PRODUTO"
sql = sql & ") VALUES (redex.seq_saida_carga.nextval,"
sql = sql & Nnull(LBLID.Caption, 0) & ","
sql = sql & Nnull(Txt_qtde.Text, 0) & "," & Nnull(DataCombo1(1).BoundText, 0) & ","

sql = sql & PPonto(Nnull(Adodc_carga.Recordset!bruto, 0) * Val(Txt_qtde))

sql = sql & " ," & PPonto(Nnull(Text3(1).Text, 0)) & "," & PPonto(Nnull(Text3(2).Text, 0)) & ","
sql = sql & PPonto(Nnull(Text3(0).Text, 0)) & "," & PPonto(Nnull(Text3(3).Text, 0)) & "," & Nnull(DataCombo1(0).BoundText, 0) & ",'"
sql = sql & Nnull(DataCombo1(0).Text, 1) & "','" & Nnull(Text2.Text, 1) & "',to_date('" & Format(MskData(0), "dd/mm/yyyy") & "','dd/mm/yyyy')"
sql = sql & "," & Nnull(Text7.Tag, 0)
sql = sql & "," & LbTalie
sql = sql & ",'" & TxtProduto & "'"
sql = sql & ")"
DB.Execute TROCA_DBO(UCase(sql))

'sql = "update " & Esquema_Op & ".tb_patio set ef='F' where autonum_patio = " & Nnull(DataCombo1(0).BoundText, 0)
'DB.Execute UCase(sql)


If Val(Txt_qtde.Text) - Val(Text4.Text) = 0 Then
    sql = "update " & Esquema_Op & ".tb_patio_cs set flag_historico=1 where autonum_pcs=" & Nnull(LBLID.Caption, 0)
    DB.Execute UCase(sql)
End If


If CargaSuzano Then
    sql = "update " & Esquema_Op & ".tb_integra_carga set estufado=1 where codbarra='" & TxtProduto & "'"
    DB.Execute UCase(sql)
End If



Adodc_carga.Refresh
Adodc_SAIDA.Refresh
limpa_tela


Call Mostra_Estufagem



End Sub

Private Sub Command2_Click()
    qtde_saida = Nnull(Adodc_SAIDA.Recordset.Fields("QTDE_SAIDA"), 0)
    nfi = Adodc_SAIDA.Recordset.Fields("autonum_nfi")
    sc = Adodc_SAIDA.Recordset.Fields("autonum_sc")
    p = Adodc_SAIDA.Recordset.Fields("autonum_patio")

    sql = "delete from " & Esquema_Op & ".TB_AMR_NF_SAIDA"
    sql = sql & " where"
    sql = sql & " autonum_nfi=" & nfi
    sql = sql & " and autonum_patio=" & p
    DB.Execute TROCA_DBO(UCase(sql))
    
    sql = "UPDATE " & Esquema_Op & ".TB_NOTAS_ITENS SET QTDE_ESTUFADA = qtde_estufada - " & qtde_saida & " WHERE AUTONUM_NFI = " & nfi
    DB.Execute TROCA_DBO(UCase(sql))
    
    sql = "delete from tb_saida_carga where autonum_sc=" & sc
    DB.Execute sql
    
    
    If CargaSuzano Then
        sql = "update " & Esquema_Op & ".tb_integra_carga set estufado=0 where codbarra='" & Adodc_SAIDA.Recordset!COD_PRODUTO & "'"
        DB.Execute UCase(sql)
    End If
    
    
'    sql = "select autonum_sc from tb_saida_carga where autonum_patio=" & p
'    PRSet Rs, sql
'    If Rs.EOF Then
        sql = "update " & Esquema_Op & ".tb_patio set ef='E' where autonum_patio = " & p
        DB.Execute UCase(sql)
'    End If
'    Rs.Close
    
'    Adodc_carga.Refresh
    Adodc_SAIDA.Refresh
    Call Limpa_Dados

    Call Mostra_Estufagem

End Sub

Private Sub DataCombo1_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 13 Then SendKeys "{TAB}"

End Sub

Private Sub DataCombo1_LostFocus(Index As Integer)

If Index = 0 Then
    LbTalie = "0"
    If Nnull(DataCombo1(0).BoundText, 0) > 0 Then
        MskData(0) = Format(Now, "dd/mm/yyyy hh:mm")
        MskData(1) = "__/__/____ __:__"
        Check1.Value = 0
        Combo1.ListIndex = -1
        
        sql = "select autonum_talie,inicio,termino, crossdocking,conferente,equipe from " & Esquema_Op & ".tb_talie where autonum_patio=" & DataCombo1(0).BoundText & " and nvl(crossdocking,0)=0 order by autonum_talie desc"
        PRSet Rs, sql
        If Not Rs.EOF Then
            
            LbTalie = Rs!autonum_talie
'            Check1.Value = Nnull(Rs!crossdocking, 0)
            MskData(0) = Format(Rs!Inicio, "dd/mm/yyyy hh:mm")
            If Not IsNull(Rs!TERMINO) Then
                MskData(1) = Format(Rs!TERMINO, "dd/mm/yyyy hh:mm")
            End If
            DataCombo5.BoundText = Nnull(Rs!CONFERENTE, 0)
            DataCombo4.BoundText = Nnull(Rs!EQUIPE, 0)
            If Format(Rs!TERMINO, "dd/mm/yyyy") <> "" Then
                If MsgBox("Já consta Talie com fim de estufagem informada." & vbCr & "Deseja reabrir o conteiner e editar a estufagem ?", vbYesNo) = vbYes Then
                    MskData(1) = "__/__/____ __:__"
                    sql = "update " & Esquema_Op & ".tb_talie set termino=null where autonum_patio=" & DataCombo1(0).BoundText
                    DB.Execute sql
                    
                    sql = "update " & Esquema_Op & ".tb_patio set ef='E' where autonum_patio = " & DataCombo1(0).BoundText
                    DB.Execute sql
                End If
            End If
        End If
        Rs.Close
        
        Frame4.Enabled = True
        
                
                
        Adodc_SAIDA.ConnectionString = ConnectionString
        sql_grid_saida = " SELECT  cc.id_conteiner,D.REFERENCE, E.DESCRICAO_EMB, F.DESC_PRODUTO, A.*, B.AUTONUM_PCS, G.RAZAO, H.DESCRICAO_NAV || '--' || I.NUM_VIAGEM AS navio_viagem"
        sql_grid_saida = sql_grid_saida & ",NF.NUM_NF"
        sql_grid_saida = sql_grid_saida & " FROM " & Esquema_Op & ".TB_SAIDA_CARGA A, " & Esquema_Op & ".TB_PATIO_CS B, " & Esquema_Op & ".TB_BOOKING_CARGA C, "
        sql_grid_saida = sql_grid_saida & Esquema_Op & ".TB_BOOKING D, " & Esquema_Op & ".TB_CAD_EMBALAGENS E, " & Esquema_Op & ".TB_CAD_PRODUTOS F, "
        sql_grid_saida = sql_grid_saida & Esquema_Op & ".TB_CAD_PARCEIROS G, " & Esquema_Op & ".tb_cad_navios h, " & Esquema_Op & ".tb_viagens i, "
        sql_grid_saida = sql_grid_saida & Esquema_Op & ".TB_NOTAS_ITENS NI, " & Esquema_Op & ".tb_NOTAS_FISCAIS NF,"
        sql_grid_saida = sql_grid_saida & Esquema_Op & ".TB_PATIO CC"
        sql_grid_saida = sql_grid_saida & " WHERE A.AUTONUM_PCS = B.AUTONUM_PCS AND B.AUTONUM_BCG = C.AUTONUM_BCG AND C.AUTONUM_BOO = D.AUTONUM_BOO AND "
        sql_grid_saida = sql_grid_saida & " A.AUTONUM_EMB = E.AUTONUM_EMB(+) AND B.AUTONUM_PRO = f.AUTONUM_PRO And D.AUTONUM_PARCEIRO = G.AUTONUM And "
        sql_grid_saida = sql_grid_saida & " D.autonum_via = i.autonum_via And i.autonum_nav = h.autonum_nav"
        sql_grid_saida = sql_grid_saida & " AND A.AUTONUM_PATIO = CC.AUTONUM_PATIO(+)"
        sql_grid_saida = sql_grid_saida & " AND A.AUTONUM_NFI = NI.AUTONUM_NFI"
        sql_grid_saida = sql_grid_saida & " AND NI.AUTONUM_NF = NF.AUTONUM_NF"
        sql_grid_saida = sql_grid_saida & " and a.autonum_patio = " & DataCombo1(0).BoundText
        sql_grid_saida = sql_grid_saida & " and a.autonum_talie = " & LbTalie
        
        sql_grid_saida = UCase(sql_grid_saida)
        Adodc_SAIDA.RecordSource = UCase(sql_grid_saida)
        Adodc_SAIDA.Refresh
        
        
        Call Mostra_Estufagem
        
        
        If MskData(0).ClipText = "" Then
            MskData(0).SetFocus
        Else
            TxtProduto.SetFocus
        End If
        
    
    End If
End If

End Sub

Private Sub DADOS_CNTR()

If Nnull(DataCombo1(0).BoundText, 1) <> "" Then
    sql = "Select * from " & Esquema_Op & ".tb_dados_conteiner a, " & Esquema_Op & ".tb_cad_iso_cntr b where a.autonum_iso = b.autonum_iso and "
    sql = sql & " a.id_conteiner ='" & Nnull(DataCombo1(0).Text, 1) & "'"
    PRSet rs_dados_cntr, sql
    If rs_dados_cntr.RecordCount > 0 Then
        Text5.Text = Nnull(rs_dados_cntr!volume_nominal, 1)
        Text6.Text = Nnull(rs_dados_cntr!maxgross, 1)
    Else
        MsgBox "Cadastro do Conteiner Incompleto!", vbInformation
    End If
    'Frame3.Enabled = True
Else
    'Frame3.Enabled = False
End If

End Sub

Private Sub CALC_VOL()

vol = (Nnull(Text3(0).Text, 0) * Nnull(Text3(1).Text, 0)) * Nnull(Text3(2).Text, 0)
vol = Nnull(vol, 0) / 1000000
Text3(4).Text = vol
If Nnull(Txt_qtde.Text, 0) > 0 Then
    vol = Nnull(vol, 0) * Nnull(Txt_qtde.Text, 0)
    Text3(3).Text = Nnull(vol, 0)
End If

End Sub


Private Sub DataCombo2_Click(Area As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub DataCombo3_Click(Area As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

Adodc2.ConnectionString = ConnectionString
Adodc2.RecordSource = "Select autonum_eqp, nome_eqp from " & Esquema_Op & ".tb_equipe where flag_ativo=1 and flag_conferente=1 order by nome_eqp "
Adodc2.Refresh
Adodc2.Recordset.ActiveConnection = Nothing

Adodc3.ConnectionString = ConnectionString
Adodc3.RecordSource = "Select autonum_eqp, nome_eqp from " & Esquema_Op & ".tb_equipe where flag_ativo=1 and flag_operador=1  order by nome_eqp "
Adodc3.Refresh
Adodc3.Recordset.ActiveConnection = Nothing


Adodc_conteiner.ConnectionString = ConnectionString
Adodc_conteiner.RecordSource = TROCA_DBO(UCase(" Select autonum_patio as id , id_conteiner from " & Esquema_Op & ".tb_patio where flag_historico = 0 order by id_conteiner"))
Adodc_conteiner.Refresh

Adodc_cliente.ConnectionString = ConnectionString
Adodc_cliente.RecordSource = TROCA_DBO(UCase("Select autonum, fantasia from " & Esquema_Op & ".tb_cad_parceiros where flag_exportador = 1 order by fantasia "))
Adodc_cliente.Refresh

Adodc_emb.ConnectionString = ConnectionString
Adodc_emb.RecordSource = TROCA_DBO(UCase("Select descricao_emb as descr, autonum_emb as id from " & Esquema_Op & ".tb_cad_embalagens order by descricao_emb "))
Adodc_emb.Refresh

Adodc_navio.ConnectionString = ConnectionString
Adodc_navio.RecordSource = TROCA_DBO(UCase("Select a.Descricao_nav " & sql_conc & " '--' " & sql_conc & " b.num_viagem descr, b.autonum_via id " & _
                           " from " & Esquema_Op & ".tb_cad_navios a, " & Esquema_Op & ".tb_viagens b where a.autonum_nav = b.autonum_nav order by a.descricao_nav "))
Adodc_navio.Refresh

posiciona_tela Me

Controlar_Acesso_Forms Me, Item_menu
End Sub


Private Sub List1_Click()
'Autonum_NFI = List1.ItemData(Index)
End Sub

Private Sub Text3_LostFocus(Index As Integer)
If Not IsNumeric(Text3(Index).Text) Then
    MsgBox "Este campo só aceita valores numéricos!", vbInformation
    Text3(Index).Text = Empty
    Text3(Index).SetFocus
End If
If Index <> 5 Then
    CALC_VOL
End If

End Sub



Private Sub Txt_qtde_LostFocus()

If Val(Nnull(Txt_qtde.Text, 0)) > Val(Nnull(Text4.Text, 0)) Then
    MsgBox "Quantidade Indisponível!", vbInformation
    Txt_qtde.Text = Empty
    Txt_qtde.SetFocus
Else
    
    Text3(5).Text = Val(Txt_qtde) * Nnull(Adodc_carga.Recordset.Fields("bruto"), 0)
    CALC_VOL

End If

End Sub

Private Sub txtInstrucao_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtInstrucao_LostFocus()
txtInstrucao.Text = UCase(txtInstrucao.Text)
End Sub

Private Sub TxtProduto_LostFocus()

    
    If TxtProduto <> "" Then
        Call Busca_Carga
        Txt_qtde.SetFocus
    End If

End Sub


Private Sub txtReserva_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtReserva_LostFocus()
txtreserva.Text = UCase(txtreserva.Text)
End Sub

