VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form Reefer 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3585
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
   ScaleHeight     =   5010
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   4980
      Left            =   0
      TabIndex        =   8
      Top             =   30
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   8784
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   14737632
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Monitoramento"
      TabPicture(0)   =   "ReeferS.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblRdx"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblOp"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblIPA"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label9"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblReeferDesligado"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblPlugOFF"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label8"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label10"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Grid2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "CNTR"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "MskCntr"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Adodc6"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Command2"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtGateIN"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtVentilacao"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtUmidade"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtEscala"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtTempRet"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtTempSup"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtVentilacaoSet"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtUmidadeSet"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtSetPoint"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Command1"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Command4"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtMIN"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtMAX"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).ControlCount=   33
      TabCaption(1)   =   "Previsão"
      TabPicture(1)   =   "ReeferS.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Check3"
      Tab(1).Control(1)=   "Check2"
      Tab(1).Control(2)=   "Check1"
      Tab(1).Control(3)=   "GridSaidas"
      Tab(1).Control(4)=   "GridEntradas"
      Tab(1).Control(5)=   "AdodcEntradas"
      Tab(1).Control(6)=   "AdodcSaidas"
      Tab(1).Control(7)=   "Shape1"
      Tab(1).Control(8)=   "Label1(1)"
      Tab(1).Control(9)=   "Label1(2)"
      Tab(1).Control(10)=   "lblTotEnt"
      Tab(1).Control(11)=   "lblTotSai"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Un. Desligadas"
      TabPicture(2)   =   "ReeferS.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GridDesligados"
      Tab(2).Control(1)=   "AdodcDesligados"
      Tab(2).Control(2)=   "lblTotDesl"
      Tab(2).Control(3)=   "Label13"
      Tab(2).ControlCount=   4
      Begin VB.TextBox txtMAX 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1740
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1080
         Width           =   750
      End
      Begin VB.TextBox txtMIN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   930
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1080
         Width           =   750
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         Caption         =   "Posicionados"
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
         Height          =   225
         Left            =   -72360
         TabIndex        =   39
         Top             =   4620
         Width           =   1005
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         Caption         =   "Com agendamento Saida"
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
         Height          =   255
         Left            =   -73740
         TabIndex        =   36
         Top             =   4650
         Width           =   1395
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Apenas os Desligados"
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
         Height          =   255
         Left            =   -74700
         TabIndex        =   35
         Top             =   4650
         Width           =   1275
      End
      Begin VB.CommandButton Command4 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2670
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   540
         Width           =   315
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SALVAR"
         Height          =   465
         Left            =   2100
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2610
         Width           =   1350
      End
      Begin VB.TextBox txtSetPoint 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1080
         Width           =   750
      End
      Begin VB.TextBox txtUmidadeSet 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2130
         Width           =   900
      End
      Begin VB.TextBox txtVentilacaoSet 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2700
         Width           =   900
      End
      Begin VB.TextBox txtTempSup 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1560
         Width           =   750
      End
      Begin VB.TextBox txtTempRet 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   900
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1560
         Width           =   780
      End
      Begin VB.TextBox txtEscala 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1710
         MaxLength       =   1
         TabIndex        =   4
         Top             =   1560
         Width           =   330
      End
      Begin VB.TextBox txtUmidade 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1020
         MaxLength       =   8
         TabIndex        =   5
         Top             =   2130
         Width           =   1020
      End
      Begin VB.TextBox txtVentilacao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1050
         MaxLength       =   8
         TabIndex        =   6
         Top             =   2700
         Width           =   990
      End
      Begin VB.TextBox txtGateIN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2130
         Width           =   1320
      End
      Begin VB.CommandButton Command2 
         Caption         =   "PLUG OFF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2250
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1560
         Width           =   960
      End
      Begin MSAdodcLib.Adodc Adodc6 
         Height          =   345
         Left            =   -90
         Top             =   7230
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   609
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
         Caption         =   "Adodc6"
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
      Begin MSMask.MaskEdBox MskCntr 
         Height          =   330
         Left            =   885
         TabIndex        =   1
         Top             =   555
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">????######-#"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox CNTR 
         Height          =   315
         Left            =   105
         TabIndex        =   0
         Top             =   555
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###-#"
         PromptChar      =   "_"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
         Bindings        =   "ReeferS.frx":0054
         Height          =   1815
         Left            =   90
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3090
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3201
         _Version        =   393216
         Cols            =   8
         RowHeightMin    =   300
         WordWrap        =   -1  'True
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "   |<Data               |^Temp. SUP|^Temp. RET|^Umidade  |^Ventilacao|<Usuario                           |"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   8
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSaidas 
         Bindings        =   "ReeferS.frx":0069
         Height          =   1815
         Left            =   -74940
         TabIndex        =   16
         Top             =   2790
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   3201
         _Version        =   393216
         Cols            =   11
         RowHeightMin    =   420
         WordWrap        =   -1  'True
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "   |<Conteiner        |^Gate IN  |^Yard     |^Prev.Saida|^SetPoint|^1ª Medic.|^Data       |<Ult. Medic.|^Data         | "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   11
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridEntradas 
         Bindings        =   "ReeferS.frx":0083
         Height          =   1935
         Left            =   -74970
         TabIndex        =   17
         Top             =   600
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   3413
         _Version        =   393216
         Cols            =   11
         RowHeightMin    =   300
         WordWrap        =   -1  'True
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "   |<Conteiner        |^Tam |<Set Point|^Atracação        |<Local Atracacao|<Entrada Prevista  "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   11
      End
      Begin MSAdodcLib.Adodc AdodcEntradas 
         Height          =   345
         Left            =   -74670
         Top             =   2100
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   609
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
         Caption         =   "Adodc6"
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
      Begin MSAdodcLib.Adodc AdodcSaidas 
         Height          =   345
         Left            =   -68610
         Top             =   7650
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   609
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
         Caption         =   "Adodc6"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridDesligados 
         Bindings        =   "ReeferS.frx":009F
         Height          =   4005
         Left            =   -74880
         TabIndex        =   37
         Top             =   450
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   7064
         _Version        =   393216
         Cols            =   11
         RowHeightMin    =   420
         WordWrap        =   -1  'True
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "   |<Conteiner        |^Gate IN  |^Yard     |^Prev.Saida|^SetPoint|^1ª Medic.|^Data       |<Ult. Medic.|^Data         | "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   11
      End
      Begin MSAdodcLib.Adodc AdodcDesligados 
         Height          =   345
         Left            =   -72540
         Top             =   270
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   609
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
         Caption         =   "Adodc6"
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MAX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   1980
         TabIndex        =   44
         Top             =   900
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MIN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   1170
         TabIndex        =   43
         Top             =   900
         Width           =   255
      End
      Begin VB.Label lblPlugOFF 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   3060
         TabIndex        =   40
         Top             =   8040
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label lblTotDesl 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   -72300
         TabIndex        =   38
         Top             =   4590
         Width           =   705
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H008080FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   -74910
         Top             =   4650
         Width           =   165
      End
      Begin VB.Label lblReeferDesligado 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REEFER DESLIGADO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   990
         TabIndex        =   34
         Top             =   330
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   33
         Top             =   330
         Width           =   6795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " CONTEINER"
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
         Height          =   195
         Index           =   0
         Left            =   -15
         TabIndex        =   32
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SET POINT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   90
         TabIndex        =   31
         Top             =   900
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UMIDADE (%)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   90
         TabIndex        =   30
         Top             =   1950
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENTILACAO m3/h"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   29
         Top             =   2520
         Width           =   1485
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TEMP SUP."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   90
         TabIndex        =   28
         Top             =   1410
         Width           =   645
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TEMP RET."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   990
         TabIndex        =   27
         Top             =   1410
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ESCALA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   1740
         TabIndex        =   26
         Top             =   1410
         Width           =   450
      End
      Begin VB.Label lblIPA 
         BackColor       =   &H0080FFFF&
         Caption         =   "0"
         Height          =   315
         Left            =   2460
         TabIndex        =   25
         Top             =   3780
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblOp 
         BackColor       =   &H0080FFFF&
         Caption         =   "0"
         Height          =   315
         Left            =   6240
         TabIndex        =   24
         Top             =   4350
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblRdx 
         BackColor       =   &H0080FFFF&
         Caption         =   "0"
         Height          =   315
         Left            =   6240
         TabIndex        =   23
         Top             =   4380
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GATE IN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   2310
         TabIndex        =   22
         Top             =   1950
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UNIDADES PREVISTAS - ENTRADA PATIO"
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
         Left            =   -74910
         TabIndex        =   21
         Top             =   390
         Width           =   2430
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UNIDADES EM ESTOQUE "
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
         Left            =   -74910
         TabIndex        =   20
         Top             =   2610
         Width           =   1485
      End
      Begin VB.Label lblTotEnt 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   -72120
         TabIndex        =   19
         Top             =   360
         Width           =   405
      End
      Begin VB.Label lblTotSai 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   225
         Left            =   -72120
         TabIndex        =   18
         Top             =   2580
         Width           =   435
      End
   End
End
Attribute VB_Name = "Reefer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
Carrega_Grid_Saidas_Previstas

End Sub

Private Sub Check2_Click()
Carrega_Grid_Saidas_Previstas
End Sub

Private Sub Check3_Click()
Carrega_Grid_Saidas_Previstas
End Sub

Private Sub CNTR_Change()
If InStr(1, Me.CNTR, "_") = 0 Then
    Call Busca_Cntr
End If
End Sub

Private Sub Command1_Click()
If Valida_Dados Then
    SQl = "INSERT INTO OPERADOR..TB_MONITORING ("
    SQl = SQl & " [DT_MONITORING],"
    SQl = SQl & " [ID_CONTEINER],"
    SQl = SQl & " [AUTONUM_CNTR_IPA],"
    SQl = SQl & " [AUTONUM_CNTR_RDX],"
    SQl = SQl & " [AUTONUM_CNTR_OP],"
    SQl = SQl & " [TEMP_SETPOINT],"
    SQl = SQl & " [UMIDADE_SETPOINT],"
    SQl = SQl & " [VENTILACAO_SETPOINT],"
    SQl = SQl & " [TEMP_SUP],"
    SQl = SQl & " [TEMP_RET],"
    SQl = SQl & " [VENTILACAO],"
    SQl = SQl & " [UMIDADE],"
    SQl = SQl & " [ESCALA],"
    SQl = SQl & " [PATIO],"
    SQl = SQl & " [USUARIO])"
    SQl = SQl & " VALUES ("
    SQl = SQl & " GETDATE(),"
    SQl = SQl & "'" & Me.MskCntr & "',"
    SQl = SQl & Me.lblIPA.Caption & ","
    SQl = SQl & Me.lblRdx.Caption & ","
    SQl = SQl & Me.lblOp.Caption & ","
    SQl = SQl & "'" & Me.txtSetPoint.text & "',"
    SQl = SQl & "'" & Me.txtUmidadeSet.text & "',"
    SQl = SQl & "'" & Me.txtVentilacaoSet.text & "',"
    SQl = SQl & "'" & Me.txtTempSup.text & "',"
    SQl = SQl & "'" & Me.txtTempRet.text & "',"
    SQl = SQl & "'" & Me.txtVentilacao.text & "',"
    SQl = SQl & "'" & Me.txtUmidade.text & "',"
    SQl = SQl & "'" & Me.txtEscala.text & "',"
    SQl = SQl & IIf(Col_Patio = 7, 1, Col_Patio) & ","
    SQl = SQl & Cod_Usuario
    SQl = SQl & " )"
    db.Execute SQl
    
    If Val(Me.lblIPA.Caption) > 0 Then
        SQl = "UPDATE SGIPA..TB_CNTR_BL SET DT_PLUG_OFF=NULL WHERE AUTONUM=" & Val(Me.lblIPA.Caption)
        db.Execute SQl
        Me.lblReeferDesligado.Visible = False
    End If
    
    If Val(Me.lblRdx.Caption) > 0 Then
        SQl = "UPDATE REDEX..TB_PATIO SET DT_PLUG_OFF=NULL WHERE AUTONUM_PATIO=" & Val(Me.lblRdx.Caption)
        db.Execute SQl
        Me.lblReeferDesligado.Visible = False
    End If
    
    
    Carrega_Grid1
    Me.txtTempSup = ""
    Me.txtTempRet = ""
    Me.txtUmidade = ""
    Me.txtVentilacao = ""
    Me.lblPlugOFF.Caption = ""
    Me.lblPlugOFF.Visible = False
    Me.CNTR.SetFocus
End If
End Sub


Sub Busca_Dados()

idPatio = 0
If MskCntr.ClipText = "" Then Exit Sub
MskCntr.Tag = ""

Limpa_Dados
'GoTo pula
SQl = "select autonum,  data_ent_temp, sistema, temperature, scale, umidade, ventilacao,flag_reefer, flag_reefer_desligado "
SQl = SQl & " from " & Banco_Operador & "VW_INVENT_SISTEMAS where PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio) & " and id_conteiner='" & MskCntr & "'"
SQl = SQl & " AND FLAG_REEFER='Y' AND FLAG_REEFER_DESLIGADO='N'"
Set tbvw = New ADODB.Recordset
SQl = CONVERTE_SQL(SQl)
tbvw.Open SQl, db, 1, 1

If Not tbvw.EOF Then

    
    If Nnull(tbvw!flag_reefer, 1) = "N" Then
        Frmmensagem.Label1.Caption = "Contêiner não é reefer "
        Frmmensagem.Show vbModal
        'MsgBox "Unidade não encontrada no estoque"
        MskCntr = "__________-_"
        MskCntr.SetFocus
        Exit Sub

    End If

    If Nnull(tbvw!flag_reefer_desligado, 1) = "Y" Then
        Frmmensagem.Label1.Caption = "Contêiner reefer captado com carga DRY!"
        Frmmensagem.Show vbModal
        'MsgBox "Unidade não encontrada no estoque"
        MskCntr = "__________-_"
        MskCntr.SetFocus
        Exit Sub

    End If
    
    idPatio = tbvw!autonum
        
    If tbvw!Sistema = "I" Then
        Me.lblIPA = idPatio
        Dim PO$
        
        Dim tbCntrBl As New ADODB.Recordset
        SQl = "SELECT rtrim(convert(char,dt_plug_off,103)) + ' ' + rtrim(convert(char,dt_plug_off,108)) AS QUAl,tempmin,tempmax FROM SGIPA..TB_CNTR_BL WHERE AUTONUM=" & idPatio
        tbCntrBl.Open SQl, db, 1, 1
        PO$ = Nnull(tbCntrBl!Qual, 1)
        Me.txtMax = Nnull(tbCntrBl!tempmax, 1)
        Me.txtMin = Nnull(tbCntrBl!tempmin, 1)
        'PO$ = Trim(Nnull(db.Execute("SELECT rtrim(convert(char,dt_plug_off,103)) + ' ' + rtrim(convert(char,dt_plug_off,108)) AS QUAl FROM SGIPA..TB_CNTR_BL WHERE AUTONUM=" & idPatio)(0), 1))
        If PO$ <> "" Then
            Me.lblReeferDesligado.Visible = True
            Me.lblPlugOFF.Caption = "PLUG OFF:" & PO$
            Me.lblPlugOFF.Visible = True
        Else
            Me.lblReeferDesligado.Visible = False
            Me.lblPlugOFF.Caption = ""
            Me.lblPlugOFF.Visible = False
        End If
    ElseIf tbvw!Sistema = "O" Then
        Me.lblOp = idPatio
    ElseIf tbvw!Sistema = "R" Then
        Me.lblRdx = idPatio
    End If
    Me.txtGateIN = Nnull(tbvw!data_ent_temp, 1)
    
    Me.txtEscala.text = Nnull(tbvw!Scale, 1)
    Me.txtSetPoint.text = Nnull(tbvw!TEMPERATURE, 1)
    If Nnull(tbvw!Umidade, 1) <> "" Then
        Me.txtUmidadeSet.text = Nnull(tbvw!Umidade, 1)
    Else
        Me.txtUmidadeSet.text = "OFF"
    End If
    If Nnull(tbvw!Ventilacao, 1) <> "" Then
        Me.txtVentilacaoSet.text = Nnull(tbvw!Ventilacao, 1)
    Else
        Me.txtUmidadeSet.text = "CLOSED"
    End If
    
    Me.txtTempRet.text = ""
    Me.txtTempSup.text = ""
    
    Me.txtUmidade.text = ""
    Me.txtVentilacao.text = ""
    
    
    
Else

    Frmmensagem.Label1.Caption = "Contêiner não encontrado "
    Frmmensagem.Show vbModal
    'MsgBox "Unidade não encontrada no estoque"
    MskCntr = "__________-_"
    MskCntr.SetFocus
    Exit Sub

End If
tbvw.Close

    
Carrega_Grid1


Me.txtTempSup.SetFocus

End Sub

Private Sub Carrega_Grid1()
Grid2.Clear
Dim SQl As String

If Me.lblIPA.Caption <> "0" Or Me.lblRdx.Caption <> "0" Or Me.lblOp.Caption <> "0" Then

    SQl = "SELECT "
    SQl = SQl & " LTRIM(LEFT(CONVERT(CHAR,A.DT_MONITORING,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,A.DT_MONITORING,108),5))"
    SQl = SQl & " ,A.TEMP_SUP + ' ' + A.ESCALA"
    SQl = SQl & " ,A.TEMP_RET + ' ' + A.ESCALA"
    SQl = SQl & " ,A.UMIDADE "
    SQl = SQl & " ,A.VENTILACAO "
    SQl = SQl & " ,B.USUARIO "
    SQl = SQl & " ,A.AUTONUM"
    SQl = SQl & " FROM OPERADOR..TB_MONITORING A "
    SQl = SQl & " LEFT JOIN SGIPA..TB_CAD_USUARIOS B ON A.USUARIO=B.AUTONUM "
    SQl = SQl & " WHERE 0=0 "
    If Me.lblIPA.Caption <> "0" Then
        SQl = SQl & " AND A.AUTONUM_CNTR_IPA=" & Me.lblIPA.Caption
    ElseIf Me.lblRdx.Caption <> "0" Then
        SQl = SQl & " AND A.AUTONUM_CNTR_RDX=" & Me.lblRdx.Caption
    ElseIf Me.lblOp.Caption <> "0" Then
        SQl = SQl & " AND A.AUTONUM_CNTR_OP=" & Me.lblOp.Caption
    End If
    SQl = SQl & " ORDER BY A.AUTONUM DESC "
    
    Adodc6.ConnectionString = ConnectionStringAdo
    Adodc6.RecordSource = SQl
    Adodc6.Refresh
    Set Adodc6.Recordset.ActiveConnection = Nothing
End If

Grid2.FormatString = "   |<Data              |^Temp. SUP|^Temp. RET|^Umidade  |^Ventilacao|<Usuario                           |"

End Sub

Private Sub Carrega_Grid_Entradas_Previstas()
Me.lblTotEnt.Caption = "0"

Dim SQl As String
SQl = " SELECT"
SQl = SQl & " C.ID_CONTEINER,"
SQl = SQl & " C.TAMANHO,"
SQl = SQl & " C.TEMP + ' ' + ISNULL(C.ESCALA,'') AS SETPOINT,"
SQl = SQl & " REPLACE(CASE V.DT_ATRACACAO  WHEN NULL THEN LEFT(CONVERT(CHAR,V.DT_ATRACACAOP,103),10) ELSE  LEFT(CONVERT(CHAR,V.DT_ATRACACAO,103),10) END,'01/01/1900','') AS DT_ATRACACAO,"
SQl = SQl & " CASE ISNULL(A.DESCR,'-') WHEN '-' THEN AP.DESCR ELSE A.DESCR END AS ARM_ATRACACAO,"
SQl = SQl & " LTRIM(LEFT(CONVERT(CHAR,T.HR_ENTRADA_PREV,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,T.HR_ENTRADA_PREV,108),5)) AS ENTRADA_PREVISTA"
SQl = SQl & " From"
SQl = SQl & " SGIPA..TB_CNTR_BL C INNER JOIN"
SQl = SQl & " SGIPA..DTE_TB_VIAGENS V ON C.VIAGEM=V.VIAGEM LEFT JOIN"
SQl = SQl & " SGIPA..DTE_TB_ARMAZENS AP ON V.ARM_ATRACACAOP=AP.CODE LEFT JOIN"
SQl = SQl & " SGIPA..DTE_TB_ARMAZENS A ON V.ARM_ATRACACAO=A.CODE LEFT JOIN"
SQl = SQl & " SGIPA..TB_TTIMEUNT T ON C.ID_CONTEINER=T.SIGLA AND C.GMCI=T.GMCI"
SQl = SQl & " Where c.FLAG_TERMINAL = 0 And c.GMCI Is Not Null"
SQl = SQl & " AND C.PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio)
SQl = SQl & " AND ISNULL(C.TEMP,' ')<>' '"
SQl = SQl & " and isnull(c.flag_desligado,0)=0"
SQl = SQl & " order by c.id_conteiner "


AdodcEntradas.ConnectionString = ConnectionStringAdo
AdodcEntradas.RecordSource = SQl
AdodcEntradas.Refresh
Me.lblTotEnt.Caption = Me.AdodcEntradas.Recordset.RecordCount

Set AdodcEntradas.Recordset.ActiveConnection = Nothing


GridEntradas.FormatString = "  |<Conteiner           |^Tam |^Set Point|^Atracação        |<Local Atracacao|<Entrada Prevista  "



End Sub

Private Sub Carrega_Grid_Saidas_Previstas()

Me.lblTotSai.Caption = "0"

Dim SQl As String
SQl = " SELECT"
SQl = SQl & " C.ID_CONTEINER,"
SQl = SQl & " LTRIM(LEFT(CONVERT(CHAR,C.DT_ENTRADA,103),10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,C.DT_ENTRADA,108),5)) AS DT_ENTRADA,"
SQl = SQl & " C.YARD,"
SQl = SQl & " LTRIM(LEFT(CONVERT(CHAR,GD.PERIODO_INICIAL,103),10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,GD.PERIODO_INICIAL,108),5))  PERIODO_INICIAL,"
SQl = SQl & " C.TEMP + ' ' + ISNULL(C.ESCALA,'') AS SETPOINT,"
SQl = SQl & " M1.TEMP_RET,"
SQl = SQl & " LTRIM(LEFT(CONVERT(CHAR,M1.DT_MONITORING,103),10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,M1.DT_MONITORING,108),5)) AS DT1 ,"
SQl = SQl & " M2.TEMP_RET,"
SQl = SQl & " LTRIM(LEFT(CONVERT(CHAR,M2.DT_MONITORING,103),10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,M2.DT_MONITORING,108),5)) AS DT2, "
SQl = SQl & " c.dt_plug_off "


SQl = SQl & " From"
SQl = SQl & " SGIPA..TB_CNTR_BL C LEFT JOIN"
SQl = SQl & " (SELECT MIN(AUTONUM) AS A_PRIM ,AUTONUM_CNTR_IPA  FROM OPERADOR..TB_MONITORING GROUP BY AUTONUM_CNTR_IPA ) PRIM"
SQl = SQl & " ON C.AUTONUM=PRIM.AUTONUM_CNTR_IPA LEFT JOIN"
SQl = SQl & " OPERADOR..TB_MONITORING M1 ON PRIM.A_PRIM=M1.AUTONUM"
SQl = SQl & " Left Join"
SQl = SQl & " (SELECT MAX(AUTONUM) AS A_ULT ,AUTONUM_CNTR_IPA  FROM OPERADOR..TB_MONITORING GROUP BY AUTONUM_CNTR_IPA ) ULT"
SQl = SQl & " ON C.AUTONUM=ULT.AUTONUM_CNTR_IPA LEFT JOIN"
SQl = SQl & " OPERADOR..TB_MONITORING M2 ON ULT.A_ULT=M2.AUTONUM LEFT JOIN"
SQl = SQl & " OPERADOR..TB_GD_RESERVA GD ON C.AUTONUM_GD_RESERVA=GD.AUTONUM_GD_RESERVA"
SQl = SQl & " Where"
SQl = SQl & " c.Patio  " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio)
SQl = SQl & " AND C.FLAG_TERMINAL=1 AND C.FLAG_HISTORICO=0 AND C.DT_ENTRADA IS NOT NULL AND C.FLAG_DESLIGADO=0"
SQl = SQl & " AND ISNULL(C.TEMP,' ')<>' '"
If Me.Check1.Value = 1 Then
    SQl = SQl & " and c.dt_plug_off is not null"
End If
If Me.Check2.Value = 1 Then
    SQl = SQl & " and GD.PERIODO_INICIAL is not null"
End If

If Me.Check3.Value = 1 Then
    SQl = SQl & " and (c.yard LIKE 'CF%' OR C.YARD LIKE 'CONF%') "
End If

SQl = SQl & " ORDER BY C.ID_CONTEINER "

AdodcSaidas.ConnectionString = ConnectionStringAdo
AdodcSaidas.RecordSource = SQl
AdodcSaidas.Refresh
Me.lblTotSai.Caption = Me.AdodcSaidas.Recordset.RecordCount

Set AdodcSaidas.Recordset.ActiveConnection = Nothing

Me.GridSaidas.FormatString = "  |<Conteiner           |^Gate IN       |^Yard     |^Prev.Saida   |^SetPoint|^1ª Medic.|^Data            |^Ult. Medic.|^Data             |"
Me.GridSaidas.COLWIDTH(Me.GridSaidas.Cols - 1) = 0

For i = 1 To GridSaidas.Rows - 1
    If GridSaidas.TextMatrix(i, GridSaidas.Cols - 1) <> "" Then
        GridSaidas.Row = i
        GridSaidas.col = 1
        GridSaidas.CellBackColor = Me.Shape1.FillColor
    End If
    
  
Next i


End Sub

Private Sub Carrega_Grid_Desligados()

Me.lblTotDesl.Caption = "0"

Dim SQl As String
SQl = " SELECT"
SQl = SQl & " C.ID_CONTEINER,"
SQl = SQl & " LTRIM(LEFT(CONVERT(CHAR,C.DT_ENTRADA,103),10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,C.DT_ENTRADA,108),5)) AS DT_ENTRADA,"
SQl = SQl & " C.YARD,"
SQl = SQl & " LTRIM(LEFT(CONVERT(CHAR,GD.PERIODO_INICIAL,103),10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,GD.PERIODO_INICIAL,108),5))  PERIODO_INICIAL,"
SQl = SQl & " C.TEMP + ' ' + ISNULL(C.ESCALA,'') AS SETPOINT,"
SQl = SQl & " M1.TEMP_RET,"
SQl = SQl & " LTRIM(LEFT(CONVERT(CHAR,M1.DT_MONITORING,103),10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,M1.DT_MONITORING,108),5)) AS DT1 ,"
SQl = SQl & " M2.TEMP_RET,"
SQl = SQl & " LTRIM(LEFT(CONVERT(CHAR,M2.DT_MONITORING,103),10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,M2.DT_MONITORING,108),5)) AS DT2, "
SQl = SQl & " LTRIM(LEFT(CONVERT(CHAR,c.DT_plug_off,103),10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,c.dt_plug_off,108),5)) "


SQl = SQl & " From"
SQl = SQl & " SGIPA..TB_CNTR_BL C LEFT JOIN"
SQl = SQl & " (SELECT MIN(AUTONUM) AS A_PRIM ,AUTONUM_CNTR_IPA  FROM OPERADOR..TB_MONITORING GROUP BY AUTONUM_CNTR_IPA ) PRIM"
SQl = SQl & " ON C.AUTONUM=PRIM.AUTONUM_CNTR_IPA LEFT JOIN"
SQl = SQl & " OPERADOR..TB_MONITORING M1 ON PRIM.A_PRIM=M1.AUTONUM"
SQl = SQl & " Left Join"
SQl = SQl & " (SELECT MAX(AUTONUM) AS A_ULT ,AUTONUM_CNTR_IPA  FROM OPERADOR..TB_MONITORING GROUP BY AUTONUM_CNTR_IPA ) ULT"
SQl = SQl & " ON C.AUTONUM=ULT.AUTONUM_CNTR_IPA LEFT JOIN"
SQl = SQl & " OPERADOR..TB_MONITORING M2 ON ULT.A_ULT=M2.AUTONUM LEFT JOIN"
SQl = SQl & " OPERADOR..TB_GD_RESERVA GD ON C.AUTONUM_GD_RESERVA=GD.AUTONUM_GD_RESERVA"
SQl = SQl & " Where"
SQl = SQl & " c.Patio " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio)
SQl = SQl & " AND C.FLAG_TERMINAL=1 AND C.FLAG_HISTORICO=0 AND C.DT_ENTRADA IS NOT NULL AND C.FLAG_DESLIGADO=0"
SQl = SQl & " AND ISNULL(C.TEMP,' ')<>' '"
SQl = SQl & " and c.dt_plug_off is not null"
SQl = SQl & " ORDER BY c.dt_plug_off desc "

AdodcDesligados.ConnectionString = ConnectionStringAdo
AdodcDesligados.RecordSource = SQl
AdodcDesligados.Refresh
Me.lblTotDesl.Caption = Me.AdodcDesligados.Recordset.RecordCount

Set AdodcDesligados.Recordset.ActiveConnection = Nothing

Me.GridDesligados.FormatString = "  |<Conteiner           |^Gate IN       |^Yard     |^Prev.Saida   |^SetPoint|^1ª Medic.|^Data            |^Ult. Medic.|^Data             |^Data Plug OFF  "




End Sub


Sub Busca_Cntr()

    If CNTR.ClipText = "" Then Exit Sub
    MskCntr = "__________-_"
    
    Dim Regs As Integer
    SQl = "SELECT COUNT(1) FROM " & Banco_Operador & "VW_INVENT_SISTEMAS WHERE PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio) & " AND SUBSTR(ID_CONTEINER,8,5)='" & CNTR & "' and FLAG_REEFER='Y' AND FLAG_REEFER_DESLIGADO='N' "
    SQl = CONVERTE_SQL(SQl)
    Regs = Nnull(db.Execute(SQl).Fields(0), 0)
    
    If Regs = 0 Then
        'MsgBox "Nenhum contêiner com este final", vbExclamation
        Frmmensagem.Label1.Caption = "Nenhum contêiner Reefer com este final"
        Frmmensagem.Show vbModal

        CNTR = "___-_"
        CNTR.SetFocus
        Exit Sub
    End If
    
    If Regs = 1 Then
        
        SQl = "select id_conteiner from " & Banco_Operador & "VW_INVENT_SISTEMAS where PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio) & " and SUBSTR(ID_CONTEINER,8,5)='" & CNTR & "'"
        SQl = CONVERTE_SQL(SQl)
        MskCntr = Nnull(db.Execute(SQl).Fields(0), 1)
    
    Else
        
        Frmmensagem.Label1.Caption = "Existe mais de 1  conteiner reefer com este final. Favor informar a identificação completa"
        Frmmensagem.Show vbModal
        CNTR = "___-_"
        MskCntr.SetFocus
        Exit Sub
    
    End If
    
    CNTR = "___-_"
    
    If MskCntr.ClipText <> "" Then Call Busca_Dados
    

End Sub


Private Sub Command2_Click()
If Me.lblReeferDesligado.Visible = False Then
    If Val(Me.lblIPA.Caption) > 0 Then
        SQl = "UPDATE SGIPA..TB_CNTR_BL SET DT_PLUG_OFF=GETDATE() WHERE AUTONUM=" & Val(Me.lblIPA.Caption)
        db.Execute SQl
        Busca_Dados
    End If
    
End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyAscii = 13 Then Sendkeys "{tab}"
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub MskCntr_LostFocus()
Call Busca_Dados
End Sub

Private Sub Limpa_Monitora()
    Limpa_Dados
    CNTR = "___-_"
    Me.MskCntr = "__________-_"
    CNTR.SetFocus
End Sub


Private Sub Limpa_Dados()

Me.Grid2.Clear
lblReeferDesligado.Visible = False
Me.txtEscala.text = ""
Me.txtSetPoint.text = ""
Me.txtTempRet.text = ""
Me.txtTempSup.text = ""
Me.txtUmidade.text = ""
Me.txtUmidadeSet.text = ""
Me.txtVentilacao.text = ""
Me.txtVentilacaoSet.text = ""
Me.lblIPA.Caption = "0"
Me.lblOp.Caption = "0"
Me.lblRdx.Caption = "0"
Me.txtGateIN = ""
Me.txtMax = ""
Me.txtMin = ""


End Sub


Function Valida_Dados() As Boolean

If Me.lblIPA.Caption = "0" And Me.lblOp.Caption = "0" And Me.lblRdx.Caption = "0" Then
    MsgBox "Contêiner inválido", vbExclamation
    Valida_Dados = False
    Frmmensagem.Label1.Caption = "Contêiner inválido"
    Frmmensagem.Show vbModal
    Exit Function
End If

If Me.txtTempSup.text = "" Then
    Valida_Dados = False
    Frmmensagem.Label1.Caption = "Temperatura fornecida não informada"
    Frmmensagem.Show vbModal
    Exit Function
End If

If Me.txtTempRet.text = "" Then
    Valida_Dados = False
    Frmmensagem.Label1.Caption = "Temperatura de retorno não informada"
    Frmmensagem.Show vbModal
    Exit Function
End If

If Me.txtUmidade.text = "" Then
    Valida_Dados = False
    Frmmensagem.Label1.Caption = "Umidade não informada"
    Frmmensagem.Show vbModal
    Exit Function
End If

If Me.txtVentilacao.text = "" Then
    Valida_Dados = False
    Frmmensagem.Label1.Caption = "Ventilacao não informada"
    Frmmensagem.Show vbModal
    Exit Function
End If

If Me.txtEscala.text = "" Then
    Valida_Dados = False
    Frmmensagem.Label1.Caption = "Escala não informada"
    Frmmensagem.Show vbModal
    Exit Function
End If

If Abs(Val(Trim(Me.txtTempSup.text)) - Val(Trim(Me.txtSetPoint.text))) > 2 Then
    Col_Mensagem = "Atenção. Temperatura fornecida(Supply) muito divergente da temperatura de Set Point"
    frmMensagem2.Show 1
    If Col_Cancelado = True Then
        Valida_Dados = False
        Exit Function
    End If
End If
Valida_Dados = True

End Function

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 1 Then
    Carrega_Grid_Entradas_Previstas
    Carrega_Grid_Saidas_Previstas
    
End If
If SSTab1.Tab = 2 Then
    Carrega_Grid_Desligados
End If
End Sub

Private Sub txtEscala_LostFocus()
txtEscala = UCase(txtEscala)
If txtEscala <> "" Then
    If txtEscala <> "F" And txtEscala <> "C" Then txtEscala = ""
End If
End Sub

Private Sub txtTempRet_LostFocus()
txtTempRet = UCase(txtTempRet)
txtTempRet = Consiste_Temp(txtTempRet)
End Sub

Private Sub txtTempSup_LostFocus()
txtTempSup = UCase(txtTempSup)
txtTempSup = Consiste_Temp(txtTempSup)
End Sub

Private Sub txtUmidade_LostFocus()
txtUmidade = UCase(txtUmidade)
txtUmidade = Consiste_Umidade(txtUmidade)

End Sub

Private Sub txtVentilacao_LostFocus()
txtVentilacao = UCase(txtVentilacao)
txtVentilacao = Consiste_Ventilacao(txtVentilacao)
End Sub
