VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Reefer 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   8490
      Left            =   30
      TabIndex        =   8
      Top             =   30
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   14975
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   14737632
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Monitoramento"
      TabPicture(0)   =   "Reefer.frx":0000
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
      Tab(0).Control(13)=   "Grid2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "CNTR"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "MskCntr"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Adodc6"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Command2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtGateIN"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtVentilacao"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtUmidade"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtEscala"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtTempRet"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtTempSup"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtVentilacaoSet"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtUmidadeSet"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtSetPoint"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Command1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Command4"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "Previsão"
      TabPicture(1)   =   "Reefer.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Check3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Check2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Check1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "GridSaidas"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "GridEntradas"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "AdodcEntradas"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "AdodcSaidas"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Shape1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label1(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(2)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lblTotEnt"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lblTotSai"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Un. Desligadas"
      TabPicture(2)   =   "Reefer.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label13"
      Tab(2).Control(1)=   "lblTotDesl"
      Tab(2).Control(2)=   "AdodcDesligados"
      Tab(2).Control(3)=   "GridDesligados"
      Tab(2).ControlCount=   4
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         Caption         =   "Posicionados"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -69720
         TabIndex        =   39
         Top             =   8160
         Width           =   1755
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         Caption         =   "Com agendamento Saida"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -72420
         TabIndex        =   36
         Top             =   8160
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Apenas os Desligados"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74580
         TabIndex        =   35
         Top             =   8160
         Width           =   1995
      End
      Begin VB.CommandButton Command4 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   6150
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   600
         Width           =   705
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SALVAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3870
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   3660
         Width           =   1410
      End
      Begin VB.TextBox txtSetPoint 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1590
         Width           =   1320
      End
      Begin VB.TextBox txtUmidadeSet 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2610
         Width           =   1710
      End
      Begin VB.TextBox txtVentilacaoSet 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3690
         Width           =   1710
      End
      Begin VB.TextBox txtTempSup 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1500
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1590
         Width           =   1410
      End
      Begin VB.TextBox txtTempRet 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2970
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1590
         Width           =   1410
      End
      Begin VB.TextBox txtEscala 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4590
         MaxLength       =   1
         TabIndex        =   4
         Top             =   1590
         Width           =   570
      End
      Begin VB.TextBox txtUmidade 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1950
         MaxLength       =   8
         TabIndex        =   5
         Top             =   2610
         Width           =   1830
      End
      Begin VB.TextBox txtVentilacao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1950
         MaxLength       =   8
         TabIndex        =   6
         Top             =   3690
         Width           =   1830
      End
      Begin VB.TextBox txtGateIN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3870
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2610
         Width           =   3030
      End
      Begin VB.CommandButton Command2 
         Caption         =   "PLUG OFF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   5430
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1410
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
         Height          =   630
         Left            =   1875
         TabIndex        =   1
         Top             =   645
         Width           =   4170
         _ExtentX        =   7355
         _ExtentY        =   1111
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   26.25
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
         Height          =   615
         Left            =   105
         TabIndex        =   0
         Top             =   645
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   26.25
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
         Bindings        =   "Reefer.frx":0054
         Height          =   3405
         Left            =   90
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   4530
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   6006
         _Version        =   393216
         Cols            =   8
         RowHeightMin    =   420
         WordWrap        =   -1  'True
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "   |<Data                      |^Temp. SUP|^Temp. RET|^Umidade  |^Ventilacao|<Usuario                           |"
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
         _Band(0).Cols   =   8
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridSaidas 
         Bindings        =   "Reefer.frx":0069
         Height          =   3495
         Left            =   -74910
         TabIndex        =   16
         Top             =   4590
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   6165
         _Version        =   393216
         Cols            =   11
         RowHeightMin    =   420
         WordWrap        =   -1  'True
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "   |<Conteiner        |^Gate IN  |^Yard     |^Prev.Saida|^SetPoint|^1ª Medic.|^Data       |<Ult. Medic.|^Data         | "
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridEntradas 
         Bindings        =   "Reefer.frx":0083
         Height          =   3165
         Left            =   -74970
         TabIndex        =   17
         Top             =   840
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   5583
         _Version        =   393216
         Cols            =   11
         RowHeightMin    =   420
         WordWrap        =   -1  'True
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "   |<Conteiner        |^Tam |<Set Point|^Atracação        |<Local Atracacao|<Entrada Prevista  "
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
      Begin MSAdodcLib.Adodc AdodcEntradas 
         Height          =   345
         Left            =   -74640
         Top             =   3240
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
         Bindings        =   "Reefer.frx":009F
         Height          =   7365
         Left            =   -74880
         TabIndex        =   37
         Top             =   540
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   12991
         _Version        =   393216
         Cols            =   11
         RowHeightMin    =   420
         WordWrap        =   -1  'True
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "   |<Conteiner        |^Gate IN  |^Yard     |^Prev.Saida|^SetPoint|^1ª Medic.|^Data       |<Ult. Medic.|^Data         | "
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
         Left            =   -68730
         TabIndex        =   38
         Top             =   7980
         Width           =   705
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H008080FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   -74910
         Top             =   8130
         Width           =   255
      End
      Begin VB.Label lblReeferDesligado 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REEFER DESLIGADO"
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
         Left            =   2880
         TabIndex        =   34
         Top             =   300
         Visible         =   0   'False
         Width           =   2955
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
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   0
         Left            =   825
         TabIndex        =   32
         Top             =   330
         Width           =   1755
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SET POINT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   31
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UMIDADE (%)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   30
         Top             =   2340
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENTILACAO m3/h"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   29
         Top             =   3390
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TEMP SUP."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   28
         Top             =   1320
         Width           =   1305
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TEMP RET."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   27
         Top             =   1320
         Width           =   1290
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ESCALA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   26
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label lblIPA 
         BackColor       =   &H0080FFFF&
         Caption         =   "0"
         Height          =   315
         Left            =   6240
         TabIndex        =   25
         Top             =   4380
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
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4050
         TabIndex        =   22
         Top             =   2340
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UNIDADES PREVISTAS - ENTRADA PATIO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   1
         Left            =   -74910
         TabIndex        =   21
         Top             =   450
         Width           =   5760
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UNIDADES EM ESTOQUE "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   2
         Left            =   -73470
         TabIndex        =   20
         Top             =   4140
         Width           =   3570
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
         Left            =   -68880
         TabIndex        =   19
         Top             =   420
         Width           =   705
      End
      Begin VB.Label lblTotSai 
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
         Left            =   -68850
         TabIndex        =   18
         Top             =   4140
         Width           =   705
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
    Sql = "INSERT INTO OPERADOR..TB_MONITORING ("
    Sql = Sql & " [DT_MONITORING],"
    Sql = Sql & " [ID_CONTEINER],"
    Sql = Sql & " [AUTONUM_CNTR_IPA],"
    Sql = Sql & " [AUTONUM_CNTR_RDX],"
    Sql = Sql & " [AUTONUM_CNTR_OP],"
    Sql = Sql & " [TEMP_SETPOINT],"
    Sql = Sql & " [UMIDADE_SETPOINT],"
    Sql = Sql & " [VENTILACAO_SETPOINT],"
    Sql = Sql & " [TEMP_SUP],"
    Sql = Sql & " [TEMP_RET],"
    Sql = Sql & " [VENTILACAO],"
    Sql = Sql & " [UMIDADE],"
    Sql = Sql & " [ESCALA],"
    Sql = Sql & " [PATIO],"
    Sql = Sql & " [USUARIO])"
    Sql = Sql & " VALUES ("
    Sql = Sql & " GETDATE(),"
    Sql = Sql & "'" & Me.MskCntr & "',"
    Sql = Sql & Me.lblIPA.Caption & ","
    Sql = Sql & Me.lblRdx.Caption & ","
    Sql = Sql & Me.lblOp.Caption & ","
    Sql = Sql & "'" & Me.txtSetPoint.text & "',"
    Sql = Sql & "'" & Me.txtUmidadeSet.text & "',"
    Sql = Sql & "'" & Me.txtVentilacaoSet.text & "',"
    Sql = Sql & "'" & Me.txtTempSup.text & "',"
    Sql = Sql & "'" & Me.txtTempRet.text & "',"
    Sql = Sql & "'" & Me.txtVentilacao.text & "',"
    Sql = Sql & "'" & Me.txtUmidade.text & "',"
    Sql = Sql & "'" & Me.txtEscala.text & "',"
    Sql = Sql & Col_Patio & ","
    Sql = Sql & Cod_Usuario
    Sql = Sql & " )"
    db.Execute Sql
    
    If Val(Me.lblIPA.Caption) > 0 Then
        Sql = "UPDATE SGIPA..TB_CNTR_BL SET DT_PLUG_OFF=NULL WHERE AUTONUM=" & Val(Me.lblIPA.Caption)
        db.Execute Sql
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
Sql = "select autonum,  data_ent_temp, sistema, temperature, scale, umidade, ventilacao,flag_reefer, flag_reefer_desligado "
Sql = Sql & " from " & Banco_Operador & "VW_INVENT_SISTEMAS where patio=" & Col_Patio & " and id_conteiner='" & MskCntr & "'"
'Sql = Sql & " AND FLAG_REEFER='Y' AND FLAG_REEFER_DESLIGADO='N'"
Set TBvw = New ADODB.Recordset
Sql = CONVERTE_SQL(Sql)
TBvw.Open Sql, db, 1, 1

If Not TBvw.EOF Then

    
    If Nnull(TBvw!flag_reefer, 1) = "N" Then
        Frmmensagem.Label1.Caption = "Contêiner não é reefer "
        Frmmensagem.Show vbModal
        'MsgBox "Unidade não encontrada no estoque"
        MskCntr = "__________-_"
        MskCntr.SetFocus
        Exit Sub

    End If

    If Nnull(TBvw!flag_reefer_desligado, 1) = "Y" Then
        Frmmensagem.Label1.Caption = "Contêiner reefer captado com carga DRY!"
        Frmmensagem.Show vbModal
        'MsgBox "Unidade não encontrada no estoque"
        MskCntr = "__________-_"
        MskCntr.SetFocus
        Exit Sub

    End If
    
    idPatio = TBvw!AUTONUM
        
    If TBvw!Sistema = "I" Then
        Me.lblIPA = idPatio
        Dim PO$
        PO$ = Trim(Nnull(db.Execute("SELECT rtrim(convert(char,dt_plug_off,103)) + ' ' + rtrim(convert(char,dt_plug_off,108)) AS QUAl FROM SGIPA..TB_CNTR_BL WHERE AUTONUM=" & idPatio)(0), 1))
        If PO$ <> "" Then
            Me.lblReeferDesligado.Visible = True
            Me.lblPlugOFF.Caption = "PLUG OFF:" & PO$
            Me.lblPlugOFF.Visible = True
        Else
            Me.lblReeferDesligado.Visible = False
            Me.lblPlugOFF.Caption = ""
            Me.lblPlugOFF.Visible = False
        End If
    ElseIf TBvw!Sistema = "O" Then
        Me.lblOp = idPatio
    ElseIf TBvw!Sistema = "R" Then
        Me.lblRdx = idPatio
    End If
    Me.txtGateIN = Nnull(TBvw!data_ent_temp, 1)
    
    Me.txtEscala.text = Nnull(TBvw!Scale, 1)
    Me.txtSetPoint.text = Nnull(TBvw!TEMPERATURE, 1)
    If Nnull(TBvw!Umidade, 1) <> "" Then
        Me.txtUmidadeSet.text = Nnull(TBvw!Umidade, 1)
    Else
        Me.txtUmidadeSet.text = "OFF"
    End If
    If Nnull(TBvw!Ventilacao, 1) <> "" Then
        Me.txtVentilacaoSet.text = Nnull(TBvw!Ventilacao, 1)
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
TBvw.Close

    
Carrega_Grid1


Me.txtTempSup.SetFocus

End Sub

Private Sub Carrega_Grid1()
Grid2.Clear
Dim Sql As String

If Me.lblIPA.Caption <> "0" Or Me.lblRdx.Caption <> "0" Or Me.lblOp.Caption <> "0" Then

    Sql = "SELECT "
    Sql = Sql & " LTRIM(LEFT(CONVERT(CHAR,A.DT_MONITORING,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,A.DT_MONITORING,108),5))"
    Sql = Sql & " ,A.TEMP_SUP + ' ' + A.ESCALA"
    Sql = Sql & " ,A.TEMP_RET + ' ' + A.ESCALA"
    Sql = Sql & " ,A.UMIDADE "
    Sql = Sql & " ,A.VENTILACAO "
    Sql = Sql & " ,B.USUARIO "
    Sql = Sql & " ,A.AUTONUM"
    Sql = Sql & " FROM OPERADOR..TB_MONITORING A "
    Sql = Sql & " LEFT JOIN SGIPA..TB_CAD_USUARIOS B ON A.USUARIO=B.AUTONUM "
    Sql = Sql & " WHERE 0=0 "
    If Me.lblIPA.Caption <> "0" Then
        Sql = Sql & " AND A.AUTONUM_CNTR_IPA=" & Me.lblIPA.Caption
    ElseIf Me.lblRdx.Caption <> "0" Then
        Sql = Sql & " AND A.AUTONUM_CNTR_RDX=" & Me.lblRdx.Caption
    ElseIf Me.lblOp.Caption <> "0" Then
        Sql = Sql & " AND A.AUTONUM_CNTR_OP=" & Me.lblOp.Caption
    End If
    Sql = Sql & " ORDER BY A.AUTONUM DESC "
    
    Adodc6.ConnectionString = ConnectionStringAdo
    Adodc6.RecordSource = Sql
    Adodc6.Refresh
    Set Adodc6.Recordset.ActiveConnection = Nothing
End If

Grid2.FormatString = "   |<Data                      |^Temp. SUP|^Temp. RET|^Umidade  |^Ventilacao|<Usuario                           |"

End Sub

Private Sub Carrega_Grid_Entradas_Previstas()
Me.lblTotEnt.Caption = "0"

Dim Sql As String
Sql = " SELECT"
Sql = Sql & " C.ID_CONTEINER,"
Sql = Sql & " C.TAMANHO,"
Sql = Sql & " C.TEMP + ' ' + ISNULL(C.ESCALA,'') AS SETPOINT,"
Sql = Sql & " REPLACE(CASE V.DT_ATRACACAO  WHEN NULL THEN LEFT(CONVERT(CHAR,V.DT_ATRACACAOP,103),10) ELSE  LEFT(CONVERT(CHAR,V.DT_ATRACACAO,103),10) END,'01/01/1900','') AS DT_ATRACACAO,"
Sql = Sql & " CASE ISNULL(A.DESCR,'-') WHEN '-' THEN AP.DESCR ELSE A.DESCR END AS ARM_ATRACACAO,"
Sql = Sql & " LTRIM(LEFT(CONVERT(CHAR,T.HR_ENTRADA_PREV,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,T.HR_ENTRADA_PREV,108),5)) AS ENTRADA_PREVISTA"
Sql = Sql & " From"
Sql = Sql & " SGIPA..TB_CNTR_BL C INNER JOIN"
Sql = Sql & " SGIPA..DTE_TB_VIAGENS V ON C.VIAGEM=V.VIAGEM LEFT JOIN"
Sql = Sql & " SGIPA..DTE_TB_ARMAZENS AP ON V.ARM_ATRACACAOP=AP.CODE LEFT JOIN"
Sql = Sql & " SGIPA..DTE_TB_ARMAZENS A ON V.ARM_ATRACACAO=A.CODE LEFT JOIN"
Sql = Sql & " SGIPA..TB_TTIMEUNT T ON C.ID_CONTEINER=T.SIGLA AND C.GMCI=T.GMCI"
Sql = Sql & " Where c.FLAG_TERMINAL = 0 And c.GMCI Is Not Null"
Sql = Sql & " AND C.PATIO=" & Col_Patio
Sql = Sql & " AND ISNULL(C.TEMP,' ')<>' '"
Sql = Sql & " and isnull(c.flag_desligado,0)=0"
Sql = Sql & " order by c.id_conteiner "


AdodcEntradas.ConnectionString = ConnectionStringAdo
AdodcEntradas.RecordSource = Sql
AdodcEntradas.Refresh
Me.lblTotEnt.Caption = Me.AdodcEntradas.Recordset.RecordCount

Set AdodcEntradas.Recordset.ActiveConnection = Nothing


GridEntradas.FormatString = "  |<Conteiner           |^Tam |^Set Point|^Atracação        |<Local Atracacao|<Entrada Prevista  "



End Sub

Private Sub Carrega_Grid_Saidas_Previstas()

Me.lblTotSai.Caption = "0"

Dim Sql As String
Sql = " SELECT"
Sql = Sql & " C.ID_CONTEINER,"
Sql = Sql & " LTRIM(LEFT(CONVERT(CHAR,C.DT_ENTRADA,103),10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,C.DT_ENTRADA,108),5)) AS DT_ENTRADA,"
Sql = Sql & " C.YARD,"
Sql = Sql & " LTRIM(LEFT(CONVERT(CHAR,GD.PERIODO_INICIAL,103),10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,GD.PERIODO_INICIAL,108),5))  PERIODO_INICIAL,"
Sql = Sql & " C.TEMP + ' ' + ISNULL(C.ESCALA,'') AS SETPOINT,"
Sql = Sql & " M1.TEMP_RET,"
Sql = Sql & " LTRIM(LEFT(CONVERT(CHAR,M1.DT_MONITORING,103),10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,M1.DT_MONITORING,108),5)) AS DT1 ,"
Sql = Sql & " M2.TEMP_RET,"
Sql = Sql & " LTRIM(LEFT(CONVERT(CHAR,M2.DT_MONITORING,103),10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,M2.DT_MONITORING,108),5)) AS DT2, "
Sql = Sql & " c.dt_plug_off "


Sql = Sql & " From"
Sql = Sql & " SGIPA..TB_CNTR_BL C LEFT JOIN"
Sql = Sql & " (SELECT MIN(AUTONUM) AS A_PRIM ,AUTONUM_CNTR_IPA  FROM OPERADOR..TB_MONITORING GROUP BY AUTONUM_CNTR_IPA ) PRIM"
Sql = Sql & " ON C.AUTONUM=PRIM.AUTONUM_CNTR_IPA LEFT JOIN"
Sql = Sql & " OPERADOR..TB_MONITORING M1 ON PRIM.A_PRIM=M1.AUTONUM"
Sql = Sql & " Left Join"
Sql = Sql & " (SELECT MAX(AUTONUM) AS A_ULT ,AUTONUM_CNTR_IPA  FROM OPERADOR..TB_MONITORING GROUP BY AUTONUM_CNTR_IPA ) ULT"
Sql = Sql & " ON C.AUTONUM=ULT.AUTONUM_CNTR_IPA LEFT JOIN"
Sql = Sql & " OPERADOR..TB_MONITORING M2 ON ULT.A_ULT=M2.AUTONUM LEFT JOIN"
Sql = Sql & " OPERADOR..TB_GD_RESERVA GD ON C.AUTONUM_GD_RESERVA=GD.AUTONUM_GD_RESERVA"
Sql = Sql & " Where"
Sql = Sql & " c.Patio = " & Col_Patio
Sql = Sql & " AND C.FLAG_TERMINAL=1 AND C.FLAG_HISTORICO=0 AND C.DT_ENTRADA IS NOT NULL AND C.FLAG_DESLIGADO=0"
Sql = Sql & " AND ISNULL(C.TEMP,' ')<>' '"
If Me.Check1.Value = 1 Then
    Sql = Sql & " and c.dt_plug_off is not null"
End If
If Me.Check2.Value = 1 Then
    Sql = Sql & " and GD.PERIODO_INICIAL is not null"
End If

If Me.Check3.Value = 1 Then
    Sql = Sql & " and (c.yard LIKE 'CF%' OR C.YARD LIKE 'CONF%') "
End If

Sql = Sql & " ORDER BY C.ID_CONTEINER "

AdodcSaidas.ConnectionString = ConnectionStringAdo
AdodcSaidas.RecordSource = Sql
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

Dim Sql As String
Sql = " SELECT"
Sql = Sql & " C.ID_CONTEINER,"
Sql = Sql & " LTRIM(LEFT(CONVERT(CHAR,C.DT_ENTRADA,103),10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,C.DT_ENTRADA,108),5)) AS DT_ENTRADA,"
Sql = Sql & " C.YARD,"
Sql = Sql & " LTRIM(LEFT(CONVERT(CHAR,GD.PERIODO_INICIAL,103),10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,GD.PERIODO_INICIAL,108),5))  PERIODO_INICIAL,"
Sql = Sql & " C.TEMP + ' ' + ISNULL(C.ESCALA,'') AS SETPOINT,"
Sql = Sql & " M1.TEMP_RET,"
Sql = Sql & " LTRIM(LEFT(CONVERT(CHAR,M1.DT_MONITORING,103),10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,M1.DT_MONITORING,108),5)) AS DT1 ,"
Sql = Sql & " M2.TEMP_RET,"
Sql = Sql & " LTRIM(LEFT(CONVERT(CHAR,M2.DT_MONITORING,103),10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,M2.DT_MONITORING,108),5)) AS DT2, "
Sql = Sql & " LTRIM(LEFT(CONVERT(CHAR,c.DT_plug_off,103),10)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,c.dt_plug_off,108),5)) "


Sql = Sql & " From"
Sql = Sql & " SGIPA..TB_CNTR_BL C LEFT JOIN"
Sql = Sql & " (SELECT MIN(AUTONUM) AS A_PRIM ,AUTONUM_CNTR_IPA  FROM OPERADOR..TB_MONITORING GROUP BY AUTONUM_CNTR_IPA ) PRIM"
Sql = Sql & " ON C.AUTONUM=PRIM.AUTONUM_CNTR_IPA LEFT JOIN"
Sql = Sql & " OPERADOR..TB_MONITORING M1 ON PRIM.A_PRIM=M1.AUTONUM"
Sql = Sql & " Left Join"
Sql = Sql & " (SELECT MAX(AUTONUM) AS A_ULT ,AUTONUM_CNTR_IPA  FROM OPERADOR..TB_MONITORING GROUP BY AUTONUM_CNTR_IPA ) ULT"
Sql = Sql & " ON C.AUTONUM=ULT.AUTONUM_CNTR_IPA LEFT JOIN"
Sql = Sql & " OPERADOR..TB_MONITORING M2 ON ULT.A_ULT=M2.AUTONUM LEFT JOIN"
Sql = Sql & " OPERADOR..TB_GD_RESERVA GD ON C.AUTONUM_GD_RESERVA=GD.AUTONUM_GD_RESERVA"
Sql = Sql & " Where"
Sql = Sql & " c.Patio = " & Col_Patio
Sql = Sql & " AND C.FLAG_TERMINAL=1 AND C.FLAG_HISTORICO=0 AND C.DT_ENTRADA IS NOT NULL AND C.FLAG_DESLIGADO=0"
Sql = Sql & " AND ISNULL(C.TEMP,' ')<>' '"
Sql = Sql & " and c.dt_plug_off is not null"
Sql = Sql & " ORDER BY c.dt_plug_off desc "

AdodcDesligados.ConnectionString = ConnectionStringAdo
AdodcDesligados.RecordSource = Sql
AdodcDesligados.Refresh
Me.lblTotDesl.Caption = Me.AdodcDesligados.Recordset.RecordCount

Set AdodcDesligados.Recordset.ActiveConnection = Nothing

Me.GridDesligados.FormatString = "  |<Conteiner           |^Gate IN       |^Yard     |^Prev.Saida   |^SetPoint|^1ª Medic.|^Data            |^Ult. Medic.|^Data             |^Data Plug OFF  "




End Sub


Sub Busca_Cntr()

    If CNTR.ClipText = "" Then Exit Sub
    MskCntr = "__________-_"
    
    Dim Regs As Integer
    Sql = "SELECT COUNT(1) FROM " & Banco_Operador & "VW_INVENT_SISTEMAS WHERE PATIO=" & Col_Patio & " AND SUBSTR(ID_CONTEINER,8,5)='" & CNTR & "'"
    Sql = Sql & " and FLAG_REEFER='Y' AND FLAG_REEFER_DESLIGADO='N'"
    Sql = CONVERTE_SQL(Sql)
    Regs = Nnull(db.Execute(Sql).Fields(0), 0)
    
    If Regs = 0 Then
        'MsgBox "Nenhum contêiner com este final", vbExclamation
        Frmmensagem.Label1.Caption = "Nenhum contêiner com este final"
        Frmmensagem.Show vbModal

        CNTR = "___-_"
        CNTR.SetFocus
        Exit Sub
    End If
    
    If Regs = 1 Then
        
        Sql = "select id_conteiner from " & Banco_Operador & "VW_INVENT_SISTEMAS where patio=" & Col_Patio & " and SUBSTR(ID_CONTEINER,8,5)='" & CNTR & "'"
        Sql = Sql & " and FLAG_REEFER='Y' AND FLAG_REEFER_DESLIGADO='N'"
        Sql = CONVERTE_SQL(Sql)
        MskCntr = Nnull(db.Execute(Sql).Fields(0), 1)
    
    Else
        
        Frmmensagem.Label1.Caption = "Existe mais de 1 conteiner com este final. Favor informar a identificação completa"
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
        Sql = "UPDATE SGIPA..TB_CNTR_BL SET DT_PLUG_OFF=GETDATE() WHERE AUTONUM=" & Val(Me.lblIPA.Caption)
        db.Execute Sql
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
