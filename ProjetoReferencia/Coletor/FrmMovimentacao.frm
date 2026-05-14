VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FrmMovimentacao 
   BorderStyle     =   0  'None
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc31 
      Height          =   465
      Left            =   7290
      Top             =   4290
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   820
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
      Caption         =   "Adodc7"
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
   Begin VB.TextBox txtPOD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7725
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2610
      Visible         =   0   'False
      Width           =   435
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8430
      Left            =   30
      TabIndex        =   5
      Top             =   60
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   14870
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   14737632
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Moviment."
      TabPicture(0)   =   "FrmMovimentacao.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(7)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(8)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(9)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblReeferDesligado"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "LBLDTA"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblDestinadoDesova"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtPatio"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Grid2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "CNTR"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "MskCntr"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Command1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Command4"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "DataCombo2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Command2"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TXTYardAtual"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "mskDestino"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1(0)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Grid1"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Option1(0)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Option1(1)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Adodc5"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Frame2"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "Histórico"
      TabPicture(1)   =   "FrmMovimentacao.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(2)"
      Tab(1).Control(1)=   "DataGrid1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Dados CNTR"
      TabPicture(2)   =   "FrmMovimentacao.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtSIGVIG"
      Tab(2).Control(1)=   "Frame1"
      Tab(2).Control(2)=   "txtTam"
      Tab(2).Control(3)=   "txtTipo"
      Tab(2).Control(4)=   "txtEF"
      Tab(2).Control(5)=   "TxtNavio"
      Tab(2).Control(6)=   "TxtGateIN"
      Tab(2).Control(7)=   "txtBruto"
      Tab(2).Control(8)=   "txtIMO"
      Tab(2).Control(9)=   "txtLacres"
      Tab(2).Control(10)=   "DataGrid2"
      Tab(2).Control(11)=   "DataCombo1"
      Tab(2).Control(12)=   "DataCombo3"
      Tab(2).Control(13)=   "Label17"
      Tab(2).Control(14)=   "Label1(5)"
      Tab(2).Control(15)=   "Label1(6)"
      Tab(2).Control(16)=   "Label3"
      Tab(2).Control(17)=   "Label4"
      Tab(2).Control(18)=   "Label5"
      Tab(2).Control(19)=   "Label6"
      Tab(2).Control(20)=   "Label8"
      Tab(2).Control(21)=   "Label10"
      Tab(2).Control(22)=   "Label11"
      Tab(2).Control(23)=   "Label12"
      Tab(2).Control(24)=   "Label13"
      Tab(2).ControlCount=   25
      Begin VB.TextBox txtSIGVIG 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   5160
         Width           =   5400
      End
      Begin VB.Frame Frame2 
         Caption         =   "Câmera"
         Height          =   1005
         Left            =   60
         TabIndex        =   66
         Top             =   3300
         Visible         =   0   'False
         Width           =   6855
         Begin VB.CommandButton Command3 
            Caption         =   "Salvar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   4440
            TabIndex        =   68
            Top             =   210
            Width           =   2325
         End
         Begin MSDataListLib.DataCombo DataCombo4 
            Bindings        =   "FrmMovimentacao.frx":0054
            Height          =   405
            Left            =   120
            TabIndex        =   67
            Top             =   390
            Width           =   4230
            _ExtentX        =   7461
            _ExtentY        =   714
            _Version        =   393216
            MatchEntry      =   -1  'True
            Appearance      =   0
            Style           =   2
            ListField       =   "descr"
            BoundColumn     =   "autonum"
            Text            =   "DataCombo4"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "DADOS ENTRADA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74880
         TabIndex        =   49
         Top             =   3150
         Width           =   6765
         Begin VB.CheckBox Check1 
            Caption         =   "LACRES CONFERIDOS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2400
            TabIndex        =   65
            Top             =   510
            Width           =   4005
         End
         Begin VB.TextBox TXTgmci 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   6120
            MaxLength       =   20
            TabIndex        =   62
            Top             =   960
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.TextBox txtLACRE3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   5130
            MaxLength       =   20
            TabIndex        =   61
            Top             =   1350
            Width           =   1605
         End
         Begin VB.TextBox txtLACRE2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   3420
            MaxLength       =   20
            TabIndex        =   60
            Top             =   1350
            Width           =   1605
         End
         Begin VB.TextBox txtLACRE1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   1770
            MaxLength       =   20
            TabIndex        =   59
            Top             =   1350
            Width           =   1605
         End
         Begin VB.TextBox TxtLacreIPA 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   90
            MaxLength       =   20
            TabIndex        =   57
            Top             =   1350
            Width           =   1605
         End
         Begin VB.TextBox txtMaxGross 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   5400
            MaxLength       =   5
            TabIndex        =   55
            Top             =   -90
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.TextBox txtTara 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   930
            MaxLength       =   4
            TabIndex        =   53
            Top             =   390
            Width           =   1035
         End
         Begin VB.TextBox txtISO 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   5760
            MaxLength       =   4
            TabIndex        =   51
            Top             =   -150
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DEMAIS LACRES"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2820
            TabIndex        =   58
            Top             =   960
            Width           =   2355
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LACRE IPA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   60
            TabIndex        =   56
            Top             =   960
            Width           =   1545
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MAXGROSS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   5310
            TabIndex        =   54
            Top             =   -30
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TARA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   60
            TabIndex        =   52
            Top             =   450
            Width           =   795
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ISO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5670
            TabIndex        =   50
            Top             =   120
            Visible         =   0   'False
            Width           =   555
         End
      End
      Begin MSAdodcLib.Adodc Adodc5 
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
         Caption         =   "Adodc5"
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
      Begin VB.OptionButton Option1 
         Caption         =   "DESTINO"
         Height          =   345
         Index           =   1
         Left            =   5970
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   4980
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ATUAL"
         Height          =   345
         Index           =   0
         Left            =   5970
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   4620
         Value           =   -1  'True
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
         Bindings        =   "FrmMovimentacao.frx":006A
         Height          =   2115
         Left            =   60
         TabIndex        =   43
         Top             =   5340
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   3731
         _Version        =   393216
         RowHeightMin    =   480
         WordWrap        =   -1  'True
         Appearance      =   0
         FormatString    =   "|<Motivo                   "
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
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox txtTam 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -74820
         TabIndex        =   27
         Top             =   855
         Width           =   675
      End
      Begin VB.TextBox txtTipo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -74130
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   855
         Width           =   705
      End
      Begin VB.TextBox txtEF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -73410
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   855
         Width           =   480
      End
      Begin VB.TextBox TxtNavio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -72915
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   855
         Width           =   4875
      End
      Begin VB.TextBox TxtGateIN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74850
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1740
         Width           =   2520
      End
      Begin VB.TextBox txtBruto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -74850
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2535
         Width           =   1110
      End
      Begin VB.TextBox txtIMO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -73695
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2535
         Width           =   645
      End
      Begin VB.TextBox txtLacres 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -73005
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2535
         Width           =   4770
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1515
         Index           =   0
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   3090
         Width           =   6780
      End
      Begin VB.TextBox mskDestino 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2130
         MaxLength       =   7
         TabIndex        =   2
         Top             =   1590
         Width           =   2235
      End
      Begin VB.TextBox TXTYardAtual 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1590
         Width           =   1950
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Limpar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   180
         TabIndex        =   12
         Top             =   7530
         Width           =   2280
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "FrmMovimentacao.frx":007F
         Height          =   405
         Left            =   1170
         TabIndex        =   3
         Top             =   2310
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   714
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         ListField       =   "descricao"
         BoundColumn     =   "autonum"
         Text            =   "DataCombo2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         TabIndex        =   9
         Top             =   600
         Width           =   705
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmMovimentacao.frx":0094
         Height          =   6780
         Left            =   -74955
         TabIndex        =   7
         Top             =   1110
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   11959
         _Version        =   393216
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   19
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "data"
            Caption         =   "Data"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/mm/yyyy hh:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "origem"
            Caption         =   "De"
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
            DataField       =   "destino"
            Caption         =   "Para"
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
            DataField       =   "nmusuario"
            Caption         =   "Usuario"
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
            DataField       =   "identificacao"
            Caption         =   "Equipamento"
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
            DataField       =   "desc_motivo"
            Caption         =   "Motivo"
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
            DataField       =   "veiculo"
            Caption         =   "Veiculo"
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
               ColumnWidth     =   1755,213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   900,284
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   945,071
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2264,882
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2924,788
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   2564,788
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Salvar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   5490
         TabIndex        =   4
         Top             =   1620
         Width           =   1410
      End
      Begin MSMask.MaskEdBox MskCntr 
         Height          =   630
         Left            =   1935
         TabIndex        =   1
         Top             =   645
         Width           =   4020
         _ExtentX        =   7091
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   2520
         Left            =   -74880
         TabIndex        =   16
         Top             =   5730
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   4445
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   34
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "FrmMovimentacao.frx":00A9
         Height          =   465
         Left            =   -72300
         TabIndex        =   36
         Top             =   1740
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   820
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         ListField       =   "identificacao"
         BoundColumn     =   "autonum"
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "FrmMovimentacao.frx":00BE
         Height          =   465
         Left            =   -70650
         TabIndex        =   37
         Top             =   1740
         Visible         =   0   'False
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   820
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         ListField       =   "identificacao"
         BoundColumn     =   "autonum"
         Text            =   "DataCombo3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
         Bindings        =   "FrmMovimentacao.frx":00D3
         Height          =   2985
         Left            =   2790
         TabIndex        =   46
         Top             =   5340
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   5265
         _Version        =   393216
         Cols            =   4
         RowHeightMin    =   480
         WordWrap        =   -1  'True
         Appearance      =   0
         FormatString    =   "|<Conteiner   |<Posição|<Motivo                     "
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
         _Band(0).Cols   =   4
      End
      Begin VB.TextBox txtPatio 
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
         Height          =   675
         Left            =   4410
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   1590
         Width           =   1050
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SIGVIG"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -74730
         TabIndex        =   70
         Top             =   5220
         Width           =   1065
      End
      Begin VB.Label lblDestinadoDesova 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESTINADO DESOVA"
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
         Left            =   3930
         TabIndex        =   64
         Top             =   330
         Visible         =   0   'False
         Width           =   2985
      End
      Begin VB.Label LBLDTA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DTA"
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
         TabIndex        =   63
         Top             =   330
         Visible         =   0   'False
         Width           =   600
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
         Left            =   4020
         TabIndex        =   48
         Top             =   1260
         Visible         =   0   'False
         Width           =   2955
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CARREGAMENTO REGISTRADO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Index           =   9
         Left            =   60
         TabIndex        =   47
         Top             =   2670
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FIADA ATUAL"
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
         Index           =   8
         Left            =   2790
         TabIndex        =   42
         Top             =   4980
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONTEINER"
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
         Index           =   7
         Left            =   210
         TabIndex        =   41
         Top             =   5010
         Width           =   1680
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MOVIMENTOS PENDENTES"
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
         Index           =   3
         Left            =   1350
         TabIndex        =   40
         Top             =   4650
         Width           =   3855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stacker"
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
         Index           =   5
         Left            =   -72300
         TabIndex        =   39
         Top             =   1410
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Truck"
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
         Index           =   6
         Left            =   -70710
         TabIndex        =   38
         Top             =   1410
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tam"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74775
         TabIndex        =   35
         Top             =   465
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74010
         TabIndex        =   34
         Top             =   465
         Width           =   600
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73305
         TabIndex        =   33
         Top             =   480
         Width           =   345
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Navio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72660
         TabIndex        =   32
         Top             =   510
         Width           =   4560
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GATE IN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74160
         TabIndex        =   31
         Top             =   1380
         Width           =   1200
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BRUTO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -74730
         TabIndex        =   30
         Top             =   2220
         Width           =   885
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -73605
         TabIndex        =   29
         Top             =   2220
         Width           =   525
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LACRES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -72525
         TabIndex        =   28
         Top             =   2205
         Width           =   975
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
         TabIndex        =   17
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
         TabIndex        =   15
         Top             =   330
         Width           =   1755
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "POS. ATUAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   13
         Top             =   1260
         Width           =   1725
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2190
         TabIndex        =   11
         Top             =   2220
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
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
         Index           =   4
         Left            =   150
         TabIndex        =   10
         Top             =   2340
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Origem"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   -74775
         TabIndex        =   8
         Top             =   615
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESTINO"
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
         Left            =   2550
         TabIndex        =   6
         Top             =   1260
         Width           =   1290
      End
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   3675
      Top             =   2700
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   3690
      Top             =   2175
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
      Caption         =   "Adodc4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Left            =   3750
      Top             =   1785
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8895
      Top             =   1035
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   345
      Left            =   2250
      Top             =   8100
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
      Caption         =   "Adodc5"
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
End
Attribute VB_Name = "FrmMovimentacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim idPatio As Long
'Const Banco_Em_Uso = "ORACLE"
'Const Banco_Operador = "OPERADOR."

Sub Atualiza_Posicao()


If idPatio = 0 Then
    Frmmensagem.Label1.Caption = "Um conteiner deve ser informado para pesquisa!"
    Frmmensagem.Show vbModal
    'MsgBox "Um conteiner deve ser informado para pesquisa", vbExclamation
    MskCntr.SetFocus
    Exit Sub
End If


If idPatio > 0 And MskCntr.Tag = "I" Then
    'SE FOR O PRIMEIRO MOVIMENTO - PEDIR QUE CLIQUEM NO LACRE CONFERIDO
    Dim Sql As String
    Sql = "SELECT AUTONUM FROM OPERADOR..TB_HIST_SHIFTING WHERE CNTR=" & idPatio & " AND TIPO='I' AND ORIGEM<>'GATEIN'"
    Set tbhist = New ADODB.Recordset
    tbhist.Open Sql, db, 1, 1
    If tbhist.EOF Then
        If Me.Check1.Value = 0 Then
            MsgBox "Informe que os lacres foram conferidos", vbInformation
            Exit Sub
        End If
    End If
End If
    
    
If Trim(Me.txtLACRE1.text) = "" And Trim(Me.txtLACRE2.text) = "" And Trim(Me.txtLACRE3.text) = "" Then
    MsgBox "Pelo menos 1 lacre de origem deve ser informado", vbInformation
    Exit Sub
End If

    
'If DataCombo1.BoundText = "" Then
'    MsgBox "Informe a stacker utilizada", vbExclamation
'    DataCombo1.SetFocus
'    Exit Sub
'End If
If DataCombo2.BoundText = "" Then
    Frmmensagem.Label1.Caption = "Informe o motivo da movimentação!"
    Frmmensagem.Show vbModal
    'MsgBox "Informe o motivo da movimentação", vbExclamation
    DataCombo2.SetFocus
    Exit Sub
End If
        
If mskDestino = "" Then
    Frmmensagem.Label1.Caption = "Destino não Informado!"
    Frmmensagem.Show vbModal
    'MsgBox "Destino não informado", vbExclamation
    mskDestino.SetFocus
    Exit Sub
End If

If mskDestino = Me.TXTYardAtual Then
    Frmmensagem.Label1.Caption = "Destino e Origem iguais!"
    Frmmensagem.Show vbModal
    'MsgBox "Destino e Origem iguais", vbExclamation
    mskDestino.SetFocus
    Exit Sub
End If




        
Dim Saida As Boolean
Saida = False
If mskDestino = "SAIDA" Then
    If MsgBox("Confirma Saida da Unidade ?", vbYesNo + vbQuestion) = vbNo Then
        MsgBox "Operação Cancelada"
        mskDestino = ""
        mskDestino.SetFocus
        Exit Sub
    End If
    Saida = True
End If



Dim Valida As Boolean
Dim tb1 As New ADODB.Recordset


'Existe a posicao
If mskDestino <> "SAIDA" Then

    If Me.txtTara.text = "" Or Me.txtTara.text = "0" Then
        Frmmensagem.Label1.Caption = "Informe a tara do conteiner"
        Frmmensagem.Show vbModal
        Exit Sub
    End If
    
'    If Me.txtMaxGross.Text = "" Or Me.txtMaxGross.Text = "0" Then
'        Frmmensagem.Label1.Caption = "Informe o MaxGross do conteiner"
'        Frmmensagem.Show vbModal
'        Exit Sub
'    End If
    
'    If Me.txtISO.Text = "" Then
'        Frmmensagem.Label1.Caption = "Informe o código ISO do conteiner"
'        Frmmensagem.Show vbModal
'        Exit Sub
'    End If
    
    If Me.txtISO <> "" Then
        Dim ValT$
        ValT$ = Valida_Tara(Me.txtISO, Me.txtTara, Me.txtTam)
        
        If ValT$ <> "" Then
            Frmmensagem.Label1.Caption = ValT
            Frmmensagem.Show vbModal
            Exit Sub
        End If
    End If



    PRSet tb1, "SELECT VALIDA,FLAG_BLOQUEIO,MOTIVO_BLOQUEIO FROM " & Banco_Operador & "TB_YARD WHERE PATIO=" & Col_Patio & " AND YARD='" & mskDestino & "'"
    If tb1.EOF Then
        Frmmensagem.Label1.Caption = "Posição não encontrada!"
        Frmmensagem.Show vbModal
    
        'MsgBox "Posição não encontrada", vbExclamation:
        Exit Sub
    Else
        If tb1!Valida = 1 Then Valida = True
        If tb1!Valida = 0 Then Valida = False
        
         If Nnull(tb1!FLAG_BLOQUEIO, 0) = 1 Then
            Frmmensagem.Label1.Caption = "Posição Bloqueada (" & Nnull(tb1!motivo_bloqueio, 1) & ")"
            Frmmensagem.Show vbModal
            'MsgBox "Posição Bloqueada (" & Nnull(tb1!motivo_bloqueio, 1) & ")", vbCritical:
            Exit Sub
        End If
        
    End If
    
    
    
    
    
    If Valida Then
    
    
        If Nnull(tb1!FLAG_BLOQUEIO, 0) = 1 Then
            Frmmensagem.Label1.Caption = "Posição Bloqueada (" & Nnull(tb1!motivo_bloqueio, 1) & ")"
            Frmmensagem.Show vbModal
            'MsgBox "Posição Bloqueada (" & Nnull(tb1!motivo_bloqueio, 1) & ")", vbCritical:
            Exit Sub
        End If
    
        Dim SqlYArd$
    
    
        Dim ColQuadraD$
        Dim ColFiadaD$
        Dim ColLastroD$
        Dim ColAlturaD$
        Dim ColPilhaD$
        
        Dim ColQuadraO$
        Dim ColFiadaO$
        Dim ColLastroO$
        Dim ColAlturaO$
        Dim ColPilhaO$
        
        If Col_Tam_Pos_Yard = 6 Then
            ColQuadraD$ = Left(mskDestino, 1)
            ColFiadaD$ = Mid$(mskDestino, 2, 2)
            ColLastroD$ = Mid$(mskDestino, 4, 2)
            ColAlturaD$ = Right(mskDestino, 1)
        Else
            ColQuadraD$ = Left(mskDestino, 2)
            ColFiadaD$ = Mid$(mskDestino, 3, 2)
            ColLastroD$ = Mid$(mskDestino, 5, 2)
            ColAlturaD$ = Right(mskDestino, 1)
        End If
        ColPilhaD = ColQuadraD + ColFiadaD + ColLastroD
            
        If Val(ColFiadaD) Mod 2 <> 0 And txtTam = "40" Then
            Frmmensagem.Label1.Caption = "Posição inválida Cntr de 40 deve estar em fiada par!"
            Frmmensagem.Show vbModal
            'MsgBox "Posição inválida Cntr de 40 deve estar em fiada par!", vbCritical, "PÁTIO"
            mskDestino.SetFocus
            Exit Sub
        End If
    
        If Val(ColFiadaD) Mod 2 = 0 And txtTam = "20" Then
            Frmmensagem.Label1.Caption = "Posição inválida Cntr de 20' deve estar em fiada ímpar!"
            Frmmensagem.Show vbModal
            'MsgBox "Posição inválida Cntr de 20' deve estar em fiada ímpar!", vbCritical, "PÁTIO"
            mskDestino.SetFocus
            Exit Sub
        End If
        
        
        If Nnull(txtTam, 0) = 40 Then
            PosAnt = ColQuadraD & Right$(Format(Val(ColFiadaD) - 2, "00"), 2) & ColLastroD & ColAlturaD
            PosPos$ = ColQuadraD & Right$(Format(Val(ColFiadaD) + 2, "00"), 2) & ColLastroD & ColAlturaD
                'Nao pode ter um Cntr 40 na anterior
            Set Tbh = New ADODB.Recordset
            Tbh.Open "Select id_conteiner From " & Banco_Operador & "VW_INVENT_SISTEMAS WHERE Yard='" & PosAnt & "' and patio=" & Col_Patio & " and tamanho=40 ", db, 3, 3
            If Not Tbh.EOF Then
                Frmmensagem.Label1.Caption = "Container em concomitancia com a unidade  " & Tbh!Id_Conteiner
                Frmmensagem.Show vbModal
                    
                    'MsgBox "Container em concomitancia com a unidade  " & Tbh!Id_Conteiner, vbCritical:
                Exit Sub
            End If
                'Nao pode ter um Cntr 40 na posterior
            Set Tbh = New ADODB.Recordset
            Tbh.Open "Select id_conteiner From " & Banco_Operador & "VW_INVENT_SISTEMAS WHERE Yard='" & PosPos$ & "' and tamanho=40 and patio=" & Col_Patio, db, 3, 3
            If Not Tbh.EOF Then
                Frmmensagem.Label1.Caption = "Container em concomitancia com a unidade  " & Tbh!Id_Conteiner
                Frmmensagem.Show vbModal
                
                'MsgBox "Container em concomitancia com a unidade  " & Tbh!Id_Conteiner, vbCritical:
                Exit Sub
            End If
            
            PosAnt = ColQuadraD & Right$(Format(Val(ColFiadaD) - 1, "00"), 2) & ColLastroD & ColAlturaD
            PosPos$ = ColQuadraD & Right$(Format(Val(ColFiadaD) + 1, "00"), 2) & ColLastroD & ColAlturaD
            
            'Nao pode ter um Cntr 20 na anterior
            Set Tbh = New ADODB.Recordset
            Tbh.Open "Select id_conteiner From " & Banco_Operador & "VW_INVENT_SISTEMAS WHERE Yard='" & PosAnt & "'  and tamanho=20 and patio=" & Col_Patio, db, 3, 3
            If Not Tbh.EOF Then
                Frmmensagem.Label1.Caption = "Container em concomitancia com a unidade  " & Tbh!Id_Conteiner
                Frmmensagem.Show vbModal
                'MsgBox "Container em concomitancia com a unidade  " & Tbh!Id_Conteiner, vbCritical:
                Exit Sub
            End If
            'Nao pode ter um Cntr 20 na posterior
            Set Tbh = New ADODB.Recordset
            Tbh.Open "Select id_conteiner From " & Banco_Operador & "VW_INVENT_SISTEMAS WHERE Yard='" & PosPos$ & "'  and tamanho=20 and patio=" & Col_Patio, db, 3, 3
            If Not Tbh.EOF Then
                Frmmensagem.Label1.Caption = "Container em concomitancia com a unidade  " & Tbh!Id_Conteiner
                Frmmensagem.Show vbModal
                
                'MsgBox "Container em concomitancia com a unidade  " & Tbh!Id_Conteiner, vbCritical:
                Exit Sub
            End If
        End If
        
        If Nnull(txtTam, 0) = 20 Then
            PosAnt = ColQuadraD & Right$(Format(Val(ColFiadaD) - 1, "00"), 2) & ColLastroD & ColAlturaD
            PosPos$ = ColQuadraD & Right$(Format(Val(ColFiadaD) + 1, "00"), 2) & ColLastroD & ColAlturaD
            'Nao pode ter um Cntr 40 na anterior
            Set Tbh = New ADODB.Recordset
            Tbh.Open "Select id_conteiner From " & Banco_Operador & "VW_INVENT_SISTEMAS WHERE Yard='" & PosAnt & "' and tamanho=20 and patio=" & Col_Patio & "  and autonum<>" & idPatio, db, 3, 3
            If Not Tbh.EOF Then
                Frmmensagem.Label1.Caption = "Container em concomitancia com a unidade  " & Tbh!Id_Conteiner
                Frmmensagem.Show vbModal
                'MsgBox "Container em concomitancia com a unidade  " & Tbh!Id_Conteiner, vbCritical:
                Exit Sub
            End If
            'Nao pode ter um Cntr 40 na posterior
            Set Tbh = New ADODB.Recordset
            Tbh.Open "Select id_conteiner From " & Banco_Operador & "VW_INVENT_SISTEMAS WHERE Yard='" & PosPos$ & "' and tamanho=40 and patio=" & Col_Patio & " and autonum<>" & idPatio, db, 3, 3
            If Not Tbh.EOF Then
                Frmmensagem.Label1.Caption = "Container em concomitancia com a unidade  " & Tbh!Id_Conteiner
                Frmmensagem.Show vbModal
               ' MsgBox "Container em concomitancia com a unidade  " & Tbh!Id_Conteiner, vbCritical:
               Exit Sub
            End If
            'Nao pode ter um Cntr 40 na anterior
            Set Tbh = New ADODB.Recordset
            Tbh.Open "Select id_conteiner From " & Banco_Operador & "VW_INVENT_SISTEMAS WHERE Yard='" & PosAnt & "' and tamanho=40 and patio=" & Col_Patio & " and autonum<>" & idPatio, db, 3, 3
            If Not Tbh.EOF Then
                Frmmensagem.Label1.Caption = "Container em concomitancia com a unidade  " & Tbh!Id_Conteiner
                Frmmensagem.Show vbModal
                'MsgBox "Container em concomitancia com a unidade  " & Tbh!Id_Conteiner, vbCritical:
                Exit Sub
            End If
        End If
        
            'Verificar se CNTR de 20 em pilha de 40
            
        If Col_Flag_Pilha_2040 = 0 Then
            If mskDestino <> "" Then
                If Nnull(txtTam, 0) = 20 Then
                    Set Tbh = New ADODB.Recordset
                    Tbh.Open "SELECT Yard, Autonum From " & Banco_Operador & "VW_INVENT_SISTEMAS WHERE Yard LIKE '" & ColPilhaD & "%' and tamanho=40  and patio=" & Col_Patio & " and autonum <> " & idPatio, db, 3, 3
                    If Not Tbh.EOF Then
                        If Nnull(Tbh!AUTONUM, 0) <> Val(idPatio) Then
                            Frmmensagem.Label1.Caption = "Esta pilha já contém cntr de 40 pés"
                            Frmmensagem.Show vbModal
                            'MsgBox "Esta pilha já contém cntr de 40 pés", vbCritical
                            Exit Sub
                        End If
                    End If
                    Tbh.Close
                    PosAnt = ColQuadraD & Format(Val(ColFiadaD) - 1, "00") & ColLastroD
                    
                    Set Tbh = New ADODB.Recordset
                    Tbh.Open "SELECT Yard, Autonum From " & Banco_Operador & "VW_INVENT_SISTEMAS WHERE Yard LIKE '" & PosAnt & "%' and tamanho=40 and patio=" & Col_Patio & " and autonum<>" & idPatio, db, 3, 3
                    If Not Tbh.EOF Then
                        If Nnull(Tbh!AUTONUM, 0) <> Val(idPatio) Then
                            Frmmensagem.Label1.Caption = "Esta pilha já contém cntr de 40 pés"
                            Frmmensagem.Show vbModal
                            'MsgBox "Esta pilha já contém cntr de 40 pés", vbCritical
                            Exit Sub
                        End If
                    End If
                    Tbh.Close
                End If
            End If
            
            'Verificar se CNTR de 40 em pilha de 20
            If mskDestino <> "" Then
                If Nnull(txtTam, 0) = 40 Then
                    Set Tbh = New ADODB.Recordset
                    Tbh.Open "SELECT Yard, Autonum From " & Banco_Operador & "VW_INVENT_SISTEMAS WHERE Yard LIKE '" & ColPilhaD & "%' and tamanho=20  and patio=" & Col_Patio & " and autonum<>" & idPatio, db, 3, 3
                    If Not Tbh.EOF Then
                        If Nnull(Tbh!AUTONUM, 0) <> Val(idPatio) Then
                            Frmmensagem.Label1.Caption = "Esta pilha já contém cntr de 20 pés"
                            Frmmensagem.Show vbModal
                           ' MsgBox "Esta pilha já contém cntr de 20 pés", vbCritical
                            Exit Sub
                        End If
                    End If
                    Tbh.Close
                    PosAnt = ColQuadraD & Format(Val(ColFiadaD) + 1, "00") & ColLastroD
                    Set Tbh = New ADODB.Recordset
                    Tbh.Open "SELECT Yard, Autonum From " & Banco_Operador & "VW_INVENT_SISTEMAS WHERE Yard LIKE '" & PosAnt & "%' and tamanho=20  and patio=" & Col_Patio & " and autonum<>" & idPatio, db, 3, 3
                    If Not Tbh.EOF Then
                        If Nnull(Tbh!AUTONUM, 0) <> Val(idPatio) Then
                            Frmmensagem.Label1.Caption = "Esta pilha já contém cntr de 20 pés"
                            Frmmensagem.Show vbModal
                            'MsgBox "Esta pilha já contém cntr de 20 pés", vbCritical
                            Exit Sub
                        End If
                    End If
                    Tbh.Close
                End If
            End If
        End If
                        
        
    
        If mskDestino <> "" Then
            Set Tbh = New ADODB.Recordset
            Tbh.Open "SELECT Yard, Autonum, ID_CONTEINER From " & Banco_Operador & "VW_INVENT_SISTEMAS WHERE Yard='" & Nnull(mskDestino, 1) & "' and patio=" & Col_Patio & " and autonum<>" & Nnull(idPatio, 0), db, 3, 3
            If Not Tbh.EOF Then
                If Nnull(Tbh!AUTONUM, 0) <> Val(idPatio) Then
                    Frmmensagem.Label1.Caption = "Existe um container alocado nesta posição (" & Nnull(Tbh!Id_Conteiner, 1) & ")"
                    Frmmensagem.Show vbModal
                    Exit Sub
                End If
            End If
            Tbh.Close
        
            ' Verifica Altura
            PosAnt = ColPilhaD & Right$(Format(Val(ColAlturaD) - 1, "0"), 1)
    
            Dim AlturaAntBloq As Integer
            AlturaAntBloq = 0
        
            Set TbA = New ADODB.Recordset
            TbA.Open "Select flag_bloqueio from " & Banco_Operador & "tb_yard where patio=" & Col_Patio & " and yard='" & PosAnt & "' and flag_bloqueio=1", db, 1, 1
            If Not TbA.EOF Then AlturaAntBloq = 1
            TbA.Close
    
            If Right(PosAnt, 1) <> 0 Then
                'Tem que ter um Cntr na anterior
                'Operador
                Set Tbh = New ADODB.Recordset
                Tbh.Open "Select id_conteiner From " & Banco_Operador & "VW_INVENT_SISTEMAS WHERE Yard='" & PosAnt & "' and patio=" & Col_Patio & " AND id_conteiner<>'" & Me.MskCntr & "'", db, 3, 3
                If Tbh.EOF Then
                    If AlturaAntBloq = 0 Then
                         Frmmensagem.Label1.Caption = "Falta Container na Altura Anterior "
                         Frmmensagem.Show vbModal
                        'MsgBox "Falta Container na Altura Anterior ", vbCritical:
                        Exit Sub
                    End If
                End If
            End If
        
            
        End If
        
        
        
    End If
    
    
    If mskDestino <> "" Then
    
    
        Set tbregras = New ADODB.Recordset
        
        Sql = " SELECT"
        Sql = Sql & " ISNULL(max(B1.OBS),'OK') as  MENSAGEM "
        Sql = Sql & " FROM    INTELOPER..TB_IP_ATRIBUTO_YARD A1 "
        Sql = Sql & " Inner Join"
        Sql = Sql & " INTELOPER..TB_IP_ATRIBUTOS B1"
        Sql = Sql & " ON A1.AUTONUM_ATR = B1.AUTONUM_ATR,"
        Sql = Sql & " (  SELECT MAX (A.PATIO) AS PATIO,"
        Sql = Sql & " MAX (A.YARD) AS YARD,"
        Sql = Sql & " MAX (A.AUTONUM_ATR) AS AUTONUM_ATR,"
        Sql = Sql & " MAX (B.DESCR_ATR) AS DESCR_ATR,"
        Sql = Sql & " MAX (A.STATUS_ATRIB) AS STATUS_ATRIB,"
        Sql = Sql & " MAX (a.DT_INI_ATR) As DT_INI_ATR"
        Sql = Sql & " FROM INTELOPER..TB_IP_ATRIBUTO_YARD A"
        Sql = Sql & " Inner Join"
        Sql = Sql & " INTELOPER..TB_IP_ATRIBUTOS b"
        Sql = Sql & " ON A.AUTONUM_ATR = B.AUTONUM_ATR"
        Sql = Sql & " Where a.DT_INI_ATR <= GETDATE() "
        Sql = Sql & " AND (A.DT_FIM_ATR >= GETDATE() OR A.DT_FIM_ATR IS NULL)"
        Sql = Sql & " GROUP BY A.PATIO,"
        Sql = Sql & " A.YARD,"
        Sql = Sql & " A.AUTONUM_ATR,"
        Sql = Sql & " B.DESCR_ATR,"
        Sql = Sql & " A.STATUS_ATRIB) C1,"
        Sql = Sql & " (SELECT * FROM OPERADOR..VW_INVENT_SISTEMAS WHERE ID_CONTEINER='" & Me.MskCntr & "') I"
        Sql = Sql & " Where A1.Patio = C1.Patio"
        Sql = Sql & " AND A1.YARD = C1.YARD"
        Sql = Sql & " AND A1.AUTONUM_ATR = C1.AUTONUM_ATR"
        Sql = Sql & " AND A1.STATUS_ATRIB = C1.STATUS_ATRIB"
        Sql = Sql & " AND A1.DT_INI_ATR = C1.DT_INI_ATR"
        Sql = Sql & " AND A1.PATIO=I.PATIO"
        Sql = Sql & " AND (A1.YARD = '" & Me.mskDestino & "' OR A1.YARD='" & ColQuadraD & "')"
        Sql = Sql & " AND ( (    A1.STATUS_ATRIB = 'R'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 1"
        Sql = Sql & " AND A1.VLR_ATRIB = I.VIAGEM)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'E'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 1"
        Sql = Sql & " AND A1.VLR_ATRIB <> I.VIAGEM)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'R'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 3"
        Sql = Sql & " AND A1.VLR_ATRIB = I.LINE)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'E'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 3"
        Sql = Sql & " AND A1.VLR_ATRIB <> I.LINE)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'R'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 4"
        Sql = Sql & " AND A1.VLR_ATRIB = I.TAMANHO)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'E'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 4"
        Sql = Sql & " AND A1.VLR_ATRIB <> I.TAMANHO)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'R'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 5"
        Sql = Sql & " AND A1.VLR_ATRIB = I.TIPOBASICO)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'E'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 5"
        Sql = Sql & " AND A1.VLR_ATRIB <> I.TIPOBASICO)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'R'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 6"
        Sql = Sql & " AND A1.VLR_ATRIB = I.POD)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'E'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 6"
        Sql = Sql & " AND A1.VLR_ATRIB <> I.POD)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'R'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 7"
        Sql = Sql & " AND A1.VLR_ATRIB = I.EF)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'E'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 7"
        Sql = Sql & " AND A1.VLR_ATRIB <> I.EF)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'R'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 8"
        Sql = Sql & " AND A1.VLR_ATRIB = I.FLAG_TRS)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'E'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 8"
        Sql = Sql & " AND A1.VLR_ATRIB <> I.FLAG_TRS)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'R'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 9"
        Sql = Sql & " AND A1.VLR_ATRIB = I.DLV_TERM)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'E'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 9"
        Sql = Sql & " AND A1.VLR_ATRIB <> I.DLV_TERM)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'R'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 10"
        Sql = Sql & " AND A1.VLR_ATRIB = I.DOC_TYPE)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'E'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 10"
        Sql = Sql & " AND A1.VLR_ATRIB <> I.DOC_TYPE)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'R'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 11"
        Sql = Sql & " AND A1.VLR_ATRIB = I.CNTR_REG)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'E'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 11"
        Sql = Sql & " AND A1.VLR_ATRIB <> I.CNTR_REG)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'R'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 12"
        Sql = Sql & " AND A1.VLR_ATRIB = I.IMO1)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'E'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 12"
        Sql = Sql & " AND A1.VLR_ATRIB <> I.IMO1)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'R'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 13"
        Sql = Sql & " AND A1.VLR_ATRIB = CASE I.FLAG_REEFER WHEN 'Y' THEN 1 ELSE 0 END )"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'E'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 13"
        Sql = Sql & " AND A1.VLR_ATRIB <> CASE I.FLAG_REEFER WHEN 'Y' THEN 1 ELSE 0 END )"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'R'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 14"
        Sql = Sql & " AND A1.VLR_ATRIB = CASE I.FLAG_OOG WHEN 'Y' THEN 1 ELSE 0 END )"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'E'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 14"
        Sql = Sql & " AND A1.VLR_ATRIB <> CASE I.FLAG_OOG WHEN 'Y' THEN 1 ELSE 0 END )"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'R'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 15"
        Sql = Sql & " AND A1.VLR_ATRIB = CASE I.FLAG_SPC WHEN 'Y' THEN 1 ELSE 0 END )"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'E'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 15"
        Sql = Sql & " AND A1.VLR_ATRIB <> CASE I.FLAG_SPC WHEN 'Y' THEN 1 ELSE 0 END )"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'R'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 16"
        Sql = Sql & " AND A1.VLR_ATRIB = CASE I.FLAG_VIP WHEN 'Y' THEN 1 ELSE 0 END )"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'E'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 16"
        Sql = Sql & " AND A1.VLR_ATRIB <> CASE I.FLAG_VIP WHEN 'Y' THEN 1 ELSE 0 END )"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'R'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 17"
        Sql = Sql & " AND A1.VLR_ATRIB = I.GWT_CODE)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'E'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 17"
        Sql = Sql & " AND A1.VLR_ATRIB <> I.GWT_CODE)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'R'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 19"
        Sql = Sql & " AND A1.VLR_ATRIB = I.FINALITY)"
        Sql = Sql & " OR (    A1.STATUS_ATRIB = 'E'"
        Sql = Sql & " AND A1.AUTONUM_ATR = 19"
        Sql = Sql & " AND A1.VLR_ATRIB <> I.FINALITY) "
        Sql = Sql & " OR (A1.STATUS_ATRIB = 'R' AND A1.AUTONUM_ATR = 21 AND A1.VLR_ATRIB =  CASE  RTRIM(LTRIM(I.IMO1)) WHEN '' THEN 0 ELSE 1 END ) OR (    A1.STATUS_ATRIB = 'E' AND A1.AUTONUM_ATR = 21 AND A1.VLR_ATRIB <> CASE   RTRIM(LTRIM(I.IMO1)) WHEN '' THEN 0 ELSE 1 END ) )"
        
        tbregras.Open Sql, db, 1, 1
        If Not tbregras.EOF Then
            If Nnull(tbregras!Mensagem, 1) <> "OK" Then
                Frmmensagem.Label1.Caption = "Regra de pátio não permite este local! " & Nnull(tbregras!Mensagem, 1)
                Frmmensagem.Show vbModal
                'MsgBox "Esta pilha já contém cntr de 20 pés", vbCritical
                Exit Sub
            End If
        End If
        tbregras.Close
        
        
        If idPatio > 0 And MskCntr.Tag = "I" And Valida Then
            If ColAlturaD = "1" Then
                Sql = "SELECT 1 "
                Sql = Sql & " From "
                Sql = Sql & " SGIPA..TB_CNTR_BL C "
                Sql = Sql & " Where"
                Sql = Sql & " C.AUTONUM=" & idPatio
                Sql = Sql & " AND C.FLAG_CNTR_BLOQ_CHAO_PATIO=1 "
                Set TB_BC = New ADODB.Recordset
                TB_BC.Open Sql, db, 1, 1
                If Not TB_BC.EOF Then
                    Col_Mensagem = "Atenção, existe Bloqueio de chão cadastrado para o Conteiner ! "
                    Col_Imagem = 3
                    frmMensagem2.Show 1
                    If Col_Cancelado Then Exit Sub
                End If
                TB_BC.Close
            End If
        End If
        
    End If

    
    
    tb1.Close





End If


    
'Se saiu de posicao valida nao pode ter conteiner acima
Sql = "SELECT V.VALIDA,A.YARD FROM " & Banco_Operador & "VW_INVENT_SISTEMAS A "
Sql = Sql & " INNER JOIN " & Banco_Operador & "TB_YARD V ON A.PATIO=V.PATIO AND A.YARD=V.YARD "
Sql = Sql & " Where a.Patio = " & Col_Patio & " AND A.ID_CONTEINER='" & MskCntr & "'"
Set tbant = New ADODB.Recordset
tbant.Open Sql, db, 1, 1
If Not tbant.EOF Then
    If Nnull(tbant!Valida, 0) = 1 Then
        If Col_Tam_Pos_Yard = 6 Then
            ColQuadraO$ = Left(tbant!Yard, 1)
            ColFiadaO$ = Mid$(tbant!Yard, 2, 2)
            ColLastroO$ = Mid$(tbant!Yard, 4, 2)
            ColAlturaO$ = Right(tbant!Yard, 1)
        Else
            ColQuadraO$ = Left(tbant!Yard, 2)
            ColFiadaO$ = Mid$(tbant!Yard, 3, 2)
            ColLastroO$ = Mid$(tbant!Yard, 5, 2)
            ColAlturaO$ = Right(tbant!Yard, 1)
        End If
        ColPilhaO$ = ColQuadraO$ + ColFiadaO$ + ColLastroO$
    
        Sql = "SELECT AUTONUM, ID_CONTEINER FROM " & Banco_Operador & "VW_INVENT_SISTEMAS WHERE patio=" & Col_Patio & " AND YARD='" & ColPilhaO & Trim(str(Val(ColAlturaO + 1))) & "'"
        Set tb1 = New ADODB.Recordset
        tb1.Open Sql, db, 1, 1
        If Not tb1.EOF Then
             Frmmensagem.Label1.Caption = "Existe um conteiner acima, impossivel movimentar ( " & tb1!Id_Conteiner & " )"
             Frmmensagem.Show vbModal
            'MsgBox "Existe um conteiner acima, impossivel movimentar ( " & tb1!Id_Conteiner & " )", vbExclamation
            Exit Sub
        End If
        tb1.Close
    End If
End If


''Verifica a obrigatoriedade de empilhadeira e carreta
'Dim SolEq As Boolean
'SolEq = False
'
'sql = "SELECT FLAG_TRUCK_MOV_COLETOR FROM OPERADOR.TB_PATIOS WHERE AUTONUM=1"
'Set tbpt = New ADODB.Recordset
'tbpt.Open sql, DB, 1, 1
'If Not tbpt.EOF Then
'    If Nnull(tbpt!flag_truck_mov_coletor, 0) = 1 Then
'        SolEq = True
'    End If
'End If
'tbpt.Close
'
'If SolEq = False Then
'    If Me.Grid2.TextMatrix(Grid2.Row, 3) = "TCD1" Then
'        SolEq = True
'    End If
'    If Me.Grid2.TextMatrix(Grid2.Row, 3) = "TCD2" Then
'        SolEq = True
'    End If
'    If Me.Grid2.TextMatrix(Grid2.Row, 3) = "TCD3" Then
'        SolEq = True
'    End If
'    If mskDestino = "NAVIO" Then
'        SolEq = True
'    End If
'End If
'
'
'If SolEq Then
'    If Me.dcCarreta = "" Then
'        MsgBox "É necessário informar a frota para este movimento", vbExclamation
'        Me.dcCarreta.SetFocus
'        Exit Sub
'    End If
'    If Me.dcEmp = "" Then
'        MsgBox "É necessário informar a frota para este movimento", vbExclamation
'        Me.dcEmp.SetFocus
'        Exit Sub
'    End If
'End If



On Error GoTo Trata_Erro_Atualizacao


db.BeginTrans

SQL2 = "INSERT INTO " & Banco_Operador & "TB_HIST_SHIFTING ("
If Banco_Em_Uso = "ORACLE" Then SQL2 = SQL2 & "AUTONUM,"
SQL2 = SQL2 & "CNTR,ORIGEM,DESTINO,DATA,TIPO,USUARIO,MOTIVO"
'SQL2 = SQL2 & ",ID_TRANSPORTADORA,AUTONUM_FROTA_CARRETA,AUTONUM_FROTA_EMPILHADEIRA"
SQL2 = SQL2 & ") VALUES "
SQL2 = SQL2 & "("
If Banco_Em_Uso = "ORACLE" Then SQL2 = SQL2 & Banco_Operador & "SEQ_HIST_SHIFTING.NEXTVAL,"
SQL2 = SQL2 & idPatio & ","
SQL2 = SQL2 & "'" & TXTYardAtual & "',"
SQL2 = SQL2 & "'" & mskDestino & "',"
SQL2 = SQL2 & "GETDATE(),"
SQL2 = SQL2 & "'" & MskCntr.Tag & "',"
If CodEmpSis = "BAN" Then
    SQL2 = SQL2 & Cod_Usuario
Else
    SQL2 = SQL2 & nUser
End If
SQL2 = SQL2 & "," & DataCombo2.BoundText
'SQL2 = SQL2 & ",0"
'SQL2 = SQL2 & ",0"
'SQL2 = SQL2 & ",0"
SQL2 = SQL2 & ")"

Dim SQL3 As String
SQL3 = ""

Dim FlagDivLacre As Boolean
FlagDivLacre = False

If MskCntr.Tag = "I" Then


    Sql = "UPDATE " & Banco_Sgipa & "tb_cntr_bl SET "
    Sql = Sql & " YARD='" & mskDestino & "'"
    'If Me.txtISO.text <> "" Then Sql = Sql & " ,ISO='" & Me.txtISO.text & "'"
    'If Me.txtMaxGross.text <> "" Then Sql = Sql & " ,MAXGROSS=" & Val(Me.txtMaxGross.text)
    Sql = Sql & " ,TARA=" & Val(Me.txtTara.text)
    Sql = Sql & " ,TARA_REAL_EST='R'"
    Sql = Sql & " ,LACRE_IPA='" & Me.txtLacreIPA.text & "'"
    Sql = Sql & " ,LACRE1_GATE='" & Me.txtLACRE1.text & "'"
    Sql = Sql & " ,LACRE2_GATE='" & Me.txtLACRE2.text & "'"
    Sql = Sql & " ,LACRE3_GATE='" & Me.txtLACRE3.text & "'"
    Sql = Sql & " ,CAMERA_POSICIONAMENTO="
    If DataCombo4.BoundText = "" Then
        Sql = Sql & "NULL"
    Else
        Sql = Sql & Me.DataCombo4.BoundText
    End If
    Sql = Sql & " WHERE Autonum=" & idPatio
    
    SQL3 = "UPDATE " & Banco_Operador & "tb_GATE_DADOS SET "
    SQL3 = SQL3 & " LACREIPA='" & Me.txtLacreIPA.text & "'"
    SQL3 = SQL3 & " ,LACRE1='" & Me.txtLACRE1.text & "'"
    SQL3 = SQL3 & " ,LACRE2='" & Me.txtLACRE2.text & "'"
    SQL3 = SQL3 & " ,LACRE3='" & Me.txtLACRE3.text & "'"
    SQL3 = SQL3 & " WHERE CONTEINER='" & Me.MskCntr.text & "' "
    SQL3 = SQL3 & " AND GMCI='" & Me.TXTgmci.text & "' "
    
    'rotina de divergencia de lacre alterada em 24/02/17.
    'se apenas 1 dos lacres estiver correto, nao gerar divergencia
    FlagDivLacre = True
'    If (FlagDivLacre) And (UCase(Trim(Me.txtLACRE1.text)) = UCase(Trim(Me.txtLACRE1.Tag)) Or UCase(Trim(Me.txtLACRE1.text)) = UCase(Trim(Me.txtLACRE2.Tag)) Or UCase(Trim(Me.txtLACRE1.text)) = UCase(Trim(Me.txtLACRE3.Tag))) Then FlagDivLacre = False
'    If (FlagDivLacre) And (UCase(Trim(Me.txtLACRE2.text)) = UCase(Trim(Me.txtLACRE1.Tag)) Or UCase(Trim(Me.txtLACRE2.text)) = UCase(Trim(Me.txtLACRE2.Tag)) Or UCase(Trim(Me.txtLACRE2.text)) = UCase(Trim(Me.txtLACRE3.Tag))) Then FlagDivLacre = False
'    If (FlagDivLacre) And (UCase(Trim(Me.txtLACRE3.text)) = UCase(Trim(Me.txtLACRE1.Tag)) Or UCase(Trim(Me.txtLACRE3.text)) = UCase(Trim(Me.txtLACRE2.Tag)) Or UCase(Trim(Me.txtLACRE3.text)) = UCase(Trim(Me.txtLACRE3.Tag))) Then FlagDivLacre = False
    
    Dim Lacres_Manif As String
    Lacres_Manif = ""
    If Trim(Me.txtLACRE1.Tag) <> "" Then Lacres_Manif = Lacres_Manif & "***" & Trim(Me.txtLACRE1.Tag)
    If Trim(Me.txtLACRE2.Tag) <> "" Then Lacres_Manif = Lacres_Manif & "***" & Trim(Me.txtLACRE2.Tag)
    If Trim(Me.txtLACRE3.Tag) <> "" Then Lacres_Manif = Lacres_Manif & "***" & Trim(Me.txtLACRE3.Tag)
    Lacres_Manif = UCase(Lacres_Manif)
    
    If (FlagDivLacre) And UCase(Trim(Me.txtLACRE1.text)) <> "" And InStr(1, Lacres_Manif, UCase(Trim(Me.txtLACRE1.text))) <> 0 Then FlagDivLacre = False
    If (FlagDivLacre) And UCase(Trim(Me.txtLACRE2.text)) <> "" And InStr(1, Lacres_Manif, UCase(Trim(Me.txtLACRE2.text))) <> 0 Then FlagDivLacre = False
    If (FlagDivLacre) And UCase(Trim(Me.txtLACRE3.text)) <> "" And InStr(1, Lacres_Manif, UCase(Trim(Me.txtLACRE3.text))) <> 0 Then FlagDivLacre = False
    
    '----------------------------------------------------
    
    
ElseIf MskCntr.Tag = "A" Then
    Sql = "UPDATE " & Banco_Sgipa & "tb_armazens_ipa SET "
    Sql = Sql & " YARD='" & mskDestino & "'"
 
    Sql = Sql & " WHERE Autonum=" & idPatio
Else

    Sql = "UPDATE redex.tb_patio SET "
    Sql = Sql & " YARD='" & mskDestino & "'"
    Sql = Sql & " WHERE Autonum_patio=" & idPatio
End If

Executa Sql
Executa SQL2
If SQL3 <> "" Then
    Executa SQL3
End If


If FlagDivLacre = True Then
    Dim Xlocal As String, Xtipo As String, XComplemento As String
    Xlocal = "12": Xtipo = "32": XComplemento = "8"
        
    Sql = "select autonum from " & Banco_Sgipa & "tb_avarias_conteiner where"
    Sql = Sql & " local = '" & Nnull(Xlocal, 1) & "'"
    Sql = Sql & " and tipo = '" & Nnull(Xtipo, 1) & "'"
    Sql = Sql & " and complemento = '" & Nnull(XComplemento, 1) & "'"
    Sql = Sql & " and cntr = " & idPatio
    
    PRSet Rs, Sql
    If Rs.EOF Then
        Sql = "INSERT INTO " & Banco_Sgipa & "TB_AVARIAS_CONTEINER ("
        Sql = Sql & "CNTR,LOCAL,TIPO,COMPLEMENTO,COMPL_IPA,PROVIDENCIA) "
        Sql = Sql & "VALUES ("
        Sql = Sql & idPatio & ",'" & Nnull(Xlocal, 1) & "','" & Nnull(Xtipo, 1) & "','" & Nnull(XComplemento, 1) & "','" & Left(wDiverg, 35) & "',11)"
        Executa Sql
    End If
    Rs.Close
End If


Dim Segreg$
Segreg = Verifica_Regras_Seg_Imo_Delta(MskCntr, mskDestino)
If Segreg <> "" Then
    db.RollbackTrans
    Frmmensagem.Label1.Caption = "SEGREGACAO IMO !!!! " & Segreg
    Frmmensagem.Show vbModal
    Exit Sub
End If

'Verifica se local pode ter IMO

Dim AlocaImo$
AlocaImo = Valida_Aloca_Imo(MskCntr, mskDestino, Val(Me.txtTam), 1)
If AlocaImo <> "" Then
    db.RollbackTrans
    Frmmensagem.Label1.Caption = AlocaImo
    Frmmensagem.Show vbModal
    Exit Sub
End If



db.CommitTrans


Limpa_Dados
Me.CNTR = "___-_"
Me.MskCntr = "__________-_"
CNTR.SetFocus

On Error GoTo 0
Exit Sub

Trata_Erro_Atualizacao:
Frmmensagem.Label1.Caption = "Erro durante atualização de posição" & vbCr & "Comunique o TI" & vbCr & Error$
Frmmensagem.Show vbModal
'MsgBox "Erro durante atualização de posição" & vbCr & "Comunique o TI" & vbCr & Error$
db.RollbackTrans


End Sub


Sub Busca_Cntr()

    If CNTR.ClipText = "" Then Exit Sub
    MskCntr = "__________-_"
    
    Dim Regs As Integer
'    sql = "SELECT COUNT(0) FROM " & Banco_Operador & "VW_INVENT_SISTEMAS WHERE PATIO=4 AND SUBSTR(ID_CONTEINER,8,5)='" & CNTR & "'"
    Sql = "SELECT COUNT(1) FROM " & Banco_Operador & "VW_INVENT_SISTEMAS WHERE PATIO=" & Col_Patio & " AND SUBSTR(ID_CONTEINER,8,5)='" & CNTR & "'"
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
'        sql = "select id_conteiner from " & Banco_Operador & "VW_INVENT_SISTEMAS where patio=4 and SUBSTR(ID_CONTEINER,8,5)='" & CNTR & "'"
        Sql = "select id_conteiner from " & Banco_Operador & "VW_INVENT_SISTEMAS where patio=" & Col_Patio & " and SUBSTR(ID_CONTEINER,8,5)='" & CNTR & "'"
        Sql = CONVERTE_SQL(Sql)
        MskCntr = Nnull(db.Execute(Sql).Fields(0), 1)
    Else
        Frmmensagem.Label1.Caption = "Existe mais de 1 conteiner com este final. Favor informar a identificação completa"
        Frmmensagem.Show vbModal
        'MsgBox "Existe mais de 1 conteiner com este final. Favor informar a identificação completa", vbExclamation
        CNTR = "___-_"
        MskCntr.SetFocus
        Exit Sub
    End If
    CNTR = "___-_"
    
    If MskCntr.ClipText <> "" Then Call Busca_Dados
    

End Sub


Sub Busca_Dados()

idPatio = 0
If MskCntr.ClipText = "" Then Exit Sub

Label1(2) = MskCntr
MskCntr.Tag = ""

Limpa_Dados
'GoTo pula
Sql = "select autonum, ef, nome, viagem, tamanho, tipobasico, data_ent_temp,yard,pod,imo1,lacre1,lacre2,lacre3,lacre4,lacre5,lacre6,lacre7,bruto, sistema, flag_reefer, FLAG_REEFER_DESLIGADO, FLAG_DTA, FLAG_DESTINADO_DESOVA,INDICADOR "
Sql = Sql & " from " & Banco_Operador & "VW_INVENT_SISTEMAS where patio=" & Col_Patio & " and id_conteiner='" & MskCntr & "'"
Set tbvw = New ADODB.Recordset
Sql = CONVERTE_SQL(Sql)
tbvw.Open Sql, db, 1, 1

If Not tbvw.EOF Then
    idPatio = tbvw!AUTONUM
    Me.txtEF = tbvw!EF
    Me.txtGateIN = Format(tbvw!data_ent_temp, "dd/mm/yyyy")
    Me.TxtNavio = Nnull(tbvw!Nome, 1) & " " & Nnull(tbvw!Viagem, 1)
    Me.txtPOD = Nnull(tbvw!pod, 1)
    Me.txtBruto = Nnull(tbvw!Bruto, 1)
    Me.txtLacres = Nnull(tbvw!Lacre1, 1) & " " & Nnull(tbvw!Lacre2, 1) & " " & Nnull(tbvw!Lacre3, 1) & " " & Nnull(tbvw!Lacre4, 1) & " " & Nnull(tbvw!lacre5, 1) & " " & Nnull(tbvw!lacre6, 1) & " " & Nnull(tbvw!lacre7, 1)
    Me.txtIMO = Nnull(tbvw!IMO1, 1)
    Me.txtTam = Nnull(tbvw!Tamanho, 0)
    Me.txtTipo = Nnull(tbvw!tipobasico, 1)
    Me.TXTYardAtual = Nnull(tbvw!Yard, 1)
    Me.MskCntr.Tag = Nnull(tbvw!Sistema, 1)
    Me.lblReeferDesligado.Caption = ""
    
    If tbvw!flag_reefer_desligado = "Y" Then
        Me.lblReeferDesligado.Visible = True
        Me.lblReeferDesligado.ForeColor = vbRed
        Me.lblReeferDesligado.Caption = "REEFER DESLIGADO"
    End If
    
    If tbvw!flag_reefer_desligado = "N" And tbvw!flag_reefer = "Y" Then
        Me.lblReeferDesligado.Visible = True
        Me.lblReeferDesligado.ForeColor = vbBlue
        Me.lblReeferDesligado.Caption = "REEFER LIGADO"
    End If
    
    
    If tbvw!flag_DTA = "Y" Then
        Me.LBLDTA.Visible = True
    End If
    
    If tbvw!flag_DESTINADO_DESOVA = 1 Then
        Me.lblDestinadoDesova.Visible = True
    End If
    
    
    Dim DescSigVig As String
    DescSigVig = ""
    
    Set TBCNTR = New ADODB.Recordset
    TBCNTR.Open "SELECT TARA,ISO,MAXGROSS,LACRE_ORIGEM,LACRE2,LACRE3,LACRE_IPA,LACRE1_GATE,LACRE2_GATE,LACRE3_GATE,GMCI,descricao_sigvig FROM sgipa..tb_cntr_bl where autonum=" & idPatio, db, 1, 1
    If Not TBCNTR.EOF Then
        Me.txtTara.text = Nnull(TBCNTR!Tara, 0)
        Me.txtISO.text = Nnull(TBCNTR!ISO, 1)
        Me.txtMaxGross.text = Nnull(TBCNTR!maxgross, 0)
        Me.txtLacreIPA.text = Nnull(TBCNTR!Lacre_ipa, 1)
        
        Me.txtLACRE1.text = Nnull(TBCNTR!lacre1_gate, 1)
        Me.txtLACRE1.Tag = Nnull(TBCNTR!lacre_origem, 1)
        
        Me.txtLACRE2.text = Nnull(TBCNTR!lacre2_gate, 1)
        Me.txtLACRE2.Tag = Nnull(TBCNTR!Lacre2, 1)
        
        Me.txtLACRE3.text = Nnull(TBCNTR!lacre3_gate, 1)
        Me.txtLACRE3.Tag = Nnull(TBCNTR!Lacre3, 1)
        
        Me.TXTgmci.text = Nnull(TBCNTR!GMCI, 1)
        DescSigVig = Nnull(TBCNTR!descricao_sigvig, 1)
        
    End If
    TBCNTR.Close
    
    If Nnull(tbvw!Indicador, 1) <> "" Then
        Sql = "SELECT AUTONUM FROM SGIPA..TB_CAD_PARCEIROS WHERE FANTASIA='" & Nnull(tbvw!Indicador, 1) & "' AND FLAG_ALERTA_COL_INDICADOR=1 "
        Set TBAL = New ADODB.Recordset
        TBAL.Open Sql, db, 1, 1
        If Not TBAL.EOF Then
            Label1(9).Caption = "CARGA ESPECIAL : " & Nnull(tbvw!Indicador, 1)
            Label1(9).Visible = True
        End If
        TBAL.Close
    End If
    
    
    
   
    If DescSigVig = "" Then
        Sql = "select vw.descricao from sgipa.dbo.VW_SIGVIG_CNTR_ESTOQUE vw "
        Sql = Sql & " where vw.CNTR = " & idPatio
        DescSigVig = Nnull(db.Execute(Sql).Fields(0), 1)
    End If
    
    If DescSigVig = "" Then
        WSQL = Empty
        WSQL = "SELECT sit.DESCRICAO AS situacao"
        WSQL = WSQL & " FROM dbo.TB_RET_UN_SIGVIG AS un"
        WSQL = WSQL & " INNER JOIN dbo.TB_SITUACAO_CARGA_SIGVIG AS sit ON un.COD_SIT = sit.AUTONUM"
        WSQL = WSQL & " INNER JOIN dbo.TB_SIGVIG_PACOTE AS pct ON un.AUTONUM_PCT = pct.AUTONUM"
        WSQL = WSQL & " INNER JOIN dbo.TB_CNTR_BL AS c ON un.ID_CONTEINER = c.ID_CONTEINER"
        WSQL = WSQL & "  LEFT JOIN dbo.TB_AMR_CNTR_BL AS a ON c.AUTONUM = a.CNTR AND pct.LOTE = a.BL"
        WSQL = WSQL & "  LEFT JOIN dbo.TB_BL AS bl ON pct.LOTE = bl.AUTONUM"
        WSQL = WSQL & "  LEFT JOIN dbo.TB_SITUACAO_BL_SIGVIG AS bls ON pct.SITUACAO_BL = bls.AUTONUM"
        WSQL = WSQL & "  LEFT OUTER JOIN dbo.TB_MSG_UN_SIGVIG AS msg ON un.AUTONUM = msg.AUTONUM_RET_UN"
        WSQL = WSQL & " Where C.autonum=" & idPatio
        WSQL = WSQL & " ORDER BY PCT.DATA_RETORNO DESC"
            
        DescSigVig = Nnull(db.Execute(WSQL).Fields(0), 1)
        
    End If
    
    
    Me.txtSIGVIG.text = DescSigVig
    
Else
    Frmmensagem.Label1.Caption = "Unidade não encontrada no estoque "
    Frmmensagem.Show vbModal
    'MsgBox "Unidade não encontrada no estoque"
    MskCntr = "__________-_"
    MskCntr.SetFocus
    Exit Sub
End If
tbvw.Close

If MskCntr.Tag = "R" Then
    Sql = "select hs.*,usu.usuario as nmusuario, cm.descricao as desc_motivo"
    Sql = Sql & " from redex.tb_patio cc"
    Sql = Sql & " inner join " & Banco_Operador & "tb_hist_shifting hs on cc.autonum_patio = hs.cntr and hs.tipo='R'"
    Sql = Sql & " left  join redex.tb_cad_usuarios usu on hs.usuario = usu.autonum_usu"
    Sql = Sql & " left join " & Banco_Operador & "tb_cad_motivo cm on hs.motivo = cm.autonum"
    Sql = Sql & " where cc.id_conteiner='" & MskCntr & "'"
    Sql = Sql & " order by hs.data"
ElseIf Me.MskCntr.Tag = "A" Then
    Sql = "select hs.*,usu.usuario as nmusuario, cm.descricao as desc_motivo"
    Sql = Sql & " from " & Banco_Sgipa & "tb_armazens_ipa cc"
    Sql = Sql & " inner join " & Banco_Operador & "tb_hist_shifting hs on cc.autonum = hs.cntr and hs.tipo='A'"
    Sql = Sql & " left  join " & Banco_Sgipa & "tb_cad_usuarios usu on hs.usuario = usu.autonum"
    Sql = Sql & " left join " & Banco_Operador & "tb_cad_motivo cm on hs.motivo = cm.autonum"
    Sql = Sql & " where cc.descr='" & MskCntr & "'"
    Sql = Sql & " order by hs.data"

Else

    Sql = "select hs.*,usu.usuario as nmusuario, cm.descricao as desc_motivo"
    Sql = Sql & " from " & Banco_Sgipa & "tb_cntr_bl cc"
    Sql = Sql & " inner join " & Banco_Operador & "tb_hist_shifting hs on cc.autonum = hs.cntr and hs.tipo='I'"
    Sql = Sql & " left  join " & Banco_Sgipa & "tb_cad_usuarios usu on hs.usuario = usu.autonum"
    Sql = Sql & " left join " & Banco_Operador & "tb_cad_motivo cm on hs.motivo = cm.autonum"
    Sql = Sql & " where cc.id_conteiner='" & MskCntr & "'"
    Sql = Sql & " order by hs.data"
End If

Adodc1.ConnectionString = ConnectionStringAdo
Adodc1.RecordSource = UCase(Sql)
Adodc1.ConnectionString = ConnectionStringAdo: Adodc1.Refresh
Set Adodc1.Recordset.ActiveConnection = Nothing

pula:
Set RsAmr = New ADODB.Recordset
Sql = " select tb_bl.autonum as lote, tb_bl.numero as Bl"
Sql = Sql & " from tb_bl, tb_amr_cntr_bl ,tb_cntr_bl"
Sql = Sql & " where tb_bl.autonum= tb_amr_cntr_bl.bl"
Sql = Sql & " and TB_AMR_CNTR_BL.CNTR=TB_CNTR_BL.AUTONUM "
Sql = Sql & " and tb_cntr_bl.id_conteiner='" & MskCntr & "'"
Sql = Sql & " and tb_bl.flag_ativo=1"

RsAmr.Open Sql, db
Set DataGrid2.DataSource = RsAmr

Label13.Caption = MskCntr

Text1(0) = ""

'busca informacoes complementares
IdCntr = Nnull(db.Execute("select max(autonum) from " & Banco_Sgipa & "tb_cntr_bl where id_conteiner='" & MskCntr & "'").Fields(0), 0)
If IdCntr <> 0 Then
    'verifica GR
'    Sql = "select count(*) from " & Banco_Sgipa & "tb_gr_bl gr"
'    Sql = Sql & " inner join " & Banco_Sgipa & "tb_amr_cntr_bl amr on gr.bl=amr.bl"
'    Sql = Sql & " where amr.cntr=" & IdCntr
'    If Nnull(db.Execute(Sql).Fields(0), 0) = 0 Then
'        Text1 = "NAO CONSTA GR" & vbCrLf
'    Else
'        Sql = "select count(*) from " & Banco_Sgipa & "tb_gr_bl gr"
'        Sql = Sql & " inner join " & Banco_Sgipa & "tb_amr_cntr_bl amr on gr.bl=amr.bl"
'        Sql = Sql & " where amr.cntr=" & IdCntr & " and gr.flag_gr_paga=0"
'        If Nnull(db.Execute(Sql).Fields(0), 0) = 0 Then
'            Text1 = "PENDENCIA DE RECEBIMENTO DE GR" & vbCrLf
'        Else
'            Text1 = "GR PAGA" & vbCrLf
'        End If
'    End If
    
    
    
    
    
    
    
    'verifica bloqueio CNTR
    bloqueio = False
    Sql = "select flag_bloqueio_manual, flag_bloqueio from tb_cntr_bl WHERE autonum=" & IdCntr
    PRSet Rs, Sql
    If Nnull(Rs!FLAG_BLOQUEIO_MANUAL, 0) = 1 Then bloqueio = True: Text1(0) = Text1(0) & "BLOQUEIO MANUAL DE CNTR" & vbCrLf
    If Nnull(Rs!FLAG_BLOQUEIO, 0) = 1 Then bloqueio = True: Text1(0) = Text1(0) & "BLOQUEIO ALFANDEGA (CNTR)" & vbCrLf
    Rs.Close
    
    'verifica bloqueio bl
    Sql = "select count(*) from " & Banco_Sgipa & "tb_bl bl"
    Sql = Sql & " inner join " & Banco_Sgipa & "tb_amr_cntr_bl amr on bl.autonum=amr.bl"
    Sql = Sql & " where amr.cntr=" & IdCntr & " and bl.flag_bloqueio_manual=1"
    If Nnull(db.Execute(Sql).Fields(0), 0) <> 0 Then
        bloqueio = True: Text1(0) = Text1(0) & "BLOQUEIO MANUAL DE BL" & vbCrLf
    End If

    Sql = "select count(*) from " & Banco_Sgipa & "tb_bl bl"
    Sql = Sql & " inner join " & Banco_Sgipa & "tb_amr_cntr_bl amr on bl.autonum=amr.bl"
    Sql = Sql & " where amr.cntr=" & IdCntr & " and bl.flag_bloqueio=1"
    If Nnull(db.Execute(Sql).Fields(0), 0) <> 0 Then
        bloqueio = True: Text1(0) = Text1(0) & "BLOQUEIO ALFANDEGA DE BL" & vbCrLf
    End If
    If Not bloqueio Then Text1(0) = Text1(0) & "NAO CONSTA BLOQUEIO" & vbCrLf
    
    
    'VER REGISTRO
    Sql = "select count(*) from " & Banco_Sgipa & "tb_registro_saida_cntr sc"
    Sql = Sql & " inner join " & Banco_Sgipa & "tb_ordem_carregamento oc on sc.ordem_carreg = oc.autonum"
    Sql = Sql & " where sc.cntr=" & IdCntr & " and isnull(oc.flag_saida,0)=0"
    If Nnull(db.Execute(Sql).Fields(0), 0) <> 0 Then
        Text1(0) = Text1(0) & "CARREGAMENTO REGISTRADO" & vbCrLf
        Label1(9).Caption = "CARREGAMENTO REGISTRADO"
        Label1(9).Visible = True
    Else
        Text1(0) = Text1(0) & "NAO CONSTA REGISTRO" & vbCrLf
    End If
    
    'DOCUMENTO
    Sql = "SELECT td.descr, bl.dt_desembaraco, bl.flag_desembaraco"
    Sql = Sql & " FROM " & Banco_Sgipa & "TB_BL BL"
    Sql = Sql & " INNER JOIN " & Banco_Sgipa & "tb_tipos_documentos td on bl.tipo_documento=td.code"
    Sql = Sql & " inner join " & Banco_Sgipa & "tb_amr_cntr_bl amr on amr.bl=bl.autonum"
    Sql = Sql & " where amr.cntr=" & IdCntr
    PRSet Rs, Sql
    If Not Rs.EOF Then
        Text1(0) = Text1(0) & "DOCUMENTO : " & Rs!Descr & vbCrLf
    End If
    
    If Nnull(Rs!flag_desembaraco, 0) = 1 Then
        Text1(0) = Text1(0) & "Carga Desembaraçada em " & Format(Rs!dt_desembaraco, "dd/mm hh:mm") & vbCrLf
    End If
    
    Rs.Close
    
    
    'verifica motivo posicionamento
    Sql = "select mp.descr || ' em ' || to_char(ap.dt_prevista,'dd/mm/yyyy hh24:mi') as descr from tb_agendamento_posicao ap"
    Sql = Sql & " inner join tb_agenda_posicao_motivo apm on ap.autonum = apm.autonum_agenda_posicao"
    Sql = Sql & " inner join tb_motivo_posicao mp on apm.motivo_posicao = mp.code"
    Sql = Sql & " Where ap.CNTR = " & IdCntr & " And ap.cancelado = 0"
    Sql = Sql & " union all"
    Sql = Sql & " select 'Agendamento Desova em ' || to_char(dt_prevista,'dd/mm/yyyy hh24:mi') as descr from " & Banco_Sgipa & "tb_agendamento_desova where cntr=" & IdCntr
    
    PRSet Rs, Sql
    If Not Rs.EOF Then
        Text1(0) = Text1(0) & "Motivos de Posicionamento" & vbCrLf
        Do While Not Rs.EOF
            Text1(0) = Text1(0) & UCase(Rs!Descr) & vbCrLf
            Rs.MoveNext
        Loop
    Else
        Text1(0) = Text1(0) & "Não consta Agendamento" & vbCrLf
    End If
    Rs.Close
    
    
End If
    
    

Carrega_Grid1
Carrega_Grid2

mskDestino.SetFocus

End Sub
Private Sub Carrega_Grid1()
Grid1.Clear
Dim Sql As String
Sql = "SELECT MOTIVO,LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,108),5)) AS DT_PREV FROM OPERADOR..VW_PROX_MOV_CNTR WHERE ID_CONTEINER='" & Me.MskCntr & "'"
Sql = Sql & " UNION "
Sql = Sql & " SELECT MOTIVO,' ' AS DT_PREV FROM OPERADOR..VW_PROX_MOV_CNTR_COL WHERE ID_CONTEINER='" & Me.MskCntr & "'"

Adodc5.ConnectionString = ConnectionStringAdo
Adodc5.RecordSource = Sql
Adodc5.Refresh
Set Adodc5.Recordset.ActiveConnection = Nothing

Grid1.FormatString = "|<Motivo                    |<Data            "
End Sub
Private Sub Carrega_Grid2()

Grid2.Clear

Dim PosValida As Boolean
Dim PosBusca As String

Sql = ""

If Me.Option1(0).Value = True Then
    PosBusca = Me.TXTYardAtual.text
End If
If Me.Option1(1).Value = True Then
    PosBusca = Me.mskDestino.text
End If

If PosBusca <> "" Then

    Set TBval = New ADODB.Recordset
    TBval.Open "Select valida from operador..tb_yard where yard='" & PosBusca & "'", db, 1, 1
    If TBval.RecordCount <> 0 Then
    
        If TBval!Valida = 1 Then
            
            Sql = "SELECT ID_CONTEINER,YARD,MOTIVO,LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,108),5)),SUBSTRING(YARD,2,2) AS PILHA FROM OPERADOR..VW_PROX_MOV_CNTR WHERE PATIO=" & Col_Patio & " AND YARD LIKE "
            If Me.Option1(0).Value = True Then
                Sql = Sql & "'" & Mid$(PosBusca, 1, 3) & "%'"
            Else
                If Len(PosBusca) = 1 Then
                    Sql = Sql & "'" & Mid$(PosBusca, 1, 1) & "%'"
                ElseIf Len(PosBusca) = 2 Then
                    Sql = Sql & "'" & Mid$(PosBusca, 1, 2) & "%'"
                ElseIf Len(PosBusca) >= 3 Then
                    Sql = Sql & "'" & Mid$(PosBusca, 1, 3) & "%'"
                End If
            End If
            Sql = Sql & " and id_conteiner<>'" & Me.MskCntr & "' AND YARD NOT IN (SELECT YARD FROM OPERADOR..TB_YARD WHERE VALIDA=0) "
            'ORDER BY SUBSTRING(YARD,2,2), ID_CONTEINER "
            
            
            Sql = Sql & " UNION "
            Sql = Sql & " SELECT ID_CONTEINER,YARD,MOTIVO,LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,108),5)),SUBSTRING(YARD,2,2) AS PILHA FROM OPERADOR..VW_PROX_MOV_CNTR_COL WHERE PATIO=" & Col_Patio & " AND YARD LIKE "
            If Me.Option1(0).Value = True Then
                Sql = Sql & "'" & Mid$(PosBusca, 1, 3) & "%'"
            Else
                If Len(PosBusca) = 1 Then
                    Sql = Sql & "'" & Mid$(PosBusca, 1, 1) & "%'"
                ElseIf Len(PosBusca) = 2 Then
                    Sql = Sql & "'" & Mid$(PosBusca, 1, 2) & "%'"
                ElseIf Len(PosBusca) >= 3 Then
                    Sql = Sql & "'" & Mid$(PosBusca, 1, 3) & "%'"
                End If
            End If
            Sql = Sql & " and id_conteiner<>'" & Me.MskCntr & "' AND YARD NOT IN (SELECT YARD FROM OPERADOR..TB_YARD WHERE VALIDA=0) "
            Sql = Sql & " ORDER BY PILHA , ID_CONTEINER "
            
            
        Else
            Sql = "SELECT ID_CONTEINER,YARD,MOTIVO,LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,108),5)) FROM OPERADOR..VW_PROX_MOV_CNTR WHERE YARD='" & PosBusca & "' and id_conteiner<>'" & Me.MskCntr & "'"
            'ORDER BY ID_CONTEINER "
            Sql = Sql & " UNION "
            Sql = Sql & " SELECT ID_CONTEINER,YARD,MOTIVO,NULL FROM OPERADOR..VW_PROX_MOV_CNTR_COL WHERE YARD='" & PosBusca & "' and id_conteiner<>'" & Me.MskCntr & "'"
            Sql = Sql & " ORDER BY ID_CONTEINER "
            
        End If
        
    Else
    
        Sql = "SELECT ID_CONTEINER,YARD,MOTIVO,LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,108),5)),SUBSTRING(YARD,2,2) AS PILHA FROM OPERADOR..VW_PROX_MOV_CNTR WHERE PATIO=" & Col_Patio
        
        If Len(PosBusca) = 1 Then
            Sql = Sql & " AND YARD LIKE '" & Mid$(PosBusca, 1, 1) & "%'"
        ElseIf Len(PosBusca) = 2 Then
            Sql = Sql & " AND YARD LIKE '" & Mid$(PosBusca, 1, 2) & "%'"
        ElseIf Len(PosBusca) >= 3 Then
            If Val(Mid$(PosBusca, 2, 2)) <> 0 Then
                If Val(Mid$(PosBusca, 2, 2)) Mod 2 <> 0 Then
                    Sql = Sql & " AND (YARD LIKE '" & Mid$(PosBusca, 1, 3) & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 1) & Format(Val(Mid$(PosBusca, 2, 2)) + 1, "00") & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 1) & Format(Val(Mid$(PosBusca, 2, 2)) - 1, "00") & "%')"
                Else
                     Sql = Sql & " AND (YARD LIKE '" & Mid$(PosBusca, 1, 3) & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 1) & Format(Val(Mid$(PosBusca, 2, 2)) + 1, "00") & "%'  OR YARD LIKE '" & Mid$(PosBusca, 1, 1) & Format(Val(Mid$(PosBusca, 2, 2)) + 2, "00") & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 1) & Format(Val(Mid$(PosBusca, 2, 2)) - 2, "00") & "%'  )"
                End If
            Else
                Sql = Sql & " AND YARD LIKE '" & Mid$(PosBusca, 1, 3) & "%'"
            End If
        End If
        Sql = Sql & " and id_conteiner<>'" & Me.MskCntr & "' "
        'ORDER BY SUBSTRING(YARD,2,2), ID_CONTEINER "
        
        
        Sql = Sql & " UNION "
        
        Sql = Sql & " SELECT ID_CONTEINER,YARD,MOTIVO,NULL,SUBSTRING(YARD,2,2) AS PILHA FROM OPERADOR..VW_PROX_MOV_CNTR_COL WHERE PATIO=" & Col_Patio
        
        If Len(PosBusca) = 1 Then
            Sql = Sql & " AND YARD LIKE '" & Mid$(PosBusca, 1, 1) & "%'"
        ElseIf Len(PosBusca) = 2 Then
            Sql = Sql & " AND YARD LIKE '" & Mid$(PosBusca, 1, 2) & "%'"
        ElseIf Len(PosBusca) >= 3 Then
            If Val(Mid$(PosBusca, 2, 2)) <> 0 Then
                If Val(Mid$(PosBusca, 2, 2)) Mod 2 <> 0 Then
                    Sql = Sql & " AND (YARD LIKE '" & Mid$(PosBusca, 1, 3) & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 1) & Format(Val(Mid$(PosBusca, 2, 2)) + 1, "00") & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 1) & Format(Val(Mid$(PosBusca, 2, 2)) - 1, "00") & "%')"
                Else
                     Sql = Sql & " AND (YARD LIKE '" & Mid$(PosBusca, 1, 3) & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 1) & Format(Val(Mid$(PosBusca, 2, 2)) + 1, "00") & "%'  OR YARD LIKE '" & Mid$(PosBusca, 1, 1) & Format(Val(Mid$(PosBusca, 2, 2)) + 2, "00") & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 1) & Format(Val(Mid$(PosBusca, 2, 2)) - 2, "00") & "%'  )"
                End If
            Else
                Sql = Sql & " AND YARD LIKE '" & Mid$(PosBusca, 1, 3) & "%'"
            End If
        End If
        'Sql = Sql & " and id_conteiner<>'" & Me.MskCntr & "' ORDER BY SUBSTRING(YARD,2,2), ID_CONTEINER "
        Sql = Sql & " and id_conteiner<>'" & Me.MskCntr & "' ORDER BY PILHA , ID_CONTEINER "
        
    End If
        
    Adodc6.ConnectionString = ConnectionStringAdo
    Adodc6.RecordSource = Sql
    Adodc6.Refresh
    Set Adodc6.Recordset.ActiveConnection = Nothing
       
    Grid2.FormatString = "|<Conteiner         |^Posição |<Motivo                      |<Data          "

    TBval.Close
End If

End Sub

Private Sub Limpa_Dados()

mskDestino = ""
Label2 = ""
DataCombo1.BoundText = ""
DataCombo2.BoundText = ""
DataCombo3.BoundText = ""
Me.txtEF = ""
Me.txtGateIN = ""
Me.TxtNavio = ""
Me.txtPOD = ""
Me.txtTam = ""
Me.txtTipo = ""
Me.TXTYardAtual = ""
Me.txtBruto = ""
Me.txtLacres = ""
Me.txtIMO = ""
Grid1.Clear
Grid2.Clear
Me.Text1(0) = ""
Label1(9).Visible = False
Label1(9).Caption = ""
lblReeferDesligado.Visible = False
lblDestinadoDesova.Visible = False
LBLDTA.Visible = False
Me.txtISO.text = ""
Me.txtTara.text = ""
Me.txtMaxGross.text = ""
Me.txtLACRE1.text = ""
Me.txtLACRE1.Tag = ""
Me.txtLACRE2.text = ""
Me.txtLACRE2.Tag = ""
Me.txtLACRE3.text = ""
Me.txtLACRE3.Tag = ""
Me.txtLacreIPA.text = ""
Me.TXTgmci.text = ""
Me.Check1.Value = 0

End Sub






Private Sub CNTR_Change()
If InStr(1, Me.CNTR, "_") = 0 Then
    Call Busca_Cntr
End If
End Sub

Private Sub CNTR_LostFocus()
'Call Busca_Cntr
End Sub


Private Sub Command1_Click()

    
If DataCombo2.BoundText = "" Then
    MsgBox "Informe o motivo"
    Exit Sub
End If
Sql = "select count(*) from operador.dbo.tb_cad_motivo where autonum=" & DataCombo2.BoundText & " and flag_camera=1"
If Nnull(db.Execute(Sql).Fields(0), 0) <> 0 Then
    DataCombo4.BoundText = ""
    With Frame2
        .Left = 105
        .Width = 6855
        .Top = 1290
        .Height = 1005
        .Visible = True
        If Me.mskDestino.text <> "" Then
            Me.DataCombo4.BoundText = Nnull(db.Execute("SELECT ISNULL(AUTONUM_CAMERA,0) FROM OPERADOR..TB_YARD WHERE PATIO=" & Col_Patio & " AND YARD='" & Me.mskDestino.text & "'")(0), 0)
        End If
        
    End With
    Exit Sub
Else
    Call Atualiza_Posicao
End If
    
    

End Sub




Private Sub Command2_Click()
Limpa_Dados
CNTR = "___-_"
Me.MskCntr = "__________-_"
CNTR.SetFocus
End Sub


Private Sub Command3_Click()

If DataCombo4.BoundText = "" Then
    MsgBox "Selecione uma camera"
    Exit Sub
End If

Frame2.Visible = False

Call Atualiza_Posicao




End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub DataCombo3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        DataCombo3.BoundText = ""
    End If


End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Sendkeys "{tab}"

End Sub

Private Sub Form_Load()


ConnectionStringAdo = ConnectionString


'sql = "SELECT autonum,identificacao FROM " & banco_operador & "TB_FROTA WHERE PATIO=4 and flag_ativo=1 and tipo_veiculo='E' ORDER BY IDENTIFICACAO"
'Adodc2.ConnectionString = ConnectionStringAdo
'Adodc2.RecordSource = UCase(sql)
'Adodc2.ConnectionString = ConnectionStringAdo: Adodc2.Refresh
'Set Adodc2.Recordset.ActiveConnection = Nothing

'sql = "SELECT autonum,placa_c || ' / ' || placa_carreta IDENTIFICACAO FROM " & banco_operador & "TB_FROTA WHERE PATIO=4 and flag_ativo=1 and tipo_veiculo='C' ORDER BY placa_c"
'Adodc4.ConnectionString = ConnectionStringAdo
'Adodc4.RecordSource = UCase(sql)
'Adodc4.ConnectionString = ConnectionStringAdo: Adodc4.Refresh
'Set Adodc4.Recordset.ActiveConnection = Nothing

'sql = "SELECT autonum,descricao FROM " & banco_operador & "TB_cad_motivo where flag_ativo=1 and flag_redex=1 ORDER BY descricao"
Sql = "SELECT autonum,descricao FROM " & Banco_Operador & "TB_cad_motivo where flag_ativo=1 ORDER BY descricao"
Adodc3.ConnectionString = ConnectionStringAdo
Adodc3.RecordSource = UCase(Sql)
Adodc3.ConnectionString = ConnectionStringAdo: Adodc3.Refresh
Set Adodc3.Recordset.ActiveConnection = Nothing


Sql = "SELECT autonum,descr FROM " & Banco_Operador & "TB_CAMERAS WHERE PATIO=" & Col_Patio & " ORDER BY descr"
Adodc31.ConnectionString = ConnectionStringAdo
Adodc31.RecordSource = UCase(Sql)
Adodc31.ConnectionString = ConnectionStringAdo: Adodc31.Refresh
Set Adodc31.Recordset.ActiveConnection = Nothing


Me.TxtPatio.text = db.Execute("SELECT DESCR_RESUMIDO FROM OPERADOR..TB_PATIOS WHERE AUTONUM=" & Col_Patio)(0)

End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub MskCntr_LostFocus()

Call Busca_Dados

End Sub

Function Valida_Tara(ISO As String, Peso As Long, Tamanho As Integer) As String
Valida_Tara = ""
Dim RsTara As New ADODB.Recordset
Dim CC As String
Dim Tam As Byte
Sql = "select tamanho,tipo,nvl(pesonominal,0) tara from " & Banco_Operador & "tb_cad_conteiners where ISO='" & ISO & "'"
PRSet RsTara, Sql
If RsTara.EOF Then
    Valida_Tara = "Codigo ISO não cadastrado"
    RsTara.Close
    Exit Function
End If

CC = Nnull(RsTara!Tipo, 1)
Tam = Nnull(RsTara!Tamanho, 0)
RsTara.Close

If Tam <> Tamanho Then
    Valida_Tara = "Tamanho divergente do codigo ISO"
    Exit Function
End If

Sql = "select tara_min_20, tara_max_20,tara_min_40, tara_max_40 from " & Banco_Sgipa & "dte_tb_tipos_conteiner where codigo='" & CC & "'"
PRSet RsTara, Sql
If RsTara.EOF Then
    Valida_Tara = "Tipo de conteiner ( " & CC & " ) não encontrado na tabela dte_tb_tipos_conteiner"
    RsTara.Close
    Exit Function
End If
If Peso < Nnull(RsTara("tara_min_" & Tam), 0) Or Peso > Nnull(RsTara("tara_max_" & Tam), 0) Then
    Valida_Tara = "Tara inválida" & vbCr & "Tara deve estar entre " & Nnull(RsTara("tara_min_" & Tam), 0) & " e " & Nnull(RsTara("tara_max_" & Tam), 0)
    RsTara.Close
    Exit Function
End If
RsTara.Close


End Function



Private Sub mskDestino_Change()

If Len(Me.mskDestino) >= 1 And Option1(1).Value = True Then
    Carrega_Grid2
End If

If Len(Me.mskDestino) = 0 And Option1(1).Value = True Then
    Grid2.Clear
End If

End Sub

Private Sub mskDestino_LostFocus()
mskDestino = UCase(mskDestino)

End Sub

Private Sub Option1_Click(Index As Integer)
Carrega_Grid2
If Index = 0 Then Me.Label1(8) = "FIADA ATUAL"
If Index = 1 Then Me.Label1(8) = "Q/FIADA DESTINO"
End Sub

