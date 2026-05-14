VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FrmMovimentacao 
   BorderStyle     =   0  'None
   Caption         =   "Movimentação"
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
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
   ScaleHeight     =   4950
   ScaleWidth      =   3540
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
      Height          =   4980
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   8784
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   14737632
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Moviment."
      TabPicture(0)   =   "FrmMovimentacaoS.frx":0000
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
      Tab(0).Control(5)=   "Label1(7)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(8)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(9)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblReeferDesligado"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "LBLDTA"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblDestinadoDesova"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(10)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label16"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(11)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtSetPoint"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtMin"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtMax"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "CNTR"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "MskCntr"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Command1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Command4"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "DataCombo2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Command2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "TXTYardAtual"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "mskDestino"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text1(0)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Grid1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Option1(0)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Option1(1)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Adodc5"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "frmFrota"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtPatio"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Grid2"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Frame2"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Frame5"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtEF2"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).ControlCount=   36
      TabCaption(1)   =   "Histórico"
      TabPicture(1)   =   "FrmMovimentacaoS.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(3)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "GridHist_Avaria"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "DataGrid1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Dados CNTR"
      TabPicture(2)   =   "FrmMovimentacaoS.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label13"
      Tab(2).Control(1)=   "Label12"
      Tab(2).Control(2)=   "Label11"
      Tab(2).Control(3)=   "Label10"
      Tab(2).Control(4)=   "Label8"
      Tab(2).Control(5)=   "Label6"
      Tab(2).Control(6)=   "Label5"
      Tab(2).Control(7)=   "Label4"
      Tab(2).Control(8)=   "Label3"
      Tab(2).Control(9)=   "Label1(6)"
      Tab(2).Control(10)=   "Label1(5)"
      Tab(2).Control(11)=   "Label17"
      Tab(2).Control(12)=   "DataCombo3"
      Tab(2).Control(13)=   "DataCombo1"
      Tab(2).Control(14)=   "DataGrid2"
      Tab(2).Control(15)=   "txtLacres"
      Tab(2).Control(16)=   "txtIMO"
      Tab(2).Control(17)=   "txtBruto"
      Tab(2).Control(18)=   "TxtGateIN"
      Tab(2).Control(19)=   "TxtNavio"
      Tab(2).Control(20)=   "txtEF"
      Tab(2).Control(21)=   "txtTipo"
      Tab(2).Control(22)=   "txtTam"
      Tab(2).Control(23)=   "Frame1"
      Tab(2).Control(24)=   "txtSIGVIG"
      Tab(2).ControlCount=   25
      TabCaption(3)   =   "Avarias"
      TabPicture(3)   =   "FrmMovimentacaoS.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdExcluirAvaria"
      Tab(3).Control(1)=   "cmdFinalizarAvaria"
      Tab(3).Control(2)=   "GridAvarias"
      Tab(3).Control(3)=   "cmdIncluirAvaria"
      Tab(3).Control(4)=   "Frame4"
      Tab(3).Control(5)=   "Frame3"
      Tab(3).Control(6)=   "Label18"
      Tab(3).Control(7)=   "Shape1"
      Tab(3).ControlCount=   8
      Begin VB.TextBox txtEF2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   96
         Text            =   "-"
         Top             =   510
         Width           =   300
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00808080&
         Height          =   1395
         Left            =   60
         TabIndex        =   92
         Top             =   3075
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton Command5 
            Caption         =   "OK"
            Height          =   405
            Left            =   2745
            TabIndex        =   94
            Top             =   495
            Width           =   450
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Transporte Interno"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   93
            Top             =   585
            Width           =   2430
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Câmera"
         Height          =   915
         Left            =   300
         TabIndex        =   64
         Top             =   3600
         Visible         =   0   'False
         Width           =   3465
         Begin VB.CommandButton Command3 
            Caption         =   "Salvar"
            Height          =   435
            Left            =   2220
            TabIndex        =   66
            Top             =   300
            Width           =   825
         End
         Begin MSDataListLib.DataCombo DataCombo4 
            Bindings        =   "FrmMovimentacaoS.frx":0070
            Height          =   315
            Left            =   120
            TabIndex        =   65
            Top             =   390
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Appearance      =   0
            Style           =   2
            ListField       =   "descr"
            BoundColumn     =   "autonum"
            Text            =   "DataCombo4"
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
      End
      Begin VB.CommandButton cmdExcluirAvaria 
         Caption         =   "EXCLUIR"
         Height          =   345
         Left            =   -72930
         TabIndex        =   84
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdFinalizarAvaria 
         Caption         =   "FINALIZAR"
         Height          =   345
         Left            =   -74940
         TabIndex        =   83
         Top             =   4530
         Width           =   1485
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridAvarias 
         Height          =   1905
         Left            =   -74910
         TabIndex        =   82
         Top             =   2550
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3360
         _Version        =   393216
         Cols            =   3
         RowHeightMin    =   300
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "    |<Localização                              |<Tipo                           "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin VB.CommandButton cmdIncluirAvaria 
         Caption         =   "INCLUIR"
         Height          =   345
         Left            =   -74910
         TabIndex        =   81
         Top             =   2160
         Width           =   1515
      End
      Begin VB.Frame Frame4 
         Caption         =   "TIPO"
         Height          =   1755
         Left            =   -72990
         TabIndex        =   78
         Top             =   360
         Width           =   1485
         Begin VB.ListBox ListTipo 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1410
            Left            =   60
            Style           =   1  'Checkbox
            TabIndex        =   80
            Top             =   270
            Width           =   1365
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "LOCAL"
         Height          =   1755
         Left            =   -74970
         TabIndex        =   77
         Top             =   360
         Width           =   1965
         Begin VB.ListBox ListLocal 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1410
            Left            =   60
            Style           =   1  'Checkbox
            TabIndex        =   79
            Top             =   270
            Width           =   1845
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
         Bindings        =   "FrmMovimentacaoS.frx":0086
         Height          =   1425
         Left            =   60
         TabIndex        =   45
         Top             =   3060
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   2514
         _Version        =   393216
         Cols            =   4
         RowHeightMin    =   330
         WordWrap        =   -1  'True
         Appearance      =   0
         FormatString    =   "|<Conteiner   |<Posição|<Motivo                     "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   1050
         Width           =   570
      End
      Begin VB.Frame frmFrota 
         Caption         =   "Escolha do veículo "
         Height          =   915
         Left            =   390
         TabIndex        =   69
         Top             =   3000
         Visible         =   0   'False
         Width           =   3405
         Begin MSAdodcLib.Adodc AdodcCam 
            Height          =   330
            Left            =   2400
            Top             =   90
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
         Begin MSDataListLib.DataCombo dcCamRetVz 
            Bindings        =   "FrmMovimentacaoS.frx":009B
            Height          =   315
            Left            =   30
            TabIndex        =   70
            Top             =   360
            Width           =   3315
            _ExtentX        =   5847
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Appearance      =   0
            Style           =   2
            ListField       =   "Display"
            BoundColumn     =   "autonum"
            Text            =   "DataCombo2"
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
      End
      Begin VB.TextBox txtSIGVIG 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   -74160
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   3360
         Width           =   2610
      End
      Begin VB.Frame Frame1 
         Caption         =   "DADOS ENTRADA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   -74880
         TabIndex        =   48
         Top             =   1860
         Width           =   3375
         Begin VB.TextBox TxtLacreLona 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1830
            MaxLength       =   20
            TabIndex        =   97
            Top             =   1140
            Width           =   1455
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "LACRES CONFERIDOS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   1320
            TabIndex        =   63
            Top             =   200
            Width           =   2115
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
            TabIndex        =   60
            Top             =   960
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.TextBox txtLACRE3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   180
            MaxLength       =   20
            TabIndex        =   59
            Top             =   1140
            Width           =   1455
         End
         Begin VB.TextBox txtLACRE2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1830
            MaxLength       =   20
            TabIndex        =   58
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtLACRE1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   180
            MaxLength       =   20
            TabIndex        =   57
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox TxtLacreIPA 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1830
            MaxLength       =   20
            TabIndex        =   56
            Top             =   550
            Width           =   1455
         End
         Begin VB.TextBox txtMaxGross 
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
            Height          =   525
            Left            =   5400
            MaxLength       =   5
            TabIndex        =   54
            Top             =   -90
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.TextBox txtTara 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   570
            MaxLength       =   4
            TabIndex        =   52
            Top             =   240
            Width           =   555
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
            TabIndex        =   50
            Top             =   -150
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "L"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   1695
            TabIndex        =   98
            Top             =   1110
            Width           =   105
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   30
            TabIndex        =   75
            Top             =   1170
            Width           =   135
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   1695
            TabIndex        =   74
            Top             =   810
            Width           =   135
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   30
            TabIndex        =   73
            Top             =   870
            Width           =   135
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   1695
            TabIndex        =   72
            Top             =   510
            Width           =   135
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LACRES"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   165
            TabIndex        =   55
            Top             =   600
            Width           =   645
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
            TabIndex        =   53
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
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   51
            Top             =   270
            Width           =   495
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
            TabIndex        =   49
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
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2610
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   4500
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ATUAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1905
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   4500
         Value           =   -1  'True
         Width           =   675
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
         Bindings        =   "FrmMovimentacaoS.frx":00B2
         Height          =   1275
         Left            =   60
         TabIndex        =   42
         Top             =   3210
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   2249
         _Version        =   393216
         RowHeightMin    =   300
         WordWrap        =   -1  'True
         Appearance      =   0
         FormatString    =   "|<Motivo                   "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
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
         Height          =   285
         Left            =   -74880
         TabIndex        =   27
         Top             =   585
         Width           =   405
      End
      Begin VB.TextBox txtTipo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   -74430
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   585
         Width           =   465
      End
      Begin VB.TextBox txtEF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   -73920
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   585
         Width           =   270
      End
      Begin VB.TextBox TxtNavio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   -73605
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   585
         Width           =   1815
      End
      Begin VB.TextBox TxtGateIN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   -74880
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1080
         Width           =   1080
      End
      Begin VB.TextBox txtBruto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   -74880
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1575
         Width           =   630
      End
      Begin VB.TextBox txtIMO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   -74205
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1575
         Width           =   525
      End
      Begin VB.TextBox txtLacres 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   -73635
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1575
         Width           =   2100
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   0
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   2190
         Width           =   3420
      End
      Begin VB.TextBox mskDestino 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   960
         MaxLength       =   7
         TabIndex        =   2
         Top             =   1050
         Width           =   915
      End
      Begin VB.TextBox TXTYardAtual 
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
         Height          =   300
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1050
         Width           =   870
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Limpar"
         Height          =   315
         Left            =   60
         TabIndex        =   12
         Top             =   4500
         Width           =   885
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "FrmMovimentacaoS.frx":00C7
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Top             =   1380
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         ListField       =   "descricao"
         BoundColumn     =   "autonum"
         Text            =   "DataCombo2"
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
         Height          =   375
         Left            =   3180
         TabIndex        =   9
         Top             =   480
         Width           =   315
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "FrmMovimentacaoS.frx":00DC
         Height          =   2220
         Left            =   -74910
         TabIndex        =   7
         Top             =   660
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   3916
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   34
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
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
               Format          =   "dd/mm/yy hh:mm"
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
               Alignment       =   2
               WrapText        =   -1  'True
               ColumnWidth     =   764,787
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               WrapText        =   -1  'True
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               WrapText        =   -1  'True
               ColumnWidth     =   675,213
            EndProperty
            BeginProperty Column03 
               WrapText        =   -1  'True
               ColumnWidth     =   2264,882
            EndProperty
            BeginProperty Column04 
               WrapText        =   -1  'True
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2924,788
            EndProperty
            BeginProperty Column06 
               WrapText        =   -1  'True
               ColumnWidth     =   2564,788
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Salvar"
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
         Left            =   2520
         TabIndex        =   4
         Top             =   1020
         Width           =   930
      End
      Begin MSMask.MaskEdBox MskCntr 
         Height          =   360
         Left            =   960
         TabIndex        =   1
         Top             =   510
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   635
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
         Height          =   360
         Left            =   60
         TabIndex        =   0
         Top             =   495
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         Height          =   1200
         Left            =   -74880
         TabIndex        =   16
         Top             =   3660
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   2117
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   18
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
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
               ColumnWidth     =   1005,165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1995,024
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "FrmMovimentacaoS.frx":00F1
         Height          =   315
         Left            =   -73770
         TabIndex        =   36
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         ListField       =   "identificacao"
         BoundColumn     =   "autonum"
         Text            =   "DataCombo1"
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
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "FrmMovimentacaoS.frx":0106
         Height          =   315
         Left            =   -72660
         TabIndex        =   37
         Top             =   1080
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         ListField       =   "identificacao"
         BoundColumn     =   "autonum"
         Text            =   "DataCombo3"
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
      Begin VB.TextBox txtMax 
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
         Left            =   3030
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   1710
         Width           =   420
      End
      Begin VB.TextBox txtMin 
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
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   1710
         Width           =   540
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   88
         Top             =   1710
         Width           =   600
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridHist_Avaria 
         Height          =   1665
         Left            =   -74910
         TabIndex        =   91
         Top             =   3210
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   2937
         _Version        =   393216
         Cols            =   4
         RowHeightMin    =   300
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "    |<Data                          |<Avaria                      |<Usuário                               "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " EF"
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
         Index           =   11
         Left            =   2850
         TabIndex        =   95
         Top             =   300
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Avarias"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   -74880
         TabIndex        =   90
         Top             =   2940
         Width           =   750
      End
      Begin VB.Label Label16 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2880
         TabIndex        =   89
         Top             =   1680
         Width           =   225
      End
      Begin VB.Label Label18 
         Caption         =   "Informado pelo Op. Port."
         Height          =   405
         Left            =   -72900
         TabIndex        =   85
         Top             =   4500
         Width           =   1305
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   -73260
         Top             =   4530
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Temperatura:"
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
         Index           =   10
         Left            =   420
         TabIndex        =   76
         Top             =   1740
         Width           =   1170
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SIGVIG"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74850
         TabIndex        =   67
         Top             =   3420
         Width           =   600
      End
      Begin VB.Label lblDestinadoDesova 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEST.DESOVA/MAPA"
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
         Left            =   1050
         TabIndex        =   62
         Top             =   300
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.Label LBLDTA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DTA"
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
         Left            =   1050
         TabIndex        =   61
         Top             =   300
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblReeferDesligado 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RF. DESLIG."
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
         Left            =   2460
         TabIndex        =   47
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CARREGAMENTO REGISTRADO"
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
         Index           =   9
         Left            =   120
         TabIndex        =   46
         Top             =   2010
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FIADA ATUAL"
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
         Index           =   8
         Left            =   2040
         TabIndex        =   41
         Top             =   3030
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONTEINER"
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
         Index           =   7
         Left            =   240
         TabIndex        =   40
         Top             =   3030
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stacker"
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
         Index           =   5
         Left            =   -73740
         TabIndex        =   39
         Top             =   870
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Truck"
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
         Index           =   6
         Left            =   -72660
         TabIndex        =   38
         Top             =   870
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tam"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74910
         TabIndex        =   35
         Top             =   345
         Width           =   405
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74400
         TabIndex        =   34
         Top             =   345
         Width           =   360
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73890
         TabIndex        =   33
         Top             =   360
         Width           =   195
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Navio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73590
         TabIndex        =   32
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GATE IN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74880
         TabIndex        =   31
         Top             =   870
         Width           =   660
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BRUTO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74880
         TabIndex        =   30
         Top             =   1380
         Width           =   570
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74115
         TabIndex        =   29
         Top             =   1380
         Width           =   345
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LACRES MANIFESTADOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -73605
         TabIndex        =   28
         Top             =   1395
         Width           =   1980
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -73050
         TabIndex        =   17
         Top             =   300
         Width           =   1545
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
         Left            =   30
         TabIndex        =   15
         Top             =   300
         Width           =   960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "POS. ATUAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   30
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1710
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
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
         Index           =   4
         Left            =   300
         TabIndex        =   10
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Origem"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   -74895
         TabIndex        =   8
         Top             =   375
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESTINO"
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
         Index           =   1
         Left            =   1080
         TabIndex        =   6
         Top             =   840
         Width           =   720
      End
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   6360
      Top             =   7920
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
      Left            =   9780
      Top             =   9450
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
      Left            =   9600
      Top             =   9330
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

Sub Atualiza_Posicao()

CntrVazio = 0

If idPatio = 0 Then
    Frmmensagem.Label1.Caption = "Um conteiner deve ser informado para pesquisa!"
    Frmmensagem.Show vbModal
    MskCntr.SetFocus
    Exit Sub
End If

Dim SQl As String

    Dim ST As New ADODB.Recordset
    SQl = "SELECT isnull(flag_reefer_sem_temp,0) as flag_reefer_sem_temp, isnull(flag_desligado,0) as flag_desligado FROM SGIPA.DBO.TB_CNTR_BL WHERE AUTONUM=" & idPatio & " and isnull(flag_reefer_sem_temp,0)<>0"
    PRSet ST, SQl
    If Not ST.EOF Then
        If ST!flag_reefer_sem_temp = 1 And ST!flag_desligado = 0 Then
            Frmmensagem.Label1.Caption = "Conteiner Reefer sem temperatura - Posicionamento bloqueado"
            Frmmensagem.Show vbModal
            Exit Sub
        End If
    End If


If idPatio > 0 And MskCntr.Tag = "I" Then
    SQl = "SELECT AUTONUM FROM OPERADOR..TB_HIST_SHIFTING WHERE CNTR=" & idPatio & " AND TIPO='I' AND ORIGEM<>'GATEIN'"
    Set tbhist = New ADODB.Recordset
    tbhist.Open SQl, db, 1, 1
    If tbhist.EOF Then
        If Me.Check1.Value = 0 Then
            Frmmensagem.Label1.Caption = "Informe que os lacres foram conferidos"
            Frmmensagem.Show vbModal
            Exit Sub
        End If
    End If
End If



If Me.MskCntr.Tag = "I" Then
    If Trim(Me.TxtLacreIPA.Text) = "" And Trim(Me.txtLACRE1.Text) = "" And Trim(Me.txtLACRE2.Text) = "" And Trim(Me.txtLACRE3.Text) = "" Then
        MsgBox "Pelo menos 1 lacre de origem deve ser informado", vbInformation
        Exit Sub
    End If
End If

If Me.MskCntr.Tag = "I" And Me.txtEF = "F" Then
    If Trim(Me.TxtLacreLona.Text) = "" And Me.txtTipo.Text = "OT" Then
        MsgBox "Informe o Lacre Lona ", vbInformation
        Exit Sub
    End If
End If

 
If DataCombo2.BoundText = "" Then
    Frmmensagem.Label1.Caption = "Informe o motivo da movimentação!"
    Frmmensagem.Show vbModal
 
    DataCombo2.SetFocus
    Exit Sub
End If
        
If mskDestino = "" Then
    Frmmensagem.Label1.Caption = "Destino não Informado!"
    Frmmensagem.Show vbModal
 
    mskDestino.SetFocus
    Exit Sub
End If

If mskDestino = Me.TXTYardAtual Then
    Frmmensagem.Label1.Caption = "Destino e Origem iguais!"
    Frmmensagem.Show vbModal
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

'Verifica Bloqueio Escaner Cntr
If Saida = True And idPatio > 0 And MskCntr.Tag = "I" Then
    Dim Flag_Bloqueio_Scanner As Boolean
    SQl = "Select SGIPA.DBO.FN_BLOQUEIO_SCANNER(" & idPatio & ",'SAIDA') "
    Flag_Bloqueio_Scanner = db.Execute(SQl)(0)
    If Flag_Bloqueio_Scanner = True Then
        Frmmensagem.Label1.Caption = "Esta unidade tem BLOQUEIO de ESCANER"
        Frmmensagem.Show vbModal
        mskDestino.SetFocus
        Exit Sub
    End If

End If


Dim Auto As Long

'Se o Destino é BK precisa estar autorizado
If Me.mskDestino = "BK" And idPatio > 0 And MskCntr.Tag = "I" Then
    ModoRepesagem = 1
    SQl = " SELECT ISNULL(FLAG_SEL_TRANSF_BK,0) AS QUAL FROM SGIPA..TB_CNTR_BL WHERE AUTONUM= " & idPatio
    If db.Execute(SQl)(0) = 0 Then
        Frmmensagem.Label1.Caption = "Unidade sem autorização para Transferência BK"
        Frmmensagem.Show vbModal
        mskDestino.SetFocus
        Exit Sub
    End If
    SQl = "SELECT  ISNULL(MAX(C.AUTONUM_GD_RESERVA),0) FROM SGIPA..TB_CNTR_BL A  "
    SQl = SQl + "INNER JOIN  OPERADOR..TB_GD_RESERVA B ON A.AUTONUM_GD_RESERVA=B.AUTONUM_GD_RESERVA   "
    SQl = SQl + "INNER JOIN  OPERADOR..TB_GD_RESERVA C ON B.PERIODO_INICIAL=C.PERIODO_INICIAL AND B.SERVICO_GATE=C.SERVICO_GATE AND C.PATIO=2 "
    SQl = SQl + "LEFT JOIN ( SELECT COUNT(1) AG ,AUTONUM_GD_RESERVA FROM SGIPA..TB_CNTR_BL WHERE  "
    SQl = SQl + "FLAG_TERMINAL=1 AND FLAG_HISTORICO=0 GROUP BY  AUTONUM_GD_RESERVA) D ON  C.AUTONUM_GD_RESERVA=D.AUTONUM_GD_RESERVA  OR B.PATIO=2 "
    SQl = SQl + "WHERE  ISNULL(C.LIMITE_MOVIMENTOS,0)>ISNULL(D.AG,0) AND "
    SQl = SQl + " AUTONUM= " & idPatio
    Auto = db.Execute(SQl)(0)
    If Auto > 0 Then
       SQl = "UPDATE SGIPA..TB_CNTR_BL SET AUTONUM_GD_RESERVA=" & Auto & " WHERE AUTONUM= " & idPatio
       Executa SQl
    Else
        SQl = " SELECT COUNT(1) QUAL  FROM SGIPA..VW_AGG_AGENDAMENTOS  WHERE AUTONUMCNTR= " & idPatio
        If db.Execute(SQl)(0) <> 0 Then
            Frmmensagem.Label1.Caption = "Unidade com Agendamento, Transferência BK não autorizada "
            SQl = "SELECT   ISNULL(MAX('Limite Janela BK ' + CONVERT(VARCHAR,ISNULL(C.LIMITE_MOVIMENTOS,0))+' Agendados ' + CONVERT(VARCHAR,ISNULL(D.AG,0)) ),'')  FROM SGIPA..TB_CNTR_BL A  "
            SQl = SQl + "INNER JOIN  OPERADOR..TB_GD_RESERVA B ON A.AUTONUM_GD_RESERVA=B.AUTONUM_GD_RESERVA   "
            SQl = SQl + "INNER JOIN  OPERADOR..TB_GD_RESERVA C ON B.PERIODO_INICIAL=C.PERIODO_INICIAL AND B.SERVICO_GATE=C.SERVICO_GATE AND C.PATIO=2 "
            SQl = SQl + "LEFT JOIN ( SELECT COUNT(1) AG ,AUTONUM_GD_RESERVA FROM SGIPA..TB_CNTR_BL WHERE  "
            SQl = SQl + "FLAG_TERMINAL=1 AND FLAG_HISTORICO=0 GROUP BY  AUTONUM_GD_RESERVA) D ON  C.AUTONUM_GD_RESERVA=D.AUTONUM_GD_RESERVA  "
            SQl = SQl + "WHERE  "
            SQl = SQl + " AUTONUM= " & idPatio
            TEXTO = db.Execute(SQl)(0)
            Frmmensagem.Label1.Caption = Frmmensagem.Label1.Caption + " " + Nnull(TEXTO, 1)
            Frmmensagem.Show vbModal
            mskDestino.SetFocus
            Exit Sub
        End If
    End If
    P_Cntr = idPatio
    P_BL = 0
   ' FrmVeiculoServ.Show vbModal
    P_Cntr = 0
End If

If Me.mskDestino = "CLIA" And idPatio > 0 And MskCntr.Tag = "I" Then
    ModoRepesagem = 1
    SQl = " SELECT ISNULL(FLAG_SEL_TRANSF_CLIA,0) AS QUAL FROM SGIPA..TB_CNTR_BL WHERE AUTONUM= " & idPatio
    If db.Execute(SQl)(0) = 0 Then
        Frmmensagem.Label1.Caption = "Unidade sem autorização para Transferência CLIA"
        Frmmensagem.Show vbModal
        mskDestino.SetFocus
        Exit Sub
    End If
    P_Cntr = idPatio
    P_BL = 0
    'FrmVeiculoServ.Show vbModal
    P_Cntr = 0
End If


If Me.mskDestino.Text <> "DTA" Then
If Me.TxtLacreIPA.Text = "" Then
    If Me.txtLACRE1.Text <> "" Or Me.txtLACRE2.Text <> "" Or Me.txtLACRE3.Text <> "" Then
        Frmmensagem.Label1.Caption = "Lacres informados incorretamente"
        Frmmensagem.Show vbModal
        Exit Sub
    End If
End If

If Me.txtLACRE1.Text = "" Then
    If Me.txtLACRE2.Text <> "" Or Me.txtLACRE3.Text <> "" Then
        Frmmensagem.Label1.Caption = "Lacres informados incorretamente"
        Frmmensagem.Show vbModal
        Exit Sub
    End If
End If

If Me.txtLACRE2.Text = "" Then
    If Me.txtLACRE3.Text <> "" Then
        Frmmensagem.Label1.Caption = "Lacres informados incorretamente"
        Frmmensagem.Show vbModal
        Exit Sub
    End If
End If
End If
    
If Me.txtEF = "E" And idPatio > 0 And mskDestino = "SAIDA" Then
    P_Cntr = 0
    LoteDTA = 0
    P_BL = 0
    CntrVazio = idPatio
    FrmVeiculoServ.Show vbModal
End If

If Me.mskDestino.Text = "DTA" And idPatio > 0 Then
    LoteDTA = 0
    P_Cntr = idPatio
    FrmVeiculoServ.Show vbModal
    P_BL = 0
    LoteDTA = 0
    Me.mskDestino.Text = "CAM"
End If


'Dim FlagAbcDestino As Integer
'
'FlagAbcDestino = 0
'
'
'SQl = "select isnull(max(flag_abc),0) from operador..tb_yard where patio=" & Col_Patio & " and (yard='" & Me.mskDestino.text & "' or yard_eq='" & Me.mskDestino.text & "')"
'FlagAbcDestino = db.Execute(SQl)(0)


Set tbvw = New ADODB.Recordset

SQl = "select [Operador].[dbo].[fnMovimentaCntr]('" & MskCntr & "'," & Col_Patio & "," & Val(Me.txtTam.Text) & ",'" & Me.MskCntr.Tag & "','" & Me.TXTYardAtual.Text & "','" & Me.mskDestino.Text & "'," & Cod_Usuario & "," & Me.Check1.Tag & ")"

tbvw.Open SQl, db, 1, 1
If Not tbvw.EOF Then
    
    If tbvw.Fields(0) <> "OK" Then
        Frmmensagem.Label1.Caption = tbvw.Fields(0)
        Frmmensagem.Show vbModal
        Exit Sub
    Else
        
        Dim cmd As New ADODB.Command
        
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = ConnectionString
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "[Operador].[dbo].[pcMovimentaCntr]"
        cmd.Parameters.Append cmd.CreateParameter("Autonum", adInteger, adParamInput, 8, idPatio)
        cmd.Parameters.Append cmd.CreateParameter("Sigla", adVarChar, adParamInput, 12, Me.MskCntr.Text)
        cmd.Parameters.Append cmd.CreateParameter("Gmci", adVarChar, adParamInput, 15, Me.TXTgmci.Text)
        cmd.Parameters.Append cmd.CreateParameter("Sistema", adVarChar, adParamInput, 1, Me.MskCntr.Tag)
        cmd.Parameters.Append cmd.CreateParameter("YardOrigem", adVarChar, adParamInput, 10, Me.TXTYardAtual.Text)
        cmd.Parameters.Append cmd.CreateParameter("YardDestino", adVarChar, adParamInput, 10, Me.mskDestino.Text)
        cmd.Parameters.Append cmd.CreateParameter("Tara", adInteger, adParamInput, 8, Val(Me.txtTara.Text))
        cmd.Parameters.Append cmd.CreateParameter("Lacre_Ipa", adVarChar, adParamInput, 20, Me.TxtLacreIPA.Text)
        cmd.Parameters.Append cmd.CreateParameter("Lacre1_Gate", adVarChar, adParamInput, 20, Me.txtLACRE1.Text)
        cmd.Parameters.Append cmd.CreateParameter("Lacre2_Gate", adVarChar, adParamInput, 20, Me.txtLACRE2.Text)
        cmd.Parameters.Append cmd.CreateParameter("Lacre3_Gate", adVarChar, adParamInput, 20, Me.txtLACRE3.Text)
        cmd.Parameters.Append cmd.CreateParameter("Lacre_Lona", adVarChar, adParamInput, 20, Me.TxtLacreLona.Text)
        cmd.Parameters.Append cmd.CreateParameter("Camera_Posicionamento", adInteger, adParamInput, 5, Nnull(Me.DataCombo4.BoundText, 0))
        cmd.Parameters.Append cmd.CreateParameter("Usuario", adInteger, adParamInput, 8, Cod_Usuario)
        cmd.Parameters.Append cmd.CreateParameter("Motivo", adInteger, adParamInput, 8, Nnull(Me.DataCombo2.BoundText, 0))
   
        
        cmd.Execute
     
        Set cmd = Nothing
        
    End If
    
    
    Dim FlagDivLacre As Boolean
    FlagDivLacre = False
    
    If MskCntr.Tag = "I" Then
    
        FlagDivLacre = True
     
        Dim Lacres_Manif As String
        Lacres_Manif = ""
        If Trim(Me.TxtLacreIPA.Tag) <> "" Then Lacres_Manif = Lacres_Manif & "***" & Trim(Me.TxtLacreIPA.Tag)
        If Trim(Me.txtLACRE1.Tag) <> "" Then Lacres_Manif = Lacres_Manif & "***" & Trim(Me.txtLACRE1.Tag)
        If Trim(Me.txtLACRE2.Tag) <> "" Then Lacres_Manif = Lacres_Manif & "***" & Trim(Me.txtLACRE2.Tag)
        If Trim(Me.txtLACRE3.Tag) <> "" Then Lacres_Manif = Lacres_Manif & "***" & Trim(Me.txtLACRE3.Tag)
        Lacres_Manif = UCase(Lacres_Manif)
        
        If (FlagDivLacre) And UCase(Trim(Me.TxtLacreIPA.Text)) <> "" And InStr(1, Lacres_Manif, UCase(Trim(Me.TxtLacreIPA.Text))) <> 0 Then FlagDivLacre = False
        If (FlagDivLacre) And UCase(Trim(Me.txtLACRE1.Text)) <> "" And InStr(1, Lacres_Manif, UCase(Trim(Me.txtLACRE1.Text))) <> 0 Then FlagDivLacre = False
        If (FlagDivLacre) And UCase(Trim(Me.txtLACRE2.Text)) <> "" And InStr(1, Lacres_Manif, UCase(Trim(Me.txtLACRE2.Text))) <> 0 Then FlagDivLacre = False
        If (FlagDivLacre) And UCase(Trim(Me.txtLACRE3.Text)) <> "" And InStr(1, Lacres_Manif, UCase(Trim(Me.txtLACRE3.Text))) <> 0 Then FlagDivLacre = False
    End If
        
    
    If FlagDivLacre = True Then
        Dim Xlocal As String, Xtipo As String, XComplemento As String
        Xlocal = "12": Xtipo = "32": XComplemento = "8"
            
        SQl = "select autonum from " & Banco_Sgipa & "tb_avarias_conteiner where"
        SQl = SQl & " local = '" & Nnull(Xlocal, 1) & "'"
        SQl = SQl & " and tipo = '" & Nnull(Xtipo, 1) & "'"
        SQl = SQl & " and complemento = '" & Nnull(XComplemento, 1) & "'"
        SQl = SQl & " and cntr = " & idPatio
        
        PRSet Rs, SQl
        If Rs.EOF Then
            SQl = "INSERT INTO " & Banco_Sgipa & "TB_AVARIAS_CONTEINER ("
            SQl = SQl & "CNTR,LOCAL,TIPO,COMPLEMENTO,COMPL_IPA,PROVIDENCIA) "
            SQl = SQl & "VALUES ("
            SQl = SQl & idPatio & ",'" & Nnull(Xlocal, 1) & "','" & Nnull(Xtipo, 1) & "','" & Nnull(XComplemento, 1) & "','" & Left(wDiverg, 35) & "',11)"
            Executa SQl
        End If
        Rs.Close
    End If
    
    If Me.txtIMO.Text <> "" Then
        Dim Segreg$
        Segreg = Verifica_Regras_Seg_Imo_Delta(MskCntr, mskDestino)
        If Segreg <> "" Then
            'retorna
             Dim cmd2 As New ADODB.Command
            cmd2.ActiveConnection = ConnectionString
            cmd2.CommandType = adCmdStoredProc
            cmd2.CommandText = "[Operador].[dbo].[pcMovimentaCntr]"
            cmd2.Parameters.Append cmd2.CreateParameter("Autonum", adInteger, adParamInput, 8, idPatio)
            cmd2.Parameters.Append cmd2.CreateParameter("Sigla", adVarChar, adParamInput, 12, Me.MskCntr.Text)
            cmd2.Parameters.Append cmd2.CreateParameter("Gmci", adVarChar, adParamInput, 15, Me.TXTgmci.Text)
            cmd2.Parameters.Append cmd2.CreateParameter("Sistema", adVarChar, adParamInput, 1, Me.MskCntr.Tag)
            cmd2.Parameters.Append cmd2.CreateParameter("YardOrigem", adVarChar, adParamInput, 10, Me.mskDestino.Text)
            cmd2.Parameters.Append cmd2.CreateParameter("YardDestino", adVarChar, adParamInput, 10, Me.TXTYardAtual.Text)
            cmd2.Parameters.Append cmd2.CreateParameter("Tara", adInteger, adParamInput, 8, Val(Me.txtTara.Text))
            cmd2.Parameters.Append cmd2.CreateParameter("Lacre_Ipa", adVarChar, adParamInput, 20, Me.TxtLacreIPA.Text)
            cmd2.Parameters.Append cmd2.CreateParameter("Lacre1_Gate", adVarChar, adParamInput, 20, Me.txtLACRE1.Text)
            cmd2.Parameters.Append cmd2.CreateParameter("Lacre2_Gate", adVarChar, adParamInput, 20, Me.txtLACRE2.Text)
            cmd2.Parameters.Append cmd2.CreateParameter("Lacre3_Gate", adVarChar, adParamInput, 20, Me.txtLACRE3.Text)
            cmd2.Parameters.Append cmd2.CreateParameter("Lacre_Lona", adVarChar, adParamInput, 20, Me.TxtLacreLona.Text)
            cmd2.Parameters.Append cmd2.CreateParameter("Camera_Posicionamento", adInteger, adParamInput, 5, Nnull(Me.DataCombo4.BoundText, 0))
            cmd2.Parameters.Append cmd2.CreateParameter("Usuario", adInteger, adParamInput, 8, Cod_Usuario)
            cmd2.Parameters.Append cmd2.CreateParameter("Motivo", adInteger, adParamInput, 8, Nnull(Me.DataCombo2.BoundText, 0))
            cmd2.Execute
         
            Set cmd2 = Nothing
            Frmmensagem.Label1.Caption = "SEGREGACAO IMO !!!! " & Segreg
            Frmmensagem.Show vbModal
            Exit Sub
        End If
    End If
    
    Limpa_Dados
    Me.CNTR = "___-_"
    Me.MskCntr = "__________-_"
    CNTR.SetFocus

End If



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

    'Sql = "SELECT COUNT(1) FROM " & Banco_Operador & "VW_INVENT_SISTEMAS WHERE PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio) & " AND SUBSTR(ID_CONTEINER,8,5)='" & CNTR & "'"
   ' Sql = CONVERTE_SQL(Sql)
    
    SQl = "SELECT * FROM operador.dbo.fnInvent_Sistemas_YardF('" & CNTR & "'," & Col_Patio & ")"
    Dim tbRegs As New ADODB.Recordset
    tbRegs.Open SQl, db, 1, 1
    'Regs = Nnull(db.Execute(Sql).Fields(0), 0)
    If tbRegs.RecordCount = 0 Then
        'MsgBox "Nenhum contêiner com este final", vbExclamation
        Frmmensagem.Label1.Caption = "Nenhum contêiner com este final"
        Frmmensagem.Show vbModal

        CNTR = "___-_"
        CNTR.SetFocus
        Exit Sub
    End If
    If tbRegs.RecordCount = 1 Then
        MskCntr = Nnull(tbRegs!Id_Conteiner, 1)
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
Set tbvw = New ADODB.Recordset
SQl = "select * FROM operador.dbo.fnInvent_Sistemas_Col('" & MskCntr & "'," & Col_Patio & ")"
tbvw.Open SQl, db, 1, 1

If Not tbvw.EOF Then
    idPatio = tbvw!autonum
    Me.txtEF = tbvw!EF
    Me.txtEF2 = tbvw!EF
    Me.TxtGateIN = Format(tbvw!data_ent_temp, "dd/mm/yyyy")
    Me.TxtNavio = Nnull(tbvw!Nome, 1) & " " & Nnull(tbvw!Viagem, 1)
    Me.txtPOD = Nnull(tbvw!pod, 1)
    Me.txtBruto = Nnull(tbvw!Bruto, 1)
    Me.txtLacres = Nnull(tbvw!Lacre1, 1) & " " & Nnull(tbvw!Lacre2, 1) & " " & Nnull(tbvw!Lacre3, 1) & " " & Nnull(tbvw!Lacre4, 1) & " " & Nnull(tbvw!lacre5, 1) & " " & Nnull(tbvw!lacre6, 1) & " " & Nnull(tbvw!lacre7, 1)
    Me.txtIMO = Nnull(tbvw!IMO1, 1)
    Me.txtTam = Nnull(tbvw!Tamanho, 0)
    Me.txtTipo = Nnull(tbvw!TipoBasico, 1)
    Me.TXTYardAtual = Nnull(tbvw!Yard, 1)
    Me.MskCntr.Tag = Nnull(tbvw!Sistema, 1)
    Me.lblReeferDesligado.Caption = ""
    Me.Check1.Tag = Nnull(tbvw!FLAG_CNTR_BLOQ_CHAO_PATIO, 0)

    
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
     
    Me.lblDestinadoDesova.Visible = False
    Me.lblDestinadoDesova.Caption = "DESTINADO DESOVA"
     
    If tbvw!flag_DESTINADO_DESOVA = 1 Or tbvw!Regime = "HP" Or tbvw!Regime = "PP" Then
        Me.lblDestinadoDesova.Visible = True
    End If
    
    If InStr(tbvw!DESCRICAO_SIGVIG, "SELECIONADO") > 0 Then
       If InStr(tbvw!DESCRICAO_SIGVIG, "NÃO") = 0 Then
          If tbvw!flag_DESTINADO_DESOVA = 1 Or tbvw!Regime = "HP" Or tbvw!Regime = "PP" Then
             Me.lblDestinadoDesova.Visible = True
             Me.lblDestinadoDesova.Caption = "DEST.DESOVA/MAPA"
          Else
             Me.lblDestinadoDesova.Visible = True
             Me.lblDestinadoDesova.Caption = "SELECIONADO MAPA"
          End If
       End If
    End If
    
    Dim DescSigVig As String
    DescSigVig = ""
    
        Me.txtTara.Text = Nnull(tbvw!Tara, 0)
        Me.txtISO.Text = Nnull(tbvw!ISO, 1)
        Me.txtMaxGross.Text = Nnull(tbvw!maxgross, 0)
        
        Me.TxtLacreIPA.Text = Nnull(tbvw!Lacre_ipa, 1)
        Me.TxtLacreIPA.Tag = Nnull(tbvw!lacre_origem, 1)
        
        Me.txtLACRE1.Text = Nnull(tbvw!lacre1_gate, 1)
        Me.txtLACRE1.Tag = Nnull(tbvw!Lacre2, 1)
        
        Me.txtLACRE2.Text = Nnull(tbvw!lacre2_gate, 1)
        Me.txtLACRE2.Tag = Nnull(tbvw!Lacre3, 1)
        
        Me.txtLACRE3.Text = Nnull(tbvw!lacre3_gate, 1)
        Me.txtLACRE3.Tag = Nnull(tbvw!Lacre4, 1)
        
        Me.TxtLacreLona.Text = Nnull(tbvw!lacre_lona, 1)
        Me.TxtLacreLona.Tag = Nnull(tbvw!lacre_lona, 1)
        
        Me.TXTgmci.Text = Nnull(tbvw!GMCI, 1)
        DescSigVig = Nnull(tbvw!DESCRICAO_SIGVIG, 1)
        
        Me.txtSetPoint.Text = Nnull(tbvw!TEMP, 1)
        Me.txtMax.Text = Nnull(tbvw!tempmax, 1)
        Me.txtMin.Text = Nnull(tbvw!tempmin, 1)
        
    
    Text1(0) = ""
   
    
    If tbvw!Sistema = "A" Then
        Text1(0) = "CONTEINER TEMPORARIO" & vbCrLf
        Text1(0) = Text1(0) & Nnull(db.Execute("SELECT B.DESCR  FROM SGIPA..TB_ARMAZENS_IPA A INNER JOIN TB_CAD_TIPO_CT B ON A.TIPO_CT=B.AUTONUM WHERE A.AUTONUM=" & tbvw!autonum)(0), 1) & vbCrLf
        Text1(0) = Text1(0) & "OCUPACAO: " & Nnull(db.Execute("SELECT REPLACE(ISNULL(A.PERC_OCUPACAO_CT,0),'%','')  FROM SGIPA..TB_ARMAZENS_IPA A INNER JOIN TB_CAD_TIPO_CT B ON A.TIPO_CT=B.AUTONUM WHERE A.AUTONUM=" & tbvw!autonum)(0), 1) & "%" & vbCrLf
        
    End If
    
    If tbvw!Sistema = "R" Then
        Text1(0) = "CONTEINER EXPORTACAO" & vbCrLf
    End If
    
    
    If tbvw!Sistema = "I" Then
        Text1(0) = "Imp:" & Nnull(tbvw!Fantasia, 1) & vbCrLf
        bloqueio = False
        If Me.txtEF <> "E" Then
            If Nnull(tbvw!FLAG_BLOQUEIO_MANUAL, 0) = 1 Then bloqueio = True: Text1(0) = Text1(0) & "BLOQUEIO MANUAL DE CNTR" & vbCrLf
            If Nnull(tbvw!FLAG_BLOQUEIO, 0) = 1 Then bloqueio = True: Text1(0) = Text1(0) & "BLOQUEIO ALFANDEGA (CNTR)" & vbCrLf
            If Nnull(tbvw!FLAG_ALERTA_COL_INDICADOR, 1) = 1 Then
                Label1(9).Caption = "CARGA ESPECIAL : " & Nnull(tbvw!Indicador, 1)
                Label1(9).Visible = True
            End If
            'verifica bloqueio bl
            SQl = "select count(*) from " & Banco_Sgipa & "tb_bl bl"
            SQl = SQl & " inner join " & Banco_Sgipa & "tb_amr_cntr_bl amr on bl.autonum=amr.bl"
            SQl = SQl & " where amr.cntr=" & tbvw!autonum & " and bl.flag_bloqueio_manual=1"
            If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then
                bloqueio = True: Text1(0) = Text1(0) & "BLOQUEIO MANUAL DE BL" & vbCrLf
            End If
            SQl = "select count(*) from " & Banco_Sgipa & "tb_bl bl"
            SQl = SQl & " inner join " & Banco_Sgipa & "tb_amr_cntr_bl amr on bl.autonum=amr.bl"
            SQl = SQl & " where amr.cntr=" & tbvw!autonum & " and bl.flag_bloqueio=1"
            If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then
                bloqueio = True: Text1(0) = Text1(0) & "BLOQUEIO ALFANDEGA DE BL" & vbCrLf
            End If
            If Not bloqueio Then Text1(0) = Text1(0) & "NAO CONSTA BLOQUEIO" & vbCrLf
            'VER REGISTRO
            SQl = "select count(*) from " & Banco_Sgipa & "tb_registro_saida_cntr sc"
            SQl = SQl & " inner join " & Banco_Sgipa & "tb_ordem_carregamento oc on sc.ordem_carreg = oc.autonum"
            SQl = SQl & " where sc.cntr=" & tbvw!autonum & " and isnull(oc.flag_saida,0)=0"
            If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then
                Text1(0) = Text1(0) & "CARREGAMENTO REGISTRADO" & vbCrLf
                Label1(9).Caption = "CARREGAMENTO REGISTRADO"
                Label1(9).Visible = True
            Else
                Text1(0) = Text1(0) & "NAO CONSTA REGISTRO" & vbCrLf
            End If
            If Trim(Nnull(tbvw!Num_Documento, 1)) <> "" Then Text1(0) = Text1(0) & "DOCUMENTO : " & tbvw!Num_Documento & vbCrLf
            If Trim(Nnull(tbvw!Num_Documento, 1)) <> "" Then
                Text1(0) = Text1(0) & "Carga Desembaraçada em " & tbvw!Dt_Desembaraco & vbCrLf
            End If
            If DescSigVig = "" Then
                SQl = "select vw.descricao from sgipa.dbo.VW_SIGVIG_CNTR_ESTOQUE vw "
                SQl = SQl & " where vw.CNTR = " & idPatio
                DescSigVig = Nnull(db.Execute(SQl).Fields(0), 1)
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
            Me.txtSIGVIG.Text = DescSigVig
        End If
        'verifica motivo posicionamento
        SQl = "Select x.descr from ("
        SQl = SQl & "select mp.descr || ' em ' || to_char(ap.dt_prevista,'dd/mm/yyyy hh24:mi') as descr, ap.dt_prevista from tb_agendamento_posicao ap"
        SQl = SQl & " inner join tb_agenda_posicao_motivo apm on ap.autonum = apm.autonum_agenda_posicao"
        SQl = SQl & " inner join tb_motivo_posicao mp on apm.motivo_posicao = mp.code"
        SQl = SQl & " Where ap.CNTR = " & tbvw!autonum & " And ap.cancelado = 0"
        SQl = SQl & " union all"
        SQl = SQl & " select 'Agendamento Desova em ' || to_char(dt_prevista,'dd/mm/yyyy hh24:mi') as descr,dt_prevista from " & Banco_Sgipa & "tb_agendamento_desova where cntr=" & tbvw!autonum
        SQl = SQl & ") x order by dt_prevista desc"
        
        PRSet Rs, SQl
        If Not Rs.EOF Then
            Text1(0) = Text1(0) & "Motivos de Posicionamento" & vbCrLf
            Do While Not Rs.EOF
                Text1(0) = Text1(0) & UCase(Rs!Descr) & vbCrLf
                Rs.MoveNext
            Loop
        Else
            Text1(0) = Text1(0) & "Não consta Agendamento" & vbCrLf
        End If
    End If

       
    
Else
    Frmmensagem.Label1.Caption = "Unidade não encontrada no estoque "
    Frmmensagem.Show vbModal
    'MsgBox "Unidade não encontrada no estoque"
    MskCntr = "__________-_"
    MskCntr.SetFocus
    Exit Sub
End If




pula:
Set RsAmr = New ADODB.Recordset
SQl = " select tb_bl.autonum as lote, tb_bl.numero as Bl"
SQl = SQl & " from tb_bl, tb_amr_cntr_bl ,tb_cntr_bl"
SQl = SQl & " where tb_bl.autonum= tb_amr_cntr_bl.bl"
SQl = SQl & " and TB_AMR_CNTR_BL.CNTR=TB_CNTR_BL.AUTONUM "
SQl = SQl & " and tb_cntr_bl.id_conteiner='" & MskCntr & "'"
SQl = SQl & " and tb_bl.flag_ativo=1"

RsAmr.Open SQl, db
Set DataGrid2.DataSource = RsAmr

Label13.Caption = MskCntr

    

Carrega_Grid1
Carrega_Grid2


SetaControlesAV



    Dim ST As New ADODB.Recordset
    SQl = "SELECT isnull(a.flag_reefer_sem_temp,0) as flag_reefer_sem_temp, isnull(a.flag_desligado,0) as flag_desligado"
    SQl = SQl & " FROM SGIPA.DBO.TB_CNTR_BL a"
    SQl = SQl & " WHERE a.AUTONUM=" & idPatio & " and isnull(a.flag_reefer_sem_temp,0)<>0"
    PRSet ST, SQl
    If Not ST.EOF Then
        If ST!flag_reefer_sem_temp = 1 And ST!flag_desligado = 0 Then
            Frmmensagem.Label1.Caption = "Conteiner Reefer sem temperatura - Posicionamento bloqueado"
            Frmmensagem.Show vbModal
        Else
            Frmmensagem.Label1.Caption = "Conteiner Reefer DESLIGADO"
            Frmmensagem.Show vbModal
        End If
    End If


    SQl = "SELECT count(*) FROM SGIPA.DBO.TB_CNTR_BL a"
    SQl = SQl & " inner join SGIPA.DBO.TB_amr_CNTR_BL B on a.autonum=b.cntr"
    SQl = SQl & " inner join SGIPA.DBO.TB_BL c on b.bl=c.autonum"
    SQl = SQl & " inner join SGIPA.DBO.TB_cad_parceiros d on c.importador=d.autonum"
    SQl = SQl & " WHERE a.AUTONUM=" & idPatio & " and d.flag_transp_interno=1"
    If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then
        Frame5.Visible = True
    End If



mskDestino.SetFocus

End Sub
Private Sub Carrega_Grid1()
Grid1.Clear
Dim SQl As String
SQl = "SELECT MOTIVO,LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,108),5)) AS DT_PREV FROM OPERADOR..VW_PROX_MOV_CNTR WHERE ID_CONTEINER='" & Me.MskCntr & "'"
SQl = SQl & " UNION "
SQl = SQl & " SELECT MOTIVO,' ' AS DT_PREV FROM OPERADOR..VW_PROX_MOV_CNTR_COL WHERE ID_CONTEINER='" & Me.MskCntr & "'"

Adodc5.ConnectionString = ConnectionStringAdo
Adodc5.RecordSource = SQl
Adodc5.Refresh
Set Adodc5.Recordset.ActiveConnection = Nothing

Grid1.FormatString = "|<Motivo                    |<Data            "
End Sub
Private Sub Carrega_Grid2()

Grid2.Clear

Dim PosValida As Boolean
Dim PosBusca As String

SQl = ""

If Me.Option1(0).Value = True Then
    PosBusca = Me.TXTYardAtual.Text
End If
If Me.Option1(1).Value = True Then
    PosBusca = Me.mskDestino.Text
End If

If PosBusca <> "" Then

    Set TBval = New ADODB.Recordset
    
    If Col_Patio <> 3 Then
        TBval.Open "Select yard,valida from operador..tb_yard where patio in (1,7) and yard='" & PosBusca & "'", db, 1, 1
    Else
        TBval.Open "Select yard,valida from operador..tb_yard where patio=3 and (yard='" & PosBusca & "' or yard_eq='" & PosBusca & "')", db, 1, 1
    End If
    If TBval.RecordCount <> 0 Then
    
        PosBusca = Nnull(TBval!Yard, 1)
    
        If TBval!Valida = 1 Then
            If Col_Patio <> 3 Then
                SQl = "SELECT ID_CONTEINER,YARD,LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,108),5)),MOTIVO,SUBSTRING(YARD,2,2) AS PILHA FROM sgipa..VW_PROX_MOV_CNTR_all WHERE PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio) & " AND YARD LIKE "
            Else
                SQl = "SELECT ID_CONTEINER,YARD,LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,108),5)),MOTIVO,SUBSTRING(YARD,3,2) AS PILHA FROM sgipa..VW_PROX_MOV_CNTR_all WHERE PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio) & " AND YARD LIKE "
            End If
            
            If Me.Option1(0).Value = True Then
                If Col_Patio <> 3 Then
                    SQl = SQl & "'" & Mid$(PosBusca, 1, 3) & "%'"
                Else
                    SQl = SQl & "'" & Mid$(PosBusca, 1, 4) & "%'"
                End If
            Else
                If Len(PosBusca) = 1 Then
                    SQl = SQl & "'" & Mid$(PosBusca, 1, 1) & "%'"
                ElseIf Len(PosBusca) = 2 Then
                    SQl = SQl & "'" & Mid$(PosBusca, 1, 2) & "%'"
                ElseIf Len(PosBusca) >= 3 Then
                    If Col_Patio <> 3 Then
                        SQl = SQl & "'" & Mid$(PosBusca, 1, 3) & "%'"
                    Else
                        SQl = SQl & "'" & Mid$(PosBusca, 1, 4) & "%'"
                    End If
                End If
            End If
            
            
            SQl = SQl & " and id_conteiner<>'" & Me.MskCntr & "'"
            'AND YARD NOT IN (SELECT YARD FROM OPERADOR..TB_YARD WHERE VALIDA=0) "
            'ORDER BY SUBSTRING(YARD,2,2), ID_CONTEINER "
            
            
            SQl = SQl & " UNION "
            If Col_Patio <> 3 Then
                SQl = SQl & " SELECT ID_CONTEINER,YARD,LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,108),5)),MOTIVO,SUBSTRING(YARD,2,2) AS PILHA FROM OPERADOR..VW_PROX_MOV_CNTR_COL WHERE PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio) & " AND YARD LIKE "
            Else
                SQl = SQl & " SELECT ID_CONTEINER,YARD,LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,108),5)),MOTIVO,SUBSTRING(YARD,3,2) AS PILHA FROM OPERADOR..VW_PROX_MOV_CNTR_COL WHERE PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio) & " AND YARD LIKE "
            End If
            If Me.Option1(0).Value = True Then
                If Col_Patio <> 3 Then
                    SQl = SQl & "'" & Mid$(PosBusca, 1, 3) & "%'"
                Else
                    SQl = SQl & "'" & Mid$(PosBusca, 1, 4) & "%'"
                End If
            Else
                If Len(PosBusca) = 1 Then
                    SQl = SQl & "'" & Mid$(PosBusca, 1, 1) & "%'"
                ElseIf Len(PosBusca) = 2 Then
                    SQl = SQl & "'" & Mid$(PosBusca, 1, 2) & "%'"
                ElseIf Len(PosBusca) >= 3 Then
                    If Col_Patio <> 3 Then
                        SQl = SQl & "'" & Mid$(PosBusca, 1, 3) & "%'"
                    Else
                        SQl = SQl & "'" & Mid$(PosBusca, 1, 4) & "%'"
                    End If
                End If
            End If
            SQl = SQl & " and id_conteiner<>'" & Me.MskCntr & "'"
            'AND YARD NOT IN (SELECT YARD FROM OPERADOR..TB_YARD WHERE VALIDA=0) "
            SQl = SQl & " ORDER BY PILHA , ID_CONTEINER "
            
            
        Else
            SQl = "SELECT ID_CONTEINER,YARD,LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,108),5)),MOTIVO FROM sgipa..VW_PROX_MOV_CNTR_ALL WHERE YARD='" & PosBusca & "' and id_conteiner<>'" & Me.MskCntr & "'"
            'ORDER BY ID_CONTEINER "
            SQl = SQl & " UNION "
            SQl = SQl & " SELECT ID_CONTEINER,YARD,NULL,MOTIVO FROM OPERADOR..VW_PROX_MOV_CNTR_COL WHERE YARD='" & PosBusca & "' and id_conteiner<>'" & Me.MskCntr & "'"
            SQl = SQl & " ORDER BY ID_CONTEINER "
            
        End If
        
    Else
    
        If Col_Patio <> 3 Then
            SQl = "SELECT ID_CONTEINER,YARD,LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,108),5)),MOTIVO,SUBSTRING(YARD,2,2) AS PILHA FROM SGIPA..VW_PROX_MOV_CNTR_ALL WHERE PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio)
        Else
'           Sql = "SELECT ID_CONTEINER,YARD,MOTIVO,LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,108),5)),SUBSTRING(YARD,3,2) AS PILHA FROM OPERADOR..VW_PROX_MOV_CNTR WHERE PATIO=" & IIf(Col_Patio = 7, 1, Col_Patio)
            SQl = "SELECT ID_CONTEINER,YARD,LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,103),5)) + ' ' + LTRIM(LEFT(CONVERT(CHAR,DT_PREVISTA,108),5)),MOTIVO,SUBSTRING(YARD,3,2) AS PILHA FROM SGIPA..VW_PROX_MOV_CNTR_ALL WHERE PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio)
        End If
        
        If Len(PosBusca) = 1 Then
            SQl = SQl & " AND YARD LIKE '" & Mid$(PosBusca, 1, 1) & "%'"
        ElseIf Len(PosBusca) = 2 Then
            SQl = SQl & " AND YARD LIKE '" & Mid$(PosBusca, 1, 2) & "%'"
        ElseIf Len(PosBusca) >= 3 Then
            If Col_Patio <> 3 Then
                If Val(Mid$(PosBusca, 2, 2)) <> 0 Then
                    If Val(Mid$(PosBusca, 2, 2)) Mod 2 <> 0 Then
                        SQl = SQl & " AND (YARD LIKE '" & Mid$(PosBusca, 1, 3) & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 1) & Format(Val(Mid$(PosBusca, 2, 2)) + 1, "00") & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 1) & Format(Val(Mid$(PosBusca, 2, 2)) - 1, "00") & "%')"
                    Else
                         SQl = SQl & " AND (YARD LIKE '" & Mid$(PosBusca, 1, 3) & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 1) & Format(Val(Mid$(PosBusca, 2, 2)) + 1, "00") & "%'  OR YARD LIKE '" & Mid$(PosBusca, 1, 1) & Format(Val(Mid$(PosBusca, 2, 2)) + 2, "00") & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 1) & Format(Val(Mid$(PosBusca, 2, 2)) - 2, "00") & "%'  )"
                    End If
                Else
                    SQl = SQl & " AND YARD LIKE '" & Mid$(PosBusca, 1, 3) & "%'"
                End If
            Else
                If Val(Mid$(PosBusca, 3, 2)) <> 0 Then
                    If Val(Mid$(PosBusca, 3, 2)) Mod 2 <> 0 Then
                        SQl = SQl & " AND (YARD LIKE '" & Mid$(PosBusca, 1, 4) & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 2) & Format(Val(Mid$(PosBusca, 3, 2)) + 1, "00") & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 2) & Format(Val(Mid$(PosBusca, 3, 2)) - 1, "00") & "%')"
                    Else
                         SQl = SQl & " AND (YARD LIKE '" & Mid$(PosBusca, 1, 4) & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 2) & Format(Val(Mid$(PosBusca, 3, 2)) + 1, "00") & "%'  OR YARD LIKE '" & Mid$(PosBusca, 1, 2) & Format(Val(Mid$(PosBusca, 3, 2)) + 2, "00") & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 2) & Format(Val(Mid$(PosBusca, 3, 2)) - 2, "00") & "%'  )"
                    End If
                Else
                    SQl = SQl & " AND YARD LIKE '" & Mid$(PosBusca, 1, 4) & "%'"
                End If
            End If
        End If
        SQl = SQl & " and id_conteiner<>'" & Me.MskCntr & "' "
        'ORDER BY SUBSTRING(YARD,2,2), ID_CONTEINER "
        
        
        SQl = SQl & " UNION "
        
        If Col_Patio <> 3 Then
            SQl = SQl & " SELECT ID_CONTEINER,YARD,NULL,MOTIVO,SUBSTRING(YARD,2,2) AS PILHA FROM OPERADOR..VW_PROX_MOV_CNTR_COL WHERE PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio)
        Else
            SQl = SQl & " SELECT ID_CONTEINER,YARD,NULL,MOTIVO,SUBSTRING(YARD,3,2) AS PILHA FROM OPERADOR..VW_PROX_MOV_CNTR_COL WHERE PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio)
        End If
        
        If Len(PosBusca) = 1 Then
            SQl = SQl & " AND YARD LIKE '" & Mid$(PosBusca, 1, 1) & "%'"
        ElseIf Len(PosBusca) = 2 Then
            SQl = SQl & " AND YARD LIKE '" & Mid$(PosBusca, 1, 2) & "%'"
        ElseIf Len(PosBusca) >= 3 Then
            If Col_Patio <> 3 Then
                If Val(Mid$(PosBusca, 2, 2)) <> 0 Then
                    If Val(Mid$(PosBusca, 2, 2)) Mod 2 <> 0 Then
                        SQl = SQl & " AND (YARD LIKE '" & Mid$(PosBusca, 1, 3) & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 1) & Format(Val(Mid$(PosBusca, 2, 2)) + 1, "00") & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 1) & Format(Val(Mid$(PosBusca, 2, 2)) - 1, "00") & "%')"
                    Else
                         SQl = SQl & " AND (YARD LIKE '" & Mid$(PosBusca, 1, 3) & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 1) & Format(Val(Mid$(PosBusca, 2, 2)) + 1, "00") & "%'  OR YARD LIKE '" & Mid$(PosBusca, 1, 1) & Format(Val(Mid$(PosBusca, 2, 2)) + 2, "00") & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 1) & Format(Val(Mid$(PosBusca, 2, 2)) - 2, "00") & "%'  )"
                    End If
                Else
                    SQl = SQl & " AND YARD LIKE '" & Mid$(PosBusca, 1, 3) & "%'"
                End If
            Else
                If Val(Mid$(PosBusca, 3, 2)) <> 0 Then
                    If Val(Mid$(PosBusca, 3, 2)) Mod 2 <> 0 Then
                        SQl = SQl & " AND (YARD LIKE '" & Mid$(PosBusca, 1, 4) & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 2) & Format(Val(Mid$(PosBusca, 3, 2)) + 1, "00") & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 2) & Format(Val(Mid$(PosBusca, 3, 2)) - 1, "00") & "%')"
                    Else
                         SQl = SQl & " AND (YARD LIKE '" & Mid$(PosBusca, 1, 4) & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 2) & Format(Val(Mid$(PosBusca, 3, 2)) + 1, "00") & "%'  OR YARD LIKE '" & Mid$(PosBusca, 1, 2) & Format(Val(Mid$(PosBusca, 3, 2)) + 2, "00") & "%' OR YARD LIKE '" & Mid$(PosBusca, 1, 2) & Format(Val(Mid$(PosBusca, 3, 2)) - 2, "00") & "%'  )"
                    End If
                Else
                    SQl = SQl & " AND YARD LIKE '" & Mid$(PosBusca, 1, 4) & "%'"
                End If
            End If
        End If
        'Sql = Sql & " and id_conteiner<>'" & Me.MskCntr & "' ORDER BY SUBSTRING(YARD,2,2), ID_CONTEINER "
        SQl = SQl & " and id_conteiner<>'" & Me.MskCntr & "' ORDER BY PILHA , ID_CONTEINER "
        
    End If
        
    Adodc6.ConnectionString = ConnectionStringAdo
    Adodc6.RecordSource = SQl
    Adodc6.Refresh
    Set Adodc6.Recordset.ActiveConnection = Nothing
       
    Grid2.FormatString = "|^CONTEINER         |^POSICAO    |<DATA          |<MOTIVO                                            "

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
Me.txtEF2 = ""
Me.TxtGateIN = ""
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
Me.txtISO.Text = ""
Me.txtTara.Text = ""
Me.txtMaxGross.Text = ""
Me.txtLACRE1.Text = ""
Me.txtLACRE1.Tag = ""
Me.txtLACRE2.Text = ""
Me.txtLACRE2.Tag = ""
Me.txtLACRE3.Text = ""
Me.txtLACRE3.Tag = ""
Me.TxtLacreIPA.Text = ""
Me.TxtLacreLona.Text = ""
Me.TXTgmci.Text = ""
Me.Check1.Value = 0
Me.Check1.Tag = 0
Me.txtSetPoint.Text = ""
Me.txtMax.Text = ""
Me.txtMin.Text = ""


End Sub






Private Sub cmdExcluirAvaria_Click()
If Me.GridAvarias.Rows > 0 Then
    If Val(GridAvarias.TextMatrix(Me.GridAvarias.Row, Me.GridAvarias.Cols - 1)) <> 0 Then
    
        If Verifica_Finalizado(idPatio) = False Then
    
    
            Col_Mensagem = "Deseja excluir a avaria selecionada"
            Col_Imagem = 3
            frmMensagem2.Show vbModal
            If Col_Cancelado Then Exit Sub
        
        
            SQl = "DELETE FROM TB_AVARIAS_CONTEINER WHERE AUTONUM=" & Val(GridAvarias.TextMatrix(Me.GridAvarias.Row, Me.GridAvarias.Cols - 1))
            db.Execute (SQl)
            
            CarregaAvarias
        Else
        
            Frmmensagem.Label1.Caption = "Já consta finalização do Patio para a unidade"
            Frmmensagem.Show vbModal

        End If
    
    Else
    
        Frmmensagem.Label1.Caption = "Avaria apontada pelo Operador "
        Frmmensagem.Show vbModal
        
        Exit Sub
    
    End If
End If
End Sub

Private Sub cmdFinalizarAvaria_Click()

Dim SQl As String

Dim Cod_Etapa As String
'Joga codigo 2

'Sql = " SELECT COUNT(0) AS QTDE FROM VW_EGMCI_AVARIAS_TTL WHERE "
'Sql = Sql & " AUTONUMCNTR= " & idPatio
'Sql = Sql & " AND AUTONUM>0"
'Set tbavterm = New ADODB.Recordset
'tbavterm.Open Sql, db, 1, 1
'Cod_Etapa = 3
'If Not tbavterm.EOF Then
'    If tbavterm!QTDE > 0 Then Cod_Etapa = 4
'End If
'tbavterm.Close

If Verifica_Finalizado(idPatio) = False Then

    Col_Mensagem = "Confirma a finalizacao das avarias do conteiner " & Me.MskCntr.Text
    Col_Imagem = 3
    frmMensagem2.Show vbModal
    If Col_Cancelado Then Exit Sub


    SQl = " INSERT INTO TB_ETAPAS_AVARIA ("
    SQl = SQl & " AUTONUMCNTR, COD_ETAPA, DT_ETAPA) VALUES ("
    SQl = SQl & idPatio & ",2,Getdate())"
    db.Execute (SQl)
    
Else

    Frmmensagem.Label1.Caption = "Já consta finalização do Patio para a unidade"
    Frmmensagem.Show vbModal

End If
SetaControlesAV

End Sub

Private Sub SetaControlesAV()

If idPatio > 0 Then
    Dim SQl As String
    If Verifica_Finalizado(idPatio) = True Then
        'JA FINALIZOU
        Me.cmdExcluirAvaria.Enabled = False
        Me.cmdFinalizarAvaria.Enabled = False
        Me.cmdIncluirAvaria.Enabled = False
    Else
        If cmdExcluirAvaria.Tag = "1" Then Me.cmdExcluirAvaria.Enabled = True
        If cmdFinalizarAvaria.Tag = "1" Then Me.cmdFinalizarAvaria.Enabled = True
        If cmdIncluirAvaria.Tag = "1" Then Me.cmdIncluirAvaria.Enabled = True
    End If
End If

End Sub


Private Sub cmdIncluirAvaria_Click()
Dim Contar1 As Byte
Dim Contar2 As Byte

If Verifica_Finalizado(idPatio) = False Then

 
    For j = 0 To Me.ListLocal.ListCount - 1
        If Me.ListLocal.Selected(j) = True Then
            Contar1 = Contar1 + 1
        End If
    Next j
    
    For i = 0 To Me.ListTipo.ListCount - 1
        If Me.ListTipo.Selected(i) = True Then
            Contar2 = Contar2 + 1
        End If
    Next i
  
    If Contar1 = 0 Then
        Frmmensagem.Label1.Caption = "Nenhum Local selecionado"
        Frmmensagem.Show vbModal
        Exit Sub
    End If
    
    If Contar2 = 0 Then
        Frmmensagem.Label1.Caption = "Nenhum tipo selecionado"
        Frmmensagem.Show vbModal
        Exit Sub
    End If
    
    
    
    If Contar1 > 1 Then
        Frmmensagem.Label1.Caption = "Informe apenas um local"
        Frmmensagem.Show vbModal
        Exit Sub
    End If
    

    For j = 0 To Me.ListLocal.ListCount - 1
        If Me.ListLocal.Selected(j) = True Then
            For i = 0 To Me.ListTipo.ListCount - 1
                If Me.ListTipo.Selected(i) = True Then
            
                
            
                    Dim SQl As String
                    Dim Xtipo As String
                    Dim Xlocal As String
                    Xtipo = db.Execute("select code from sgipa..dte_tb_avarias where descr='" & Me.ListTipo.List(i) & "' and ident='T'")(0)
                    Xlocal = db.Execute("select code from sgipa..dte_tb_avarias where descr='" & Me.ListLocal.List(j) & "' and ident='L'")(0)
                    
                    SQl = "INSERT INTO TB_AVARIAS_CONTEINER(CNTR,LOCAL,TIPO,COMPLEMENTO,PROVIDENCIA,DT_CAD_AVARIA,USU_CAD_AVARIA) SELECT " & idPatio & ",'" & Xlocal & "','" & Xtipo & "',3,11,GETDATE()," & Cod_Usuario
                    SQl = SQl & " WHERE NOT EXISTS (SELECT AUTONUM FROM TB_AVARIAS_CONTEINER WHERE CNTR=" & idPatio & " AND LOCAL='" & Xlocal & "' AND TIPO='" & Xtipo & "')"
                    
                    db.Execute (SQl)
                
                End If
            Next i
        End If
    Next j
    CarregaAvarias
    
    Carrega_ListLocal
    Carrega_ListTipo

Else

    Frmmensagem.Label1.Caption = "Já consta finalização do Patio para a unidade"
    Frmmensagem.Show vbModal

End If
End Sub

Function Verifica_Finalizado(AutonumCntr As Long) As Boolean

Dim SQl As String
SQl = "Select autonum from tb_etapas_avaria where autonumcntr=" & AutonumCntr
SQl = SQl & " and cod_etapa in (2,3,4)"
Set tb1 = New ADODB.Recordset
tb1.Open SQl, db, 1, 1
If Not tb1.EOF Then
    Verifica_Finalizado = True
Else
    Verifica_Finalizado = False
End If

End Function


Private Sub CNTR_Change()
If InStr(1, Me.CNTR, "_") = 0 Then
    Call Busca_Cntr
End If
End Sub

Private Sub CNTR_LostFocus()
'Call Busca_Cntr
End Sub


Private Sub Command1_Click()
ModoRepesagem = 2
    
If DataCombo2.BoundText = "" Then
    MsgBox "Informe o motivo"
    Exit Sub
End If

SQl = "select count(*) from operador.dbo.tb_cad_motivo where autonum=" & DataCombo2.BoundText & " and flag_camera=1"

If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then
    
    DataCombo4.BoundText = ""
    With Frame2
        .Left = 60
        .Width = 3500
        .Top = 2250
        .Height = 915
        .Visible = True
        If Me.mskDestino.Text <> "" Then
            
        Me.DataCombo4.BoundText = Nnull(db.Execute("SELECT ISNULL(AUTONUM_CAMERA,0) FROM OPERADOR..TB_YARD WHERE PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio) & " AND (YARD='" & Me.mskDestino.Text & "' OR YARD_EQ='" & Me.mskDestino.Text & "')")(0), 0)
            
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

Private Sub Command5_Click()
Frame5.Visible = False
End Sub

Private Sub Command6_Click()


End Sub

Private Sub DataCombo3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        DataCombo3.BoundText = ""
    End If


End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then SendKeys "{tab}"

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
SQl = "SELECT autonum,descricao FROM " & Banco_Operador & "TB_cad_motivo where flag_ativo=1 ORDER BY descricao"
Adodc3.ConnectionString = ConnectionStringAdo
Adodc3.RecordSource = UCase(SQl)
Adodc3.ConnectionString = ConnectionStringAdo: Adodc3.Refresh
Set Adodc3.Recordset.ActiveConnection = Nothing


SQl = "SELECT autonum,descr FROM " & Banco_Operador & "TB_CAMERAS WHERE PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio) & " ORDER BY descr"
Adodc31.ConnectionString = ConnectionStringAdo
Adodc31.RecordSource = UCase(SQl)
Adodc31.ConnectionString = ConnectionStringAdo: Adodc31.Refresh
Set Adodc31.Recordset.ActiveConnection = Nothing


Me.txtPatio.Text = db.Execute("SELECT DESCR_RESUMIDO FROM OPERADOR..TB_PATIOS WHERE AUTONUM " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio))(0)


'PERMISSAO DE ACESSO NAS AVARIAS

SQl = "SELECT COUNT(1) CONTAR"
SQl = SQl & " From"
SQl = SQl & " TB_SYS_FUNCOES"
SQl = SQl & " ,TB_SYS_GRP_PERMISSOES"
SQl = SQl & " ,TB_SYS_USER_GRUPOS"
SQl = SQl & " ,TB_CAD_USUARIOS"
SQl = SQl & " Where"
SQl = SQl & " ((tb_cad_usuarios.autonum = tb_sys_user_grupos.autonumuser)"
SQl = SQl & " AND (tb_sys_funcoes.codfunc = tb_sys_grp_permissoes.codfunc)"
SQl = SQl & " AND (tb_sys_grp_permissoes.codgrupo = tb_sys_user_grupos.codgrupo)"
SQl = SQl & " AND (tb_cad_usuarios.AUTONUM = " & Cod_Usuario & ")"
SQl = SQl & " AND (tb_sys_funcoes.CODFUNC = 998)"
SQl = SQl & " AND tb_sys_grp_permissoes.CODTIPOPERM=5" 'ACESSO
SQl = SQl & " )"

If db.Execute(SQl)(0) = 0 Then
    Me.SSTab1.TabEnabled(3) = False
    
End If


SQl = "SELECT COUNT(1) CONTAR"
SQl = SQl & " From"
SQl = SQl & " TB_SYS_FUNCOES"
SQl = SQl & " ,TB_SYS_GRP_PERMISSOES"
SQl = SQl & " ,TB_SYS_USER_GRUPOS"
SQl = SQl & " ,TB_CAD_USUARIOS"
SQl = SQl & " Where"
SQl = SQl & " ((tb_cad_usuarios.autonum = tb_sys_user_grupos.autonumuser)"
SQl = SQl & " AND (tb_sys_funcoes.codfunc = tb_sys_grp_permissoes.codfunc)"
SQl = SQl & " AND (tb_sys_grp_permissoes.codgrupo = tb_sys_user_grupos.codgrupo)"
SQl = SQl & " AND (tb_cad_usuarios.AUTONUM = " & Cod_Usuario & ")"
SQl = SQl & " AND (tb_sys_funcoes.CODFUNC = 998)"
SQl = SQl & " AND tb_sys_grp_permissoes.CODTIPOPERM=3" 'EXCLUIR
SQl = SQl & " )"

If db.Execute(SQl)(0) = 0 Then
    Me.cmdExcluirAvaria.Visible = False
    Me.cmdExcluirAvaria.Tag = "0"
Else
     Me.cmdExcluirAvaria.Tag = "1"
End If

SQl = "SELECT COUNT(1) CONTAR"
SQl = SQl & " From"
SQl = SQl & " TB_SYS_FUNCOES"
SQl = SQl & " ,TB_SYS_GRP_PERMISSOES"
SQl = SQl & " ,TB_SYS_USER_GRUPOS"
SQl = SQl & " ,TB_CAD_USUARIOS"
SQl = SQl & " Where"
SQl = SQl & " ((tb_cad_usuarios.autonum = tb_sys_user_grupos.autonumuser)"
SQl = SQl & " AND (tb_sys_funcoes.codfunc = tb_sys_grp_permissoes.codfunc)"
SQl = SQl & " AND (tb_sys_grp_permissoes.codgrupo = tb_sys_user_grupos.codgrupo)"
SQl = SQl & " AND (tb_cad_usuarios.AUTONUM = " & Cod_Usuario & ")"
SQl = SQl & " AND (tb_sys_funcoes.CODFUNC = 998)"
SQl = SQl & " AND tb_sys_grp_permissoes.CODTIPOPERM=1" 'INCLUIR
SQl = SQl & " )"

If db.Execute(SQl)(0) = 0 Then
    Me.cmdIncluirAvaria.Visible = False
    Me.cmdIncluirAvaria.Tag = "0"
Else
    Me.cmdIncluirAvaria.Tag = "1"
End If

SQl = "SELECT COUNT(1) CONTAR"
SQl = SQl & " From"
SQl = SQl & " TB_SYS_FUNCOES"
SQl = SQl & " ,TB_SYS_GRP_PERMISSOES"
SQl = SQl & " ,TB_SYS_USER_GRUPOS"
SQl = SQl & " ,TB_CAD_USUARIOS"
SQl = SQl & " Where"
SQl = SQl & " ((tb_cad_usuarios.autonum = tb_sys_user_grupos.autonumuser)"
SQl = SQl & " AND (tb_sys_funcoes.codfunc = tb_sys_grp_permissoes.codfunc)"
SQl = SQl & " AND (tb_sys_grp_permissoes.codgrupo = tb_sys_user_grupos.codgrupo)"
SQl = SQl & " AND (tb_cad_usuarios.AUTONUM = " & Cod_Usuario & ")"
SQl = SQl & " AND (tb_sys_funcoes.CODFUNC = 999)"
SQl = SQl & " AND tb_sys_grp_permissoes.CODTIPOPERM=1" 'FINALIZAR AVARIA
SQl = SQl & " )"

If db.Execute(SQl)(0) = 0 Then
    Me.cmdFinalizarAvaria.Visible = False
    Me.cmdFinalizarAvaria.Tag = "0"
Else
    Me.cmdFinalizarAvaria.Tag = "1"
End If




End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub GridHist_Avaria_Click()
Dim SQl As String

End Sub

Private Sub MskCntr_LostFocus()

Call Busca_Dados

End Sub

Function Valida_Tara(ISO As String, Peso As Long, Tamanho As Integer) As String
Valida_Tara = ""
Dim RsTara As New ADODB.Recordset
Dim CC As String
Dim Tam As Byte
SQl = "select tamanho,tipo,nvl(pesonominal,0) tara from " & Banco_Operador & "tb_cad_conteiners where ISO='" & ISO & "'"
PRSet RsTara, SQl
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

SQl = "select tara_min_20, tara_max_20,tara_min_40, tara_max_40 from " & Banco_Sgipa & "dte_tb_tipos_conteiner where codigo='" & CC & "'"
PRSet RsTara, SQl
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
If Me.TXTYardAtual.Text <> "" And Me.MskCntr.Text <> "" Then
    If Me.TXTYardAtual = "" Or Me.TXTYardAtual = "GATEIN" Then
        If Me.mskDestino.Text <> "SAIDA" Then
            Me.DataCombo2.BoundText = 14
        End If
    End If
    If Me.TXTYardAtual <> "" And Me.MskCntr.Text <> "SAIDA" Then
        Me.DataCombo2.BoundText = 1
    End If
    If Me.mskDestino.Text = "SAIDA" Then
        Me.DataCombo2.BoundText = 8
    End If
End If

End Sub

Private Sub Option1_Click(Index As Integer)
Carrega_Grid2
If Index = 0 Then Me.Label1(8) = "FIADA ATUAL"
If Index = 1 Then Me.Label1(8) = "Q/FIADA DESTINO"
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 3 Then
    Carrega_ListLocal
    Carrega_ListTipo
    CarregaAvarias
End If

If SSTab1.Tab = 1 Then
    Carrega_Hist_Avaria
     If MskCntr.Tag = "R" Then
        SQl = "select hs.*,usu.usuario as nmusuario, cm.descricao as desc_motivo"
        SQl = SQl & " from redex..tb_patio cc"
        SQl = SQl & " inner join " & Banco_Operador & "tb_hist_shifting hs on cc.autonum_patio = hs.cntr and hs.tipo='R'"
        SQl = SQl & " left  join redex..tb_cad_usuarios usu on hs.usuario = usu.autonum_usu"
        SQl = SQl & " left join " & Banco_Operador & "tb_cad_motivo cm on hs.motivo = cm.autonum"
        SQl = SQl & " where cc.id_conteiner='" & MskCntr & "'"
        SQl = SQl & " order by hs.data DESC"
    ElseIf Me.MskCntr.Tag = "A" Then
        SQl = "select hs.*,usu.usuario as nmusuario, cm.descricao as desc_motivo"
        SQl = SQl & " from " & Banco_Sgipa & "tb_armazens_ipa cc"
        SQl = SQl & " inner join " & Banco_Operador & "tb_hist_shifting hs on cc.autonum = hs.cntr and hs.tipo='A'"
        SQl = SQl & " left  join " & Banco_Sgipa & "tb_cad_usuarios usu on hs.usuario = usu.autonum"
        SQl = SQl & " left join " & Banco_Operador & "tb_cad_motivo cm on hs.motivo = cm.autonum"
        SQl = SQl & " where cc.descr='" & MskCntr & "'"
        SQl = SQl & " order by hs.data DESC"
    Else
       SQl = "select hs.*,usu.usuario as nmusuario, cm.descricao as desc_motivo"
        SQl = SQl & " from " & Banco_Sgipa & "tb_cntr_bl cc"
        SQl = SQl & " inner join " & Banco_Operador & "tb_hist_shifting hs on cc.autonum = hs.cntr and hs.tipo='I'"
        SQl = SQl & " left  join " & Banco_Sgipa & "tb_cad_usuarios usu on hs.usuario = usu.autonum"
        SQl = SQl & " left join " & Banco_Operador & "tb_cad_motivo cm on hs.motivo = cm.autonum"
        SQl = SQl & " where cc.id_conteiner='" & MskCntr & "'"
        SQl = SQl & " order by hs.data DESC"
    End If
    Adodc1.ConnectionString = ConnectionStringAdo
    Adodc1.RecordSource = UCase(SQl)
    Adodc1.ConnectionString = ConnectionStringAdo: Adodc1.Refresh
    Set Adodc1.Recordset.ActiveConnection = Nothing

End If


End Sub

Private Sub Carrega_Hist_Avaria()

Me.GridHist_Avaria.Clear

Dim SQl As String

SQl = "select (RTRIM(B.DESCR) + ' ' + RTRIM(C.DESCR)) AS AVARIA,"
SQl = SQl & " ISNULL(u.usuario,' ') as USUARIO,"
SQl = SQl & " RTRIM(ISNULL(CONVERT(CHAR, A.DT_CAD_AVARIA,103),'')) + ' ' + ISNULL(CONVERT(CHAR, A.DT_CAD_AVARIA,108),'') AS DT_AVARIA"
SQl = SQl & " From"
SQl = SQl & " sgipa.dbo.tb_avarias_conteiner a inner join"
SQl = SQl & " sgipa.dbo.dte_tb_avarias b on a.local=b.code and B.ident='L' inner join"
SQl = SQl & " sgipa.dbo.dte_tb_avarias c on a.tipo=c.code and C.IDENT='T' left join"
SQl = SQl & " sgipa.dbo.tb_cad_usuarios u on a.usu_cad_avaria=u.autonum"
SQl = SQl & " Where"
SQl = SQl & " a.CNTR = " & idPatio

Me.GridHist_Avaria.Clear
Me.GridHist_Avaria.FormatString = "    |<Data                          |<Avaria                      |<Usuário                               "


Dim tb1 As ADODB.Recordset
PRSet tb1, SQl

If tb1.RecordCount > 0 Then
    Me.GridHist_Avaria.Rows = tb1.RecordCount + 1
Else
    Me.GridHist_Avaria.Rows = 2
End If

Dim i As Integer
i = 0

While Not tb1.EOF
    i = i + 1
    Me.GridHist_Avaria.TextMatrix(i, 1) = Nnull(tb1!dt_avaria, 1)
    Me.GridHist_Avaria.TextMatrix(i, 2) = Nnull(tb1!avaria, 1)
    Me.GridHist_Avaria.TextMatrix(i, 3) = Nnull(tb1!usuario, 1)
    
    tb1.MoveNext
Wend



tb1.Close


End Sub


Private Sub CarregaAvarias()
Dim SQl As String
SQl = " select AUTONUM,AUTONUMCNTR,LOCAL,"
SQl = SQl & " TIPO,ID_CONTEINER,Viagem "
SQl = SQl & " from VW_EGMCI_AVARIAS_TTL "
SQl = SQl & " where "
SQl = SQl & " autonumcntr=" & idPatio
Me.GridAvarias.Clear
Me.GridAvarias.FormatString = "    |<Localização                              |<Tipo                           |"


Dim tb1 As ADODB.Recordset
PRSet tb1, SQl

If tb1.RecordCount > 0 Then
    Me.GridAvarias.Rows = tb1.RecordCount + 1
Else
    Me.GridAvarias.Rows = 2
End If

Dim i As Integer
i = 0

While Not tb1.EOF
    i = i + 1
    Me.GridAvarias.TextMatrix(i, 1) = Nnull(tb1!local, 1)
    Me.GridAvarias.TextMatrix(i, 2) = Nnull(tb1!Tipo, 1)
    Me.GridAvarias.TextMatrix(i, 3) = Nnull(tb1!autonum, 0)
    If Nnull(tb1!autonum, 0) = 0 Then
        Me.GridAvarias.Row = i
        For j = 1 To Me.GridAvarias.Cols - 1
            Me.GridAvarias.Col = j
            Me.GridAvarias.CellBackColor = Me.Shape1.FillColor
        Next
    End If
    tb1.MoveNext
Wend



tb1.Close



End Sub


Private Sub Carrega_ListLocal()

Me.ListLocal.Clear
Dim SQl As String
Dim tb1 As ADODB.Recordset

SQl = "Select descr from sgipa..dte_tb_avarias where ident='L' order by descr"
PRSet tb1, SQl
While Not tb1.EOF
    Me.ListLocal.AddItem (Nnull(tb1!Descr, 1))
    tb1.MoveNext
Wend


End Sub
Private Sub Carrega_ListTipo()

Me.ListTipo.Clear
Dim SQl As String
Dim tb1 As ADODB.Recordset

SQl = "Select descr from sgipa..dte_tb_avarias where ident='T' order by descr"
PRSet tb1, SQl
While Not tb1.EOF
    Me.ListTipo.AddItem (Nnull(tb1!Descr, 1))
    tb1.MoveNext
Wend

End Sub

Private Sub TXTYardAtual_LostFocus()
If Me.TXTYardAtual.Text <> "" And Me.MskCntr.Text <> "" Then
    If Me.TXTYardAtual = "" Or Me.TXTYardAtual = "GATEIN" Then
        If Me.mskDestino.Text <> "SAIDA" Then
            Me.DataCombo2.BoundText = 14
        End If
    End If
    If Me.TXTYardAtual <> "" And Me.MskCntr.Text <> "SAIDA" Then
        Me.DataCombo2.BoundText = 1
    End If
    If Me.mskDestino.Text = "SAIDA" Then
        Me.DataCombo2.BoundText = 8
    End If
End If

End Sub
