VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form Inventario_CSPatio 
   Appearance      =   0  'Flat
   BackColor       =   &H00BA9F8D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventario"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3600
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Placa Retorno DTA"
      Height          =   1215
      Left            =   210
      TabIndex        =   43
      Top             =   1200
      Visible         =   0   'False
      Width           =   3345
      Begin VB.CommandButton Command2 
         Caption         =   "X"
         Height          =   255
         Left            =   2985
         TabIndex        =   45
         Top             =   150
         Width           =   300
      End
      Begin VB.TextBox txtPlacaDta 
         Height          =   405
         Left            =   780
         MaxLength       =   8
         TabIndex        =   44
         Top             =   375
         Width           =   1785
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "[F3] DTA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1380
      TabIndex        =   42
      Top             =   4170
      Width           =   600
   End
   Begin VB.CommandButton cmdEscape 
      Caption         =   "[ESC]SAIR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2790
      TabIndex        =   40
      Top             =   4170
      Width           =   795
   End
   Begin VB.CommandButton cmdSalva 
      Caption         =   "[F5]SALVA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1995
      TabIndex        =   39
      Top             =   4170
      Width           =   765
   End
   Begin VB.CommandButton cmdLimpa 
      Caption         =   "[F2]LIMPA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   675
      TabIndex        =   38
      Top             =   4170
      Width           =   675
   End
   Begin VB.CommandButton cmdHistorico 
      Caption         =   "[F1]HIST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   30
      TabIndex        =   37
      Top             =   4170
      Width           =   615
   End
   Begin VB.Frame frmHist 
      Appearance      =   0  'Flat
      BackColor       =   &H00BA9F8D&
      Caption         =   "Histórico da Movimentação"
      ForeColor       =   &H00FFFFFF&
      Height          =   4155
      Left            =   3810
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   3585
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   2130
         Top             =   3120
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
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
         Bindings        =   "InventarioCSPatio.frx":0000
         Height          =   3225
         Left            =   90
         TabIndex        =   36
         Top             =   270
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   5689
         _Version        =   393216
         Cols            =   4
         FormatString    =   " |^Armazem|^Local             |^Data                         "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
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
   End
   Begin VB.TextBox txtAutonum_Cs_Yard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3270
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.TextBox txtQtdeM 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3150
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   180
      Width           =   405
   End
   Begin VB.TextBox txtMarcante 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1155
      MaxLength       =   9
      TabIndex        =   0
      Top             =   180
      Width           =   2025
   End
   Begin MSAdodcLib.Adodc AdodcItem 
      Height          =   330
      Left            =   8520
      Top             =   990
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
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtCritica 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   645
      Left            =   15
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3525
      Width           =   3585
   End
   Begin VB.TextBox txtMOV 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox txtIMO 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2910
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2370
      Width           =   645
   End
   Begin VB.TextBox txtLocal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2940
      Width           =   1035
   End
   Begin VB.TextBox txtQtde 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1020
      Width           =   825
   End
   Begin VB.TextBox txtQtdePos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2940
      Width           =   405
   End
   Begin VB.TextBox txtLocalPos 
      Alignment       =   2  'Center
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
      Height          =   300
      Left            =   2610
      TabIndex        =   4
      Top             =   2940
      Width           =   915
   End
   Begin VB.TextBox txtMarca 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtConteiner 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2370
      Width           =   1365
   End
   Begin VB.TextBox txtEntrada 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2100
      Width           =   2415
   End
   Begin VB.TextBox txtMercadoria 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1290
      Width           =   2415
   End
   Begin VB.TextBox txtEmbalagem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1950
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1605
   End
   Begin VB.TextBox txtCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1830
      Width           =   2415
   End
   Begin VB.TextBox txtBL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1950
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   450
      Width           =   1605
   End
   Begin VB.TextBox txtLote 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1140
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   1
      Top             =   450
      Width           =   825
   End
   Begin MSDataListLib.DataCombo cbItem 
      Bindings        =   "InventarioCSPatio.frx":0015
      Height          =   270
      Left            =   1140
      TabIndex        =   2
      Top             =   750
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   476
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      Appearance      =   0
      Style           =   2
      BackColor       =   16777215
      ListField       =   "Display"
      BoundColumn     =   "Autonum"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo cbMotivoPos 
      Bindings        =   "InventarioCSPatio.frx":002D
      Height          =   270
      Left            =   1140
      TabIndex        =   5
      Top             =   3240
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   476
      _Version        =   393216
      MatchEntry      =   -1  'True
      Appearance      =   0
      Style           =   2
      BackColor       =   16777215
      ListField       =   "Display"
      BoundColumn     =   "Autonum"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdodcArm 
      Height          =   330
      Left            =   5520
      Top             =   4320
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
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc AdodcMotivo 
      Height          =   330
      Left            =   5430
      Top             =   5010
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
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label LbArmPadrao 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "..............................."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   45
      TabIndex        =   41
      Top             =   3075
      Width           =   1065
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MARCANTE/QTDE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   17
      Left            =   -60
      TabIndex        =   32
      Top             =   300
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MOTIVO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   -30
      TabIndex        =   30
      Top             =   3300
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MOV. AGEND."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   14
      Left            =   -60
      TabIndex        =   28
      Top             =   2700
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "IMO "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   12
      Left            =   2400
      TabIndex        =   26
      Top             =   2460
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   -30
      TabIndex        =   23
      Top             =   810
      Width           =   1155
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[F1]HISTORICO     [F2]LIMPA    [F5]SALVA      [ESC]-SAI "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   90
      TabIndex        =   22
      Top             =   4170
      Width           =   3390
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MARCA "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   11
      Left            =   -30
      TabIndex        =   20
      Top             =   1650
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "QTDE / LOCAL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Index           =   8
      Left            =   330
      TabIndex        =   13
      Top             =   2925
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CONTÊINER "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   7
      Left            =   0
      TabIndex        =   12
      Top             =   2460
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " ENT / DESOVA "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   6
      Left            =   0
      TabIndex        =   11
      Top             =   2190
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MERCADORIA "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   -30
      TabIndex        =   10
      Top             =   1350
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "IMPORTADOR "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   -30
      TabIndex        =   9
      Top             =   1920
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "QTDE / EMB."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   -60
      TabIndex        =   8
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LOTE/BL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   -60
      TabIndex        =   7
      Top             =   540
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MOVIMENTACAO - CARGA SOLTA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BA9F8D&
      Height          =   195
      Index           =   32
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3570
   End
End
Attribute VB_Name = "Inventario_CSPatio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function Carrega_Arm_Padrao(Lote As String) As String

Dim SQl As String
SQl = "select c.descr from sgipa..tb_bl bl"
SQl = SQl & " inner join sgipa..TB_TIPO_sub_classe b on bl.sub_classe = b.autonum"
SQl = SQl & " inner join sgipa..tb_armazens_ipa  c on b.local = c.AUTONUM"
SQl = SQl & " where bl.autonum=" & Lote
Dim Rs As New ADODB.Recordset
PRSet Rs, SQl
If Not Rs.EOF Then
    Carrega_Arm_Padrao = Rs!Descr
Else
    Carrega_Arm_Padrao = ""
End If
Rs.Close

End Function


Private Sub cbItem_Change()
If Val(Me.cbItem.BoundText) > 0 Then
    Carrega_Dados_Item (Me.cbItem.BoundText)
End If
End Sub

Private Sub cbItem_Click(Area As Integer)

'Carrega_Dados_Item (Me.cbItem.BoundText)

End Sub

Private Sub cbItem_KeyDown(KeyCode As Integer, Shift As Integer)

'Carrega_Dados_Item (Me.cbItem.BoundText)

End Sub

Private Sub Carrega_Dados_Item(ByVal id_Gravacao As String)

Me.txtQtde.Text = ""
Me.txtEmbalagem.Text = ""
Me.txtLocal.Text = ""

SQl = "SELECT "
SQl = SQl & " QTDE,"
SQl = SQl & " EMBALAGEM,"
SQl = SQl & " (DESCR_ARMAZEM + ' ' + POSICAO) AS LOCAL "
SQl = SQl & " "
SQl = SQl & " FROM OPERADOR..VW_INVENT_ARMAZEM"
SQl = SQl & " WHERE ID_GRAVACAO='" & id_Gravacao & "'"

Set tb1 = New ADODB.Recordset
tb1.Open SQl, db, 1, 1
If Not tb1.EOF Then
    Me.txtQtde = Nnull(tb1!QTDE, 0)
    Me.txtEmbalagem = Nnull(tb1!embalagem, 1)
    Me.txtLocal = Nnull(tb1!local, 1)
    If Trim(Me.txtLocal.Text) = "" Or Trim(Me.txtLocal.Text) = "-" Then Me.cbMotivoPos.BoundText = 14
End If
tb1.Close



End Sub


Private Sub cmdEscape_Click()
If Me.frmHist.Visible = False Then
    Unload Me
Else
    Me.frmHist.Visible = False
End If
End Sub

Private Sub cmdHistorico_Click()
If Me.frmHist.Visible = True Then
    frmHist.Visible = False
Else
    frmHist.Visible = True
    Carrega_Hist
End If
End Sub

Private Sub cmdLimpa_Click()
Limpa
Me.txtMarcante.Text = ""
End Sub

Private Sub cmdSalva_Click()
ModoRepesagem = 1

If txtLote.Text <> "" Then
    Dim SQl As String
    SQl = "select isnull(flag_aviso_temperatura_coletor,0) from sgipa..tb_bl bl"
    SQl = SQl & " inner join operador..tb_patios p on bl.patio=p.autonum"
    SQl = SQl & " where bl.autonum=" & txtLote.Text
    SQl = SQl & " AND ISNULL(BL.SUB_CLASSE,0)<>0"
    
    If Nnull(db.Execute(SQl).Fields(0), 0) = 1 Then
        X = MsgBox("CONFIRA CONDIÇÕES DE TEMPERATURA E UMIDADE", vbExclamation)
    End If


End If

Salvar
If Me.txtCritica.ForeColor = vbGreen Then Me.txtMarcante.Text = ""
End Sub

Private Sub Command1_Click()
txtPlacaDta.Text = ""
Frame1.Visible = True



End Sub

Private Sub Command2_Click()
Frame1.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Me.frmHist.Visible = False Then
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
    If KeyCode = vbKeyF2 Then
       
        Limpa
        Me.txtMarcante.Text = ""
    End If
    If KeyCode = vbKeyF3 Then
        txtPlacaDta.Text = ""

        Frame1.Visible = True
    End If
    
    If KeyCode = vbKeyF5 Then
        Salvar
        If Me.txtCritica.ForeColor = vbGreen Then Me.txtMarcante.Text = ""
    End If
    
    
    If KeyCode = vbKeyF1 Then
        frmHist.Visible = True
        Carrega_Hist
    End If

    
Else
    
    If KeyCode = vbKeyEscape Then
        frmHist.Visible = False
    End If

    If KeyCode = vbKeyF1 Then
        frmHist.Visible = False
    End If

End If



End Sub
Private Sub Carrega_Hist()

SQl = " SELECT ISNULL(B.DESCR,''),upper(isnull(A.YARD,'')),rtrim(ltrim(convert(char,A.DT_MOV,3))) + ' ' + left(rtrim(ltrim(convert(char,A.DT_MOV,108))),5) ,U.USUARIO FROM SGIPA..TB_HIST_SHIFTING_CS A LEFT JOIN SGIPA..TB_ARMAZENS_IPA B"
SQl = SQl & " ON A.ARMAZEM=B.AUTONUM "
SQl = SQl & " LEFT JOIN SGIPA..TB_CAD_USUARIOS U ON A.USUARIO=U.AUTONUM "
SQl = SQl & " WHERE A.MARCANTE=" & Val(Nnull(Me.txtMarcante, 0))
SQl = SQl & " ORDER BY A.DT_MOV DESC "

Me.Adodc1.ConnectionString = ConnectionString
Me.Adodc1.RecordSource = SQl
Me.Adodc1.Refresh
Set Me.Adodc1.Recordset.ActiveConnection = Nothing

Me.Grid1.FormatString = " |^Armazem|^Local       |^Data               |<Usuario             "


End Sub
Private Sub Salvar(Optional Placa As String = "")
 
 If Validar() Then

    Dim SQl As String
    Dim SQL2 As String
    
    Dim AutonumCs As Long
    Dim AutonumCY As Long

    Dim Tipo As String
    
    Dim FlagChronos As Boolean
    FlagChronos = True
    
    
    If Me.Label2(32).Tag <> "C" Then FlagChronos = False
    
    If Me.Label2(32).Tag <> "M" Then 'RECONHECIMENTO
        
        If FlagChronos Then
        
            Tipo = Left(Me.cbItem.BoundText, 1)
        
            If Tipo = "C" Then
                AutonumCs = Val(Mid(Me.cbItem.BoundText, 2))
            ElseIf Tipo = "Y" Then
                AutonumCY = Val(Mid(Me.cbItem.BoundText, 2))
                SQl = "SELECT AUTONUM_CS FROM SGIPA..TB_CARGA_SOLTA_YARD WHERE AUTONUM=" & AutonumCY
                        
                Set Rst = New ADODB.Recordset
                Rst.Open SQl, db, 1, 1
                If Not Rst.EOF Then
                    AutonumCs = Nnull(Rst!AUTONUM_CS, 0)
                End If
                Rst.Close
            End If
        
        Else
            
            Tipo = Left(Me.txtLote.Tag, 1)
        
            If Tipo = "V" Then
                AutonumCs = Val(Mid(Me.txtLote.Tag, 2))
            ElseIf Tipo = "W" Then
                AutonumCY = Val(Mid(Me.txtLote.Tag, 2))
                SQl = "SELECT AUTONUM_CS FROM SGIPA..TB_CARGA_SOLTA_YARD_AC WHERE AUTONUM=" & AutonumCY
                        
                Set Rst = New ADODB.Recordset
                Rst.Open SQl, db, 1, 1
                If Not Rst.EOF Then
                    AutonumCs = Nnull(Rst!AUTONUM_CS, 0)
                End If
                Rst.Close
            End If
            
        End If
        
    
    End If
        
        
    'Sql = "INSERT INTO "

    'If FlagChronos Then
    '    Sql = Sql & "SGIPA..TB_CARGA_SOLTA_YARD"
    'Else
    '    Sql = Sql & "SGIPA..TB_CARGA_SOLTA_YARD_AC"
    'End If
    
    'Sql = Sql & "(AUTONUM_CS,PATIO,ARMAZEM,YARD,ORIGEM,QUANTIDADE,MOTIVO,AUDIT_94,FL_FRENTE,FL_FUNDO,FL_LE,FL_LD"
    
    'Sql = Sql & ") VALUES "
    'Sql = Sql & "(" & AutonumCs & "," & IIf(Col_Patio = 7, 1, Col_Patio) & ",0,'" & Me.txtLocalPos.text & "','I'," & Val(Me.txtQtdePos.text) & "," & Me.cbMotivoPos.BoundText
    'Sql = Sql & ",0,0,0,0,0"
    'Sql = Sql & ")"
    
    
    'SQL2 = ""
    
    'Dim SQL3 As String
    'Dim SQL4 As String
        
    
    
    'db.BeginTrans
        
    'db.Execute Sql
        
    'Dim Autonum_Cs_Yard As Long
    

    'If FlagChronos Then
    '    Autonum_Cs_Yard = db.Execute("SELECT IDENT_CURRENT('TB_CARGA_SOLTA_YARD')")(0)
    'Else
    '    Autonum_Cs_Yard = db.Execute("SELECT IDENT_CURRENT('TB_CARGA_SOLTA_YARD_AC')")(0)
    'End If
        
        
    'If SQL2 <> "" Then db.Execute SQL2
    'If SQL3 <> "" Then db.Execute SQL3
    'If SQL4 <> "" Then db.Execute SQL4
        
    
        
'    If Me.Label2(32).Tag <> "M" Then
'
'
'        If Me.txtMarcante <> "" Then
'            Dim SqlM As String
'            SqlM = "UPDATE SGIPA..TB_MARCANTES SET "
'
'            If FlagChronos Then
'                SqlM = SqlM & " AUTONUM_CARGA=" & AutonumCs & ""
'                SqlM = SqlM & " ,AUTONUM_CS_YARD=" & Autonum_Cs_Yard
'            Else
'                SqlM = SqlM & " AUTONUM_CARGA_AC=" & AutonumCs & ""
'                SqlM = SqlM & " ,AUTONUM_CS_YARD_AC=" & Autonum_Cs_Yard
'            End If
'
'            SqlM = SqlM & ",VOLUMES=" & Val(Me.txtQtdePos.text) & " ,DT_ASSOCIACAO=GETDATE()"
'
'            SqlM = SqlM & " WHERE AUTONUM= " & Val(Me.txtMarcante)
'            'db.Execute SqlM
'        End If
'
'    End If
                
        
        
    'Sql = "INSERT INTO SGIPA..TB_HIST_SHIFTING_CS(MARCANTE,ARMAZEM,PATIO,YARD,DT_MOV,USUARIO) VALUES ("
    'Sql = Sql & Val(Me.txtMarcante.text) & ","
    'Sql = Sql & "0,"
    'Sql = Sql & IIf(Col_Patio = 7, 1, Col_Patio) & ","
    'Sql = Sql & "'" & Me.txtLocalPos.text & "',"
    'Sql = Sql & "GETDATE()," & Cod_Usuario & ")"
    'db.Execute Sql
        
    
    'db.CommitTrans
    
    
    
    'verifica se armazem possui bloqueio de IMO Inflamavel
    If txtLocal.Text = "ARM 22" And Nnull(AutonumCs, 0) > 0 Then
        SQl = "SELECT imo FROM SGIPA..TB_CARGA_SOLTA WHERE autonum=" & AutonumCs & " and isnull(imo,'')<>''"
        Dim RsI As New ADODB.Recordset
        PRSet RsI, SQl
        If Not RsI.EOF Then
            If RsI!imo <> "4.3" Then
                Frmmensagem.Label1.Caption = " Local Bloqueado, carga Inflamável"
                Frmmensagem.Show vbModal
                Exit Sub
            Else
                If UCase(Me.txtLocalPos.Text) <> "L10" Then
                    Frmmensagem.Label1.Caption = " Local Bloqueado, carga Inflamável (L10)"
                    Frmmensagem.Show vbModal
                    Exit Sub
                End If
            End If
        End If
        Rs.Close
    End If
    
    
    'Se o Destino é BK precisa estar autorizado
If Me.txtLocalPos.Text = "BK" And AutonumCs > 0 Then
    SQl = " SELECT ISNULL(MAX(AUTONUM),0)  LOTE FROM SGIPA..TB_BL WHERE ISNULL(FLAG_SEL_TRANSF_BK,0)=1 AND AUTONUM IN ( SELECT BL FROM TB_CARGA_SOLTA WHERE AUTONUM= " & AutonumCs & ")"
    P_BL = db.Execute(SQl)(0)
    If P_BL = 0 Then
        Frmmensagem.Label1.Caption = "lote sem autorização para Transferência BK"
        Frmmensagem.Show vbModal
        mskDestino.SetFocus
        Exit Sub
    End If
    P_Cntr = 0
    FrmVeiculoServ.Show vbModal
    P_BL = 0
End If




If Me.txtLocalPos.Text = "DTA" And AutonumCs > 0 Then
    SQl = "SELECT BL FROM TB_CARGA_SOLTA WHERE AUTONUM= " & AutonumCs
    P_BL = db.Execute(SQl)(0)
    LoteDTA = 0
    If P_BL = 0 Then
        Frmmensagem.Label1.Caption = "lote não enocntrado"
        Frmmensagem.Show vbModal
        mskDestino.SetFocus
        Exit Sub
    End If
    LoteDTA = AutonumCs
    P_Cntr = 0
    
    FrmVeiculoServ.Show vbModal
    
    P_BL = 0
    LoteDTA = 0
    Me.txtLocalPos.Text = "CAM"
End If

    
    
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = ConnectionString
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "[Sgipa].[dbo].[pcMovimentaCS_Sequence]"
    cmd.Parameters.Append cmd.CreateParameter("Tag", adVarChar, adParamInput, 1, Me.Label2(32).Tag)
    cmd.Parameters.Append cmd.CreateParameter("Tipo", adVarChar, adParamInput, 1, Tipo)
    cmd.Parameters.Append cmd.CreateParameter("FlagChronos", adInteger, adParamInput, 1, IIf(FlagChronos, 1, 0))
    cmd.Parameters.Append cmd.CreateParameter("AutonumCs", adInteger, adParamInput, 10, AutonumCs)
    cmd.Parameters.Append cmd.CreateParameter("Patio", adInteger, adParamInput, 2, IIf(Col_Patio = 7, 1, Col_Patio))
    cmd.Parameters.Append cmd.CreateParameter("Yard", adVarChar, adParamInput, 10, Me.txtLocalPos.Text)
    cmd.Parameters.Append cmd.CreateParameter("Quantidade", adInteger, adParamInput, 8, Val(Me.txtQtdePos.Text))
    cmd.Parameters.Append cmd.CreateParameter("Motivo", adInteger, adParamInput, 8, Me.cbMotivoPos.BoundText)
    cmd.Parameters.Append cmd.CreateParameter("Marcante", adInteger, adParamInput, 8, Val(Me.txtMarcante.Text))
    cmd.Parameters.Append cmd.CreateParameter("Usuario", adInteger, adParamInput, 8, Cod_Usuario)
    cmd.Execute
     
    Set cmd = Nothing
    
    
    
    Me.txtCritica = "POSICIONADA COM SUCESSO"
    Me.txtCritica.ForeColor = vbGreen
      
    Limpa
    
    Me.txtLote.Text = ""
    Me.txtLote.Tag = ""
    Me.txtMarcante.SetFocus

Else
    
    If Nnull(Me.txtCritica.Tag, 1) <> "" Then
        
        Dim X As Integer
        
        X = MsgBox("Deseja movimentar o marcante " & Nnull(Me.txtCritica.Tag, 1), 36 + vbDefaultButton2)
        If X = vbYes Then
            Limpa
            Me.txtMarcante.Text = Nnull(Me.txtCritica.Tag, 1)
            Call txtMarcante_Change
        End If
        
    End If

End If
End Sub


Private Sub Form_Load()
Me.Top = 0
Me.Left = 0

Limpa
CarregarArmazens
CarregarMotivos



End Sub

Sub Carrega_Dados_Lote(ByVal QualLote As String, ByVal Autonum_Cs_Yard As Long)


Set tb1 = New ADODB.Recordset
SQl = "SELECT * FROM sgipa.dbo.fnInvent_Sistemas_Arm_CollOTE(" & QualLote & "," & Val(Me.txtMarcante) & ")"
tb1.Open SQl, db, 1, 1
        Limpa

If Not tb1.EOF Then

    Me.txtMercadoria.Text = Nnull(tb1!MERCADORIA, 1)
    Me.txtMarca.Text = Nnull(tb1!Marca, 1)
    Me.txtEntrada.Text = Nnull(tb1!Entrada, 1)
    Me.txtConteiner.Text = Nnull(tb1!CNTR_DESOVA, 1)
    Me.txtCliente.Text = Nnull(tb1!Importador, 1)
    
    
    Me.txtMOV.Text = Nnull(tb1!MOTIVO_PROX_MVTO, 1)
    If Nnull(tb1!imo, 1) <> "0" Then
        Me.txtIMO.Text = Nnull(tb1!imo, 1)
    Else
        Me.txtIMO.Text = ""
    End If
    
    Me.txtBL.Text = Nnull(tb1!BL, 1)
    Me.txtLocal.Text = Nnull(tb1!DESCR_ARMAZEM, 1) & " " & Nnull(tb1!Posicao, 1)
    
    If Trim(Replace(Trim(Me.txtLocal.Text), "-", "")) = "" Then
        Me.cbMotivoPos.BoundText = 14
    End If
    

    Call CarregarItens(QualLote, Nnull(Me.txtMarcante.Text, 0), Autonum_Cs_Yard)
    
    
   
    
Else
    Me.txtCritica = "Lote não encontrado"
    Me.txtCritica.ForeColor = &H80C0FF
End If
tb1.Close

    
End Sub

Private Sub CarregarItens(ByVal QualLote As String, Optional Marcante As Long = 0, Optional Autonum_Cs_Yard As Long = 0)

Dim SQl As String

'Set tb1 = New ADODB.Recordset

'tb1.Open Sql, db, 1, 1

'Sql = " select ID_GRAVACAO as autonum, ( REPLACE(rtrim(convert(char,QTDE)),'.000','') +  '/' + rtrim(convert(char,QTDE_CAPTADA)) + ' ' + ISNULL(EMBALAGEM, ' ') + ' ' +  ISNULL(DESCR_ARMAZEM,' ') + ' ' + ISNULL(POSICAO, ' ')  ) AS DISPLAY"
'Sql = Sql & " FROM OPERADOR..VW_INVENT_ARMAZEM"
'Sql = Sql & " Where"
'Sql = Sql & " lote = " & QualLote
'
'If Marcante > 0 Then
'    Sql = Sql & " and Marcante =" & Marcante
'End If
'
'If Autonum_Cs_Yard > 0 Then
'    Sql = Sql & " and Autonum_CS_Yard =" & Autonum_Cs_Yard
'Else
'    Sql = Sql & " and Autonum_CS_Yard = 0 "
'End If
'
'Sql = Sql & "  AND QTDE>0 ORDER BY DESCR_ARMAZEM,POSICAO"



SQl = "SELECT * FROM sgipa.dbo.fnInvent_Sistemas_Arm_ColItem(" & QualLote & "," & Val(Me.txtMarcante) & "," & Autonum_Cs_Yard & ")"
Me.AdodcItem.ConnectionString = ConnectionString
Me.AdodcItem.RecordSource = SQl
Me.AdodcItem.Refresh

If Marcante > 0 And Me.AdodcItem.Recordset.RecordCount > 0 Then
    Me.cbItem.BoundText = Me.AdodcItem.Recordset("AUTONUM")
    Me.txtLocalPos.SetFocus
End If

'    Sql = " select ID_GRAVACAO as autonum, ( REPLACE(rtrim(convert(char,QTDE)),'.000','') +  '/' + rtrim(convert(char,QTDE_CAPTADA)) + ' ' + ISNULL(EMBALAGEM, ' ') + ' ' +  ISNULL(DESCR_ARMAZEM,' ') + ' ' + ISNULL(POSICAO, ' ')  ) AS DISPLAY"
'    Sql = Sql & " FROM OPERADOR..VW_INVENT_ARMAZEM"
'    Sql = Sql & " Where"
'    Sql = Sql & " lote = " & QualLote
'    Sql = Sql & " and Autonum_CS_Yard = 0 "
'    Sql = Sql & " AND QTDE>0 ORDER BY DESCR_ARMAZEM,POSICAO"
'
'    Me.AdodcItem.ConnectionString = ConnectionString
'    Me.AdodcItem.RecordSource = Sql
'    Me.AdodcItem.Refresh
'
'
'End If

Set Me.AdodcItem.Recordset.ActiveConnection = Nothing






End Sub


Private Sub CarregarArmazens()

SQl = " select Autonum as autonum, Descr AS DISPLAY "
SQl = SQl & " FROM sgipa..tb_armazens_ipa "
SQl = SQl & " WHERE "
If Col_Patio <> 1 And Col_Patio <> 7 Then
    SQl = SQl & " patio=" & Col_Patio & " and ( isnull(flag_ct,0)=0 or (flag_ct=1 and FLAG_historico=0)) "
Else
    SQl = SQl & " patio in (1,7) and ( isnull(flag_ct,0)=0 or (flag_ct=1 and FLAG_historico=0)) "
End If
SQl = SQl & " ORDER BY DESCR "

Me.AdodcArm.ConnectionString = ConnectionString
Me.AdodcArm.RecordSource = SQl
Me.AdodcArm.Refresh
Set Me.AdodcArm.Recordset.ActiveConnection = Nothing

End Sub

Private Sub CarregarMotivos()

SQl = " select Autonum as autonum, Descricao AS DISPLAY "
SQl = SQl & " FROM operador..tb_cad_motivo "
SQl = SQl & " ORDER BY descricao "

Me.AdodcMotivo.ConnectionString = ConnectionString
Me.AdodcMotivo.RecordSource = SQl
Me.AdodcMotivo.Refresh
Set Me.AdodcMotivo.Recordset.ActiveConnection = Nothing


End Sub

Private Sub Limpa()

Me.txtCliente.Text = ""
Me.txtConteiner.Text = ""
Me.txtEmbalagem.Text = ""
Me.txtEntrada.Text = ""
Me.txtLocal.Text = ""
Me.txtMarca.Text = ""
Me.txtMercadoria.Text = ""
Me.txtQtde.Text = ""
Me.txtQtdeM.Text = ""
Me.txtQtdePos.Text = ""



Me.txtIMO.Text = ""
Me.txtBL.Text = ""
Me.txtMOV.Text = ""
Me.txtLocalPos.Text = ""
Me.txtLocalPos.Tag = ""

Me.cbItem.Text = ""

Me.cbMotivoPos.Text = ""

Me.txtAutonum_Cs_Yard.Text = ""
Me.LbArmPadrao = ""


SetaNormal

End Sub
Function Validar() As Boolean

Validar = True
Me.txtCritica.ForeColor = vbGreen
Me.txtLocalPos.Tag = ""

Me.txtCritica.Tag = ""

If Me.txtLote.Text = "" Then
    Validar = False
    Me.txtCritica.Text = "INFORME O LOTE"
    Me.txtCritica.ForeColor = &H80C0FF
    Exit Function
End If

If Me.Label2(32).Tag = "C" Then
    If Me.cbItem.BoundText = "" Then
        Validar = False
        Me.txtCritica.Text = "ITEM NAO INFORMADO"
        Me.txtCritica.ForeColor = &H80C0FF
        Exit Function
    End If
End If

If Me.Label2(32).Tag <> "M" Then  'CARGAS EM RECONHECIMENTO

    If Nnull(Me.txtQtdePos.Text, 0) = 0 Then
        Validar = False
        Me.txtCritica.Text = "QTDE NAO INFORMADA"
        Me.txtCritica.ForeColor = &H80C0FF
        Exit Function
    End If
    
            
    If Val(Me.txtQtdePos.Text) > Val(Me.txtQtde.Text) Then
        Validar = False
        Me.txtCritica.Text = "QTDE INDISPONIVEL"
        Me.txtCritica.ForeColor = &H80C0FF
        Exit Function
    End If
    
End If


If Me.txtLocalPos.Text = "" Then  'nao ct
    Validar = False
    Me.txtCritica.Text = "LOCAL NAO INFORMADO"
    Me.txtCritica.ForeColor = &H80C0FF
    Exit Function
End If

SQl = "SELECT AUTONUM FROM OPERADOR..TB_YARD WHERE PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio) & " AND YARD='" & Me.txtLocalPos.Text & "'"
Dim tbL As New ADODB.Recordset
tbL.Open SQl, db, 1, 1
If tbL.EOF Then
    Validar = False
    Me.txtCritica.Text = "LOCAL INVÁLIDO"
    Me.txtCritica.ForeColor = &H80C0FF
    Exit Function
End If
tbL.Close


'If Me.txtLocalPos.text = "CAM" Then
'    Validar = False
'    Me.txtCritica.text = "CARREGAMENTO DO IPA INVALIDO"
'    Me.txtCritica.ForeColor = &H80C0FF
'    Exit Function
'End If


If Nnull(Me.cbMotivoPos.BoundText, 0) = 0 Then
    Validar = False
    Me.txtCritica.Text = "MOTIVO NAO INFORMADO"
    Me.txtCritica.ForeColor = &H80C0FF
    Exit Function
End If
       

If Me.txtMarcante.Text <> "" Then
    SQl = "SELECT AUTONUM,AUTONUM_CARGA,ISNULL(AUTONUM_CARGA_AC,0) AS AUTONUM_CARGA_AC FROM SGIPA..TB_MARCANTES WHERE AUTONUM=" & Val(Me.txtMarcante)
    Set TbM = New ADODB.Recordset
    TbM.Open SQl, db, 1, 1
    
    If TbM.EOF Then
        Validar = False
        Me.txtCritica.Text = "MARCANTE INVALIDO"
        Me.txtCritica.ForeColor = &H80C0FF
        Exit Function
    Else
    
        
        If Me.Label2(32).Tag <> "M" Then
        
            Dim AutonumCs As Long
            Dim AutonumCY As Long
        
            Dim Tipo As String
            Tipo = Left(Me.cbItem.BoundText, 1)
        
            
            If TbM!AUTONUM_CARGA_AC = 0 Then
            
                If Tipo = "C" Then
                    AutonumCs = Val(Mid(Me.cbItem.BoundText, 2))
                ElseIf Tipo = "Y" Then
                    AutonumCY = Val(Mid(Me.cbItem.BoundText, 2))
                    SQl = "SELECT AUTONUM_CS FROM SGIPA..TB_CARGA_SOLTA_YARD WHERE AUTONUM=" & AutonumCY
                            
                    Set Rst = New ADODB.Recordset
                    Rst.Open SQl, db, 1, 1
                    If Not Rst.EOF Then
                        AutonumCs = Nnull(Rst!AUTONUM_CS, 0)
                    End If
                    Rst.Close
                End If
        
                If Nnull(TbM!autonum_carga, 0) > 0 Then
                    If Nnull(TbM!autonum_carga, 0) <> AutonumCs Then
                        Validar = False
                        Me.txtCritica.Text = "MARCANTE VINC.A OUTRA CARGA"
                        Me.txtCritica.ForeColor = &H80C0FF
                        Exit Function
                    End If
                End If
            End If
        
        End If
    
    End If
    TbM.Close
End If


If Nnull(Me.txtQtdeM.Text, 0) > 0 Then
    If Nnull(Me.txtQtdePos.Text, 0) <> Nnull(Me.txtQtdeM.Text, 0) Then
        Validar = False
        Me.txtCritica.Text = "DIV. COM A QTDE MARCANTE"
        Me.txtCritica.ForeColor = &H80C0FF
        Exit Function
    End If
End If


Dim RetAnvisa As String
SQl = "SELECT [dbo].[FN_VALIDA_ANVISA_ARMAZEM]()"
RetAnvisa = Nnull(db.Execute(SQl).Fields(0), 1)
If RetAnvisa <> "OK" Then
    Validar = False
    Me.txtCritica.Text = RetAnvisa
    Me.txtCritica.ForeColor = &H80C0FF
    Exit Function
End If



End Function

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub txtLocalPos_LostFocus()
txtLocalPos = UCase(txtLocalPos)
End Sub

Private Sub txtMarcante_Change()

Me.Label2(32).ForeColor = &HBA9F8D
Me.Label2(32).Tag = "C"


If Len(Me.txtMarcante) = 9 Then

    Me.LbArmPadrao = ""
    Set tb1 = New ADODB.Recordset
    SQl = "SELECT * FROM sgipa.dbo.fnInvent_Sistemas_Arm_Col_PATIO(" & Val(Me.txtMarcante) & ")"
    tb1.Open SQl, db, 1, 1
    If Not tb1.EOF Then
    
    
     If tb1!Patio <> 0 Then
        
            If tb1!Patio <> 1 And tb1!Patio <> 7 Then
        
                If Col_Patio <> tb1!Patio Then
                    
                    Me.Label2(32).ForeColor = vbRed
                    Me.Label2(32).Caption = "PATIO DIVERGENTE AO USUARIO LOGADO"
                    Exit Sub
                           
                    
                End If
            
            Else
            
            
                Dim Col_Patio2 As Integer
                If Col_Patio = 1 Then Col_Patio2 = 7
                If Col_Patio = 7 Then Col_Patio2 = 1
                
                
                If (Col_Patio <> tb1!Patio) And (Col_Patio2 <> tb1!Patio) Then
                    
                    Me.Label2(32).ForeColor = vbRed
                    Me.Label2(32).Caption = "PATIO DIVERGENTE AO USUARIO LOGADO"
                    Exit Sub
                           
                    
                End If
                
                
            End If
        
        End If
       
       
       If tb1!AUTONUM_CEXP = 0 Then
       
            If tb1!FLAG_RECONHECIMENTO = 0 Then
        
                If tb1!autonum_carga > 0 Then
                
                    Me.Label2(32).ForeColor = &HBA9F8D
                    Me.Label2(32).Tag = "C"
                
                    Set Tb2 = New ADODB.Recordset
                    Tb2.Open "SELECT BL,ITEM FROM SGIPA..TB_CARGA_SOLTA WHERE AUTONUM=" & tb1!autonum_carga, db, 1, 1
                    If Not Tb2.EOF Then
                        Call Carrega_Dados_Lote(Nnull(Tb2!BL, 0), Nnull(tb1!Autonum_Cs_Yard, 0))
                        Me.txtLote.Text = Nnull(Tb2!BL, 0)
                         'Me.cbItem.SetFocus
                         
                        Me.txtQtde.Text = Nnull(tb1!VOLUMES, 0)
                        Me.txtQtdeM.Text = Nnull(tb1!VOLUMES, 0)
                        Me.txtQtdePos.Text = Nnull(tb1!VOLUMES, 0)
                        Me.txtAutonum_Cs_Yard.Text = Nnull(tb1!Autonum_Cs_Yard, 0)
                        Me.txtLocalPos.SetFocus
                         
                        Me.LbArmPadrao.Caption = Carrega_Arm_Padrao(Me.txtLote.Text)
                         
                    Else
                        Limpa
                    End If
                 
              
             
                Else
             
                    Me.Label2(32).ForeColor = vbRed
                    Me.Label2(32).Tag = "Z"
                
                    Set Tb2 = New ADODB.Recordset
                    Tb2.Open "SELECT LOTE,BL,EMBALAGEM,IMPORTADOR,NAVIO,QTDE,AUTONUMCS,MERCADORIA,DATA_ENTRADA,TIPO_DOC,DESCR_ARMAZEM,POSICAO,ID_GRAVACAO FROM OPERADOR..VW_INVENT_ARMAZEM WHERE MARCANTE= " & Val(Me.txtMarcante), db, 1, 1
                 
                    If Not Tb2.EOF Then
                     
                        Me.txtLote.Text = Nnull(Tb2!Lote, 0)
                        Me.txtLote.Tag = Nnull(Tb2!id_Gravacao, 1)
                                        
                        Me.txtMercadoria.Text = Nnull(Tb2!MERCADORIA, 1)
                        Me.txtMarca.Text = ""
                        Me.txtEntrada.Text = Nnull(Tb2!Data_Entrada, 1)
                        Me.txtConteiner.Text = ""
                        Me.txtCliente.Text = Nnull(Tb2!Importador, 1)
                        Me.txtEmbalagem.Text = Nnull(Tb2!embalagem, 1)
                        
                        
                        
                        Me.txtMOV.Text = ""
                        
                        Me.txtIMO.Text = ""
                        
                        Me.txtBL.Text = Nnull(Tb2!BL, 1)
                        Me.txtLocal.Text = Nnull(Tb2!DESCR_ARMAZEM, 1) & " " & Nnull(Tb2!Posicao, 1)
                        If Trim(Me.txtLocal.Text) = "" Or Trim(Me.txtLocal.Text) = "-" Then Me.cbMotivoPos.BoundText = 14
                                     
                        Me.LbArmPadrao.Caption = Carrega_Arm_Padrao(Me.txtLote.Text)
                                     
                    Else
                 
                        Limpa
                 
                    End If
                 
                    Me.txtQtde.Text = Nnull(tb1!VOLUMES, 0)
                    Me.txtQtdeM.Text = Nnull(tb1!VOLUMES, 0)
                    Me.txtQtdePos.Text = Nnull(tb1!VOLUMES, 0)
                    Me.txtAutonum_Cs_Yard.Text = Nnull(tb1!Autonum_Cs_Yard_AC, 0)
                    Me.txtLocalPos.SetFocus
             
                End If
         
            Else 'CARGA EM RECONHECIMENTO
         
                Me.Label2(32).ForeColor = vbRed
                Me.Label2(32).Tag = "M"
                Me.txtConteiner.Text = tb1!ID_CONTEINER_IMPRESSO
                Me.txtLote.Text = tb1!LOTE_IMPRESSO
             
                Dim ArmD As String
                ArmD = Nnull(db.Execute(" SELECT DESCR FROM SGIPA..TB_ARMAZENS_IPA WHERE AUTONUM=" & Nnull(tb1!ARMAZEM_REC, 0))(0), 1)
                Me.txtLocal.Text = ArmD & " " & Nnull(tb1!YARD_REC, 1)
             
             
                Me.txtQtde.Text = "0"
                Me.txtQtdeM.Text = "0"
                Me.txtQtdePos.Text = "0"
                Me.txtAutonum_Cs_Yard.Text = "0"
                Me.txtLocalPos.SetFocus
         
            End If
    
        End If
         
        
    Else
        Limpa
    End If
Else
    Limpa
End If
End Sub


Private Sub SetaNormal()
Me.txtLocalPos.Visible = True



End Sub

Private Sub SetaCT()
Me.txtLocalPos.Visible = False

End Sub

Private Sub txtPlacaDta_Change()

If Len(txtPlacaDta.Text) = 7 Then
    Dim wPlaca As String
    wPlaca = Mid(txtPlacaDta, 1, 3) & "-" & Mid(txtPlacaDta, 4, 4)
    wPlaca = UCase(wPlaca)

    SQl = "select autonum_cs from sgipa..TB_GOS_AMR_CHRONOS WHERE FUNCAO_GATE=20 and patio=" & Col_Patio & " and gos_vehicle_number='" & wPlaca & "'"
    Dim Rs As New ADODB.Recordset
    PRSet Rs, SQl
    If Not Rs.EOF Then
        SQl = "SELECT autonum FROM sgipa..TB_MARCANTES WHERE autonum_carga=" & Rs!AUTONUM_CS
        PRSet Rs, SQl
        txtMarcante.Text = Format(Rs!autonum, "000000000")
        txtLocalPos.Text = "DTA"
        cbMotivoPos = "REPESAGEM"
        
        
        If txtLote.Text <> "" Then
            SQl = "select isnull(flag_aviso_temperatura_coletor,0) from sgipa..tb_bl bl"
            SQl = SQl & " inner join operador..tb_patios p on bl.patio=p.autonum"
            SQl = SQl & " where bl.autonum=" & txtLote.Text
            SQl = SQl & " AND ISNULL(BL.SUB_CLASSE,0)<>0"
            
            If Nnull(db.Execute(SQl).Fields(0), 0) = 1 Then
                MsgBox ("CONFIRA CONDIÇÕES DE TEMPERATURA E UMIDADE")
                Exit Sub
            End If
        End If
        
        Me.Frame1.Visible = False
        Salvar (wPlaca)
        If Me.txtCritica.ForeColor = vbGreen Then Me.txtMarcante.Text = ""
    Else
        MsgBox ("sem processo")
    End If

End If



End Sub


