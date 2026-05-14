VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form Inventario_CS 
   Appearance      =   0  'Flat
   BackColor       =   &H00BA9F8D&
   BorderStyle     =   0  'None
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3585
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
   ScaleHeight     =   4425
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   2850
      TabIndex        =   57
      Top             =   4170
      Width           =   675
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
      Left            =   2040
      TabIndex        =   56
      Top             =   4170
      Width           =   765
   End
   Begin VB.CommandButton cmdLimpa 
      BackColor       =   &H008080FF&
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
      Left            =   1110
      MaskColor       =   &H008080FF&
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   4170
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton cmdHistorico 
      Caption         =   "[F1]HISTORICO"
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
      TabIndex        =   54
      Top             =   4170
      Width           =   1005
   End
   Begin VB.Frame frmHist 
      Appearance      =   0  'Flat
      BackColor       =   &H00BA9F8D&
      Caption         =   "Histórico da Movimentação"
      ForeColor       =   &H00FFFFFF&
      Height          =   4185
      Left            =   0
      TabIndex        =   51
      Top             =   -30
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
         Bindings        =   "InventarioCS.frx":0000
         Height          =   3225
         Left            =   120
         TabIndex        =   52
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
   Begin VB.ComboBox cbOC 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      ItemData        =   "InventarioCS.frx":0015
      Left            =   2775
      List            =   "InventarioCS.frx":0025
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   2775
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CheckBox Check4 
      Appearance      =   0  'Flat
      BackColor       =   &H00BA9F8D&
      Caption         =   "LD"
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
      Left            =   3060
      TabIndex        =   49
      Top             =   3030
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CheckBox Check3 
      Appearance      =   0  'Flat
      BackColor       =   &H00BA9F8D&
      Caption         =   "LE"
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
      Left            =   2520
      TabIndex        =   48
      Top             =   3030
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H00BA9F8D&
      Caption         =   "FUNDO"
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
      Left            =   1830
      TabIndex        =   47
      Top             =   3030
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00BA9F8D&
      Caption         =   "FRENTE"
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
      Left            =   1140
      TabIndex        =   46
      Top             =   3030
      Visible         =   0   'False
      Width           =   705
   End
   Begin MSDataListLib.DataCombo cbArm 
      Bindings        =   "InventarioCS.frx":003E
      Height          =   270
      Left            =   1530
      TabIndex        =   5
      Top             =   2760
      Width           =   1245
      _ExtentX        =   2196
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
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   3210
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
      Height          =   225
      Left            =   3150
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   180
      Width           =   405
   End
   Begin VB.TextBox txtPrateleira 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      IMEMode         =   3  'DISABLE
      Left            =   1140
      MaxLength       =   12
      TabIndex        =   3
      Top             =   2550
      Width           =   2415
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
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   1140
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
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   3510
      Width           =   3585
   End
   Begin VB.TextBox txtCargaNoLocal 
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
      Height          =   255
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3030
      Width           =   2415
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
      Height          =   225
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2340
      Width           =   2415
   End
   Begin VB.TextBox txtCANAL 
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
      Left            =   9870
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1065
   End
   Begin VB.TextBox txtDOC 
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
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1365
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
      Height          =   225
      Left            =   2910
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2130
      Width           =   645
   End
   Begin VB.TextBox txtNVOCC 
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
      Left            =   8790
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3600
      Width           =   2415
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
      Height          =   225
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1050
      Width           =   2415
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
      Height          =   225
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   840
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
      Height          =   270
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2760
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
      Height          =   270
      Left            =   2730
      TabIndex        =   6
      Top             =   2760
      Width           =   825
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
      Height          =   225
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1500
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
      Height          =   225
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2130
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
      Height          =   225
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1920
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
      Height          =   255
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1260
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
      Height          =   225
      Left            =   1950
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   840
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
      Height          =   225
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1710
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
      Height          =   225
      Left            =   1950
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   390
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
      Height          =   225
      Left            =   1140
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   1
      Top             =   390
      Width           =   825
   End
   Begin MSDataListLib.DataCombo cbItem 
      Bindings        =   "InventarioCS.frx":0055
      Height          =   270
      Left            =   1140
      TabIndex        =   2
      Top             =   600
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
      Bindings        =   "InventarioCS.frx":006D
      Height          =   270
      Left            =   1140
      TabIndex        =   7
      Top             =   3270
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
   Begin VB.Label LBLEXPORTACAO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EXP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   240
      Left            =   90
      TabIndex        =   53
      Top             =   390
      Width           =   390
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
      TabIndex        =   43
      Top             =   210
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ETQ  PRATELEIRA"
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
      Index           =   16
      Left            =   -30
      TabIndex        =   42
      Top             =   2580
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
      TabIndex        =   40
      Top             =   3300
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CARGA LOCAL "
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
      Index           =   10
      Left            =   0
      TabIndex        =   38
      Top             =   3060
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
      Left            =   -30
      TabIndex        =   36
      Top             =   2370
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DOC. / CANAL "
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
      Index           =   13
      Left            =   7380
      TabIndex        =   33
      Top             =   3150
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
      TabIndex        =   31
      Top             =   2190
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NVOCC "
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
      Index           =   9
      Left            =   7620
      TabIndex        =   29
      Top             =   3600
      Width           =   1155
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
      TabIndex        =   26
      Top             =   630
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
      TabIndex        =   25
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
      TabIndex        =   23
      Top             =   1500
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   285
      Index           =   8
      Left            =   -30
      TabIndex        =   16
      Top             =   2820
      Width           =   1155
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
      TabIndex        =   15
      Top             =   2160
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
      TabIndex        =   14
      Top             =   1950
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
      TabIndex        =   13
      Top             =   1260
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
      TabIndex        =   12
      Top             =   1710
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LOCAL "
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
      Index           =   3
      Left            =   -30
      TabIndex        =   11
      Top             =   1050
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
      Left            =   -30
      TabIndex        =   10
      Top             =   840
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
      TabIndex        =   9
      Top             =   420
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "INVENTÁRIO DE CARGA SOLTA"
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
      TabIndex        =   8
      Top             =   0
      Width           =   3570
   End
End
Attribute VB_Name = "Inventario_CS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HorarioNumero As Date
Dim HorarioEnter As Date


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

Me.txtQtde.text = ""
Me.txtEmbalagem.text = ""
Me.TXTLOCAL.text = ""

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
    Me.TXTLOCAL = Nnull(tb1!local, 1)
    If Trim(Me.TXTLOCAL.text) = "" Or Trim(Me.TXTLOCAL.text) = "-" Then
        Me.cbMotivoPos.BoundText = 14
    Else
        Me.cbMotivoPos.BoundText = 1
    End If
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
Dim X As Integer
X = MsgBox("Deseja limpar os dados", 36 + vbDefaultButton2)
If X = vbYes Then
    Limpa
    Me.TXTMARCANTE.text = ""
    'txtMarcante.Enabled = True
    'txtMarcante.Tag = 0

    'HorarioNumero = Now
    'HorarioEnter = Now
End If
End Sub

Private Sub cmdSalva_Click()
Salvar
If Me.txtCritica.ForeColor = vbGreen Then Me.TXTMARCANTE.text = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Me.frmHist.Visible = False Then
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
    If KeyCode = vbKeyF2 Then
       
        Limpa
        Me.TXTMARCANTE.text = ""
    End If
    
    If KeyCode = vbKeyF5 Then
        Salvar
        If Me.txtCritica.ForeColor = vbGreen Then Me.TXTMARCANTE.text = ""
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
SQl = SQl & " WHERE A.MARCANTE=" & Val(Nnull(Me.TXTMARCANTE, 0))
SQl = SQl & " ORDER BY A.DT_MOV DESC "

Me.Adodc1.ConnectionString = ConnectionString
Me.Adodc1.RecordSource = SQl
Me.Adodc1.Refresh
Set Me.Adodc1.Recordset.ActiveConnection = Nothing

Me.Grid1.FormatString = " |^Armazem|^Local       |^Data               |<Usuario             "


End Sub
Private Sub Salvar()
 
 If Validar() Then

    Dim SQl As String
    Dim SQL2 As String
    
    Dim AutonumCs As Long
    Dim AutonumCY As Long

    Dim Tipo As String
    
    Dim FlagChronos As Boolean
    FlagChronos = True
    
    
    If Me.Label2(32).Tag <> "C" And Me.Label2(32).Tag <> "M" Then
        FlagChronos = False
    End If
    
    If Me.LBLEXPORTACAO.Visible = False Then
    
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
                
            Tipo = Left(Me.TXTLOTE.Tag, 1)
        
            If Tipo = "V" Then
                AutonumCs = Val(Mid(Me.TXTLOTE.Tag, 2))
            ElseIf Tipo = "W" Then
                AutonumCY = Val(Mid(Me.TXTLOTE.Tag, 2))
                SQl = "SELECT AUTONUM_CS FROM SGIPA..TB_CARGA_SOLTA_YARD_AC WHERE AUTONUM=" & AutonumCY
                        
                Set Rst = New ADODB.Recordset
                Rst.Open SQl, db, 1, 1
                If Not Rst.EOF Then
                    AutonumCs = Nnull(Rst!AUTONUM_CS, 0)
                End If
                Rst.Close
            End If
                
                
        End If
    
    Else
        
        Tipo = Left(Me.TXTLOTE.Tag, 1)
        If Tipo = "R" Then
            AutonumCs = Val(Mid(Me.TXTLOTE.Tag, 2))
        ElseIf Tipo = "T" Then
            AutonumCY = Val(Mid(Me.TXTLOTE.Tag, 2))
            SQl = "SELECT AUTONUM_CEXP FROM SGIPA..TB_CARGA_SOLTA_YARD_CEXP WHERE AUTONUM=" & AutonumCY
                    
            Set Rst = New ADODB.Recordset
            Rst.Open SQl, db, 1, 1
            If Not Rst.EOF Then
                AutonumCs = Nnull(Rst!AUTONUM_CEXP, 0)
            End If
            Rst.Close
        End If
    
    End If
        
    
    If Me.txtLocalPos.text <> "" Then
        
        SQl = "SELECT AUTONUM FROM "
        If Me.LBLEXPORTACAO.Visible = False Then
            If FlagChronos Then
                SQl = SQl & " SGIPA..TB_CARGA_SOLTA_YARD "
            Else
                SQl = SQl & " SGIPA..TB_CARGA_SOLTA_YARD_AC "
            End If
            SQl = SQl & " WHERE AUTONUM_CS= " & AutonumCs
        Else
            SQl = SQl & " SGIPA..TB_CARGA_SOLTA_YARD_CEXP "
            SQl = SQl & " WHERE AUTONUM_CEXP= " & AutonumCs
        End If
        
        If Me.txtLocalPos.Tag = "" Then
            SQl = SQl & " AND ARMAZEM=" & Me.cbArm.BoundText & " AND YARD='" & Me.txtLocalPos.text & "'"
        Else
            SQl = SQl & " AND ARMAZEM=" & Me.cbArm.BoundText & " AND YARD='" & Me.txtLocalPos.Tag & "'"
        End If
            
    Else
    
        SQl = "SELECT AUTONUM FROM "
        If Me.LBLEXPORTACAO.Visible = False Then
            If FlagChronos Then
                SQl = SQl & " SGIPA..TB_CARGA_SOLTA_YARD "
            Else
                SQl = SQl & " SGIPA..TB_CARGA_SOLTA_YARD_AC "
            End If
            SQl = SQl & " WHERE AUTONUM_CS= " & AutonumCs
        Else
            SQl = SQl & " SGIPA..TB_CARGA_SOLTA_YARD_CEXP "
            SQl = SQl & " WHERE AUTONUM_CEXP= " & AutonumCs
        End If
        SQl = SQl & " AND ARMAZEM= " & Me.cbArm.BoundText
        SQl = SQl & " AND YARD IS NULL "
        
    End If

    Set Rst5 = New ADODB.Recordset
    Rst5.Open SQl, db, 1, 1
    
    If Rst5.EOF Then
        If Me.txtLocalPos.text <> "" Then
            
            SQl = "INSERT INTO "
            If Me.LBLEXPORTACAO.Visible = False Then
                If FlagChronos Then
                    SQl = SQl & "SGIPA..TB_CARGA_SOLTA_YARD"
                Else
                    SQl = SQl & "SGIPA..TB_CARGA_SOLTA_YARD_AC"
                End If
                SQl = SQl & "(AUTONUM_CS,ARMAZEM,YARD,ORIGEM,QUANTIDADE,MOTIVO,AUDIT_94,FL_FRENTE,FL_FUNDO,FL_LE,FL_LD"
            Else
                SQl = SQl & "SGIPA..TB_CARGA_SOLTA_YARD_CEXP"
                SQl = SQl & "(AUTONUM_CEXP,ARMAZEM,YARD,ORIGEM,QUANTIDADE,MOTIVO,AUDIT_94,FL_FRENTE,FL_FUNDO,FL_LE,FL_LD"
            End If
            
            SQl = SQl & ") VALUES "
            SQl = SQl & "(" & AutonumCs & "," & Me.cbArm.BoundText & ",'" & IIf(Me.txtLocalPos.Tag <> "", Me.txtLocalPos.Tag, Me.txtLocalPos.text) & "','I'," & Val(Me.txtQtdePos.text) & "," & Me.cbMotivoPos.BoundText
            SQl = SQl & ",0,0,0,0,0"
            SQl = SQl & ")"
            
        Else
            
            SQl = "INSERT INTO "
            If Me.LBLEXPORTACAO.Visible = False Then
                If FlagChronos Then
                    SQl = SQl & "SGIPA..TB_CARGA_SOLTA_YARD "
                Else
                    SQl = SQl & "SGIPA..TB_CARGA_SOLTA_YARD_AC "
                End If
                SQl = SQl & " (AUTONUM_CS,ARMAZEM,ORIGEM,QUANTIDADE,MOTIVO"
            Else
                   SQl = SQl & "SGIPA..TB_CARGA_SOLTA_YARD_CEXP "
            End If
            If Me.Check1.Visible = True Then
                SQl = SQl & ",FL_FRENTE,FL_FUNDO,FL_LE,FL_LD"
            End If
            SQl = SQl & ") VALUES "
            
            SQl = SQl & "(" & AutonumCs & "," & Me.cbArm.BoundText & ",'I'," & Val(Me.txtQtdePos.text) & "," & Me.cbMotivoPos.BoundText & ""
            
            If Me.Check1.Visible = True Then
                SQl = SQl & "," & Check1.Value & "," & Check2.Value & "," & Check3.Value & "," & Check4.Value
            End If
            SQl = SQl & ")"
                
        End If
        
    Else
         
         If Me.txtLocalPos.text <> "" Then
            SQl = "INSERT INTO "
            
            If Me.LBLEXPORTACAO.Visible = False Then
                If FlagChronos Then
                    SQl = SQl & "SGIPA..TB_CARGA_SOLTA_YARD "
                Else
                    SQl = SQl & "SGIPA..TB_CARGA_SOLTA_YARD_AC "
                End If
                SQl = SQl & "(AUTONUM_CS,ARMAZEM,YARD,ORIGEM,QUANTIDADE,MOTIVO,AUDIT_94,FL_FRENTE,FL_FUNDO,FL_LE,FL_LD) VALUES "
            Else
                SQl = SQl & "SGIPA..TB_CARGA_SOLTA_YARD_CEXP "
                SQl = SQl & "(AUTONUM_CEXP,ARMAZEM,YARD,ORIGEM,QUANTIDADE,MOTIVO,AUDIT_94,FL_FRENTE,FL_FUNDO,FL_LE,FL_LD) VALUES "
            End If
            SQl = SQl & "(" & AutonumCs & "," & Me.cbArm.BoundText & ",'" & IIf(Me.txtLocalPos.Tag <> "", Me.txtLocalPos.Tag, Me.txtLocalPos.text) & "','I'," & Val(Me.txtQtdePos.text) & "," & Me.cbMotivoPos.BoundText & ",0,0,0,0,0)"
            
        Else
            SQl = "INSERT INTO "
            If Me.LBLEXPORTACAO.Visible = False Then
                If FlagChronos Then
                    SQl = SQl & "SGIPA..TB_CARGA_SOLTA_YARD "
                Else
                    SQl = SQl & "SGIPA..TB_CARGA_SOLTA_YARD_AC "
                End If
                SQl = SQl & "(AUTONUM_CS,ARMAZEM,ORIGEM,QUANTIDADE,MOTIVO"
            Else
                SQl = SQl & "SGIPA..TB_CARGA_SOLTA_YARD_CEXP "
                SQl = SQl & "(AUTONUM_CEXP,ARMAZEM,ORIGEM,QUANTIDADE,MOTIVO"
            End If
            If Me.Check1.Visible = True Then
                SQl = SQl & ",FL_FRENTE,FL_FUNDO,FL_LE,FL_LD"
            End If
            SQl = SQl & ") VALUES "
            SQl = SQl & "(" & AutonumCs & "," & Me.cbArm.BoundText & ",'I'," & Val(Me.txtQtdePos.text) & "," & Me.cbMotivoPos.BoundText & ""
            If Me.Check1.Visible = True Then
                SQl = SQl & "," & Check1.Value & "," & Check2.Value & "," & Check3.Value & "," & Check4.Value
            End If
            SQl = SQl & ")"
                
        End If
         
        'Nao soma mais, pois quantidade associada ao marcante
        'Sql = "UPDATE SGIPA..TB_CARGA_SOLTA_YARD SET QUANTIDADE=QUANTIDADE +  " & Val(Me.txtQtdePos.Text)
        'Sql = Sql & " ,MOTIVO=" & Me.cbMotivoPos.BoundText & " "
        'Sql = Sql & " WHERE AUTONUM=" & Rst5!AUTONUM
            
    End If
    
    Rst5.Close
    
    SQL2 = ""
    
    Dim SQL3 As String
    Dim SQL4 As String
        
    If Tipo = "Y" Then
        SQL2 = "UPDATE SGIPA..TB_CARGA_SOLTA_YARD SET QUANTIDADE=QUANTIDADE - " & Val(Me.txtQtdePos.text)
        If Me.Check1.Visible = True Then
            SQL2 = SQL2 & ",FL_FRENTE=" & Check1.Value & ",FL_FUNDO=" & Check2.Value & ",FL_LE=" & Check3.Value & ",FL_LD=" & Check4.Value
        End If
        SQL2 = SQL2 & " WHERE AUTONUM=" & AutonumCY
    End If
    
    If Tipo = "W" Then
        SQL2 = "UPDATE SGIPA..TB_CARGA_SOLTA_YARD_AC SET QUANTIDADE=QUANTIDADE - " & Val(Me.txtQtdePos.text)
        If Me.Check1.Visible = True Then
            SQL2 = SQL2 & ",FL_FRENTE=" & Check1.Value & ",FL_FUNDO=" & Check2.Value & ",FL_LE=" & Check3.Value & ",FL_LD=" & Check4.Value
        End If
        SQL2 = SQL2 & " WHERE AUTONUM=" & AutonumCY
    End If
    
    If Tipo = "T" Then
        SQL2 = "UPDATE SGIPA..TB_CARGA_SOLTA_YARD_CEXP SET QUANTIDADE=QUANTIDADE - " & Val(Me.txtQtdePos.text)
        If Me.Check1.Visible = True Then
            SQL2 = SQL2 & ",FL_FRENTE=" & Check1.Value & ",FL_FUNDO=" & Check2.Value & ",FL_LE=" & Check3.Value & ",FL_LD=" & Check4.Value
        End If
        SQL2 = SQL2 & " WHERE AUTONUM=" & AutonumCY
    End If
    
    
    If Me.LBLEXPORTACAO.Visible = True Then
        If Me.txtLocalPos.text = "CAM" Then
            SQL3 = "UPDATE SGIPA..TB_CARGA_EXPORTACAO SET QUANTIDADE_SAIDA=ISNULL(QUANTIDADE_SAIDA,0) + " & Val(Me.txtQtdePos.text) & " WHERE AUTONUM=" & AutonumCs
            SQL4 = "UPDATE SGIPA..TB_CARGA_EXPORTACAO SET DT_SAIDA=GETDATE() WHERE ISNULL(QUANTIDADE_SAIDA,0) >=QUANTIDADE AND AUTONUM=" & AutonumCs
        ElseIf Right(Me.TXTLOCAL.text, 3) = "CAM" And Me.txtLocalPos.text <> "CAM" Then
            SQL3 = "UPDATE SGIPA..TB_CARGA_EXPORTACAO SET QUANTIDADE_SAIDA=ISNULL(QUANTIDADE_SAIDA,0) - " & Val(Me.txtQtdePos.text) & " WHERE AUTONUM=" & AutonumCs
            SQL4 = "UPDATE SGIPA..TB_CARGA_EXPORTACAO SET DT_SAIDA=NULL WHERE ISNULL(QUANTIDADE_SAIDA,0)<QUANTIDADE AND AUTONUM=" & AutonumCs
            
        End If
    End If
    
    
    'db.BeginTrans
        
    'db.Execute Sql
        
    Dim Autonum_Cs_Yard As Long
    
    If Me.LBLEXPORTACAO.Visible = False Then
        If FlagChronos Then
            ''Autonum_Cs_Yard = db.Execute("SELECT IDENT_CURRENT('TB_CARGA_SOLTA_YARD')")(0)
            'Autonum_Cs_Yard = db.Execute("SELECT MAX(AUTONUM) FROM TB_CARGA_SOLTA_YARD WHERE AUTONUM_CS=" & AutonumCs)(0)
        Else
            'Autonum_Cs_Yard = db.Execute("SELECT IDENT_CURRENT('TB_CARGA_SOLTA_YARD_AC')")(0)
        End If
    Else
        'Autonum_Cs_Yard = db.Execute("SELECT IDENT_CURRENT('TB_CARGA_SOLTA_YARD_CEXP')")(0)
    End If
        
        
    'If SQL2 <> "" Then db.Execute SQL2
    'If SQL3 <> "" Then db.Execute SQL3
    'If SQL4 <> "" Then db.Execute SQL4
        
    If Me.TXTMARCANTE <> "" Then
        Dim SqlM As String
        SqlM = "UPDATE SGIPA..TB_MARCANTES SET "
        If Me.LBLEXPORTACAO.Visible = False Then
            If FlagChronos Then
                SqlM = SqlM & " AUTONUM_CARGA=" & AutonumCs & ""
                SqlM = SqlM & " ,AUTONUM_CS_YARD=" & Autonum_Cs_Yard
            Else
                SqlM = SqlM & " AUTONUM_CARGA_AC=" & AutonumCs & ""
                SqlM = SqlM & " ,AUTONUM_CS_YARD_AC=" & Autonum_Cs_Yard
            End If
        Else
            SqlM = SqlM & " AUTONUM_CEXP=" & AutonumCs & ""
            SqlM = SqlM & " ,AUTONUM_CS_YARD_CEXP=" & Autonum_Cs_Yard
        End If
        SqlM = SqlM & ",VOLUMES=" & Val(Me.txtQtdePos.text) & " ,DT_ASSOCIACAO=GETDATE()"
        
        SqlM = SqlM & " WHERE AUTONUM= " & Val(Me.TXTMARCANTE)
        'db.Execute SqlM
    End If
    
    
    If Me.Label2(32).Tag = "M" Then
        SqlM = "UPDATE SGIPA..TB_MARCANTES SET "
        SqlM = SqlM & " DT_ASSOCIACAO=GETDATE()"
        SqlM = SqlM & " ,ARMAZEM_REC=" & Me.cbArm.BoundText
        SqlM = SqlM & " ,YARD_REC='" & Me.txtLocalPos.text & "'"
        
        If FlagChronos Then
            SqlM = SqlM & " ,AUTONUM_CARGA=" & AutonumCs & ""
            SqlM = SqlM & " ,AUTONUM_CS_YARD=" & Autonum_Cs_Yard
        Else
            SqlM = SqlM & " ,AUTONUM_CARGA_AC=" & AutonumCs & ""
            SqlM = SqlM & " ,AUTONUM_CS_YARD_AC=" & Autonum_Cs_Yard
        End If
        SqlM = SqlM & ",VOLUMES=" & Val(Me.txtQtdePos.text)
        SqlM = SqlM & " WHERE AUTONUM= " & Val(Me.TXTMARCANTE)
        'db.Execute SqlM
    End If

        
    If Me.cbOC.Visible = True Then
        'db.Execute "UPDATE SGIPA..TB_ARMAZENS_IPA SET PERC_OCUPACAO_CT='" & Me.cbOC.text & "' WHERE AUTONUM=" & Me.cbArm.BoundText
    End If
        
        
'    Sql = "INSERT INTO SGIPA..TB_HIST_SHIFTING_CS(MARCANTE,ARMAZEM,YARD,DT_MOV,USUARIO) VALUES ("
'    Sql = Sql & Val(Me.txtMarcante.text) & ","
'    Sql = Sql & Me.cbArm.BoundText & ","
'    Sql = Sql & "'" & Me.txtLocalPos.text & "',"
'    Sql = Sql & "GETDATE()," & Cod_Usuario & ")"
    'db.Execute Sql
    
    
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = ConnectionString
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "[Sgipa].[dbo].[pcMovimentaCS2]"
    cmd.Parameters.Append cmd.CreateParameter("Tag", adVarChar, adParamInput, 1, Me.Label2(32).Tag)
    cmd.Parameters.Append cmd.CreateParameter("Tipo", adVarChar, adParamInput, 1, Tipo)
    cmd.Parameters.Append cmd.CreateParameter("FlagChronos", adInteger, adParamInput, 1, IIf(FlagChronos, 1, 0))
    cmd.Parameters.Append cmd.CreateParameter("AutonumCs", adInteger, adParamInput, 10, AutonumCs)
    cmd.Parameters.Append cmd.CreateParameter("Patio", adInteger, adParamInput, 2, IIf(Col_Patio = 7, 1, Col_Patio))
    cmd.Parameters.Append cmd.CreateParameter("Yard", adVarChar, adParamInput, 10, Me.txtLocalPos.text)
    cmd.Parameters.Append cmd.CreateParameter("Quantidade", adInteger, adParamInput, 8, Val(Me.txtQtdePos.text))
    cmd.Parameters.Append cmd.CreateParameter("Motivo", adInteger, adParamInput, 8, Me.cbMotivoPos.BoundText)
    cmd.Parameters.Append cmd.CreateParameter("Marcante", adInteger, adParamInput, 8, Val(Me.TXTMARCANTE.text))
    cmd.Parameters.Append cmd.CreateParameter("Usuario", adInteger, adParamInput, 8, Cod_Usuario)
    
    cmd.Parameters.Append cmd.CreateParameter("FlagExp", adInteger, adParamInput, 8, IIf(Me.LBLEXPORTACAO.Visible = True, 1, 0))
    cmd.Parameters.Append cmd.CreateParameter("Armazem", adInteger, adParamInput, 8, Nnull(Me.cbArm.BoundText, 0))
    cmd.Parameters.Append cmd.CreateParameter("Fl_Frente", adInteger, adParamInput, 8, Me.Check1.Value)
    cmd.Parameters.Append cmd.CreateParameter("Fl_Fundo", adInteger, adParamInput, 8, Me.Check2.Value)
    cmd.Parameters.Append cmd.CreateParameter("Fl_LE", adInteger, adParamInput, 8, Me.Check3.Value)
    cmd.Parameters.Append cmd.CreateParameter("Fl_LD", adInteger, adParamInput, 8, Me.Check4.Value)
    cmd.Parameters.Append cmd.CreateParameter("Autonumcy", adInteger, adParamInput, 8, AutonumCY)
    cmd.Parameters.Append cmd.CreateParameter("YardAtual", adVarChar, adParamInput, 30, Me.TXTLOCAL.text)
    cmd.Parameters.Append cmd.CreateParameter("FlagCT", adInteger, adParamInput, 8, IIf(cbOC.Visible = True, 1, 0))
    cmd.Parameters.Append cmd.CreateParameter("PercentCT", adInteger, adParamInput, 8, IIf(cbOC.Visible = True, Val(Me.cbOC.text), 0))
    cmd.Execute
     
    Set cmd = Nothing
    
        
    If FlagChronos Then
        If Trim(Me.txtIMO.text) <> "" Then
            Dim Segreg$
            Segreg = Verifica_Regras_Seg_Imo_DeltaCS(Me.TXTMARCANTE, Me.txtLocalPos.text, Me.cbArm.BoundText)
            If Segreg <> "" Then
                
                Me.txtCritica = "SEGREGACAO IMO !!!! " & Segreg
                Me.txtCritica.ForeColor = &H80C0FF
                
                SQl = "UPDATE SGIPA..TB_CARGA_SOLTA_YARD SET YARD='' WHERE AUTONUM IN (SELECT MAX(AUTONUM) FROM TB_CARGA_SOLTA_YARD WHERE AUTONUM_CS=" & AutonumCs & ")"
                db.Execute (SQl)
                
                'Frmmensagem.Show vbModal
                Exit Sub
            End If
        End If
    End If
    
    'db.CommitTrans
    
    Me.txtCritica = "CARGA POSICIONADA COM SUCESSO"
    Me.txtCritica.ForeColor = vbGreen
      
    Limpa
    
    Me.TXTLOTE.text = ""
    Me.TXTLOTE.Tag = ""
    Me.TXTMARCANTE.SetFocus

Else
    
    If Nnull(Me.txtCritica.Tag, 1) <> "" Then
        
        Dim X As Integer
        
        X = MsgBox("Deseja movimentar o marcante " & Nnull(Me.txtCritica.Tag, 1), 36 + vbDefaultButton2)
        If X = vbYes Then
            Limpa
            Me.TXTMARCANTE.text = Nnull(Me.txtCritica.Tag, 1)
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

'Sql = "SELECT "
'Sql = Sql & " MERCADORIA,"
'Sql = Sql & " MARCA,"
'Sql = Sql & " CONVERT(CHAR,DATA_ENTRADA,0) AS ENTRADA,"
'Sql = Sql & " IMPORTADOR,"
'Sql = Sql & " CNTR_DESOVA, "
'Sql = Sql & " TIPO_DOC,DESCR_ARMAZEM, POSICAO,"
'Sql = Sql & " CASE CANAL_ALF WHEN 0 THEN 'AMARELO' WHEN 1 THEN 'VERMELHO' WHEN 2 THEN 'VERDE' WHEN 3 THEN 'CINZA' WHEN 9 THEN '' END AS CANAL_ALF,"
'Sql = Sql & " MOTIVO_PROX_MVTO + ' ' + CONVERT(CHAR,HORAS_PROX_MVTO,0) AS MOTIVO_PROX_MVTO ,"
'Sql = Sql & " VOLUME,"
'Sql = Sql & " IMO,"
'Sql = Sql & " NVOCC,"
'Sql = Sql & " BL"
'Sql = Sql & " FROM OPERADOR..VW_INVENT_ARMAZEM"
'Sql = Sql & " WHERE LOTE='" & QualLote & "'"
'If Me.txtMarcante.text <> "" Then
'    Sql = Sql & " AND MARCANTE=" & Val(Me.txtMarcante.text)
'    Sql = Sql & " AND QTDE>0 "
'End If
SQl = "SELECT * FROM sgipa.dbo.fnInvent_Sistemas_Arm_CollOTE(" & QualLote & "," & Val(Me.TXTMARCANTE) & ")"
Set tb1 = New ADODB.Recordset
tb1.Open SQl, db, 1, 1

        Limpa

If Not tb1.EOF Then

    Me.txtMercadoria.text = Nnull(tb1!MERCADORIA, 1)
    Me.txtMarca.text = Nnull(tb1!Marca, 1)
    Me.txtEntrada.text = Nnull(tb1!Entrada, 1)
    Me.txtConteiner.text = Nnull(tb1!CNTR_DESOVA, 1)
    Me.txtCliente.text = Nnull(tb1!Importador, 1)
    
    If Nnull(tb1!Tipo_Doc, 1) <> "NOT DEFINED YET" Then
        Me.txtDOC.text = Nnull(tb1!Tipo_Doc, 1)
    Else
        Me.txtDOC.text = ""
    End If
    Me.txtCANAL.text = Nnull(tb1!CANAL_ALF, 1)
    Me.txtMOV.text = Nnull(tb1!MOTIVO_PROX_MVTO, 1)
    If Nnull(tb1!IMO, 1) <> "0" Then
        Me.txtIMO.text = Nnull(tb1!IMO, 1)
    Else
        Me.txtIMO.text = ""
    End If
    Me.txtNVOCC.text = Nnull(tb1!NVOCC, 1)
    Me.txtBL.text = Nnull(tb1!BL, 1)
    Me.TXTLOCAL.text = Nnull(tb1!DESCR_ARMAZEM, 1) & " " & Nnull(tb1!Posicao, 1)
    
    If Trim(Replace(Trim(Me.TXTLOCAL.text), "-", "")) = "" Then
        Me.cbMotivoPos.BoundText = 14
    Else
        Me.cbMotivoPos.BoundText = 1
    End If
    

    Call CarregarItens(QualLote, Nnull(Me.TXTMARCANTE.text, 0), Autonum_Cs_Yard)
    
    
   
    
Else
    Me.txtCritica = "Lote não encontrado"
    Me.txtCritica.ForeColor = &H80C0FF
End If
tb1.Close

    
End Sub

Private Sub CarregarItens(ByVal QualLote As String, Optional Marcante As Long = 0, Optional Autonum_Cs_Yard As Long = 0)

Dim SQl As String

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

SQl = "SELECT * FROM sgipa.dbo.fnInvent_Sistemas_Arm_ColItem(" & QualLote & "," & Val(Me.TXTMARCANTE) & "," & Autonum_Cs_Yard & ")"
Me.AdodcItem.ConnectionString = ConnectionString
Me.AdodcItem.RecordSource = SQl
Me.AdodcItem.Refresh

If Marcante > 0 And Me.AdodcItem.Recordset.RecordCount > 0 Then

    Me.cbItem.BoundText = Me.AdodcItem.Recordset("AUTONUM")
    
     
    Me.txtPrateleira.SetFocus
    
    
'Else

 '   Sql = " select ID_GRAVACAO as autonum, ( REPLACE(rtrim(convert(char,QTDE)),'.000','') +  '/' + rtrim(convert(char,QTDE_CAPTADA)) + ' ' + ISNULL(EMBALAGEM, ' ') + ' ' +  ISNULL(DESCR_ARMAZEM,' ') + ' ' + ISNULL(POSICAO, ' ')  ) AS DISPLAY"
 '   Sql = Sql & " FROM OPERADOR..VW_INVENT_ARMAZEM"
 '   Sql = Sql & " Where"
 '   Sql = Sql & " lote = " & QualLote
 '   Sql = Sql & " and Autonum_CS_Yard = 0 "
 '   Sql = Sql & " AND QTDE>0 ORDER BY DESCR_ARMAZEM,POSICAO"
    
 '   Me.AdodcItem.ConnectionString = ConnectionString
 '   Me.AdodcItem.RecordSource = Sql
 '   Me.AdodcItem.Refresh


End If

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

Me.txtCliente.text = ""
Me.txtConteiner.text = ""
Me.txtEmbalagem.text = ""
Me.txtEntrada.text = ""
Me.TXTLOCAL.text = ""
Me.txtMarca.text = ""
Me.txtMercadoria.text = ""
Me.txtQtde.text = ""
Me.txtQtdeM.text = ""
Me.txtQtdePos.text = ""
Me.txtNVOCC.text = ""
Me.txtDOC.text = ""
Me.txtCANAL.text = ""
Me.txtIMO.text = ""
Me.txtBL.text = ""
Me.txtMOV.text = ""
Me.txtLocalPos.text = ""
Me.txtLocalPos.Tag = ""
Me.txtCargaNoLocal.text = ""
Me.cbItem.text = ""
Me.cbArm.text = ""
Me.cbMotivoPos.text = ""
Me.txtPrateleira.text = ""
Me.txtAutonum_Cs_Yard.text = ""
Me.LBLEXPORTACAO.Visible = False
Me.Label2(32).Caption = "INVENTARIO DE CARGA SOLTA"
Me.TXTLOTE.text = ""



SetaNormal

End Sub
Function Validar() As Boolean

Validar = True
Me.txtCritica.ForeColor = vbGreen
Me.txtLocalPos.Tag = ""

Me.txtCritica.Tag = ""

If Me.TXTLOTE.text = "" Then
    Validar = False
    Me.txtCritica.text = "INFORME O LOTE"
    Me.txtCritica.ForeColor = &H80C0FF
    Exit Function
End If

If Me.Label2(32).Tag = "C" And Me.LBLEXPORTACAO.Visible = False Then
    If Me.cbItem.BoundText = "" Then
        Validar = False
        Me.txtCritica.text = "ITEM NAO INFORMADO"
        Me.txtCritica.ForeColor = &H80C0FF
        Exit Function
    End If
End If

'If Me.Label2(32).Tag <> "M" Then  'CARGAS EM RECONHECIMENTO

    If Nnull(Me.txtQtdePos.text, 0) = 0 Then
        Validar = False
        Me.txtCritica.text = "QUANTIDADE NAO INFORMADA"
        Me.txtCritica.ForeColor = &H80C0FF
        Exit Function
    End If
    
            
    If Val(Me.txtQtdePos.text) > Val(Me.txtQtde.text) Then
        Validar = False
        Me.txtCritica.text = "QUANTIDADE INDISPONIVEL"
        Me.txtCritica.ForeColor = &H80C0FF
        Exit Function
    End If
    
'End If

If Nnull(Me.cbArm.BoundText, 0) = 0 Then
    Validar = False
    Me.txtCritica.text = "ARMAZEM NAO INFORMADO"
    Me.txtCritica.ForeColor = &H80C0FF
    Exit Function
End If




'valida IMO no arm.22
If cbArm.BoundText = 22 Then
    SQl = "SELECT imo FROM SGIPA..TB_CARGA_SOLTA WHERE bl=" & TXTLOTE.text & " and isnull(imo,'')<>''"
    Dim RsI As New ADODB.Recordset
    PRSet RsI, SQl
    If Not RsI.EOF Then
        If RsI!IMO = "2.1" Or RsI!IMO = "3" Or RsI!IMO = "4.1" Or RsI!IMO = "4.2" Or RsI!IMO = "4.3" Or RsI!IMO = "3.0" Then
            If RsI!IMO <> "4.3" Then
                
                If InStr("L1-L2-L3-L4-L5-L6-L7-L8-L9-L10", UCase(txtLocalPos.text)) = 0 Then
                
                    Validar = False
                    Me.txtCritica.text = "Local Bloqueado, carga Inflamável"
                    Me.txtCritica.ForeColor = &H80C0FF
                    Exit Function
                End If
            Else
                If UCase(txtLocalPos.text) <> "L10" Then
                    Validar = False
                    Me.txtCritica.text = "Local Bloqueado, carga Inflamável (L10)"
                    Me.txtCritica.ForeColor = &H80C0FF
                    Exit Function
                End If
            End If
        End If
    End If
    RsI.Close
End If


If Me.txtLocalPos.text = "" And Me.Check1.Visible = False Then 'nao ct
    Validar = False
    Me.txtCritica.text = "LOCAL NAO INFORMADO"
    Me.txtCritica.ForeColor = &H80C0FF
    Exit Function
End If

If Me.txtLocalPos.text = "CAM" And Me.LBLEXPORTACAO.Visible = False Then
    Validar = False
    Me.txtCritica.text = "CARREGAMENTO DO IPA INVALIDO"
    Me.txtCritica.ForeColor = &H80C0FF
    Exit Function
End If


If Nnull(Me.cbMotivoPos.BoundText, 0) = 0 Then
    Validar = False
    Me.txtCritica.text = "MOTIVO NAO INFORMADO"
    Me.txtCritica.ForeColor = &H80C0FF
    Exit Function
End If
       









If Me.TXTMARCANTE.text <> "" Then
    SQl = "SELECT AUTONUM,AUTONUM_CARGA,ISNULL(AUTONUM_CARGA_AC,0) AS AUTONUM_CARGA_AC FROM SGIPA..TB_MARCANTES WHERE AUTONUM=" & Val(Me.TXTMARCANTE)
    Set TbM = New ADODB.Recordset
    TbM.Open SQl, db, 1, 1
    
    If TbM.EOF Then
        Validar = False
        Me.txtCritica.text = "MARCANTE INVALIDO"
        Me.txtCritica.ForeColor = &H80C0FF
        Exit Function
    Else
    
        
        'If Me.Label2(32).Tag <> "M" Then
        
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
                        Me.txtCritica.text = "MARCANTE ASSOCIADO A OUTRA CARGA"
                        Me.txtCritica.ForeColor = &H80C0FF
                        Exit Function
                    End If
                End If
            End If
        
        'End If
    
    End If
    TbM.Close
End If


If Nnull(Me.txtQtdeM.text, 0) > 0 Then
    If Nnull(Me.txtQtdePos.text, 0) <> Nnull(Me.txtQtdeM.text, 0) Then
        Validar = False
        Me.txtCritica.text = "DIVERGENCIA COM A QTDE MARCANTE"
        Me.txtCritica.ForeColor = &H80C0FF
        Exit Function
    End If
End If


If Me.Check1.Visible = False Then 'apenas ct

    Dim LocalCorrigido As String
    
    If Len(Me.txtLocalPos.text) = 4 Then
        Set tb10 = New ADODB.Recordset
        tb10.Open "SELECT autonum from sgipa..tb_yard_cs where yard='" & Me.txtLocalPos.text & "'", db, 1, 1
        If tb10.EOF Then
            LocalCorrigido = Left$(Me.txtLocalPos, 3) & "01" & Right$(Me.txtLocalPos, 1)
        End If
        tb10.Close
    End If
    
    
    Set Rst = New ADODB.Recordset
    
    
    Dim PosicaoPilha As Boolean
    PosicaoPilha = False
    
    SQl = "SELECT "
    SQl = SQl & " AUTONUM,VALIDA,FLAG_BLOQUEIO,LIMITE_MARCANTES,YARD "
    SQl = SQl & " FROM SGIPA..TB_YARD_CS "
    SQl = SQl & " WHERE ARMAZEM=" & Me.cbArm.BoundText & "  AND YARD='" & Me.txtLocalPos.text & "'"
    
    Rst.Open SQl, db, 1, 1
    If Rst.EOF Then
        If LocalCorrigido <> "" Then
            Rst.Close
            SQl = "SELECT "
            SQl = SQl & " AUTONUM,VALIDA,FLAG_BLOQUEIO,LIMITE_MARCANTES,YARD "
            SQl = SQl & " FROM SGIPA..TB_YARD_CS "
            SQl = SQl & " WHERE ARMAZEM=" & Me.cbArm.BoundText & "  AND YARD='" & LocalCorrigido & "'"
            Rst.Open SQl, db, 1, 1
            If Rst.EOF Then
                Validar = False
                Me.txtCritica.text = "LOCAL NAO ENCONTRADO"
                Me.txtCritica.ForeColor = &H80C0FF
                Exit Function
            Else
                If Nnull(Rst!Valida, 0) = 1 Then
                    PosicaoPilha = True
                Else
                    PosicaoPilha = False
                End If
                
                If Nnull(Rst!FLAG_BLOQUEIO, 0) = 1 Then
                    Validar = False
                    Me.txtCritica.text = "LOCAL BLOQUEADO"
                    Me.txtCritica.ForeColor = &H80C0FF
                    Exit Function
                End If
                
                
                
                Me.txtLocalPos.Tag = LocalCorrigido
                
            End If
        Else
            Validar = False
            Me.txtCritica.text = "LOCAL NAO ENCONTRADO"
            Me.txtCritica.ForeColor = &H80C0FF
            Exit Function
        End If
    Else
        If Nnull(Rst!Valida, 0) = 1 Then
            
            PosicaoPilha = True
        
        Else
        
            PosicaoPilha = False
            
            If Nnull(Rst!LIMITE_MARCANTES, 0) > 0 Then
                SQl = "SELECT COUNT(0) as QTOS FROM OPERADOR..VW_INVENT_ARMAZEM WHERE AUTONUM_ARMAZEM= " & Me.cbArm.BoundText
                SQl = SQl & " AND POSICAO='" & Rst!Yard & "'"
                Set tbq = New ADODB.Recordset
                tbq.Open SQl, db, 1, 1
                If Not tbq.EOF Then
                    If Nnull(tbq!qtos, 0) > Nnull(Rst!LIMITE_MARCANTES, 0) Then
                        Validar = False
                        Me.txtCritica.text = "LIMITE DE CARGA NO LOCAL"
                        Me.txtCritica.ForeColor = &H80C0FF
                        Exit Function
                    End If
                End If
                tbq.Close
                    
            End If
            
            
        End If
        
        If Nnull(Rst!FLAG_BLOQUEIO, 0) = 1 Then
            Validar = False
            Me.txtCritica.text = "LOCAL BLOQUEADO"
            Me.txtCritica.ForeColor = &H80C0FF
            Exit Function
        End If
        
    End If
    Rst.Close
    
    'Se existe outra Carga no Local
    
'    If PosicaoPilha Then
'
'        If db.Execute("SELECT ISNULL(MAX(FLAG_CS_PERMITE_CONC),0) FROM SGIPA..DTE_TB_PARAMETROS")(0) = 0 Then
'
'
'            If Right(IIf(Me.txtLocalPos.Tag <> "", Me.txtLocalPos.Tag, Me.txtLocalPos.text), 1) <> "1" Then
'
'                SQl = "SELECT LOTE, MARCANTE FROM OPERADOR..VW_INVENT_ARMAZEM WHERE AUTONUM_ARMAZEM=" & Me.cbArm.BoundText & " AND POSICAO='" & IIf(Me.txtLocalPos.Tag <> "", Me.txtLocalPos.Tag, Me.txtLocalPos.text) & "' AND MARCANTE<>" & Val(Me.TXTMARCANTE.text) & " AND QTDE>0 "
'                SQl = SQl & " AND LOTE<> " & Me.TXTLOTE
'
'                Set tb1 = New ADODB.Recordset
'                tb1.Open SQl, db, 1, 1
'                If Not tb1.EOF Then
'
'                    Validar = False
'                    Me.txtCritica.text = "OUTRO LOTE NESTE LOCAL. MC: " & Format(tb1!Marcante, "000000000")
'                    Me.txtCritica.ForeColor = &H80C0FF
'                    Me.txtCritica.Tag = Format(tb1!Marcante, "000000000")
'                    Exit Function
'
'                End If
'                tb1.Close
'            End If
'        End If
'
'
'    End If
    
End If





Dim RetAnvisa As String
SQl = "SELECT [dbo].[FN_VALIDA_ANVISA_ARMAZEM](" & Me.TXTLOTE.text & "," & Me.cbArm.BoundText & ",'" & Me.txtLocalPos.text & "')"
RetAnvisa = Nnull(db.Execute(SQl).Fields(0), 1)
If RetAnvisa <> "OK" Then
    Validar = False
    Me.txtCritica.text = RetAnvisa
    Me.txtCritica.ForeColor = &H80C0FF
    Exit Function
End If



Dim RetMisturaLote As String
SQl = "SELECT [OPERADOR].[dbo].[FN_VALIDA_LOTE_PRATELEIRA]('" & Me.TXTLOTE.text & "'," & Me.cbArm.BoundText & ",'" & Me.txtLocalPos.text & "')"
RetMisturaLote = Nnull(db.Execute(SQl).Fields(0), 1)
If RetMisturaLote <> "OK" Then
    Validar = False
    Me.txtCritica.text = "Outro Lote na posicao " & RetMisturaLote
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



'If Len(Me.txtMarcante) = 9 Then
'
'    Carrega_Dados
'
'
'Else
'    Limpa
'End If
End Sub

Private Sub Carrega_Dados()

    Me.txtCritica.text = ""
    

    Me.LBLEXPORTACAO.Visible = False
    
    Set tb1 = New ADODB.Recordset
    'Sql = "SELECT AUTONUM, ISNULL(AUTONUM_CARGA,0) AS AUTONUM_CARGA, ISNULL(AUTONUM_CARGA_AC,0) AS AUTONUM_CARGA_AC ,VOLUMES,ISNULL(AUTONUM_CS_YARD,0) AS AUTONUM_CS_YARD,ISNULL(AUTONUM_CS_YARD_AC,0) AS AUTONUM_CS_YARD_AC, ISNULL(FLAG_RECONHECIMENTO,0) AS FLAG_RECONHECIMENTO, "
    'Sql = Sql & " ISNULL(LOTE_IMPRESSO,0) AS LOTE_IMPRESSO, ISNULL(ID_CONTEINER_IMPRESSO,' ') AS ID_CONTEINER_IMPRESSO, "
    'Sql = Sql & " ISNULL(ARMAZEM_REC,0) AS ARMAZEM_REC, ISNULL(YARD_REC,' ') AS YARD_REC,ISNULL(AUTONUM_CEXP,0) AS AUTONUM_CEXP "
    'Sql = Sql & " FROM SGIPA..TB_MARCANTES WHERE AUTONUM=" & Val(Me.txtMarcante) & " AND (ISNULL(AUTONUM_CARGA,0)>0 or ISNULL(AUTONUM_CARGA_AC,0)>0 OR ISNULL(FLAG_RECONHECIMENTO,0)=1  or ISNULL(AUTONUM_CEXP,0)>0) "
    SQl = "SELECT * FROM sgipa.dbo.fnInvent_Sistemas_Arm_Col_patio(" & Val(Me.TXTMARCANTE) & ")"
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
       
            If tb1!autonum_carga > 0 Then
                
                Me.Label2(32).ForeColor = &HBA9F8D
                Me.Label2(32).Tag = "C"
                
                Set Tb2 = New ADODB.Recordset
                Tb2.Open "SELECT BL,ITEM FROM SGIPA..TB_CARGA_SOLTA WHERE AUTONUM=" & tb1!autonum_carga, db, 1, 1
                If Not Tb2.EOF Then
                    Call Carrega_Dados_Lote(Nnull(Tb2!BL, 0), Nnull(tb1!Autonum_Cs_Yard, 0))
                    Me.TXTLOTE.text = Nnull(Tb2!BL, 0)
                    Me.txtQtde.text = Nnull(tb1!VOLUMES, 0)
                    Me.txtQtdeM.text = Nnull(tb1!VOLUMES, 0)
                    Me.txtQtdePos.text = Nnull(tb1!VOLUMES, 0)
                    Me.txtAutonum_Cs_Yard.text = Nnull(tb1!Autonum_Cs_Yard, 0)
                    Me.txtPrateleira.SetFocus
                         
                         
                       If tb1!FLAG_RECONHECIMENTO = 1 Then
                       Me.Label2(32).ForeColor = vbRed
                       Me.Label2(32).Caption = "EM RECONHECIMENTO"
                       Me.Label2(32).Tag = "M"
                       Me.txtConteiner.text = tb1!ID_CONTEINER_IMPRESSO
                       Me.TXTLOTE.text = tb1!LOTE_IMPRESSO
                    
                        Dim ArmD As String
                        ArmD = Nnull(db.Execute(" SELECT DESCR FROM SGIPA..TB_ARMAZENS_IPA WHERE AUTONUM=" & Nnull(tb1!ARMAZEM_REC, 0))(0), 1)
                        Me.TXTLOCAL.text = ArmD & " " & Nnull(tb1!YARD_REC, 1)
                    End If
                         
                Else
                    Limpa
                End If
            
            Else
             
                Me.Label2(32).ForeColor = vbRed
                Me.Label2(32).Tag = "Z"
                
                Set Tb2 = New ADODB.Recordset
                Tb2.Open "SELECT LOTE,BL,EMBALAGEM,IMPORTADOR,NAVIO,QTDE,AUTONUMCS,MERCADORIA,DATA_ENTRADA,TIPO_DOC,DESCR_ARMAZEM,POSICAO,ID_GRAVACAO FROM OPERADOR..VW_INVENT_ARMAZEM WHERE MARCANTE= " & Val(Me.TXTMARCANTE), db, 1, 1
                 
                If Not Tb2.EOF Then
                     
                    Me.TXTLOTE.text = Nnull(Tb2!Lote, 0)
                    Me.TXTLOTE.Tag = Nnull(Tb2!id_Gravacao, 1)
                                        
                    Me.txtMercadoria.text = Nnull(Tb2!MERCADORIA, 1)
                    Me.txtMarca.text = ""
                    Me.txtEntrada.text = Nnull(Tb2!Data_Entrada, 1)
                    Me.txtConteiner.text = ""
                    Me.txtCliente.text = Nnull(Tb2!Importador, 1)
                    Me.txtEmbalagem.text = Nnull(Tb2!embalagem, 1)
                    
                    If Nnull(Tb2!Tipo_Doc, 1) <> "NOT DEFINED YET" Then
                        Me.txtDOC.text = Nnull(Tb2!Tipo_Doc, 1)
                    Else
                        Me.txtDOC.text = ""
                    End If
                    Me.txtCANAL.text = ""
                    Me.txtMOV.text = ""
                    
                    Me.txtIMO.text = ""
                    Me.txtNVOCC.text = ""
                    Me.txtBL.text = Nnull(Tb2!BL, 1)
                    Me.TXTLOCAL.text = Nnull(Tb2!DESCR_ARMAZEM, 1) & " " & Nnull(Tb2!Posicao, 1)
                    If Trim(Me.TXTLOCAL.text) = "" Or Trim(Me.TXTLOCAL.text) = "-" Then Me.cbMotivoPos.BoundText = 14
                                     
                Else
                 
                    Limpa
                 
                End If
                 
                Me.txtQtde.text = Nnull(tb1!VOLUMES, 0)
                Me.txtQtdeM.text = Nnull(tb1!VOLUMES, 0)
                Me.txtQtdePos.text = Nnull(tb1!VOLUMES, 0)
                Me.txtAutonum_Cs_Yard.text = Nnull(tb1!Autonum_Cs_Yard_AC, 0)
                Me.txtPrateleira.SetFocus
             
            End If
    
        Else 'CARGA DE EXPORTACAO
        
        
            Set Tb2 = New ADODB.Recordset
            Tb2.Open "SELECT LOTE,BL,EMBALAGEM,IMPORTADOR,NAVIO,QTDE,AUTONUMCS,MERCADORIA,DATA_ENTRADA,TIPO_DOC,DESCR_ARMAZEM,POSICAO,ID_GRAVACAO FROM OPERADOR..VW_INVENT_ARMAZEM WHERE MARCANTE= " & Val(Me.TXTMARCANTE), db, 1, 1
                 
            If Not Tb2.EOF Then
                     
                Me.LBLEXPORTACAO.Visible = True
                     
                Me.TXTLOTE.text = Nnull(Tb2!Lote, 1)
                Me.TXTLOTE.Tag = Nnull(Tb2!id_Gravacao, 1)
                                        
                Me.txtMercadoria.text = Nnull(Tb2!MERCADORIA, 1)
                Me.txtMarca.text = ""
                Me.txtEntrada.text = Nnull(Tb2!Data_Entrada, 1)
                Me.txtConteiner.text = ""
                Me.txtCliente.text = Nnull(Tb2!Importador, 1)
                Me.txtEmbalagem.text = Nnull(Tb2!embalagem, 1)
                Me.txtDOC.text = ""
                Me.txtCANAL.text = ""
                Me.txtMOV.text = ""
                Me.txtCliente.text = Nnull(Tb2!Importador, 0)
                        
                Me.txtIMO.text = ""
                Me.txtNVOCC.text = ""
                Me.txtBL.text = Nnull(Tb2!BL, 1)
                Me.TXTLOCAL.text = Nnull(Tb2!DESCR_ARMAZEM, 1) & " " & Nnull(Tb2!Posicao, 1)
                If Trim(Me.TXTLOCAL.text) = "" Or Trim(Me.TXTLOCAL.text) = "-" Then
                    Me.cbMotivoPos.BoundText = 14
                Else
                     Me.cbMotivoPos.BoundText = 1
                End If
                
                Me.txtQtde.text = Nnull(tb1!VOLUMES, 0)
                Me.txtQtdeM.text = Nnull(tb1!VOLUMES, 0)
                Me.txtQtdePos.text = Nnull(tb1!VOLUMES, 0)
                Me.txtAutonum_Cs_Yard.text = Nnull(tb1!Autonum_Cs_Yard, 0)
                
                Me.txtPrateleira.SetFocus

                                     
            Else
                 
                Limpa
             
            End If
        
        End If
    
        
    Else
        Limpa
    End If
End Sub


Private Sub txtMarcante_GotFocus()
'Me.txtMarcante.Tag = "0"
End Sub

Private Sub txtMarcante_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'    Me.txtMarcante.Tag = "1"
'    HorarioEnter = Now
'End If
End Sub

Private Sub txtMarcante_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    Sendkeys "{TAB}"
'Else
'    If Len(txtMarcante.text) = 0 Then
'        HorarioNumero = Now
'    End If
'End If

End Sub

Private Sub txtMarcante_LostFocus()
'Dim DifSeg As Long
'DifSeg = Abs(DateDiff("s", HorarioEnter, HorarioNumero))


'If txtMarcante.Tag <> "1" Or DifSeg > 2 Then
'    X = MsgBox("Campo não permite digitação manual", vbExclamation)
'    Me.txtMarcante.text = ""
'Else
'    Carrega_Dados
'    txtMarcante.Enabled = False
    
'End If

Carrega_Dados


End Sub

Private Sub txtPrateleira_Change()
Dim SQl As String

If Len(Me.txtPrateleira.text) = 9 Then

    Set TBP = New ADODB.Recordset
    Dim AutonumArm$
    Dim YardCs As String
    
    SetaNormal
    
    
    If Col_Patio = 1 Or Col_Patio = 7 Then
        AutonumArm$ = Left(Me.txtPrateleira.text, 2)
    Else
        AutonumArm$ = Nnull(db.Execute("SELECT max(AUTONUM) FROM SGIPA..TB_ARMAZENS_IPA WHERE PATIO=" & Col_Patio & " AND AUTONUM_ARM_PATIO=" & Val(Left(Me.txtPrateleira.text, 2)))(0), 0)
    End If
    
    YardCs = UCase(Replace(Mid$(Me.txtPrateleira.text, 3), "-", ""))
    
    '01L011
    '01COFRE
    
    Dim xArm As String
    Dim Xpos As String
    xArm = AutonumArm$
    Xpos = Replace(UCase(Mid(Me.txtPrateleira.text, 3)), "-", "")
    Set tb1 = New ADODB.Recordset
    tb1.Open "SELECT AUTONUM FROM sgipa..tb_yard_cs where armazem=" & Val(xArm) & " and yard='" & Xpos & "'", db, 1, 1
    If Not tb1.EOF Then
        YardCs = Xpos
    Else
        YardCs = Left$(YardCs, 3) & "01" & Right$(YardCs, 1)
        tb1.Close
        tb1.Open "SELECT AUTONUM FROM sgipa..tb_yard_cs where armazem=" & Val(xArm) & " and yard='" & Left$(YardCs, 3) & "01" & Right$(YardCs, 1) & "'", db, 1, 1
        If Not tb1.EOF Then
            YardCs = Left$(YardCs, 3) & "01" & Right$(YardCs, 1)
        Else
            YardCs = ""
        End If
    End If
    tb1.Close
    
    
    
    Me.cbArm.BoundText = Val(AutonumArm)
    Me.txtLocalPos.text = YardCs

ElseIf Len(Me.txtPrateleira.text) = 10 Then

    Set TBP = New ADODB.Recordset
    'Dim AutonumArm$
    'Dim YardCs As String
    
    SetaNormal
    
    
    If Col_Patio = 1 Or Col_Patio = 7 Then
        AutonumArm$ = Left(Me.txtPrateleira.text, 2)
    Else
        AutonumArm$ = Nnull(db.Execute("SELECT max(AUTONUM) FROM SGIPA..TB_ARMAZENS_IPA WHERE PATIO=" & Col_Patio & " AND AUTONUM_ARM_PATIO=" & Val(Left(Me.txtPrateleira.text, 2)))(0), 0)
    End If
    
    YardCs = UCase(Replace(Mid$(Me.txtPrateleira.text, 3), "-", ""))
    
    '01LA011
    '01COFRE
    
    'Dim xArm As String
    'Dim Xpos As String
    xArm = AutonumArm$
    Xpos = Replace(UCase(Mid(Me.txtPrateleira.text, 3)), "-", "")
    Set tb1 = New ADODB.Recordset
    tb1.Open "SELECT AUTONUM FROM sgipa..tb_yard_cs where armazem=" & Val(xArm) & " and yard='" & Xpos & "'", db, 1, 1
    If Not tb1.EOF Then
        YardCs = Xpos
    Else
        YardCs = Left$(YardCs, 4) & "01" & Right$(YardCs, 1)
        tb1.Close
        tb1.Open "SELECT AUTONUM FROM sgipa..tb_yard_cs where armazem=" & Val(xArm) & " and yard='" & Left$(YardCs, 4) & "01" & Right$(YardCs, 1) & "'", db, 1, 1
        If Not tb1.EOF Then
            YardCs = Left$(YardCs, 4) & "01" & Right$(YardCs, 1)
        Else
            YardCs = ""
        End If
    End If
    tb1.Close
    
    
    
    Me.cbArm.BoundText = Val(AutonumArm)
    Me.txtLocalPos.text = YardCs

ElseIf Len(Me.txtPrateleira.text) >= 4 And Len(Me.txtPrateleira.text) < 12 And InStr(1, "-", Me.txtPrateleira.text) = 0 Then

    SetaNormal
    

    If Col_Patio = 1 Or Col_Patio = 7 Then
        AutonumArm$ = Left(Me.txtPrateleira.text, 2)
    Else
        AutonumArm$ = Nnull(db.Execute("SELECT max(AUTONUM) FROM SGIPA..TB_ARMAZENS_IPA WHERE PATIO=" & Col_Patio & " AND AUTONUM_ARM_PATIO=" & Val(Left(Me.txtPrateleira.text, 2)))(0), 0)
    End If
    
    YardCs = UCase(Replace(Mid$(Me.txtPrateleira.text, 3), "-", ""))
    
    xArm = AutonumArm$
    Xpos = Replace(UCase(Mid(Me.txtPrateleira.text, 3)), "-", "")
    
    Set tb1 = New ADODB.Recordset
    tb1.Open "SELECT AUTONUM FROM sgipa..tb_yard_cs where armazem=" & Val(xArm) & " and yard='" & Xpos & "'", db, 1, 1
    If Not tb1.EOF Then
        YardCs = Xpos
    Else
        If Col_Patio = 1 Or Col_Patio = 7 Then
            tb1.Close
            tb1.Open "SELECT AUTONUM FROM sgipa..tb_yard_cs where armazem=" & Val(xArm) & " and yard='" & Left$(YardCs, 3) & "01" & Right$(YardCs, 1) & "'", db, 1, 1
            If Not tb1.EOF Then
                YardCs = Left$(YardCs, 3) & "01" & Right$(YardCs, 1)
            Else
                YardCs = ""
            End If
        Else
            tb1.Close
            tb1.Open "SELECT AUTONUM FROM sgipa..tb_yard_cs where armazem=" & Val(xArm) & " and yard='" & Left$(YardCs, 4) & "01" & Right$(YardCs, 1) & "'", db, 1, 1
            If Not tb1.EOF Then
                YardCs = Left$(YardCs, 4) & "01" & Right$(YardCs, 1)
            Else
                YardCs = ""
            End If
        End If
    End If
    tb1.Close
    
    Me.cbArm.BoundText = Val(AutonumArm)
    Me.txtLocalPos.text = YardCs

ElseIf Len(Me.txtPrateleira.text) = 12 And InStr(1, Me.txtPrateleira.text, "-") = 11 Then 'conteiner armazem

   
    SetaNormal
    
    Set tbCT = New ADODB.Recordset
    tbCT.Open "SELECT autonum FROM SGIPA..tb_armazens_ipa where descr='" & Me.txtPrateleira.text & "' and flag_ct=1 and flag_historico=0", db, 1, 1
    If Not tbCT.EOF Then
        SetaCT
        Me.cbArm.BoundText = Val(tbCT!AUTONUM)
        Me.TXTLOCAL.text = ""
    End If
    
ElseIf Len(Me.txtPrateleira.text) = 12 And InStr(1, Me.txtPrateleira.text, "-") = 3 Then
    
    Set TBP = New ADODB.Recordset
    
    
    SetaNormal
    
    
    If Col_Patio = 1 Or Col_Patio = 7 Then
        AutonumArm$ = Left(Me.txtPrateleira.text, 2)
    Else
        AutonumArm$ = Nnull(db.Execute("SELECT max(AUTONUM) FROM SGIPA..TB_ARMAZENS_IPA WHERE PATIO=" & Col_Patio & " AND AUTONUM_ARM_PATIO=" & Val(Left(Me.txtPrateleira.text, 2)))(0), 0)
    End If
    
    YardCs = UCase(Replace(Mid$(Me.txtPrateleira.text, 3), "-", ""))
    
    '01L011
    '01COFRE
    
    
    xArm = AutonumArm$
    Xpos = Replace(UCase(Mid(Me.txtPrateleira.text, 3)), "-", "")
    
    Set tb1 = New ADODB.Recordset
    tb1.Open "SELECT AUTONUM FROM sgipa..tb_yard_cs where armazem=" & Val(xArm) & " and yard='" & Xpos & "'", db, 1, 1
    If Not tb1.EOF Then
        YardCs = Xpos
    Else
        YardCs = ""
    End If
    tb1.Close
    
    
    
    Me.cbArm.BoundText = Val(AutonumArm)
    Me.txtLocalPos.text = YardCs

Else


    Me.cbArm.BoundText = 0
    Me.txtLocalPos.text = ""

End If

End Sub

Private Sub txtPrateleira_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbEnter Then

    If Len(Me.txtPrateleira.text) = 9 Then
    
        Set TBP = New ADODB.Recordset
        Dim AutonumArm$
        Dim YardCs As String
        
        AutonumArm$ = Left(Me.txtPrateleira.text, 2)
        YardCs = UCase(Replace(Mid$(Me.txtPrateleira.text, 3), "-", ""))
        YardCs = Left$(YardCs, 3) & "01" & Right$(YardCs, 1)
        
        Me.cbArm.BoundText = Val(AutonumArm)
        Me.txtLocalPos.text = YardCs
    
    ElseIf Len(Me.txtPrateleira.text) = 10 Then
    
        Set TBP = New ADODB.Recordset
        'Dim AutonumArm$
        'Dim YardCs As String
        
        AutonumArm$ = Left(Me.txtPrateleira.text, 2)
        YardCs = UCase(Replace(Mid$(Me.txtPrateleira.text, 3), "-", ""))
        YardCs = Left$(YardCs, 4) & "01" & Right$(YardCs, 1)
        
        Me.cbArm.BoundText = Val(AutonumArm)
        Me.txtLocalPos.text = YardCs
    
    Else
    
        Me.cbArm.BoundText = 0
        Me.txtLocalPos.text = ""
    
    End If
    
End If

End Sub

Private Sub SetaNormal()
Me.txtLocalPos.Visible = True
Me.Label2(10).Visible = True
Me.txtCargaNoLocal.Visible = True
Me.Check1.Visible = False
Me.Check2.Visible = False
Me.Check3.Visible = False
Me.Check4.Visible = False

Me.Check1.Value = 0
Me.Check2.Value = 0
Me.Check3.Value = 0
Me.Check4.Value = 0
Me.cbOC.Visible = False
Me.cbOC.ListIndex = -1

End Sub

Private Sub SetaCT()
Me.txtLocalPos.Visible = False
Me.Label2(10).Visible = False
Me.txtCargaNoLocal.Visible = False
Me.Check1.Visible = True
Me.Check2.Visible = True
Me.Check3.Visible = True
Me.Check4.Visible = True
Me.txtLocalPos.text = ""
Me.cbOC.Visible = True

End Sub
