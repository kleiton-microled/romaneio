VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form DesovaPier 
   Appearance      =   0  'Flat
   BackColor       =   &H00BA9F8D&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3615
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
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtViagem 
      Height          =   315
      Left            =   3480
      TabIndex        =   48
      Text            =   "Text1"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtdt 
      Height          =   315
      Left            =   1200
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00BA9F8D&
      ForeColor       =   &H80000008&
      Height          =   5025
      Left            =   -30
      TabIndex        =   2
      Top             =   -90
      Width           =   3675
      Begin VB.OptionButton Option2 
         BackColor       =   &H00BA9F8D&
         Caption         =   "DESOVA "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   1
         Left            =   1860
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   150
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00BA9F8D&
         Caption         =   "CONTÊINER"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   150
         Value           =   -1  'True
         Width           =   1665
      End
      Begin VB.Frame FrameCntr 
         Appearance      =   0  'Flat
         BackColor       =   &H00BA9F8D&
         ForeColor       =   &H80000008&
         Height          =   4515
         Left            =   30
         TabIndex        =   18
         Top             =   480
         Width           =   3615
         Begin VB.CommandButton cmdSair 
            Caption         =   "[ESC]-SAIR"
            Height          =   345
            Left            =   2220
            TabIndex        =   105
            Top             =   3840
            Width           =   975
         End
         Begin VB.CommandButton cmdTermino 
            Caption         =   "[F2]TERMINO"
            Height          =   345
            Left            =   1080
            TabIndex        =   104
            Top             =   3840
            Width           =   1125
         End
         Begin VB.CommandButton cmdInicio 
            Caption         =   "[F1]INICIO"
            Height          =   345
            Left            =   60
            TabIndex        =   103
            Top             =   3840
            Width           =   1005
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   2760
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
         Begin MSAdodcLib.Adodc AdodcFiscal 
            Height          =   330
            Left            =   4680
            Top             =   4740
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
         Begin VB.TextBox txtLacreIPA 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   3510
            Width           =   1695
         End
         Begin VB.TextBox txtLacreOrigem 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   3240
            Width           =   1695
         End
         Begin VB.OptionButton Option3 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            Caption         =   "MECAN"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   2370
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   2970
            Width           =   825
         End
         Begin VB.OptionButton Option3 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            Caption         =   "MANUAL"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   1500
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   2970
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.TextBox txtConferente 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   1590
            Width           =   1695
         End
         Begin VB.TextBox txtTermino 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtInicio 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1050
            Width           =   1695
         End
         Begin VB.TextBox txtYard 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   780
            Width           =   1695
         End
         Begin VB.TextBox txtTamTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
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
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   510
            Width           =   1695
         End
         Begin MSMask.MaskEdBox MskCntr 
            Height          =   360
            Left            =   1500
            TabIndex        =   1
            Top             =   150
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">&&&&&&&&&&-&"
            PromptChar      =   "_"
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            ForeColor       =   &H80000008&
            Height          =   1065
            Left            =   90
            TabIndex        =   76
            Top             =   1890
            Width           =   3105
            Begin VB.CommandButton Command1 
               Caption         =   "?"
               Height          =   315
               Left            =   2820
               TabIndex        =   85
               Top             =   690
               Width           =   255
            End
            Begin MSDataListLib.DataCombo DBCombo10 
               Bindings        =   "DesovaPier.frx":0000
               Height          =   345
               Left            =   30
               TabIndex        =   83
               Top             =   660
               Width           =   2745
               _ExtentX        =   4842
               _ExtentY        =   609
               _Version        =   393216
               Enabled         =   0   'False
               MatchEntry      =   -1  'True
               ListField       =   "Display"
               BoundColumn     =   "Autonum"
               Text            =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.OptionButton Option4 
               Appearance      =   0  'Flat
               BackColor       =   &H00BA9F8D&
               Caption         =   "ALF."
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   5
               Left            =   2460
               Style           =   1  'Graphical
               TabIndex        =   82
               Top             =   330
               Width           =   615
            End
            Begin VB.OptionButton Option4 
               Appearance      =   0  'Flat
               BackColor       =   &H00BA9F8D&
               Caption         =   "DESP."
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   4
               Left            =   1860
               Style           =   1  'Graphical
               TabIndex        =   81
               Top             =   330
               Width           =   615
            End
            Begin VB.OptionButton Option4 
               Appearance      =   0  'Flat
               BackColor       =   &H00BA9F8D&
               Caption         =   "NVOCC"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   3
               Left            =   1170
               Style           =   1  'Graphical
               TabIndex        =   80
               Top             =   330
               Width           =   705
            End
            Begin VB.OptionButton Option4 
               Appearance      =   0  'Flat
               BackColor       =   &H00BA9F8D&
               Caption         =   "ARM."
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   2
               Left            =   600
               Style           =   1  'Graphical
               TabIndex        =   79
               Top             =   330
               Width           =   585
            End
            Begin VB.OptionButton Option4 
               Appearance      =   0  'Flat
               BackColor       =   &H00BA9F8D&
               Caption         =   "IMP."
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   1
               Left            =   30
               Style           =   1  'Graphical
               TabIndex        =   78
               Top             =   330
               Width           =   585
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "SOLICITANTE"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   3
               Left            =   0
               TabIndex        =   77
               Top             =   90
               Width           =   3075
            End
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LACRE IPA  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   21
            Left            =   90
            TabIndex        =   35
            Top             =   3510
            Width           =   1395
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LACRE ORIGEM  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   20
            Left            =   90
            TabIndex        =   33
            Top             =   3240
            Width           =   1395
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "RECURSO  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   19
            Left            =   90
            TabIndex        =   30
            Top             =   2940
            Width           =   1395
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "COLABORADOR"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   17
            Left            =   90
            TabIndex        =   28
            Top             =   1590
            Width           =   1395
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TÉRMINO  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   16
            Left            =   90
            TabIndex        =   26
            Top             =   1320
            Width           =   1395
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "INÍCIO OP."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   15
            Left            =   90
            TabIndex        =   24
            Top             =   1050
            Width           =   1395
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "POS. PATIO  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   13
            Left            =   90
            TabIndex        =   22
            Top             =   780
            Width           =   1395
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "TAM/TIPO  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   12
            Left            =   90
            TabIndex        =   20
            Top             =   510
            Width           =   1395
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CONTÊINER  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   11
            Left            =   90
            TabIndex        =   19
            Top             =   150
            Width           =   1395
         End
      End
      Begin VB.Frame frmResumo 
         Appearance      =   0  'Flat
         BackColor       =   &H00BA9F8D&
         ForeColor       =   &H80000008&
         Height          =   4515
         Left            =   30
         TabIndex        =   89
         Top             =   480
         Visible         =   0   'False
         Width           =   3615
         Begin VB.CommandButton cmdCarregaItem 
            Caption         =   "[F4] CARREGA O ITEM SELECIONADO"
            Height          =   285
            Left            =   90
            TabIndex        =   106
            Top             =   3150
            Width           =   3435
         End
         Begin VB.TextBox TxtObs 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   960
            Left            =   480
            Locked          =   -1  'True
            MaxLength       =   400
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   101
            Top             =   3480
            Width           =   3075
         End
         Begin MSComctlLib.ListView ListView1 
            CausesValidation=   0   'False
            Height          =   2685
            Left            =   60
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   420
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   4736
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            TextBackground  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "OBS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   9
            Left            =   -30
            TabIndex        =   102
            Top             =   3720
            Width           =   435
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LOTES - CONTÊINER"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00BA9F8D&
            Height          =   285
            Index           =   8
            Left            =   30
            TabIndex        =   90
            Top             =   120
            Width           =   3495
         End
      End
      Begin VB.Frame FrameDesova 
         Appearance      =   0  'Flat
         BackColor       =   &H00BA9F8D&
         ForeColor       =   &H80000008&
         Height          =   4485
         Left            =   30
         TabIndex        =   37
         Top             =   480
         Visible         =   0   'False
         Width           =   3585
         Begin VB.Frame frmMarcantes 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            ForeColor       =   &H80000008&
            Height          =   2565
            Left            =   0
            TabIndex        =   93
            Top             =   1950
            Visible         =   0   'False
            Width           =   3525
            Begin VB.CommandButton cmdRemoveMc 
               Caption         =   "[F5]REMOVE"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2460
               TabIndex        =   108
               Top             =   540
               Width           =   1035
            End
            Begin VB.CommandButton cmdAdicionaMc 
               Caption         =   "[F4]ADIC."
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1590
               TabIndex        =   107
               Top             =   540
               Width           =   855
            End
            Begin VB.TextBox txtMarcante 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               IMEMode         =   3  'DISABLE
               Left            =   930
               TabIndex        =   96
               Top             =   120
               Width           =   2565
            End
            Begin VB.TextBox txtQtdeMarcante 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   930
               MaxLength       =   4
               TabIndex        =   95
               Top             =   510
               Width           =   645
            End
            Begin VB.TextBox txtPrateleira 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               IMEMode         =   3  'DISABLE
               Left            =   930
               MaxLength       =   10
               PasswordChar    =   "|"
               TabIndex        =   94
               Top             =   840
               Width           =   2565
            End
            Begin MSDataListLib.DataList DlMarcante 
               Bindings        =   "DesovaPier.frx":0015
               DataSource      =   "AdodcMarcante"
               Height          =   1230
               Left            =   30
               TabIndex        =   97
               Top             =   1260
               Width           =   3435
               _ExtentX        =   6059
               _ExtentY        =   2170
               _Version        =   393216
               MatchEntry      =   -1  'True
               Appearance      =   0
               BackColor       =   -2147483633
               ListField       =   "Display"
               BoundColumn     =   "Autonum"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSAdodcLib.Adodc AdodcMarcante 
               Height          =   330
               Left            =   1650
               Top             =   1260
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
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "MARCANTE "
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   405
               Index           =   1
               Left            =   30
               TabIndex        =   100
               Top             =   120
               Width           =   915
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "QTDE "
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   330
               Index           =   2
               Left            =   30
               TabIndex        =   99
               Top             =   510
               Width           =   915
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
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
               Height          =   405
               Index           =   7
               Left            =   30
               TabIndex        =   98
               Top             =   840
               Width           =   915
            End
         End
         Begin VB.CheckBox ckReconhecimento 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            Caption         =   "RECONHECIMENTO CARGA"
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
            Height          =   270
            Left            =   30
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   930
            Width           =   1455
         End
         Begin VB.TextBox txtBrutoAP 
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
            Height          =   270
            Left            =   2220
            TabIndex        =   9
            Top             =   930
            Width           =   1305
         End
         Begin VB.TextBox txtPesoAvariado 
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
            Left            =   2820
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   1980
            Width           =   705
         End
         Begin VB.TextBox txtQtdeAv 
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
            Left            =   3090
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   1740
            Width           =   435
         End
         Begin VB.CheckBox ChkAv 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            Caption         =   "IDFA"
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
            Index           =   10
            Left            =   2670
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   2700
            Width           =   885
         End
         Begin VB.CheckBox ChkAv 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            Caption         =   "IDFA"
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
            Height          =   165
            Index           =   11
            Left            =   2670
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   2910
            Width           =   885
         End
         Begin VB.TextBox txtConsertado 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
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
            Index           =   10
            Left            =   2520
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   2670
            Width           =   195
         End
         Begin VB.TextBox txtConsertado 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
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
            Index           =   11
            Left            =   2520
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   2880
            Width           =   165
         End
         Begin VB.CheckBox ChkAv 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            Caption         =   "IDFA"
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
            Index           =   8
            Left            =   2670
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   2250
            Width           =   885
         End
         Begin VB.CheckBox ChkAv 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            Caption         =   "IDFA"
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
            Index           =   9
            Left            =   2670
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   2460
            Width           =   885
         End
         Begin VB.TextBox txtConsertado 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
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
            Index           =   8
            Left            =   2520
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   2250
            Width           =   165
         End
         Begin VB.TextBox txtConsertado 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
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
            Index           =   9
            Left            =   2520
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   2460
            Width           =   165
         End
         Begin VB.CheckBox ChkAv 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            Caption         =   "IDFA"
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
            Left            =   1440
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   2880
            Width           =   1035
         End
         Begin VB.TextBox txtConsertado 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
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
            Index           =   7
            Left            =   1260
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   2880
            Width           =   195
         End
         Begin VB.CheckBox ChkAv 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            Caption         =   "IDFA"
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
            Index           =   5
            Left            =   1440
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   2460
            Width           =   1065
         End
         Begin VB.CheckBox ChkAv 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            Caption         =   "IDFA"
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
            Index           =   6
            Left            =   1440
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   2670
            Width           =   1005
         End
         Begin VB.TextBox txtConsertado 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
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
            Index           =   5
            Left            =   1260
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   2460
            Width           =   195
         End
         Begin VB.TextBox txtConsertado 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
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
            Index           =   6
            Left            =   1260
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   2670
            Width           =   195
         End
         Begin VB.CheckBox ChkAv 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            Caption         =   "IDFA"
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
            Index           =   4
            Left            =   1440
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   2250
            Width           =   1095
         End
         Begin VB.CheckBox ChkAv 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            Caption         =   "IDFA"
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
            Index           =   2
            Left            =   210
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   2670
            Width           =   1080
         End
         Begin VB.CheckBox ChkAv 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            Caption         =   "IDFA"
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
            Index           =   0
            Left            =   210
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   2250
            Width           =   1050
         End
         Begin VB.CheckBox ChkAv 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            Caption         =   "IDFA"
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
            Index           =   1
            Left            =   210
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   2460
            Width           =   1080
         End
         Begin VB.CheckBox ChkAv 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            Caption         =   "IDFA"
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
            Index           =   3
            Left            =   210
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   2880
            Width           =   1080
         End
         Begin VB.TextBox txtConsertado 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
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
            Index           =   0
            Left            =   60
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   2250
            Width           =   165
         End
         Begin VB.TextBox txtConsertado 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
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
            Index           =   1
            Left            =   60
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   2460
            Width           =   165
         End
         Begin VB.TextBox txtConsertado 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
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
            Index           =   2
            Left            =   60
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   2670
            Width           =   165
         End
         Begin VB.TextBox txtConsertado 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
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
            Index           =   3
            Left            =   60
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   2880
            Width           =   165
         End
         Begin VB.TextBox txtConsertado 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
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
            Index           =   4
            Left            =   1260
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   2280
            Width           =   195
         End
         Begin VB.TextBox txtItem 
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
            Height          =   285
            Left            =   3270
            MaxLength       =   3
            TabIndex        =   4
            Top             =   120
            Width           =   285
         End
         Begin VB.TextBox txtQtde 
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
            Height          =   285
            Left            =   450
            TabIndex        =   5
            Top             =   390
            Width           =   555
         End
         Begin VB.TextBox txtVolume 
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
            Height          =   270
            Left            =   960
            TabIndex        =   7
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox txtBruto 
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
            Height          =   270
            Left            =   2220
            TabIndex        =   8
            Top             =   660
            Width           =   1305
         End
         Begin VB.TextBox txtMarca 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   960
            MaxLength       =   50
            TabIndex        =   10
            Top             =   1200
            Width           =   2565
         End
         Begin VB.TextBox txtMercadoria 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   960
            MaxLength       =   50
            TabIndex        =   11
            Top             =   1470
            Width           =   2565
         End
         Begin MSDataListLib.DataCombo dcEmbalagem 
            Bindings        =   "DesovaPier.frx":0031
            Height          =   285
            Left            =   1380
            TabIndex        =   6
            Top             =   390
            Width           =   2175
            _ExtentX        =   3836
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
         Begin MSDataListLib.DataCombo dcLote 
            Bindings        =   "DesovaPier.frx":004E
            Height          =   285
            Left            =   450
            TabIndex        =   3
            Top             =   120
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   503
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
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSAdodcLib.Adodc AdodcLote 
            Height          =   330
            Left            =   4020
            Top             =   -60
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
         Begin MSAdodcLib.Adodc AdodcEmbalagem 
            Height          =   330
            Left            =   4140
            Top             =   300
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
         Begin MSAdodcLib.Adodc AdodcGenero 
            Height          =   330
            Left            =   3270
            Top             =   120
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
         Begin MSAdodcLib.Adodc AdodcDesovados 
            Height          =   330
            Left            =   4500
            Top             =   4260
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
         Begin VB.CheckBox ckIDFA 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            Caption         =   "FALTA"
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
            Left            =   1530
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1740
            Width           =   645
         End
         Begin VB.CheckBox ckAcrescimo 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            Caption         =   "ACRESC."
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
            Left            =   810
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1770
            Width           =   765
         End
         Begin VB.CheckBox ckAvaria 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            Caption         =   "AVARIA-Qtd:"
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
            Left            =   2160
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1740
            Width           =   1005
         End
         Begin VB.CheckBox ckMadeira 
            Appearance      =   0  'Flat
            BackColor       =   &H00BA9F8D&
            Caption         =   "MADEIRA"
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
            Height          =   210
            Left            =   30
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1770
            Width           =   795
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridDesovados 
            Bindings        =   "DesovaPier.frx":0066
            Height          =   885
            Left            =   60
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   3120
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   1561
            _Version        =   393216
            Rows            =   4
            Cols            =   6
            RowHeightMin    =   240
            GridLines       =   0
            GridLinesFixed  =   0
            SelectionMode   =   1
            Appearance      =   0
            FormatString    =   "   |<Item  |^Qtde|<Embalagem   |<Marca    |<Mercadoria "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
         End
         Begin VB.CommandButton cmdDesovaItem 
            Caption         =   "[F1] - ADICIONA"
            Height          =   225
            Left            =   90
            TabIndex        =   109
            Top             =   4020
            Width           =   1665
         End
         Begin VB.CommandButton cmdRemoveDesova 
            Caption         =   "[F2] - REMOVE"
            Height          =   225
            Left            =   1830
            TabIndex        =   110
            Top             =   4020
            Width           =   1635
         End
         Begin VB.CommandButton cmdMarcantes 
            Caption         =   "[F3] - MARCANTES"
            Height          =   225
            Left            =   120
            TabIndex        =   111
            Top             =   4260
            Width           =   1665
         End
         Begin VB.CommandButton cmdResumo 
            Caption         =   "[F4] - RESUMO/OBS"
            Height          =   225
            Left            =   1800
            TabIndex        =   112
            Top             =   4260
            Width           =   1665
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "[F3]MARCANTES     [F4]RESUMO/OBS"
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
            Index           =   1
            Left            =   210
            TabIndex        =   92
            Top             =   4230
            Width           =   2925
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   6
            Left            =   30
            TabIndex        =   87
            Top             =   930
            Width           =   2325
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "PB. APURADO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Index           =   5
            Left            =   1470
            TabIndex        =   86
            Top             =   900
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "PESO Av."
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   4
            Left            =   2100
            TabIndex        =   84
            Top             =   2010
            Width           =   675
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "AVARIAS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   0
            Left            =   30
            TabIndex        =   75
            Top             =   2010
            Width           =   2115
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "[F1]ADICIONA     [F2]REMOVE "
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
            Index           =   0
            Left            =   510
            TabIndex        =   46
            Top             =   4050
            Width           =   2430
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MERCAD."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   26
            Left            =   30
            TabIndex        =   45
            Top             =   1470
            Width           =   945
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "MARCA "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   27
            Left            =   30
            TabIndex        =   44
            Top             =   1200
            Width           =   945
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "VOLUME  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   29
            Left            =   30
            TabIndex        =   43
            Top             =   660
            Width           =   945
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "EMB."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   30
            Left            =   990
            TabIndex        =   42
            Top             =   390
            Width           =   405
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " ITEM  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   31
            Left            =   2850
            TabIndex        =   41
            Top             =   120
            Width           =   435
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LOTE  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   32
            Left            =   30
            TabIndex        =   40
            Top             =   120
            Width           =   435
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "QTDE  "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   33
            Left            =   30
            TabIndex        =   39
            Top             =   390
            Width           =   435
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "P. BRUTO "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   34
            Left            =   1440
            TabIndex        =   38
            Top             =   660
            Width           =   795
         End
      End
   End
End
Attribute VB_Name = "DesovaPier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Carrega_Marcantes(AutonumBL As Long, Item As Integer)

Sql = "SELECT RIGHT('00000000' + RTRIM(CONVERT(CHAR,AUTONUM)),9) + ' Qtde:' + LTRIM(RTRIM(CONVERT(CHAR,ISNULL(VOLUMES,0))) + ' Local: ' + ISNULL(ETQ_PRATELEIRA,'')) as DISPLAY,AUTONUM    "
Sql = Sql & " FROM "
Sql = Sql & " SGIPA..TB_MARCANTES "
Sql = Sql & " WHERE CNTR=" & Me.MskCntr.Tag
Sql = Sql & " AND BL=" & AutonumBL
Sql = Sql & " AND ITEM=" & Item
Sql = Sql & " AND VOLUMES>0 "
Sql = Sql & " ORDER BY AUTONUM"


Me.AdodcMarcante.ConnectionString = ConnectionString
Me.AdodcMarcante.RecordSource = Sql
Me.AdodcMarcante.Refresh
Set Me.AdodcMarcante.Recordset.ActiveConnection = Nothing

End Sub

Private Sub Carrega_Avarias_Col()
For i = 0 To 11
    ChkAv(i).Value = 0
    ChkAv(i).Visible = False
    ChkAv(i).Tag = 0
    ChkAv(i).Caption = ""
    txtConsertado(i).Visible = False
    txtConsertado(i).Enabled = False
    txtConsertado(i).BackColor = vbRed
Next i

Set tb1 = New ADODB.Recordset
Dim Sql As String
Sql = " SELECT "
Sql = Sql & " DESCR_AVARIA, "
Sql = Sql & " LOCAL, "
Sql = Sql & " TIPO, "
Sql = Sql & " COMPLEMENTO, "
Sql = Sql & " PROVIDENCIA, "
Sql = Sql & " PROVIDENCIA_C, "
Sql = Sql & " ID_AVARIA_COL "
Sql = Sql & " FROM SGIPA..TB_CONV_AVARIAS "
Sql = Sql & " ORDER BY DESCR_AVARIA"

tb1.Open Sql, db, 1, 1


i = 0
While Not tb1.EOF
    If i < 12 Then
    
        Me.ChkAv(i).Value = 0
        Me.ChkAv(i).Visible = True
        Me.ChkAv(i).Tag = Nnull(tb1!id_avaria_col, 0)
        Me.ChkAv(i).Caption = Nnull(tb1!descr_avaria, 1)
        Me.txtConsertado(i).Visible = True
        Me.txtConsertado(i).Enabled = IIf(Nnull(tb1!providencia_c, 1) = "", False, True)
        
    End If
    i = i + 1
    tb1.MoveNext
Wend

End Sub




Sub cmdAdicionaMc_Click()
If Me.TXTMARCANTE = "" Then
    Col_Mensagem = "Informe o marcante"
    Col_Imagem = 1
    frmMensagem2.Show 1
    Exit Sub
End If


If Trim(Me.txtPrateleira) = "" Then
    If db.Execute("SELECT FLAG_POSICIONA_CARGA_DESOVA FROM OPERADOR..TB_PATIOS WHERE AUTONUM=" & Col_Patio)(0) = 1 Then
        Col_Mensagem = "Informe o LOCAL de armazenagem"
        Col_Imagem = 1
        frmMensagem2.Show 1
        Exit Sub
    End If
End If

'Verifica se o Marcante é valido
Sql = "SELECT AUTONUM,AUTONUM_CARGA,DT_ASSOCIACAO,LOTE_IMPRESSO,ID_CONTEINER_IMPRESSO,ISNULL(FLAG_RECONHECIMENTO,0) AS FLAG_RECONHECIMENTO FROM SGIPA..TB_MARCANTES WHERE "
Sql = Sql & " AUTONUM= " & Replace(Val(Me.TXTMARCANTE), ",", "") & ""

Set tb1 = New ADODB.Recordset
tb1.Open Sql, db, 1, 1
If tb1.EOF Then
    Col_Mensagem = "Nr de marcante inválido"
    frmMensagem2.Show vbModal
    Me.TXTMARCANTE = ""
    Exit Sub
Else
    If Nnull(tb1!DT_ASSOCIACAO, 1) <> "" Then
        Col_Mensagem = "Nr de marcante já utilizado"
        frmMensagem2.Show vbModal
        Me.TXTMARCANTE = ""
        Exit Sub
    End If
    'If Nnull(tb1!FLAG_RECONHECIMENTO, 0) = 1 Then
    '    Col_Mensagem = "Marcante para carga em reconhecimento"
     '   frmMensagem2.Show vbModal
     '   Me.TXTMARCANTE = ""
     '   Exit Sub
    'End If
    If Me.ckReconhecimento.Value = 1 And Nnull(tb1!FLAG_RECONHECIMENTO, 0) = 0 Then
        Col_Mensagem = "Este marcante NÃO é de reconhecimento"
        frmMensagem2.Show vbModal
        Me.TXTMARCANTE = ""
        Exit Sub
    End If
    
    If Me.ckReconhecimento.Value = 0 And Nnull(tb1!FLAG_RECONHECIMENTO, 0) = 1 Then
        Col_Mensagem = "Este marcante é de reconhecimento"
        frmMensagem2.Show vbModal
        Me.TXTMARCANTE = ""
        Exit Sub
    End If
    
End If


If Nnull(tb1!LOTE_IMPRESSO, 0) <> 0 Then
    If Val(Nnull(tb1!LOTE_IMPRESSO, 0)) <> Val(Nnull(Me.dcLote.BoundText, 0)) Then
        Col_Mensagem = "Este marcante é de outro lote"
        frmMensagem2.Show vbModal
        Me.TXTMARCANTE = ""
        Exit Sub
    End If
End If


If Trim(Nnull(tb1!ID_CONTEINER_IMPRESSO, 1)) <> Me.MskCntr.text Then
    Col_Mensagem = "Este marcante é de outro conteiner"
    frmMensagem2.Show vbModal
    Me.TXTMARCANTE = ""
End If
    
Dim QualItem As Integer
If Nnull(Me.txtItem, 0) <> 0 Then
    QualItem = Nnull(Me.txtItem, 0)
End If
    
If QualItem = 0 Then
    QualItem = Val(Nnull(Me.GridDesovados.TextMatrix(Me.GridDesovados.Row, 1), 0))
End If
    
 If QualItem = 0 Then
    Col_Mensagem = "Escolha o item do bl"
    Col_Imagem = 1
    frmMensagem2.Show 1
    Exit Sub
End If
    
If Nnull(Me.txtQtdeMarcante, 0) = 0 Then
    Col_Mensagem = "Informe a quantidade associada ao marcante"
    Col_Imagem = 1
    frmMensagem2.Show 1
    Exit Sub
End If

If Val(Me.txtQtdeMarcante) > Val(Me.txtQtde) Then
    Col_Mensagem = "Quantidade inválida"
    Col_Imagem = 1
    frmMensagem2.Show 1
    Me.txtQtdeMarcante = ""
    Exit Sub
End If


'SE INFORMOU A ETIQUETA DE POSICAO
Dim AutonumArm$
Dim YardCs As String
AutonumArm = 0
YardCs = ""





If Me.txtPrateleira.text <> "" Then
    
    If Len(Me.txtPrateleira.text) = 9 Then

        Set TBP = New ADODB.Recordset
        
        If Col_Patio = 1 Or Col_Patio = 7 Then
            AutonumArm$ = Left(Me.txtPrateleira.text, 2)
        Else
            AutonumArm$ = Nnull(db.Execute("SELECT max(AUTONUM) FROM SGIPA..TB_ARMAZENS_IPA WHERE PATIO=" & Col_Patio & " AND AUTONUM_ARM_PATIO=" & Val(Left(Me.txtPrateleira.text, 2)))(0), 0)
        End If
        
    
        
        YardCs = UCase(Replace(Mid$(Me.txtPrateleira.text, 3), "-", ""))

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
    ElseIf Len(Me.txtPrateleira.text) = 10 Then

        Set TBP = New ADODB.Recordset
    
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
    ElseIf Len(Me.txtPrateleira.text) >= 4 And InStr(1, "-", Me.txtPrateleira.text) = 0 Then
        
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
            tb1.Close
            tb1.Open "SELECT AUTONUM FROM sgipa..tb_yard_cs where armazem=" & Val(xArm) & " and yard='" & Left$(YardCs, 3) & "01" & Right$(YardCs, 1) & "'", db, 1, 1
            If Not tb1.EOF Then
                YardCs = Left$(YardCs, 3) & "01" & Right$(YardCs, 1)
            Else
                YardCs = ""
            End If
        End If
    End If
    


    If YardCs = "" Then
        Col_Mensagem = "Local inválido"
        Col_Imagem = 1
        frmMensagem2.Show 1
        Me.txtPrateleira.text = ""
        Exit Sub
    Else
        
        Set tb1 = New ADODB.Recordset
        tb1.Open "SELECT VALIDA,FLAG_BLOQUEIO FROM SGIPA..TB_YARD_CS WHERE ARMAZEM=" & Val(xArm) & " AND YARD='" & YardCs & "'", db, 1, 1
        If Nnull(tb1!FLAG_BLOQUEIO, 0) = 1 Then
            Col_Mensagem = "Local bloqueado"
            Col_Imagem = 1
            frmMensagem2.Show 1
            Me.txtPrateleira.text = ""
        End If
    
        If Nnull(tb1!Valida, 0) = 1 Then
            If db.Execute("SELECT FLAG_CS_PERMITE_CONC FROM SGIPA..DTE_TB_PARAMETROS")(0) = 0 Then
                Sql = "SELECT MARCANTE FROM OPERADOR..VW_INVENT_ARMAZEM WHERE AUTONUM_ARMAZEM=" & Val(xArm) & " AND POSICAO='" & YardCs & "'"
                Set tb1 = New ADODB.Recordset
                tb1.Open Sql, db, 1, 1
                If Not tb1.EOF Then
                    Col_Mensagem = "Local ocupado por outra carga"
                    Col_Imagem = 1
                    frmMensagem2.Show 1
                    Me.txtPrateleira.text = ""
                End If
                tb1.Close
            End If
            
        End If
    
    End If

End If

Sql = " UPDATE SGIPA..TB_MARCANTES SET CNTR=" & Me.MskCntr.Tag & ", BL=" & Nnull(Me.dcLote.BoundText, 0) & ",ITEM=" & Nnull(Me.txtItem, 0) & " ,AUTONUM_CARGA=" & Nnull(AutonumCs, 0) & ",VOLUMES = " & Val(Me.txtQtdeMarcante) & "  ,DT_ASSOCIACAO=GETDATE(),ETQ_PRATELEIRA='" & Me.txtPrateleira.text & "', ARMAZEM=" & Nnull(xArm, 0) & ",YARD='" & YardCs & "' WHERE AUTONUM=" & Val(Me.TXTMARCANTE)
db.Execute Sql

Call Carrega_Marcantes(Nnull(Me.dcLote.BoundText, 0), Nnull(Me.txtItem, 0))

Me.TXTMARCANTE = ""
Me.txtQtdeMarcante = ""
Me.txtPrateleira = ""
Me.TXTMARCANTE.SetFocus


End Sub

Sub cmdCarregaItem_Click()
If Me.ListView1.SelectedItem > 0 Then
    Me.dcLote.BoundText = Me.ListView1.ListItems(Val(Me.ListView1.SelectedItem)).SubItems(2)
    Me.txtItem = Me.ListView1.ListItems(Val(Me.ListView1.SelectedItem)).SubItems(4)
    Call Valida_Item(Nnull(Me.dcLote.BoundText, 0), Nnull(Me.txtItem, 0))
End If
Me.FrameDesova.Visible = True
Me.frmResumo.Visible = False
End Sub

Sub cmdDesovaItem_Click()
  Dim Avaria_Cs As Long

If Nnull(Me.dcLote.BoundText, 0) = 0 Then
    Col_Mensagem = "Lote não informado"
    
    frmMensagem2.Show vbModal
    Exit Sub
End If

If Nnull(Me.txtItem, 0) = 0 Then
    Col_Mensagem = "Item não informado"
    
    frmMensagem2.Show vbModal
    Exit Sub
End If

If Me.ckReconhecimento.Value = 0 Then

    If Nnull(Me.txtBruto, 0) = 0 Then
        Col_Mensagem = "Peso Bruto não informado"
        
        frmMensagem2.Show vbModal
        Exit Sub
    End If
    
    If Nnull(Me.dcEmbalagem, 0) = 0 Then
        Col_Mensagem = "Embalagem não informada"
        
        frmMensagem2.Show vbModal
        Exit Sub
    End If
    
    If Nnull(Me.txtQtde, 0) = 0 Then
        Col_Mensagem = "Quantidade não informada"
        
        frmMensagem2.Show vbModal
        Exit Sub
    End If
    
    If Nnull(Me.txtQtdeAv, 1) = "" Then
        Col_Mensagem = "Quantidade avariada não informada"
        
        frmMensagem2.Show vbModal
        Exit Sub
    End If
    
    
    If Nnull(Me.txtPesoAvariado, 1) = "" Then
        Col_Mensagem = "Peso avariado não informado"
        
        frmMensagem2.Show vbModal
        Exit Sub
    End If
    
    
'    If Nnull(Me.txtBrutoAP, 0) = 0 And Me.txtBrutoAP.Enabled = True Then
'        Col_Mensagem = "BL DTA - informe Peso Bruto apurado"
'        frmMensagem2.Show vbModal
'        Exit Sub
'    End If
    
    '01/08/2025
    If Nnull(Me.txtBrutoAP, 0) = 0 And Me.txtBrutoAP.Enabled = True Then
        Col_Mensagem = "BL - informe Peso Bruto apurado"
        frmMensagem2.Show vbModal
        Exit Sub
    End If
    
    'Soma das quantidades de marcantes deve ser igual a quantidade desovada
    Sql = "SELECT SUM(VOLUMES) FROM SGIPA..TB_MARCANTES WHERE BL=" & Me.dcLote.BoundText & " AND ITEM=" & Val(Me.txtItem)
    Dim Qtde_Marcante As Long
    Qtde_Marcante = Nnull(db.Execute(Sql)(0), 0)
    
    If Val(Me.txtQtde) <> Qtde_Marcante Then
        Col_Mensagem = "Quantidade informada em marcantes diferente da quantidade desovada"
        Col_Imagem = 1
        frmMensagem2.Show vbModal
        Exit Sub
    End If
    
    
    If Nnull(Me.txtMarca, 1) = "" Then
        Col_Mensagem = "Marca não informada"
        
        frmMensagem2.Show vbModal
        Exit Sub
    End If
    
    If Nnull(Me.txtMercadoria, 1) = "" Then
        Col_Mensagem = "Mercadoria não informada"
        
        frmMensagem2.Show vbModal
        Exit Sub
    End If
    
Else

    If Nnull(Me.txtQtde, 0) = 0 Then
        Col_Mensagem = "Quantidade não informada"
        
        frmMensagem2.Show vbModal
        Exit Sub
    End If
    
    If Nnull(Me.txtQtdeAv, 1) = "" Then
        Col_Mensagem = "Quantidade avariada não informada"
        
        frmMensagem2.Show vbModal
        Exit Sub
    End If
    
    If Nnull(Me.txtPesoAvariado, 1) = "" Then
        Col_Mensagem = "Para carga aguardando reconhecimento, não informe a Peso avariado"
        
        frmMensagem2.Show vbModal
        Exit Sub
    End If

End If
       
                             
Dim Amr As Long
Sql = "select autonum from sgipa..tb_amr_cntr_bl where cntr=" & Me.MskCntr.Tag & " and bl=" & Me.dcLote.BoundText
Amr = Nnull(db.Execute(Sql)(0), 0)

Dim QtdeDesovada As Long
QtdeDesovada = Val(Me.txtQtde) + Nnull(db.Execute("SELECT SUM(QUANTIDADE) FROM SGIPA..TB_CARGA_SOLTA WHERE AMR=" & Amr & " AND ITEM=" & Me.txtItem)(0), 0)
    
    
If Val(Nnull(Me.txtQtde.Tag, 0)) > Val(QtdeDesovada) And Me.ckIDFA.Value = 0 Then
    Col_Mensagem = "Atenção, consta falta de " & Val(Nnull(Me.txtQtde.Tag, 0)) - Val(QtdeDesovada) & " volume(s) "
    Col_Imagem = 3
    frmMensagem2.Show 1
    If Col_Cancelado Then Exit Sub
End If

If Val(Nnull(Me.txtQtde.Tag, 0)) < Val(QtdeDesovada) And Me.ckAcrescimo.Value = 0 Then
    Col_Mensagem = "Atenção, consta acréscimo de " & Abs(Val(Nnull(Me.txtQtde.Tag, 0)) - Val(QtdeDesovada)) & " volume(s) "
    Col_Imagem = 3
    frmMensagem2.Show 1
    
    If Col_Cancelado Then Exit Sub
    'Me.ckAcrescimo.Value = 1
End If
       
        
Sql = "INSERT INTO SGIPA..TB_CARGA_SOLTA (AMR,BL,CNTR,VIAGEM,"
Sql = Sql & " ITEM,"
Sql = Sql & " QUANTIDADE,"
Sql = Sql & " QUANTIDADE_REAL, "
Sql = Sql & " EMBALAGEM,"
Sql = Sql & " Mercadoria,"
Sql = Sql & " Marca,"
Sql = Sql & " Genero,"
Sql = Sql & " PESO_BRUTO,"
Sql = Sql & " PESO_AP_DTA,"
Sql = Sql & " Volume, "
Sql = Sql & " FLAG_TERMINAL,FLAG_HISTORICO,FLAG_CHEGADA_TOTAL,QUANTIDADE_SAIDA,DT,"
Sql = Sql & " FLAG_AVARIADO,"
Sql = Sql & " FLAG_MADEIRA,"
Sql = Sql & " FLAG_ACRESCIMO,"
Sql = Sql & " IDFA,"
Sql = Sql & " PATIO) VALUES (" & Amr & ","
Sql = Sql & Me.dcLote.BoundText & " ," & Me.MskCntr.Tag & " , '" & Me.txtViagem & "', "
Sql = Sql & Val(Me.txtItem) & ","  'ITEM
Sql = Sql & Val(Me.txtQtde) & "," 'QUANTIDADE
Sql = Sql & Val(Me.txtQtde) & "," 'QUANTIDADE_REAL
Sql = Sql & Me.dcEmbalagem.BoundText & ","  'EMBALAGEM
Sql = Sql & "'" & UCase(Me.txtMercadoria) & "',"  'MERCADORIA
Sql = Sql & "'" & UCase(Me.txtMarca) & "'," 'MARCA
'Sql = Sql & Nnull(Me.dcGenero.BoundText, 0) & "," 'GENERO
Sql = Sql & "0,"
Sql = Sql & Replace(Me.txtBruto, ",", ".") & ","  'PESO_BRUTO
If Me.txtBrutoAP.text <> "" Then
    Sql = Sql & Replace(Me.txtBrutoAP, ",", ".") & ","    'PESO_BRUTO DTA
Else
    Sql = Sql & "0,"
End If
Sql = Sql & Nnull(Replace(Me.txtVolume, ",", "."), 0) & ","  'VOLUME
Sql = Sql & " 1,0,1,0,'" & Me.txtdt & "', "
Sql = Sql & Me.ckAvaria.Value & ","
Sql = Sql & Me.ckMadeira.Value & ","
Sql = Sql & Me.ckAcrescimo.Value & ","
Sql = Sql & Me.ckIDFA.Value & ","
Sql = Sql & db.Execute("select patio from sgipa..tb_bl where autonum=" & Me.dcLote.BoundText)(0)
Sql = Sql & ")"
db.Execute Sql
    
Dim AutonumCs As Long
AutonumCs = db.Execute("SELECT ISNULL(MAX(AUTONUM),0) FROM SGIPA..TB_CARGA_SOLTA WHERE AMR=" & Amr & " AND ITEM=" & Val(txtItem))(0)
    
            
For i = 0 To 11
    If Me.ChkAv(i).Value = 1 Then
    
    
        Sql = " INSERT INTO SGIPA..SEQ_AVARIAS_CS DEFAULT VALUES "
        db.Execute Sql
    
        Avaria_Cs = db.Execute("SELECT IDENT_CURRENT('SEQ_AVARIAS_CS')")(0)
    
        Sql = "INSERT INTO SGIPA..TB_AVARIAS_CS (AUTONUM,"
        Sql = Sql & " AUTONUMCS, "
        Sql = Sql & " LOCAL,"
        Sql = Sql & " TIPO,"
        Sql = Sql & " COMPLEMENTO,"
        Sql = Sql & " PROVIDENCIA,"
        Sql = Sql & " DT_AVARIA, "
        Sql = Sql & " ID_AVARIA_COL,QUANTIDADE_AVARIADA,PESO_AVARIADO) "
        Sql = Sql & " SELECT " & Avaria_Cs & ","
        Sql = Sql & AutonumCs & ", LOCAL, TIPO,COMPLEMENTO,"
        If Me.txtConsertado(i).BackColor = vbRed Or Me.txtConsertado(i).BackColor = vbMagenta Then
            Sql = Sql & "PROVIDENCIA,"
        Else
            Sql = Sql & "PROVIDENCIA_C,"
        End If
        Sql = Sql & "GETDATE(),"
        Sql = Sql & "ID_AVARIA_COL ," & Val(Me.txtQtdeAv.text) & "," & Val(Me.txtPesoAvariado)
        Sql = Sql & " FROM SGIPA..TB_CONV_AVARIAS WHERE ID_AVARIA_COL=" & Me.ChkAv(i).Tag
        db.Execute Sql
        
        If Me.txtConsertado(i).BackColor = vbMagenta Then
        
            Sql = " INSERT INTO SGIPA..SEQ_AVARIAS_CS DEFAULT VALUES "
            db.Execute Sql
            Avaria_Cs = db.Execute("SELECT IDENT_CURRENT('SEQ_AVARIAS_CS')")(0)
            Sql = "INSERT INTO SGIPA..TB_AVARIAS_CS (AUTONUM,"
            Sql = Sql & " AUTONUMCS, "
            Sql = Sql & " LOCAL,"
            Sql = Sql & " TIPO,"
            Sql = Sql & " COMPLEMENTO,"
            Sql = Sql & " PROVIDENCIA,"
            Sql = Sql & " DT_AVARIA, "
            Sql = Sql & " ID_AVARIA_COL,QUANTIDADE_AVARIADA,PESO_AVARIADO) "
            Sql = Sql & " SELECT " & Avaria_Cs & ","
            Sql = Sql & AutonumCs & ", LOCAL, TIPO,3,"
            Sql = Sql & "PROVIDENCIA,"
            Sql = Sql & "GETDATE(),"
            Sql = Sql & "ID_AVARIA_COL ," & Val(Me.txtQtdeAv.text) & "," & Val(Me.txtPesoAvariado)
            Sql = Sql & " FROM SGIPA..TB_CONV_AVARIAS WHERE ID_AVARIA_COL=" & Me.ChkAv(i).Tag
            db.Execute Sql
        End If
        
    End If
Next i
    
Sql = " UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CARGA=" & AutonumCs & " WHERE CNTR=" & Me.MskCntr.Tag & " AND BL=" & Me.dcLote.BoundText & " AND ITEM=" & Val(Me.txtItem)
db.Execute Sql
    
'Marcantes já com posicao
Sql = "SELECT AUTONUM, ARMAZEM, YARD, VOLUMES FROM SGIPA..TB_MARCANTES WHERE AUTONUM_CARGA=" & AutonumCs & " and isnull(ARMAZEM,0)>0"
Set TBP = New ADODB.Recordset
TBP.Open Sql, db, 1, 1
While Not TBP.EOF

    Sql = "INSERT INTO SGIPA..TB_CARGA_SOLTA_YARD (AUTONUM_CS,ARMAZEM,YARD,ORIGEM,QUANTIDADE,MOTIVO) VALUES "
    Sql = Sql & "(" & AutonumCs & "," & Nnull(TBP!ARMAZEM, 0) & ",'" & Nnull(TBP!Yard, 1) & "','I'," & Nnull(TBP!VOLUMES, 0) & ",7)"
    db.Execute Sql

    Dim Autonum_Cs_Yard As Long
    'Autonum_Cs_Yard = db.Execute("SELECT IDENT_CURRENT('TB_CARGA_SOLTA_YARD')")(0)
    Autonum_Cs_Yard = db.Execute("SELECT ISNULL(MAX(AUTONUM),0) FROM TB_CARGA_SOLTA_YARD WHERE AUTONUM_CS=" & AutonumCs)(0)
    
    Sql = "UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CS_YARD=" & Autonum_Cs_Yard & " WHERE AUTONUM=" & TBP!AUTONUM
    db.Execute Sql

    TBP.MoveNext
Wend
TBP.Close

    
Sql = "UPDATE SGIPA..TB_CARGA_CNTR SET FLAG_RECONHECIMENTO_CARGA=" & Me.ckReconhecimento.Value & " WHERE   "
Sql = Sql & " BL=" & Me.dcLote.BoundText
Sql = Sql & " AND ITEM=" & Me.txtItem
Sql = Sql & " AND ID_CONTEINER='" & Me.MskCntr & "'"
db.Execute Sql
    
Call COL_OBS_Coletor(Me.dcLote.BoundText)
            
Col_Mensagem = "ITEM DESOVADO"
Col_Imagem = 2
frmMensagem2.Show 1
    
Dim LoteAnt As Long
LoteAnt = Nnull(Me.dcLote.BoundText, 0)
Carrega_Lotes
Me.dcLote.BoundText = LoteAnt
Carrega_ItensDesovados (Nnull(Me.dcLote.BoundText, 0))


'PARAMETRO
If db.Execute("SELECT FLAG_INVENT_CS_POS_DESOVA FROM DTE_TB_PARAMETROS")(0) = 1 Then
    Load Inventario_CS
    Inventario_CS.TXTLOTE = Me.dcLote.BoundText
    Call Inventario_CS.Carrega_Dados_Lote(Trim(str(Me.dcLote.BoundText)), 0)
    Inventario_CS.cbItem.BoundText = AutonumCs
    Inventario_CS.Show 1
End If
    


Limpa_Item
    
End Sub

Sub cmdInicio_Click()

If Nnull(Me.MskCntr.Tag, 0) = 0 Then
    Col_Mensagem = "Contêiner inválido"
    
    frmMensagem2.Show vbModal
    Exit Sub
End If


Dim IdSol As Integer
IdSol = 0
For i = 1 To 5
    If Me.Option4(i).Value = True Then IdSol = i
Next i


'    If IdSol = 0 Then
'        Col_Mensagem = "Informe o tipo de solicitante"
'        frmMensagem2.Show vbModal
'        Exit Sub
'    End If
'
'
'    If Val(Me.DBCombo10.BoundText) = 0 Then
'        Col_Mensagem = "Informe o solicitante"
'        frmMensagem2.Show vbModal
'        Exit Sub
'    End If


If Me.txtInicio = "" Then

    Dim Agora As Date
    Agora = db.Execute("SELECT GETDATE()")(0)
    Me.txtInicio = Format(Agora, "dd/mm/yyyy hh:mm")
    Sql = "UPDATE SGIPA..TB_CNTR_BL SET CONFERENTE_DESOVA='" & Col_Login & "',FLAG_DESOVA_IMPORTADOR=" & IdSol & ",AUTONUM_SOLICITOU_DESOVA=" & Val(Nnull(Me.DBCombo10.BoundText, 0)) & ",DT_INICIO_DESOVA=GETDATE(), FLAG_DESOVA_MANUAL=" & IIf(Me.Option3(0).Value = True, 1, 0) & " "
    Sql = Sql & " Where autonum = " & Me.MskCntr.Tag
    
    db.Execute Sql

Else
    Col_Mensagem = "Início de operação de desova já informado"
    
    frmMensagem2.Show vbModal
    Exit Sub
End If
    
    

End Sub

Sub cmdMarcantes_Click()
Me.frmMarcantes.Visible = True
Call Carrega_Marcantes(Nnull(Me.dcLote.BoundText, 0), Nnull(Me.txtItem, 0))
Me.TXTMARCANTE.SetFocus
End Sub

Sub cmdRemoveDesova_Click()

If Val(Me.GridDesovados.TextMatrix(Me.GridDesovados.Row, Me.GridDesovados.Cols - 1)) > 0 Then
    Col_Mensagem = "Deseja excluir o item desovado ? "
    Col_Imagem = 3
    frmMensagem2.Show 1
    
    If Col_Cancelado Then Exit Sub

    
    Sql = " DELETE FROM SGIPA..TB_CARGA_SOLTA WHERE AUTONUM=" & Val(Me.GridDesovados.TextMatrix(Me.GridDesovados.Row, Me.GridDesovados.Cols - 1))
    db.Execute Sql
    
    Sql = "DELETE FROM SGIPA..TB_AVARIAS_CS WHERE AUTONUMCS=" & Val(Me.GridDesovados.TextMatrix(Me.GridDesovados.Row, Me.GridDesovados.Cols - 1))
    db.Execute Sql
    
    Sql = "UPDATE SGIPA..TB_MARCANTES SET DT_ASSOCIACAO=NULL, AUTONUM_CARGA=0,CNTR=0, ITEM=0, BL=0, ETQ_PRATELEIRA='' , ARMAZEM=0, YARD='' WHERE AUTONUM_CARGA=" & Val(Me.GridDesovados.TextMatrix(Me.GridDesovados.Row, Me.GridDesovados.Cols - 1)) & " AND AUTONUM_CARGA>0"
    db.Execute Sql
    
    Limpa_Item
    
    Carrega_ItensDesovados (Nnull(Me.dcLote.BoundText, 0))
    
Else
    Col_Mensagem = "Nenhum item selecionado"
    Col_Imagem = 1
    frmMensagem2.Show 1

End If
    
End Sub

Sub cmdRemoveMc_Click()
If Nnull(Me.DlMarcante.BoundText, 0) = 0 Then
    Col_Mensagem = "Escolha o marcante a ser removido"
    Col_Imagem = 1
    frmMensagem2.Show 1
    Exit Sub
End If

Sql = "UPDATE SGIPA..TB_MARCANTES SET VOLUMES=0, DT_ASSOCIACAO=NULL,BL=0,CNTR=0,ITEM=0, ETQ_PRATELEIRA='', ARMAZEM=0, YARD='' WHERE AUTONUM=" & Val(Nnull(Me.DlMarcante.BoundText, 0))
db.Execute Sql

Call Carrega_Marcantes(Nnull(Me.dcLote.BoundText, 0), Nnull(Me.txtItem, 0))
Me.TXTMARCANTE = ""
Me.txtQtdeMarcante = ""
Me.txtPrateleira = ""
End Sub

Private Sub cmdResumo_Click()

Me.FrameDesova.Visible = False
Me.frmResumo.Visible = True
Carrega_Resumo
    
End Sub

Private Sub cmdSair_Click()
If IsLoaded("frmMensagem2") = True Then
    Unload frmMensagem2
Else
    Unload Me
End If

End Sub

Sub cmdTermino_Click()
If Nnull(Me.MskCntr.Tag, 0) = 0 Then
    Col_Mensagem = "Contêiner inválido"
    Col_Imagem = 1
    frmMensagem2.Show 1
    Exit Sub
End If

If Me.txtInicio = "" Then
    Col_Mensagem = "Início de operação de desova ainda não informado"
    Col_Imagem = 1
    frmMensagem2.Show 1
    Exit Sub
End If


'Parametro pra ver ser precisa tudo estar com marcante
Set tbp1 = New ADODB.Recordset
tbp1.Open "Select FLAG_MARCANTES_FINALIZA_DESOVA from sgipa..tb_parametros", db, 1, 1
If Not tbp1.EOF Then
    If Nnull(tbp1!FLAG_MARCANTES_FINALIZA_DESOVA, 0) = 1 Then
        'Compara se todos os itens desovados do cntr tem marcante associado
        Sql = "SELECT"
        Sql = Sql & " c.BL , c.Item"
        Sql = Sql & " From"
        Sql = Sql & " SGIPA..TB_CARGA_SOLTA C LEFT JOIN"
        Sql = Sql & " SGIPA..TB_MARCANTES M ON C.AUTONUM=M.AUTONUM_CARGA"
        Sql = Sql & " Where c.CNTR = " & Me.MskCntr.Tag
        Sql = Sql & " AND M.AUTONUM IS NULL"
        Set TbM = New ADODB.Recordset
        TbM.Open Sql, db, 1, 1
        If Not TbM.EOF Then
            Col_Mensagem = "Existe item desovado sem Marcante Associado, Lote " & TbM!BL
            Col_Imagem = 1
            frmMensagem2.Show 1
        End If
        TbM.Close
    End If
End If
tbp1.Close


Sql = "SELECT LOTE FROM SGIPA..VW_DESOVA_COL_LOTES3 WHERE "
Sql = Sql & " AUTONUM_CNTR=" & Me.MskCntr.Tag & " AND"
Sql = Sql & " RTRIM(FLAG_DESOVADO) NOT IN ('D','X')"
Set tbTodos = New ADODB.Recordset
tbTodos.Open Sql, db, 1, 1
If Not tbTodos.EOF Then
    Col_Mensagem = "Existe Lote não desovado " & tbTodos!Lote
    Col_Imagem = 1
    frmMensagem2.Show 1
    Exit Sub
End If
tbTodos.Close


If Me.txtTermino = "" Then

    Agora = db.Execute("SELECT GETDATE()")(0)
    Me.txtTermino = Format(Agora, "dd/mm/yyyy hh:mm")
    Sql = "UPDATE SGIPA..TB_CNTR_BL SET flag_desovado=1, DT_FIM_DESOVA=GETDATE(), FLAG_DESOVA_MANUAL=" & IIf(Me.Option3(0).Value = True, 1, 0) & " WHERE autonum=" & Me.MskCntr.Tag
    db.Execute Sql
    
Else
    Col_Mensagem = "Término de operação de desova já informado"
    frmMensagem2.Show vbModal
    Exit Sub
End If
    
End Sub

Private Sub Command1_Click()
Carrega_Sol
End Sub

Private Sub dcEmbalagem_GotFocus()
dcEmbalagem.BackColor = &HC0FFFF
End Sub

Private Sub dcEmbalagem_LostFocus()
dcEmbalagem.BackColor = &HFFFFFF
End Sub

Private Sub dcGenero_GotFocus()
'dcGenero.BackColor = &HC0FFFF
End Sub

Private Sub dcGenero_LostFocus()
'dcGenero.BackColor = &HFFFFFF
End Sub

Private Sub dcLote_Change()
Carrega_ItensDesovados (Val(Nnull(Me.dcLote.BoundText, 0)))
Me.TxtObs.text = "Lote : " & Val(Nnull(Me.dcLote.BoundText, 0)) & " " & Nnull(db.Execute("SELECT OBS_DESOVA FROM SGIPA..TB_BL WHERE AUTONUM=" & Val(Nnull(Me.dcLote.BoundText, 0)))(0), 1)

End Sub

Private Sub dcLote_Click(Area As Integer)
'Carrega_ItensDesovados (Val(Nnull(Me.dcLote.BoundText, 0)))
End Sub
Private Sub Carrega_ItensDesovados(AutonumBL As Long)
Me.GridDesovados.Clear
'MsgBox AutonumBl
Sql = "SELECT ITEM,QUANTIDADE,DESCR_EMBALAGEM,MARCA,MERCADORIA,AUTONUM_CS "
Sql = Sql & " FROM "
Sql = Sql & " SGIPA..VW_DESOVA_COL_ITENS_DESOVADOS "
Sql = Sql & " WHERE LOTE=" & AutonumBL
Sql = Sql & " ORDER BY ITEM, QUANTIDADE "

Me.AdodcDesovados.ConnectionString = ConnectionString
Me.AdodcDesovados.RecordSource = Sql
Me.AdodcDesovados.Refresh
Set Me.AdodcDesovados.Recordset.ActiveConnection = Nothing

Me.GridDesovados.FormatString = "   |<Item  |^Qtde|<Embalagem   |<Marca               |<Mercadoria                          |"

Me.TxtObs.text = "Lote : " & AutonumBL & " " & Nnull(db.Execute("SELECT OBS_DESOVA FROM SGIPA..TB_BL WHERE AUTONUM=" & Val(Nnull(AutonumBL, 0)))(0), 1)

End Sub

Private Sub dcLote_GotFocus()
dcLote.BackColor = &HC0FFFF

End Sub

Private Sub dcLote_LostFocus()
dcLote.BackColor = &HFFFFFF
Me.Label2(6).Caption = ""
Me.txtBrutoAP.Enabled = False
Me.txtBrutoAP.text = ""

If Val(Me.dcLote.BoundText) > 0 Then
    Dim Sql As String
    'Sql = "SELECT autonum from sgipa..tb_bl where autonum=" & Val(Me.dcLote.BoundText) & " and flag_hubPort=1 "
    Sql = "SELECT autonum from sgipa..tb_bl where autonum=" & Val(Me.dcLote.BoundText)
    Set tb1 = New ADODB.Recordset
    tb1.Open Sql, db, 1, 1
    If Not tb1.EOF Then
        Me.txtBrutoAP.Enabled = True
       ' Me.Label2(6).Caption = "BL-DTA"
    End If
    tb1.Close
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Sendkeys "{TAB}"
End If

If KeyCode = vbKeyF4 And Me.FrameDesova.Visible = True And Me.frmMarcantes.Visible = False And Me.frmResumo.Visible = False Then
   Call cmdResumo_Click
   Exit Sub
End If


If KeyCode = vbKeyF4 And Me.frmResumo.Visible = True Then
    Call cmdCarregaItem_Click
    Exit Sub
End If



If KeyCode = vbKeyEscape And Me.frmResumo.Visible = True Then
    Me.frmResumo.Visible = False
    Me.FrameDesova.Visible = True
    Exit Sub
End If


If KeyCode = vbKeyEscape And Me.frmMarcantes.Visible = False Then
    If IsLoaded("frmMensagem2") = True Then
        Unload frmMensagem2
    Else
        Unload Me
    End If
    Exit Sub
End If


If KeyCode = vbKeyEscape And Me.frmMarcantes.Visible = True Then
    Me.frmMarcantes.Visible = False
    Exit Sub
End If


If KeyCode = vbKeyF3 Then
    If Me.FrameDesova.Visible = True Then
        Call cmdMarcantes_Click
        
        Exit Sub
    End If
End If

If KeyCode = vbKeyEscape Then
    If Me.FrameDesova.Visible = True Then
        If Me.frmMarcantes.Visible = True Then
            Me.frmMarcantes.Visible = False
            Exit Sub
        End If
    End If
End If


If KeyCode = vbKeyF1 And Me.FrameCntr.Visible = True Then

    Call cmdInicio_Click
    
End If


'DESOVA DE ITEM
If KeyCode = vbKeyF1 And Me.FrameDesova.Visible = True And Me.frmMarcantes.Visible = False Then

    Call cmdDesovaItem_Click
    
End If

'REMOVE ITEM DESOVADO
If KeyCode = vbKeyF2 And Me.FrameDesova.Visible = True Then

    Call cmdRemoveDesova_Click

    
End If

If KeyCode = vbKeyF2 And Me.FrameCntr.Visible = True Then

    Call cmdTermino_Click
    
End If


'Marcantes ADICIONANDO
If KeyCode = vbKeyF4 Then
    If Me.FrameDesova.Visible = True Then
        If Me.frmMarcantes.Visible = True Then
            Call cmdAdicionaMc_Click
        End If
    End If
End If

'MARCANTE - REMOVENDO
If KeyCode = vbKeyF5 And Me.FrameDesova.Visible = True And Me.frmMarcantes.Visible = True Then
    Call cmdRemoveMc_Click
End If



End Sub


Private Sub Form_Load()
Me.Top = 0
Me.Left = 0

Sql = " select autonum, descr AS DISPLAY "
Sql = Sql & " FROM SGIPA..TB_CAD_GRUPO_PRODUTOS "
Sql = Sql & " ORDER BY DESCR "


Sql = " select code as autonum, descr AS DISPLAY "
Sql = Sql & " FROM SGIPA..dte_tb_embalagens "
Sql = Sql & " ORDER BY DESCR "

Me.AdodcEmbalagem.ConnectionString = ConnectionString
Me.AdodcEmbalagem.RecordSource = Sql
Me.AdodcEmbalagem.Refresh
Set Me.AdodcEmbalagem.Recordset.ActiveConnection = Nothing


Sql = " Select matricula as autonum, nome as dISPLAY from sgipa..DTE_TB_FISCAIS order by nome"
Me.AdodcFiscal.ConnectionString = ConnectionString
Me.AdodcFiscal.RecordSource = Sql
Me.AdodcFiscal.Refresh
Set Me.AdodcFiscal.Recordset.ActiveConnection = Nothing

Carrega_Avarias_Col


'wsIntegraChronos_url$ = db.Execute("SELECT URL_WSINTEGRACHRONOS FROM OPERADOR.DTE_TB_PARAMETROS_SISTEMA")(0)

'If wsIntegraChronos_url$ = "" Then
'    MsgBox "Url de integração WebService não cadastrado", vbExclamation
'    Exit Sub
'End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

'Verifica se SortKey é a mesma que a atual
If ListView1.SortKey <> ColumnHeader.Index - 1 Then
'Quando clicar em uma coluna define sortkey para indice -1
   ListView1.SortKey = ColumnHeader.Index - 1
   ListView1.SortOrder = lvwAscending
Else
  'Se a coluna ja esta selecionada entao altera a
   'propr. SetOrder para ser o oposto da coluna em uso
  ListView1.SortOrder = IIf(ListView1.SortOrder = lvwAscending, lvwDescending, lvwAscending)
End If

'Define a propriedade Sorted para utilizar a ordem atual
ListView1.Sorted = True
End Sub

Private Sub MskCntr_Change()
If InStr(1, MskCntr, "_") = 0 Then
    ValidaCntr (Me.MskCntr)
Else
    LimpaCntr
End If
End Sub

Private Sub LimpaCntr()

Me.txtConferente = ""
Me.txtInicio = ""
Me.txtTermino = ""
Me.txtLacreIPA = ""
Me.txtLacreOrigem = ""
Me.txtTamTipo = ""
Me.txtYard = ""
Me.Option3(0).Value = False
Me.Option3(1).Value = False
Me.txtViagem = ""
Me.txtdt = ""
Me.MskCntr.Tag = 0

Me.Option4(1).Value = False
Me.Option4(2).Value = False
Me.Option4(3).Value = False
Me.Option4(4).Value = False
Me.Option4(5).Value = False
Me.DBCombo10.text = ""

End Sub

Private Sub ValidaCntr(IDConteiner$)
'CONTEINER EM HISTORICO
'CONTEINER NO PATIO DO CONFERENTE LOGADO
'CONTEINER EM POSICAO DE DESOVA NO PATIO
'DESOVA AINDA NAO ENVIADA PARA ALFANDEGA
'REGIME PIER
'DATA DE TERMINO DA DESOVA AINDA NAO INFORMADA OU COM ALGUM BL EM RECONHECIMENTO DE CARGA
'CONTEINER COM AGENDAMENTO DE ABERTURA

Sql = "SELECT 0 as reconhecimento,C.AUTONUM,C.YARD,C.PATIO,C.DT_INICIO_DESOVA,C.DT_FIM_DESOVA,1 AS FLAG_LOCAL_ABERTURA,C.REGIME, "
Sql = Sql & " C.ID_AGENDAMENTO,C.FLAG_DESOVA_MANUAL, "
Sql = Sql & " C.TAMANHO, TC.CODIGO AS TIPOBASICO,C.LACRE_ORIGEM,C.LACRE_IPA,'-' AS CAMERA,0 AS ID_CAMERA, "
Sql = Sql & " C.DT,C.VIAGEM,C.FLAG_DESOVA_IMPORTADOR,C.AUTONUM_SOLICITOU_DESOVA  "
Sql = Sql & " FROM "
Sql = Sql & " SGIPA..TB_CNTR_BL C "
Sql = Sql & " LEFT Join "
Sql = Sql & " SGIPA..dte_tb_tipos_conteiner TC ON C.TIPO=TC.CODE"
Sql = Sql & " WHERE "
Sql = Sql & " C.FLAG_TERMINAL=1 "
Sql = Sql & " AND C.FLAG_HISTORICO=0 "
Sql = Sql & " AND C.ID_CONTEINER='" & IDConteiner & "'"
Sql = Sql & " UNION "
Sql = Sql & " SELECT DISTINCT 1 as reconhecimento,C.AUTONUM,C.YARD,C.PATIO,C.DT_INICIO_DESOVA,C.DT_FIM_DESOVA,1 AS FLAG_LOCAL_ABERTURA,C.REGIME, "
Sql = Sql & " C.ID_AGENDAMENTO,C.FLAG_DESOVA_MANUAL, "
Sql = Sql & " C.TAMANHO, TC.CODIGO AS TIPOBASICO,C.LACRE_ORIGEM,C.LACRE_IPA,'-' AS CAMERA,0 AS ID_CAMERA, "
Sql = Sql & " C.DT,C.VIAGEM,C.FLAG_DESOVA_IMPORTADOR,C.AUTONUM_SOLICITOU_DESOVA  "
Sql = Sql & " FROM "
Sql = Sql & " SGIPA..TB_CNTR_BL C INNER JOIN SGIPA..TB_AMR_CNTR_BL A ON C.AUTONUM=A.CNTR "
Sql = Sql & " LEFT Join "
Sql = Sql & " SGIPA..dte_tb_tipos_conteiner TC ON C.TIPO=TC.CODE"
Sql = Sql & " INNER JOIN SGIPA..TB_CARGA_CNTR CC ON C.ID_CONTEINER=CC.ID_CONTEINER AND A.BL=CC.BL"
Sql = Sql & " WHERE "
Sql = Sql & " C.FLAG_TERMINAL=1 "
Sql = Sql & " AND C.FLAG_HISTORICO=1 "
Sql = Sql & " AND C.ID_CONTEINER='" & IDConteiner & "' AND CC.FLAG_RECONHECIMENTO_CARGA=1 "



Set tb1 = New ADODB.Recordset
tb1.Open Sql, db, 1, 1

If tb1.EOF Then
    Col_Mensagem = "Contêiner não encontrado em estoque"
    Col_Imagem = 1
    frmMensagem2.Show 1
    Exit Sub
End If


If Nnull(tb1!FLAG_DESOVA_IMPORTADOR, 0) > 0 Then
    Me.Option4(Nnull(tb1!FLAG_DESOVA_IMPORTADOR, 0)).Value = True
    For i = 1 To 5
        Me.Option4(i).Enabled = False
    Next i
Else
    For i = 1 To 5
        Me.Option4(i).Enabled = True
        Me.Option4(i).Value = False
    Next i

End If


If Nnull(tb1!AUTONUM_SOLICITOU_DESOVA, 0) > 0 Then
    Me.DBCombo10.Enabled = False
    Me.Command1.Enabled = False
    Call Carrega_Sol(Nnull(tb1!AUTONUM_SOLICITOU_DESOVA, 0))
    Me.DBCombo10.BoundText = Nnull(tb1!AUTONUM_SOLICITOU_DESOVA, 0)
Else
    Me.DBCombo10.Enabled = True
    Me.Command1.Enabled = True
End If
'19/06 - incluida chamada para carregar lista de parceiros
'        Sem carregar a lista, sistema apresenta msg de erro ao preencher o campo.



If Col_Patio <> 1 And Col_Patio <> 7 Then
    If Val(tb1!Patio) <> Col_Patio Then
        Col_Mensagem = "Contêiner se encontra em pátio divergente ao que o conferente está logado"
        Col_Imagem = 1
        frmMensagem2.Show 1
        Exit Sub
    End If
Else
    If Val(tb1!Patio) <> 7 And Val(tb1!Patio) <> 1 Then
        Col_Mensagem = "Contêiner se encontra em pátio divergente ao que o conferente está logado"
        Col_Imagem = 1
        frmMensagem2.Show 1
        Exit Sub
    End If
End If

'If Nnull(tb1!FLAG_LOCAL_ABERTURA, 0) = 0 Then
'    Col_Mensagem = "Local onde o Contêiner se encontra no pátio (" & Nnull(tb1!yard, 1) & ") não configurado como área de abertura"
'    Col_Imagem = 1
'    frmMensagem2.Show 1
'    Exit Sub
'End If


If tb1!Reconhecimento = 0 Then

    Col_Reconhecimento = 0

    Sql = "SELECT BL.AUTONUM FROM "
    Sql = Sql & " SGIPA..TB_BL BL "
    Sql = Sql & " INNER JOIN SGIPA..TB_AMR_CNTR_BL A ON "
    Sql = Sql & " BL.AUTONUM=A.BL"
    Sql = Sql & " AND A.CNTR= " & tb1!AUTONUM
    Sql = Sql & " AND ISNULL(BL.AUDIT_DESOVA,0)>0 "
    Set Tbh = New ADODB.Recordset
    Tbh.Open Sql, db, 1, 1
    If Not Tbh.EOF Then
        Col_Mensagem = "Existe lote atrelado ao conteiner (" & Nnull(Tbh!AUTONUM, 0) & ") com desova informada para a alfândega"
        Col_Imagem = 1
        frmMensagem2.Show 1
        Exit Sub
    End If
    Tbh.Close
Else
    Col_Reconhecimento = 1
End If

If Right(Nnull(tb1!Regime, 1), 1) <> "P" Then
    Col_Mensagem = "Regime do conteiner (" & Nnull(tb1!Regime, 1) & ") não é Pier"
    Col_Imagem = 1
    frmMensagem2.Show 1
    Exit Sub
End If


If tb1!Reconhecimento = 0 Then
    If Nnull(tb1!DT_FIM_DESOVA, 1) <> "" Then
        'Verifica se existe algum bl em reconhecimento de carga
        Sql = "SELECT AUTONUM FROM SGIPA..TB_CARGA_CNTR WHERE ID_CONTEINER='" & IDConteiner & "' AND FLAG_RECONHECIMENTO_CARGA=1 "
        Sql = Sql & " AND BL IN (SELECT BL FROM SGIPA..TB_AMR_CNTR_BL WHERE CNTR=" & tb1!AUTONUM & ")"
        Col_Mensagem = "Data de término da desova já informada (" & Format(tb1!DT_FIM_DESOVA, "dd/mm/yy hh:mm") & ")"
        Col_Imagem = 1
        frmMensagem2.Show 1
        Exit Sub
    End If
End If

If Nnull(tb1!ID_AGENDAMENTO, 0) = 0 Then
    Col_Mensagem = "Contêiner não agendado para abertura "
    Col_Imagem = 1
    frmMensagem2.Show 1
    Exit Sub
End If

'CARREGANDO OS DADOS
Me.MskCntr.Tag = tb1!AUTONUM
Me.txtInicio = Format(tb1!dt_inicio_desova, "DD/MM/YYYY HH:MM")
Me.txtTermino = Format(tb1!DT_FIM_DESOVA, "DD/MM/YYYY HH:MM")

Me.txtTamTipo = Nnull(tb1!Tamanho, 0) & Nnull(tb1!TipoBasico, 1)
Me.txtYard = Nnull(tb1!Yard, 1)
Me.txtLacreIPA = Nnull(tb1!Lacre_ipa, 1)
Me.txtLacreOrigem = Nnull(tb1!lacre_origem, 1)



Me.txtConferente = Usuario_Sistema

Me.txtViagem = Nnull(tb1!Viagem, 1)
Me.txtdt = Nnull(tb1!DT, 1)


If Nnull(tb1!FLAG_DESOVA_MANUAL, 0) = 1 Then
    Me.Option3(0).Value = True
Else
    Me.Option3(1).Value = True
End If

End Sub

Private Sub Option2_Click(Index As Integer)
If Index = 1 Then

    If Nnull(Me.MskCntr.Tag, 0) = 0 Then
        Col_Mensagem = "Contêiner inválido"
        Col_Imagem = 1
        frmMensagem2.Show 1
        Option2(0).Value = True
        Exit Sub
    End If


    If Me.txtTermino <> "" And Col_Reconhecimento = 0 Then
        Col_Mensagem = "Término da operação de desova já informado"
        Col_Imagem = 1
        frmMensagem2.Show 1
        Option2(0).Value = True
        Exit Sub
    End If

    If Me.txtInicio = "" Then
        Col_Mensagem = "Não foi informado o início da operação de desova"
        Col_Imagem = 1
        frmMensagem2.Show 1
        Option2(0).Value = True
        Exit Sub
    Else
        Me.FrameCntr.Visible = False
        Me.FrameDesova.Visible = True
        Carrega_Lotes
        Me.dcLote.SetFocus
        
        Dim colx As ColumnHeader

        'Inclui as colunas
        
        Set colx = ListView1.ColumnHeaders.Add(, , "i")
        colx.Width = 0
        colx.Alignment = lvwColumnLeft
        
        Set colx = ListView1.ColumnHeaders.Add(, , "  ")
        colx.Width = 280
        colx.Alignment = lvwColumnCenter
        
        Set colx = ListView1.ColumnHeaders.Add(, , "Lote")
        colx.Width = 650
        colx.Alignment = lvwColumnLeft
        
        Set colx = ListView1.ColumnHeaders.Add(, , "BL")
        colx.Width = 1000
        colx.Alignment = lvwColumnLeft
        
        Set colx = ListView1.ColumnHeaders.Add(, , "ITM")
        colx.Width = 250
        colx.Alignment = lvwColumnLeft
        
        Set colx = ListView1.ColumnHeaders.Add(, , "Qte")
        colx.Width = 600
        colx.Alignment = lvwColumnCenter
        
    
        Set colx = ListView1.ColumnHeaders.Add(, , "EMBALAGEM")
        colx.Width = 1000
        colx.Alignment = lvwColumnLeft
        
        
        Set colx = ListView1.ColumnHeaders.Add(, , "MERCADORIA")
        colx.Width = 1000
        colx.Alignment = lvwColumnLeft
        
        
        Set colx = ListView1.ColumnHeaders.Add(, , "MARCA")
        colx.Width = 1000
        colx.Alignment = lvwColumnLeft
        
    End If
    
Else

    Me.FrameCntr.Visible = True
    Me.FrameDesova.Visible = False
    
    
    

    
    
End If
End Sub

Private Sub Carrega_Lotes()

If Col_Reconhecimento = 0 Then
    'Sql = " select LOTE as autonum, '(' + RTRIM(CONVERT(CHAR,LOTE)) +  ')' + ' ' + NUMERO_BL AS DISPLAY "
    Sql = " Select LOTE As autonum, '[ ' + FLAG_DESOVADO + '] (' + RTRIM(CONVERT(CHAR,LOTE)) + ')' + ' ' + NUMERO_BL AS DISPLAY"
    Sql = Sql & " FROM SGIPA..VW_DESOVA_COL_LOTES3 "
    Sql = Sql & " WHERE "
    Sql = Sql & " AUTONUM_CNTR=" & Me.MskCntr.Tag
    Sql = Sql & " ORDER BY LOTE "
Else
    'Sql = " select LOTE as autonum, '(' + RTRIM(CONVERT(CHAR,LOTE)) +  ')' + ' ' + NUMERO_BL AS DISPLAY "
    Sql = " Select LOTE As autonum, '[ ' + FLAG_DESOVADO + '] (' + RTRIM(CONVERT(CHAR,LOTE)) + ')' + ' ' + NUMERO_BL AS DISPLAY"
    Sql = Sql & " FROM SGIPA..VW_DESOVA_COL_LOTES_REC3 "
    Sql = Sql & " WHERE "
    Sql = Sql & " AUTONUM_CNTR=" & Me.MskCntr.Tag
    Sql = Sql & " ORDER BY LOTE "

End If

Me.AdodcLote.ConnectionString = ConnectionString
Me.AdodcLote.RecordSource = Sql
Me.AdodcLote.Refresh
Set Me.AdodcLote.Recordset.ActiveConnection = Nothing


End Sub

Private Sub Option4_Click(Index As Integer)
SolDesova = Index
DBCombo10.text = ""

End Sub

Private Sub Carrega_Sol(Optional ID As Long = 0)


 'If Me.DBCombo10.Text = "" Then Exit Sub

 Dim Rst As New ADODB.Recordset
 
 If ID = 0 Then
 
     If Option4(1) = True Then
    
        PRSet Rst, "Select Autonum ,Razao ||' ' ||Cgc Display from " & Banco_Sgipa & "TB_CAD_PARCEIROS where flag_Importador=1 and razao like '%" & UCase(DBCombo10.text) & "%' and flag_ativo=1 order by Display"
     End If
     If Option4(2) = True Then
        
        PRSet Rst, "Select Autonum ,Razao ||' ' ||Cgc Display from " & Banco_Sgipa & "TB_CAD_PARCEIROS where flag_armador=1 and razao like '%" & UCase(DBCombo10.text) & "%' and flag_ativo=1 order by Display"
     End If
     If Option4(3) = True Then
        
        PRSet Rst, "Select Autonum ,Razao ||' ' ||Cgc Display from " & Banco_Sgipa & "TB_CAD_PARCEIROS where flag_Nvocc=1 and razao like '%" & UCase(DBCombo10.text) & "%' and flag_ativo=1 order by Display"
     End If
     If Option4(4) = True Then
        
        PRSet Rst, "Select Autonum ,Razao ||' ' ||Cgc Display from " & Banco_Sgipa & "TB_CAD_PARCEIROS where flag_despachante=1 and razao like '%" & UCase(DBCombo10.text) & "%' and flag_ativo=1 order by Display"
     End If
     If Option4(5) = True Then
        PRSet Rst, "Select 0 autonum ,'' display from dual"
    End If

Else
        PRSet Rst, "Select Autonum ,Razao ||' ' ||Cgc Display from " & Banco_Sgipa & "TB_CAD_PARCEIROS where autonum=" & ID

End If
 
 MontaCombo Me.DBCombo10, Rst, "display", "autonum"
End Sub


Private Sub txtBruto_GotFocus()
txtBruto.BackColor = &HC0FFFF
a = Marca(Me)

End Sub

Private Sub txtBruto_KeyPress(KeyAscii As Integer)
KeyAscii = SoNumeros(KeyAscii)
End Sub

Private Sub txtBruto_LostFocus()
txtBruto.BackColor = &HFFFFFF
End Sub

Private Sub txtBrutoAP_GotFocus()
txtBrutoAP.BackColor = &HC0FFFF
a = Marca(Me)
End Sub

Private Sub txtBrutoAP_KeyPress(KeyAscii As Integer)
KeyAscii = SoNumeros(KeyAscii)
End Sub

Private Sub txtBrutoAP_LostFocus()
txtBrutoAP.BackColor = &HFFFFFF
End Sub

Private Sub txtConsertado_Click(Index As Integer)
If txtConsertado(Index).BackColor = vbRed Then
    txtConsertado(Index).BackColor = vbMagenta
    Me.ChkAv(Index).Value = 1
End If
If txtConsertado(Index).BackColor = vbMagenta Then
    txtConsertado(Index).BackColor = vbGreen
    Me.ChkAv(Index).Value = 1
End If
If txtConsertado(Index).BackColor = vbGreen Then
    txtConsertado(Index).BackColor = vbRed
End If


End Sub

Private Sub txtItem_GotFocus()
txtItem.BackColor = &HC0FFFF
End Sub

Private Sub txtItem_LostFocus()
If Me.txtItem <> "" Then
    If Nnull(Me.dcLote.BoundText, 0) = 0 Then
        Col_Mensagem = "Primeiro escolha o lote"
        Col_Imagem = 1
        frmMensagem2.Show 1
        Limpa_Item
    Else
        
        Call Valida_Item(Nnull(Me.dcLote.BoundText, 0), Nnull(Me.txtItem, 0))
    End If
End If

txtItem.BackColor = &HFFFFFF

End Sub

Private Sub Limpa_Item()

Me.txtItem = ""
Me.txtBruto = ""
Me.txtMarca = ""
Me.TXTMARCANTE = ""
Me.txtPrateleira = ""
Me.txtMercadoria = ""
Me.txtQtde = ""
Me.txtQtde.Tag = "0"

Me.txtQtdeAv = "0"
Me.txtPesoAvariado = "0"

Me.txtVolume = ""
'Me.dcGenero = ""
Me.dcEmbalagem = ""
Me.ckAvaria.Value = 0
Me.ckAcrescimo.Value = 0
Me.ckIDFA.Value = 0
Me.ckMadeira.Value = 0
Me.ckReconhecimento.Value = 0

Carrega_Avarias_Col

End Sub

Private Sub Valida_Item(AutonumBL As Long, Item As Integer)
'Procura no desovado, se nao achar no manifestado

Limpa_Item

Dim NovoItem As Boolean

Sql = "SELECT "
Sql = Sql & " ID_CONTEINER,"
Sql = Sql & " AUTONUM_CNTR,"
Sql = Sql & " LOTE,"
Sql = Sql & " ITEM,"
Sql = Sql & " QUANTIDADE,"
Sql = Sql & " COD_EMBALAGEM,"
Sql = Sql & " DESCR_EMBALAGEM,"
Sql = Sql & " MERCADORIA,"
Sql = Sql & " MARCA,"
Sql = Sql & " PESO_BRUTO,"
Sql = Sql & " FLAG_AVARIA,"
Sql = Sql & " FLAG_MADEIRA,"
Sql = Sql & " VIAGEM,"
Sql = Sql & " DT,"
Sql = Sql & " Patio "
Sql = Sql & " FROM SGIPA..VW_DESOVA_COL_ITENS"
Sql = Sql & " WHERE LOTE=" & AutonumBL
Sql = Sql & " AND ITEM=" & Item
Set tb1 = New ADODB.Recordset
tb1.Open Sql, db, 1, 1
If tb1.EOF Then

    Col_Mensagem = "Item não encontrado no lote selecionado. Deseja criar um novo item "
    Col_Imagem = 3
    frmMensagem2.Show 1
    If Col_Cancelado = True Then
        Limpa_Item
        Exit Sub
    Else
        NovoItem = True
        
        tb1.Close
        Sql = "SELECT "
        Sql = Sql & " ID_CONTEINER,"
        Sql = Sql & " AUTONUM_CNTR,"
        Sql = Sql & " LOTE,"
        Sql = Sql & " ITEM,"
        Sql = Sql & " QUANTIDADE,"
        Sql = Sql & " COD_EMBALAGEM,"
        Sql = Sql & " DESCR_EMBALAGEM,"
        Sql = Sql & " MERCADORIA,"
        Sql = Sql & " MARCA,"
        Sql = Sql & " PESO_BRUTO,"
        Sql = Sql & " FLAG_AVARIA,"
        Sql = Sql & " FLAG_MADEIRA,"
        Sql = Sql & " VIAGEM,"
        Sql = Sql & " DT,"
        Sql = Sql & " Patio "
        Sql = Sql & " FROM SGIPA..VW_DESOVA_COL_ITENS"
        Sql = Sql & " WHERE LOTE=" & AutonumBL
        Set tb1 = New ADODB.Recordset
        tb1.Open Sql, db, 1, 1
        If Not tb1.EOF Then
            GoSub Preenche_Item
        End If
        tb1.Close
        
    End If

Else

    NovoItem = False
    GoSub Preenche_Item


End If


Exit Sub

Preenche_Item:
Me.txtItem = Item
If Not NovoItem Then
    Me.dcEmbalagem.BoundText = tb1!COD_EMBALAGEM
    Me.dcEmbalagem.Tag = tb1!COD_EMBALAGEM
Else
    Me.dcEmbalagem = ""
    Me.dcEmbalagem.Tag = 0
End If
Me.txtBruto = Nnull(tb1!PESO_BRUTO, 0)
Me.txtMarca = Nnull(tb1!Marca, 1)
Me.txtMercadoria = Nnull(tb1!MERCADORIA, 1)
Me.txtQtde = Nnull(tb1!QUANTIDADE, 0)
Me.txtQtde.Tag = Nnull(tb1!QUANTIDADE, 0)
Me.txtVolume = ""
'Me.dcGenero.BoundText = 0
Me.ckAvaria.Value = Nnull(tb1!FLAG_AVARIA, 0)
Me.ckAcrescimo.Value = 0
Me.ckIDFA.Value = 0
Me.ckMadeira.Value = Nnull(tb1!FLAG_MADEIRA, 0)
Return

End Sub

Private Sub txtMarca_GotFocus()
txtMarca.BackColor = &HC0FFFF
End Sub

Private Sub txtMarca_LostFocus()
txtMarca.BackColor = &HFFFFFF
End Sub

Private Sub txtMarcante_GotFocus()
TXTMARCANTE = ""
TXTMARCANTE.BackColor = &HC0FFFF
End Sub

Private Sub txtMarcante_LostFocus()
TXTMARCANTE.BackColor = &HFFFFFF
End Sub

Private Sub txtMercadoria_GotFocus()
txtMercadoria.BackColor = &HC0FFFF
End Sub

Private Sub txtMercadoria_LostFocus()
txtMercadoria.BackColor = &HFFFFFF
End Sub

Private Sub txtPesoAvariado_GotFocus()
a = Marca(Me)
End Sub

Private Sub txtPesoAvariado_KeyPress(KeyAscii As Integer)
KeyAscii = SoNumeros09(KeyAscii)
End Sub

Private Sub txtPrateleira_GotFocus()
txtPrateleira = ""
txtPrateleira.BackColor = &HC0FFFF
a = Marca(Me)
End Sub

Private Sub txtPrateleira_LostFocus()
txtPrateleira.BackColor = vbWhite
txtPrateleira = UCase(txtPrateleira)
End Sub

Private Sub txtQtde_GotFocus()
txtQtde.BackColor = &HC0FFFF
a = Marca(Me)
End Sub

Private Sub txtQtde_KeyPress(KeyAscii As Integer)
KeyAscii = SoNumeros09(KeyAscii)
End Sub

Private Sub txtQtde_LostFocus()
txtQtde.BackColor = &HFFFFFF
End Sub

Private Sub txtQtdeAv_GotFocus()
a = Marca(Me)
End Sub

Private Sub txtQtdeMarcante_GotFocus()
txtQtdeMarcante.BackColor = &HC0FFFF
End Sub

Private Sub txtQtdeMarcante_KeyPress(KeyAscii As Integer)
KeyAscii = SoNumeros09(KeyAscii)
End Sub

Private Sub txtQtdeMarcante_LostFocus()
txtQtdeMarcante.BackColor = &HFFFFFF
End Sub

Private Sub txtQtdeAv_KeyPress(KeyAscii As Integer)
KeyAscii = SoNumeros09(KeyAscii)
End Sub

Private Sub txtVolume_GotFocus()
txtVolume.BackColor = &HC0FFFF
a = Marca(Me)
End Sub

Private Sub txtVolume_KeyPress(KeyAscii As Integer)
KeyAscii = SoNumeros(KeyAscii)
End Sub

Private Sub txtVolume_LostFocus()
txtVolume.BackColor = &HFFFFFF
End Sub


Private Sub Carrega_Resumo()
'Lote,BL,ITM,Q.M,Q.D,EMBALAGEM
Dim Sql As String
Sql = " SELECT A.LOTE,BL.NUMERO,A.ITEM,A.QUANTIDADE,ISNULL(B.QUANTIDADE,0) as qtdeD,A.DESCR_EMBALAGEM,A.MARCA,A.MERCADORIA "
Sql = Sql & " From"
Sql = Sql & " SGIPA..VW_DESOVA_COL_ITENS A INNER JOIN"
Sql = Sql & " SGIPA..TB_BL BL ON A.LOTE=BL.AUTONUM"
Sql = Sql & " LEFT JOIN SGIPA..VW_DESOVA_COL_ITENS_DESOVADOS B"
Sql = Sql & " ON A.AUTONUM_CNTR = B.AUTONUM_CNTR AND"
Sql = Sql & " A.LOTE=B.LOTE AND "
Sql = Sql & " a.Item = b.Item"
Sql = Sql & " WHERE A.AUTONUM_CNTR=" & Me.MskCntr.Tag
Sql = Sql & " ORDER BY A.LOTE,A.ITEM"

Me.ListView1.ListItems.Clear
Dim itmx As ListItem

Dim tb1 As New ADODB.Recordset
tb1.Open Sql, db, 1, 1
i = 1

    
While Not tb1.EOF
    
    Set itmx = ListView1.ListItems.Add(, , i)
    If Nnull(tb1!qtdeD, 0) = Nnull(tb1!QUANTIDADE, 0) Then
        itmx.SubItems(1) = "X"
    Else
        itmx.SubItems(1) = " "
    End If
    itmx.SubItems(2) = Nnull(tb1!Lote, 1)
    itmx.SubItems(3) = Nnull(tb1!Numero, 1)
    itmx.SubItems(4) = Nnull(tb1!Item, 1)
    itmx.SubItems(5) = Nnull(tb1!qtdeD, 0) & "/" & Nnull(tb1!QUANTIDADE, 0)
    itmx.SubItems(6) = Nnull(tb1!DESCR_EMBALAGEM, 1)
    itmx.SubItems(7) = Nnull(tb1!MERCADORIA, 1)
    itmx.SubItems(8) = Nnull(tb1!Marca, 1)
    
    i = i + 1
    tb1.MoveNext
Wend

tb1.Close

End Sub
