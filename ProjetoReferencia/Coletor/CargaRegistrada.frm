VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form CargaRegistrada 
   Caption         =   "Carga Registrada"
   ClientHeight    =   6330
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtRegistro 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1035
      TabIndex        =   10
      Top             =   90
      Width           =   1620
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   270
      TabIndex        =   0
      Top             =   1635
      Width           =   8310
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   3255
         TabIndex        =   6
         Top             =   315
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Busca"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7230
         TabIndex        =   4
         Top             =   300
         Width           =   930
      End
      Begin MSMask.MaskEdBox MskPlaca 
         Height          =   300
         Left            =   690
         TabIndex        =   3
         Top             =   300
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">???-#A##"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Height          =   315
         Left            =   3255
         TabIndex        =   8
         Top             =   660
         Width           =   3180
         _ExtentX        =   5609
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instrução"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   2310
         TabIndex        =   7
         Top             =   690
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reserva"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2415
         TabIndex        =   2
         Top             =   330
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Placa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   315
         Width           =   540
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   60
      TabIndex        =   9
      Top             =   105
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   165
      TabIndex        =   5
      Top             =   2925
      Width           =   675
   End
End
Attribute VB_Name = "CargaRegistrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label2_Click()

End Sub


