VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTalieColetor 
   BackColor       =   &H00C1EAF9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Talie de descarga / estufagem"
   ClientHeight    =   3735
   ClientLeft      =   285
   ClientTop       =   420
   ClientWidth     =   3240
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   3240
   Begin TabDlg.SSTab SSTab1 
      Height          =   3720
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   6562
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Talie"
      TabPicture(0)   =   "FrmTalieColetor.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameTalie"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmd(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmd(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmd(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmd(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmd(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmd(11)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CmdFinaliza"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Descarga"
      TabPicture(1)   =   "FrmTalieColetor.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frameitem"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Consulta"
      TabPicture(2)   =   "FrmTalieColetor.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "framenota"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "frameitens"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Filtro"
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
         Height          =   3360
         Left            =   45
         TabIndex        =   79
         Top             =   315
         Visible         =   0   'False
         Width           =   3150
         Begin VB.CommandButton cmdFiltro 
            Appearance      =   0  'Flat
            Height          =   450
            Left            =   150
            MaskColor       =   &H00E3F7FF&
            Picture         =   "FrmTalieColetor.frx":0054
            Style           =   1  'Graphical
            TabIndex        =   83
            ToolTipText     =   "LOCALIZA REGISTRO ESPECÍFICO"
            Top             =   2280
            Width           =   2835
         End
         Begin VB.TextBox FiltroInstrucao 
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
            Height          =   315
            Index           =   6
            Left            =   900
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1740
         End
         Begin VB.TextBox FiltroReserva 
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
            Height          =   315
            Index           =   0
            Left            =   885
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   1095
            Width           =   1770
         End
         Begin VB.CommandButton Command1 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2790
            TabIndex        =   80
            Top             =   165
            Width           =   330
         End
         Begin MSMask.MaskEdBox FiltroCntr 
            Height          =   300
            Left            =   900
            TabIndex        =   84
            Top             =   1800
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">????######-#"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskFiltroPlaca 
            Height          =   300
            Left            =   885
            TabIndex        =   85
            Top             =   765
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">???-#A##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Instrução"
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
            Index           =   10
            Left            =   165
            TabIndex        =   89
            Top             =   1500
            Width           =   570
         End
         Begin VB.Label Label1 
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
            Height          =   165
            Index           =   11
            Left            =   165
            TabIndex        =   88
            Top             =   1155
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Conteiner"
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
            Index           =   12
            Left            =   165
            TabIndex        =   87
            Top             =   1830
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Placa"
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
            Index           =   15
            Left            =   180
            TabIndex        =   86
            Top             =   825
            Width           =   330
         End
      End
      Begin VB.CommandButton CmdFinaliza 
         Caption         =   "F"
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
         Height          =   330
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   96
         ToolTipText     =   "Finaliza Talie"
         Top             =   3300
         Width           =   285
      End
      Begin VB.CommandButton cmd 
         Appearance      =   0  'Flat
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   1515
         MaskColor       =   &H00E3F7FF&
         Style           =   1  'Graphical
         TabIndex        =   95
         ToolTipText     =   "DESFAZ ALTERAÇÕES PENDENTES"
         Top             =   3300
         Width           =   285
      End
      Begin VB.CommandButton cmd 
         Appearance      =   0  'Flat
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   60
         MaskColor       =   &H00E3F7FF&
         Style           =   1  'Graphical
         TabIndex        =   94
         Tag             =   "1"
         ToolTipText     =   "INSERE NOVO REGISTRO"
         Top             =   3300
         Width           =   285
      End
      Begin VB.CommandButton cmd 
         Appearance      =   0  'Flat
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   360
         MaskColor       =   &H00E3F7FF&
         Style           =   1  'Graphical
         TabIndex        =   93
         Tag             =   "2"
         ToolTipText     =   "EDITA REGISTRO ATUAL"
         Top             =   3300
         Width           =   285
      End
      Begin VB.CommandButton cmd 
         Appearance      =   0  'Flat
         Caption         =   "G"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   930
         MaskColor       =   &H00E3F7FF&
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "SALVA DADOS NO REGISTRO ATUAL"
         Top             =   3300
         Width           =   285
      End
      Begin VB.CommandButton cmd 
         Appearance      =   0  'Flat
         Caption         =   "C"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   1230
         MaskColor       =   &H00E3F7FF&
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   "DESFAZ ALTERAÇÕES PENDENTES"
         Top             =   3300
         Width           =   285
      End
      Begin VB.CommandButton cmd 
         Appearance      =   0  'Flat
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   645
         MaskColor       =   &H00E3F7FF&
         Style           =   1  'Graphical
         TabIndex        =   90
         Tag             =   "3"
         ToolTipText     =   "ELIMINA REGISTRO ATUAL"
         Top             =   3300
         Width           =   285
      End
      Begin VB.Frame FrameTalie 
         BackColor       =   &H00C1EAF9&
         BorderStyle     =   0  'None
         Caption         =   "Talie"
         Enabled         =   0   'False
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
         Height          =   3315
         Left            =   30
         TabIndex        =   51
         Top             =   315
         Width           =   3150
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   195
            Left            =   2895
            TabIndex        =   100
            Top             =   315
            Width           =   225
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C1EAF9&
            BorderStyle     =   0  'None
            Height          =   525
            Left            =   45
            TabIndex        =   60
            Top             =   600
            Width           =   3165
            Begin MSMask.MaskEdBox MskData 
               Height          =   255
               Index           =   0
               Left            =   795
               TabIndex        =   61
               Top             =   0
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   450
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
            Begin MSMask.MaskEdBox MskCntr 
               Height          =   255
               Left            =   1725
               TabIndex        =   62
               Top             =   270
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   450
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   12
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">????######-#"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskPlaca 
               Height          =   255
               Left            =   795
               TabIndex        =   63
               Top             =   270
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   450
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">???-#A##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskData 
               Height          =   255
               Index           =   1
               Left            =   2280
               TabIndex        =   99
               Top             =   0
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Placa/Cntr"
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
               Left            =   90
               TabIndex        =   65
               Top             =   315
               Width           =   645
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "INICIO"
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
               Index           =   0
               Left            =   105
               TabIndex        =   64
               Top             =   45
               Width           =   510
            End
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00C1EAF9&
            Caption         =   "Estuf. Completa"
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
            Left            =   1920
            TabIndex        =   58
            Top             =   2640
            Width           =   900
         End
         Begin VB.TextBox TxtRegistro 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   840
            TabIndex        =   57
            Top             =   300
            Width           =   705
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C1EAF9&
            Enabled         =   0   'False
            Height          =   345
            Left            =   30
            TabIndex        =   53
            Top             =   2595
            Width           =   1800
            Begin VB.OptionButton Option1 
               BackColor       =   &H00C1EAF9&
               Caption         =   "CD"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   150
               Left            =   45
               TabIndex        =   56
               Top             =   135
               Width           =   495
            End
            Begin VB.OptionButton Option2 
               BackColor       =   &H00C1EAF9&
               Caption         =   "DA"
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
               Left            =   600
               TabIndex        =   55
               Top             =   120
               Value           =   -1  'True
               Width           =   495
            End
            Begin VB.OptionButton Option3 
               BackColor       =   &H00C1EAF9&
               Caption         =   "CCS"
               Enabled         =   0   'False
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
               Left            =   1155
               TabIndex        =   54
               Top             =   120
               Width           =   510
            End
         End
         Begin VB.ComboBox Combo1 
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
            ItemData        =   "FrmTalieColetor.frx":1156
            Left            =   840
            List            =   "FrmTalieColetor.frx":1160
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   2235
            Width           =   2130
         End
         Begin MSDataListLib.DataCombo DcReserva 
            Bindings        =   "FrmTalieColetor.frx":117A
            Height          =   285
            Left            =   840
            TabIndex        =   59
            Top             =   1410
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ListField       =   "reference"
            BoundColumn     =   "autonum_boo"
            Text            =   "DataCombo3"
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
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "FrmTalieColetor.frx":118F
            Height          =   285
            Left            =   840
            TabIndex        =   66
            Top             =   1680
            Width           =   2100
            _ExtentX        =   3704
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
         Begin MSDataListLib.DataCombo DataCombo2 
            Bindings        =   "FrmTalieColetor.frx":11A4
            Height          =   285
            Left            =   840
            TabIndex        =   67
            Top             =   1950
            Width           =   2100
            _ExtentX        =   3704
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
         Begin MSDataListLib.DataCombo DCombo2 
            Bindings        =   "FrmTalieColetor.frx":11B9
            Height          =   285
            Index           =   2
            Left            =   840
            TabIndex        =   68
            Top             =   1140
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   503
            _Version        =   393216
            Locked          =   -1  'True
            MatchEntry      =   -1  'True
            Appearance      =   0
            Style           =   2
            ListField       =   "fantasia"
            BoundColumn     =   "autonum_parceiro"
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
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "SUZANO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Left            =   1650
            TabIndex        =   102
            Top             =   375
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label LbTalie 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   225
            Left            =   840
            TabIndex        =   78
            Top             =   0
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CONF."
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
            Index           =   2
            Left            =   105
            TabIndex        =   77
            Top             =   1710
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EQUIPE"
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
            Index           =   3
            Left            =   105
            TabIndex        =   76
            Top             =   1995
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CLIENTE"
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
            Index           =   4
            Left            =   120
            TabIndex        =   75
            Top             =   1200
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RESERVA"
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
            Left            =   120
            TabIndex        =   74
            Top             =   1440
            Width           =   660
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "CROSSDOCKING"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Left            =   1605
            TabIndex        =   73
            Top             =   180
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TALIE FECHADO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Left            =   1530
            TabIndex        =   72
            Top             =   15
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Registro"
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
            Index           =   17
            Left            =   120
            TabIndex        =   71
            Top             =   315
            Width           =   600
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MODO"
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
            Left            =   105
            TabIndex        =   70
            Top             =   2280
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Talie"
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
            Index           =   5
            Left            =   105
            TabIndex        =   69
            Top             =   0
            Width           =   330
         End
      End
      Begin VB.Frame Frameitem 
         BackColor       =   &H00C1EAF9&
         BorderStyle     =   0  'None
         Caption         =   "Itens Descarregados"
         Enabled         =   0   'False
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
         Height          =   3300
         Left            =   -74970
         TabIndex        =   19
         Top             =   315
         Width           =   3150
         Begin VB.TextBox Text3 
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
            Height          =   630
            Left            =   0
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   101
            TabStop         =   0   'False
            Text            =   "FrmTalieColetor.frx":11CE
            Top             =   15
            Width           =   3135
         End
         Begin VB.CommandButton cmd 
            Appearance      =   0  'Flat
            Caption         =   "NF"
            Height          =   350
            Index           =   12
            Left            =   2670
            MaskColor       =   &H00E3F7FF&
            Style           =   1  'Graphical
            TabIndex        =   97
            ToolTipText     =   "SALVA DADOS NO REGISTRO ATUAL"
            Top             =   2925
            Width           =   420
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00C1EAF9&
            Height          =   2325
            Left            =   0
            TabIndex        =   20
            Top             =   585
            Visible         =   0   'False
            Width           =   3135
            Begin VB.TextBox Text2 
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
               Left            =   975
               TabIndex        =   28
               Top             =   180
               Width           =   990
            End
            Begin VB.CommandButton Command2 
               Appearance      =   0  'Flat
               Height          =   435
               Left            =   2610
               MaskColor       =   &H00E3F7FF&
               Picture         =   "FrmTalieColetor.frx":11D8
               Style           =   1  'Graphical
               TabIndex        =   27
               ToolTipText     =   "Próximo"
               Top             =   120
               Width           =   420
            End
            Begin VB.OptionButton Option4 
               BackColor       =   &H00C1EAF9&
               Caption         =   "Tot"
               Height          =   195
               Index           =   3
               Left            =   45
               TabIndex        =   26
               Top             =   2100
               Width           =   570
            End
            Begin VB.OptionButton Option4 
               BackColor       =   &H00C1EAF9&
               Caption         =   "Av"
               Height          =   195
               Index           =   2
               Left            =   1335
               TabIndex        =   25
               Top             =   2100
               Width           =   570
            End
            Begin VB.OptionButton Option4 
               BackColor       =   &H00C1EAF9&
               Caption         =   "Fal"
               Height          =   195
               Index           =   1
               Left            =   1920
               TabIndex        =   24
               Top             =   2100
               Width           =   570
            End
            Begin VB.OptionButton Option4 
               BackColor       =   &H00C1EAF9&
               Caption         =   "Acr"
               Height          =   195
               Index           =   0
               Left            =   690
               TabIndex        =   23
               Top             =   2100
               Width           =   570
            End
            Begin VB.OptionButton Option4 
               BackColor       =   &H00C1EAF9&
               Caption         =   "Sld"
               Height          =   195
               Index           =   4
               Left            =   2505
               TabIndex        =   22
               Top             =   2100
               Value           =   -1  'True
               Width           =   570
            End
            Begin VB.TextBox Text2 
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
               Index           =   13
               Left            =   2010
               TabIndex        =   21
               Text            =   "1"
               Top             =   180
               Width           =   225
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
               Height          =   915
               Left            =   45
               TabIndex        =   29
               Top             =   555
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   1614
               _Version        =   393216
               AllowBigSelection=   0   'False
               AllowUserResizing=   3
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
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nota Fiscal"
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
               Left            =   120
               TabIndex        =   30
               Top             =   210
               Width           =   780
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00C1EAF9&
            Height          =   1905
            Left            =   0
            TabIndex        =   37
            Top             =   585
            Visible         =   0   'False
            Width           =   3135
            Begin VB.TextBox Text2 
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
               Index           =   2
               Left            =   795
               TabIndex        =   44
               Top             =   1170
               Width           =   810
            End
            Begin VB.TextBox Text2 
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
               Index           =   12
               Left            =   795
               TabIndex        =   41
               Top             =   495
               Width           =   855
            End
            Begin VB.TextBox Text2 
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
               Index           =   10
               Left            =   2070
               TabIndex        =   40
               Top             =   135
               Width           =   495
            End
            Begin VB.TextBox Text2 
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
               Index           =   9
               Left            =   1425
               TabIndex        =   39
               Top             =   135
               Width           =   480
            End
            Begin VB.TextBox Text2 
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
               Index           =   8
               Left            =   780
               TabIndex        =   38
               Top             =   135
               Width           =   495
            End
            Begin VB.TextBox Text2 
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
               Height          =   300
               Index           =   14
               Left            =   795
               TabIndex        =   43
               Top             =   825
               Width           =   1785
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Qtde :"
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
               Left            =   375
               TabIndex        =   98
               Top             =   1215
               Width           =   405
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Peso :"
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
               Index           =   16
               Left            =   390
               TabIndex        =   48
               Top             =   525
               Width           =   390
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "X"
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
               Index           =   14
               Left            =   1305
               TabIndex        =   47
               Top             =   195
               Width           =   75
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "X"
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
               Index           =   13
               Left            =   1950
               TabIndex        =   46
               Top             =   210
               Width           =   75
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "C x L x A :"
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
               Index           =   9
               Left            =   120
               TabIndex        =   45
               Top             =   195
               Width           =   660
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Produto :"
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
               Index           =   2
               Left            =   210
               TabIndex        =   42
               Top             =   900
               Width           =   570
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C1EAF9&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   60
            TabIndex        =   31
            Top             =   2940
            Width           =   1545
            Begin VB.CommandButton cmd 
               Appearance      =   0  'Flat
               Caption         =   "G"
               Enabled         =   0   'False
               Height          =   350
               Index           =   7
               Left            =   1215
               MaskColor       =   &H00E3F7FF&
               Style           =   1  'Graphical
               TabIndex        =   36
               ToolTipText     =   "SALVA DADOS NO REGISTRO ATUAL"
               Top             =   0
               Width           =   285
            End
            Begin VB.CommandButton cmd 
               Appearance      =   0  'Flat
               Caption         =   "N"
               Height          =   350
               Index           =   6
               Left            =   0
               MaskColor       =   &H00E3F7FF&
               Style           =   1  'Graphical
               TabIndex        =   35
               Tag             =   "1"
               ToolTipText     =   "INSERE NOVO REGISTRO"
               Top             =   0
               Width           =   285
            End
            Begin VB.CommandButton cmd 
               Appearance      =   0  'Flat
               Caption         =   "E"
               Height          =   350
               Index           =   9
               Left            =   600
               MaskColor       =   &H00E3F7FF&
               Style           =   1  'Graphical
               TabIndex        =   34
               Tag             =   "3"
               ToolTipText     =   "ELIMINA REGISTRO ATUAL"
               Top             =   0
               Width           =   285
            End
            Begin VB.CommandButton cmd 
               Appearance      =   0  'Flat
               Caption         =   "A"
               Height          =   350
               Index           =   8
               Left            =   300
               MaskColor       =   &H00E3F7FF&
               Style           =   1  'Graphical
               TabIndex        =   33
               Tag             =   "2"
               ToolTipText     =   "EDITA REGISTRO ATUAL"
               Top             =   0
               Width           =   285
            End
            Begin VB.CommandButton cmd 
               Appearance      =   0  'Flat
               Caption         =   "C"
               Enabled         =   0   'False
               Height          =   350
               Index           =   10
               Left            =   900
               MaskColor       =   &H00E3F7FF&
               Style           =   1  'Graphical
               TabIndex        =   32
               ToolTipText     =   "DESFAZ ALTERAÇÕES PENDENTES"
               Top             =   0
               Width           =   285
            End
         End
         Begin VB.Label LbItem 
            BackColor       =   &H00E0E0E0&
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
            Left            =   8805
            TabIndex        =   50
            Top             =   120
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label LbNFI 
            BackColor       =   &H00E0E0E0&
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
            Left            =   8040
            TabIndex        =   49
            Top             =   180
            Visible         =   0   'False
            Width           =   555
         End
      End
      Begin VB.Frame frameitens 
         BackColor       =   &H00C1EAF9&
         Caption         =   "Itens"
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
         Height          =   1365
         Left            =   -74985
         TabIndex        =   17
         Top             =   2325
         Width           =   3180
         Begin MSDataGridLib.DataGrid DataGrid2 
            Bindings        =   "FrmTalieColetor.frx":1942
            Height          =   1125
            Left            =   30
            TabIndex        =   18
            Top             =   180
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   1984
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   13
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
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   14
            BeginProperty Column00 
               DataField       =   "num_nf"
               Caption         =   "NF"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "serie_nf"
               Caption         =   "SR"
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
               DataField       =   "item"
               Caption         =   "Item"
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
               DataField       =   "Qtde_descarga"
               Caption         =   "Descarga"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column04 
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
            BeginProperty Column05 
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
            BeginProperty Column06 
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
            BeginProperty Column07 
               DataField       =   "desc_produto"
               Caption         =   "Produto"
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
               DataField       =   "tipo_descarga"
               Caption         =   "Tipo Descarga"
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
               DataField       =   "diferenca"
               Caption         =   "Diferença"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column10 
               DataField       =   "comprimento"
               Caption         =   "Comp."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column11 
               DataField       =   "largura"
               Caption         =   "Larg."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column12 
               DataField       =   "altura"
               Caption         =   "Alt."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column13 
               DataField       =   "peso"
               Caption         =   "Peso"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   1
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   884,976
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   299,906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   524,976
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   764,787
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   915,024
               EndProperty
               BeginProperty Column05 
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   2039,811
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   1934,929
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   1184,882
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   989,858
               EndProperty
               BeginProperty Column10 
               EndProperty
               BeginProperty Column11 
               EndProperty
               BeginProperty Column12 
               EndProperty
               BeginProperty Column13 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame framenota 
         BackColor       =   &H00C1EAF9&
         Caption         =   "Talie"
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
         Height          =   2010
         Left            =   -74985
         TabIndex        =   15
         Top             =   315
         Width           =   3165
         Begin MSAdodcLib.Adodc Dc_Talie 
            Height          =   330
            Left            =   1665
            Top             =   810
            Visible         =   0   'False
            Width           =   2355
            _ExtentX        =   4154
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
            Caption         =   "Dc_Talie"
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
            Bindings        =   "FrmTalieColetor.frx":1957
            Height          =   1755
            Left            =   45
            TabIndex        =   16
            Top             =   195
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   3096
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   13
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
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "Inicio"
               Caption         =   "Inicio"
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
               DataField       =   "Placa"
               Caption         =   "Placa"
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
            BeginProperty Column03 
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
            BeginProperty Column04 
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
            BeginProperty Column05 
               DataField       =   "Fantasia"
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
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1695,118
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1184,882
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739,906
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1800
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1769,953
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   3555,213
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc dcExportador 
            Height          =   330
            Left            =   2400
            Top             =   1305
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
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
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3480
      TabIndex        =   13
      Top             =   2640
      Width           =   2010
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   4
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3885
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.TextBox Text2 
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
      Height          =   315
      Index           =   5
      Left            =   945
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3975
      Width           =   975
   End
   Begin VB.TextBox Text2 
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
      Height          =   315
      Index           =   6
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4770
      Width           =   1800
   End
   Begin VB.TextBox Text2 
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
      Height          =   315
      Index           =   7
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   5100
      Width           =   1110
   End
   Begin VB.TextBox Text2 
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
      Height          =   315
      Index           =   11
      Left            =   1275
      TabIndex        =   3
      Top             =   5490
      Width           =   1110
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C1EAF9&
      Caption         =   "CrossDocking"
      Height          =   195
      Left            =   3525
      TabIndex        =   2
      Top             =   405
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Height          =   465
      Index           =   5
      Left            =   3270
      MaskColor       =   &H00E3F7FF&
      Picture         =   "FrmTalieColetor.frx":196E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "SAIR DO MÓDULO"
      Top             =   4365
      Width           =   465
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4845
      Top             =   975
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   420
      Left            =   4680
      Top             =   1395
      Visible         =   0   'False
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   741
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   4515
      Top             =   1980
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   330
      Left            =   3450
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
      Caption         =   "Adodc5"
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
      Left            =   3240
      Top             =   2130
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
      Left            =   3360
      Top             =   1740
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde. NF"
      Height          =   195
      Index           =   0
      Left            =   195
      TabIndex        =   11
      Top             =   4005
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde.Descarga"
      Height          =   195
      Index           =   1
      Left            =   225
      TabIndex        =   10
      Top             =   4350
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Descarga"
      Height          =   195
      Index           =   3
      Left            =   150
      TabIndex        =   9
      Top             =   4800
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Diferença"
      Height          =   195
      Index           =   4
      Left            =   330
      TabIndex        =   8
      Top             =   5175
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estufagem"
      Height          =   195
      Index           =   5
      Left            =   135
      TabIndex        =   7
      Top             =   5550
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Término"
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
      Index           =   1
      Left            =   3435
      TabIndex        =   1
      Top             =   3240
      Width           =   570
   End
End
Attribute VB_Name = "FrmTalieColetor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Nota_item As Boolean
Dim rs_aux As New ADODB.Recordset
Dim TxtReg As Integer
Dim ACRESC As Integer
Dim FALTA As Integer
Dim Novo As Boolean
Sub Atualiza_Item()
        sql = "SELECT a.codproduto,a.obs,b.qtde,a.qtde_estufagem,a.autonum_ti,c.num_nf,c.serie_nf,d.lote,a.qtde_descarga,a.tipo_descarga, a.diferenca"
        sql = sql & ", f.descricao_emb, e.desc_produto, d.instrucao, b.item, a.comprimento, a.largura, a.altura, a.peso"
        sql = sql & " FROM"
        sql = sql & " " & Esquema_Op & ".tb_talie_item a "
        sql = sql & " inner join " & Esquema_Op & ".tb_notas_itens b on a.autonum_nfi = b.autonum_nfi"
        sql = sql & " inner join " & Esquema_Op & ".tb_notas_fiscais c on b.autonum_nf = c.autonum_nf"
        sql = sql & " inner join " & Esquema_Op & ".tb_booking_carga d on b.autonum_bcg = d.autonum_bcg"
        sql = sql & " inner join " & Esquema_Op & ".tb_cad_produtos e on d.autonum_pro = e.autonum_pro"
        sql = sql & " inner join " & Esquema_Op & ".tb_cad_embalagens f on d.autonum_emb = f.autonum_emb"
        sql = sql & " inner join " & Esquema_Op & ".tb_talie g on a.autonum_talie = g.autonum_talie"
        sql = sql & " where a.autonum_talie=" & Nnull(LbTalie, 0)
'        sql = sql & " and g.flag_descarga=1"
        sql = sql & " order by d.instrucao,c.num_nf, b.item"
        
        Adodc6.ConnectionString = ConnectionString
        Adodc6.RecordSource = UCase(sql)
        Adodc6.Refresh
        
        
End Sub


Private Sub Limpa(Tela)
    
    Select Case Tela
        Case 1
            MskData(0) = "__/__/____ __:__"
            MskData(1) = "__:__"
            MskPlaca = "___-____"
            Option1.Value = 0
            Option2.Value = 0
            Option3.Value = 0
            Check1.Value = 0
            DataCombo1.BoundText = ""
            DataCombo2.BoundText = ""
            MskCntr = "__________-_"
            MskCntr.Tag = ""
            LbTalie = ""
            MskCntr = "__________-_"
            
            DCombo2(2).BoundText = ""
            MskCntr.Tag = ""
            
            DcReserva.BoundText = ""
            Combo1.ListIndex = -1
        Case 2
'            Text2(0) = ""
            Text2(13) = ""
            Text2(14) = ""
            Text2(2) = ""
            LbItem = ""
            Grid1.Clear
    End Select
    
End Sub


Sub Monta_Combo_Reservas(placa, Data As String)

        sql = "select e.autonum_parceiro,e.reference, a.autonum, e.autonum_boo "
        sql = sql & " from " & Esquema_Op & ".tb_gate_new a"
        sql = sql & " inner join " & Esquema_Op & ".tb_registro b on a.autonum = b.autonum_gate"
        sql = sql & " inner join " & Esquema_Op & ".tb_booking e on b.autonum_boo = e.autonum_boo"
        sql = sql & " where a.placa='" & placa & "' AND"
        
        sql = sql & " ("
        sql = sql & " TO_CHAR(a.dt_gate_in,'DD/MM/YYYY') = '" & Data & "'"
        sql = sql & " or"
        sql = sql & " TO_char(a.dt_gate_in,'DD/MM/YYYY') = '" & Format(DateAdd("d", -1, CDate(Data)), "DD/MM/YYYY") & "'"
        sql = sql & " )"
        
        Adodc4.ConnectionString = ConnectionString
        Adodc4.RecordSource = UCase(sql)
        Adodc4.Refresh


End Sub

Private Sub Adodc6_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)



    Limpa (2)
    If Not Adodc6.Recordset.EOF Then
        Me.Frame4.Visible = False
        Me.Frame6.Visible = True
        Me.Frame6.Enabled = True
        
'        Frame6.Width = 11445
'        Frame6.Height = 1785
'        Frame6.Left = 120
'        Frame6.Top = 225
        
        Me.Frameitem.Visible = True
        Me.Frameitem.Enabled = True
    
        Text2(5) = Adodc6.Recordset!qtde
        Text2(2) = Nnull(Adodc6.Recordset!qtde_descarga, 0)
        Text2(6) = Nnull(Adodc6.Recordset!tipo_descarga, 1)
'        Text2(3) = Nnull(Adodc6.Recordset!obs, 1)
        Text2(14) = Nnull(Adodc6.Recordset!CODPRODUTO, 1)
    
        Text2(8) = Nnull(Adodc6.Recordset!Comprimento, 0)
        Text2(9) = Nnull(Adodc6.Recordset!Largura, 0)
        Text2(10) = Nnull(Adodc6.Recordset!Altura, 0)
        Text2(12) = Nnull(Adodc6.Recordset!Peso, 0)
        Text2(11) = Nnull(Adodc6.Recordset!qtde_estufagem, 0)
    
        LbItem = Adodc6.Recordset!autonum_ti
    
    Else
        Me.Frame6.Visible = False
        Me.Frame4.Visible = True
    End If

End Sub

Private Sub Cmd_Click(Index As Integer)

Select Case Index

    Case 0
        Limpa (1)
        Limpa (2)
        Grid1.Enabled = False
        
        FrameTalie.Enabled = True
        Frameitem.Enabled = False
        Frame3.Enabled = False
        MskData(0) = Format(Now, "dd/mm/yyyy hh:mm")
        Seta_Controles
        CmdFinaliza.Enabled = False
        Combo1.ListIndex = 1
    
        If DD Then
            Option1.Value = True
        Else
            Option2.Value = True
        End If
        
        TxtRegistro = ""
        TxtRegistro.SetFocus
    
    Case 1
        If Not Dc_Talie.Recordset.EOF Then
            Frame3.Enabled = True
            FrameTalie.Enabled = True
            Frameitem.Enabled = True
            Seta_Controles
        
            CmdFinaliza.Enabled = True
            If LbTalie <> "" Then
                If Nnull(DB.Execute("select flag_fechado from " & Esquema_Op & ".tb_talie where autonum_talie=" & LbTalie).Fields(0), 0) = 1 Then
                    CmdFinaliza.Enabled = False
                End If
            End If
        Else
            MsgBox "Favor selecionar um registro para edição.", vbInformation, "Talie"
            FrameTalie.Enabled = False
            Frameitem.Enabled = False
            Exit Sub
        End If
    Case 2
        If Combo1 = "" Then
            MsgBox "Informe o modo de operação"
            Combo1.SetFocus
            Exit Sub
        End If
        If MskData(0).ClipText = "" Then
            MsgBox "Favor informar a data de inicio"
            MskData(0).SetFocus
            Exit Sub
        End If
'        If MskData(1).ClipText = "" Then
'            MsgBox "Favor informar a data de término"
'            MskData(1).SetFocus
'            Exit Sub
'        End If
        If Not IsDate(MskData(0)) Then
            MsgBox "Data Inválida"
            MskData(0).SetFocus
            Exit Sub
        End If
'        If MskData(1).ClipText <> "" Then
'            If Not IsDate(MskData(1)) Then
'                MsgBox "Data Inválida"
'                MskData(0).SetFocus
'                Exit Sub
'            End If
'
'            If CDate(MskData(1)) < CDate(MskData(0)) Then
'                MsgBox "Período Inválido"
'                MskData(0).SetFocus
'                Exit Sub
'            End If
'        End If
        
        If Not Option1.Value And Not Option2.Value And Not Option3.Value Then
            MsgBox "Selecione o tipo de operação"
            Exit Sub
        End If
        If DataCombo1.BoundText = "" Then
            MsgBox "Conferente não Informado"
            DataCombo1.SetFocus
            Exit Sub
        End If
        If DataCombo2.BoundText = "" Then
            MsgBox "Equipe não Informada"
            DataCombo2.SetFocus
            Exit Sub
        End If
        If Option1.Value = True Then
            If MskCntr.ClipText = "" Then
                MsgBox "Conteiner não informado"
                MskCntr.SetFocus
                Exit Sub
            End If
        End If
        If DCombo2(2).BoundText = "" Then
            MsgBox "Cliente não informado"
            If DCombo2(2).Enabled = True Then DCombo2(2).SetFocus
            Exit Sub
        End If
        If DcReserva.BoundText = "" Then
            MsgBox "Reserva não Identificada"
            Exit Sub
        End If
            
        If MskPlaca.ClipText = "" Then
            MsgBox "Informe a Placa do veicul descarregado"
            Exit Sub
        End If
        
        If LbTalie = "" Then
            
            sql = "select autonum_talie from " & Esquema_Op & ".tb_talie where autonum_boo=" & DcReserva.BoundText & " and placa='" & MskPlaca & "' and nvl(flag_fechado,0)=0"
            PRSet Rs, sql
            If Not Rs.EOF Then
                MsgBox "Já existe talie para esta placa"
                Rs.Close
                Exit Sub
            End If
            
            sql = "select redex.seq_talie.nextval from dual"
            PRSet Rs, sql
            TxtRegistro = Rs.Fields(0)
            LbTalie.Caption = TxtRegistro
            LbTalie.Refresh
            Rs.Close
            
            
            sql = "Insert into " & Esquema_Op & ".tb_talie (autonum_talie,"
            sql = sql & "autonum_patio,inicio,termino,flag_descarga,flag_estufagem,flag_carregamento"
            sql = sql & ",crossdocking,conferente,equipe,autonum_boo,forma_operacao,placa,AUTONUM_GATE,FLAG_COMPLETO,AUTONUM_REG"
            sql = sql & ") values ( " & TxtRegistro & ","
            If Option1.Value = True Then
                sql = sql & MskCntr.Tag
            Else
                sql = sql & "NULL"
            End If
            sql = sql & ",to_date('" & Format(MskData(0), "dd/mm/yyyy hh:mm") & "','dd/mm/yyyy HH24:MI')"
'            If MskData(1).ClipText <> "" Then
'                sql = sql & ",'" & Format(MskData(1), "yyyy-mm-dd hh:mm") & "'"
'            Else
                sql = sql & ",NULL"
'            End If
            sql = sql & "," & IIf(Option2.Value, 1, 0)
            sql = sql & "," & IIf(Option1.Value, 1, 0)
            sql = sql & "," & IIf(Option3.Value, 1, 0)
            sql = sql & "," & Check1.Value
            sql = sql & "," & DataCombo1.BoundText
            sql = sql & "," & DataCombo2.BoundText
            sql = sql & "," & DcReserva.BoundText
            sql = sql & ",'" & UCase(Left(Combo1, 1)) & "'"
            If MskPlaca.ClipText <> "" Then
                sql = sql & ",'" & MskPlaca & "'"
                sql = sql & "," & MskPlaca.Tag
            Else
                sql = sql & ",null"
                sql = sql & ",0"
            End If
            sql = sql & "," & Check2.Value
            sql = sql & "," & Nnull(TxtRegistro, 0)
            sql = sql & ")"
            DB.Execute sql
            

            FrameTalie.Enabled = True
            Frameitem.Enabled = True
            Dc_Talie.Refresh
            Dc_Talie.Recordset.Find "autonum_talie =" & TxtRegistro
            
            LbTalie.Caption = TxtRegistro
            LbTalie.Refresh
            
            Seta_Controles
        
        Else
            
            
            If Dc_Talie.Recordset!FLAG_FECHADO = 1 Then
                MsgBox "Talie Fechado, alteração não permitida"
                Exit Sub
            End If
            
            Rec = Nnull(LbTalie, 0)
            
            sql = "update " & Esquema_Op & ".tb_talie set"
            If Option1.Value = True Then
                sql = sql & " autonum_patio=" & MskCntr.Tag
            Else
                sql = sql & " autonum_patio=0"
            End If
            sql = sql & ",inicio=to_date('" & Format(MskData(0), "dd/mm/yyyy hh:mm") & "','dd/mm/yyyy hh24:mi')"
            
'            If MskData(1).ClipText <> "" Then
'                sql = sql & ",termino='" & Format(MskData(1), "yyyy-mm-dd hh:mm") & "'"
'            Else
'                sql = sql & ",termino=NULL"
'            End If
'            sql = sql & ",termino='" & Format(MskData(1), "yyyy-mm-dd hh:mm") & "'"
            
            sql = sql & ",flag_descarga=" & IIf(Option2.Value, 1, 0)
            sql = sql & ",flag_estufagem=" & IIf(Option1.Value, 1, 0)
            sql = sql & ",flag_carregamento=" & IIf(Option3.Value, 1, 0)
            sql = sql & ",crossdocking=" & Check1.Value
            sql = sql & ",FLAG_COMPLETO=" & Check2.Value
            sql = sql & ",conferente=" & DataCombo1.BoundText
            sql = sql & ",equipe=" & DataCombo2.BoundText
            sql = sql & ",autonum_boo=" & DcReserva.BoundText
            sql = sql & ",forma_operacao='" & UCase(Left(Combo1, 1)) & "'"
            sql = sql & ",AUTONUM_REG=" & Nnull(TxtRegistro, 0)
            sql = sql & " where autonum_talie = " & LbTalie
            DB.Execute UCase(sql)
            
            
            FrameTalie.Enabled = True
            Frameitem.Enabled = True
            Dc_Talie.Refresh
            Dc_Talie.Recordset.Find "autonum_talie =" & Rec
            Seta_Controles
                        
        End If
        
        If MskCntr.Tag <> "" And MskCntr.Tag <> "0" Then
            If Check2.Value = 1 Then
                sql = "update " & Esquema_Op & ".tb_patio set ef='F' where autonum_patio = " & MskCntr.Tag
            Else
                sql = "update " & Esquema_Op & ".tb_patio set ef='E' where autonum_patio = " & MskCntr.Tag
            End If
            DB.Execute UCase(sql)
        End If
        
        Call Atualiza_Item
        
        
        Call Evento_06(LbTalie, "O", "", "")
        Call Evento_07(LbTalie, "O")
        
    Case 3
        Limpa (1)
        Limpa (2)
        FrameTalie.Enabled = False
        Frameitem.Enabled = False
        Seta_Controles
        
    Case 4
        If Nnull(LbTalie, 0) = 0 Then
            MsgBox "Selecione um registro para exclusão"
            Exit Sub
        End If
        
        
        If Dc_Talie.Recordset!FLAG_FECHADO = 1 Then
            MsgBox "Não é possivel excluir talie com fechamento confirmado"
            Exit Sub
        End If
        
        
        If MsgBox("Deseja Excluir o talie ?", vbYesNo) = vbYes Then
            DB.Execute "Delete from " & Esquema_Op & ".tb_talie_item where autonum_talie = " & Nnull(LbTalie, 0)
            DB.Execute "Delete from " & Esquema_Op & ".tb_talie where autonum_talie = " & Nnull(LbTalie, 0)
            MsgBox "Registro Excluido!"
        End If
        Limpa (1)
        Limpa (2)
        Dc_Talie.Refresh

    Case 5
        Unload Me
    
    Case 6 'ADICIONAR ITEM
        If Val(LbTalie) = 0 Then
            MsgBox "Selecione um talie existente"
            Exit Sub
        End If
        
        sql = "select count(*) from " & Esquema_Op & ".tb_talie where autonum_talie=" & LbTalie & " and flag_fechado =1"
        sql = UCase(sql)
        If DB.Execute(sql).Fields(0) <> 0 Then
            MsgBox "Talie Fechado - Operação Cancelada"
            Exit Sub
        End If
        
        sql = "select count(*) from " & Esquema_Op & ".tb_talie"
        sql = sql & " where autonum_talie=" & LbTalie
        sql = sql & " and not termino is null"
        sql = UCase(sql)
        If DB.Execute(sql).Fields(0) <> 0 Then
            If MsgBox("Consta término para este talie" & vbCr & "Deseja reabrir ?", vbQuestion + vbYesNo) = vbNo Then
                MsgBox "Operação cancelada"
                Exit Sub
            Else
                sql = "update " & Esquema_Op & ".tb_talie set termino = null"
                sql = sql & " where autonum_talie=" & LbTalie
                sql = UCase(sql)
                DB.Execute sql
            End If
        End If

        Limpa (2)
        FrameTalie.Enabled = False
        
        Frame6.Visible = False
        Frame4.Visible = True
        Frame4.Enabled = True
        
'        Frame4.Left = 0
'        Frame4.Top = 0
'        Frame4.Width = 11775
'        Frame4.Height = 2610
        
        Text2(0).SetFocus
        
        cmd(6).Enabled = False
        cmd(8).Enabled = False
        cmd(9).Enabled = False
        cmd(10).Enabled = True
        cmd(7).Enabled = True
        Novo = True
        
    Case 7 'GRAVA O ITEM
        If Nnull(LbTalie, 0) = 0 Then
            MsgBox "Um talie deve estar selecionado"
            Exit Sub
        End If



        sql = "select count(*) from " & Esquema_Op & ".tb_talie where autonum_talie=" & LbTalie & " and flag_fechado =1"
        sql = UCase(sql)
        If DB.Execute(sql).Fields(0) <> 0 Then
            MsgBox "Talie Fechado - Operação Cancelada"
            Exit Sub
        End If

'        If Text2(0) = "" Then
'            MsgBox "Nota Fiscal não informada"
'            Text2(0).SetFocus
'            Exit Sub
'        End If
'        If Text2(1) = "" Then
'            MsgBox "Qtde Descarga não informada"
'            Text2(1).SetFocus
'            Exit Sub
'        End If
'
'        If Val(Text2(3)) = 0 And Option1.Value = True Then
'            MsgBox "Quantidade Estufada não informada"
'            Text2(3).SetFocus
'            Exit Sub
'        End If
'
'        If Option1.Value = True And Val(Text2(3)) > Val(Text2(1)) Then
'            MsgBox "Quantidade estufada inválida"
'            Text2(3).SetFocus
'            Exit Sub
'        End If

'        If Dc_Talie.Recordset!FLAG_FECHADO = 1 Then
'            MsgBox "Talie Fechado, alteração não permitida"
'            Exit Sub
'        End If
            
'        If CargaSuzano And TxtBarra = "" Then
'            MsgBox "Carga Suzano - Obrigatorio validar carga"
'            TxtBarra.SetFocus
'            Exit Sub
'        End If
           
        Dim Chave As String
        If Nnull(LbItem, 0) = 0 Then
            For I = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(I, 1) = "" Then
                    MsgBox "Produto não informado - item " & I
                    Grid1.Col = 1
                    Grid1.Row = I
                    Grid1.TextMatrix(I, 1) = ""
                    Exit Sub
                End If
                
                Chave = Grid1.TextMatrix(I, 1)
                If CargaSuzano Then
                    sql = "select count(*) from " & Esquema_Op & ".tb_integra_reg ir"
                    sql = sql & " inner join " & Esquema_Op & ".tb_integra_carga ic on ir.id_reg = ic.id_reg"
                    sql = sql & " where"
                    sql = sql & " ir.autonum_boo = " & Val(Text2(0).Tag)
                    sql = sql & " and ic.codbarra='" & Chave & "'"
                    sql = UCase(sql)
                    If Nnull(DB.Execute(sql).Fields(0), 0) = 0 Then
                        MsgBox "Produto não encontrado como carga esperada" & vbCr & "Verifique com o operacional"
                        TxtBarra = ""
                        Exit Sub
                    End If
                End If
                
                sql = "select count(*) from " & Esquema_Op & ".tb_talie_item"
                sql = sql & " where codproduto = '" & Chave & "'"
                If DB.Execute(UCase(sql)).Fields(0) <> 0 Then
                    MsgBox "Produto já cadastrado - item " & I
                    Grid1.Col = 1
                    Grid1.Row = I
                    Grid1.TextMatrix(I, 1) = ""
                    Exit Sub
                End If
            Next
        Else
            Chave = Text2(14)
            If CargaSuzano Then
                sql = "select count(*) from " & Esquema_Op & ".tb_integra_reg ir"
                sql = sql & " inner join " & Esquema_Op & ".tb_integra_carga ic on ir.id_reg = ic.id_reg"
                sql = sql & " where"
                sql = sql & " ir.autonum_boo = " & Val(Text2(0).Tag)
                sql = sql & " and ic.codbarra='" & Chave & "'"
                sql = UCase(sql)
                If Nnull(DB.Execute(sql).Fields(0), 0) = 0 Then
                    MsgBox "Produto não encontrado como carga esperada" & vbCr & "Verifique com o operacional"
                    TxtBarra = ""
                    Exit Sub
                End If
            End If
            
            sql = "select count(*) from " & Esquema_Op & ".tb_talie_item"
            sql = sql & " where codproduto = '" & Chave & "'"
            sql = sql & " and autonum_ti<>" & LbItem
            If DB.Execute(UCase(sql)).Fields(0) <> 0 Then
                MsgBox "Produto já cadastrado - item " & I
                Grid1.Col = 1
                Grid1.Row = I
                Grid1.TextMatrix(I, 1) = ""
                Exit Sub
            End If
        End If
                
        If Nnull(LbItem, 0) = 0 Then
            For I = 1 To Grid1.Rows - 1
                If Option1.Value = True Then
                    If Val(Nnull(Replace(Grid1.TextMatrix(I, 16), ",", "."), 0)) <> 0 And Val(Nnull(Replace(Grid1.TextMatrix(I, 16), ",", "."), 0)) > Val(Grid1.TextMatrix(I, 8)) Then
                        MsgBox "Qtde de Estufagem não pode ser superior a qtde. descarregada"
                        Exit Sub
                    End If
                    If Val(Nnull(Replace(Grid1.TextMatrix(I, 16), ",", "."), 0)) = 0 Then
                        Grid1.TextMatrix(I, 16) = Grid1.TextMatrix(I, 8)
                    End If
                Else
                    If Val(Nnull(Replace(Grid1.TextMatrix(I, 16), ",", "."), 0)) <> 0 And Val(Nnull(Replace(Grid1.TextMatrix(I, 16), ",", "."), 0)) > Val(Grid1.TextMatrix(I, 8)) Then
                        MsgBox "Qtde de Estufagem não pode ser superior a qtde. descarregada"
                        Exit Sub
                    End If
                End If
            Next
            
            For I = 1 To Grid1.Rows - 1
                sql = "Insert into " & Esquema_Op & ".tb_talie_item (autonum_ti,CODPRODUTO,autonum_talie,autonum_nfi,qtde_descarga"
                sql = sql & ",tipo_descarga,diferenca,obs,qtde_disponivel,comprimento,largura,altura,peso,qtde_estufagem"
                sql = sql & ") values ( redex.seq_talie_item.nextval,"
                sql = sql & " '" & Nnull(Grid1.TextMatrix(I, 1), 1) & "'"
                sql = sql & "," & LbTalie
                sql = sql & "," & Grid1.TextMatrix(I, 2)
                sql = sql & "," & Val(Grid1.TextMatrix(I, 8))
                sql = sql & ",'" & Grid1.TextMatrix(I, 9) & "'"
                sql = sql & "," & Val(Grid1.TextMatrix(I, 10))
                sql = sql & ",'" & Grid1.TextMatrix(I, 11) & "'"
                sql = sql & "," & Val(Grid1.TextMatrix(I, 4))
                sql = sql & "," & Nnull(Replace(Grid1.TextMatrix(I, 12), ",", "."), 0)
                sql = sql & "," & Nnull(Replace(Grid1.TextMatrix(I, 13), ",", "."), 0)
                sql = sql & "," & Nnull(Replace(Grid1.TextMatrix(I, 14), ",", "."), 0)
                sql = sql & "," & Nnull(Replace(Grid1.TextMatrix(I, 15), ",", "."), 0)
                sql = sql & "," & Nnull(Replace(Grid1.TextMatrix(I, 16), ",", "."), 0)
                sql = sql & ")"
                DB.Execute UCase(sql)
            Next
        Else
            
            If Val(Text2(11)) <> 0 And Val(Text2(11)) > Val(Text2(0)) Then
                MsgBox "Qtde de Estufagem não pode ser superior a qtde. descarregada"
                Exit Sub
            End If
            
            
            sql = "update " & Esquema_Op & ".tb_talie_item set " 'OBS = '" & Text2(3) & "',"
            sql = sql & " comprimento=" & Nnull(Replace(Text2(8), ",", "."), 0)
            sql = sql & ",largura=" & Nnull(Replace(Text2(9), ",", "."), 0)
            sql = sql & ",altura=" & Nnull(Replace(Text2(10), ",", "."), 0)
            sql = sql & ",peso=" & Nnull(Replace(Text2(12), ",", "."), 0)
            sql = sql & ",qtde_estufagem=" & Nnull(Replace(Text2(11), ",", "."), 0)
            sql = sql & ",codproduto='" & Nnull(Text2(14), 1) & "'"
            sql = sql & ",qtde_descarga=" & Nnull(Text2(2), 0)
            sql = sql & " where autonum_ti=" & LbItem
            DB.Execute UCase(sql)
                               
        End If
        
        
        If CargaSuzano Then
            sql = "update " & Esquema_Op & ".tb_integra_carga set recebido=1 where codbarra='" & Chave & "'"
            DB.Execute UCase(sql)
        End If
        
'        If Nnull(ACRESC, 0) <> 0 Or Nnull(FALTA, 0) <> 0 Then
'            sql = "UPDATE " & Esquema_Op & ".TB_NOTAS_ITENS SET"
'            sql = sql & " QTDE_ACRESC = isnull(QTDE_ACRESC,0) + " & ACRESC
'            sql = sql & ",QTDE_FALTA = isnull(QTDE_FALTA,0) + " & FALTA
'            sql = sql & " WHERE AUTONUM_NFI = " & Nnull(DcItem.BoundText, 0)
'            DB.Execute UCase(sql)
'        End If
        
              
        Adodc6.Refresh
               
        
        Frame4.Visible = False
        cmd(6).Enabled = True
        cmd(8).Enabled = True
        cmd(9).Enabled = True
        cmd(10).Enabled = False
        cmd(7).Enabled = False
        
        If Novo Then
            Call Cmd_Click(6)
            Call Command2_Click
        Else
            Call Cmd_Click(10)
            SSTab1.Tab = 2
        End If

        
        'cmd(6).SetFocus
        
    Case 8
        Novo = False
        
        Frame6.Visible = True
        
        cmd(6).Enabled = False
        cmd(7).Enabled = True
        cmd(8).Enabled = False
        cmd(9).Enabled = False
        cmd(10).Enabled = True
'        Text2(0).SetFocus
    
    Case 9
        If Nnull(LbItem, 0) <> 0 Then
            If MsgBox("Deseja Excluir o Item?", vbYesNo) = vbYes Then
                    
                If CargaSuzano Then
                    sql = "select codproduto from " & Esquema_Op & ".tb_talie_item where autonum_ti = " & Nnull(LbItem, 0)
                    sql = UCase(sql)
                    Chave = Nnull(DB.Execute(sql).Fields(0), 1)
                    If Chave <> "" Then
                        sql = "update " & Esquema_Op & ".tb_integra_carga set recebido=0 where codbarra='" & Chave & "'"
                        DB.Execute UCase(sql)
                    End If
                End If
                    
                DB.Execute "Delete from " & Esquema_Op & ".tb_talie_item where autonum_ti = " & Nnull(LbItem, 0)
                    'atualiza acresc e falta
'                    DB.Execute "UPDATE " & Esquema_Op & ".TB_NOTAS_ITENS SET QTDE_ACRESC = 0 WHERE AUTONUM_NFI = " & LbNFI
'                    DB.Execute "UPDATE " & Esquema_Op & ".TB_NOTAS_ITENS SET QTDE_FALTA = 0 WHERE AUTONUM_NFI = " & LbNFI
                Call Atualiza_Item
                MsgBox "Registro Excluido!"
                Limpa (2)
            End If
        Else
            MsgBox "É Necessário haver um registro atual para excluir", vbInformation
        End If


    Case 10
        Limpa (2)
        cmd(6).Enabled = True
        cmd(7).Enabled = False
        cmd(8).Enabled = True
        cmd(9).Enabled = True
        cmd(10).Enabled = False

    Case 11

        Frameitem.Visible = False
        Frame1.Visible = True
        FiltroReserva(0).SetFocus

    Case 12
        Cmd_Click (10)
        Text2(0) = ""
        Call Cmd_Click(6)
        
        


End Select

End Sub

Private Sub cmdFiltro_Click()

Dc_Talie.ConnectionString = ConnectionString
sql = "SELECT c.id_conteiner,b.reference,b.instrucao,d.fantasia,b.autonum_parceiro,A.AUTONUM_REG"
sql = sql & ",a.AUTONUM_TALIE , a.AUTONUM_PATIO, a.Placa, a.Inicio, a.TERMINO, a.FLAG_DESCARGA, a.FLAG_ESTUFAGEM"
sql = sql & ",a.CROSSDOCKING, a.CONFERENTE ,a.EQUIPE, a.AUTONUM_BOO,a.FLAG_CARREGAMENTO,A.AUTONUM_GATE,A.FLAG_FECHADO,A.FLAG_COMPLETO, a.forma_operacao"
sql = sql & " FROM"
sql = sql & " " & Esquema_Op & ".tb_talie a inner join " & Esquema_Op & ".tb_booking b on a.autonum_boo = b.autonum_boo"
sql = sql & " left join " & Esquema_Op & ".tb_patio c on a.autonum_patio = c.autonum_patio"
sql = sql & " inner join " & Esquema_Op & ".tb_cad_parceiros d on b.autonum_parceiro = d.autonum"
sql = sql & " where a.autonum_talie <> 0 "
If Trim(FiltroReserva(0)) <> "" Then
  sql = sql & " and b.reference = '" & Trim(FiltroReserva(0)) & "'"
End If
If Trim(FiltroInstrucao(6)) <> "" Then
  sql = sql & " and b.instrucao = '" & Trim(FiltroInstrucao(6)) & "'"
End If
If FiltroCntr.ClipText <> "" Then
  sql = sql & " and c.id_conteiner = '" & FiltroCntr & "'"
End If
If MskFiltroPlaca.ClipText <> "" Then
  sql = sql & " and a.placa = '" & MskFiltroPlaca & "'"
End If


sql = sql & " and (a.flag_descarga=1 or a.crossdocking=1)"

sql = sql & " order by a.inicio desc, d.fantasia, c.id_conteiner, b.reference, b.instrucao"
Dc_Talie.RecordSource = UCase(sql)
Dc_Talie.Refresh

Frame1.Visible = False

End Sub

Private Sub CmdFinaliza_Click()

    If Val(LbTalie) <> 0 Then
        sql = "update " & Esquema_Op & ".tb_talie set termino=to_date('" & Format(Now, "dd/mm/yyyy hh:mm") & "','dd/mm/yyyy hh24:mi') where autonum_talie=" & LbTalie
        DB.Execute UCase(sql)
    Else
        MsgBox "Selecione um talie existente para fechamento"
    End If

End Sub

Private Sub Command1_Click()
Frame1.Visible = False
End Sub

Private Sub Command2_Click()
    If Text2(0) = "" Then
        MsgBox "Informe a NF"
        Text2(0).SetFocus
        Exit Sub
    End If
    
    sql = "SELECT count(*) FROM " & Esquema_Op & ".TB_NOTAS_FISCAIS A"
    sql = sql & " inner join " & Esquema_Op & ".tb_booking b on a.autonum_boo = b.autonum_boo"
    sql = sql & " WHERE A.NUM_NF = '" & Text2(0) & "' AND b.AUTONUM_parceiro = " & DCombo2(2).BoundText
    If Nnull(DB.Execute(UCase(sql)).Fields(0), 0) > 1 Then
        MsgBox "Atenção - Encontrada mais de uma NF com a mesma série"
    End If
        
    sql = "SELECT a.serie_nf FROM " & Esquema_Op & ".TB_NOTAS_FISCAIS A"
    sql = sql & " inner join " & Esquema_Op & ".tb_booking b on a.autonum_boo = b.autonum_boo"
    sql = sql & " WHERE A.NUM_NF = '" & Text2(0) & "' AND b.AUTONUM_parceiro = " & DCombo2(2).BoundText
        
    Text2(13) = Nnull(DB.Execute(UCase(sql)).Fields(0), 1)
    
    Dim Tb As New ADODB.Recordset
    sql = "SELECT a.AUTONUM_NF, a.autonum_boo FROM " & Esquema_Op & ".TB_NOTAS_FISCAIS A"
    sql = sql & " inner join " & Esquema_Op & ".tb_booking b on a.autonum_boo = b.autonum_boo"
    sql = sql & " WHERE A.NUM_NF = '" & Text2(0) & "' and a.serie_nf = '" & Text2(13) & "' AND b.AUTONUM_parceiro = " & DCombo2(2).BoundText
    PRSet Tb, sql
    If Tb.EOF Then
        MsgBox "Nota fiscal não encontrada no cadastro de NF.", vbCritical
        Text2(0).SetFocus
        Exit Sub
    End If
        
    Text2(0).Tag = Tb!autonum_boo
        
    Grid1.Enabled = True
    Grid1.Clear
    
    sql = Empty
    sql = sql & " SELECT '' AS CODPRODUTO,A.AUTONUM_NFI AS ID, A.ITEM "
    
    sql = sql & ", case when A.QTDE - nvl((SELECT SUM(QTDE_DESCARGA) FROM TB_TALIE_ITEM WHERE AUTONUM_NFI = A.AUTONUM_NFI),0) < 0 then 0 else A.QTDE - nvl((SELECT SUM(QTDE_DESCARGA) FROM TB_TALIE_ITEM WHERE AUTONUM_NFI = A.AUTONUM_NFI),0) end AS QTDE"
    
    sql = sql & ", C.DESCRICAO_EMB ,D.DESC_PRODUTO, e.instrucao "
    
'    sql = sql & ", case when A.QTDE - nvl((SELECT SUM(QTDE_DESCARGA) FROM TB_TALIE_ITEM WHERE AUTONUM_NFI = A.AUTONUM_NFI),0) < 0 then 0 else A.QTDE - nvl((SELECT SUM(QTDE_DESCARGA) FROM TB_TALIE_ITEM WHERE AUTONUM_NFI = A.AUTONUM_NFI),0) end AS DESCARGA"
    sql = sql & ", 1 AS DESCARGA"
    
'    sql = sql & ", 'TOTAL' as TIPO, 0 as DIFERENCA, '' AS OBS"
    sql = sql & ", 'SALDO' as TIPO"
    
'    sql = sql & ", 0 as DIFERENCA"
    sql = sql & ", case when A.QTDE - (nvl((SELECT SUM(QTDE_DESCARGA) FROM TB_TALIE_ITEM WHERE AUTONUM_NFI = A.AUTONUM_NFI),0)+1) < 0 then 0 else A.QTDE - (nvl((SELECT SUM(QTDE_DESCARGA) FROM TB_TALIE_ITEM WHERE AUTONUM_NFI = A.AUTONUM_NFI),0)+1) end AS DIFERENCA"
    
    sql = sql & ", '' AS OBS"
    
    sql = sql & ", 0 as comprimento, 0 as largura, 0 as altura"
    
    sql = sql & ", (A.PESO_BRUTO / A.QTDE) as peso"
    
    If Option1.Value = True Then
        sql = sql & ", case when A.QTDE - nvl((SELECT SUM(QTDE_DESCARGA) FROM TB_TALIE_ITEM WHERE AUTONUM_NFI = A.AUTONUM_NFI),0) < 0 then 0 else A.QTDE - nvl((SELECT SUM(QTDE_DESCARGA) FROM TB_TALIE_ITEM WHERE AUTONUM_NFI = A.AUTONUM_NFI),0) end AS ESTUFAGEM"
    Else
        sql = sql & ", 0 as ESTUFAGEM"
    End If
    sql = sql & ",nvl(a.peso_bruto,0) AS peso_bruto"
    sql = sql & ",A.QTDE"
    
    sql = sql & " FROM " & Esquema_Op & ".TB_NOTAS_FISCAIS B "
    sql = sql & " LEFT JOIN " & Esquema_Op & ".TB_NOTAS_ITENS A ON A.AUTONUM_NF = B.AUTONUM_NF"
    sql = sql & " LEFT OUTER JOIN " & Esquema_Op & ".TB_CAD_EMBALAGENS C ON A.AUTONUM_EMB = C.AUTONUM_EMB"
    sql = sql & " LEFT OUTER JOIN " & Esquema_Op & ".TB_CAD_PRODUTOS D ON A.AUTONUM_PRO = D.AUTONUM_PRO"
    sql = sql & " LEFT OUTER JOIN " & Esquema_Op & ".TB_booking_carga e ON A.AUTONUM_bcg = e.AUTONUM_bcg"
    sql = sql & " WHERE B.NUM_NF = '" & Text2(0) & "' and b.serie_nf = '" & Text2(13) & "' AND B.AUTONUM_BOO = " & Text2(0).Tag
    Adodc1.ConnectionString = ConnectionString
    Adodc1.RecordSource = UCase(sql)
    Adodc1.Refresh
    L = 0
    Grid1.FixedRows = 0
    Grid1.Clear
    Grid1.Rows = 1
    Grid1.Cols = 19
    If Not Adodc1.Recordset.EOF Then
        If IsNull(Adodc1.Recordset!Id) Then
            MsgBox "Atenção - Não foi encontrado item para NF informada"
        End If
    End If
    Do While Not Adodc1.Recordset.EOF
        L = L + 1
        Grid1.AddItem ""
        For I = 0 To Adodc1.Recordset.Fields.Count - 1
            Grid1.TextMatrix(L, I + 1) = Nnull(Adodc1.Recordset.Fields(I), 1)
        Next
        Adodc1.Recordset.MoveNext
    Loop
    Adodc1.Recordset.Close
        
    Grid1.FixedRows = 1
    sql = "|COD.PRODUTO|ID|^ITEM|>QTDE|<EMBALAGEM|<PRODUTO|<INSTRUÇÃO|>QTDE DESCARGA|^TIPO DESCARGA|>DIFERENÇA|<OBSERVAÇÃO|>COMPRIMENTO|>LARGURA|>ALTURA|>PESO|>ESTUFAGEM|>PESO TOTAL|>QTDE. TOTAL"
    Grid1.FormatString = sql
    Grid1.ColWidth(0) = 200
    Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 0
    Grid1.ColWidth(3) = 400
    Grid1.ColWidth(4) = 800
    Grid1.ColWidth(5) = 1700
    Grid1.ColWidth(6) = 1700
    Grid1.ColWidth(7) = 1400
    Grid1.ColWidth(8) = 1400
    Grid1.ColWidth(9) = 1400
    Grid1.ColWidth(10) = 900
    Grid1.ColWidth(11) = 5000
    Grid1.ColWidth(12) = 1200
    Grid1.ColWidth(13) = 900
    Grid1.ColWidth(14) = 900
    Grid1.ColWidth(15) = 900
    Grid1.ColWidth(16) = 1200
    Grid1.ColWidth(17) = 1200
    Grid1.ColWidth(18) = 1200
    
    
    Option4(0).Enabled = False
    Option4(1).Enabled = False
    Option4(2).Enabled = False
    Option4(3).Enabled = False
    Option4(4).Enabled = False
    
    If Grid1.Rows > 1 Then
        Grid1.Row = 1
        Grid1.Col = 1
        SendKeys "{TAB}"
    End If
    
    

End Sub

Private Sub Command3_Click()

Resp = InputBox("Senha")
If Resp <> "MLED" Then
    MsgBox "Senha inválida"
    Exit Sub
End If

Dim Rs1 As New ADODB.Recordset


sql = "SELECT COUNT(*) FROM " & Esquema_Op & ".TB_TALIE_ITEM WHERE AUTONUM_TALIE=" & LbTalie & " AND not CODPRODUTO IS NULL"
PRSet Rs1, sql
qt = Nnull(Rs1.Fields(0), 0)
Rs1.Close

sql = "SELECT COUNT(*) FROM " & Esquema_Op & ".TB_PATIO_CS WHERE TALIE_DESCARGA=" & LbTalie & " AND CODPRODUTO IS NULL"
PRSet Rs1, sql
qcs = Nnull(Rs1.Fields(0), 0)
Rs1.Close

If qt <> qcs Then
    MsgBox "Divergencia na qtde de itens"
    Exit Sub
End If

sql = "SELECT * FROM " & Esquema_Op & ".TB_TALIE_ITEM WHERE AUTONUM_TALIE=" & LbTalie
PRSet Rs1, sql
Do While Not Rs1.EOF
    
    PCS = DB.Execute("SELECT MIN(AUTONUM_PCS) FROM " & Esquema_Op & ".TB_PATIO_CS WHERE TALIE_DESCARGA=" & LbTalie & " and codproduto is null").Fields(0)
    If Nnull(PCS, 0) <> 0 Then
        sql = "UPDATE " & Esquema_Op & ".TB_PATIO_CS SET CODPRODUTO='" & Rs1!CODPRODUTO & "' WHERE AUTONUM_PCS=" & PCS
        DB.Execute UCase(sql)
    End If
    Rs1.MoveNext

Loop
Rs1.Close
MsgBox "FIM"







End Sub

Private Sub DataGrid1_Click()
    Call Atualiza_Item
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)

WCampo = Empty
If WCampo <> DataGrid1.Columns(ColIndex).DataField Then
    WCampo = DataGrid1.Columns(ColIndex).DataField
    WLabel = DataGrid1.Columns(ColIndex).Caption
    
    sql = "SELECT c.id_conteiner,b.reference,b.instrucao,d.fantasia,b.autonum_parceiro,A.AUTONUM_REG"
    sql = sql & ",a.AUTONUM_TALIE , a.AUTONUM_PATIO, a.Placa, a.Inicio, a.TERMINO, a.FLAG_DESCARGA, a.FLAG_ESTUFAGEM, a.flag_fechado"
    sql = sql & ",a.CROSSDOCKING, a.CONFERENTE ,a.EQUIPE, a.AUTONUM_BOO,a.FLAG_CARREGAMENTO,A.AUTONUM_GATE, a.flag_completo"
    sql = sql & " FROM"
    sql = sql & " " & Esquema_Op & ".tb_talie a inner join " & Esquema_Op & ".tb_booking b on a.autonum_boo = b.autonum_boo"
    sql = sql & " left join " & Esquema_Op & ".tb_patio c on a.autonum_patio = c.autonum_patio"
    sql = sql & " inner join " & Esquema_Op & ".tb_cad_parceiros d on b.autonum_parceiro = d.autonum"
    sql = sql & " order by " & WCampo
    Dc_Talie.ConnectionString = ConnectionString
    Dc_Talie.RecordSource = sql
    Dc_Talie.Refresh

End If

End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Call Atualiza_Item
End Sub

Private Sub DataGrid2_HeadClick(ByVal ColIndex As Integer)

WCampo = Empty
If WCampo <> DataGrid2.Columns(ColIndex).DataField Then
    WCampo = DataGrid2.Columns(ColIndex).DataField
    WLabel = DataGrid2.Columns(ColIndex).Caption
    sql = "SELECT a.codproduto,a.autonum_nfi,a.nf,a.qtde_descarga,a.lote,a.qtde_estufagem,a.autonum_ti,a.autonum_talie,d.qtde,d.qtde_acresc,d.qtde_falta,"
    sql = sql & " to_char(d.qtde) || '   ' || b.descricao_emb || '   ' || c.desc_produto as display, a.falta, a.acrescimo "
    sql = sql & " FROM"
    sql = sql & " " & Esquema_Op & ".tb_talie_item a inner join " & Esquema_Op & ".tb_cad_embalagens b on a.autonum_emb = b.autonum_emb"
    sql = sql & " inner join " & Esquema_Op & ".tb_cad_produtos c on a.autonum_pro = c.autonum_pro"
    sql = sql & " left outer join " & Esquema_Op & ".tb_notas_itens d on a.autonum_nfi = d.autonum_nfi"
    sql = sql & " where a.autonum_talie=" & Nnull(LbTalie, 0)
    sql = sql & " order by " & WCampo
    Adodc6.ConnectionString = ConnectionString
    Adodc6.RecordSource = UCase(sql)
    Adodc6.Refresh
   
End If

End Sub

Private Sub Dc_Talie_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

    If Not Dc_Talie.Recordset.EOF Then
        
        Label6.Visible = False
        CargaSuzano = False
        
        Me.Frameitem.Enabled = True
        
        If Dc_Talie.Recordset!FLAG_FECHADO = 1 Then
            Label5.Visible = True
            CmdFinaliza.Enabled = False
        Else
            Label5.Visible = False
            CmdFinaliza.Enabled = True
        End If
        
        If Dc_Talie.Recordset!FLAG_DESCARGA = 1 Then Option2.Value = 1
        If Dc_Talie.Recordset!FLAG_ESTUFAGEM = 1 Then Option1.Value = 1
        If Dc_Talie.Recordset!flag_carregamento = 1 Then Option3.Value = 1
        If Dc_Talie.Recordset!crossdocking = 1 Then
            Check1.Value = 1
            Label4.Visible = True
        Else
            Check1.Value = 0
            Label4.Visible = False
        End If
        
        Check2.Value = Nnull(Dc_Talie.Recordset!flag_completo, 0)
        
        TxtRegistro = Nnull(Dc_Talie.Recordset!autonum_REG, 0)
        MskCntr.Tag = Nnull(Dc_Talie.Recordset!autonum_patio, 1)
        If Nnull(Dc_Talie.Recordset!ID_CONTEINER, 1) <> "" Then
            MskCntr = Nnull(Dc_Talie.Recordset!ID_CONTEINER, 1)
        Else
            MskCntr = "__________-_"
        End If
        MskData(0) = Format(Dc_Talie.Recordset!Inicio, "dd/mm/yyyy hh:mm"): MskData(0).Refresh
        
        If Not IsNull(Dc_Talie.Recordset!TERMINO) Then
            MskData(1) = Format(Dc_Talie.Recordset!TERMINO, "hh:mm")
        Else
            MskData(1) = "__:__"
        End If
        MskData(1).Refresh

        
        If Nnull(Dc_Talie.Recordset!placa, 1) <> "" Then
            MskPlaca = Nnull(Dc_Talie.Recordset!placa, 1)
        Else
            MskPlaca = "___-____"
        End If
        
        MskPlaca.Tag = Nnull(Dc_Talie.Recordset!autonum_gate, 0)
        
        If Nnull(Dc_Talie.Recordset!forma_operacao, 1) = "" Then Combo1.ListIndex = -1
        If Nnull(Dc_Talie.Recordset!forma_operacao, 1) = "A" Then Combo1.ListIndex = 0
        If Nnull(Dc_Talie.Recordset!forma_operacao, 1) = "M" Then Combo1.ListIndex = 1
        
        sql = "select reference, autonum_boo from " & Esquema_Op & ".tb_booking where autonum_boo = " & Dc_Talie.Recordset!autonum_boo
        Adodc4.ConnectionString = ConnectionString
        Adodc4.RecordSource = UCase(sql)
        Adodc4.Refresh
        DcReserva.BoundText = Nnull(Dc_Talie.Recordset!autonum_boo, 1)
        
        LbTalie = Nnull(Dc_Talie.Recordset!autonum_talie, 1)
        Check1.Value = Nnull(Dc_Talie.Recordset!crossdocking, 0)
        
        DataCombo1.BoundText = Nnull(Dc_Talie.Recordset!CONFERENTE, 0)
        DataCombo2.BoundText = Nnull(Dc_Talie.Recordset!EQUIPE, 0)
        DCombo2(2).BoundText = Nnull(Dc_Talie.Recordset!AUTONUM_PARCEIRO, 0)
        
        
        If DcReserva.BoundText <> "" Then
          sql = Empty
          sql = sql & " SELECT A.AUTONUM_NFI AS ID, to_char(A.QTDE) || '    ' || C.DESCRICAO_EMB || '    ' || D.DESC_PRODUTO AS DISPLAY, A.AUTONUM_EMB,A.AUTONUM_PRO, A.QTDE"
          sql = sql & " FROM " & Esquema_Op & ".TB_NOTAS_ITENS A "
          sql = sql & " INNER JOIN " & Esquema_Op & ".TB_NOTAS_FISCAIS B ON A.AUTONUM_NF = B.AUTONUM_NF"
          sql = sql & " LEFT OUTER JOIN " & Esquema_Op & ".TB_CAD_EMBALAGENS C ON A.AUTONUM_EMB = C.AUTONUM_EMB"
          sql = sql & " LEFT OUTER JOIN " & Esquema_Op & ".TB_CAD_PRODUTOS D ON A.AUTONUM_PRO = D.AUTONUM_PRO"
          sql = sql & " WHERE B.NUM_NF = '" & Text2(0) & "' and b.serie_nf = '" & Text2(13) & "' AND B.AUTONUM_BOO = " & DcReserva.BoundText & ""
          Adodc1.ConnectionString = ConnectionString
          Adodc1.RecordSource = UCase(sql)
          Adodc1.Refresh
                
            
            sql = "select flag_suzano from " & Esquema_Op & ".tb_booking where autonum_boo=" & Nnull(DcReserva.BoundText, 0)
            sql = UCase(sql)
            If Nnull(DB.Execute(sql).Fields(0), 0) = 1 Then
                Label6.Visible = True
                CargaSuzano = True
            End If
            
        End If
    
    
    End If

End Sub

Private Sub DCombo2_KeyPress(Index As Integer, KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub




Private Sub FiltroInstrucao_LostFocus(Index As Integer)
FiltroInstrucao(6) = UCase(FiltroInstrucao(6))
End Sub

Private Sub FiltroReserva_LostFocus(Index As Integer)
FiltroReserva(0) = UCase(FiltroReserva(0))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case vbKeyF2
        Call CmdFinaliza_Click
    Case vbKeyF5
        Call Cmd_Click(0)
    Case vbKeyF6
        Call Cmd_Click(1)
    Case vbKeyF7
        Call Cmd_Click(4)
    Case vbKeyF8
        Call Cmd_Click(2)
    
    Case vbKeyF9
        Call Cmd_Click(6)
    Case vbKeyF10
        Call Cmd_Click(8)
    Case vbKeyF11
        Call Cmd_Click(9)
    Case vbKeyF12
        Call Cmd_Click(7)
    
    Case vbKeyF4
        If Frameitem.Enabled = True Then
            Call Cmd_Click(10)
        Else
            Call Cmd_Click(3)
        End If
    Case vbKeyF3
        If Frame1.Visible = False Then
            Frameitem.Visible = False
            Frame1.Visible = True
            FiltroReserva(0).SetFocus
        Else
            Call cmdFiltro_Click
            Frame1.Visible = False
            Frameitem.Visible = True
        End If
End Select




End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then SendKeys "{TAB}"

End Sub


Private Sub Form_Load()


Grid1.Cols = 6


Nota_item = False

Adodc5.ConnectionString = ConnectionString
Adodc5.RecordSource = "Select autonum autonum_parceiro, fantasia from " & Esquema_Op & ".tb_cad_parceiros where flag_exportador=1 and flag_ativo=1 order by fantasia"
Adodc5.Refresh
Adodc5.Recordset.ActiveConnection = Nothing

'Adodc1.ConnectionString = ConnectionString
'Adodc1.RecordSource = "Select autonum_emb, descricao_emb from " & Esquema_Op & ".tb_cad_embalagens order by descricao_emb "
'Adodc1.Refresh
'Adodc1.Recordset.ActiveConnection = Nothing

Adodc2.ConnectionString = ConnectionString
Adodc2.RecordSource = "Select autonum_eqp, nome_eqp from " & Esquema_Op & ".tb_equipe where flag_ativo=1 and flag_conferente=1 order by nome_eqp "
Adodc2.Refresh
Adodc2.Recordset.ActiveConnection = Nothing

Adodc3.ConnectionString = ConnectionString
Adodc3.RecordSource = "Select autonum_eqp, nome_eqp from " & Esquema_Op & ".tb_equipe where flag_ativo=1 and flag_operador=1  order by nome_eqp "
Adodc3.Refresh
Adodc3.Recordset.ActiveConnection = Nothing


Dc_Talie.ConnectionString = ConnectionString
sql = "SELECT c.id_conteiner,b.reference,b.instrucao,d.fantasia,b.autonum_parceiro,A.AUTONUM_REG"
sql = sql & ",a.AUTONUM_TALIE , a.AUTONUM_PATIO, a.Placa, a.Inicio, a.TERMINO, a.FLAG_DESCARGA, a.FLAG_ESTUFAGEM"
sql = sql & ",a.CROSSDOCKING, a.CONFERENTE ,a.EQUIPE, a.AUTONUM_BOO,a.FLAG_CARREGAMENTO, A.AUTONUM_GATE,a.flag_fechado, A.FLAG_COMPLETO, a.forma_operacao"
sql = sql & " FROM"
sql = sql & " " & Esquema_Op & ".tb_talie a inner join " & Esquema_Op & ".tb_booking b on a.autonum_boo = b.autonum_boo"
sql = sql & " left join " & Esquema_Op & ".tb_patio c on a.autonum_patio = c.autonum_patio"
sql = sql & " inner join " & Esquema_Op & ".tb_cad_parceiros d on b.autonum_parceiro = d.autonum"
sql = sql & " where "
If CD Then
    sql = sql & " flag_descarga=1 and crossdocking=0"
Else
    sql = sql & " flag_descarga=0 and crossdocking=1"
End If
sql = sql & " order by a.inicio desc, d.fantasia, c.id_conteiner, b.reference, b.instrucao"
Dc_Talie.RecordSource = UCase(sql)
Dc_Talie.Refresh
'Dc_Talie.Recordset.ActiveConnection = Nothing


Me.Left = (MDI.Width - Me.Width) / 2
Me.Top = (MDI.Height - Me.Height) / 2

End Sub












Private Sub Grid1_Click()
    Select Case Grid1.TextMatrix(Grid1.Row, 9)
    Case "ACRESCIMO"
        Option4(0).Value = True
    Case "FALTA"
        Option4(1).Value = True
    Case "AVARIA"
        Option4(2).Value = True
    Case "TOTAL"
        Option4(3).Value = True
    Case "SALDO"
        Option4(4).Value = True
    End Select



End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)

If Grid1.Col = 8 Then

    Option4(0).Enabled = False
    Option4(1).Enabled = False
    Option4(2).Enabled = False
    Option4(3).Enabled = False
    Option4(4).Enabled = False
    
    If KeyAscii = vbKeyBack And Grid1.TextMatrix(Grid1.Row, 8) <> "" Then
        Grid1.TextMatrix(Grid1.Row, 8) = Left(Grid1.TextMatrix(Grid1.Row, 8), Len(Grid1.TextMatrix(Grid1.Row, 8)) - 1)
        
        If Val(Grid1.TextMatrix(Grid1.Row, 8)) < Val(Grid1.TextMatrix(Grid1.Row, 4)) Then
            Grid1.TextMatrix(Grid1.Row, 9) = "FALTA"
            Grid1.TextMatrix(Grid1.Row, 10) = Val(Grid1.TextMatrix(Grid1.Row, 4)) - Val(Grid1.TextMatrix(Grid1.Row, 8))
            Option4(1).Value = True
            Option4(1).Enabled = True
            Option4(2).Enabled = True
            Option4(4).Enabled = True
        End If
        If Val(Grid1.TextMatrix(Grid1.Row, 8)) > Val(Grid1.TextMatrix(Grid1.Row, 4)) Then
            Grid1.TextMatrix(Grid1.Row, 9) = "ACRESCIMO"
            Grid1.TextMatrix(Grid1.Row, 10) = Val(Grid1.TextMatrix(Grid1.Row, 8)) - Val(Grid1.TextMatrix(Grid1.Row, 4))
            Option4(0).Value = True
        End If
        If Val(Grid1.TextMatrix(Grid1.Row, 8)) = Val(Grid1.TextMatrix(Grid1.Row, 4)) Then
            Grid1.TextMatrix(Grid1.Row, 9) = "TOTAL"
            Grid1.TextMatrix(Grid1.Row, 10) = 0
            Option4(3).Value = True
        End If
        
                
        Exit Sub
    End If
    
    If Chr(KeyAscii) >= "0" And Chr(KeyAscii) <= "9" Then
        Grid1.TextMatrix(Grid1.Row, 8) = Grid1.TextMatrix(Grid1.Row, 8) & Chr(KeyAscii)
        
        If Val(Grid1.TextMatrix(Grid1.Row, 8)) < Val(Grid1.TextMatrix(Grid1.Row, 4)) Then
            Grid1.TextMatrix(Grid1.Row, 9) = "FALTA"
            Grid1.TextMatrix(Grid1.Row, 10) = Val(Grid1.TextMatrix(Grid1.Row, 4)) - Val(Grid1.TextMatrix(Grid1.Row, 8))
            Option4(1).Value = True
            Option4(1).Enabled = True
            Option4(2).Enabled = True
            Option4(4).Enabled = True
        End If
        If Val(Grid1.TextMatrix(Grid1.Row, 8)) > Val(Grid1.TextMatrix(Grid1.Row, 4)) Then
            Grid1.TextMatrix(Grid1.Row, 9) = "ACRESCIMO"
            Grid1.TextMatrix(Grid1.Row, 10) = Val(Grid1.TextMatrix(Grid1.Row, 8)) - Val(Grid1.TextMatrix(Grid1.Row, 4))
            Option4(0).Value = True
        End If
        If Val(Grid1.TextMatrix(Grid1.Row, 8)) = Val(Grid1.TextMatrix(Grid1.Row, 4)) Then
            Grid1.TextMatrix(Grid1.Row, 9) = "TOTAL"
            Grid1.TextMatrix(Grid1.Row, 10) = 0
            Option4(3).Value = True
        End If
        
        Exit Sub
    End If

End If



If (Grid1.Col >= 12 And Grid1.Col <= 16) Then

    If KeyAscii = vbKeyBack And Grid1.TextMatrix(Grid1.Row, Grid1.Col) <> "" Then
        Grid1.TextMatrix(Grid1.Row, Grid1.Col) = Left(Grid1.TextMatrix(Grid1.Row, Grid1.Col), Len(Grid1.TextMatrix(Grid1.Row, Grid1.Col)) - 1)
        Exit Sub
    End If
    If Chr(KeyAscii) = "," Or (Chr(KeyAscii) >= "0" And Chr(KeyAscii) <= "9") Then
        Grid1.TextMatrix(Grid1.Row, Grid1.Col) = Grid1.TextMatrix(Grid1.Row, Grid1.Col) & Chr(KeyAscii)
        Exit Sub
    End If

End If

If (Grid1.Col = 1) Then
    TxtBarra = ""

    If KeyAscii = vbKeyBack And Grid1.TextMatrix(Grid1.Row, Grid1.Col) <> "" Then
        Grid1.TextMatrix(Grid1.Row, Grid1.Col) = Left(Grid1.TextMatrix(Grid1.Row, Grid1.Col), Len(Grid1.TextMatrix(Grid1.Row, Grid1.Col)) - 1)
        Exit Sub
    End If
    If (Chr(KeyAscii) >= "0" And Chr(KeyAscii) <= "9") Or (Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z") Then
        Grid1.TextMatrix(Grid1.Row, Grid1.Col) = Grid1.TextMatrix(Grid1.Row, Grid1.Col) & Chr(KeyAscii)
        Exit Sub
    End If

End If



End Sub


Private Sub MskCntr_LostFocus()

    
    If Option1.Value = True Then
        If MskCntr.ClipText = "" Then Exit Sub

        MskCntr.Tag = ""
'        Text2(4) = ""
'        Text2(4).Tag = ""
'        Text2(5) = ""
'        DCombo2(2).BoundText = ""

'        sql = "select autonum_patio, autonum_bcg from " & Esquema_Op & ".tb_patio where id_conteiner='" & MskCntr & "' And flag_historico = 0"
        sql = "select autonum_patio, autonum_bcg from " & Esquema_Op & ".tb_patio where id_conteiner='" & MskCntr & "' AND nvl(flag_historico,0)=0"
        PRSet Rs, sql
        If Rs.EOF Then
            MsgBox "Conteiner não encontrado no pátio"
            Rs.Close
            Exit Sub
        End If
        MskCntr.Tag = Rs!autonum_patio
    
        sql = "select a.reference,a.instrucao,a.autonum_parceiro,a.autonum_boo"
        sql = sql & " from " & Esquema_Op & ".tb_booking a, " & Esquema_Op & ".tb_booking_carga b"
        sql = sql & " where "
        sql = sql & " b.autonum_boo=a.autonum_boo"
        sql = sql & " and b.autonum_bcg=" & Nnull(Rs!AUTONUM_BCG, 0)
        Rs.Close
        PRSet Rs, sql
        If Not Rs.EOF Then
            If Rs!AUTONUM_PARCEIRO <> DCombo2(2).BoundText Then
                Rs.Close
                MsgBox "Conteiner selecionado pertence a outro cliente"
                MskCntr = "__________-_"
                Exit Sub
            End If
            If Rs!autonum_boo <> DcReserva.BoundText Then
                
                If MsgBox("Conteiner selecionado pertence a outra reserva" & vbCr & "Deseja realizar a troca da reserva?", vbYesNo) = vbNo Then
                    Rs.Close
                    MskCntr = "__________-_"
                    Exit Sub
                Else
                    sql = "select reference, autonum_boo from " & Esquema_Op & ".tb_booking where autonum_boo = " & Rs!autonum_boo
                    Adodc4.ConnectionString = ConnectionString
                    Adodc4.RecordSource = UCase(sql)
                    Adodc4.Refresh
                    
                    DcReserva.BoundText = Nnull(Rs!autonum_boo, 0)
                End If
            End If
            
'            DCombo2(2).BoundText = Rs!autonum_parceiro
'            Text2(4) = Nnull(Rs!reference, 1)
'            Text2(4).Tag = Nnull(Rs!autonum_boo, 1)
'            Text2(5) = Nnull(Rs!Instrucao, 1)
        End If
        Rs.Close
    End If
    
End Sub


Private Sub MskData_Change(Index As Integer)
    
'    If Index = 0 And Option2.Value = True Then
'        MskPlaca = "___-____"
'        DCombo2(2).BoundText = ""
'        DcReserva.BoundText = ""
'    End If
    
End Sub

Private Sub MskPlaca_LostFocus()

    
'    If Option2.Value = True Then
        DCombo2(2).BoundText = ""
        Text2(5) = ""
        DcReserva.BoundText = ""
        
        
        sql = "select e.autonum_parceiro,e.reference, a.autonum, e.autonum_boo, e.flag_suzano "
        sql = sql & " from " & Esquema_Op & ".tb_gate_new a"
        sql = sql & " inner join " & Esquema_Op & ".tb_registro b on a.autonum = b.autonum_gate"
        sql = sql & " inner join " & Esquema_Op & ".tb_booking e on b.autonum_boo = e.autonum_boo"
        sql = sql & " where a.autonum = "
        sql = sql & " nvl((select max(autonum) from " & Esquema_Op & ".tb_gate_new where placa='" & MskPlaca & "' and flag_gate_in=1),0)"
        PRSet Rs, sql
        If Rs.EOF Then
            MsgBox "Não consta entrada para esta placa/data"
            Rs.Close
            Exit Sub
        End If
        MskPlaca.Tag = Rs!Autonum
        DCombo2(2).BoundText = Rs!AUTONUM_PARCEIRO
        
'        Text2(4) = Nnull(Rs!reference, 1)
'        Text2(4).Tag = Nnull(Rs!autonum_boo, 1)
        
        Call Monta_Combo_Reservas(MskPlaca, Left(MskData(0), 10))
        DcReserva.BoundText = Nnull(Rs!autonum_boo, 0)
        If Nnull(Rs!flag_suzano, 0) = 1 Then
            Label6.Visible = True
            CargaSuzano = True
        Else
            Label6.Visible = False
            CargaSuzano = False
        End If
        
        
        Rs.Close
'    End If

End Sub


Private Sub Option1_Click()
    
    If LbTalie = "" Then
        Frame3.Enabled = True
    '    Text2(4).Enabled = False
        
        DCombo2(2).Enabled = False
        DCombo2(2).BoundText = ""
        DcReserva.BoundText = ""
        
        Label4.Visible = True
        Check1.Value = 1
        MskCntr.Enabled = True
    End If


End Sub


Private Sub Option2_Click()
    
    
    If LbTalie = "" Then
        Frame3.Enabled = True
        MskData(0) = "__/__/____ __:__"
        MskData(1) = "__:__"
        MskPlaca = "___-____"
        MskCntr = "__________-_"
        DCombo2(2).BoundText = ""
        DcReserva.BoundText = ""
        Check1.Value = 0
    End If
    
    Text2(4).Enabled = False
    DCombo2(2).Enabled = False
    Label4.Visible = False
    MskCntr.Enabled = False





End Sub


Private Sub Option3_Click()
    
    Frame3.Enabled = True
    Text2(4).Enabled = True
    Text2(5).Enabled = True
    DCombo2(2).Enabled = True
    DCombo2(2).BoundText = ""
    DcReserva.BoundText = ""
    Text2(5) = ""
    Label4.Visible = False
    Check1.Value = 0
    MskCntr.Enabled = False




End Sub


Private Sub Option4_Click(Index As Integer)
    Select Case Index
    Case 0
        Grid1.TextMatrix(Grid1.Row, 9) = "ACRESCIMO"
    Case 1
        Grid1.TextMatrix(Grid1.Row, 9) = "FALTA"
    Case 2
        Grid1.TextMatrix(Grid1.Row, 9) = "AVARIA"
    Case 3
        Grid1.TextMatrix(Grid1.Row, 9) = "TOTAL"
    Case 4
        Grid1.TextMatrix(Grid1.Row, 9) = "SALDO"
    End Select
End Sub




Private Sub SSTab1_Click(PreviousTab As Integer)

    If SSTab1.Tab = 1 Then
        Text3 = ""
        If Nnull(LbTalie, 0) <> 0 Then
            sql = "select tl.placa,tl.inicio, pcc.id_conteiner, boo.reference from " & Esquema_Op & ".tb_talie tl"
            sql = sql & " left join " & Esquema_Op & ".tb_patio pcc on tl.autonum_patio = pcc.autonum_patio"
            sql = sql & " inner join " & Esquema_Op & ".tb_booking boo on tl.autonum_boo = boo.autonum_boo"
            sql = sql & " where autonum_talie=" & LbTalie
            PRSet Rs, sql
            If Rs.EOF Then
                MsgBox "Talie não encontrado"
                Rs.Close
                Exit Sub
            End If
            Text3 = Rs!placa & "   " & Format(Rs!Inicio, "dd/mm/yyyy hh:mm")
            Text3 = Text3 & vbCr & vbLf & "Reserva : " & Rs!reFERENCE
            If Nnull(Rs!ID_CONTEINER, 1) <> "" Then
                Text3 = Text3 & vbCr & vbLf & Rs!ID_CONTEINER
            End If
            Rs.Close
        End If
    End If

End Sub

Private Sub Text2_Change(Index As Integer)

If Index = 1 Then

    Grid1.TextMatrix(Grid1.Row, 11) = Text2(1)

End If


If Index = 2 Then

    Text2(7) = Abs(Val(Text2(2)) - Val(Text2(5)))
    
    If Val(Text2(2)) < Val(Text2(5)) Then
        Text2(6) = "FALTA"
    End If
    If Val(Text2(2)) > Val(Text2(5)) Then
        Text2(6) = "ACRESCIMO"
    End If
    If Val(Text2(2)) = Val(Text2(5)) Then
        Text2(6) = "TOTAL"
    End If

End If


End Sub

Private Sub Seta_Controles()
For I = 0 To 4
    cmd(I).Enabled = Not cmd(I).Enabled
Next I
End Sub

Private Sub Text2_LostFocus(Index As Integer)
    Dim Tb As New ADODB.Recordset

    Select Case Index
    Case 17
        cmd(2).SetFocus
    Case 3
        If cmd(7).Enabled = True Then
            cmd(7).SetFocus
        End If
    Case 0
'        If Text2(0) <> "" Then
'            PRSet Tb, "SELECT AUTONUM_NF FROM " & Esquema_Op & ".TB_NOTAS_FISCAIS A WHERE  A.NUM_NF = '" & Text2(0) & "' AND A.AUTONUM_BOO = " & Text2(4).Tag & " "
'            If Tb.EOF Then
'                MsgBox "Nota fiscal inválida.", vbCritical
'                Text2(0).SetFocus
'                Exit Sub
'            End If
'        End If
'        sql = Empty
'        sql = sql & " SELECT A.AUTONUM_NFI AS ID, A.ITEM , A.QTDE, C.DESCRICAO_EMB ,D.DESC_PRODUTO, 0 AS DESCARGA"
'        sql = sql & " FROM " & Esquema_Op & ".TB_NOTAS_ITENS A "
'        sql = sql & " INNER JOIN " & Esquema_Op & ".TB_NOTAS_FISCAIS B ON A.AUTONUM_NF = B.AUTONUM_NF"
'        sql = sql & " LEFT OUTER JOIN " & Esquema_Op & ".TB_CAD_EMBALAGENS C ON A.AUTONUM_EMB = C.AUTONUM_EMB"
'        sql = sql & " LEFT OUTER JOIN " & Esquema_Op & ".TB_CAD_PRODUTOS D ON A.AUTONUM_PRO = D.AUTONUM_PRO"
'        sql = sql & " WHERE B.NUM_NF = '" & Text2(0) & "' AND B.AUTONUM_BOO = " & Text2(4).Tag & ""
'        Adodc1.ConnectionString = ConnectionString
'        Adodc1.RecordSource = sql
'        Adodc1.Refresh
'
'
'        sql = sql & "ID|ITEM|QTDE|EMBALAGEM|PRODUTO|QTDE_DESCARGA"
'        Grid1.Cols = 8
'        Grid1.FormatString = sql
    
    Case 1
        FALTA = 0
        ACRESC = 0
'        Dim QI As Long
'        QI = Nnull(DB.Execute("select qtde from " & Esquema_Op & ".tb_notas_itens where autonum_nfi=" & DcItem.BoundText).Fields(0), 0)
'        qd = Nnull(DB.Execute("select sum(qtde_descarga) from " & Esquema_Op & ".tb_talie_item where autonum_nfi=" & DcItem.BoundText).Fields(0), 0)
'        If Val(Text2(1)) > (QI - qd) Then
'            ACRESC = Val(Text2(1)) - (QI - qd)
'            Text1.Text = "ACRÉSCIMO DE " & ACRESC & ""
'        End If
'        If Val(Text2(1)) < (QI - qd) Then
'            FALTA = (QI - qd) - Val(Text2(1))
'            Text1.Text = "FALTA " & FALTA & ""
'        End If
    End Select
    
End Sub


Private Sub TxtBarra_LostFocus()

    If TxtBarra <> "" Then
        If CargaSuzano Then
            sql = "select count(*) from " & Esquema_Op & ".tb_integra_reg ir"
            sql = sql & " inner join " & Esquema_Op & ".tb_integra_carga ic on ir.id_reg = ic.id_reg"
            sql = sql & " where"
            sql = sql & " ir.autonum_boo = " & Val(Text2(0).Tag)
            sql = sql & " and ic.codbarra='" & TxtBarra & "'"
            sql = UCase(sql)
            If Nnull(DB.Execute(sql).Fields(0), 0) = 0 Then
                MsgBox "Produto não encontrado como carga esperada" & vbCr & "Verifique com o operacional"
                TxtBarra = ""
                Exit Sub
            End If
        End If
    End If
    


End Sub


Private Sub TxtRegistro_LostFocus()
    
    If TxtRegistro <> "" Then
    
        Label6.Visible = False
    
        Dim Rsx As New ADODB.Recordset
        sql = "select reg.* from"
        sql = sql & " " & Esquema_Op & ".tb_registro reg"
        sql = sql & " where"
        sql = sql & " autonum_reg=" & TxtRegistro
        PRSet Rsx, sql
        If Rsx.EOF Then
            MsgBox "Registro não encontrado"
            Rsx.Close
            TxtRegistro = ""
            Exit Sub
        End If
        
        MskData(0) = Format(Now, "dd/mm/yyyy hh:mm")
        MskPlaca = Rsx!placa
        Call MskPlaca_LostFocus
        
    End If
    
End Sub


