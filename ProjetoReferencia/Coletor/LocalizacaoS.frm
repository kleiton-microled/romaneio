VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form Localizacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Localização"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3570
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
   ScaleHeight     =   5010
   ScaleWidth      =   3570
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5115
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   3555
      Begin VB.CommandButton CmdSaida 
         Caption         =   "SAIDA"
         Height          =   405
         Left            =   2640
         TabIndex        =   32
         Top             =   1950
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtResponsavel 
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
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1590
         Width           =   2685
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   4050
         Top             =   7470
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
      Begin VB.TextBox txtTransportadora 
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
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1350
         Width           =   2685
      End
      Begin VB.TextBox txtRua 
         Alignment       =   2  'Center
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
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1110
         Width           =   615
      End
      Begin VB.TextBox txtQuadra 
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
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1110
         Width           =   1695
      End
      Begin VB.TextBox txtCarga 
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
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   870
         Width           =   2685
      End
      Begin VB.TextBox TxtPatio 
         Alignment       =   2  'Center
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "@"
         Top             =   180
         Width           =   495
      End
      Begin VB.TextBox txtGateIN 
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   660
         Width           =   975
      End
      Begin VB.TextBox txtLibera 
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
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   630
         Width           =   855
      End
      Begin VB.TextBox txtCarreta 
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
         Height          =   255
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox txtCavalo 
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
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   420
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "FILTRO"
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   60
         TabIndex        =   3
         Top             =   1830
         Width           =   2475
         Begin VB.CommandButton Command1 
            Caption         =   "Filtrar"
            Height          =   375
            Left            =   1770
            TabIndex        =   23
            Top             =   150
            Width           =   675
         End
         Begin MSMask.MaskEdBox mskPlaca 
            Height          =   315
            Left            =   630
            TabIndex        =   22
            Top             =   180
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">???-#A##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PLACA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   30
            TabIndex        =   21
            Top             =   210
            Width           =   540
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
         Bindings        =   "LocalizacaoS.frx":0000
         Height          =   2385
         Left            =   60
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2400
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   4207
         _Version        =   393216
         Cols            =   11
         RowHeightMin    =   390
         WordWrap        =   -1  'True
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   $"LocalizacaoS.frx":0015
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
         _Band(0).Cols   =   11
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RESPONS."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   165
         Index           =   2
         Left            =   30
         TabIndex        =   30
         Top             =   1590
         Width           =   750
      End
      Begin VB.Label LBLTOTCS 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   1530
         TabIndex        =   29
         Top             =   4800
         Width           =   525
      End
      Begin VB.Label LBLTOTCNTR 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   630
         TabIndex        =   28
         Top             =   4800
         Width           =   525
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOT"
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
         Left            =   2160
         TabIndex        =   27
         Top             =   4830
         Width           =   330
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   1290
         TabIndex        =   26
         Top             =   4860
         Width           =   255
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CNTR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   120
         TabIndex        =   25
         Top             =   4830
         Width           =   420
      End
      Begin VB.Label lblTotEnt 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   2550
         TabIndex        =   24
         Top             =   4800
         Width           =   585
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRANSP."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   165
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   1380
         Width           =   675
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RUA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   2520
         TabIndex        =   17
         Top             =   1140
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QUADRA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   120
         TabIndex        =   15
         Top             =   1140
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CARGA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   240
         TabIndex        =   12
         Top             =   900
         Width           =   555
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   1800
         TabIndex        =   10
         Top             =   660
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LIB. ENT."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   90
         TabIndex        =   8
         Top             =   660
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CARRETA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   1740
         TabIndex        =   6
         Top             =   450
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CAVALO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   120
         TabIndex        =   4
         Top             =   450
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CARREGAMENTOS -  PATIO "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   360
         TabIndex        =   1
         Top             =   180
         Width           =   2160
      End
   End
End
Attribute VB_Name = "Localizacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdSaida_Click()
'Se saiu de posicao valida nao pode ter conteiner acima

If Me.txtQuadra.text = "SAIDA" Then
    Col_Mensagem = "Conteiner já posicionado em SAIDA"
    Frmmensagem.Label1.Caption = Col_Mensagem
    Frmmensagem.Show vbModal
    Exit Sub
End If
    



SQl = "SELECT V.VALIDA,A.YARD,A.AUTONUM,A.ID_CONTEINER FROM " & Banco_Operador & "VW_INVENT_SISTEMAS A "
SQl = SQl & " INNER JOIN " & Banco_Operador & "TB_YARD V ON A.PATIO=V.PATIO AND A.YARD=V.YARD "
SQl = SQl & " Where a.PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio) & " AND A.ID_CONTEINER='" & Me.txtCarga.text & "'"
Set tbant = New ADODB.Recordset
tbant.Open SQl, db, 1, 1
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
    
        SQl = "SELECT AUTONUM, ID_CONTEINER FROM " & Banco_Operador & "VW_INVENT_SISTEMAS WHERE PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio) & " AND YARD='" & ColPilhaO & Trim(str(Val(ColAlturaO + 1))) & "'"
        Set tb1 = New ADODB.Recordset
        tb1.Open SQl, db, 1, 1
        If Not tb1.EOF Then
             Frmmensagem.Label1.Caption = "Existe um conteiner acima, impossivel movimentar ( " & tb1!Id_Conteiner & " )"
             Frmmensagem.Show vbModal
            Exit Sub
        End If
        tb1.Close
    End If
End If


Col_Mensagem = "Confirma o carregamento da unidade"
Col_Imagem = 3
frmMensagem2.Show vbModal
If Col_Cancelado = True Then
    Exit Sub
End If



SQL2 = "INSERT INTO " & Banco_Operador & "TB_HIST_SHIFTING ("
SQL2 = SQL2 & "CNTR,ORIGEM,DESTINO,DATA,TIPO,USUARIO,MOTIVO"
SQL2 = SQL2 & ") VALUES "
SQL2 = SQL2 & "("
SQL2 = SQL2 & Grid1.TextMatrix(Me.Grid1.Row, Grid1.Cols - 1) & ","
SQL2 = SQL2 & "'" & TXTYardAtual & "',"
SQL2 = SQL2 & "'SAIDA',"
SQL2 = SQL2 & "GETDATE(),"
SQL2 = SQL2 & "'I',"
SQL2 = SQL2 & Cod_Usuario
SQL2 = SQL2 & ",0"
SQL2 = SQL2 & ")"


SQl = "UPDATE " & Banco_Sgipa & "tb_cntr_bl SET "
SQl = SQl & " YARD='SAIDA'"
SQl = SQl & " WHERE Autonum=" & Grid1.TextMatrix(Me.Grid1.Row, Grid1.Cols - 1)
    
    
Executa SQl
Executa SQL2

Carrega_Grid1



End Sub

Private Sub Command1_Click()
Carrega_Grid1
End Sub

Private Sub Form_Activate()
'Carrega_Grid1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Me.TxtPatio.text = Nnull(db.Execute("Select descr_resumido from operador..tb_patios where autonum " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio))(0), 1)
Carrega_Grid1
End Sub

Private Sub Carrega_Grid1()

Screen.MousePointer = 11
Dim SQl As String
Me.Grid1.Visible = False

SQl = " SELECT   gn.placa,  gn.carreta ,   "
SQl = SQl + " convert(varchar(5),o.dt_lib_ent_cam,103) + ' ' +  "
SQl = SQl + " convert(varchar(5),o.dt_lib_ent_cam,108) dt_lib_ent_cam,  "
SQl = SQl + "  convert(varchar(5),gn.dt_gate_in,103) + ' ' +  "
SQl = SQl + "  convert(varchar(5),gn.dt_gate_in,108) dt_gate_in,   "
SQl = SQl + "  cc.id_conteiner,   cc.yard,   isnull(Rem.temp_rua,' ') as rua,  "
SQl = SQl + "    T.fantasia as Nometransportadora, u.usuario, 1 AS FLAG_CNTR, "
SQl = SQl + "    cc.autonum as autonum_cntr    FROM operador..tb_gate_new gn   "
SQl = SQl + "     inner join operador..tb_amr_gate amr on gn.autonum = amr.gate  "
SQl = SQl + "       inner join sgipa..TB_CNTR_BL cc on amr.CNTR_ipa=cc.AUTONUM          "
SQl = SQl + "       inner join operador..TB_CAD_TRANSPORTADORAS T  "
SQl = SQl + "       on gn.id_TRANSPORTADORA=T.AUTONUM    "
SQl = SQl + "       inner join sgipa..TB_ORDEM_CARREGAMENTO O on  "
SQl = SQl + "       Amr.ID_OC=O.AUTONUM    "
SQl = SQl + "       left join operador..tb_remocao rem  "
SQl = SQl + "       on cc.autonum=rem.AUTONUM and rem.sistema='I'   "
SQl = SQl + "        left join sgipa..tb_cad_usuarios u  "
SQl = SQl + "        on o.usu_lib_ent_cam=u.autonum     "
SQl = SQl + "        Where   gn.dt_gate_out is null and gn.dt_gate_in>getdate()-3 "
SQl = SQl & "   and cc.patio " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio)
If Me.mskPlaca.ClipText <> "" Then SQl = SQl & " AND (GN.PLACA='" & Me.mskPlaca & "' OR GN.CARRETA='" & Me.mskPlaca & "')"
SQl = SQl + "           Union All   SELECT    max(gn.placa),  "
SQl = SQl + "             max(gn.carreta) ,   max(convert(varchar "
SQl = SQl + "(5),loc.dt_lib_ent_cam,103) + ' ' + convert(varchar(5), "
SQl = SQl + "loc.dt_lib_ent_cam,108)) dt_lib_ent_cam,   "
SQl = SQl + " max(convert(varchar(5),gn.dt_gate_in,103) + ' ' +  "
SQl = SQl + " convert(varchar(5),gn.dt_gate_in,108)) dt_gate_in, "
SQl = SQl + "    convert(varchar,sum(LOC.quantidade)) + ' vols ' as carga  , "
SQl = SQl + "       max(ISNULL(LOC.ARMAZEM,' ') + '-' + ISNULL(LOC.YARD,' '))  "
SQl = SQl + "       AS YARD,   ' ' as rua,   max(T.fantasia)  "
SQl = SQl + "       as Nometransportadora, max(u.usuario) as usuario ,  "
SQl = SQl + "       0 AS FLAG_CNTR,0 as autonum_cntr    "
SQl = SQl + "       FROM operador..tb_gate_new gn    "
SQl = SQl + "       inner join operador..tb_amr_gate amr  "
SQl = SQl + "       on gn.autonum = amr.gate   "
SQl = SQl + "       inner join operador..TB_CAD_TRANSPORTADORAS T  "
SQl = SQl + "       on gn.id_TRANSPORTADORA=T.AUTONUM              "
SQl = SQl + "         LEFT JOIN sgipa..VW_LOCAL_OC LOC "
SQl = SQl + "          ON Amr.ID_OC=LOC.AUTONUMORDEM     "
SQl = SQl + "          left join sgipa..tb_cad_usuarios u  "
SQl = SQl + "          on LOC.usu_lib_ent_cam=u.autonum   "
SQl = SQl + "  Where   gn.dt_gate_out is null and gn.dt_gate_in>getdate()-3  "
SQl = SQl & "   and loc.patio " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio)
If Me.mskPlaca.ClipText <> "" Then SQl = SQl & " AND (GN.PLACA='" & Me.mskPlaca & "' OR GN.CARRETA='" & Me.mskPlaca & "')"

Adodc1.ConnectionString = ConnectionStringAdo
Adodc1.RecordSource = SQl
Adodc1.Refresh

Me.lblTotEnt.Caption = Me.Adodc1.Recordset.RecordCount

Set Adodc1.Recordset.ActiveConnection = Nothing

Me.Grid1.FormatString = "   |^Cavalo     |^Carreta    |^Lib. Ent.|^Gate IN|^Carga               |^Quadra    |^Rua |<Transportadora                        |<Responsavel                      ||"

Me.Grid1.Visible = True

Dim TOTCNTR As Integer
Dim TOTCS As Integer

TOTCNTR = 0
TOTCS = 0

For i = 1 To Grid1.Rows - 1
    If Grid1.TextMatrix(i, Grid1.Cols - 2) = 1 Then
        TOTCNTR = TOTCNTR + 1
    Else
        TOTCS = TOTCS + 1
    End If
Next i

Me.LBLTOTCNTR = TOTCNTR
Me.LBLTOTCS = TOTCS
Screen.MousePointer = 0
Me.Grid1.COLWIDTH(2) = 0
Me.Grid1.COLWIDTH(Me.Grid1.Cols - 1) = 0
Me.Grid1.COLWIDTH(Me.Grid1.Cols - 2) = 0

End Sub

Private Sub Grid1_Click()
popula
End Sub


Private Sub popula()
If Me.Grid1.Row > 0 Then
    Me.txtCavalo.text = Me.Grid1.TextMatrix(Me.Grid1.Row, 1)
    Me.txtCarreta.text = Me.Grid1.TextMatrix(Me.Grid1.Row, 2)
    Me.txtLibera.text = Me.Grid1.TextMatrix(Me.Grid1.Row, 3)
    Me.txtGateIN.text = Me.Grid1.TextMatrix(Me.Grid1.Row, 4)
    Me.txtCarga.text = Me.Grid1.TextMatrix(Me.Grid1.Row, 5)
    Me.txtQuadra.text = Me.Grid1.TextMatrix(Me.Grid1.Row, 6)
    Me.txtRua.text = Me.Grid1.TextMatrix(Me.Grid1.Row, 7)
    Me.txtTransportadora.text = Me.Grid1.TextMatrix(Me.Grid1.Row, 8)
    Me.txtResponsavel = Me.Grid1.TextMatrix(Me.Grid1.Row, 9)
    If Grid1.TextMatrix(Me.Grid1.Row, Grid1.Cols - 2) = 1 Then
        Me.CmdSaida.Visible = True
    Else
        Me.CmdSaida.Visible = False
    End If
End If
End Sub

Private Sub Grid1_RowColChange()
popula
End Sub

