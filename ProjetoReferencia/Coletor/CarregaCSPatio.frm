VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form CarregaCSPatio 
   Appearance      =   0  'Flat
   BackColor       =   &H00BA9F8D&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3690
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdF2 
      Caption         =   "[F2] DOC"
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
      Left            =   1920
      TabIndex        =   32
      Top             =   4140
      Width           =   800
   End
   Begin VB.CommandButton cmdF5 
      Caption         =   "[F5] CANCELA"
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
      Left            =   2700
      TabIndex        =   31
      Top             =   4140
      Width           =   950
   End
   Begin VB.CommandButton cmdF4 
      Caption         =   "[F4] CARREGA"
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
      Left            =   980
      TabIndex        =   30
      Top             =   4140
      Width           =   920
   End
   Begin VB.CommandButton cmdF1 
      Caption         =   "[F1] LOCAIS"
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
      Left            =   30
      TabIndex        =   29
      Top             =   4140
      Width           =   945
   End
   Begin VB.Frame frmLocais 
      Appearance      =   0  'Flat
      BackColor       =   &H00BA9F8D&
      ForeColor       =   &H00000000&
      Height          =   3915
      Left            =   30
      TabIndex        =   23
      Top             =   180
      Visible         =   0   'False
      Width           =   3645
      Begin VB.TextBox txt2Lote 
         Alignment       =   2  'Center
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
         Height          =   330
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   540
         Width           =   1215
      End
      Begin VB.TextBox txt2Ordem 
         Alignment       =   2  'Center
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
         Height          =   330
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   180
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvLocal 
         Height          =   2835
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   5001
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   1620
         TabIndex        =   25
         Top             =   540
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem de Carregamento :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   210
         Width           =   1890
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   345
      Left            =   3210
      TabIndex        =   21
      ToolTipText     =   "Refresh"
      Top             =   210
      Width           =   345
   End
   Begin VB.TextBox TXTLOCAL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   345
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3540
      Width           =   885
   End
   Begin VB.TextBox txtArmazem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   345
      Left            =   1830
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3540
      Width           =   915
   End
   Begin VB.TextBox txtCritica 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3900
      Width           =   3585
   End
   Begin VB.TextBox txtAutonum_Cs_Yard 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3060
      MaxLength       =   9
      TabIndex        =   16
      Top             =   2670
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAutonumCS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3060
      MaxLength       =   9
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2850
      Top             =   -180
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox TXTQTDE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   510
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3540
      Width           =   765
   End
   Begin VB.TextBox TXTLOTE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   345
      Left            =   2700
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3210
      Width           =   885
   End
   Begin VB.TextBox txtMarcante 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   840
      MaxLength       =   9
      TabIndex        =   8
      Top             =   3210
      Width           =   1305
   End
   Begin MSComctlLib.ListView lvOC 
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "CarregaCSPatio.frx":0000
      Height          =   345
      Left            =   420
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   180
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   609
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "Display"
      Text            =   "DataCombo1"
   End
   Begin MSComctlLib.ListView lvCm 
      Height          =   1065
      Left            =   0
      TabIndex        =   6
      Top             =   2130
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   1879
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[F1]-LOCAIS"
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
      Height          =   225
      Index           =   3
      Left            =   30
      TabIndex        =   22
      Top             =   4110
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LOCAL"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   6
      Left            =   1260
      TabIndex        =   19
      Top             =   3540
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[F4]-CARREGA ITEM"
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
      Height          =   225
      Index           =   4
      Left            =   1020
      TabIndex        =   14
      Top             =   4110
      Width           =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[F5]-CANCELA"
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
      Height          =   225
      Index           =   2
      Left            =   2610
      TabIndex        =   13
      Top             =   4110
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "QTDE"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   3
      Left            =   0
      TabIndex        =   10
      Top             =   3540
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LOTE"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   2
      Left            =   2130
      TabIndex        =   9
      Top             =   3210
      Width           =   585
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MARCANTE"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   3210
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CARREGAMENTOS"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   1920
      Width           =   3585
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ORDENS"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   510
      Width           =   3585
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "VEIC."
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   0
      Left            =   30
      TabIndex        =   1
      Top             =   210
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CARREGAMENTO CARGA SOLTA"
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
      TabIndex        =   0
      Top             =   0
      Width           =   3600
   End
End
Attribute VB_Name = "CarregaCSPatio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub cmdF1_Click()
If Not (Me.lvOC.SelectedItem Is Nothing) Then
    Me.txt2Lote.text = Me.lvOC.ListItems(Val(Me.lvOC.SelectedItem)).SubItems(2)
    Me.txt2Ordem.text = Me.lvOC.ListItems(Val(Me.lvOC.SelectedItem)).SubItems(1)
    Carrega_LvLocal
    Me.frmLocais.Visible = True
End If
End Sub

Private Sub cmdF2_Click()
  Dim Tb_Cv As ADODB.Recordset
  Dim TestaEntrada As Boolean
  Dim Resp As Integer
  Dim nIDBL As Long
  Dim X As String
  Dim TbNovo As ADODB.Recordset
  Dim Dir_Sistema As String
  Dim VLTL As Boolean

If Val(TxtLote.text) = 0 Then
        Me.txtCritica.text = "Escolha uma carga"
        Me.txtCritica.ForeColor = vbRed
        Exit Sub
    End If
    
   
VLTL = False
   Dir_Sistema = App.Path & "\ConsultaDocumentos\ConsultaDocColetor.exe "
    SQl = "SELECT isnull(max(flag_ltl),0) FROM TB_SOLICITACAO_LTL A INNER JOIN TB_BL B ON A.LOTE = B.AUTONUM  WHERE b.autonum=" & Val(TxtLote.text)
    If Nnull(db.Execute(SQl).Fields(0), 0) = 1 Then
         VLTL = True
         TestaEntrada = True
         Call Shell(Dir_Sistema & TrabDsn & " " & UsuBanco & " " & PasBanco & " " & Cod_Usuario & " " & Val(TxtLote.text) & " " & TestaEntrada & " " & VLTL, vbNormalFocus)
    Else
        VLTL = False
        SQl = "SELECT max(autonum) AUTONUM  FROM TB_AG_CS  WHERE LOTE=" & Val(TxtLote.text)
        PRSet TbNovo, SQl
        If Nnull(TbNovo!autonum, 0) > 0 Then
            TestaEntrada = True
            Call Shell(Dir_Sistema & TrabDsn & " " & UsuBanco & " " & PasBanco & " " & Cod_Usuario & " " & Nnull(TbNovo!autonum, 0) & " " & TestaEntrada & " " & VLTL, vbNormalFocus)
        End If
    End If

  
 


End Sub

Sub cmdF4_Click()
If Not (Me.lvOC.SelectedItem Is Nothing) Then
    If Nnull(Val(Me.lvOC.SelectedItem), 0) = 0 Then
        Me.txtCritica.text = "Ordem de carregamento não selecionada"
        Me.txtCritica.ForeColor = vbRed
        Exit Sub
    End If
    
    OrdemAnt = Nnull(Val(Me.lvOC.SelectedItem), 0)
    
    
    If Me.txtMarcante.text = "" Then
        Me.txtCritica.text = "Marcante não informado/inválido"
        Me.txtCritica.ForeColor = vbRed
        Exit Sub
    End If
    
    
    If Me.txtAutonumCS.text <> Nnull(Me.lvOC.ListItems(Val(Me.lvOC.SelectedItem)).SubItems(7), 0) Then
        Me.txtCritica.text = "Carga Divergente da Ordem de carregamento"
        Me.txtCritica.ForeColor = vbRed
        Exit Sub
    End If
    
    
    If Me.TXTLOCAL.text = "CAM" Then
        Me.txtCritica.text = "Carga já carregada"
        Me.txtCritica.ForeColor = vbRed
        Exit Sub
    End If
    
    
    If Val(Me.TXTQTDE) > Val(Me.TXTQTDE.Tag) Then
        Me.txtCritica.text = "Quantidade indisponível"
        Me.txtCritica.ForeColor = vbRed
        Exit Sub
    End If
    
    SQl = "UPDATE SGIPA..TB_MARCANTES SET PLACA_C='" & Left(Me.DataCombo1.text, 8) & "'  WHERE AUTONUM=" & Val(Me.txtMarcante.text)
    db.Execute SQl
    
    db.BeginTrans
    
    If Val(Me.TXTQTDE) = Val(Me.TXTQTDE.Tag) Then
    
        SQl = "UPDATE SGIPA..TB_CARGA_SOLTA_YARD SET QUANTIDADE=0 WHERE AUTONUM=" & Me.txtAutonum_Cs_Yard
        db.Execute SQl
        
        SQl = "INSERT INTO SGIPA..TB_CARGA_SOLTA_YARD (AUTONUM_CS,ARMAZEM,YARD,ORIGEM,QUANTIDADE,MOTIVO) VALUES "
        SQl = SQl & "(" & Me.txtAutonumCS.text & "," & Me.txtArmazem.Tag & ",'CAM','I'," & Me.TXTQTDE.text & ",8)"
        db.Execute SQl
    Else
    
        SQl = "UPDATE SGIPA..TB_CARGA_SOLTA_YARD SET QUANTIDADE=QUANTIDADE - " & Val(Me.TXTQTDE) & " WHERE AUTONUM=" & Me.txtAutonum_Cs_Yard
        db.Execute SQl
        
        SQl = "INSERT INTO SGIPA..TB_CARGA_SOLTA_YARD (AUTONUM_CS,ARMAZEM,YARD,ORIGEM,QUANTIDADE,MOTIVO) VALUES "
        SQl = SQl & "(" & Me.txtAutonumCS.text & "," & Me.txtArmazem.Tag & ",'CAM','I'," & Me.TXTQTDE.text & ",8)"
        db.Execute SQl
    
    End If
    
    Dim Autonum_Cs_Yard As Long
    Autonum_Cs_Yard = db.Execute("SELECT IDENT_CURRENT('TB_CARGA_SOLTA_YARD')")(0)
    
    If Val(Me.TXTQTDE) = Val(Me.TXTQTDE.Tag) Then
        'Sql = "UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CS_YARD=" & Autonum_Cs_Yard & ",PLACA_C='" & Left(Me.DataCombo1.text, 8) & "' WHERE AUTONUM=" & Val(Me.txtMarcante.text)
        SQl = "UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CS_YARD=" & Autonum_Cs_Yard & " WHERE AUTONUM=" & Val(Me.txtMarcante.text)
        db.Execute SQl
    Else
        'Sql = "UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CS_YARD=" & Autonum_Cs_Yard & ",PLACA_C='" & Left(Me.DataCombo1.text, 8) & "',VOLUMES=" & Val(Me.TXTQTDE) & " WHERE AUTONUM=" & Val(Me.txtMarcante.text)
        SQl = "UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CS_YARD=" & Autonum_Cs_Yard & " ,VOLUMES=" & Val(Me.TXTQTDE) & " WHERE AUTONUM=" & Val(Me.txtMarcante.text)
        db.Execute SQl
        
        
        SQl = "UPDATE SGIPA..tb_temp_marcante SET QTDE=" & Val(Me.TXTQTDE) & " WHERE MARCANTE='" & Me.txtMarcante.text & "'"
        db.Execute SQl
        
    End If
    
    SQl = "INSERT INTO SGIPA..TB_HIST_SHIFTING_CS(MARCANTE,ARMAZEM,YARD,DT_MOV,USUARIO) VALUES ("
    SQl = SQl & Val(Me.txtMarcante.text) & ","
    SQl = SQl & "0,"
    SQl = SQl & "'CAM',"
    SQl = SQl & "GETDATE()," & Cod_Usuario & ")"
    db.Execute SQl
    
    
    db.CommitTrans
    
    If Val(Me.TXTQTDE) = Val(Me.TXTQTDE.Tag) Then
        Me.txtCritica.text = "ITEM CARREGADO"
    Else
        Me.txtCritica.text = "ITEM PARCIALMENTE CARREGADO"
    End If
    
    Me.txtCritica.ForeColor = vbGreen
    
    
    Carrega_LvCarregamentos
    
    
    
    Carrega_LvOrdem
    Set lvOC.SelectedItem = lvOC.ListItems(OrdemAnt)
    
    
    Limpa
    Me.txtMarcante.text = ""
    Me.txtMarcante.SetFocus
End If
End Sub

Sub cmdF5_Click()
If Me.txtMarcante.text = "" Then
    Me.txtCritica.text = "Marcante não informado/inválido"
    Me.txtCritica.ForeColor = vbRed
    Exit Sub
End If


If Me.TXTLOCAL.text <> "CAM" Then
    Me.txtCritica.text = "Carga ainda não carregada"
    Me.txtCritica.ForeColor = vbRed
    Exit Sub
End If

db.BeginTrans
SQl = "UPDATE SGIPA..TB_CARGA_SOLTA_YARD SET QUANTIDADE=0 WHERE AUTONUM=" & Me.txtAutonum_Cs_Yard
db.Execute SQl

SQl = "INSERT INTO SGIPA..TB_CARGA_SOLTA_YARD (AUTONUM_CS,ARMAZEM,YARD,ORIGEM,QUANTIDADE,MOTIVO) VALUES "
SQl = SQl & "(" & Me.txtAutonumCS.text & "," & Me.txtArmazem.Tag & ",'CANCC','I'," & Me.TXTQTDE.text & ",8)"
db.Execute SQl



Autonum_Cs_Yard = db.Execute("SELECT IDENT_CURRENT('TB_CARGA_SOLTA_YARD')")(0)

SQl = "UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CS_YARD=" & Autonum_Cs_Yard & ",PLACA_C=NULL WHERE AUTONUM=" & Val(Me.txtMarcante.text)
db.Execute SQl

db.CommitTrans

Me.txtCritica.text = "ITEM CANCELADO"
Me.txtCritica.ForeColor = vbGreen

Limpa
Me.txtMarcante.text = ""
Me.txtMarcante.SetFocus

Carrega_LvCarregamentos
Carrega_LvOrdem

End Sub

Private Sub Command1_Click()
Carrega_Veiculos
End Sub

Private Sub DataCombo1_Change()
Carrega_LvOrdem
Carrega_LvCarregamentos
End Sub

Private Sub Carrega_LvOrdem()

Dim i As Integer
Me.lvOC.ListItems.Clear

Dim itmx As ListItem
Dim SQl As String

SQl = "SELECT "
SQl = SQl & " A.PLACA_C,"
SQl = SQl & " A.PLACA_CARRETA,"
SQl = SQl & " A.MODELO,"
SQl = SQl & " A.ORDEM_CARREG,"
SQl = SQl & " A.NUM_OC,"
SQl = SQl & " A.QUANTIDADE,"
SQl = SQl & " A.AUTONUMCS,"
SQl = SQl & " A.LOTE,"
SQl = SQl & " A.ITEM,"
SQl = SQl & " A.EMBALAGEM, ISNULL(B.QTDE_CARREGADA,0) AS QTDE_CARREGADA from SGIPA..VW_CAM_CARREGAMENTO A LEFT JOIN ("
SQl = SQl & " SELECT  SUM(volumes) AS QTDE_CARREGADA, AUTONUM_CARGA AS AUTONUMCS "
SQl = SQl & " from sgipa..tb_marcantes M INNER JOIN SGIPA..TB_CARGA_SOLTA_YARD Y ON M.AUTONUM_CS_YARD=Y.AUTONUM "
SQl = SQl & " WHERE "
SQl = SQl & " M.VOLUMES>0 AND Y.YARD='CAM' AND (M.PLACA_C IS NULL OR M.PLACA_C='" & Left$(Me.DataCombo1.text, 8) & "') "
SQl = SQl & " AND y.PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio) & ""
SQl = SQl & " GROUP BY M.AUTONUM_CARGA "
SQl = SQl & " ) B ON A.AUTONUMCS=B.AUTONUMCS "
SQl = SQl & " WHERE A.PLACA_C='" & Left$(Me.DataCombo1.text, 8) & "'"
SQl = SQl & " ORDER BY A.LOTE,A.ITEM "


i = 1
Set tb1 = New ADODB.Recordset
tb1.Open SQl, db, 1, 1

While Not tb1.EOF
    
    Set itmx = lvOC.ListItems.Add(, , i)
    
    itmx.SubItems(1) = Nnull(tb1!num_oc, 1)
    itmx.SubItems(2) = Nnull(tb1!Lote, 1)
    itmx.SubItems(3) = Nnull(tb1!Item, 0)
    itmx.SubItems(4) = Nnull(tb1!QTDE_CARREGADA, 1) & "/" & Nnull(tb1!quantidade, 1)
    If Nnull(tb1!QTDE_CARREGADA, 1) = Nnull(tb1!quantidade, 1) Then
   '     itmx.ForeColor.SubItems(4) = vbGreen
    End If
'
    If Nnull(tb1!QTDE_CARREGADA, 1) <> Nnull(tb1!quantidade, 1) Then
    '        itmx.ForeColor.SubItems(4) = vbRed
    End If
    

    itmx.SubItems(5) = Nnull(tb1!embalagem, 1)
    
    itmx.SubItems(6) = Nnull(tb1!ordem_carreg, 0)
    itmx.SubItems(7) = Nnull(tb1!AutonumCs, 0)
    
    i = i + 1
    tb1.MoveNext
Wend
tb1.Close



End Sub

Private Sub Carrega_LvLocal()

Dim i As Integer
Me.lvLocal.ListItems.Clear

Dim itmx As ListItem
Dim SQl As String

SQl = "SELECT MARCANTE, QTDE, DESCR_ARMAZEM, POSICAO FROM SGIPA..VW_INVENT_ARMAZEM "
'Sql = Sql & " WHERE LOTE=" & Me.txt2Lote.text & " "
SQl = SQl & " WHERE AUTONUMCS=" & Nnull(Me.lvOC.ListItems(Val(Me.lvOC.SelectedItem)).SubItems(7), 0)
SQl = SQl & " and qtde>0 ORDER BY DESCR_ARMAZEM, POSICAO, MARCANTE "


i = 1
Set tb1 = New ADODB.Recordset
tb1.Open SQl, db, 1, 1

While Not tb1.EOF
    
    Set itmx = lvLocal.ListItems.Add(, , i)
    itmx.SubItems(1) = Format(Nnull(tb1!Marcante, 0), "000000000")
    itmx.SubItems(2) = Nnull(tb1!QTDE, 0)
    itmx.SubItems(3) = Nnull(tb1!DESCR_ARMAZEM, 1) & " " & Nnull(tb1!Posicao, 1)
    
    i = i + 1
    tb1.MoveNext
Wend
tb1.Close



End Sub



Private Sub Carrega_LvCarregamentos()



Me.lvCm.ListItems.Clear

Dim itmx As ListItem
Dim SQl As String

'MsgBox Me.lvOC.ListItems(Val(Me.lvOC.SelectedItem)).SubItems(7)


'If Me.lvCm.SelectedItem Is Nothing Then Exit Sub

If Nnull(Val(Me.lvOC.SelectedItem), 0) > 0 Then

    SQl = "SELECT "
    SQl = SQl & " M.AUTONUM AS MARCANTE,"
    SQl = SQl & " M.VOLUMES AS QTDE "
    SQl = SQl & " from SGIPA..TB_MARCANTES M INNER JOIN SGIPA..TB_CARGA_SOLTA_YARD Y ON M.AUTONUM_CS_YARD=Y.AUTONUM "
    SQl = SQl & " WHERE M.AUTONUM_CARGA=" & Nnull(Me.lvOC.ListItems(Val(Me.lvOC.SelectedItem)).SubItems(7), 0)
    SQl = SQl & " AND (M.PLACA_C IS NULL OR M.PLACA_C='" & Left$(Me.DataCombo1.text, 8) & "')"
    SQl = SQl & " AND Y.YARD='CAM' ORDER BY M.AUTONUM "
    
    
        
    Set tb1 = New ADODB.Recordset
    tb1.Open SQl, db, 1, 1
    Dim i As Integer
    
    i = 1
    While Not tb1.EOF
        Set itmx = lvCm.ListItems.Add(, , i)
        itmx.SubItems(1) = Right$("000000000" & tb1!Marcante, 9)
        itmx.SubItems(2) = Nnull(tb1!QTDE, 1)
        i = i + 1
        tb1.MoveNext
    Wend

    Me.lvOC.ListItems(Val(Me.lvOC.SelectedItem)).Selected = True

End If


End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim SQl As String
Dim OrdemAnt As Long




If Me.frmLocais.Visible = False Then

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If


    If KeyCode = vbKeyF4 Then
        Call cmdF4_Click
    End If
    
    
    If KeyCode = vbKeyF5 Then
        Call cmdF5_Click
    End If
    
    If KeyCode = vbKeyF1 Then
        Call cmdF1_Click
    End If
    If KeyCode = vbKeyF2 Then
        Call cmdF2_Click
    End If

Else

    If KeyCode = vbKeyEscape Then
        Me.frmLocais.Visible = False
    End If
    
End If


End Sub




Private Sub Form_Load()
Me.Top = 0
Me.Left = 0

Dim colx As ColumnHeader

'Inclui as colunas


Set colx = lvOC.ColumnHeaders.Add(, , "none")
colx.Width = 0


Set colx = lvOC.ColumnHeaders.Add(, , "#ORDEM")
colx.Width = 850


Set colx = lvOC.ColumnHeaders.Add(, , "LOTE")
colx.Width = 750
colx.Alignment = lvwColumnCenter

Set colx = lvOC.ColumnHeaders.Add(, , "ITEM")
colx.Width = 0
colx.Alignment = lvwColumnCenter

Set colx = lvOC.ColumnHeaders.Add(, , "QTDE")
colx.Width = 600
colx.Alignment = lvwColumnCenter

Set colx = lvOC.ColumnHeaders.Add(, , "EMBALAGEM")
colx.Width = 1200
colx.Alignment = lvwColumnLeft

Set colx = lvOC.ColumnHeaders.Add(, , "OC")
colx.Width = 0
colx.Alignment = lvwColumnLeft

Set colx = lvOC.ColumnHeaders.Add(, , "CS")
colx.Width = 0
colx.Alignment = lvwColumnLeft



Set colx = Me.lvCm.ColumnHeaders.Add(, , "NONE")
colx.Width = 0
colx.Alignment = lvwColumnLeft


Set colx = Me.lvCm.ColumnHeaders.Add(, , "MARCANTE")
colx.Width = 1800
colx.Alignment = lvwColumnLeft

Set colx = lvCm.ColumnHeaders.Add(, , "QTDE")
colx.Width = 800
colx.Alignment = lvwColumnCenter


Set colx = Me.lvLocal.ColumnHeaders.Add(, , "NONE")
colx.Width = 0
colx.Alignment = lvwColumnLeft


Set colx = Me.lvLocal.ColumnHeaders.Add(, , "MARCANTE")
colx.Width = 1050
colx.Alignment = lvwColumnLeft

Set colx = lvLocal.ColumnHeaders.Add(, , "QTDE")
colx.Width = 600
colx.Alignment = lvwColumnCenter

Set colx = lvLocal.ColumnHeaders.Add(, , "LOCAL ")
colx.Width = 1800
colx.Alignment = lvwColumnCenter




Carrega_Veiculos

End Sub


Private Sub Carrega_Veiculos()

SQl = " SELECT DISTINCT PLACA_C + ' ' + PLACA_CARRETA + ' - ' + ISNULL(MODELO,'') AS DISPLAY FROM SGIPA..VW_CAM_CARREGAMENTO WHERE PATIO " & IIf(Col_Patio = 1 Or Col_Patio = 7, " in (1,7) ", " = " & Col_Patio) & " ORDER BY 1 "
Me.Adodc1.ConnectionString = ConnectionString
Me.Adodc1.RecordSource = SQl
Me.Adodc1.Refresh
Set Me.Adodc1.Recordset.ActiveConnection = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub lvOC_Click()
Carrega_LvCarregamentos
End Sub

Private Sub lvOC_KeyUp(KeyCode As Integer, Shift As Integer)
Carrega_LvCarregamentos
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub txtMarcante_Change()
If Len(Me.txtMarcante) = 9 Then

    SQl = " SELECT"
    SQl = SQl & " M.AUTONUM AS MARCANTE,"
    SQl = SQl & " S.AUTONUM AS AUTONUMCS,"
    SQl = SQl & " S.BL AS LOTE,"
    SQl = SQl & " S.ITEM,"
    SQl = SQl & " M.VOLUMES AS QUANTIDADE,"
    SQl = SQl & " E.DESCR AS EMBALAGEM,"
    SQl = SQl & " S.MERCADORIA,"
    SQl = SQl & " S.MARCA,"
    SQl = SQl & " S.ARMAZEM_IPA AS AUTONUM_ARMAZEM,"
    SQl = SQl & " A.DESCR AS DESCR_ARMAZEM,"
    SQl = SQl & " S.CNTR AS AUTONUMCNTR,"
    SQl = SQl & " c.Id_Conteiner,"
    SQl = SQl & " Y.YARD AS POSICAO, "
    SQl = SQl & " M.AUTONUM_CS_YARD "
    SQl = SQl & " From"
    SQl = SQl & " TB_CARGA_SOLTA S INNER JOIN"
    SQl = SQl & " DTE_TB_EMBALAGENS E ON S.EMBALAGEM=E.CODE LEFT JOIN"
    SQl = SQl & " TB_CNTR_BL C ON S.CNTR=C.AUTONUM INNER JOIN"
    SQl = SQl & " TB_MARCANTES M ON S.AUTONUM=M.AUTONUM_CARGA LEFT JOIN"
    SQl = SQl & " TB_CARGA_SOLTA_YARD Y ON M.AUTONUM_CS_YARD=Y.AUTONUM LEFT JOIN "
    SQl = SQl & " TB_ARMAZENS_IPA A ON Y.ARMAZEM= A.AUTONUM "
    SQl = SQl & " Where"
    SQl = SQl & " M.AUTONUM=" & Val(Me.txtMarcante)
    Set tb1 = New ADODB.Recordset
    tb1.Open SQl, db, 1, 1
    If Not tb1.EOF Then
        Me.TxtLote.text = Nnull(tb1!Lote, 0)
        Me.TXTQTDE.text = Nnull(tb1!quantidade, 0)
        Me.TXTQTDE.Tag = Nnull(tb1!quantidade, 0)
        Me.txtAutonum_Cs_Yard = Nnull(tb1!Autonum_Cs_Yard, 0)
        Me.txtAutonumCS = Nnull(tb1!AutonumCs, 0)
        Me.TXTLOCAL.text = Nnull(tb1!Posicao, 1)
        Me.txtArmazem.text = Nnull(tb1!DESCR_ARMAZEM, 1)
        Me.txtArmazem.Tag = Nnull(tb1!Autonum_Armazem, 0)
    Else
        Limpa
    End If
tb1.Close
End If
    
    
'    Set tb1 = New ADODB.Recordset
'    SQL = "SELECT AUTONUM_ARMAZEM,"
'    SQL = SQL & "DESCR_ARMAZEM, "
'    SQL = SQL & "LOTE,"
'    SQL = SQL & "QTDE,"
'    SQL = SQL & "POSICAO,"
'    SQL = SQL & "AUTONUMCS,"
'    SQL = SQL & "AUTONUM_CS_YARD "
'    SQL = SQL & " FROM OPERADOR..VW_INVENT_ARMAZEM WHERE MARCANTE=" & Val(Me.txtMarcante) & " AND QTDE>0 "
'    tb1.Open SQL, db, 1, 1
'    If Not tb1.EOF Then
'        Me.TXTLOTE.Text = Nnull(tb1!lote, 0)
'        Me.TXTQTDE.Text = Nnull(tb1!QTDE, 0)
'        Me.txtAutonum_Cs_Yard = Nnull(tb1!Autonum_Cs_Yard, 0)
'        Me.txtAutonumCS = Nnull(tb1!autonumCs, 0)
'        Me.TXTLOCAL.Text = Nnull(tb1!Posicao, 1)
'        Me.txtArmazem.Text = Nnull(tb1!DESCR_ARMAZEM, 1)
'        Me.txtArmazem.Tag = Nnull(tb1!AUTONUM_ARMAZEM, 0)
'    Else
'        Limpa
'
'    End If
'Else
'    Limpa
'End If

End Sub

Private Sub Limpa()
    Me.TXTQTDE = "0"
    Me.TXTQTDE.Tag = "0"
    Me.TxtLote.text = ""
    Me.txtAutonum_Cs_Yard = ""
    Me.txtAutonumCS = ""
    Me.txtArmazem.text = ""
    Me.txtArmazem.Tag = 0
    Me.TXTLOCAL.text = ""
    txtCritica.text = ""
    
End Sub

