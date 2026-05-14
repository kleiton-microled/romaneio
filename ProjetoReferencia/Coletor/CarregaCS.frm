VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form CarregaCS 
   Appearance      =   0  'Flat
   BackColor       =   &H00BA9F8D&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3705
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
   ScaleHeight     =   4545
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "[F3] LIMPAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4140
      Width           =   645
   End
   Begin VB.Frame frmLocais 
      Appearance      =   0  'Flat
      BackColor       =   &H00BA9F8D&
      ForeColor       =   &H00000000&
      Height          =   3915
      Left            =   0
      TabIndex        =   23
      Top             =   120
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
         Top             =   990
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
      Height          =   375
      Left            =   2160
      TabIndex        =   32
      Top             =   4140
      Width           =   615
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
      Height          =   375
      Left            =   2835
      TabIndex        =   31
      Top             =   4140
      Width           =   795
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
      Height          =   375
      Left            =   1410
      TabIndex        =   30
      Top             =   4140
      Width           =   705
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
      Height          =   375
      Left            =   30
      TabIndex        =   29
      Top             =   4140
      Width           =   645
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
      Bindings        =   "CarregaCS.frx":0000
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
      Height          =   255
      Index           =   2
      Left            =   2610
      TabIndex        =   13
      Top             =   4110
      Width           =   975
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
Attribute VB_Name = "CarregaCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HorarioNumero As Date
Dim HorarioEnter As Date
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

If Val(TXTLOTE.text) = 0 Then
        Me.txtCritica.text = "Escolha uma carga"
        Me.txtCritica.ForeColor = vbRed
        Exit Sub
    End If
    
   
VLTL = False
   Dir_Sistema = App.Path & "\ConsultaDocumentos\ConsultaDocColetor.exe "
    Sql = "SELECT isnull(max(flag_ltl),0) FROM TB_SOLICITACAO_LTL A INNER JOIN TB_BL B ON A.LOTE = B.AUTONUM  WHERE b.autonum=" & Val(TXTLOTE.text)
    If Nnull(db.Execute(Sql).Fields(0), 0) = 1 Then
         VLTL = True
         TestaEntrada = True
         Call Shell(Dir_Sistema & TrabDsn & " " & UsuBanco & " " & PasBanco & " " & Cod_Usuario & " " & Val(TXTLOTE.text) & " " & TestaEntrada & " " & VLTL, vbNormalFocus)
    Else
        VLTL = False
        Sql = "SELECT max(autonum) AUTONUM  FROM TB_AG_CS  WHERE LOTE=" & Val(TXTLOTE.text)
        PRSet TbNovo, Sql
        If Nnull(TbNovo!AUTONUM, 0) > 0 Then
            TestaEntrada = True
            Call Shell(Dir_Sistema & TrabDsn & " " & UsuBanco & " " & PasBanco & " " & Cod_Usuario & " " & Nnull(TbNovo!AUTONUM, 0) & " " & TestaEntrada & " " & VLTL, vbNormalFocus)
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
    
    
    If Me.TXTMARCANTE.text = "" Then
        Me.txtCritica.text = "Marcante não informado/inválido"
        Me.txtCritica.ForeColor = vbRed
        Exit Sub
    End If
    
    
    If Me.TXTAUTONUMCS.text <> Nnull(Me.lvOC.ListItems(Val(Me.lvOC.SelectedItem)).SubItems(7), 0) Then
        Me.txtCritica.text = "Carga Divergente da Ordem de carregamento"
        Me.txtCritica.ForeColor = vbRed
        Exit Sub
    End If
    
    
    If Me.TXTLOCAL.text = "CAM" Then
        Me.txtCritica.text = "Carga já carregada"
        Me.txtCritica.ForeColor = vbRed
        Exit Sub
    End If
    
    
    If Val(Me.txtQtde) > Val(Me.txtQtde.Tag) Then
        Me.txtCritica.text = "Quantidade indisponível"
        Me.txtCritica.ForeColor = vbRed
        Exit Sub
    End If
    
    Dim rsaux As ADODB.Recordset
    Dim Lote As String
    
    Sql = "SELECT bl from tb_carga_solta where autonum=" & Me.TXTAUTONUMCS.text
    PRSet rsaux, Sql
    If Not rsaux.EOF Then
      Sql = "SELECT DBO.FC_VALIDA_SAIDA_CARGA(" & rsaux!BL & "," & Nnull(Val(Me.lvOC.SelectedItem), 0) & "," & Me.TXTAUTONUMCS.text & ") AS RETORNO"
      PRSet rsaux, Sql
      If rsaux!Retorno <> "OK" Then
          Me.txtCritica.text = rsaux!Retorno
          Me.txtCritica.ForeColor = vbRed
          Exit Sub
      End If
    Else
       Me.txtCritica.text = "Lote não encontrado"
       Me.txtCritica.ForeColor = vbRed
       Exit Sub
    End If
    
    Sql = "UPDATE SGIPA..TB_MARCANTES SET ORDEM=" & Nnull(Val(Me.lvOC.SelectedItem), 0) & ", PLACA_C='" & Left(Me.DataCombo1.text, 8) & "'  WHERE AUTONUM=" & Val(Me.TXTMARCANTE.text)
    db.Execute Sql
    
    db.BeginTrans
    
    Dim Autonum_Cs_Yard As Long
    'Autonum_Cs_Yard = db.Execute("SELECT NEXT VALUE FOR SEQ_CARGA_SOLTA_YARD")(0)
    
    If Val(Me.txtQtde) = Val(Me.txtQtde.Tag) Then
    
        Sql = "UPDATE SGIPA..TB_CARGA_SOLTA_YARD SET QUANTIDADE=0 WHERE AUTONUM=" & Me.txtAutonum_Cs_Yard
        db.Execute Sql
        
        'Sql = "INSERT INTO SGIPA..TB_CARGA_SOLTA_YARD (autonum,AUTONUM_CS,ARMAZEM,YARD,ORIGEM,QUANTIDADE,MOTIVO) VALUES "
        'Sql = Sql & "(" & Autonum_Cs_Yard & "," & Me.txtAutonumCS.text & "," & Me.txtArmazem.Tag & ",'CAM','I'," & Me.TXTQTDE.text & ",8)"
        
        Sql = "INSERT INTO SGIPA..TB_CARGA_SOLTA_YARD (AUTONUM_CS,ARMAZEM,YARD,ORIGEM,QUANTIDADE,MOTIVO) VALUES "
        Sql = Sql & "(" & Me.TXTAUTONUMCS.text & "," & Me.txtArmazem.Tag & ",'CAM','I'," & Me.txtQtde.text & ",8)"
        
        db.Execute Sql
    Else
    
        Sql = "UPDATE SGIPA..TB_CARGA_SOLTA_YARD SET QUANTIDADE=QUANTIDADE - " & Val(Me.txtQtde) & " WHERE AUTONUM=" & Me.txtAutonum_Cs_Yard
        db.Execute Sql
        'Sql = "INSERT INTO SGIPA..TB_CARGA_SOLTA_YARD (autonum, AUTONUM_CS,ARMAZEM,YARD,ORIGEM,QUANTIDADE,MOTIVO) VALUES "
        'Sql = Sql & "(" & Autonum_Cs_Yard & "," & Me.txtAutonumCS.text & "," & Me.txtArmazem.Tag & ",'CAM','I'," & Me.TXTQTDE.text & ",8)"
        Sql = "INSERT INTO SGIPA..TB_CARGA_SOLTA_YARD (AUTONUM_CS,ARMAZEM,YARD,ORIGEM,QUANTIDADE,MOTIVO) VALUES "
        Sql = Sql & "(" & Me.TXTAUTONUMCS.text & "," & Me.txtArmazem.Tag & ",'CAM','I'," & Me.txtQtde.text & ",8)"
        db.Execute Sql
    
    End If
    
    'Autonum_Cs_Yard = db.Execute("SELECT IDENT_CURRENT('TB_CARGA_SOLTA_YARD')")(0)
    Autonum_Cs_Yard = db.Execute("SELECT ISNULL(MAX(AUTONUM),0) FROM TB_CARGA_SOLTA_YARD WHERE AUTONUM_CS=" & Me.TXTAUTONUMCS.text)(0)
    
    If Val(Me.txtQtde) = Val(Me.txtQtde.Tag) Then
        'Sql = "UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CS_YARD=" & Autonum_Cs_Yard & ",PLACA_C='" & Left(Me.DataCombo1.text, 8) & "' WHERE AUTONUM=" & Val(Me.txtMarcante.text)
        Sql = "UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CS_YARD=" & Autonum_Cs_Yard & " WHERE AUTONUM=" & Val(Me.TXTMARCANTE.text)
        db.Execute Sql
    Else
        'Sql = "UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CS_YARD=" & Autonum_Cs_Yard & ",PLACA_C='" & Left(Me.DataCombo1.text, 8) & "',VOLUMES=" & Val(Me.TXTQTDE) & " WHERE AUTONUM=" & Val(Me.txtMarcante.text)
        Sql = "UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CS_YARD=" & Autonum_Cs_Yard & " ,VOLUMES=" & Val(Me.txtQtde) & " WHERE AUTONUM=" & Val(Me.TXTMARCANTE.text)
        db.Execute Sql
        Sql = "UPDATE SGIPA..tb_temp_marcante SET QTDE=" & Val(Me.txtQtde) & " WHERE MARCANTE='" & Me.TXTMARCANTE.text & "'"
        db.Execute Sql
    End If
    
    Sql = "INSERT INTO SGIPA..TB_HIST_SHIFTING_CS(MARCANTE,ARMAZEM,YARD,DT_MOV,USUARIO) VALUES ("
    Sql = Sql & Val(Me.TXTMARCANTE.text) & ","
    Sql = Sql & "0,"
    Sql = Sql & "'CAM',"
    Sql = Sql & "GETDATE()," & Cod_Usuario & ")"
    db.Execute Sql
    db.CommitTrans
    If Val(Me.txtQtde) = Val(Me.txtQtde.Tag) Then
        Me.txtCritica.text = "ITEM CARREGADO"
    Else
        Me.txtCritica.text = "ITEM PARCIALMENTE CARREGADO"
    End If
    
    Me.txtCritica.ForeColor = vbGreen
    Carrega_LvCarregamentos
    
    
    
    Carrega_LvOrdem
    Set lvOC.SelectedItem = lvOC.ListItems(OrdemAnt)
    
    
    Limpa
    Me.TXTMARCANTE.text = ""
    Me.TXTMARCANTE.SetFocus
End If
End Sub

Sub cmdF5_Click()
If Me.TXTMARCANTE.text = "" Then
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
Sql = "UPDATE SGIPA..TB_CARGA_SOLTA_YARD SET QUANTIDADE=0 WHERE AUTONUM=" & Me.txtAutonum_Cs_Yard
db.Execute Sql

Sql = "INSERT INTO SGIPA..TB_CARGA_SOLTA_YARD (AUTONUM_CS,ARMAZEM,YARD,ORIGEM,QUANTIDADE,MOTIVO) VALUES "
Sql = Sql & "(" & Me.TXTAUTONUMCS.text & "," & Me.txtArmazem.Tag & ",'CANCC','I'," & Me.txtQtde.text & ",8)"
db.Execute Sql



'Autonum_Cs_Yard = db.Execute("SELECT IDENT_CURRENT('TB_CARGA_SOLTA_YARD')")(0)
Autonum_Cs_Yard = db.Execute("SELECT ISNULL(MAX(AUTONUM),0) FROM TB_CARGA_SOLTA_YARD WHERE AUTONUM_CS=" & Me.TXTAUTONUMCS.text)(0)

Sql = "UPDATE SGIPA..TB_MARCANTES SET AUTONUM_CS_YARD=" & Autonum_Cs_Yard & ",PLACA_C=NULL WHERE AUTONUM=" & Val(Me.TXTMARCANTE.text)
db.Execute Sql

db.CommitTrans

Me.txtCritica.text = "ITEM CANCELADO"
Me.txtCritica.ForeColor = vbGreen

Limpa
Me.TXTMARCANTE.text = ""
Me.TXTMARCANTE.SetFocus

Carrega_LvCarregamentos
Carrega_LvOrdem

End Sub

Private Sub Command1_Click()
Carrega_Veiculos
End Sub

Private Sub Command2_Click()
Call Limpa
If Me.TXTMARCANTE.Enabled = True Then
    Me.TXTMARCANTE.SetFocus
End If

End Sub

Private Sub DataCombo1_Change()
Carrega_LvOrdem
Carrega_LvCarregamentos
End Sub

Private Sub Carrega_LvOrdem()

Dim i As Integer
Me.lvOC.ListItems.Clear

Dim itmx As ListItem
Dim Sql As String

Sql = "SELECT "
Sql = Sql & " A.PLACA_C,"
Sql = Sql & " A.PLACA_CARRETA,"
Sql = Sql & " A.MODELO,"
Sql = Sql & " A.ORDEM_CARREG,"
Sql = Sql & " A.NUM_OC,"
Sql = Sql & " A.QUANTIDADE,"
Sql = Sql & " A.AUTONUMCS,"
Sql = Sql & " A.LOTE,"
Sql = Sql & " A.ITEM,"
Sql = Sql & " A.EMBALAGEM, ISNULL(B.QTDE_CARREGADA,0) AS QTDE_CARREGADA from SGIPA..VW_CAM_CARREGAMENTO A LEFT JOIN ("
Sql = Sql & " SELECT  SUM(volumes) AS QTDE_CARREGADA, AUTONUM_CARGA AS AUTONUMCS "
Sql = Sql & " from sgipa..tb_marcantes M INNER JOIN SGIPA..TB_CARGA_SOLTA_YARD Y ON M.AUTONUM_CS_YARD=Y.AUTONUM "
Sql = Sql & " WHERE "
Sql = Sql & " M.VOLUMES>0 AND Y.YARD='CAM' AND (M.PLACA_C IS NULL OR M.PLACA_C='" & Left$(Me.DataCombo1.text, 8) & "') GROUP BY M.AUTONUM_CARGA "
Sql = Sql & " ) B ON A.AUTONUMCS=B.AUTONUMCS "
Sql = Sql & " WHERE A.PLACA_C='" & Left$(Me.DataCombo1.text, 8) & "'"
Sql = Sql & " ORDER BY A.LOTE,A.ITEM "


i = 1
Set tb1 = New ADODB.Recordset
tb1.Open Sql, db, 1, 1

While Not tb1.EOF
    
    Set itmx = lvOC.ListItems.Add(, , i)
    
    itmx.SubItems(1) = Nnull(tb1!num_oc, 1)
    itmx.SubItems(2) = Nnull(tb1!Lote, 1)
    itmx.SubItems(3) = Nnull(tb1!Item, 0)
    itmx.SubItems(4) = Nnull(tb1!QTDE_CARREGADA, 1) & "/" & Nnull(tb1!QUANTIDADE, 1)
    If Nnull(tb1!QTDE_CARREGADA, 1) = Nnull(tb1!QUANTIDADE, 1) Then
   '     itmx.ForeColor.SubItems(4) = vbGreen
    End If
'
    If Nnull(tb1!QTDE_CARREGADA, 1) <> Nnull(tb1!QUANTIDADE, 1) Then
    '        itmx.ForeColor.SubItems(4) = vbRed
    End If
    

    itmx.SubItems(5) = Nnull(tb1!embalagem, 1)
    
    itmx.SubItems(6) = Nnull(tb1!ORDEM_CARREG, 0)
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
Dim Sql As String

Sql = "SELECT MARCANTE, QTDE, DESCR_ARMAZEM, POSICAO FROM SGIPA..VW_INVENT_ARMAZEM "
'Sql = Sql & " WHERE LOTE=" & Me.txt2Lote.text & " "
Sql = Sql & " WHERE AUTONUMCS=" & Nnull(Me.lvOC.ListItems(Val(Me.lvOC.SelectedItem)).SubItems(7), 0)
Sql = Sql & " and qtde>0 ORDER BY DESCR_ARMAZEM, POSICAO, MARCANTE "


i = 1
Set tb1 = New ADODB.Recordset
tb1.Open Sql, db, 1, 1

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
Dim Sql As String

'MsgBox Me.lvOC.ListItems(Val(Me.lvOC.SelectedItem)).SubItems(7)


'If Me.lvCm.SelectedItem Is Nothing Then Exit Sub

If Nnull(Val(Me.lvOC.SelectedItem), 0) > 0 Then

    Sql = "SELECT "
    Sql = Sql & " M.AUTONUM AS MARCANTE,"
    Sql = Sql & " M.VOLUMES AS QTDE "
    Sql = Sql & " from SGIPA..TB_MARCANTES M INNER JOIN SGIPA..TB_CARGA_SOLTA_YARD Y ON M.AUTONUM_CS_YARD=Y.AUTONUM "
    Sql = Sql & " WHERE M.AUTONUM_CARGA=" & Nnull(Me.lvOC.ListItems(Val(Me.lvOC.SelectedItem)).SubItems(7), 0)
    Sql = Sql & " AND (M.PLACA_C IS NULL OR M.PLACA_C='" & Left$(Me.DataCombo1.text, 8) & "')"
    Sql = Sql & " AND Y.YARD='CAM' ORDER BY M.AUTONUM "
    
    
        
    Set tb1 = New ADODB.Recordset
    tb1.Open Sql, db, 1, 1
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
Dim Sql As String
Dim OrdemAnt As Long




If Me.frmLocais.Visible = False Then

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If


    If KeyCode = vbKeyF4 Then
        Call cmdF4_Click
    End If
    
    If KeyCode = vbKeyF3 Then
        Call Limpa
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

If Col_Patio <> 1 Then
    Sql = " SELECT DISTINCT PLACA_C + ' ' + PLACA_CARRETA + ' - ' + ISNULL(MODELO,'') AS DISPLAY FROM SGIPA..VW_CAM_CARREGAMENTO WHERE PATIO=" & Col_Patio & " ORDER BY 1 "
Else
    Sql = " SELECT DISTINCT PLACA_C + ' ' + PLACA_CARRETA + ' - ' + ISNULL(MODELO,'') AS DISPLAY FROM SGIPA..VW_CAM_CARREGAMENTO WHERE PATIO in (1,7) ORDER BY 1 "
End If
Me.Adodc1.ConnectionString = ConnectionString
Me.Adodc1.RecordSource = Sql
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
If Len(Me.TXTMARCANTE) = 9 Then
    
    
    
    Dim DifSeg As Long
    If CStr(HorarioEnter) <> "00:00:00" Then
        DifSeg = Abs(DateDiff("s", HorarioEnter, HorarioNumero))
    Else
        DifSeg = 99999
    End If
    
    If (TXTMARCANTE.Tag <> "1" Or DifSeg > 1) And Me.TXTMARCANTE.text <> "" And Wflag_carregamento_marcante_bip <> 0 Then
        X = MsgBox("Campo não permite digitação manual", vbExclamation)
        Me.TXTMARCANTE.text = ""
    Else
        Carrega_Dados
        
        If Wflag_carregamento_marcante_bip = 1 Then
            TXTMARCANTE.Enabled = False
        End If
        
    End If

    
    
    'Carrega_Dados
    
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
Private Sub Carrega_Dados()
Sql = " SELECT"
    Sql = Sql & " M.AUTONUM AS MARCANTE,"
    Sql = Sql & " S.AUTONUM AS AUTONUMCS,"
    Sql = Sql & " S.BL AS LOTE,"
    Sql = Sql & " S.ITEM,"
    Sql = Sql & " M.VOLUMES AS QUANTIDADE,"
    Sql = Sql & " E.DESCR AS EMBALAGEM,"
    Sql = Sql & " S.MERCADORIA,"
    Sql = Sql & " S.MARCA,"
    Sql = Sql & " S.ARMAZEM_IPA AS AUTONUM_ARMAZEM,"
    Sql = Sql & " A.DESCR AS DESCR_ARMAZEM,"
    Sql = Sql & " S.CNTR AS AUTONUMCNTR,"
    Sql = Sql & " c.Id_Conteiner,"
    Sql = Sql & " Y.YARD AS POSICAO, "
    Sql = Sql & " M.AUTONUM_CS_YARD "
    Sql = Sql & " From"
    Sql = Sql & " TB_CARGA_SOLTA S INNER JOIN"
    Sql = Sql & " DTE_TB_EMBALAGENS E ON S.EMBALAGEM=E.CODE LEFT JOIN"
    Sql = Sql & " TB_CNTR_BL C ON S.CNTR=C.AUTONUM INNER JOIN"
    Sql = Sql & " TB_MARCANTES M ON S.AUTONUM=M.AUTONUM_CARGA LEFT JOIN"
    Sql = Sql & " TB_CARGA_SOLTA_YARD Y ON M.AUTONUM_CS_YARD=Y.AUTONUM LEFT JOIN "
    Sql = Sql & " TB_ARMAZENS_IPA A ON Y.ARMAZEM= A.AUTONUM "
    Sql = Sql & " Where"
    Sql = Sql & " M.AUTONUM=" & Val(Me.TXTMARCANTE)
    Set tb1 = New ADODB.Recordset
    tb1.Open Sql, db, 1, 1
    If Not tb1.EOF Then
        Me.TXTLOTE.text = Nnull(tb1!Lote, 0)
        Me.txtQtde.text = Nnull(tb1!QUANTIDADE, 0)
        Me.txtQtde.Tag = Nnull(tb1!QUANTIDADE, 0)
        Me.txtAutonum_Cs_Yard = Nnull(tb1!Autonum_Cs_Yard, 0)
        Me.TXTAUTONUMCS = Nnull(tb1!AutonumCs, 0)
        Me.TXTLOCAL.text = Nnull(tb1!Posicao, 1)
        Me.txtArmazem.text = Nnull(tb1!DESCR_ARMAZEM, 1)
        Me.txtArmazem.Tag = Nnull(tb1!Autonum_Armazem, 0)
    Else
        Limpa
    End If
tb1.Close
End Sub

Private Sub txtMarcante_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Me.TXTMARCANTE.Tag = "1"
    HorarioEnter = Now
End If
End Sub

Private Sub txtMarcante_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Wflag_carregamento_marcante_bip = 0 Then
    
    Sendkeys "{TAB}"

Else

'    If Len(txtMarcante.text) = 0 Then
'        HorarioNumero = Now
'    End If
    
    If TXTMARCANTE.text = "" Or Len(TXTMARCANTE.text) < 2 Then
        HorarioNumero = Now
    End If
    
    
    If Len(TXTMARCANTE.text) > 5 Then
        HorarioEnter = Now
        Me.TXTMARCANTE.Tag = "1"
    End If
    
End If

End Sub

Private Sub Limpa()
    Me.txtQtde = "0"
    Me.txtQtde.Tag = "0"
    Me.TXTLOTE.text = ""
    Me.txtAutonum_Cs_Yard = ""
    Me.TXTAUTONUMCS = ""
    Me.txtArmazem.text = ""
    Me.txtArmazem.Tag = 0
    Me.TXTLOCAL.text = ""
    txtCritica.text = ""
    
    Me.TXTMARCANTE.text = ""
    TXTMARCANTE.Enabled = True
    TXTMARCANTE.Tag = 0

    HorarioNumero = Now
    HorarioEnter = Now
End Sub

Private Sub txtMarcante_GotFocus()
Me.TXTMARCANTE.Tag = "0"
End Sub

Private Sub txtMarcante_LostFocus()
Dim DifSeg As Long
If CStr(HorarioEnter) <> "00:00:00" Then
    DifSeg = Abs(DateDiff("s", HorarioEnter, HorarioNumero))
Else
    DifSeg = 99999
End If

If (TXTMARCANTE.Tag <> "1" Or DifSeg > 1) And Me.TXTMARCANTE.text <> "" And Wflag_carregamento_marcante_bip <> 0 Then
    X = MsgBox("Campo não permite digitação manual", vbExclamation)
    Me.TXTMARCANTE.text = ""
Else
    Carrega_Dados
    
    If Wflag_carregamento_marcante_bip = 1 Then
        TXTMARCANTE.Enabled = False
    End If
    
End If
End Sub
