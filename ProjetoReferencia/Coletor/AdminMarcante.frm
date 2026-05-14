VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form AdminMarcante 
   Appearance      =   0  'Flat
   BackColor       =   &H00BA9F8D&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4455
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
   ScaleHeight     =   4455
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEscape 
      BackColor       =   &H00E0E0E0&
      Caption         =   "[ESC]-VOLTAR"
      Height          =   315
      Left            =   1080
      MaskColor       =   &H00FF0000&
      TabIndex        =   18
      Top             =   4110
      Width           =   1365
   End
   Begin VB.CommandButton cmdSalvar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "[F5]-SALVA"
      Height          =   345
      Left            =   2400
      MaskColor       =   &H00FF0000&
      TabIndex        =   17
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdFiltra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "[F4]-FILTRA"
      Height          =   345
      Left            =   2370
      MaskColor       =   &H00FF0000&
      TabIndex        =   16
      Top             =   390
      Width           =   1095
   End
   Begin VB.TextBox TXTAUTONUMCS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Height          =   345
      Left            =   2340
      MaxLength       =   6
      TabIndex        =   15
      Top             =   750
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox TXTQTDEN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1140
      TabIndex        =   12
      Top             =   3690
      Width           =   1155
   End
   Begin VB.TextBox TXTMARCANTE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   1140
      TabIndex        =   11
      Top             =   3330
      Width           =   2355
   End
   Begin VB.TextBox TXTTOT 
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
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   2850
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2970
      Width           =   645
   End
   Begin VB.TextBox txtEmbalagem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1110
      Width           =   2295
   End
   Begin VB.TextBox txtQtde 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   750
      Width           =   1155
   End
   Begin VB.TextBox TXTLOTE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1170
      MaxLength       =   7
      TabIndex        =   1
      Top             =   390
      Width           =   1155
   End
   Begin MSComctlLib.ListView lv 
      Height          =   1245
      Left            =   60
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1680
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   2196
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "QTDE"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   6
      Left            =   30
      TabIndex        =   13
      Top             =   3690
      Width           =   1095
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
      Index           =   5
      Left            =   30
      TabIndex        =   10
      Top             =   3330
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TTL"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   4
      Left            =   2430
      TabIndex        =   9
      Top             =   2970
      Width           =   405
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MARCANTES"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1230
      TabIndex        =   7
      Top             =   1470
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EMBALAGEM"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   3
      Left            =   60
      TabIndex        =   5
      Top             =   1110
      Width           =   1095
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
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   750
      Width           =   1095
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
      Left            =   60
      TabIndex        =   2
      Top             =   390
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "MANUTENÇÃO- MARCANTES"
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
      Height          =   315
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3465
   End
End
Attribute VB_Name = "AdminMarcante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdEscape_Click()
Unload Me
End Sub

Sub cmdFiltra_Click()
If Me.TXTLOTE.text = "" Then
    Col_Mensagem = "Informe o lote"
    frmMensagem2.Show vbModal
    Exit Sub
Else
    Sql = "SELECT * FROM SGIPA..VW_WMS_DIVERGENCIA_QTDE WHERE LOTE=" & Val(Me.TXTLOTE.text)
    Set tb1 = New ADODB.Recordset
    tb1.Open Sql, db, 1, 1
    If tb1.EOF Then
        Col_Mensagem = "Lote não encontrado como divergente"
        frmMensagem2.Show vbModal
        Exit Sub
    Else
        Me.txtEmbalagem.text = Nnull(tb1!embalagem, 1)
        Me.TXTQTDE.text = Nnull(tb1!QUANTIDADE_REAL, 0)
        Me.txtAutonumCS.text = tb1!AUTONUM_CS
        Carrega_Lv
    End If
    tb1.Close
    
    
End If
End Sub

Sub cmdSalvar_Click()
If Me.txtMarcante.text = "" Then
        Col_Mensagem = "Selecione o marcante"
        frmMensagem2.Show vbModal
        Exit Sub
    End If
    
    If Me.TXTQTDEN.text = "" Then
        Col_Mensagem = "Informe a quantidade"
        frmMensagem2.Show vbModal
        Exit Sub
    End If
    
    
    Sql = "UPDATE SGIPA..TB_MARCANTES SET VOLUMES=" & Val(Me.TXTQTDEN.text) & " WHERE autonum=" & Val(Me.txtMarcante.text)
    db.Execute Sql
    
    Me.txtMarcante.text = ""
    Me.TXTQTDEN.text = ""
    
    Dim LvAnt As Integer
    LvAnt = Me.LV.SelectedItem
    Carrega_Lv
    
    On Error Resume Next
    Me.LV.SelectedItem = LvAnt
    On Error GoTo 0
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If

If KeyCode = vbKeyF4 Then
    Call cmdFiltra_Click
End If


If KeyCode = vbKeyF5 Then
    Call cmdSalvar_Click
End If



End Sub

Private Sub Carrega_Lv()

Dim i As Integer
Me.LV.ListItems.Clear



If Me.TXTLOTE.text = "" Then Exit Sub

Dim QtdeTot As Long
QtdeTot = 0

Dim itmx As ListItem
Dim Sql As String

'Sql = "SELECT "
'Sql = Sql & " C.ITEM,I.DESCR_ARMAZEM,I.POSICAO, M.VOLUMES AS QTDE, M.AUTONUM AS MARCANTE, M.AUTONUM_CARGA AS AUTONUMCS"
'Sql = Sql & " FROM "
'Sql = Sql & " SGIPA..TB_MARCANTES M INNER JOIN SGIPA..TB_CARGA_SOLTA C ON M.AUTONUM_CARGA=C.AUTONUM "
'Sql = Sql & " LEFT JOIN OPERADOR..VW_INVENT_ARMAZEM I ON M.AUTONUM=I.MARCANTE "
'Sql = Sql & " WHERE "
'Sql = Sql & " C.BL=" & Val(Me.TXTLOTE.Text)
'Sql = Sql & " and ISNULL(I.QTDE,0)>0 "
'Sql = Sql & " ORDER BY M.AUTONUM "


Sql = " SELECT  C.ITEM,I.DESCR_ARMAZEM,I.POSICAO, M.VOLUMES AS QTDE, M.AUTONUM AS MARCANTE, M.AUTONUM_CARGA AS AUTONUMCS FROM "
Sql = Sql & " SGIPA..TB_MARCANTES M INNER JOIN SGIPA..TB_CARGA_SOLTA C ON M.AUTONUM_CARGA=C.AUTONUM LEFT JOIN"
Sql = Sql & " (SELECT DESCR_ARMAZEM,POSICAO,MARCANTE FROM OPERADOR..VW_INVENT_ARMAZEM I WHERE AUTONUMCS=" & Me.txtAutonumCS & " AND QTDE>0) I"
Sql = Sql & " ON M.AUTONUM=I.MARCANTE"
Sql = Sql & " Where c.AUTONUM = " & Me.txtAutonumCS
Sql = Sql & " AND M.DT_ASSOCIACAO IS NOT NULL "
Sql = Sql & " ORDER BY M.AUTONUM "


i = 1
Set tb1 = New ADODB.Recordset
tb1.Open Sql, db, 1, 1

While Not tb1.EOF
    
    Set itmx = LV.ListItems.Add(, , i)
    itmx.SubItems(1) = Nnull(tb1!DESCR_ARMAZEM, 1)
    itmx.SubItems(2) = Nnull(tb1!Posicao, 1)
    itmx.SubItems(3) = Nnull(tb1!Item, 0)
    itmx.SubItems(4) = Nnull(tb1!QTDE, 0)
    itmx.SubItems(5) = Right("000000000" & Nnull(tb1!Marcante, 1), 9)
    itmx.SubItems(6) = Nnull(tb1!AutonumCs, 0)
    
    QtdeTot = QtdeTot + Nnull(tb1!QTDE, 0)
    
    i = i + 1
    tb1.MoveNext
Wend
tb1.Close

Me.TXTTOT.text = QtdeTot
If Me.TXTTOT.text = Me.TXTQTDE.text Then
    Me.TXTTOT.ForeColor = vbGreen
Else
    Me.TXTTOT.ForeColor = vbRed
End If

End Sub



Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Limpa

Dim colx As ColumnHeader

'Inclui as colunas


Set colx = LV.ColumnHeaders.Add(, , "none")
colx.Width = 0


Set colx = LV.ColumnHeaders.Add(, , "ARMAZEM")
colx.Width = 750

Set colx = LV.ColumnHeaders.Add(, , "POSICAO")
colx.Width = 800
colx.Alignment = lvwColumnCenter


Set colx = LV.ColumnHeaders.Add(, , "ITEM")
colx.Width = 400
colx.Alignment = lvwColumnCenter


Set colx = LV.ColumnHeaders.Add(, , "QTD")
colx.Width = 450
colx.Alignment = lvwColumnCenter

Set colx = LV.ColumnHeaders.Add(, , "MARCANTE")
colx.Width = 1200
colx.Alignment = lvwColumnLeft


Set colx = LV.ColumnHeaders.Add(, , "AUTONUMCS")
colx.Width = 0
colx.Alignment = lvwColumnLeft




End Sub



Private Sub Limpa()
Me.txtEmbalagem.text = ""
Me.TXTQTDE.text = ""
Me.txtMarcante.text = ""
Me.TXTQTDEN.text = ""
Me.TXTTOT.text = ""
Me.TXTTOT.ForeColor = vbBlack
Me.txtMarcante.text = ""
Me.txtAutonumCS.text = ""
Carrega_Lv
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub lv_Click()

If Me.LV.SelectedItem > 0 Then
     Me.txtMarcante.text = Nnull(Me.LV.ListItems(Val(Me.LV.SelectedItem)).SubItems(5), 1)
End If

End Sub

