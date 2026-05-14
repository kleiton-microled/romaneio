VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form MovimentaCS 
   Appearance      =   0  'Flat
   BackColor       =   &H00BA9F8D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimentação"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3345
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
   ScaleHeight     =   4425
   ScaleWidth      =   3345
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmMarcantes 
      Appearance      =   0  'Flat
      BackColor       =   &H00BA9F8D&
      ForeColor       =   &H80000008&
      Height          =   2025
      Left            =   0
      TabIndex        =   5
      Top             =   2100
      Width           =   3525
      Begin VB.TextBox txtQtdeMarcante 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   510
         MaxLength       =   4
         TabIndex        =   10
         Top             =   510
         Width           =   615
      End
      Begin VB.TextBox txtMarcante 
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
         Left            =   900
         MaxLength       =   9
         TabIndex        =   9
         Top             =   120
         Width           =   2385
      End
      Begin VB.ListBox lstMarcantes 
         Height          =   1035
         Left            =   30
         TabIndex        =   8
         Top             =   870
         Width           =   3225
      End
      Begin VB.CommandButton cmdAdiciona 
         Caption         =   "[F1]ADICIONA"
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
         Left            =   1110
         TabIndex        =   7
         Top             =   540
         Width           =   1125
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "[F2]REMOVE"
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
         Left            =   2250
         TabIndex        =   6
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[F2]-APAGA"
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
         Index           =   1
         Left            =   2670
         TabIndex        =   14
         Top             =   510
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[F1]-ADICIONA"
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
         Index           =   0
         Left            =   930
         TabIndex        =   13
         Top             =   630
         Width           =   1050
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
         TabIndex        =   12
         Top             =   510
         Width           =   495
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
         TabIndex        =   11
         Top             =   120
         Width           =   885
      End
   End
   Begin VB.TextBox TXTLOTE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Top             =   270
      Width           =   855
   End
   Begin VB.CommandButton cmdFiltrar 
      Caption         =   "[F4] FILTRAR"
      Height          =   285
      Left            =   2190
      TabIndex        =   3
      Top             =   270
      Width           =   1155
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "[F5] SALVAR"
      Height          =   285
      Left            =   450
      TabIndex        =   2
      Top             =   4140
      Width           =   1185
   End
   Begin VB.CommandButton cmdEscape 
      Caption         =   "[ESC]SAIR"
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   4140
      Width           =   1125
   End
   Begin MSAdodcLib.Adodc AdodcCntr 
      Height          =   405
      Left            =   2520
      Top             =   870
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   714
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
   Begin MSComctlLib.ListView lv 
      Height          =   1485
      Left            =   30
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   570
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   2619
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
      Caption         =   "LOTE"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   2
      Left            =   30
      TabIndex        =   16
      Top             =   270
      Width           =   585
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ASSOCIAÇÃO DE MARCANTES"
      ForeColor       =   &H00BA9F8D&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3345
   End
End
Attribute VB_Name = "MovimentaCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub cmdFiltrar_Click()
Carrega_Lv
End Sub

Sub cmdRemove_Click()

SQl = "DELETE FROM sgipa..TB_TEMP_MARCANTE "
SQl = SQl & " WHERE AUTONUM_CS= " & Nnull(Me.lv.ListItems(Val(Me.lv.SelectedItem)).SubItems(7), 0)
SQl = SQl & " AND MARCANTE= " & Val(Left(Me.lstMarcantes.Text, 9))
db.Execute SQl
    
Carrega_Marcantes_Temp
    
End Sub

Sub cmdSalvar_Click()

If Me.lv.ListItems.Count = 0 Then
    Col_Mensagem = "Carga nao selecionada"
    frmMensagem2.Show vbModal
    Exit Sub
End If


Dim AutonumCs As Long
Dim qteItem As Long
AutonumCs = Nnull(Me.lv.ListItems(Val(Me.lv.SelectedItem)).SubItems(7), 0)
qteItem = Nnull(Me.lv.ListItems(Val(Me.lv.SelectedItem)).SubItems(3), 0)
Dim AutonumMarcante As Long
Dim qteM As Long
qteM = 0

Dim QtdeCs As Integer
QtdeCs = Nnull(db.Execute("SELECT QUANTIDADE_REAL FROM SGIPA..TB_CARGA_SOLTA WHERE AUTONUM=" & AutonumCs)(0), 0)

For i = 0 To Me.lstMarcantes.ListCount - 1
    qteM = qteM + Val(Mid$(Me.lstMarcantes.List(i), 10))
Next i

If qteM <> qteItem Then
    Col_Mensagem = "Quantidades divergentes"
    frmMensagem2.Show vbModal
    Exit Sub
Else
    For i = 0 To Me.lstMarcantes.ListCount - 1
        AutonumMarcante = Val(Mid$(Me.lstMarcantes.List(i), 1, 9))
        qteM = Val(Mid$(Me.lstMarcantes.List(i), 10))
        
        Dim QtdeJaMarcada As Long
        QtdeJaMarcada = 0
        Set tb1 = New ADODB.Recordset

        tb1.Open "SELECT SUM(VOLUMES) AS QTOS FROM SGIPA..TB_MARCANTES WHERE AUTONUM_CARGA= " & AutonumCs, db, 1, 1
            
        If Not tb1.EOF Then
            QtdeJaMarcada = Nnull(tb1!qtos, 0)
        End If
        tb1.Close
        
        If QtdeJaMarcada + qteM <= QtdeCs Then
        
            SQl = "UPDATE SGIPA..TB_MARCANTES SET DT_ASSOCIACAO=GETDATE(),AUTONUM_CARGA=" & AutonumCs
            SQl = SQl & " ,VOLUMES=" & qteM
            SQl = SQl & " ,BL=" & Nnull(Me.lv.ListItems(Val(Me.lv.SelectedItem)).SubItems(1), 0)
            SQl = SQl & " ,ITEM=" & Nnull(Me.lv.ListItems(Val(Me.lv.SelectedItem)).SubItems(2), 0)
            'Sql = Sql & " ,CNTR=" & Nnull(Me.lv.ListItems(Val(Me.lv.SelectedItem)).SubItems(8), 0)
            SQl = SQl & " WHERE "
            SQl = SQl & " AUTONUM=" & AutonumMarcante
    
            db.Execute SQl
        End If
    Next i
    
    Col_Mensagem = "MARCANTES ASSOCIADOS COM SUCESSO"
    Col_Imagem = 2
    Load frmMensagem2
    frmMensagem2.Command1.Visible = False
    frmMensagem2.Show 1
    Me.lv.ListItems.Clear
    Me.lstMarcantes.Clear
    Me.TxtLote.SetFocus
    
    
End If

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    Carrega_Lv
End If

If KeyCode = vbKeyEscape Then
    Unload Me
End If

If KeyCode = vbKeyF1 Then

    Call cmdAdiciona_Click

End If

If KeyCode = vbKeyF2 Then

    Call cmdRemove_Click

    
End If

If KeyCode = vbKeyF5 Then
    
    Call cmdSalvar_Click
    
    
    
End If
                
                
    
End Sub

Sub cmdAdiciona_Click()
If Nnull(Me.txtMarcante, 0) = 0 Then
    Col_Mensagem = "Informe o marcante"
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
   
 
'Já foi
For i = 0 To Me.lstMarcantes.ListCount - 1
    If Me.txtMarcante.Text = Left$(Me.lstMarcantes.List(i), 9) Then
        Col_Mensagem = "Marcante já adicionado"
        Col_Imagem = 1
        frmMensagem2.Show 1
        Exit Sub
    End If
Next i
 
SQl = "SELECT AUTONUM,AUTONUM_CARGA,DT_ASSOCIACAO,AUTONUM_CEXP FROM SGIPA..TB_MARCANTES WHERE "
SQl = SQl & " AUTONUM= " & Val(Me.txtMarcante) & ""

Set tb1 = New ADODB.Recordset
tb1.Open SQl, db, 1, 1
If tb1.EOF Then
    Col_Mensagem = "Nr de marcante inválido"
    frmMensagem2.Show vbModal
    Me.txtMarcante = ""
    Exit Sub
Else
    If Nnull(tb1!DT_ASSOCIACAO, 1) <> "" Then
        Col_Mensagem = "Nr de marcante já utilizado"
        frmMensagem2.Show vbModal
        Exit Sub
    End If
End If




SQl = "INSERT INTO sgipa..TB_TEMP_MARCANTE(AUTONUM_CS,MARCANTE,QTDE) VALUES ("
SQl = SQl & Nnull(Me.lv.ListItems(Val(Me.lv.SelectedItem)).SubItems(7), 0) & ","
SQl = SQl & "'" & Me.txtMarcante.Text & "',"
SQl = SQl & Val(Me.txtQtdeMarcante.Text)
SQl = SQl & ")"
db.Execute SQl


Carrega_Marcantes_Temp

'Me.lstMarcantes.AddItem Me.txtMarcante.text & " " & Me.txtQtdeMarcante.text
Me.txtMarcante.Text = ""
Me.txtQtdeMarcante.Text = ""
Me.txtMarcante.SetFocus
End Sub

Private Sub cmdEscape_Click()
Unload Me
End Sub





Private Sub Carrega_Lv()

Dim i As Integer
Me.lv.ListItems.Clear

Dim itmx As ListItem
Dim SQl As String

SQl = "SELECT "
SQl = SQl & " LOTE,ITEM,QUANTIDADE,EMBALAGEM,ID_CONTEINER,MERCADORIA, AUTONUMCS , AUTONUMCNTR "
SQl = SQl & " FROM "
SQl = SQl & " SGIPA..VW_WMS_SEM_MARCANTE "
SQl = SQl & " WHERE 0=0 AND QUANTIDADE>0 "
If Me.TxtLote.Text = "" Then
    SQl = SQl & " AND FLAG_HISTORICO=0"
Else
    If Me.TxtLote.Text <> "" Then
        SQl = SQl & " AND LOTE=" & Val(Me.TxtLote.Text)
    End If
    
End If

SQl = SQl & " ORDER BY LOTE "

i = 1
Set tb1 = New ADODB.Recordset
tb1.Open SQl, db, 1, 1

If tb1.EOF Then

    tb1.Close
    SQl = "SELECT "
    SQl = SQl & " LOTE,ITEM,QUANTIDADE,EMBALAGEM,ID_CONTEINER,MERCADORIA, AUTONUM_CEXP AS AUTONUMCS , AUTONUMCNTR "
    SQl = SQl & " FROM "
    SQl = SQl & " SGIPA..VW_WMS_SEM_MARCANTE_CEXP "
    SQl = SQl & " WHERE 0=0 AND QUANTIDADE>0 "
    If Me.TxtLote.Text = "" Then
        SQl = SQl & " AND FLAG_HISTORICO=0"
    Else
        If Me.TxtLote.Text <> "" Then
            SQl = SQl & " AND LOTE='" & Me.TxtLote.Text & "'"
        End If
    End If
    
    SQl = SQl & " ORDER BY LOTE "
    
    Set tb1 = New ADODB.Recordset
    tb1.Open SQl, db, 1, 1

    

End If


While Not tb1.EOF
    
    Set itmx = lv.ListItems.Add(, , i)
    
    itmx.SubItems(1) = Nnull(tb1!Lote, 1)
    itmx.SubItems(2) = Nnull(tb1!Item, 1)
    itmx.SubItems(3) = Nnull(tb1!quantidade, 0)
    itmx.SubItems(4) = Nnull(tb1!embalagem, 1)
    itmx.SubItems(5) = Nnull(tb1!Id_Conteiner, 1)
    itmx.SubItems(6) = Nnull(tb1!MERCADORIA, 1)
    itmx.SubItems(7) = Nnull(tb1!AutonumCs, 0)
    itmx.SubItems(8) = Nnull(tb1!AutonumCntr, 0)
    
    i = i + 1
    tb1.MoveNext
Wend
tb1.Close



End Sub


Private Sub Carrega_Marcantes_Temp()
Dim SQl As String

On Error Resume Next

Me.lstMarcantes.Clear

SQl = "SELECT MARCANTE,QTDE FROM SGIPA..TB_TEMP_MARCANTE "
SQl = SQl & " WHERE AUTONUM_CS=" & Nnull(Me.lv.ListItems(Val(Me.lv.SelectedItem)).SubItems(7), 0)
SQl = SQl & " ORDER BY MARCANTE"
    
Set tbTemp = New ADODB.Recordset
tbTemp.Open SQl, db, 1, 1

While Not tbTemp.EOF
    'Me.lstMarcantes.AddItem Me.TXTMARCANTE.text & " " & Me.txtQtdeMarcante.text
    Me.lstMarcantes.AddItem Format(Nnull(tbTemp!Marcante, 1), "000000000") & " " & Nnull(tbTemp!QTDE, 0)
    tbTemp.MoveNext
Wend
tbTemp.Close


End Sub


Private Sub Form_Load()


Me.Top = 0
Me.Left = 0

Dim colx As ColumnHeader

'Inclui as colunas


Set colx = lv.ColumnHeaders.Add(, , "none")
colx.Width = 0


Set colx = lv.ColumnHeaders.Add(, , "LOTE")
colx.Width = 750

Set colx = lv.ColumnHeaders.Add(, , "ITEM")
colx.Width = 400
colx.Alignment = lvwColumnCenter

Set colx = lv.ColumnHeaders.Add(, , "QTD")
colx.Width = 500
colx.Alignment = lvwColumnCenter

Set colx = lv.ColumnHeaders.Add(, , "EMBALAGEM")
colx.Width = 1200
colx.Alignment = lvwColumnLeft

Set colx = lv.ColumnHeaders.Add(, , "CONTEINER")
colx.Width = 0
colx.Alignment = lvwColumnLeft

Set colx = lv.ColumnHeaders.Add(, , "MERCADORIA")
colx.Width = 880
colx.Alignment = lvwColumnLeft

Set colx = lv.ColumnHeaders.Add(, , "AUTONUMCS")
colx.Width = 0
colx.Alignment = lvwColumnLeft

Set colx = lv.ColumnHeaders.Add(, , "AUTONUMCNTR")
colx.Width = 0
colx.Alignment = lvwColumnLeft





End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub lv_Click()
Me.lstMarcantes.Clear
Carrega_Marcantes_Temp
End Sub

Private Sub lv_KeyUp(KeyCode As Integer, Shift As Integer)
Me.lstMarcantes.Clear
End Sub

Private Sub TXTLOTE_LostFocus()

TxtLote = UCase(TxtLote)


End Sub

Private Sub txtMarcante_GotFocus()
txtMarcante = ""
txtMarcante.BackColor = &HC0FFFF
End Sub

Private Sub txtMarcante_LostFocus()
txtMarcante.BackColor = &HFFFFFF
End Sub


