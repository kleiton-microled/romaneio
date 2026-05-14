VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Marcantes 
   Appearance      =   0  'Flat
   BackColor       =   &H00BA9F8D&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4452
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
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
   ScaleHeight     =   4452
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEscape 
      Caption         =   "[ESC]SAIR"
      Height          =   285
      Left            =   2640
      TabIndex        =   20
      Top             =   4170
      Width           =   885
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "[F5] SALVAR"
      Height          =   285
      Left            =   1290
      TabIndex        =   19
      Top             =   4170
      Width           =   1185
   End
   Begin VB.CommandButton cmdFiltrar 
      Caption         =   "[F4] FILTRAR"
      Height          =   345
      Left            =   1110
      TabIndex        =   16
      Top             =   2070
      Width           =   1185
   End
   Begin MSAdodcLib.Adodc AdodcCntr 
      Height          =   405
      Left            =   2520
      Top             =   840
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   720
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
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox TXTLOTE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Frame frmMarcantes 
      Appearance      =   0  'Flat
      BackColor       =   &H00BA9F8D&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   2310
      Width           =   3525
      Begin VB.CommandButton cmdRemove 
         Caption         =   "[F2]REMOVE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2490
         TabIndex        =   18
         Top             =   540
         Width           =   1035
      End
      Begin VB.CommandButton cmdAdiciona 
         Caption         =   "[F1]ADICIONA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   17
         Top             =   540
         Width           =   1065
      End
      Begin VB.ListBox lstMarcantes 
         Height          =   660
         Left            =   30
         TabIndex        =   14
         Top             =   870
         Width           =   3435
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
         TabIndex        =   2
         Top             =   120
         Width           =   2385
      End
      Begin VB.TextBox txtQtdeMarcante 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   900
         MaxLength       =   4
         TabIndex        =   3
         Top             =   510
         Width           =   525
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
            Size            =   8.4
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
         TabIndex        =   7
         Top             =   120
         Width           =   885
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
            Size            =   8.4
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
         TabIndex        =   6
         Top             =   510
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[F1]-ADICIONA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   0
         Left            =   1530
         TabIndex        =   5
         Top             =   510
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[F2]-APAGA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
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
         TabIndex        =   4
         Top             =   510
         Width           =   810
      End
   End
   Begin MSDataListLib.DataCombo dcConteiner 
      Bindings        =   "Marcantes.frx":0000
      Height          =   300
      Left            =   1920
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   240
      Width           =   1632
      _ExtentX        =   2879
      _ExtentY        =   529
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
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lv 
      Height          =   1515
      Left            =   30
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   540
      Width           =   3525
      _ExtentX        =   6223
      _ExtentY        =   2667
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
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label LBLEXPORTACAO 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "EXP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   330
      Left            =   30
      TabIndex        =   8
      Top             =   2100
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[F4]-FILTRA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   1230
      TabIndex        =   12
      Top             =   2100
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CNTR"
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   1
      Left            =   1380
      TabIndex        =   11
      Top             =   240
      Width           =   555
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
      TabIndex        =   10
      Top             =   240
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
      TabIndex        =   9
      Top             =   0
      Width           =   3525
   End
End
Attribute VB_Name = "Marcantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False











Private Sub lvCarga_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub Carrega_Lv()

Dim i As Integer
cmdAdiciona.Enabled = False
cmdRemove.Enabled = False
Me.lv.ListItems.Clear

Dim itmx As ListItem
Dim Sql As String

Sql = "SELECT "
Sql = Sql & " LOTE,ITEM,QUANTIDADE,EMBALAGEM,ID_CONTEINER,MERCADORIA, AUTONUMCS , AUTONUMCNTR "
Sql = Sql & " FROM "
Sql = Sql & " SGIPA..VW_WMS_SEM_MARCANTE "
Sql = Sql & " WHERE 0=0 AND QUANTIDADE>0 "
If Me.TXTLOTE.text = "" And Me.dcConteiner.text = "" Then
    Sql = Sql & " AND FLAG_HISTORICO=0"
Else
    If Me.TXTLOTE.text <> "" Then
        Sql = Sql & " AND LOTE=" & Val(Me.TXTLOTE.text)
    End If
    If Me.dcConteiner.text <> "" Then
        Sql = Sql & " AND ID_CONTEINER='" & Me.dcConteiner.text & "'"
    End If

End If

Sql = Sql & " ORDER BY LOTE "

i = 1
Set tb1 = New ADODB.Recordset
tb1.Open Sql, db, 1, 1

If tb1.EOF Then

    tb1.Close
    Sql = "SELECT "
    Sql = Sql & " LOTE,ITEM,QUANTIDADE,EMBALAGEM,ID_CONTEINER,MERCADORIA, AUTONUM_CEXP AS AUTONUMCS , AUTONUMCNTR "
    Sql = Sql & " FROM "
    Sql = Sql & " SGIPA..VW_WMS_SEM_MARCANTE_CEXP "
    Sql = Sql & " WHERE 0=0 AND QUANTIDADE>0 "
    If Me.TXTLOTE.text = "" And Me.dcConteiner.text = "" Then
        Sql = Sql & " AND FLAG_HISTORICO=0"
    Else
        If Me.TXTLOTE.text <> "" Then
            Sql = Sql & " AND LOTE='" & Me.TXTLOTE.text & "'"
        End If
        If Me.dcConteiner.text <> "" Then
            Sql = Sql & " AND ID_CONTEINER='" & Me.dcConteiner.text & "'"
        End If
    
    End If
    
    Sql = Sql & " ORDER BY LOTE "
    
    Set tb1 = New ADODB.Recordset
    tb1.Open Sql, db, 1, 1
    If Not tb1.EOF Then
        Me.LBLEXPORTACAO.Visible = True
    End If

Else
    
    LBLEXPORTACAO.Visible = False
End If


While Not tb1.EOF
    
    Set itmx = lv.ListItems.Add(, , i)
    
    itmx.SubItems(1) = Nnull(tb1!Lote, 1)
    itmx.SubItems(2) = Nnull(tb1!Item, 1)
    itmx.SubItems(3) = Nnull(tb1!QUANTIDADE, 0)
    itmx.SubItems(4) = Nnull(tb1!EMBALAGEM, 1)
    itmx.SubItems(5) = Nnull(tb1!Id_Conteiner, 1)
    itmx.SubItems(6) = Nnull(tb1!MERCADORIA, 1)
    itmx.SubItems(7) = Nnull(tb1!AutonumCs, 0)
    itmx.SubItems(8) = Nnull(tb1!AutonumCntr, 0)
    i = i + 1
    cmdAdiciona.Enabled = True
    cmdRemove.Enabled = True
     tb1.MoveNext
Wend
tb1.Close



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
    If Me.txtMarcante.text = Left$(Me.lstMarcantes.List(i), 9) Then
        Col_Mensagem = "Marcante já adicionado"
        Col_Imagem = 1
        frmMensagem2.Show 1
        Exit Sub
    End If
Next i
 
Sql = "SELECT AUTONUM,AUTONUM_CARGA,DT_ASSOCIACAO,AUTONUM_CEXP FROM SGIPA..TB_MARCANTES WHERE "
Sql = Sql & " AUTONUM= " & Val(Me.txtMarcante) & ""

Set tb1 = New ADODB.Recordset
tb1.Open Sql, db, 1, 1
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

If Nnull(tb1!AUTONUM_CEXP, 0) > 0 Then
    If Me.LBLEXPORTACAO.Visible = False Then
        Col_Mensagem = "Nr de marcante de carga de exportacao"
        frmMensagem2.Show vbModal
        Exit Sub
    End If
    
    If Val(Nnull(Me.lv.ListItems(Val(Me.lv.SelectedItem)).SubItems(7), 0)) <> Val(Nnull(tb1!AUTONUM_CEXP, 0)) Then
        Col_Mensagem = "Marcante divergente ao lote exp."
        frmMensagem2.Show vbModal
        Exit Sub
    End If
End If


If Me.LBLEXPORTACAO.Visible = False Then
    Sql = "INSERT INTO sgipa..TB_TEMP_MARCANTE(AUTONUM_CS,MARCANTE,QTDE) VALUES ("
    Sql = Sql & Nnull(Me.lv.ListItems(Val(Me.lv.SelectedItem)).SubItems(7), 0) & ","
    Sql = Sql & "'" & Me.txtMarcante.text & "',"
    Sql = Sql & Val(Me.txtQtdeMarcante.text)
    Sql = Sql & ")"
    db.Execute Sql
Else
    Sql = "INSERT INTO sgipa..TB_TEMP_MARCANTE_CEXP(AUTONUM_CEXP,MARCANTE,QTDE) VALUES ("
    Sql = Sql & Nnull(Me.lv.ListItems(Val(Me.lv.SelectedItem)).SubItems(7), 0) & ","
    Sql = Sql & "'" & Me.txtMarcante.text & "',"
    Sql = Sql & Val(Me.txtQtdeMarcante.text)
    Sql = Sql & ")"
    db.Execute Sql
End If


Carrega_Marcantes_Temp

'Me.lstMarcantes.AddItem Me.txtMarcante.text & " " & Me.txtQtdeMarcante.text
Me.txtMarcante.text = ""
Me.txtQtdeMarcante.text = ""
Me.txtMarcante.SetFocus
End Sub

Private Sub cmdEscape_Click()
Unload Me
End Sub

Sub cmdFiltrar_Click()
Carrega_Lv
End Sub

Sub cmdRemove_Click()
If Me.LBLEXPORTACAO.Visible = False Then
    
    Sql = "DELETE FROM sgipa..TB_TEMP_MARCANTE "
    Sql = Sql & " WHERE AUTONUM_CS= " & Nnull(Me.lv.ListItems(Val(Me.lv.SelectedItem)).SubItems(7), 0)
    Sql = Sql & " AND MARCANTE= " & Val(Left(Me.lstMarcantes.text, 9))
    db.Execute Sql
    
Else

    Sql = "DELETE FROM sgipa..TB_TEMP_MARCANTE_CEXP "
    Sql = Sql & " WHERE AUTONUM_CEXP= " & Nnull(Me.lv.ListItems(Val(Me.lv.SelectedItem)).SubItems(7), 0)
    Sql = Sql & " AND MARCANTE= " & Val(Left(Me.lstMarcantes.text, 9))
    db.Execute Sql
End If
    
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


If Me.LBLEXPORTACAO.Visible = False Then
    QtdeCs = Nnull(db.Execute("SELECT QUANTIDADE_REAL FROM SGIPA..TB_CARGA_SOLTA WHERE AUTONUM=" & AutonumCs)(0), 0)
Else
    QtdeCs = Nnull(db.Execute("SELECT QUANTIDADE - ISNULL(QUANTIDADE_SAIDA,0) FROM SGIPA..TB_CARGA_EXPORTACAO WHERE AUTONUM=" & AutonumCs)(0), 0)
End If

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
        If Me.LBLEXPORTACAO.Visible = False Then
            tb1.Open "SELECT SUM(VOLUMES) AS QTOS FROM SGIPA..TB_MARCANTES WHERE AUTONUM_CARGA= " & AutonumCs, db, 1, 1
        Else
            tb1.Open "SELECT SUM(VOLUMES) AS QTOS FROM SGIPA..TB_MARCANTES WHERE AUTONUM_CEXP= " & AutonumCs, db, 1, 1
        End If
        If Not tb1.EOF Then
            QtdeJaMarcada = Nnull(tb1!qtos, 0)
        End If
        tb1.Close
        
        If QtdeJaMarcada + qteM <= QtdeCs Then
            If Me.LBLEXPORTACAO.Visible = False Then
                Sql = "UPDATE SGIPA..TB_MARCANTES SET DT_ASSOCIACAO=GETDATE(),AUTONUM_CARGA=" & AutonumCs
                Sql = Sql & " ,VOLUMES=" & qteM
                Sql = Sql & " ,BL=" & Nnull(Me.lv.ListItems(Val(Me.lv.SelectedItem)).SubItems(1), 0)
                Sql = Sql & " ,ITEM=" & Nnull(Me.lv.ListItems(Val(Me.lv.SelectedItem)).SubItems(2), 0)
                Sql = Sql & " ,CNTR=" & Nnull(Me.lv.ListItems(Val(Me.lv.SelectedItem)).SubItems(8), 0)
                Sql = Sql & " WHERE "
                Sql = Sql & " AUTONUM=" & AutonumMarcante
            Else
                Sql = "UPDATE SGIPA..TB_MARCANTES SET DT_ASSOCIACAO=GETDATE(),AUTONUM_CEXP=" & AutonumCs
                Sql = Sql & " ,VOLUMES=" & qteM
                Sql = Sql & " WHERE "
                Sql = Sql & " AUTONUM=" & AutonumMarcante
            End If
            db.Execute Sql
        End If
    Next i
    
    Col_Mensagem = "MARCANTES ASSOCIADOS COM SUCESSO"
    Col_Imagem = 2
    frmMensagem2.Show 1
    Me.lv.ListItems.Clear
    Me.lstMarcantes.Clear
    Me.TXTLOTE.SetFocus
    
    
End If

End Sub

Private Sub dcConteiner_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    dcConteiner.text = ""
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

Private Sub Carrega_Marcantes_Temp()
Dim Sql As String

On Error Resume Next

Me.lstMarcantes.Clear

If Me.LBLEXPORTACAO.Visible = False Then

    Sql = "SELECT MARCANTE,QTDE FROM SGIPA..TB_TEMP_MARCANTE "
    Sql = Sql & " WHERE AUTONUM_CS=" & Nnull(Me.lv.ListItems(Val(Me.lv.SelectedItem)).SubItems(7), 0)
    Sql = Sql & " ORDER BY MARCANTE"
Else

    Sql = "SELECT MARCANTE,QTDE FROM SGIPA..TB_TEMP_MARCANTE_CEXP "
    Sql = Sql & " WHERE AUTONUM_CEXP=" & Nnull(Me.lv.ListItems(Val(Me.lv.SelectedItem)).SubItems(7), 0)
    Sql = Sql & " ORDER BY MARCANTE"

End If
    
Set tbTemp = New ADODB.Recordset
tbTemp.Open Sql, db, 1, 1

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
colx.Width = 1350
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


'Sql = " select DISTINCT autonumCNTR, ID_CONTEINER  AS DISPLAY "
'Sql = Sql & " FROM SGIPA..VW_WMS_SEM_MARCANTE WHERE ID_CONTEINER IS NOT NULL "
'Sql = Sql & " ORDER BY ID_CONTEINER "
'
'Me.AdodcCntr.ConnectionString = ConnectionString
'Me.AdodcCntr.RecordSource = Sql
'Me.AdodcCntr.Refresh
'Set Me.AdodcCntr.Recordset.ActiveConnection = Nothing



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
TXTLOTE = UCase(TXTLOTE)

Sql = " select DISTINCT autonumCNTR, ID_CONTEINER  AS DISPLAY "
Sql = Sql & " FROM SGIPA..VW_WMS_SEM_MARCANTE WHERE ID_CONTEINER IS NOT NULL "
If TXTLOTE <> "" Then
    Sql = Sql & " and lote=" & TXTLOTE
End If
Sql = Sql & " ORDER BY ID_CONTEINER "

Screen.MousePointer = 11
Me.AdodcCntr.ConnectionString = ConnectionString
Me.AdodcCntr.RecordSource = Sql
Me.AdodcCntr.Refresh
Set Me.AdodcCntr.Recordset.ActiveConnection = Nothing
Screen.MousePointer = 0
dcConteiner.BoundText = ""


End Sub

Private Sub txtMarcante_GotFocus()
txtMarcante = ""
txtMarcante.BackColor = &HC0FFFF
End Sub

Private Sub txtMarcante_KeyPress(KeyAscii As Integer)
KeyAscii = SoNumeros09(KeyAscii)
End Sub

Private Sub txtMarcante_LostFocus()
txtMarcante.BackColor = &HFFFFFF
End Sub

