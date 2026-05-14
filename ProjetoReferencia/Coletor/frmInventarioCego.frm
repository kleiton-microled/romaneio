VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frmInventarioCego 
   Appearance      =   0  'Flat
   BackColor       =   &H00BA9F8D&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3480
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
   ScaleHeight     =   3840
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "[F2]EXCLUIR"
      Height          =   255
      Left            =   1290
      TabIndex        =   13
      Top             =   3570
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "[F5]SALVAR"
      Height          =   255
      Left            =   30
      TabIndex        =   12
      Top             =   3570
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdEscape 
      Caption         =   "[ESC]- SAIR"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   3570
      Width           =   1035
   End
   Begin VB.CommandButton cmdSelecionar 
      Caption         =   "[F5]-SELECIONA"
      Height          =   285
      Left            =   30
      TabIndex        =   10
      Top             =   3570
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00BA9F8D&
      Caption         =   "INVENTÁRIOS EM ABERTO"
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   0
      TabIndex        =   1
      Top             =   30
      Width           =   3435
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   2160
         Top             =   3120
         Visible         =   0   'False
         Width           =   1410
         _ExtentX        =   2487
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
         Bindings        =   "frmInventarioCego.frx":0000
         Height          =   3225
         Left            =   90
         TabIndex        =   2
         Top             =   240
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   5689
         _Version        =   393216
         Cols            =   6
         RowHeightMin    =   360
         BackColorFixed  =   12230541
         ForeColorFixed  =   16777215
         WordWrap        =   -1  'True
         FormatString    =   " |<Descricao   |^Armazem    |^Prat.|^Heap     |"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
   End
   Begin VB.TextBox txtCritica 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3390
      Width           =   3105
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Inventario:"
      ForeColor       =   &H80000008&
      Height          =   3405
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   2550
         Top             =   2970
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
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
      Begin VB.TextBox txtPrateleira 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   12
         TabIndex        =   6
         Top             =   600
         Width           =   2235
      End
      Begin VB.TextBox txtMarcante 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   4
         Top             =   210
         Width           =   2235
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
         Bindings        =   "frmInventarioCego.frx":0015
         Height          =   2355
         Left            =   60
         TabIndex        =   8
         Top             =   960
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   4154
         _Version        =   393216
         Cols            =   4
         RowHeightMin    =   330
         BackColorFixed  =   12230541
         ForeColorFixed  =   16777215
         SelectionMode   =   1
         FormatString    =   "     |<Marcante                     |<Local                 |"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00BA9F8D&
         Caption         =   "LOCAL"
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
         Height          =   375
         Index           =   16
         Left            =   60
         TabIndex        =   7
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00BA9F8D&
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
         Left            =   60
         TabIndex        =   5
         Top             =   210
         Width           =   1035
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[F5]SELECIONA    [ESC]-SAI "
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   3870
      Width           =   2115
   End
End
Attribute VB_Name = "frmInventarioCego"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public AutonumArm As Integer
Public YardCs As String


Private Sub cmdEscape_Click()
Unload Me
End Sub

Sub cmdExcluir_Click()
Sql = "DELETE FROM SGIPA..TB_INVENTARIO_ITEM WHERE AUTONUM_INV=" & Col_IdInventario & " AND MARCANTE=" & Val(Me.txtMarcante.text)
db.Execute Sql
Carrega_Grid_Carga
End Sub

Sub cmdSalvar_Click()
If Valida_Insert Then
        'Passou pelas criticas - salva
        
    Sql = "Select autonum from sgipa..tb_inventario_item where autonum_inv=" & Col_IdInventario & " and marcante=" & Val(Me.txtMarcante.text)
    Set tb1 = New ADODB.Recordset
    tb1.Open Sql, db, 1, 1
    If Not tb1.EOF Then
        Sql = "UPDATE SGIPA..TB_INVENTARIO_ITEM SET "
        Sql = Sql & " YARD='" & YardCs & "',"
        Sql = Sql & " DT_INVENT=GETDATE()"
        Sql = Sql & " WHERE AUTONUM=" & tb1!AUTONUM
    Else
        Sql = "INSERT INTO SGIPA..TB_INVENTARIO_ITEM ("
        Sql = Sql & "AUTONUM_INV,"
        Sql = Sql & "MARCANTE,"
        Sql = Sql & "YARD,"
        Sql = Sql & "DT_INVENT) VALUES ("
        Sql = Sql & Col_IdInventario & ","
        Sql = Sql & Val(Me.txtMarcante.text) & ","
        Sql = Sql & "'" & YardCs & "',"
        Sql = Sql & "GETDATE()"
        Sql = Sql & ")"
        
    End If
    db.Execute Sql
    Carrega_Grid_Carga
    Me.txtCritica.text = "SALVO"
    Me.txtCritica.ForeColor = vbGreen
    Me.txtMarcante.text = ""
    Me.txtPrateleira.text = ""
    
End If
End Sub

Sub cmdSelecionar_Click()
If Me.Grid1.Row > 0 Then
    Col_IdInventario = Me.Grid1.TextMatrix(Me.Grid1.Row, Me.Grid1.Cols - 1)
    Col_DescrInventario = Me.Grid1.TextMatrix(Me.Grid1.Row, 1)
    Me.Frame1.Visible = False
    Me.Frame2.Visible = True
    'Me.Label7.Caption = "[F5]SALVA [F2]-EXCLUIR [ESC]-SAI "
    Me.Frame2.Caption = "Inventario " & Col_DescrInventario
    Me.cmdSelecionar.Visible = False
    Me.cmdSalvar.Visible = True
    Me.cmdExcluir.Visible = True
    Carrega_Grid_Carga
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)



If KeyCode = vbKeyEscape And Frame1.Visible = True Then
    Unload Me
End If

If KeyCode = vbKeyEscape And Frame1.Visible = False Then
    Me.Frame1.Visible = True
    Me.Frame2.Visible = False
    Me.cmdSelecionar.Visible = True
    Me.cmdSalvar.Visible = False
    Me.cmdExcluir.Visible = False
    
    Me.Label7.Caption = "[F5]SELECIONA [ESC]-SAI "
End If


If KeyCode = vbKeyF5 And Me.Frame1.Visible = True Then
    If Me.Grid1.Row >= 1 Then
        Call cmdSelecionar_Click
        
    End If
End If

If KeyCode = vbKeyF2 And Me.Frame2.Visible = True And Me.txtMarcante.text <> "" Then
    Call cmdExcluir_Click
    
End If


If KeyCode = vbKeyF5 And Me.Frame2.Visible = True Then
    Call cmdSalvar_Click
    
End If


End Sub

Private Sub Carrega_Grid_Carga()
Me.Grid2.Clear

Sql = " SELECT RIGHT('000000000' + LTRIM(STR(MARCANTE)),9),YARD FROM SGIPA..TB_INVENTARIO_ITEM WHERE AUTONUM_INV= " & Col_IdInventario & " ORDER BY AUTONUM DESC"

Me.Adodc2.ConnectionString = ConnectionString
Me.Adodc2.RecordSource = Sql
Me.Adodc2.Refresh
Set Me.Adodc2.Recordset.ActiveConnection = Nothing

Me.Grid2.FormatString = "     |<Marcante                 |<Local             |"
Me.Grid2.COLWIDTH(Me.Grid2.Cols - 1) = 0

Me.txtMarcante.SetFocus

End Sub

Private Function Valida_Insert() As Boolean
Valida_Insert = True
AutonumArm = 0
YardCs = ""
Dim FlagCt As Boolean

Me.txtCritica.text = ""
Me.txtCritica.ForeColor = vbGreen

If Me.txtMarcante.text = "" Then
    Valida_Insert = False
    Me.txtCritica.text = "Informe o marcante"
    Me.txtCritica.ForeColor = vbRed
    Exit Function
End If
If Me.txtPrateleira.text = "" Then
    Valida_Insert = False
    Me.txtCritica.text = "Informe a local"
    Me.txtCritica.ForeColor = vbRed
    Exit Function
End If

If Len(Me.txtMarcante) <> 9 Then
    Valida_Insert = False
    Me.txtCritica.text = "Marcante em formato incorreto"
    Me.txtCritica.ForeColor = vbRed
    Exit Function
End If


FlagCt = False

If Len(Me.txtPrateleira.text) = 9 Then

    Set TBP = New ADODB.Recordset
    
    AutonumArm = Val(Left(Me.txtPrateleira.text, 2))
    YardCs = UCase(Replace(Mid$(Me.txtPrateleira.text, 3), "-", ""))
    
    '01L011
    '01COFRE
    
    Dim xArm As Integer
    Dim Xpos As String
    xArm = AutonumArm
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

ElseIf Len(Me.txtPrateleira.text) >= 4 And Len(Me.txtPrateleira.text) < 12 And InStr(1, "-", Me.txtPrateleira.text) = 0 Then

    AutonumArm = Val(Left(Me.txtPrateleira.text, 2))
    YardCs = UCase(Replace(Mid$(Me.txtPrateleira.text, 3), "-", ""))
    
    xArm = AutonumArm
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
    tb1.Close
    

ElseIf Len(Me.txtPrateleira.text) = 12 And InStr(1, Me.txtPrateleira.text, "-") = 11 Then

    
    Set tbCT = New ADODB.Recordset
    tbCT.Open "SELECT autonum FROM SGIPA..tb_armazens_ipa where descr='" & Me.txtPrateleira.text & "' and flag_ct=1 and flag_historico=0", db, 1, 1
    If Not tbCT.EOF Then
        AutonumArm = Val(tbCT!AUTONUM)
    End If
    YardCs = ""
    FlagCt = True
Else

    AutonumArm = 0
    YardCs = ""

End If


If FlagCt = False And YardCs = "" Then
    Valida_Insert = False
    Me.txtCritica.text = "Local inválido"
    Me.txtCritica.ForeColor = vbRed
    Exit Function
End If

If AutonumArm = 0 Then
    Valida_Insert = False
    Me.txtCritica.text = "Local inválido"
    Me.txtCritica.ForeColor = vbRed
    Exit Function

Else
    Set tbinv = New ADODB.Recordset
    tbinv.Open "Select id_armazem,prateleira,heap from sgipa..vw_invent_aberto where autonum=" & Col_IdInventario, db, 1, 1
    If tbinv.EOF Then
        Valida_Insert = False
        Me.txtCritica.text = "Inventário indisponivel"
        Me.txtCritica.ForeColor = vbRed
        Exit Function
    Else
        If tbinv!id_armazem <> AutonumArm Then
            Valida_Insert = False
            Me.txtCritica.text = "Armazem divergente do inventario"
            Me.txtCritica.ForeColor = vbRed
            Exit Function
        End If
        If Nnull(tbinv!prateleira, 1) <> "" Then
            If Left(YardCs, 1) <> Nnull(tbinv!prateleira, 1) Then
                Valida_Insert = False
                Me.txtCritica.text = "Prateleira divergente do inventario"
                Me.txtCritica.ForeColor = vbRed
                Exit Function
            End If
            
        End If
        If Nnull(tbinv!heap, 1) <> "" Then
            If Left(YardCs, 1) <> Nnull(tbinv!heap, 1) Then
                Valida_Insert = False
                Me.txtCritica.text = "Local divergente do inventario"
                Me.txtCritica.ForeColor = vbRed
                Exit Function
            End If
        End If
    End If
        
End If




End Function


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Sendkeys "{TAB}"
End If
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0

Carrega_Inventarios
End Sub

Private Sub Carrega_Inventarios()
Me.Grid1.Clear

Dim Sql As String
Sql = "Select DESCR,DESCR_ARMAZEM,PRATELEIRA,HEAP,AUTONUM FROM SGIPA..VW_INVENT_ABERTO ORDER BY DESCR"

Me.Adodc1.ConnectionString = ConnectionString
Me.Adodc1.RecordSource = Sql
Me.Adodc1.Refresh
Set Me.Adodc1.Recordset.ActiveConnection = Nothing


Me.Grid1.FormatString = "      |<Descricao   |^Armazem    |^Prat.|^Heap     |"
Me.Grid1.COLWIDTH(Me.Grid1.Cols - 1) = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

