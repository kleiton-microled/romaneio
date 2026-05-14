VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FrmDescargaArmazem 
   BackColor       =   &H00BA9F8D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descarga Armazem"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   3345
   StartUpPosition =   3  'Windows Default
   Begin MSDataListLib.DataCombo cbLote 
      Bindings        =   "FrmDescargaArmazem.frx":0000
      Height          =   315
      Left            =   630
      TabIndex        =   1
      Top             =   585
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "lote"
      BoundColumn     =   "lote"
      Text            =   "DataCombo1"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "F2 - Limpar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1785
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Novo"
      Top             =   975
      Width           =   1515
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00BA9F8D&
      Enabled         =   0   'False
      Height          =   3075
      Left            =   15
      TabIndex        =   18
      Top             =   1305
      Width           =   3300
      Begin MSDataGridLib.DataGrid Grid1 
         Bindings        =   "FrmDescargaArmazem.frx":0015
         Height          =   1140
         Left            =   90
         TabIndex        =   23
         Top             =   1890
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   2011
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "quantidade"
            Caption         =   "Qtde"
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
         BeginProperty Column01 
            DataField       =   "descr"
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
         BeginProperty Column02 
            DataField       =   "finalizado"
            Caption         =   "finalizado"
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
            DataField       =   "autonum"
            Caption         =   "Autonum"
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
               ColumnWidth     =   629,858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2684,977
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F3 - Novo"
         Height          =   360
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Limpar"
         Top             =   1500
         Width           =   1035
      End
      Begin VB.CommandButton Command4 
         Caption         =   "F6 - Excluir"
         Height          =   345
         Left            =   2175
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Remover"
         Top             =   1500
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00BA9F8D&
         Caption         =   "SIM"
         Height          =   240
         Left            =   1710
         TabIndex        =   6
         Top             =   1050
         Width           =   720
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00BA9F8D&
         Caption         =   "NÃO"
         Height          =   240
         Left            =   2505
         TabIndex        =   7
         Top             =   1050
         Width           =   720
      End
      Begin VB.CommandButton Command3 
         Caption         =   "F5 - Salvar"
         Height          =   360
         Left            =   1125
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Gravar"
         Top             =   1500
         Width           =   1035
      End
      Begin VB.TextBox TxtQuantidade 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1035
         MaxLength       =   5
         TabIndex        =   4
         Top             =   165
         Width           =   900
      End
      Begin MSDataListLib.DataCombo DcEmbalagem 
         Bindings        =   "FrmDescargaArmazem.frx":002A
         Height          =   315
         Left            =   1005
         TabIndex        =   5
         Top             =   525
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "descr"
         BoundColumn     =   "code"
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entrega Finalizada ?"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1065
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
         Height          =   195
         Left            =   135
         TabIndex        =   20
         Top             =   195
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Embalagem"
         Height          =   195
         Left            =   60
         TabIndex        =   19
         Top             =   495
         Width           =   825
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "F4 - Carregar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Pesquisar"
      Top             =   975
      Width           =   1530
   End
   Begin VB.TextBox TxtGate 
      Height          =   285
      Left            =   8565
      TabIndex        =   17
      Text            =   "0"
      Top             =   1815
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox TxtOC 
      Height          =   285
      Left            =   8640
      TabIndex        =   15
      Text            =   "0"
      Top             =   990
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox TxtLote 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3615
      TabIndex        =   11
      Top             =   2460
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox TxtAutonum 
      Height          =   285
      Left            =   3540
      TabIndex        =   13
      Text            =   "0"
      Top             =   2970
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TxtCntr 
      Height          =   285
      Left            =   8595
      TabIndex        =   12
      Text            =   "0"
      Top             =   1395
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3465
      Top             =   1905
      Visible         =   0   'False
      Width           =   2385
      _ExtentX        =   4207
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3480
      Top             =   3960
      Visible         =   0   'False
      Width           =   2370
      _ExtentX        =   4180
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
         Name            =   "Arial"
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
      Left            =   3495
      Top             =   720
      Visible         =   0   'False
      Width           =   2370
      _ExtentX        =   4180
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cbPlaca 
      Bindings        =   "FrmDescargaArmazem.frx":003F
      Height          =   315
      Left            =   615
      TabIndex        =   0
      Top             =   255
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "placa"
      BoundColumn     =   "autonum_rs"
      Text            =   "DataCombo1"
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   3480
      Top             =   240
      Visible         =   0   'False
      Width           =   2370
      _ExtentX        =   4180
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
      Caption         =   "Adodc4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "CARREGAMENTO DDC"
      Height          =   210
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   3345
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Placa"
      Height          =   195
      Left            =   90
      TabIndex        =   16
      Top             =   285
      Width           =   405
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lote"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   660
      Width           =   315
   End
End
Attribute VB_Name = "FrmDescargaArmazem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Carrega_Grid1()
    If Val(TxtGate.text) <> 0 And Val(TxtCntr) <> 0 Then
      SQl = "select a.quantidade, c.descr, A.FINALIZADO, a.autonum, A.EMBALAGEM "
      SQl = SQl & " from " & Banco_Sgipa & "tb_descarga_armazem a"
      SQl = SQl & " left join " & Banco_Operador & "tb_amr_gate b on a.autonum = b.id_descarga_armazem"
      SQl = SQl & " inner join " & Banco_Sgipa & "dte_tb_embalagens c on a.embalagem = c.code"
      SQl = SQl & " left join " & Banco_Operador & "tb_gate_new d on b.gate = d.autonum"
      SQl = SQl & " where "
      SQl = SQl & " a.cntr=" & TxtCntr
      SQl = SQl & " and a.gate=" & TxtGate
    Else
      SQl = "select null quantidade, null descr, null FINALIZADO, null autonum, null EMBALAGEM from dual"
    End If
    Adodc1.RecordSource = CONVERTE_SQL(SQl)
    Adodc1.ConnectionString = ConnectionString
    Adodc1.Refresh
End Sub


Sub Carrega_Placa()
    Dim SQl As String
    
    SQl = "select distinct a.autonum as autonum_rs, b.placa_c + ' (' + cc.id_conteiner + ')' as placa"
    SQl = SQl & " from sgipa..tb_registro_saida_cntr a"
    SQl = SQl & " inner join sgipa..tb_cntr_bl cc on a.cntr=cc.autonum"
    SQl = SQl & " inner join sgipa..tb_ordem_carregamento b on a.ordem_carreg = b.autonum"
    SQl = SQl & " where b.flag_saida=0 and b.flag_ddc=1"
    Adodc4.ConnectionString = ConnectionString
    Adodc4.RecordSource = SQl
    Adodc4.Refresh

End Sub


Sub CARREGAR()
    If Me.cbPlaca.BoundText = "" Then
        MsgBox "Placa não Informada"
        Exit Sub
    End If
    
    If Me.cbLote.BoundText = "" Then
        MsgBox "Lote não Informado"
        Exit Sub
    End If
    TxtLote = cbLote.BoundText
    
    Option1.Value = False
    Option2.Value = False
    
    SQl = "select a.* from " & Banco_Sgipa & "tb_registro_saida_cntr a where a.autonum = " & Me.cbPlaca.BoundText
    Dim Rs As New ADODB.Recordset
    PRSet Rs, SQl
    TxtOC = Rs!ordem_carreg
    TxtCntr = Rs!CNTR
    Frame1.Enabled = True
    Call Carrega_Grid1

End Sub

Sub Excluir()
        If MsgBox("Confirma exclusão do item ?", vbYesNo) = vbNo Then
            Exit Sub
        End If
        
        SQl = "select sum(a.quantidade_saida) from " & Banco_Operador & "tb_amr_gate a"
        SQl = SQl & " where id_descarga_armazem=" & TxtAutonum
        PRSet Rs, SQl
        If Rs.Fields(0) <> 0 Then
            MsgBox "Já consta saida para este item - Exclusão Cancelada"
            Rs.Close
            Exit Sub
        End If
        Rs.Close
        
        
        SQl = "delete from " & Banco_Sgipa & "tb_descarga_armazem where autonum = " & TxtAutonum
        Executa SQl
        
        TxtAutonum = 0
        TxtQuantidade = ""
        DcEmbalagem.BoundText = ""
        Call Carrega_Grid1

End Sub

Sub Limpar()
    cbPlaca.BoundText = ""
    cbLote.BoundText = ""
    Frame1.Enabled = False
    TxtAutonum = 0
    TxtCntr = "0"
    TxtGate = "0"
    TxtQuantidade = ""
    TxtLote.text = "0"
    TxtOC.text = "0"
    DcEmbalagem.BoundText = ""
    Option1.Value = False
    Option2.Value = False
    
End Sub

Private Sub MaskEdBox1_Change()

End Sub




Private Sub cbConteiner_LostFocus()


    If cbConteiner.BoundText <> "" Then
        Dim SQl As String
        
        
        SQl = "select distinct b.bl lote"
        SQl = SQl & " from Sgipa..tb_amr_cntr_bl b"
        SQl = SQl & " inner join sgipa..tb_bl a on b.bl = a.autonum"
        SQl = SQl & " where b.cntr=" & cbConteiner.BoundText & " and a.flag_ativo=1"
        Adodc3.ConnectionString = ConnectionString
        Adodc3.RecordSource = SQl
        Adodc3.Refresh
    End If



End Sub


Sub Novo()
        If Me.cbPlaca.BoundText = "" Then
            MsgBox "Placa não informada"
            Exit Sub
        End If
        If Me.cbLote.BoundText = "" Then
            MsgBox "Lote não informada"
            Exit Sub
        End If

        TxtAutonum = 0
        TxtQuantidade = ""
        DcEmbalagem.BoundText = ""
        Option1.Value = False
        Option2.Value = False
        TxtQuantidade.SetFocus

End Sub

Sub Salvar()

        If Option1.Value = False And Option2.Value = False Then
            MsgBox "Informe finalização de entrega"
            Exit Sub
        End If
        If Me.cbPlaca.BoundText = "" Then
            MsgBox "Placa não informada"
            Exit Sub
        End If
        If TxtQuantidade = "" Then TxtQuantidade = "0"
        If Val(TxtQuantidade) = 0 Then
            MsgBox "Quantidade Invalida"
            TxtQuantidade = ""
            TxtQuantidade.SetFocus
            Exit Sub
        End If
        If DcEmbalagem.BoundText = "" Then
            MsgBox "Embalagem não Informada"
            DcEmbalagem.SetFocus
            Exit Sub
        End If
        If Val(TxtLote) = 0 Then
            MsgBox "Lote não Informada"
            Exit Sub
        End If
        If Val(TxtGate) = 0 Then
            MsgBox "Gate não localizado"
            Exit Sub
        End If
        
        If TxtAutonum = "0" Then
            SQl = "insert into " & Banco_Sgipa & "tb_descarga_armazem ("
            If Banco_Em_Uso = "ORACLE" Then SQl = SQl & "autonum,"
            SQl = SQl & "cntr, quantidade, embalagem,  usuario, lote,gate,id_oc,finalizado"
            SQl = SQl & ") values ("
            If Banco_Em_Uso = "ORACLE" Then SQl = SQl & "" & Banco_Sgipa & "seq_tb_descarga_armazem.nextval,"
            SQl = SQl & "" & TxtCntr
            SQl = SQl & "," & TxtQuantidade
            SQl = SQl & ",'" & DcEmbalagem.BoundText & "'"
            SQl = SQl & "," & Cod_Usuario
            SQl = SQl & "," & TxtLote
            SQl = SQl & "," & TxtGate
            SQl = SQl & "," & TxtOC
            If Option1.Value = True Then
                SQl = SQl & ",1"
            Else
                SQl = SQl & ",0"
            End If
            SQl = SQl & ")"
        Else
            
            SQl = "select sum(a.quantidade_saida) from " & Banco_Operador & "tb_amr_gate a"
            SQl = SQl & " where id_descarga_armazem=" & TxtAutonum
            PRSet Rs, SQl
            
            SQl = "update " & Banco_Sgipa & "tb_descarga_armazem set"
            SQl = SQl & " quantidade = " & TxtQuantidade
            SQl = SQl & " ,finalizado = "
            If Option1.Value = True Then
                SQl = SQl & "1 "
            Else
                SQl = SQl & "0 "
            End If
            
            If Rs.Fields(0) <> 0 Then
                If Val(TxtQuantidade) < Rs.Fields(0) Then
                    MsgBox "Quantidade não pode ser inferior a quantidade já saida"
                    Rs.Close
                    Exit Sub
                End If
                If Rs.Fields(0) <> 0 Then
                    If DcEmbalagem.BoundText <> DcEmbalagem.Tag Then
                        MsgBox "Atenção - A embalagem não pode ser alterada pois ja existe saida para este item"
                    End If
                Else
                    SQl = SQl & ",embalagem = '" & DcEmbalagem.BoundText & "'"
                End If
            Else
                SQl = SQl & ",embalagem = '" & DcEmbalagem.BoundText & "'"
            End If
            Rs.Close
            
            SQl = SQl & " where autonum = " & TxtAutonum
        End If
        Executa SQl
        
        
        SQl = "select nvl(flag_ddc,0) flag_ddc from " & Banco_Sgipa & "tb_cntr_bl where autonum=" & TxtCntr
        PRSet Rst, SQl
        If Rst!Flag_DDC = 0 Then
            SQl = "update " & Banco_Sgipa & "tb_cntr_bl set flag_ddc=1,inicio_ddc=sysdate where autonum=" & TxtCntr
            Executa SQl
        End If
        If Option1.Value = True Then
            SQl = "update " & Banco_Sgipa & "tb_cntr_bl set flag_ddc=2 where autonum=" & TxtCntr
            Executa SQl
        End If
        Rst.Close
        
        
        Call Carrega_Grid1
        
        TxtAutonum = 0
        TxtQuantidade = ""
        DcEmbalagem.BoundText = ""
        txtMarca = ""
        TxtQuantidade.SetFocus

End Sub

Private Sub cbPlaca_LostFocus()

    cbLote.BoundText = ""
    Frame1.Enabled = False
    TxtAutonum = 0
    TxtCntr = "0"
    TxtGate = "0"
    TxtQuantidade = ""
    TxtLote.text = "0"
    TxtOC.text = "0"
    DcEmbalagem.BoundText = ""
    Option1.Value = False
    Option2.Value = False
    Carrega_Grid1
    

    Dim SQl As String
    SQl = "select 0 as lote from dual"
    Adodc3.ConnectionString = ConnectionString
    Adodc3.RecordSource = SQl
    Adodc3.Refresh
    TxtGate = 0
    Me.cbLote.BoundText = ""
    
    If cbPlaca.BoundText <> "" Then
        Dim Rs As New ADODB.Recordset
        SQl = "select gn.autonum"
        SQl = SQl & " from sgipa..tb_registro_saida_cntr a"
        SQl = SQl & " inner join sgipa..tb_ordem_carregamento oc on a.ordem_carreg = oc.autonum"
        SQl = SQl & " inner join operador..tb_gate_new gn on oc.placa_c = gn.placa"
        SQl = SQl & " where a.autonum=" & cbPlaca.BoundText & " and gn.flag_gate_in=1 and isnull(gn.flag_gate_out,0)=0"
        PRSet Rs, SQl
        If Rs.EOF Then
            MsgBox "Não consta entrada para o veiculo informado"
            Exit Sub
        End If
        TxtGate = Rs!autonum
        Rs.Close
        
        SQl = "select distinct bl.autonum as lote"
        SQl = SQl & " from sgipa..tb_registro_saida_cntr a"
        SQl = SQl & " inner join Sgipa..tb_amr_cntr_bl c on a.cntr = c.cntr"
        SQl = SQl & " inner join sgipa..tb_bl bl on c.bl = bl.autonum"
        SQl = SQl & " where a.autonum=" & cbPlaca.BoundText
        SQl = SQl & " and bl.flag_ativo=1"
        Adodc3.ConnectionString = ConnectionString
        Adodc3.RecordSource = SQl
        Adodc3.Refresh
    End If
    TxtCntr = "0"
    TxtOC.text = "0"
    TxtLote.text = "0"
End Sub


Private Sub Command1_Click()
    CARREGAR
End Sub



Private Sub Command2_Click()
    Limpar
End Sub

Private Sub Command3_Click()
    Salvar
End Sub

Private Sub Command4_Click()
    Excluir
End Sub

Private Sub Command5_Click()
    Novo
End Sub



Private Sub Command7_Click()
Unload Me
End Sub



Private Sub DataGrid1_Click()


    If Not Adodc1.Recordset.EOF Then
        Me.TxtAutonum = Adodc1.Recordset!autonum
        Me.TxtQuantidade.text = Adodc1.Recordset!quantidade
        Me.DcEmbalagem.BoundText = Adodc1.Recordset!embalagem
        
        If Nnull(Adodc1.Recordset!finaliZado, 0) = 0 Then
            Option2.Value = True
        Else
            Option1.Value = True
        End If
    End If



End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyEscape Then
    Unload Me
End If
Select Case KeyCode
Case vbKeyF4
    CARREGAR
Case vbKeyF2
    Limpar
Case vbKeyF3
    If Not Me.Frame1.Enabled Then
        MsgBox "Não Dispinivel"
        Exit Sub
    End If
    Novo
Case vbKeyF5
    If Not Me.Frame1.Enabled Then
        MsgBox "Não Dispinivel"
        Exit Sub
    End If
    Salvar
Case vbKeyF6
    If Not Me.Frame1.Enabled Then
        MsgBox "Não Dispinivel"
        Exit Sub
    End If
    Excluir
End Select


End Sub

Private Sub Form_Load()


Carrega_Placa




SQl = "SELECT code, descr from " & Banco_Sgipa & "dte_tb_embalagens order by descr"
Adodc2.RecordSource = CONVERTE_SQL(SQl)
Adodc2.ConnectionString = ConnectionString
Adodc2.Refresh

TxtCntr = ""

End Sub


Private Sub Form_Unload(Cancel As Integer)

    Unload Me

End Sub


Private Sub Grid1_Click()
    TxtAutonum = "0"
    If Not Adodc1.Recordset.EOF Then
        Dim Rs As New ADODB.Recordset
        TxtAutonum = Adodc1.Recordset!autonum
        TxtQuantidade = Adodc1.Recordset!quantidade
        DcEmbalagem.BoundText = Adodc1.Recordset!embalagem
        DcEmbalagem.Tag = Adodc1.Recordset!embalagem
        If Nnull(Adodc1.Recordset!finaliZado, 0) = 0 Then
            Me.Option2.Value = 1
        Else
            Me.Option1.Value = 1
        End If
        TxtQuantidade.SetFocus
    End If

End Sub

Private Sub MskPlaca_LostFocus()
    
    TxtGate = 0
    If mskPlaca.ClipText <> "" Then
        SQl = "select autonum from " & Banco_Operador & "tb_gate_new where placa='" & mskPlaca & "' and flag_gate_in=1 and flag_gate_out=0"
        Dim Rs As New ADODB.Recordset
        PRSet Rs, SQl
        If Rs.EOF Then
            MsgBox "Não consta entrada para o veiculo informado"
            mskPlaca = "___-____"
            mskPlaca.SetFocus
            Exit Sub
        End If
        TxtGate = Rs!autonum
        Rs.Close
    End If
    
End Sub




