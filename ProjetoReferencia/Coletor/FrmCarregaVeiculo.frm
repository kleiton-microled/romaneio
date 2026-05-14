VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FrmCarregaVeiculo 
   Caption         =   "Coletor- Veiculos"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   3585
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3735
      Top             =   3345
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3690
      Top             =   1155
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4980
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   8784
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "FrmCarregaVeiculo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Grid_Carregamento"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DC_Veiculos"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Grid_OC"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.CommandButton Command2 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   45
         TabIndex        =   8
         Top             =   4590
         Width           =   1710
      End
      Begin VB.CommandButton Command1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1875
         TabIndex        =   7
         Top             =   4575
         Width           =   1650
      End
      Begin MSDataGridLib.DataGrid Grid_OC 
         Bindings        =   "FrmCarregaVeiculo.frx":001C
         Height          =   1110
         Left            =   15
         TabIndex        =   4
         Top             =   720
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   1958
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "autonum"
            Caption         =   "OC"
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
         BeginProperty Column02 
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   945,071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1184,882
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   659,906
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo DC_Veiculos 
         Height          =   315
         Left            =   945
         TabIndex        =   2
         Top             =   90
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataGridLib.DataGrid Grid_Carregamento 
         Bindings        =   "FrmCarregaVeiculo.frx":0031
         Height          =   2310
         Left            =   15
         TabIndex        =   6
         Top             =   2160
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   4075
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "chassiS"
            Caption         =   "Chassi"
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
            DataField       =   "dt_carregamento"
            Caption         =   "Carregamento"
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
               ColumnWidth     =   1980,284
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Caption         =   "Carregamentos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   15
         TabIndex        =   5
         Top             =   1980
         Width           =   3495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Caption         =   "Ordens de Carregamento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   30
         TabIndex        =   3
         Top             =   450
         Width           =   3495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Veículo :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   150
         Width           =   690
      End
   End
End
Attribute VB_Name = "FrmCarregaVeiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    If Not Adodc2.Recordset.EOF Then
        If Nnull(Adodc2.Recordset("autonum"), 0) <> 0 Then
            If MsgBox("Remover Carregamento ?", vbYesNo) = vbNo Then Exit Sub
            Sql = "update SGIPA.DBO.TB_CARGA_SOLTA_CHASSIS set dt_carregamento=null"
            Sql = Sql & " where autonum=" & Adodc2.Recordset!autonum
            db.Execute Sql
        End If
        Adodc2.Refresh
    End If

    
    


End Sub

Private Sub Command2_Click()
    If Adodc2.Recordset.EOF Then
        MsgBox "Relacao de Chassis nao carregado"
        Exit Sub
    
    End If
    
    If Not IsNull(Adodc2.Recordset!dt_carregamento) Then
        MsgBox "Chassi ja Carregado"
        Exit Sub
    End If

    Sql = "update sgipa.dbo.tb_carga_solta_chassis set dt_carregamento=GETDATE() where autonum=" & Adodc2.Recordset!autonum
    db.Execute (Sql)
    Adodc2.Refresh
    


End Sub


Private Sub DC_Veiculos_Click(Area As Integer)


Sql = "select oc.autonum,bl.autonum as lote, sum(scs.quantidade) as quantidade"
Sql = Sql & " from sgipa.dbo.tb_ordem_carregamento oc"
Sql = Sql & " inner join sgipa.dbo.tb_registro_saida_cs scs on oc.autonum=scs.ordem_carreg"
Sql = Sql & " inner join sgipa.dbo.tb_carga_solta cs on scs.cs = cs.autonum"
Sql = Sql & " inner join sgipa.dbo.tb_bl bl on cs.bl=bl.autonum"
Sql = Sql & " where oc.autonum="
If DC_Veiculos.BoundText <> "" Then
    Sql = Sql & DC_Veiculos.BoundText
Else
    Sql = Sql & "NULL"
End If
Sql = Sql & " GROUP BY oc.autonum,bl.autonum"

Adodc1.ConnectionString = ConnectionString
Adodc1.RecordSource = Sql
Adodc1.Refresh



End Sub

Private Sub Form_Load()

Sql = "select distinct oc.autonum, oc.placa_c + ' / ' + oc.placa_carreta as placa"
Sql = Sql & " from sgipa.dbo.tb_ordem_carregamento oc"
Sql = Sql & " inner join sgipa.dbo.tb_registro_saida_cs scs on oc.autonum=scs.ordem_carreg"
Sql = Sql & " inner join sgipa.dbo.tb_carga_solta cs on scs.cs = cs.autonum"
Sql = Sql & " inner join sgipa.dbo.dte_tb_embalagens e on cs.embalagem=e.code"
Sql = Sql & " inner join sgipa.dbo.tb_bl bl on cs.bl=bl.autonum"
If Col_Patio <> 1 And Col_Patio <> 7 Then
    Sql = Sql & " where bl.patio=" & Col_Patio
Else
    Sql = Sql & " where bl.patio in (1,7)"
End If
Sql = Sql & " and e.descr='VEICULO'"
Sql = Sql & " and isnull(oc.flag_saida,0)=0"
Sql = Sql & " and oc.flag_entrada_camv=1"
Dim RsV As New ADODB.Recordset
PRSet RsV, Sql
Set Me.DC_Veiculos.RowSource = RsV
Me.DC_Veiculos.ListField = "placa"
Me.DC_Veiculos.BoundColumn = "autonum"
Me.DC_Veiculos.Refresh


Sql = "select null as autonum,null as lote, null as quantidade"
Adodc1.ConnectionString = ConnectionString
Adodc1.RecordSource = Sql
Adodc1.Refresh


End Sub

Private Sub Grid_OC_Click()

    If Not Adodc1.Recordset.EOF Then
        If Nnull(Adodc1.Recordset("autonum"), 0) <> 0 Then
            Sql = "select autonum,chassis,dt_carregamento"
            Sql = Sql & " from sgipa.dbo.tb_carga_solta_chassis"
            Sql = Sql & " where registro_saida=" & Adodc1.Recordset!autonum
        Else
            Sql = "select 0 as autonum,'' as chassis,null as dt_carregamento"
        End If
        Adodc2.ConnectionString = ConnectionString
        Adodc2.RecordSource = Sql
        Adodc2.Refresh
    End If

End Sub

Private Sub TxtChassi_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Then Exit Sub
ERRO = 1
If Chr(KeyAscii) >= "A" And Chr(KeyAscii) <= "Z" Then ERRO = 0
If Chr(KeyAscii) >= "a" And Chr(KeyAscii) <= "z" Then ERRO = 0
If Chr(KeyAscii) >= "0" And Chr(KeyAscii) <= "9" Then ERRO = 0
If ERRO = 1 Then KeyAscii = 0

End Sub


Private Sub TxtChassi_LostFocus()
TxtChassi = UCase(TxtChassi)
End Sub


