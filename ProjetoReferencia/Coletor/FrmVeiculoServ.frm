VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form FrmVeiculoServ 
   Caption         =   "Coletor- Veiculos"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   4005
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
      Left            =   3690
      Top             =   1155
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   1395
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   2461
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "FrmVeiculoServ.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DC_Veiculos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.CommandButton Command3 
         Caption         =   "Retorno DTA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         TabIndex        =   5
         Top             =   945
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Associar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         TabIndex        =   4
         Top             =   600
         Width           =   2145
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Voltar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2280
         TabIndex        =   3
         Top             =   585
         Width           =   1170
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
Attribute VB_Name = "FrmVeiculoServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Unload Me
    


End Sub

Private Sub Command2_Click()
    If Val(DC_Veiculos.BoundText) = 0 Then
        MsgBox "Escolha um Veículo"
        Exit Sub
    End If
    
    Dim Rs As New ADODB.Recordset
    
If CntrVazio > 0 And Val(DC_Veiculos.BoundText) > 0 Then

  If Val(DC_Veiculos.BoundText) = 0 Then
    MsgBox "Escolha um Placa antes de Associar"
    Exit Sub
  End If
    
    SQl = "UPDATE A SET A.motorista=B.ID_MOTORISTA,A.carreta=B.CARRETA,A.cavalo=B.PLACA,A.transportadora=B.id_transportadora,gate=b.autonum "
    SQl = SQl & " FROM  sgipa..tb_devolucao_vazio A , operador..TB_GATE_NEW  B  WHERE  ( A.cntr_carreta_1=" & CntrVazio & " OR A.cntr_carreta_2=" & CntrVazio & ")"
    SQl = SQl & " AND  B.AUTONUM=" & Val(DC_Veiculos.BoundText)
    Executa SQl
    
    SQl = "DELETE TB_GOS_AMR_CHRONOS WHERE ID_CHRONOS_GOS IN( SELECT ID_CHRONOS_GOS FROM  TB_GOS_AMR_CHRONOS G INNER JOIN "
    SQl = SQl & " sgipa..tb_devolucao_vazio I ON G.ID_ORIGEM=I.AUTONUM AND PC_ORIGEM='RETVZ' AND IN_OUT='OUT'  "
    SQl = SQl & " INNER JOIN SGIPA..TB_CNTR_BL C ON I.cntr_carreta_1=C.AUTONUM   "
    SQl = SQl & "  WHERE G.TOS_CARGO_NUMBER=C.ID_CONTEINER AND G.PATIO=C.PATIO AND I.cntr_carreta_1=" & CntrVazio & ")"
    Executa SQl
    
     
    SQl = "INSERT INTO TB_GOS_AMR_CHRONOS (ID_ORIGEM,PC_ORIGEM,IN_OUT,FUNCAO_GATE,AUTONUM_GATE,TOS_CARGO_NUMBER,TOS_VEHICLE_NUMBER,TOS_TRAILER_NUMBER,TOS_DRIVER_CPF,YARD_NAME,PATIO,ID_MOTORISTA,TOS_CARGO_NUMBER2) "
    SQl = SQl & " SELECT I.AUTONUM,'RETVZ','OUT',15,0,C.ID_CONTEINER,I.CAVALO,I.CARRETA,M.CPF,P.DESCR_RESUMIDO,PATIO,I.MOTORISTA,'__________-_'"
    SQl = SQl & " FROM sgipa..tb_devolucao_vazio I INNER JOIN OPERADOR..TB_MOTORISTAS M ON I.MOTORISTA=M.AUTONUM"
    SQl = SQl & " INNER JOIN SGIPA..TB_CNTR_BL C ON I.cntr_carreta_1=C.AUTONUM"
    SQl = SQl & " INNER JOIN OPERADOR..TB_PATIOS P ON C.PATIO=P.AUTONUM WHERE I.cntr_carreta_1=" & CntrVazio
    Executa SQl
    
Else
    
    
    If ModoRepesagem = 2 Then
    
            
            SQl = "INSERT INTO SGIPA.dbo.TB_GOS_AMR_CHRONOS_DTA"
            SQl = SQl & " (ID_CHRONOS_GOS,ID_GOS, ID_ORIGEM, PC_ORIGEM, IN_OUT, GOS_GATE_ID, GOS_WEIGHT, GOS_DRIVER_CPF, GOS_VEHICLE_NUMBER"
            SQl = SQl & " , GOS_TRAILER_NUMBER, GOS_CARGO_NUMBER, GOS_CARGO_ISO, TOS_CARGO_NUMBER, TOS_VEHICLE_NUMBER, TOS_TRAILER_NUMBER"
            SQl = SQl & " , TOS_DRIVER_CPF, FUNCAO_GATE, AUTONUM_GATE, YARD_NAME, Patio_Info, Quadra_Info, Info01, Info02, Info03, Info04"
            SQl = SQl & " , patio, ID_MOTORISTA, AUTONUM_CS, TOS_CARGO_NUMBER2"
            SQl = SQl & " )"
            SQl = SQl & " SELECT ID_CHRONOS_GOS,ID_GOS, ID_ORIGEM, PC_ORIGEM, IN_OUT, GOS_GATE_ID, GOS_WEIGHT, GOS_DRIVER_CPF, GOS_VEHICLE_NUMBER"
            SQl = SQl & " , GOS_TRAILER_NUMBER, GOS_CARGO_NUMBER, GOS_CARGO_ISO, TOS_CARGO_NUMBER, TOS_VEHICLE_NUMBER, TOS_TRAILER_NUMBER"
            SQl = SQl & " , TOS_DRIVER_CPF, FUNCAO_GATE, AUTONUM_GATE, YARD_NAME, Patio_Info, Quadra_Info, Info01, Info02, Info03, Info04"
            SQl = SQl & " , patio, ID_MOTORISTA, AUTONUM_CS, TOS_CARGO_NUMBER2"
            SQl = SQl & " FROM SGIPA.DBO.TB_GOS_AMR_CHRONOS WHERE IN_OUT='OUT' AND FUNCAO_GATE=12 AND ISNULL(AUTONUM_GATE,0)=0 AND TOS_VEHICLE_NUMBER='" & Left(Me.DC_Veiculos.Text, 8) & "'"
            db.Execute (SQl)
            
            
            SQl = "INSERT INTO SGIPA.dbo.TB_GOS_AMR_CHRONOS"
            SQl = SQl & " ( ID_GOS, ID_ORIGEM, PC_ORIGEM, IN_OUT, GOS_GATE_ID, GOS_WEIGHT, GOS_DRIVER_CPF, GOS_VEHICLE_NUMBER"
            SQl = SQl & " , GOS_TRAILER_NUMBER, GOS_CARGO_NUMBER, GOS_CARGO_ISO, TOS_CARGO_NUMBER, TOS_VEHICLE_NUMBER, TOS_TRAILER_NUMBER"
            SQl = SQl & " , TOS_DRIVER_CPF, FUNCAO_GATE, AUTONUM_GATE, YARD_NAME, Patio_Info, Quadra_Info, Info01, Info02, Info03, Info04"
            SQl = SQl & " , patio, ID_MOTORISTA, AUTONUM_CS, TOS_CARGO_NUMBER2"
            SQl = SQl & " )"
            SQl = SQl & " SELECT TOP 1 ID_GOS, ID_ORIGEM, PC_ORIGEM, IN_OUT, GOS_GATE_ID, GOS_WEIGHT, GOS_DRIVER_CPF, GOS_VEHICLE_NUMBER"
            SQl = SQl & " , GOS_TRAILER_NUMBER, GOS_CARGO_NUMBER, GOS_CARGO_ISO, TOS_CARGO_NUMBER, TOS_VEHICLE_NUMBER, TOS_TRAILER_NUMBER"
            SQl = SQl & " , TOS_DRIVER_CPF, 40 AS FUNCAO_GATE, AUTONUM_GATE, YARD_NAME, Patio_Info, Quadra_Info, Info01, Info02, Info03, Info04"
            SQl = SQl & " , patio, ID_MOTORISTA, " & LoteDTA & " AS AUTONUM_CS, TOS_CARGO_NUMBER2"
            SQl = SQl & " FROM SGIPA.DBO.TB_GOS_AMR_CHRONOS WHERE IN_OUT='OUT' AND FUNCAO_GATE=12 AND ISNULL(AUTONUM_GATE,0)=0 AND TOS_VEHICLE_NUMBER='" & Left(Me.DC_Veiculos.Text, 8) & "'"
            db.Execute (SQl)
            
            
            SQl = "DELETE FROM SGIPA.DBO.TB_GOS_AMR_CHRONOS WHERE IN_OUT='OUT' AND FUNCAO_GATE=12 AND ISNULL(AUTONUM_GATE,0)=0 AND TOS_VEHICLE_NUMBER='" & Left(Me.DC_Veiculos.Text, 8) & "'"
            db.Execute (SQl)
            
 
    
    Else
        If LoteDTA <> 0 Then
        
            
            SQl = "SELECT * FROM operador..tb_amr_gate WHERE dta='PESAR' AND gate=" & Me.DC_Veiculos.BoundText
            PRSet Rs, SQl
            If Not Rs.EOF Then
                Rs.Close
                MsgBox ("Já consta carga vinculada pendente de pesagem")
                Exit Sub
            End If
            Rs.Close
            
            SQl = "INSERT INTO SGIPA.dbo.TB_GOS_AMR_CHRONOS_DTA"
            SQl = SQl & " (ID_CHRONOS_GOS,ID_GOS, ID_ORIGEM, PC_ORIGEM, IN_OUT, GOS_GATE_ID, GOS_WEIGHT, GOS_DRIVER_CPF, GOS_VEHICLE_NUMBER"
            SQl = SQl & " , GOS_TRAILER_NUMBER, GOS_CARGO_NUMBER, GOS_CARGO_ISO, TOS_CARGO_NUMBER, TOS_VEHICLE_NUMBER, TOS_TRAILER_NUMBER"
            SQl = SQl & " , TOS_DRIVER_CPF, FUNCAO_GATE, AUTONUM_GATE, YARD_NAME, Patio_Info, Quadra_Info, Info01, Info02, Info03, Info04"
            SQl = SQl & " , patio, ID_MOTORISTA, AUTONUM_CS, TOS_CARGO_NUMBER2"
            SQl = SQl & " )"
            SQl = SQl & " SELECT ID_CHRONOS_GOS,ID_GOS, ID_ORIGEM, PC_ORIGEM, IN_OUT, GOS_GATE_ID, GOS_WEIGHT, GOS_DRIVER_CPF, GOS_VEHICLE_NUMBER"
            SQl = SQl & " , GOS_TRAILER_NUMBER, GOS_CARGO_NUMBER, GOS_CARGO_ISO, TOS_CARGO_NUMBER, TOS_VEHICLE_NUMBER, TOS_TRAILER_NUMBER"
            SQl = SQl & " , TOS_DRIVER_CPF, FUNCAO_GATE, AUTONUM_GATE, YARD_NAME, Patio_Info, Quadra_Info, Info01, Info02, Info03, Info04"
            SQl = SQl & " , patio, ID_MOTORISTA, AUTONUM_CS, TOS_CARGO_NUMBER2"
            SQl = SQl & " FROM SGIPA.DBO.TB_GOS_AMR_CHRONOS WHERE IN_OUT='OUT' AND FUNCAO_GATE=41 AND ISNULL(AUTONUM_GATE,0)=0 AND TOS_VEHICLE_NUMBER='" & Left(Me.DC_Veiculos.Text, 8) & "'"
            db.Execute (SQl)
            
            
            SQl = "INSERT INTO SGIPA.dbo.TB_GOS_AMR_CHRONOS"
            SQl = SQl & " ( ID_GOS, ID_ORIGEM, PC_ORIGEM, IN_OUT, GOS_GATE_ID, GOS_WEIGHT, GOS_DRIVER_CPF, GOS_VEHICLE_NUMBER"
            SQl = SQl & " , GOS_TRAILER_NUMBER, GOS_CARGO_NUMBER, GOS_CARGO_ISO, TOS_CARGO_NUMBER, TOS_VEHICLE_NUMBER, TOS_TRAILER_NUMBER"
            SQl = SQl & " , TOS_DRIVER_CPF, FUNCAO_GATE, AUTONUM_GATE, YARD_NAME, Patio_Info, Quadra_Info, Info01, Info02, Info03, Info04"
            SQl = SQl & " , patio, ID_MOTORISTA, AUTONUM_CS, TOS_CARGO_NUMBER2"
            SQl = SQl & " )"
            SQl = SQl & " SELECT TOP 1 ID_GOS, ID_ORIGEM, PC_ORIGEM, IN_OUT, GOS_GATE_ID, GOS_WEIGHT, GOS_DRIVER_CPF, GOS_VEHICLE_NUMBER"
            SQl = SQl & " , GOS_TRAILER_NUMBER, GOS_CARGO_NUMBER, GOS_CARGO_ISO, TOS_CARGO_NUMBER, TOS_VEHICLE_NUMBER, TOS_TRAILER_NUMBER"
            SQl = SQl & " , TOS_DRIVER_CPF, 20 AS FUNCAO_GATE, AUTONUM_GATE, YARD_NAME, Patio_Info, Quadra_Info, Info01, Info02, Info03, Info04"
            SQl = SQl & " , patio, ID_MOTORISTA, " & LoteDTA & " AS AUTONUM_CS, TOS_CARGO_NUMBER2"
            SQl = SQl & " FROM SGIPA.DBO.TB_GOS_AMR_CHRONOS WHERE IN_OUT='OUT' AND FUNCAO_GATE=41 AND ISNULL(AUTONUM_GATE,0)=0 AND TOS_VEHICLE_NUMBER='" & Left(Me.DC_Veiculos.Text, 8) & "'"
            db.Execute (SQl)
            
            
            SQl = "DELETE FROM SGIPA.DBO.TB_GOS_AMR_CHRONOS WHERE IN_OUT='OUT' AND FUNCAO_GATE=41 AND ISNULL(AUTONUM_GATE,0)=0 AND TOS_VEHICLE_NUMBER='" & Left(Me.DC_Veiculos.Text, 8) & "'"
            db.Execute (SQl)
            
            
            SQl = "update operador..tb_amr_gate set dta='PESAR' where cs_ipa=" & LoteDTA & " and gate=" & Me.DC_Veiculos.BoundText
            db.Execute (SQl)
        
        Else
        
            If P_Cntr <> 0 Then
                SQl = "SELECT COUNT(*) from sgipa..tb_gate_servico where autonum_veic=" & DC_Veiculos.BoundText & " AND DATA_ENTRADA IS NULL"
                If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then
                    MsgBox ("Consta pendencia de entrada para este veiculo/conteiner")
                    Exit Sub
                End If
                
                SQl = "update sgipa..tb_gate_servico set cntr= " & P_Cntr & "  where autonum_veic=" & DC_Veiculos.BoundText
                db.Execute (SQl)
            
            End If
            If P_BL <> 0 Then
                SQl = "SELECT COUNT(*) from sgipa..tb_gate_servico where autonum_veic=" & DC_Veiculos.BoundText & " AND DATA_ENTRADA IS NULL"
                If Nnull(db.Execute(SQl).Fields(0), 0) <> 0 Then
                    MsgBox ("Consta pendencia de entrada para este veiculo/conteiner")
                    Exit Sub
                End If
                SQl = "update sgipa..tb_gate_servico set lote= " & P_BL & "  where autonum_veic=" & DC_Veiculos.BoundText
                db.Execute (SQl)
            End If
        End If
    End If
 End If
 
 Unload Me
    
End Sub


Private Sub Command3_Click()

    If Val(DC_Veiculos.BoundText) = 0 Then
        MsgBox "Escolha um Veículo"
        Exit Sub
    End If
    
    Dim Rs As New ADODB.Recordset
    SQl = "SELECT * FROM operador..tb_amr_gate WHERE dta='PESAR' AND gate=" & Me.DC_Veiculos.BoundText
    PRSet Rs, SQl
    If Rs.EOF Then
        Rs.Close
        MsgBox ("Não consta carga vinculada pendente de pesagem")
        Exit Sub
    End If
    Rs.Close
    
    SQl = "UPDATE operador..tb_amr_gate set dta='OK' where gate=" & Me.DC_Veiculos.BoundText & " AND dta='PESAR'"
    db.Execute (SQl)
    
    
    SQl = "INSERT INTO SGIPA.dbo.TB_GOS_AMR_CHRONOS"
    SQl = SQl & " (ID_GOS, ID_ORIGEM, PC_ORIGEM, IN_OUT, GOS_GATE_ID, GOS_WEIGHT, GOS_DRIVER_CPF, GOS_VEHICLE_NUMBER"
    SQl = SQl & " , GOS_TRAILER_NUMBER, GOS_CARGO_NUMBER, GOS_CARGO_ISO, TOS_CARGO_NUMBER, TOS_VEHICLE_NUMBER, TOS_TRAILER_NUMBER"
    SQl = SQl & " , TOS_DRIVER_CPF, FUNCAO_GATE, AUTONUM_GATE, YARD_NAME, Patio_Info, Quadra_Info, Info01, Info02, Info03, Info04"
    SQl = SQl & " , patio, ID_MOTORISTA, AUTONUM_CS, TOS_CARGO_NUMBER2"
    SQl = SQl & " )"
    SQl = SQl & " SELECT ID_GOS, ID_ORIGEM, PC_ORIGEM, IN_OUT, GOS_GATE_ID, GOS_WEIGHT, GOS_DRIVER_CPF, GOS_VEHICLE_NUMBER"
    SQl = SQl & " , GOS_TRAILER_NUMBER, GOS_CARGO_NUMBER, GOS_CARGO_ISO, TOS_CARGO_NUMBER, TOS_VEHICLE_NUMBER, TOS_TRAILER_NUMBER"
    SQl = SQl & " , TOS_DRIVER_CPF, FUNCAO_GATE, AUTONUM_GATE, YARD_NAME, Patio_Info, Quadra_Info, Info01, Info02, Info03, Info04"
    SQl = SQl & " , patio, ID_MOTORISTA, AUTONUM_CS, TOS_CARGO_NUMBER2"
    SQl = SQl & " FROM SGIPA.DBO.TB_GOS_AMR_CHRONOS_DTA WHERE TOS_VEHICLE_NUMBER='" & Left(Me.DC_Veiculos.Text, 8) & "'"
    db.Execute (SQl)
    
    
    SQl = "DELETE FROM SGIPA.dbo.TB_GOS_AMR_CHRONOS_DTA WHERE TOS_VEHICLE_NUMBER='" & Left(Me.DC_Veiculos.Text, 8) & "'"
    db.Execute (SQl)
    
    SQl = "DELETE FROM SGIPA..TB_GOS_AMR_CHRONOS WHERE IN_OUT='OUT' AND TOS_VEHICLE_NUMBER='" & Left(Me.DC_Veiculos.Text, 8) & "' AND FUNCAO_GATE=20"
    db.Execute (SQl)
    
    
    Unload Me

End Sub


Private Sub Form_Load()
Dim RsV As New ADODB.Recordset


If CntrVazio > 0 Then
            Command3.Visible = False
            SQl = "select gn.autonum, gn.placa + ' / ' + gn.carreta as placa"
            SQl = SQl & " from OPERADOR..TB_GATE_NEW GN LEFT JOIN OPERADOR..TB_AMR_GATE AG ON"
            SQl = SQl & " GN.AUTONUM = AG.GATE where GN.FLAG_GATE_IN=1 AND GN.FLAG_GATE_OUT=0"
            SQl = SQl & " AND funcao_gate_entrada=5 AND AG.AUTONUM IS NULL"
            SQl = SQl & " AND PLACA IS NOT NULL AND DT_GATE_IN> GETDATE()-1"
            SQl = SQl & " AND PATIO=" & Col_Patio & " ORDER BY GN.PLACA"
Else
    If ModoRepesagem = 2 Then
            Command3.Visible = False
            SQl = "select c.autonum, c.placa + ' / ' + c.placa_carreta as placa"
            SQl = SQl & " from sgipa.dbo.TB_VEICULO_OPERACAO C"
            SQl = SQl & " where 1=1"
            SQl = SQl & " AND convert(varchar,c.DATA_Final,112) >=convert(varchar,getdate(),112)  "
            SQl = SQl & " AND   convert(varchar,c.DATA_INICIAL,112)<=convert(varchar,getdate(),112)"
            SQl = SQl & " AND FLAG_ATIVO=1 and tipo_servico=2  ORDER BY C.PLACA"
    Else
        If LoteDTA = 0 Then
            Command3.Visible = False
            SQl = "select c.autonum, c.placa + ' / ' + c.placa_carreta as placa"
            SQl = SQl & " from sgipa.dbo.TB_VEICULO_OPERACAO C"
            SQl = SQl & " where 1=1"
            SQl = SQl & " AND convert(varchar,c.DATA_Final,112) >=convert(varchar,getdate(),112)  "
            SQl = SQl & " AND   convert(varchar,c.DATA_INICIAL,112)<=convert(varchar,getdate(),112)"
            SQl = SQl & " AND FLAG_ATIVO=1 and tipo_servico=1 ORDER BY C.PLACA"
        Else
            Command3.Visible = True
            SQl = "select gn.autonum, gn.placa + ' / ' + gn.carreta as placa"
            SQl = SQl & " from OPERADOR..TB_GATE_NEW GN"
            SQl = SQl & " INNER JOIN OPERADOR..TB_AMR_GATE AG ON GN.AUTONUM = AG.GATE"
            SQl = SQl & " where GN.FLAG_GATE_IN=1 AND GN.FLAG_GATE_OUT=0 AND AG.CS_IPA=" & LoteDTA & "  ORDER BY GN.PLACA"
        End If
    End If
End If

PRSet RsV, SQl
Set Me.DC_Veiculos.RowSource = RsV
Me.DC_Veiculos.ListField = "placa"
Me.DC_Veiculos.BoundColumn = "autonum"
Me.DC_Veiculos.Refresh

If RsV.RecordCount > 0 Then
    DC_Veiculos.BoundText = RsV!autonum

End If


End Sub


