VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLibRegistro 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Consulta \ Liberação de Registro"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6720
   FillStyle       =   0  'Solid
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
   ScaleHeight     =   7845
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCpfMot 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      TabIndex        =   22
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox listaCarga 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   120
      TabIndex        =   21
      Top             =   5400
      Width           =   6375
   End
   Begin VB.TextBox txtPlacas 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3840
      TabIndex        =   20
      Top             =   3150
      Width           =   2655
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   375
      Left            =   5280
      TabIndex        =   18
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdRegistra 
      Caption         =   "Registrar"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox txtSefaz 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   16
      Top             =   4800
      Width           =   6375
   End
   Begin VB.TextBox txtSiscarga 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   14
      Top             =   3960
      Width           =   6375
   End
   Begin VB.TextBox txtMot 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   12
      Top             =   3150
      Width           =   3615
   End
   Begin VB.TextBox txtTransp 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   6375
   End
   Begin VB.TextBox txtBL 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      TabIndex        =   8
      Top             =   1680
      Width           =   4455
   End
   Begin VB.TextBox txtLote 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtPeriodo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   6375
   End
   Begin VB.CommandButton cmdFiltra 
      Caption         =   "Filtrar"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   200
      Width           =   975
   End
   Begin MSMask.MaskEdBox mskProtocolo 
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   11
      Mask            =   "######/####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblGLME 
      BackStyle       =   0  'Transparent
      Caption         =   "GLME, favor entregar documentação (guias) no Registro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   645
      Left            =   120
      TabIndex        =   23
      Top             =   7080
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Placas:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3840
      TabIndex        =   19
      Top             =   2850
      Width           =   870
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Status Sefaz:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   4440
      Width           =   1845
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Status Siscarga:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   2205
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Motorista:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   2850
      Width           =   1260
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transportadora:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   2100
      Width           =   1995
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Documento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2040
      TabIndex        =   7
      Top             =   1380
      Width           =   1770
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lote:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1380
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Período:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Protocolo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1650
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   8055
      Left            =   -120
      Top             =   -120
      Width           =   6885
   End
End
Attribute VB_Name = "frmLibRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdFiltra_Click()
Dim Rst As New ADODB.Recordset
Dim Sql As String
'Dim com As New regComex
Dim prot As String

If Val(Nnull(mskProtocolo.ClipText, 0)) = 0 Then
    Exit Sub
End If

prot = ObtemNumero(mskProtocolo.text)
prot = Right("0000000000" & prot, 10)
prot = Val(Mid(prot, 1, 6)) & "/" & Mid(prot, 7, 4)

Sql = "Select AG.PROTOCOLO, AG.PERIODO, AG.LOTE, AG.NUM_DOCUMENTO,  c.razao As TRANSPORTADORA, AG.NOME_MOTORISTA, AG.PLACA_CAVALO, AG.PLACA_CARRETA, AG.AUTONUM_CS, P.FLAG_CONSULTA_LIBERACAO_COLETOR, AG.CPFMOT "
Sql = Sql & " From " & Banco_Sgipa & "VW_AGENDA_CS ag "
Sql = Sql & " inner Join " & Banco_Operador & "tb_cad_transportadoras c ON Ag.COD_TRANSPORTADORA=C.AUTONUM "
Sql = Sql & " INNER JOIN " & Banco_Sgipa & "TB_BL BL ON AG.LOTE = BL.AUTONUM "
Sql = Sql & " INNER JOIN " & Banco_Operador & "TB_PATIOS P ON BL.PATIO = P.AUTONUM "
Sql = Sql & " WHERE PROTOCOLO ='" & prot & "' "
PRSet Rst, Sql
If Not Rst.EOF Then
    txtPeriodo.text = Rst!Periodo
    txtLote.text = Rst!Lote
    txtBL.text = Rst!Num_Documento
    txtTransp.text = Rst!Transportadora
    txtMot.text = Rst!NOME_MOTORISTA
    txtPlacas.text = Rst!PLACA_CAVALO & " - " & Rst!PLACA_CARRETA
    txtCpfMot.text = Rst!cpfMot

    Call Carrega_GridCS(Rst!AUTONUM_CS)
    msgdoc = ""
    msgIcms = ""
    
    If Not validaSaidaComexRealizada(txtLote.text) Then
        If Nnull(Trim(msgdoc), 1) <> "" Then
            txtSiscarga.text = msgdoc
        Else
            txtSiscarga.text = "NÃO ANALISADO"
        End If

        If Nnull(Trim(msgIcms), 1) <> "" Then
            txtSefaz.text = msgIcms
        Else
            txtSefaz.text = "NÃO ANALISADO"
        End If
    Else
        txtSefaz.text = "LIBERADO"
        txtSiscarga.text = "LIBERADO"
    End If

    If Nnull(Rst!FLAG_CONSULTA_LIBERACAO_COLETOR, 0) > 0 Then
        cmdRegistra.Visible = True
    Else
        cmdRegistra.Visible = False
    End If
Else
    Sql = "SELECT DISTINCT "
    Sql = Sql & " CONVERT(VARCHAR, Gd.PERIODO_INICIAL, 103) + ' ' + CONVERT(VARCHAR(5),Gd.PERIODO_INICIAL, 108) + ' - ' + CONVERT(VARCHAR, Gd.PERIODO_FINAL, 103) + ' ' + CONVERT(VARCHAR(5), Gd.PERIODO_FINAL, 108) AS periodo, "
    Sql = Sql & " BL.AUTONUM AS LOTE, BL.NUM_DOCUMENTO, t.razao AS TRANSPORTADORA,  B.NOME AS NOME_MOTORISTA, "
    Sql = Sql & " v.placa_cavalo, v.placa_carreta, P.FLAG_CONSULTA_LIBERACAO_COLETOR, b.cpf as cpfMot "
    Sql = Sql & " FROM " & Banco_Sgipa & "tb_cntr_bl c INNER JOIN  "
    Sql = Sql & " " & Banco_Sgipa & "TB_AMR_CNTR_BL Amr on c.autonum = amr.cntr inner join "
    Sql = Sql & " " & Banco_Sgipa & "TB_BL BL on Amr.bl= BL.autonum inner join "
    Sql = Sql & " " & Banco_Operador & "TB_GD_RESERVA GD on c.autonum_gd_reserva = gd.autonum_gd_reserva INNER JOIN "
    Sql = Sql & " " & Banco_Sgipa & "tb_ag_motoristas a  ON c.autonum_motorista = a.autonum LEFT JOIN "
    Sql = Sql & " " & Banco_Operador & "tb_motoristas b ON   a.cnh = b.cnh  INNER JOIN "
    Sql = Sql & " " & Banco_Sgipa & "tb_ag_veiculos v ON c.autonum_veiculo = v.autonum INNER JOIN "
    Sql = Sql & " " & Banco_Operador & "tb_cad_transportadoras t ON c.autonum_transporte_agenda = t.autonum INNER JOIN "
    Sql = Sql & " " & Banco_Operador & "TB_PATIOS P ON C.PATIO = P.AUTONUM "
    Sql = Sql & " WHERE bl.flag_ativo=1 "
    Sql = Sql & "   AND C.NUM_PROTOCOLO=" & Trim(str(Val(Mid$(prot, 1, 6))))
    Sql = Sql & "   AND C.ANO_PROTOCOLO=" & Trim(str(Val(Right$(prot, 4))))
    PRSet Rst, Sql
    If Not Rst.EOF Then
        txtPeriodo.text = Rst!Periodo
        txtLote.text = Rst!Lote
        txtBL.text = Rst!Num_Documento
        txtTransp.text = Rst!Transportadora
        txtMot.text = Rst!NOME_MOTORISTA
        txtPlacas.text = Rst!PLACA_CAVALO & " - " & Rst!PLACA_CARRETA
        txtCpfMot.text = Rst!cpfMot
        
        Call Carrega_GridCNTR(prot)

        msgdoc = ""
        msgIcms = ""
        If Not validaSaidaComexRealizada(txtLote.text) Then
            If msgdoc <> "" Then
                txtSiscarga.text = msgdoc
            Else
                txtSiscarga.text = "NÃO ANALISADO"
            End If

            If msgIcms <> "" Then
                txtSefaz.text = msgIcms
            Else
                txtSefaz.text = "NÃO ANALISADO"
            End If
        Else
            txtSefaz.text = "LIBERADO"
            txtSiscarga.text = "LIBERADO"
        End If

        If Nnull(Rst!FLAG_CONSULTA_LIBERACAO_COLETOR, 0) > 0 Then
            cmdRegistra.Visible = True
        Else
            cmdRegistra.Visible = False
        End If
    Else
        MsgBox "Protocolo não encontrado!"
        Limpa
    End If
End If
Rst.Close

If Nnull(db.Execute("SELECT ISNULL(FLAG_GLME,0) FROM TB_ETAPAS_CARGA WHERE BL =" & Nnull(txtLote.text, 0)).Fields(0), 0) > 0 Then
    lblGLME.Visible = True
Else
    lblGLME.Visible = False
End If




End Sub

Private Sub Limpa()

mskProtocolo.text = "______/____"
txtPeriodo.text = ""
txtSiscarga.text = ""
txtSefaz.text = ""
txtLote.text = ""
txtBL.text = ""
txtMot.text = ""
txtPlacas.text = ""
txtTransp.text = ""
listaCarga.Clear
txtCpfMot.text = ""

End Sub

Private Sub cmdLimpar_Click()
Limpa
End Sub

Private Sub cmdRegistra_Click()
On Error GoTo trataErro:
Dim sSql As String
Dim rsaux As New ADODB.Recordset

sSql = "SELECT P.FLAG_CONSULTA_LIBERACAO "
sSql = sSql & " FROM " & Banco_Sgipa & "TB_BL BL "
sSql = sSql & " LEFT JOIN " & Banco_Operador & "TB_PATIOS P ON BL.PATIO = P.AUTONUM "
sSql = sSql & " WHERE BL.AUTONUM =" & txtLote.text
PRSet rsaux, sSql
If Not rsaux.EOF Then
    If Nnull(rsaux!FLAG_CONSULTA_LIBERACAO, 0) = 0 Then
        MsgBox "Pátio parametrizado para não realizar consultas!", vbInformation
        Exit Sub
    End If
End If

msgIcms = ""
msgdoc = ""
If validaSaidaComex(Nnull(txtLote.text, 0)) = False Then
    If msgdoc <> "" Or msgIcms <> "" Then
        MSG = "Liberação Não autorizada (Sefaz\SISCARGA)" & vbCrLf
        If msgdoc <> "" Then MSG = MSG & "Retorno SISCARGA:" & msgdoc & vbCrLf
        If msgIcms <> "" Then MSG = MSG & "Retorno SEFAZ:" & msgIcms & vbCrLf
        MSG = MSG & "Deseja Continuar?"
        If MsgBox(MSG, vbYesNo, Me.Caption) = vbNo Then
            txtSefaz.text = msgIcms
            txtSiscarga.text = msgdoc
            Exit Sub
        End If
    Else
        MSG = "Não foi possivel consulta a liberação (Sefaz\SISCARGA)" & vbCrLf
        MSG = MSG & "Deseja Continuar?"
        If MsgBox(MSG, vbYesNo, Me.Caption) = vbNo Then
            txtSefaz.text = msgIcms
            txtSiscarga.text = msgdoc
            Exit Sub
        End If
    End If
    'MsgBox ("Saída não permitida!")
    'txtSefaz.text = msgIcms
    'txtSiscarga.text = msgDoc
    'Exit Sub
Else
    'ENTREGA SISCARGA\SEFAZ
    If registraSaidaComex(Nnull(txtLote.text, 0), 0, txtCpfMot.text) = False Then
        If msgdoc <> "" Or msgIcms <> "" Then
            MSG = "Não foi possivel registrar a Saída (Sefaz\SISCARGA)" & vbCrLf
            If msgdoc <> "" Then MSG = MSG & "Retorno SISCARGA:" & msgdoc & vbCrLf
            If msgIcms <> "" Then MSG = MSG & "Retorno SEFAZ:" & msgIcms & vbCrLf
            MSG = MSG & "Deseja Continuar?"
            If MsgBox(MSG, vbYesNo, Me.Caption) = vbNo Then
                txtSefaz.text = msgIcms
                txtSiscarga.text = msgdoc
                Exit Sub
            End If
        Else
            MSG = "Não foi possivel registrar a Saída (Sefaz\SISCARGA)" & vbCrLf
            MSG = MSG & "Deseja Continuar?"
            If MsgBox(MSG, vbYesNo, Me.Caption) = vbNo Then
                txtSefaz.text = msgIcms
                txtSiscarga.text = msgdoc
                Exit Sub
            End If
        End If
        'MsgBox ("Saída não permitida!")
        'txtSefaz.text = msgIcms
        'txtSiscarga.text = msgDoc
        'Exit Sub
    Else
        MsgBox ("Saída Registrada com sucesso no SISCARGA\SEFAZ!")
        cmdFiltra_Click
    End If
End If

Exit Sub
trataErro:
    MsgBox "Não foi possivel registrar a saida da carga!"
    Err.Clear


End Sub

Private Sub Carrega_GridCS(AutonumCs As Long)

listaCarga.Clear

Dim rsaux As New ADODB.Recordset
Dim Sql As String
Sql = "Select 'Carga Solta - ' + CONVERT(VARCHAR,A.ITEM) + ' - ' + C.DESCR AS CARGA "
Sql = Sql & " FROM " & Banco_Sgipa & "TB_CARGA_SOLTA A LEFT JOIN "
Sql = Sql & " " & Banco_Sgipa & "DTE_TB_EMBALAGENS C ON A.EMBALAGEM = C.CODE "
Sql = Sql & " WHERE A.AUTONUM= " & Val(AutonumCs)
PRSet rsaux, Sql
While Not rsaux.EOF
    listaCarga.AddItem (Nnull(rsaux!CARGA, 1))
    rsaux.MoveNext
Wend

        
End Sub
    
Private Sub Carrega_GridCNTR(protocolo As String)

listaCarga.Clear

Dim rsaux As New ADODB.Recordset
Dim Sql As String
        
Sql = " SELECT ID_CONTEINER AS CARGA FROM " & Banco_Sgipa & "TB_CNTR_BL "
Sql = Sql & " WHERE NUM_PROTOCOLO = " & Trim(str(Val(Mid$(protocolo, 1, 6))))
Sql = Sql & " And ANO_PROTOCOLO = " & Trim(str(Val(Right$(protocolo, 4))))
PRSet rsaux, Sql
While Not rsaux.EOF
    listaCarga.AddItem (Nnull(rsaux!CARGA, 1))
    rsaux.MoveNext
Wend
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub lblGLME_Click()

End Sub
