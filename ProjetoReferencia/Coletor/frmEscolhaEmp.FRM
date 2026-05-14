VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form FrmEscolhaEmpresa 
   BackColor       =   &H00A26842&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2145
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3120
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEscolhaEmp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2145
   ScaleWidth      =   3120
   Begin VB.PictureBox pBARRA 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00BA9F8D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   3120
      TabIndex        =   2
      Top             =   0
      Width           =   3120
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2730
         Picture         =   "frmEscolhaEmp.frx":0442
         Stretch         =   -1  'True
         Top             =   45
         Width           =   300
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Esolha de Empresa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E3F7FF&
         Height          =   225
         Index           =   5
         Left            =   60
         TabIndex        =   3
         Top             =   105
         Width           =   4200
      End
   End
   Begin VB.PictureBox pQUADRO 
      Appearance      =   0  'Flat
      BackColor       =   &H00BA9F8D&
      ForeColor       =   &H80000008&
      Height          =   1600
      Left            =   75
      ScaleHeight     =   1575
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   435
      Width           =   2985
      Begin VB.CommandButton cmd 
         Appearance      =   0  'Flat
         BackColor       =   &H00BA9F8D&
         Cancel          =   -1  'True
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   675
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   915
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmEscolhaEmp.frx":210C
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Tag             =   "COD_EMPRESA;COD_EMPRESA"
         Top             =   465
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ListField       =   "nome_fantasia"
         BoundColumn     =   "autonum"
         Text            =   ""
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   1680
         Top             =   960
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "EMPRESA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   240
         Width           =   3225
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   -15
      X2              =   4965
      Y1              =   375
      Y2              =   375
   End
End
Attribute VB_Name = "FrmEscolhaEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
  Adodc1.ConnectionString = ConnectionString
  Adodc1.RecordSource = "Select autonum, nome_fantasia from " & Banco_Sgipa & "tb_empresas order by nome_fantasia"
  Adodc1.Refresh
  Set Adodc1.Recordset.ActiveConnection = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Sendkeys "{TAB}"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub Image1_Click()
End
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image1.BorderStyle = 1
End Sub

Private Sub pBARRA_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Not Image1.BorderStyle = 0 Then Image1.BorderStyle = 0
End Sub

Private Sub cmd_Click(Index As Integer)
Dim p_in As String
Dim Tbpatio As ADODB.Recordset 'Tabela de Patios
        If DataCombo2 = "" Then Exit Sub
        Cod_Empresa = Me.DataCombo2.BoundText
        PRSet Tbpatio, "SELECT  * FROM " & Banco_Operador & "TB_PATIOS  WHERE (COD_EMPRESA=" & Cod_Empresa & " OR 0=" & Cod_Empresa & ")"
        If Not Tbpatio.EOF Then
              Col_Patio = 0
              Cod_Empresa_Patio = "PATIO IN("
               Do
                p_in = p_in & Tbpatio!AUTONUM
                If Col_Patio = 0 Then
                    Col_Patio = Tbpatio!AUTONUM
                    Col_Tam_Pos_Yard = Tbpatio!tam_pos_yard
                    Col_Flag_Pilha_2040 = Tbpatio!Flag_Pilha_2040
                End If
                Cod_Empresa_Patio = Cod_Empresa_Patio & Tbpatio!AUTONUM
                Tbpatio.MoveNext
                If Not Tbpatio.EOF Then
                  p_in = p_in & ","
                  Cod_Empresa_Patio = Cod_Empresa_Patio & ","
                Else
                  Cod_Empresa_Patio = Cod_Empresa_Patio & ")"
               End If
             Loop Until Tbpatio.EOF
        End If
        If db.State = 0 Then
         db.Open
        End If
 If Val(Cod_Empresa) <> 0 Then Unload Me
End Sub

