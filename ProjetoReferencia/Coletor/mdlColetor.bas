Attribute VB_Name = "mdlColetor"
Public Col_Mensagem$
Public Col_Imagem As Integer
Public Col_Patio As Integer
Public Col_Cancelado As Boolean
Public Col_Login As String
Public Col_IdInventario As Long
Public Col_DescrInventario As String
Public Col_Reconhecimento As Integer
Public Col_Tam_Pos_Yard As Integer
Public Col_Flag_Pilha_2040 As Integer
Public Patio_coletor As Integer
Public ServidorOperador As String



Sub COL_OBS_Coletor(Lote As Long)

Dim wCOL_OBS_Coletor  As String
wCOL_OBS_Coletor = ""

Dim Sql As String
Sql = "SELECT SUM(QUANTIDADE) AS QTDE, DESCR_EMBALAGEM FROM VW_DESOVA_COL_ITENS_DESOVADOS WHERE LOTE=" & Lote & " GROUP BY DESCR_EMBALAGEM "
Set tb1 = New ADODB.Recordset
tb1.Open Sql, db, 1, 1
If Not tb1.EOF Then
    wCOL_OBS_Coletor = "DESOVADOS "
    While Not tb1.EOF
        If wCOL_OBS_Coletor <> "DESOVADOS " Then wCOL_OBS_Coletor = wCOL_OBS_Coletor & ","
        wCOL_OBS_Coletor = wCOL_OBS_Coletor & " " & Nnull(tb1!QTDE, 0) & " " & Nnull(tb1!DESCR_EMBALAGEM, 1)
        tb1.MoveNext
    Wend
    
End If
tb1.Close

Sql = "UPDATE SGIPA..TB_BL SET OBS_DESOVA='" & wCOL_OBS_Coletor & "' WHERE AUTONUM=" & Lote
db.Execute Sql

End Sub


Public Function Valida_Acesso_Botao(QualUsuario As String, QualNomeBotao As String) As Boolean
Dim Sql As String

If QualUsuario = "MICROLED" Then
    Valida_Acesso_Botao = True
    Exit Function
End If

Sql = "SELECT 1 "
Sql = Sql & " FROM"
Sql = Sql & " sgipa..TB_SYS_FUNCOES,"
Sql = Sql & " sgipa..TB_SYS_GRP_PERMISSOES,"
Sql = Sql & " sgipa..TB_SYS_USER_GRUPOS, "
Sql = Sql & " sgipa..TB_CAD_USUARIOS "
Sql = Sql & " WHERE"
Sql = Sql & "((tb_cad_usuarios.autonum = tb_sys_user_grupos.autonumuser) "
Sql = Sql & " AND (tb_sys_funcoes.codfunc = tb_sys_grp_permissoes.codfunc) "
Sql = Sql & " AND (tb_sys_grp_permissoes.codgrupo = tb_sys_user_grupos.codgrupo) "
Sql = Sql & " AND (UPPER(tb_cad_usuarios.usuario) = '" & QualUsuario & "') "
Sql = Sql & " AND (tb_sys_funcoes.sistema = 'COLETOR_ARM') "
Sql = Sql & " AND (UPPER(tb_sys_funcoes.Nomeobj) = '" & QualNomeBotao & "')) "
Sql = Sql & " and (tb_sys_grp_permissoes.codtipoperm=5) "

Set tbUA = New ADODB.Recordset
tbUA.Open Sql, db, 1, 1
If Not tbUA.EOF Then
    tbUA.Close
    Valida_Acesso_Botao = True
Else
    tbUA.Close
    Valida_Acesso_Botao = False
End If

End Function
       
