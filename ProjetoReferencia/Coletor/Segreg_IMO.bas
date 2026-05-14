Attribute VB_Name = "Segreg_IMO"


















Function Verifica_Regras_Seg_Imo(Id_Conteiner As String, Yard As String, Optional Patio As Integer = 1) As String
Verifica_Regras_Seg_Imo = ""
Dim MColP(1 To 16) As Integer



MColP(1) = 1
MColP(2) = 2
MColP(3) = 3
MColP(4) = 4
MColP(5) = 5
MColP(6) = 6
MColP(7) = 7
MColP(8) = 8
MColP(9) = 9
MColP(10) = 10
MColP(11) = 11
MColP(12) = 12
MColP(13) = 13
MColP(14) = 14
MColP(15) = 15
MColP(16) = 16


If Patio = 7 Then Patio = 1

Dim Mensagem(1 To 4) As String
Mensagem(1) = "Mesma pilha."
Mensagem(2) = "Espaço de 1 cntr."
Mensagem(3) = "Espaço de 2 cntr."
Mensagem(4) = "24metros."

Dim SQl As String
'Posicao a ser colocada é de pilha ?
SQl = "SELECT 0 FROM OPERADOR..TB_YARD WHERE PATIO=" & Patio & " AND YARD='" & Yard & "' and valida=1"
Set tbValida = New ADODB.Recordset
tbValida.Open SQl, db, 1, 1
If Not tbValida.EOF Then
    SQl = "SELECT ID_CONTEINER,IMO1,IMO2,IMO3,IMO4 FROM OPERADOR..VW_INVENT_SISTEMAS_IMO "
    SQl = SQl & " WHERE ID_CONTEINER='" & Id_Conteiner & "'"
    Set tb1 = New ADODB.Recordset 'O Conteiner é IMO ?
    tb1.Open SQl, db, 1, 1
    If Not tb1.EOF Then
        For i = 1 To 4
            If Trim(Nnull(tb1.Fields("IMO" & i), 1)) <> "" Then
                If Trim(Nnull(tb1.Fields("IMO" & i), 1)) <> "0" Then
                    'Procura se tem conteiner na mesma pilha que nao obedece regras de segregacao
                    
                    
                    SQl = "Select SEGREGACAO, CLASS2 AS CLASS from OPERADOR..tb_segregacao where CLASS1='" & Replace(Trim(Nnull(tb1.Fields("IMO" & i), 1)), ".", "") & "' AND SEGREGACAO in ('1','2','3','4') "
                    SQl = SQl & " UNION "
                    SQl = SQl & "Select SEGREGACAO, CLASS1 AS CLASS from OPERADOR..tb_segregacao where CLASS2='" & Replace(Trim(Nnull(tb1.Fields("IMO" & i), 1)), ".", "") & "' AND SEGREGACAO in ('1','2','3','4') "
                    
                    
                    Set tbBusca = New ADODB.Recordset
                    tbBusca.Open SQl, db, 1, 1
                    While Not tbBusca.EOF
'                        MsgBox TBBUSCA!Class
                        For j = 1 To 4
                            SQl = "Select id_conteiner, yard, imo" & j & " as imo from OPERADOR..vw_invent_sistemas_imo where yard like '%" & Mid(Yard, 1, 5) & "%' and replace(imo" & j & ",'.','') ='" & tbBusca!Class & "' and id_conteiner <>'" & Id_Conteiner & "'"
                            Set TBIMO = New ADODB.Recordset
                            TBIMO.Open SQl, db, 1, 1
                            If Not TBIMO.EOF Then
                                While Not TBIMO.EOF
                                    Verifica_Regras_Seg_Imo = Verifica_Regras_Seg_Imo & Mensagem(1) & " Conteiner: " & TBIMO!Id_Conteiner & " Posicao: " & TBIMO!Yard & " Imo:" & tbBusca!Class & vbCrLf
                                    TBIMO.MoveNext
                                Wend
                            End If
                            TBIMO.Close
                        Next j
                        tbBusca.MoveNext
                    Wend
                    tbBusca.Close
                    
                    
                    'Procura se tem conteiner na pilha adjacente que nao obedece regras de segregacao
                    'Qual MColP está a carga
                    For k = 1 To 16
                        If MColP(k) = Val(Mid$(Yard, 4, 2)) Then
                            Exit For
                        End If
                    Next k
                    
                    Dim PilhaAnt As String
                    Dim PilhaPos As String
                    Dim FiadaAnt As String
                    Dim FiadaPos As String
                    
                    
                    If k > 1 Then
                        PilhaAnt = Mid$(Yard, 1, 3) & Format(MColP(k - 1), "00") & Right(Yard, 1)
                    Else
                        PilhaAnt = "--"
                    End If
                    PilhaPos = Mid$(Yard, 1, 3) & Format(MColP(k + 1), "00") & Right(Yard, 1)
                    
                    FiadaAnt = Mid$(Yard, 1, 1) & Format(Val(Mid$(Yard, 2, 2)) - 1, "00") & Format(MColP(k), "00") & Right(Yard, 1)
                    FiadaPos = Mid$(Yard, 1, 1) & Format(Val(Mid$(Yard, 2, 2)) + 1, "00") & Format(MColP(k), "00") & Right(Yard, 1)
                    
                    
                    
                    
                    
                    'Nao pode ter conteiner na fiada adjacente nem na pilha adjacente
                    SQl = "Select CLASS2 AS CLASS from OPERADOR..tb_segregacao where CLASS1='" & Replace(Trim(Nnull(tb1.Fields("IMO" & i), 1)), ".", "") & "' AND SEGREGACAO in ('2','3','4') "
                    SQl = SQl & " UNION "
                    SQl = SQl & "Select CLASS1 AS CLASS from OPERADOR..tb_segregacao where CLASS2='" & Replace(Trim(Nnull(tb1.Fields("IMO" & i), 1)), ".", "") & "' AND SEGREGACAO in ('2','3','4') "
                    
                    Set tbBusca = New ADODB.Recordset
                    tbBusca.Open SQl, db, 1, 1
                    While Not tbBusca.EOF
                        For j = 1 To 4
                            SQl = "Select id_conteiner, yard, imo" & j & " as imo from OPERADOR..vw_invent_sistemas_imo where yard in ('" & PilhaAnt & "','" & PilhaPos & "','" & FiadaAnt & "','" & FiadaPos & "') and replace(imo" & j & ",'.','') ='" & tbBusca!Class & "' and id_conteiner <>'" & Id_Conteiner & "'"
                            Set TBIMO = New ADODB.Recordset
                            TBIMO.Open SQl, db, 1, 1
                            While Not TBIMO.EOF
                                Verifica_Regras_Seg_Imo = Verifica_Regras_Seg_Imo & Mensagem(2) & " Conteiner: " & TBIMO!Id_Conteiner & " Posicao: " & TBIMO!Yard & " Imo:" & tbBusca!Class & vbCrLf
                                TBIMO.MoveNext
                            Wend
                            TBIMO.Close
                        Next j
                        tbBusca.MoveNext
                    Wend
                    tbBusca.Close
                    
                    
                    
                    
                    
                    Dim PilhaAntAcima As String
                    Dim PilhaAntAbaixo As String
                    Dim PilhaPosAcima As String
                    Dim PilhaPosAbaixo As String
                    
                    Dim FiadaAntAcima As String
                    Dim FiadaAntAbaixo As String
                    Dim FiadaPosAcima As String
                    Dim FiadaPosAbaixo As String
                    
                    
                    If k > 1 Then
                        PilhaAntAcima = Mid$(Yard, 1, 3) & Format(MColP(k - 1), "00") & Trim(Str(Val(Right(Yard, 1)) + 1))
                    
                        PilhaAntAbaixo = Mid$(Yard, 1, 3) & Format(MColP(k - 1), "00") & Trim(Str(Val(Right(Yard, 1)) - 1))
                    Else
                        PilhaAntAcima = ""
                    
                        PilhaAntAbaixo = ""
                    End If
                    
                    PilhaPosAcima = Mid$(Yard, 1, 3) & Format(MColP(k + 1), "00") & Trim(Str(Val(Right(Yard, 1)) + 1))
                    PilhaPosAbaixo = Mid$(Yard, 1, 3) & Format(MColP(k + 1), "00") & Trim(Str(Val(Right(Yard, 1)) - 1))
                    
                    
                    FiadaAntAcima = Mid$(Yard, 1, 1) & Format(Val(Mid$(Yard, 2, 2)) - 1, "00") & Format(MColP(k), "00") & Trim(Str(Val(Right(Yard, 1)) + 1))
                    FiadaAntAbaixo = Mid$(Yard, 1, 1) & Format(Val(Mid$(Yard, 2, 2)) - 1, "00") & Format(MColP(k), "00") & Trim(Str(Val(Right(Yard, 1)) - 1))
                    
                    FiadaPosAcima = Mid$(Yard, 1, 1) & Format(Val(Mid$(Yard, 2, 2)) + 1, "00") & Format(MColP(k), "00") & Trim(Str(Val(Right(Yard, 1)) + 1))
                    FiadaPosAbaixo = Mid$(Yard, 1, 1) & Format(Val(Mid$(Yard, 2, 2)) + 1, "00") & Format(MColP(k), "00") & Trim(Str(Val(Right(Yard, 1)) - 1))
                    
                    
                    
                    
                    SQl = "Select CLASS2 AS CLASS from OPERADOR..tb_segregacao where CLASS1='" & Replace(Trim(Nnull(tb1.Fields("IMO" & i), 1)), ".", "") & "' AND SEGREGACAO in ('3','4') "
                    SQl = SQl & " UNION "
                    SQl = SQl & "Select CLASS1 AS CLASS from OPERADOR..tb_segregacao where CLASS2='" & Replace(Trim(Nnull(tb1.Fields("IMO" & i), 1)), ".", "") & "' AND SEGREGACAO in ('3','4') "
                    
                    
                    Set tbBusca = New ADODB.Recordset
                    tbBusca.Open SQl, db, 1, 1
                    While Not tbBusca.EOF
                        For j = 1 To 4
                            SQl = "Select id_conteiner, yard, imo" & j & " as imo from OPERADOR..vw_invent_sistemas_imo where yard in ('" & PilhaAntAcima & "','" & PilhaAntAbaixo & "','" & PilhaPosAcima & "','" & PilhaPosAbaixo & "','" & FiadaAntAcima & "','" & FiadaAntAbaixo & "','" & FiadaPosAcima & "','" & FiadaPosAbaixo & "') and replace(imo" & j & ",'.','') ='" & tbBusca!Class & "' and id_conteiner <>'" & Id_Conteiner & "'"
                            Set TBIMO = New ADODB.Recordset
                            TBIMO.Open SQl, db, 1, 1
                            While Not TBIMO.EOF
                                Verifica_Regras_Seg_Imo = Verifica_Regras_Seg_Imo & Mensagem(3) & " Conteiner: " & TBIMO!Id_Conteiner & " Posicao: " & TBIMO!Yard & " Imo:" & tbBusca!Class & vbCrLf
                                TBIMO.MoveNext
                            Wend
                            TBIMO.Close
                        Next j
                        tbBusca.MoveNext
                    Wend
                    tbBusca.Close
                            
                    Dim PilhaAntAcima2 As String
                    Dim PilhaAntAbaixo2 As String
                    Dim PilhaPosAcima2 As String
                    Dim PilhaPosAbaixo2 As String
                    
                    Dim FiadaAntAcima2 As String
                    Dim FiadaAntAbaixo2 As String
                    Dim FiadaPosAcima2 As String
                    Dim FiadaPosAbaixo2 As String
                    
                    
                    If k > 1 Then
                        PilhaAntAcima2 = Mid$(Yard, 1, 3) & Format(MColP(k - 1), "00") & Trim(Str(Val(Right(Yard, 1)) + 2))
                        PilhaAntAbaixo2 = Mid$(Yard, 1, 3) & Format(MColP(k - 1), "00") & Trim(Str(Val(Right(Yard, 1)) - 2))
                    Else
                        PilhaAntAcima2 = ""
                        PilhaAntAbaixo2 = ""
                    End If
                    
                    PilhaPosAcima2 = Mid$(Yard, 1, 3) & Format(MColP(k + 1), "00") & Trim(Str(Val(Right(Yard, 1)) + 2))
                    PilhaPosAbaixo2 = Mid$(Yard, 1, 3) & Format(MColP(k + 1), "00") & Trim(Str(Val(Right(Yard, 1)) - 2))
                    
                    
                    FiadaAntAcima2 = Mid$(Yard, 1, 1) & Format(Val(Mid$(Yard, 2, 2)) - 1, "00") & Format(MColP(k), "00") & Trim(Str(Val(Right(Yard, 1)) + 2))
                    FiadaAntAbaixo2 = Mid$(Yard, 1, 1) & Format(Val(Mid$(Yard, 2, 2)) - 1, "00") & Format(MColP(k), "00") & Trim(Str(Val(Right(Yard, 1)) - 2))
                    
                    FiadaPosAcima2 = Mid$(Yard, 1, 1) & Format(Val(Mid$(Yard, 2, 2)) + 1, "00") & Format(MColP(k), "00") & Trim(Str(Val(Right(Yard, 1)) + 2))
                    FiadaPosAbaixo2 = Mid$(Yard, 1, 1) & Format(Val(Mid$(Yard, 2, 2)) + 1, "00") & Format(MColP(k), "00") & Trim(Str(Val(Right(Yard, 1)) - 2))
                    
                    
                    SQl = "Select CLASS2 AS CLASS from OPERADOR..tb_segregacao where CLASS1='" & Replace(Trim(Nnull(tb1.Fields("IMO" & i), 1)), ".", "") & "' AND SEGREGACAO in ('4') "
                    SQl = SQl & " UNION "
                    SQl = SQl & "Select CLASS1 AS CLASS from OPERADOR..tb_segregacao where CLASS2='" & Replace(Trim(Nnull(tb1.Fields("IMO" & i), 1)), ".", "") & "' AND SEGREGACAO in ('4') "
                    
                    Set tbBusca = New ADODB.Recordset
                    tbBusca.Open SQl, db, 1, 1
                    While Not tbBusca.EOF
                        For j = 1 To 4
                            SQl = "Select id_conteiner, yard, imo" & j & " as imo from OPERADOR..vw_invent_sistemas_imo where yard in ('" & PilhaAntAcima2 & "','" & PilhaAntAbaixo2 & "','" & PilhaPosAcima2 & "','" & PilhaPosAbaixo2 & "','" & FiadaAntAcima2 & "','" & FiadaAntAbaixo2 & "','" & FiadaPosAcima2 & "','" & FiadaPosAbaixo2 & "') and replace(imo" & j & ",'.','') ='" & tbBusca!Class & "' and id_conteiner <>'" & Id_Conteiner & "'"
                            Set TBIMO = New ADODB.Recordset
                            TBIMO.Open SQl, db, 1, 1
                            While Not TBIMO.EOF
                                Verifica_Regras_Seg_Imo = Verifica_Regras_Seg_Imo & Mensagem(4) & " Conteiner: " & TBIMO!Id_Conteiner & " Posicao: " & TBIMO!Yard & " Imo:" & tbBusca!class2 & vbCrLf
                                TBIMO.MoveNext
                            Wend
                            TBIMO.Close
                        Next j
                        tbBusca.MoveNext
                    Wend
                    tbBusca.Close
                    
                End If
            End If
        Next i
    End If
    tb1.Close
End If

End Function

Function Verifica_Regras_Seg_Imo_Delta(Id_Conteiner As String, Yard As String, Optional Patio As Integer = 1) As String

Verifica_Regras_Seg_Imo_Delta = ""

If Patio = 7 Then Patio = 1


Dim EspacoMin(1 To 4) As Single
SQl = "Select * from sgipa..dte_tb_parametros "
Set tb1 = New ADODB.Recordset
tb1.Open SQl, db, 1, 1
EspacoMin(1) = Nnull(tb1!DIST_IMO_SEGREG1, 0)
EspacoMin(2) = Nnull(tb1!DIST_IMO_SEGREG2, 0)
EspacoMin(3) = Nnull(tb1!DIST_IMO_SEGREG3, 0)
EspacoMin(4) = Nnull(tb1!DIST_IMO_SEGREG4, 0)
tb1.Close

Dim Mensagem(1 To 4) As String
Mensagem(1) = "Espaço inferior a " & EspacoMin(1) & "m "
Mensagem(2) = "Espaço inferior a " & EspacoMin(2) & "m "
Mensagem(3) = "Espaço inferior a " & EspacoMin(3) & "m "
Mensagem(4) = "Espaço inferior a " & EspacoMin(4) & "m "



'Posicao a ser colocada é de pilha ?
    
SQl = "SELECT ID_CONTEINER,IMO1,IMO2,IMO3,IMO4, "
SQl = SQl & " ID_CONTEINER_D, YARD_D,IMO1_D,IMO2_D,IMO3_D,IMO4_D, DIST_DELTA "
If Patio = 3 Then
    SQl = SQl & " FROM OPERADOR..VW_DIST_IMO_CLIA "
Else
    SQl = SQl & " FROM OPERADOR..VW_DIST_IMO "
End If
SQl = SQl & " WHERE ID_CONTEINER='" & Id_Conteiner & "' ORDER BY DIST_DELTA"
    
Set tb1 = New ADODB.Recordset
tb1.Open SQl, db, 1, 1
If Not tb1.EOF Then
    While Not tb1.EOF
        'MsgBox tb1!ID_CONTEINER_D
        For i = 1 To 4
            For j = 1 To 4
                If Trim(Nnull(tb1.Fields("IMO" & i), 1)) <> "" And Trim(Nnull(tb1.Fields("IMO" & j & "_D"), 1)) <> "" Then
                    If Trim(Nnull(tb1.Fields("IMO" & i), 1)) <> "0" And Trim(Nnull(tb1.Fields("IMO" & j & "_D"), 1)) <> "0" Then
                        SQl = "Select SEGREGACAO from OPERADOR..tb_segregacao where CLASS1='" & Replace(Trim(Nnull(tb1.Fields("IMO" & i), 1)), ".", "") & "'"
                        SQl = SQl & " AND CLASS2='" & Replace(Trim(Nnull(tb1.Fields("IMO" & j & "_D"), 1)), ".", "") & "'"
                        SQl = SQl & " AND SEGREGACAO IN ('1','2','3','4') "
                        SQl = SQl & " UNION "
                        SQl = SQl & " Select SEGREGACAO from OPERADOR..tb_segregacao where CLASS2='" & Replace(Trim(Nnull(tb1.Fields("IMO" & i), 1)), ".", "") & "'"
                        SQl = SQl & " AND CLASS1='" & Replace(Trim(Nnull(tb1.Fields("IMO" & j & "_D"), 1)), ".", "") & "'"
                        SQl = SQl & " AND SEGREGACAO IN ('1','2','3','4') "
                        Set tbBusca = New ADODB.Recordset
                        tbBusca.Open SQl, db, 1, 1
                        If Not tbBusca.EOF Then
                            If Nnull(tb1!dist_delta, 0) < EspacoMin(Val(tbBusca!SEGREGACAO)) Then
                                Verifica_Regras_Seg_Imo_Delta = Verifica_Regras_Seg_Imo_Delta & Mensagem(Val(tbBusca!SEGREGACAO)) & " Conteiner: " & tb1!ID_CONTEINER_D & " Posicao: " & tb1!Yard_D & " Imo:" & Trim(Nnull(tb1.Fields("IMO" & j & "_D"), 1)) & vbCrLf
                            End If
                        End If
                    End If
                End If
            Next j
        Next i
        tb1.MoveNext
    Wend
End If

End Function

Function Verifica_Regras_Seg_Imo_DeltaCS(Marcante As String, Yard As String, Autonum_Armazem As Integer) As String

Verifica_Regras_Seg_Imo_DeltaCS = ""

Dim EspacoMin(1 To 4) As Single
SQl = "Select * from sgipa..dte_tb_parametros "
Set tb1 = New ADODB.Recordset
tb1.Open SQl, db, 1, 1
EspacoMin(1) = Nnull(tb1!DIST_IMO_SEGREG_CS1, 0)
EspacoMin(2) = Nnull(tb1!DIST_IMO_SEGREG_CS2, 0)
EspacoMin(3) = Nnull(tb1!DIST_IMO_SEGREG_CS3, 0)
EspacoMin(4) = Nnull(tb1!DIST_IMO_SEGREG_CS4, 0)
tb1.Close

Dim Mensagem(1 To 4) As String
Mensagem(1) = "Espaço inferior a " & EspacoMin(1) & "m "
Mensagem(2) = "Espaço inferior a " & EspacoMin(2) & "m "
Mensagem(3) = "Espaço inferior a " & EspacoMin(3) & "m "
Mensagem(4) = "Espaço inferior a " & EspacoMin(4) & "m "



'Posicao a ser colocada é de pilha ?
    
'Sql = "SELECT MARCANTE,IMO1, "
'Sql = Sql & " MARCANTE_D, YARD_D,IMO1_D, DIST_DELTA "
'Sql = Sql & " FROM OPERADOR..VW_DIST_IMO_CS "
'Sql = Sql & " WHERE MARCANTE=" & Val(Marcante) & " ORDER BY DIST_DELTA"
    
SQl = "SELECT * FROM operador..fnVW_DIST_IMO_CS (" & Val(Marcante) & "," & Autonum_Armazem & ")"
    
Set tb1 = New ADODB.Recordset
tb1.Open SQl, db, 1, 1
If Not tb1.EOF Then
    While Not tb1.EOF
        
        For i = 1 To 1
            For j = 1 To 1
                If Trim(Nnull(tb1.Fields("IMO" & i), 1)) <> "" And Trim(Nnull(tb1.Fields("IMO" & j & "_D"), 1)) <> "" Then
                    If Trim(Nnull(tb1.Fields("IMO" & i), 1)) <> "0" And Trim(Nnull(tb1.Fields("IMO" & j & "_D"), 1)) <> "0" Then
                        SQl = "Select SEGREGACAO from OPERADOR..tb_segregacao where CLASS1='" & Replace(Trim(Nnull(tb1.Fields("IMO" & i), 1)), ".", "") & "'"
                        SQl = SQl & " AND CLASS2='" & Replace(Trim(Nnull(tb1.Fields("IMO" & j & "_D"), 1)), ".", "") & "'"
                        SQl = SQl & " AND SEGREGACAO IN ('1','2','3','4') "
                        SQl = SQl & " UNION "
                        SQl = SQl & " Select SEGREGACAO from OPERADOR..tb_segregacao where CLASS2='" & Replace(Trim(Nnull(tb1.Fields("IMO" & i), 1)), ".", "") & "'"
                        SQl = SQl & " AND CLASS1='" & Replace(Trim(Nnull(tb1.Fields("IMO" & j & "_D"), 1)), ".", "") & "'"
                        SQl = SQl & " AND SEGREGACAO IN ('1','2','3','4') "
                        Set tbBusca = New ADODB.Recordset
                        tbBusca.Open SQl, db, 1, 1
                        If Not tbBusca.EOF Then
                            If Nnull(tb1!dist_delta, 0) < EspacoMin(Val(tbBusca!SEGREGACAO)) Then
                                Verifica_Regras_Seg_Imo_DeltaCS = Verifica_Regras_Seg_Imo_DeltaCS & Mensagem(Val(tbBusca!SEGREGACAO)) & " Marcante: " & tb1!Marcante_D & " Posicao: " & tb1!Yard_D & " Imo:" & Trim(Nnull(tb1.Fields("IMO" & j & "_D"), 1)) & vbCrLf
                            End If
                        End If
                    End If
                End If
            Next j
        Next i
        tb1.MoveNext
    Wend
End If

End Function



Function Valida_Aloca_Imo(Id_Conteiner As String, Yard As String, Tamanho As Integer, Optional Patio As Integer = 1) As String
Dim SQl As String
Valida_Aloca_Imo = ""


If Patio = 7 Then Patio = 1

Dim Valida As Boolean
Valida = False
SQl = "SELECT valida FROM OPERADOR..TB_YARD WHERE PATIO=" & Patio & " AND YARD='" & Yard & "'"
Set tb1 = New ADODB.Recordset
tb1.Open SQl, db, 1, 1
If Not tb1.EOF Then
    If tb1!Valida = 0 Then Valida = False
    If tb1!Valida = 1 Then Valida = True
End If
tb1.Close

SQl = "SELECT ID_CONTEINER,IMO1,IMO2,IMO3,IMO4 FROM OPERADOR..VW_INVENT_SISTEMAS_IMO "
SQl = SQl & " WHERE ID_CONTEINER='" & Id_Conteiner & "'"
Set tb1 = New ADODB.Recordset 'O Conteiner é IMO ?
tb1.Open SQl, db, 1, 1
If Not tb1.EOF Then
    'Verifica se pode pelas regras de IMO ou nao
    
    
    If Valida Then
        'A QUADRA TODA
        SQl = " SELECT YARD FROM INTELOPER..TB_IP_ATRIBUTO_YARD"
        SQl = SQl & " WHERE AUTONUM_ATR=21 AND VLR_ATRIB=1 AND STATUS_ATRIB='R'"
        SQl = SQl & " AND DT_INI_ATR < GETDATE() AND (DT_FIM_ATR IS NULL OR DT_FIM_ATR>GETDATE())"
        SQl = SQl & " AND LEN(YARD)=1"
        If Col_Tam_Pos_Yard = 6 Then
            SQl = SQl & " AND SUBSTRING(YARD,1,1)='" & Left(Yard, 1) & "'"
        Else
            SQl = SQl & " AND SUBSTRING(YARD,1,2)='" & Left(Yard, 2) & "'"
        End If
        SQl = SQl & " Union"
        SQl = SQl & " SELECT YARD FROM INTELOPER..TB_IP_ATRIBUTO_YARD"
        SQl = SQl & " WHERE AUTONUM_ATR=21 AND VLR_ATRIB=0 AND STATUS_ATRIB='E'"
        SQl = SQl & " AND DT_INI_ATR < GETDATE() AND (DT_FIM_ATR IS NULL OR DT_FIM_ATR>GETDATE())"
        If Col_Tam_Pos_Yard = 6 Then
            SQl = SQl & " AND LEN(YARD)=1"
            SQl = SQl & " AND SUBSTRING(YARD,1,1)='" & Left(Yard, 1) & "'"
        Else
            SQl = SQl & " AND LEN(YARD)=2"
            SQl = SQl & " AND SUBSTRING(YARD,1,2)='" & Left(Yard, 2) & "'"
        End If
        
        Set TBV = New ADODB.Recordset
        TBV.Open SQl, db, 1, 1
        If Not TBV.EOF Then
            Valida_Aloca_Imo = "Quadra não permite conteiner IMO"
        End If
        
        If Valida_Aloca_Imo = "" Then
            SQl = " SELECT YARD FROM INTELOPER..TB_IP_ATRIBUTO_YARD"
            SQl = SQl & " WHERE AUTONUM_ATR=21 AND VLR_ATRIB=1 AND STATUS_ATRIB='R'"
            SQl = SQl & " AND DT_INI_ATR < GETDATE() AND (DT_FIM_ATR IS NULL OR DT_FIM_ATR>GETDATE())"
            If Tamanho = 20 Then
                SQl = SQl & " AND SUBSTRING(YARD,1,5)='" & Mid$(Yard, 1, 5) & "'"
            Else
                SQl = SQl & " AND (SUBSTRING(YARD,1,5)='" & Mid$(Yard, 1, 1) & Format(Val(Mid$(Yard, 2, 2)) + 1, "00") & Mid$(Yard, 4, 2) & "' OR SUBSTRING(YARD,1,5)='" & Mid$(Yard, 1, 1) & Format(Val(Mid$(Yard, 2, 2)) - 1, "00") & Mid$(Yard, 4, 2) & "')"
            End If
            SQl = SQl & " Union"
            SQl = SQl & " SELECT YARD FROM INTELOPER..TB_IP_ATRIBUTO_YARD"
            SQl = SQl & " WHERE AUTONUM_ATR=21 AND VLR_ATRIB=0 AND STATUS_ATRIB='E'"
            SQl = SQl & " AND DT_INI_ATR < GETDATE() AND (DT_FIM_ATR IS NULL OR DT_FIM_ATR>GETDATE())"
            If Tamanho = 20 Then
                If Col_Tam_Pos_Yard = 6 Then
                    SQl = SQl & " AND SUBSTRING(YARD,1,5)='" & Mid$(Yard, 1, 5) & "'"
                Else
                    SQl = SQl & " AND SUBSTRING(YARD,1,6)='" & Mid$(Yard, 1, 6) & "'"
                End If
            Else
                If Col_Tam_Pos_Yard = 6 Then
                    SQl = SQl & " AND (SUBSTRING(YARD,1,5)='" & Mid$(Yard, 1, 1) & Format(Val(Mid$(Yard, 2, 2)) + 1, "00") & Mid$(Yard, 4, 2) & "' OR SUBSTRING(YARD,1,5)='" & Mid$(Yard, 1, 1) & Format(Val(Mid$(Yard, 2, 2)) - 1, "00") & Mid$(Yard, 4, 2) & "')"
                Else
                    SQl = SQl & " AND (SUBSTRING(YARD,1,6)='" & Mid$(Yard, 1, 2) & Format(Val(Mid$(Yard, 3, 2)) + 1, "00") & Mid$(Yard, 5, 2) & "' OR SUBSTRING(YARD,1,6)='" & Mid$(Yard, 1, 2) & Format(Val(Mid$(Yard, 3, 2)) - 1, "00") & Mid$(Yard, 5, 2) & "')"
                End If
            End If
            Set TBV = New ADODB.Recordset
            TBV.Open SQl, db, 1, 1
            If Not TBV.EOF Then
                Valida_Aloca_Imo = "Pilha não permite conteiner IMO"
            End If
            
        End If
    
    Else
    
        SQl = " SELECT YARD FROM INTELOPER..TB_IP_ATRIBUTO_YARD"
        SQl = SQl & " WHERE AUTONUM_ATR=21 AND VLR_ATRIB=1 AND STATUS_ATRIB='R'"
        SQl = SQl & " AND DT_INI_ATR < GETDATE() AND (DT_FIM_ATR IS NULL OR DT_FIM_ATR>GETDATE())"
        SQl = SQl & " AND YARD='" & Yard & "'"
        SQl = SQl & " Union"
        SQl = SQl & " SELECT YARD FROM INTELOPER..TB_IP_ATRIBUTO_YARD"
        SQl = SQl & " WHERE AUTONUM_ATR=21 AND VLR_ATRIB=0 AND STATUS_ATRIB='E'"
        SQl = SQl & " AND DT_INI_ATR < GETDATE() AND (DT_FIM_ATR IS NULL OR DT_FIM_ATR>GETDATE())"
        SQl = SQl & " AND YARD='" & Yard & "'"
        Set TBV = New ADODB.Recordset
        TBV.Open SQl, db, 1, 1
        If Not TBV.EOF Then
            Valida_Aloca_Imo = "Heap/Rua não permite conteiner IMO"
        End If
    
    End If
    
    
    If Valida_Aloca_Imo = "" Then

        For i = 1 To 4
            If Trim(Nnull(tb1.Fields("IMO" & i), 1)) <> "" Then
                If Trim(Nnull(tb1.Fields("IMO" & i), 1)) <> "0" Then
                    'Verifica se pode pela classe IMO
                    If Valida Then
                    
                        'A QUADRA TODA
                        SQl = " SELECT YARD FROM INTELOPER..TB_IP_ATRIBUTO_YARD"
                        SQl = SQl & " WHERE AUTONUM_ATR=12 AND VLR_ATRIB='" & Trim(Nnull(tb1.Fields("IMO" & i), 1)) & "' AND STATUS_ATRIB='R'"
                        SQl = SQl & " AND DT_INI_ATR < GETDATE() AND (DT_FIM_ATR IS NULL OR DT_FIM_ATR>GETDATE())"
                        If Col_Tam_Pos_Yard = 6 Then
                            SQl = SQl & " AND LEN(YARD)=1"
                            SQl = SQl & " AND SUBSTRING(YARD,1,1)='" & Left(Yard, 1) & "'"
                        Else
                            SQl = SQl & " AND LEN(YARD)=2"
                            SQl = SQl & " AND SUBSTRING(YARD,1,2)='" & Left(Yard, 2) & "'"
                        End If
                        SQl = SQl & " Union"
                        SQl = SQl & " SELECT YARD FROM INTELOPER..TB_IP_ATRIBUTO_YARD"
                        SQl = SQl & " WHERE AUTONUM_ATR=12 AND VLR_ATRIB<>'" & Trim(Nnull(tb1.Fields("IMO" & i), 1)) & "' AND STATUS_ATRIB='E'"
                        SQl = SQl & " AND DT_INI_ATR < GETDATE() AND (DT_FIM_ATR IS NULL OR DT_FIM_ATR>GETDATE())"
                        If Col_Tam_Pos_Yard = 6 Then
                            SQl = SQl & " AND LEN(YARD)=1"
                            SQl = SQl & " AND SUBSTRING(YARD,1,1)='" & Left(Yard, 1) & "'"
                        Else
                            SQl = SQl & " AND LEN(YARD)=2"
                            SQl = SQl & " AND SUBSTRING(YARD,1,2)='" & Left(Yard, 2) & "'"
                        End If
                        Set TBV = New ADODB.Recordset
                        TBV.Open SQl, db, 1, 1
                        If Not TBV.EOF Then
                            Valida_Aloca_Imo = "Quadra não permite classe IMO " & Trim(Nnull(tb1.Fields("IMO" & i), 1))
                        End If
                        'PILHA
                        
                        SQl = " SELECT YARD FROM INTELOPER..TB_IP_ATRIBUTO_YARD"
                        SQl = SQl & " WHERE AUTONUM_ATR=12 AND VLR_ATRIB='" & Trim(Nnull(tb1.Fields("IMO" & i), 1)) & "' AND STATUS_ATRIB='R'"
                        SQl = SQl & " AND DT_INI_ATR < GETDATE() AND (DT_FIM_ATR IS NULL OR DT_FIM_ATR>GETDATE())"
                        If Tamanho = 20 Then
                            If Col_Tam_Pos_Yard = 6 Then
                                SQl = SQl & " AND SUBSTRING(YARD,1,5)='" & Mid$(Yard, 1, 5) & "'"
                            Else
                                SQl = SQl & " AND SUBSTRING(YARD,1,6)='" & Mid$(Yard, 1, 6) & "'"
                            End If
                        Else
                            If Col_Tam_Pos_Yard = 6 Then
                                SQl = SQl & " AND (SUBSTRING(YARD,1,5)='" & Mid$(Yard, 1, 1) & Format(Val(Mid$(Yard, 2, 2)) + 1, "00") & Mid$(Yard, 4, 2) & "' OR SUBSTRING(YARD,1,5)='" & Mid$(Yard, 1, 1) & Format(Val(Mid$(Yard, 2, 2)) - 1, "00") & Mid$(Yard, 4, 2) & "')"
                            Else
                                SQl = SQl & " AND (SUBSTRING(YARD,1,6)='" & Mid$(Yard, 1, 2) & Format(Val(Mid$(Yard, 3, 2)) + 1, "00") & Mid$(Yard, 5, 2) & "' OR SUBSTRING(YARD,1,6)='" & Mid$(Yard, 1, 2) & Format(Val(Mid$(Yard, 3, 2)) - 1, "00") & Mid$(Yard, 5, 2) & "')"
                            End If
                        End If
                        SQl = SQl & " Union"
                        SQl = SQl & " SELECT YARD FROM INTELOPER..TB_IP_ATRIBUTO_YARD"
                        SQl = SQl & " WHERE AUTONUM_ATR=12 AND VLR_ATRIB<>'" & Trim(Nnull(tb1.Fields("IMO" & i), 1)) & "' AND STATUS_ATRIB='E'"
                        SQl = SQl & " AND DT_INI_ATR < GETDATE() AND (DT_FIM_ATR IS NULL OR DT_FIM_ATR>GETDATE())"
                        If Tamanho = 20 Then
                            If Col_Tam_Pos_Yard = 6 Then
                                SQl = SQl & " AND SUBSTRING(YARD,1,5)='" & Mid$(Yard, 1, 5) & "'"
                            Else
                                SQl = SQl & " AND SUBSTRING(YARD,1,6)='" & Mid$(Yard, 1, 6) & "'"
                            End If
                        Else
                            If Col_Tam_Pos_Yard = 6 Then
                                SQl = SQl & " AND (SUBSTRING(YARD,1,5)='" & Mid$(Yard, 1, 1) & Format(Val(Mid$(Yard, 2, 2)) + 1, "00") & Mid$(Yard, 4, 2) & "' OR SUBSTRING(YARD,1,5)='" & Mid$(Yard, 1, 1) & Format(Val(Mid$(Yard, 2, 2)) - 1, "00") & Mid$(Yard, 4, 2) & "')"
                            Else
                                SQl = SQl & " AND (SUBSTRING(YARD,1,6)='" & Mid$(Yard, 1, 2) & Format(Val(Mid$(Yard, 3, 2)) + 1, "00") & Mid$(Yard, 5, 2) & "' OR SUBSTRING(YARD,1,6)='" & Mid$(Yard, 1, 2) & Format(Val(Mid$(Yard, 3, 2)) - 1, "00") & Mid$(Yard, 5, 2) & "')"
                            End If
                        End If
                        Set TBV = New ADODB.Recordset
                        TBV.Open SQl, db, 1, 1
                        If Not TBV.EOF Then
                            Valida_Aloca_Imo = "Pilha não permite Classe IMO " & Trim(Nnull(tb1.Fields("IMO" & i), 1))
                        End If
                        TBV.Close
                        
                                    
                    Else
                        
                        
                        SQl = " SELECT YARD FROM INTELOPER..TB_IP_ATRIBUTO_YARD"
                        SQl = SQl & " WHERE AUTONUM_ATR=12 AND VLR_ATRIB='" & Trim(Nnull(tb1.Fields("IMO" & i), 1)) & "' AND STATUS_ATRIB='R'"
                        SQl = SQl & " AND DT_INI_ATR < GETDATE() AND (DT_FIM_ATR IS NULL OR DT_FIM_ATR>GETDATE())"
                        SQl = SQl & " AND YARD='" & Yard & "'"
                        SQl = SQl & " Union"
                        SQl = SQl & " SELECT YARD FROM INTELOPER..TB_IP_ATRIBUTO_YARD"
                        SQl = SQl & " WHERE AUTONUM_ATR=12 AND VLR_ATRIB<>'" & Trim(Nnull(tb1.Fields("IMO" & i), 1)) & "' AND STATUS_ATRIB='E'"
                        SQl = SQl & " AND DT_INI_ATR < GETDATE() AND (DT_FIM_ATR IS NULL OR DT_FIM_ATR>GETDATE())"
                        SQl = SQl & " AND YARD='" & Yard & "'"
                        Set TBV = New ADODB.Recordset
                        TBV.Open SQl, db, 1, 1
                        If Not TBV.EOF Then
                            Valida_Aloca_Imo = "Heap/Rua não permite Classe IMO " & Trim(Nnull(tb1.Fields("IMO" & i), 1))
                        End If
                    
                    End If
        
                    
                End If
            End If
        Next
        
    End If

End If

End Function

