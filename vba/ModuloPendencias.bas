Attribute VB_Name = "ModuloPendencias"
Option Explicit

' Retorna 0 se a regra for eliminada por incompatibilidade, ou 1–4 conforme os campos correspondidos.
' Pontuação: +1 conta (ou em branco), +1 começa em, +1 contém, +1 regex.
' Um campo não-vazio que não corresponde elimina a regra imediatamente (retorna 0).
Function CalcularScore(contaSAP As String, textoSAP As String, _
                       contaBase As String, comecaEm As String, _
                       contem As String, regexStr As String) As Integer
    Dim score As Integer
    score = 0

    ' Conta do Razão: em branco = coringa (qualquer conta); não-vazio deve corresponder exatamente
    If contaBase <> "" Then
        If StrComp(contaSAP, contaBase, vbTextCompare) <> 0 Then
            CalcularScore = 0
            Exit Function
        End If
    End If
    score = score + 1

    ' Começa em: avaliado apenas quando não-vazio
    If comecaEm <> "" Then
        If StrComp(Left(textoSAP, Len(comecaEm)), comecaEm, vbTextCompare) <> 0 Then
            CalcularScore = 0
            Exit Function
        End If
        score = score + 1
    End If

    ' Contém: avaliado apenas quando não-vazio
    If contem <> "" Then
        If InStr(1, textoSAP, contem, vbTextCompare) = 0 Then
            CalcularScore = 0
            Exit Function
        End If
        score = score + 1
    End If

    ' Regex: avaliado apenas quando não-vazio (usa VBScript.RegExp, case-insensitive)
    If regexStr <> "" Then
        Dim re As Object
        Set re = CreateObject("VBScript.RegExp")
        re.IgnoreCase = True
        re.Global     = False
        re.Pattern    = regexStr
        If Not re.Test(textoSAP) Then
            CalcularScore = 0
            Set re = Nothing
            Exit Function
        End If
        Set re = Nothing
        score = score + 1
    End If

    CalcularScore = score
End Function

Sub ImportarSAP()
    Dim wbOrigem  As Workbook
    Dim wsOrigem  As Worksheet
    Dim wsDestino As Worksheet
    Dim arquivo   As Variant
    Dim lastRowSrc As Long
    Dim lastRowDest As Long
    Dim lastCol   As Long
    Dim arrData   As Variant

    arquivo = Application.GetOpenFilename( _
        FileFilter:="Excel (*.xlsx; *.xlsm; *.xls), *.xlsx; *.xlsm; *.xls", _
        Title:="Selecione o relatório exportado do SAP")

    If arquivo = False Then Exit Sub

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts  = False

    Set wbOrigem = Workbooks.Open(Filename:=arquivo, ReadOnly:=True, UpdateLinks:=False)
    Set wsOrigem = wbOrigem.Worksheets(1)

    lastRowSrc = wsOrigem.Cells(wsOrigem.Rows.Count, "A").End(xlUp).Row

    If lastRowSrc < 2 Then
        wbOrigem.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.DisplayAlerts  = True
        MsgBox "O arquivo selecionado não contém dados a partir da linha 2.", vbExclamation, "Arquivo vazio"
        Exit Sub
    End If

    Set wsDestino = ThisWorkbook.Worksheets("SAP")

    lastRowDest = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row
    If lastRowDest >= 2 Then wsDestino.Rows("2:" & lastRowDest).Delete

    lastCol = wsOrigem.Cells(1, wsOrigem.Columns.Count).End(xlToLeft).Column
    arrData = wsOrigem.Range(wsOrigem.Cells(2, 1), wsOrigem.Cells(lastRowSrc, lastCol)).Value

    wsDestino.Range("A2").Resize(lastRowSrc - 1, lastCol).Value = arrData

    wbOrigem.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.DisplayAlerts  = True

    Call ProcessarPendencias
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts  = True
    If Not wbOrigem Is Nothing Then wbOrigem.Close SaveChanges:=False
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro"
End Sub

Sub LimparResultados()
    Dim ws      As Worksheet
    Dim lastRow As Long

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets("SAP")

    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    If lastRow < 1 Then Exit Sub

    ws.Range("L1:N" & lastRow).ClearContents

    MsgBox "Colunas L:N limpas com sucesso.", vbInformation, "Limpeza Concluída"
    Exit Sub

ErrHandler:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro"
End Sub

Sub ProcessarPendencias()
    Dim wsSAP       As Worksheet
    Dim wsBase      As Worksheet
    Dim lastRowSAP  As Long
    Dim lastRowBase As Long
    Dim i As Long, j As Long

    ' --- Etapa 1: Validação ---
    On Error Resume Next
    Set wsSAP  = ThisWorkbook.Worksheets("SAP")
    Set wsBase = ThisWorkbook.Worksheets("BASE")
    On Error GoTo ErrHandler

    If wsSAP Is Nothing Then
        MsgBox "Aba 'SAP' não encontrada.", vbCritical, "Erro"
        Exit Sub
    End If
    If wsBase Is Nothing Then
        MsgBox "Aba 'BASE' não encontrada.", vbCritical, "Erro"
        Exit Sub
    End If

    lastRowSAP  = wsSAP.Cells(wsSAP.Rows.Count, "C").End(xlUp).Row
    lastRowBase = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row

    If lastRowSAP < 2 Then
        MsgBox "Nenhum dado encontrado na aba 'SAP' (a partir da linha 2).", vbExclamation, "Sem dados"
        Exit Sub
    End If
    If lastRowBase < 2 Then
        MsgBox "Nenhuma regra encontrada na aba 'BASE' (a partir da linha 2).", vbExclamation, "BASE vazia"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation    = xlCalculationManual

    ' --- Etapa 2: Exclusão de linhas inválidas (de baixo para cima) ---
    For i = lastRowSAP To 2 Step -1
        If Trim(CStr(wsSAP.Cells(i, 7).Value)) = "" Then
            wsSAP.Rows(i).Delete
        End If
    Next i

    ' Recalcular após exclusões
    lastRowSAP = wsSAP.Cells(wsSAP.Rows.Count, "C").End(xlUp).Row
    If lastRowSAP < 2 Then
        Application.ScreenUpdating = True
        Application.Calculation    = xlCalculationAutomatic
        MsgBox "Todas as linhas foram excluídas (Tipo de documento em branco).", vbExclamation, "Sem dados"
        Exit Sub
    End If

    ' --- Etapa 3: Carregamento em arrays ---
    Dim arrSAP()  As Variant
    Dim arrBase() As Variant
    arrSAP  = wsSAP.Range("A2:K" & lastRowSAP).Value   ' (1 To n, 1 To 11)
    arrBase = wsBase.Range("A2:E" & lastRowBase).Value  ' (1 To m, 1 To 5)

    ' --- Etapa 4: Processamento ---
    Dim totalRows  As Long
    Dim totalAtrib As Long
    Dim totalNaoId As Long
    Dim dataBase   As Date
    Dim contaSAP   As String
    Dim textoSAP   As String
    Dim dataLanc   As Variant
    Dim diasAberto As Variant
    Dim contaBase  As String
    Dim comecaEm   As String
    Dim contem     As String
    Dim regexStr   As String
    Dim areaBase   As String
    Dim bestScore  As Integer
    Dim bestArea   As String
    Dim scoreAtual As Integer

    dataBase   = Date
    totalRows  = lastRowSAP - 1
    totalAtrib = 0
    totalNaoId = 0

    Dim arrOut() As Variant
    ReDim arrOut(1 To totalRows, 1 To 3)

    For i = 1 To totalRows
        contaSAP = Trim(CStr(arrSAP(i, 3)))   ' col C — Conta do Razão
        textoSAP = Trim(CStr(arrSAP(i, 11)))  ' col K — Texto
        dataLanc = arrSAP(i, 5)               ' col E — Data de lançamento

        ' Dias em Aberto
        If IsDate(dataLanc) And CStr(dataLanc) <> "" Then
            diasAberto = CLng(dataBase) - CLng(CDate(dataLanc))
        Else
            diasAberto = "—"
        End If

        ' Busca a regra com maior pontuação na BASE
        bestScore = 0
        bestArea  = ""

        For j = 1 To lastRowBase - 1
            contaBase = Trim(CStr(arrBase(j, 1)))  ' col A — Conta do Razão
            comecaEm  = Trim(CStr(arrBase(j, 2)))  ' col B — Começa em
            contem    = Trim(CStr(arrBase(j, 3)))  ' col C — Contém
            regexStr  = Trim(CStr(arrBase(j, 4)))  ' col D — Regex
            areaBase  = Trim(CStr(arrBase(j, 5)))  ' col E — Área Responsável

            ' Ignora regras sem Área Responsável definida
            If areaBase = "" Then GoTo ProximaRegra

            scoreAtual = CalcularScore(contaSAP, textoSAP, contaBase, comecaEm, contem, regexStr)

            ' Maior score vence; empate mantém a primeira regra na ordem
            If scoreAtual > bestScore Then
                bestScore = scoreAtual
                bestArea  = areaBase
            End If

ProximaRegra:
        Next j

        ' Preenche array de saída
        arrOut(i, 1) = dataBase
        arrOut(i, 2) = diasAberto

        If bestScore > 0 Then
            arrOut(i, 3) = bestArea
            totalAtrib = totalAtrib + 1
        Else
            arrOut(i, 3) = ChrW(9888) & " Não identificado"
            totalNaoId = totalNaoId + 1
        End If
    Next i

    ' Cabeçalhos das colunas de saída
    wsSAP.Cells(1, 12).Value = "Data Base"
    wsSAP.Cells(1, 13).Value = "Dias em Aberto"
    wsSAP.Cells(1, 14).Value = "Área Responsável"

    ' Escrita em lote
    wsSAP.Range("L2:N" & lastRowSAP).Value = arrOut

    ' Formata Data Base como data curta
    wsSAP.Range("L2:L" & lastRowSAP).NumberFormat = "dd/mm/yyyy"

    Application.ScreenUpdating = True
    Application.Calculation    = xlCalculationAutomatic

    MsgBox "Concluído!" & vbCrLf & vbCrLf & _
           "Linhas analisadas:   " & totalRows & vbCrLf & _
           "Atribuídos:          " & totalAtrib & vbCrLf & _
           "Não identificados:   " & totalNaoId, _
           vbInformation, "Processamento Concluído"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation    = xlCalculationAutomatic
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro"
End Sub
