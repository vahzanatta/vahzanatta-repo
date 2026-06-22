Attribute VB_Name = "RelatorioProcessos"
Option Explicit

Sub ImportarRelatorios()
    Dim arquivoResg As Variant
    Dim arquivoMov  As Variant

    arquivoResg = Application.GetOpenFilename( _
        FileFilter:="Excel (*.xlsx; *.xlsm; *.xls), *.xlsx; *.xlsm; *.xls", _
        Title:="Selecione o relatório de RESGATES")
    If arquivoResg = False Then Exit Sub

    arquivoMov = Application.GetOpenFilename( _
        FileFilter:="Excel (*.xlsx; *.xlsm; *.xls), *.xlsx; *.xlsm; *.xls", _
        Title:="Selecione o relatório de MOVIMENTO")
    If arquivoMov = False Then Exit Sub

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts  = False

    Call ImportarParaAba(arquivoResg, "RESGATES")
    Call ImportarParaAba(arquivoMov, "MOVIMENTO")

    Application.ScreenUpdating = True
    Application.DisplayAlerts  = True

    Call AnalisarProcessos
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts  = True
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub ImportarParaAba(arquivo As Variant, nomeAba As String)
    Dim wbOrigem  As Workbook
    Dim wsOrigem  As Worksheet
    Dim wsDestino As Worksheet
    Dim lastRowSrc As Long
    Dim lastCol    As Long
    Dim arrData    As Variant

    Set wbOrigem  = Workbooks.Open(Filename:=arquivo, ReadOnly:=True, UpdateLinks:=False)
    Set wsOrigem  = wbOrigem.Worksheets(1)
    Set wsDestino = ThisWorkbook.Worksheets(nomeAba)

    lastRowSrc = wsOrigem.Cells(wsOrigem.Rows.Count, "A").End(xlUp).Row

    If lastRowSrc < 2 Then
        wbOrigem.Close SaveChanges:=False
        MsgBox "O arquivo selecionado para '" & nomeAba & "' não contém dados a partir da linha 2.", vbExclamation, "Arquivo vazio"
        Exit Sub
    End If

    Dim lastRowDest As Long
    lastRowDest = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row
    If lastRowDest >= 2 Then wsDestino.Rows("2:" & lastRowDest).Delete

    lastCol = wsOrigem.Cells(1, wsOrigem.Columns.Count).End(xlToLeft).Column
    arrData = wsOrigem.Range(wsOrigem.Cells(2, 1), wsOrigem.Cells(lastRowSrc, lastCol)).Value

    wsDestino.Range("A2").Resize(lastRowSrc - 1, lastCol).Value = arrData

    wbOrigem.Close SaveChanges:=False
End Sub

Sub AnalisarProcessos()
    Dim wsProc  As Worksheet
    Dim wsResg  As Worksheet
    Dim wsMov   As Worksheet

    On Error Resume Next
    Set wsProc = ThisWorkbook.Worksheets("PROCESSOS")
    Set wsResg = ThisWorkbook.Worksheets("RESGATES")
    Set wsMov  = ThisWorkbook.Worksheets("MOVIMENTO")
    On Error GoTo ErrHandler

    If wsProc Is Nothing Then
        MsgBox "Aba 'PROCESSOS' não encontrada.", vbCritical, "Erro"
        Exit Sub
    End If
    If wsResg Is Nothing Then
        MsgBox "Aba 'RESGATES' não encontrada.", vbCritical, "Erro"
        Exit Sub
    End If
    If wsMov Is Nothing Then
        MsgBox "Aba 'MOVIMENTO' não encontrada.", vbCritical, "Erro"
        Exit Sub
    End If

    Dim lastRowProc As Long
    Dim lastRowResg As Long
    Dim lastRowMov  As Long

    lastRowProc = wsProc.Cells(wsProc.Rows.Count, "A").End(xlUp).Row
    lastRowResg = wsResg.Cells(wsResg.Rows.Count, "A").End(xlUp).Row
    lastRowMov  = wsMov.Cells(wsMov.Rows.Count, "A").End(xlUp).Row

    If lastRowProc < 2 Then
        MsgBox "Nenhum dado encontrado na aba 'PROCESSOS' (a partir da linha 2).", vbExclamation, "Sem dados"
        Exit Sub
    End If
    If lastRowResg < 2 Then
        MsgBox "Nenhum dado encontrado na aba 'RESGATES' (a partir da linha 2).", vbExclamation, "Sem dados"
        Exit Sub
    End If
    If lastRowMov < 2 Then
        MsgBox "Nenhum dado encontrado na aba 'MOVIMENTO' (a partir da linha 2).", vbExclamation, "Sem dados"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation    = xlCalculationManual

    Dim arrProc As Variant
    Dim arrResg As Variant
    Dim arrMov  As Variant

    arrProc = wsProc.Range("A2:H" & lastRowProc).Value
    arrResg = wsResg.Range("A2:K" & lastRowResg).Value
    arrMov  = wsMov.Range("A2:K" & lastRowMov).Value

    Dim totalRows   As Long
    Dim totalResg   As Long
    Dim totalMov    As Long
    Dim totalNaoLoc As Long

    totalRows   = lastRowProc - 1
    totalResg   = 0
    totalMov    = 0
    totalNaoLoc = 0

    Dim arrOut() As Variant
    ReDim arrOut(1 To totalRows, 1 To 3)

    Dim i As Long, j As Long
    Dim processo    As String
    Dim valorRec    As Double
    Dim creditoUni  As Double
    Dim dataCredito As Variant
    Dim textoResg   As String
    Dim valorResg   As Double
    Dim dataResg    As Variant
    Dim textoMov    As String
    Dim valorMov    As Double
    Dim dataMov     As Variant
    Dim localizadoEm    As String
    Dim correspondencia As String
    Dim retorno         As String
    Dim difDias As Long
    Dim found   As Boolean

    For i = 1 To totalRows
        processo   = Trim(CStr(arrProc(i, 1)))
        dataCredito = arrProc(i, 2)
        valorRec   = 0
        creditoUni = 0

        On Error Resume Next
        If arrProc(i, 3) <> "" Then valorRec   = CDbl(arrProc(i, 3))
        If arrProc(i, 4) <> "" Then creditoUni = CDbl(arrProc(i, 4))
        On Error GoTo ErrHandler

        localizadoEm    = ""
        correspondencia = ""
        retorno         = ""
        found           = False

        ' Estágio 1 — RESGATES
        ' Passo 1a: por número do processo
        If processo <> "" Then
            For j = 1 To lastRowResg - 1
                textoResg = Trim(CStr(arrResg(j, 11)))
                If InStr(1, textoResg, processo, vbTextCompare) > 0 Then
                    dataResg        = arrResg(j, 4)
                    localizadoEm    = "Resgate"
                    correspondencia = "Número do processo"
                    retorno         = RetornoComDivergencia(dataCredito, dataResg, "")
                    found = True
                    Exit For
                End If
            Next j
        End If

        ' Passo 1b: por valor no RESGATES
        If Not found Then
            For j = 1 To lastRowResg - 1
                On Error Resume Next
                valorResg = CDbl(arrResg(j, 8))
                On Error GoTo ErrHandler

                If valorRec <> 0 And valorResg = valorRec Then
                    dataResg        = arrResg(j, 4)
                    localizadoEm    = "Resgate"
                    correspondencia = "Valor recuperado"
                    retorno         = RetornoComDivergencia(dataCredito, dataResg, "")
                    found = True
                    Exit For
                ElseIf creditoUni <> 0 And valorResg = creditoUni Then
                    dataResg        = arrResg(j, 4)
                    localizadoEm    = "Resgate"
                    correspondencia = "Crédito único"
                    retorno         = RetornoComDivergencia(dataCredito, dataResg, "")
                    found = True
                    Exit For
                End If
            Next j
        End If

        ' Estágio 2 — MOVIMENTO
        If Not found Then
            For j = 1 To lastRowMov - 1
                On Error Resume Next
                valorMov = CDbl(arrMov(j, 8))
                On Error GoTo ErrHandler

                If valorRec <> 0 And valorMov = valorRec Then
                    dataMov         = arrMov(j, 4)
                    localizadoEm    = "Movimento"
                    correspondencia = "Valor recuperado"
                    retorno         = RetornoComDivergencia(dataCredito, dataMov, "Localizado — sem número de processo vinculado")
                    found = True
                    Exit For
                ElseIf creditoUni <> 0 And valorMov = creditoUni Then
                    dataMov         = arrMov(j, 4)
                    localizadoEm    = "Movimento"
                    correspondencia = "Crédito único"
                    retorno         = RetornoComDivergencia(dataCredito, dataMov, "Localizado — sem número de processo vinculado")
                    found = True
                    Exit For
                End If
            Next j
        End If

        If Not found Then
            localizadoEm    = ""
            correspondencia = ""
            retorno         = "Não localizado"
            totalNaoLoc     = totalNaoLoc + 1
        ElseIf localizadoEm = "Resgate" Then
            totalResg = totalResg + 1
        ElseIf localizadoEm = "Movimento" Then
            totalMov = totalMov + 1
        End If

        arrOut(i, 1) = localizadoEm
        arrOut(i, 2) = correspondencia
        arrOut(i, 3) = retorno
    Next i

    wsProc.Cells(1, 6).Value = "Localizado em"
    wsProc.Cells(1, 7).Value = "Correspondência"
    wsProc.Cells(1, 8).Value = "Retorno"

    wsProc.Range("F2:H" & lastRowProc).Value = arrOut

    Application.ScreenUpdating = True
    Application.Calculation    = xlCalculationAutomatic

    MsgBox "Concluído!" & vbCrLf & vbCrLf & _
           "Linhas analisadas:        " & totalRows & vbCrLf & _
           "Localizados em Resgate:   " & totalResg & vbCrLf & _
           "Localizados em Movimento: " & totalMov & vbCrLf & _
           "Não localizados:          " & totalNaoLoc, _
           vbInformation, "Processamento Concluído"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation    = xlCalculationAutomatic
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro"
End Sub

Sub LimparAnalise()
    Dim ws      As Worksheet
    Dim lastRow As Long

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets("PROCESSOS")

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 1 Then Exit Sub

    ws.Range("F1:H" & lastRow).ClearContents
    Exit Sub

ErrHandler:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro"
End Sub

Private Function RetornoComDivergencia(dataCredito As Variant, dataDoc As Variant, prefixo As String) As String
    Dim base As String
    Dim difDias As Long

    If prefixo = "" Then
        base = "Localizado"
    Else
        base = prefixo
    End If

    If IsDate(dataCredito) And CStr(dataCredito) <> "" And _
       IsDate(dataDoc) And CStr(dataDoc) <> "" Then
        difDias = Abs(CLng(CDate(dataDoc)) - CLng(CDate(dataCredito)))
        If difDias = 0 Then
            RetornoComDivergencia = base
        Else
            If prefixo = "" Then
                RetornoComDivergencia = base & " — divergência de " & difDias & " dias"
            Else
                RetornoComDivergencia = base & ", divergência de " & difDias & " dias"
            End If
        End If
    Else
        RetornoComDivergencia = base
    End If
End Function
