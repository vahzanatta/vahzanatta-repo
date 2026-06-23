Attribute VB_Name = "RelatorioProcessos"
Option Explicit

Sub ImportarRelatorios()
    Dim arquivoResg As Variant
    Dim arquivoMov  As Variant

    MsgBox "Selecione o relat?rio de RESGATES." & vbCrLf & _
           "(Exportado do SAP ? cont?m os resgates do per?odo)", _
           vbInformation, "Passo 1 de 2"
    arquivoResg = Application.GetOpenFilename( _
        FileFilter:="Excel (*.xlsx; *.xlsm; *.xls), *.xlsx; *.xlsm; *.xls", _
        Title:="Passo 1/2 ? Selecione o relat?rio de RESGATES")
    If arquivoResg = False Then Exit Sub

    MsgBox "Selecione o relat?rio de MOVIMENTO." & vbCrLf & _
           "(Exportado do SAP ? cont?m o movimento financeiro do per?odo)", _
           vbInformation, "Passo 2 de 2"
    arquivoMov = Application.GetOpenFilename( _
        FileFilter:="Excel (*.xlsx; *.xlsm; *.xls), *.xlsx; *.xlsm; *.xls", _
        Title:="Passo 2/2 ? Selecione o relat?rio de MOVIMENTO")
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

    On Error Resume Next
    Set wsDestino = ThisWorkbook.Worksheets(nomeAba)
    On Error GoTo 0
    If wsDestino Is Nothing Then
        Set wsDestino = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDestino.Name = nomeAba
    End If

    lastRowSrc = wsOrigem.Cells(wsOrigem.Rows.Count, "A").End(xlUp).Row

    If lastRowSrc < 2 Then
        wbOrigem.Close SaveChanges:=False
        MsgBox "O arquivo selecionado para '" & nomeAba & "' n?o cont?m dados a partir da linha 2.", vbExclamation, "Arquivo vazio"
        Exit Sub
    End If

    Dim lastRowDest As Long
    lastRowDest = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row
    If lastRowDest >= 1 Then wsDestino.Rows("1:" & lastRowDest).Delete

    lastCol = wsOrigem.Cells(1, wsOrigem.Columns.Count).End(xlToLeft).Column
    arrData = wsOrigem.Range(wsOrigem.Cells(1, 1), wsOrigem.Cells(lastRowSrc, lastCol)).Value

    wsDestino.Range("A1").Resize(lastRowSrc, lastCol).Value = arrData

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
        MsgBox "Aba 'PROCESSOS' n?o encontrada.", vbCritical, "Erro"
        Exit Sub
    End If
    If wsResg Is Nothing Then
        MsgBox "Aba 'RESGATES' n?o encontrada.", vbCritical, "Erro"
        Exit Sub
    End If
    If wsMov Is Nothing Then
        MsgBox "Aba 'MOVIMENTO' n?o encontrada.", vbCritical, "Erro"
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
    ReDim arrOut(1 To totalRows, 1 To 5)

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
    Dim texto           As String
    Dim valor           As Double
    Dim difDias As Long
    Dim found          As Boolean
    Dim valorEsperado  As Double

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
        texto           = ""
        valor           = 0
        found           = False

        If processo <> "" Then
            valorEsperado = 0
            If valorRec <> 0 Then
                valorEsperado = -valorRec
            ElseIf creditoUni <> 0 Then
                valorEsperado = -creditoUni
            End If

            If valorEsperado <> 0 Then
                For j = 1 To lastRowResg - 1
                    If IsDate(dataCredito) And IsDate(arrResg(j, 4)) Then
                        If CDate(arrResg(j, 4)) < CDate(dataCredito) Then GoTo ProximoResg1a
                    End If
                    textoResg = Trim(CStr(arrResg(j, 11)))
                    If InStr(1, textoResg, processo, vbTextCompare) > 0 Then
                        If CDbl(arrResg(j, 8)) = valorEsperado Then
                            dataResg        = arrResg(j, 4)
                            localizadoEm    = "Resgate"
                            correspondencia = "N?mero do processo"
                            retorno         = RetornoComDivergencia(dataCredito, dataResg, "")
                            texto           = textoResg
                            valor           = CDbl(arrResg(j, 8))
                            found = True
                            Exit For
                        End If
                    End If
ProximoResg1a:
                Next j
            End If

            If Not found Then
                For j = 1 To lastRowResg - 1
                    If IsDate(dataCredito) And IsDate(arrResg(j, 4)) Then
                        If CDate(arrResg(j, 4)) < CDate(dataCredito) Then GoTo ProximoResg1aDiverg
                    End If
                    textoResg = Trim(CStr(arrResg(j, 11)))
                    If InStr(1, textoResg, processo, vbTextCompare) > 0 Then
                        dataResg        = arrResg(j, 4)
                        localizadoEm    = "Resgate"
                        correspondencia = "N?mero do processo"
                        retorno         = RetornoComDivergencia(dataCredito, dataResg, "")
                        texto           = textoResg
                        valor           = CDbl(arrResg(j, 8))
                        If valorEsperado <> 0 And valor <> valorEsperado Then
                            retorno = retorno & " - divergencia de valor (" & Format(valor, "#,##0.00") & ")"
                        End If
                        found = True
                        Exit For
                    End If
ProximoResg1aDiverg:
                Next j
            End If
        End If

        If Not found Then
            For j = 1 To lastRowResg - 1
                If IsDate(dataCredito) And IsDate(arrResg(j, 4)) Then
                    If CDate(arrResg(j, 4)) < CDate(dataCredito) Then GoTo ProximoResg1b
                End If
                On Error Resume Next
                valorResg = CDbl(arrResg(j, 8))
                On Error GoTo ErrHandler

                If valorRec <> 0 And valorResg = -valorRec Then
                    dataResg        = arrResg(j, 4)
                    localizadoEm    = "Resgate"
                    correspondencia = "Valor recuperado"
                    retorno         = RetornoComDivergencia(dataCredito, dataResg, "")
                    texto           = Trim(CStr(arrResg(j, 11)))
                    valor           = valorResg
                    found = True
                    Exit For
                ElseIf creditoUni <> 0 And valorResg = -creditoUni Then
                    dataResg        = arrResg(j, 4)
                    localizadoEm    = "Resgate"
                    correspondencia = "Cr?dito ?nico"
                    retorno         = RetornoComDivergencia(dataCredito, dataResg, "")
                    texto           = Trim(CStr(arrResg(j, 11)))
                    valor           = valorResg
                    found = True
                    Exit For
                End If
ProximoResg1b:
            Next j
        End If

        If Not found Then
            For j = 1 To lastRowMov - 1
                If IsDate(dataCredito) And IsDate(arrMov(j, 4)) Then
                    If CDate(arrMov(j, 4)) < CDate(dataCredito) Then GoTo ProximoMov
                End If

                On Error Resume Next
                valorMov = CDbl(arrMov(j, 8))
                On Error GoTo ErrHandler

                If valorRec <> 0 And valorMov = valorRec Then
                    dataMov         = arrMov(j, 4)
                    localizadoEm    = "Movimento"
                    correspondencia = "Valor recuperado"
                    retorno         = RetornoComDivergencia(dataCredito, dataMov, "Localizado ? sem n?mero de processo vinculado")
                    texto           = Trim(CStr(arrMov(j, 11)))
                    valor           = valorMov
                    found = True
                    Exit For
                ElseIf creditoUni <> 0 And valorMov = creditoUni Then
                    dataMov         = arrMov(j, 4)
                    localizadoEm    = "Movimento"
                    correspondencia = "Cr?dito ?nico"
                    retorno         = RetornoComDivergencia(dataCredito, dataMov, "Localizado ? sem n?mero de processo vinculado")
                    texto           = Trim(CStr(arrMov(j, 11)))
                    valor           = valorMov
                    found = True
                    Exit For
                End If
ProximoMov:
            Next j
        End If

        If Not found Then
            localizadoEm    = ""
            correspondencia = ""
            retorno         = "N?o localizado"
            totalNaoLoc     = totalNaoLoc + 1
        ElseIf localizadoEm = "Resgate" Then
            totalResg = totalResg + 1
        ElseIf localizadoEm = "Movimento" Then
            totalMov = totalMov + 1
        End If

        arrOut(i, 1) = localizadoEm
        arrOut(i, 2) = correspondencia
        arrOut(i, 3) = retorno
        arrOut(i, 4) = texto
        arrOut(i, 5) = valor
    Next i

    wsProc.Cells(1, 6).Value = "Localizado em"
    wsProc.Cells(1, 7).Value = "Correspond?ncia"
    wsProc.Cells(1, 8).Value = "Retorno"
    wsProc.Cells(1, 9).Value = "Texto"
    wsProc.Cells(1, 10).Value = "Valor"

    wsProc.Range("F2:J" & lastRowProc).Value = arrOut

    Application.ScreenUpdating = True
    Application.Calculation    = xlCalculationAutomatic

    MsgBox "Conclu?do!" & vbCrLf & vbCrLf & _
           "Linhas analisadas:        " & totalRows & vbCrLf & _
           "Localizados em Resgate:   " & totalResg & vbCrLf & _
           "Localizados em Movimento: " & totalMov & vbCrLf & _
           "N?o localizados:          " & totalNaoLoc, _
           vbInformation, "Processamento Conclu?do"
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

    ws.Range("F1:J" & lastRow).ClearContents
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

    If IsDate(dataCredito) And IsDate(dataDoc) Then
        difDias = Abs(DateDiff("d", CDate(dataCredito), CDate(dataDoc)))
        If difDias = 0 Then
            RetornoComDivergencia = base
        Else
            If prefixo = "" Then
                RetornoComDivergencia = base & " ? diverg?ncia de " & difDias & " dias"
            Else
                RetornoComDivergencia = base & ", diverg?ncia de " & difDias & " dias"
            End If
        End If
    Else
        RetornoComDivergencia = base
    End If
End Function
