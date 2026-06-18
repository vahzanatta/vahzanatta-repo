Attribute VB_Name = "ModuloPendencias"
Option Explicit

' Returns 0 if rule is eliminated by any mismatch, or 1-3 for the number of fields matched.
' Scoring: +1 for conta match (or blank conta), +1 for começa em match, +1 for contém match.
' A non-blank field that does not match immediately eliminates the rule (returns 0).
Function CalcularScore(contaSAP As String, textoSAP As String, _
                       contaBase As String, comecaEm As String, _
                       contem As String) As Integer
    Dim score As Integer
    score = 0

    ' Conta do Razão: blank = wildcard (matches any); non-blank must match exactly
    If contaBase <> "" Then
        If StrComp(contaSAP, contaBase, vbTextCompare) <> 0 Then
            CalcularScore = 0
            Exit Function
        End If
    End If
    score = score + 1

    ' Começa em: only evaluated when non-blank
    If comecaEm <> "" Then
        If StrComp(Left(textoSAP, Len(comecaEm)), comecaEm, vbTextCompare) <> 0 Then
            CalcularScore = 0
            Exit Function
        End If
        score = score + 1
    End If

    ' Contém: only evaluated when non-blank
    If contem <> "" Then
        If InStr(1, textoSAP, contem, vbTextCompare) = 0 Then
            CalcularScore = 0
            Exit Function
        End If
        score = score + 1
    End If

    CalcularScore = score
End Function

Sub LimparResultados()
    Dim ws      As Worksheet
    Dim lastRow As Long

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets("SAP")

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If lastRow < 1 Then Exit Sub

    ws.Range("K1:M" & lastRow).ClearContents

    MsgBox "Colunas K:M limpas com sucesso.", vbInformation, "Limpeza Concluída"
    Exit Sub

ErrHandler:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro"
End Sub
