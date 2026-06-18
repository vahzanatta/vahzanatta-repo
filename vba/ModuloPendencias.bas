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

Sub ProcessarPendencias()
    Dim wsSAP       As Worksheet
    Dim wsBase      As Worksheet
    Dim lastRowSAP  As Long
    Dim lastRowBase As Long
    Dim i           As Long

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

    lastRowSAP  = wsSAP.Cells(wsSAP.Rows.Count, "B").End(xlUp).Row
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
        If Trim(CStr(wsSAP.Cells(i, 6).Value)) = "" Then
            wsSAP.Rows(i).Delete
        End If
    Next i

    ' [Etapas 3 e 4 serão adicionadas na próxima tarefa]
    Application.ScreenUpdating = True
    Application.Calculation    = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation    = xlCalculationAutomatic
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro"
End Sub
