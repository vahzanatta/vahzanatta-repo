# Relatório de Pendências — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Create a VBA module that processes SAP FBL3H open items data, matches each row to a responsible area via a rules table, deletes invalid rows, and adds three calculated columns (Data Base, Dias em Aberto, Área Responsável).

**Architecture:** Single VBA module (`ModuloPendencias.bas`) with one helper function (`CalcularScore`) and two public macros (`ProcessarPendencias`, `LimparResultados`). All heavy processing runs in-memory using Variant arrays to handle 100k–250k rows efficiently. Results are written back to the sheet in a single batch operation.

**Tech Stack:** VBA (Excel), no external libraries.

## Global Constraints

- Sheet names are fixed: `"SAP"` and `"BASE"` (exact, case-sensitive as Excel matches them)
- SAP sheet: row 1 = headers, data from row 2; columns A:J are input, K:M are output
- BASE sheet: row 1 = headers, rules from row 2; columns A:D (Conta do Razão, Começa em, Contém, Área Responsável)
- All text comparisons must be case-insensitive (use `vbTextCompare`)
- No colored cell formatting anywhere in the output
- No cell-by-cell reads during processing — all data loaded into arrays before the main loop

---

## File Structure

| File | Action | Responsibility |
|------|--------|----------------|
| `vba/ModuloPendencias.bas` | Create | Full VBA module: `CalcularScore`, `ProcessarPendencias`, `LimparResultados` |

---

### Task 1: Create module and implement CalcularScore

**Files:**
- Create: `vba/ModuloPendencias.bas`

**Interfaces:**
- Produces: `Function CalcularScore(contaSAP As String, textoSAP As String, contaBase As String, comecaEm As String, contem As String) As Integer`
  - Returns `0` if the rule is eliminated by a mismatch
  - Returns `1–3` based on how many fields matched (each matched field = +1 point)

- [ ] **Step 1: Create `vba/ModuloPendencias.bas` with the module header and CalcularScore**

```vba
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
```

- [ ] **Step 2: Import the module into Excel and verify CalcularScore in the Immediate Window**

To import: open the VBE (Alt+F11) → File → Import File → select `ModuloPendencias.bas`.

Open the Immediate Window (Ctrl+G) and run each line — confirm the value after the `?`:

```
? CalcularScore("1001", "PIX TRANSF JOEL", "1001", "PIX", "TRANSF")
```
Expected: `3`

```
? CalcularScore("1001", "PIX TRANSF JOEL", "", "", "TRANSF")
```
Expected: `2`  (blank conta = wildcard, contém matches)

```
? CalcularScore("1001", "PIX TRANSF JOEL", "9999", "", "")
```
Expected: `0`  (conta mismatch → eliminated)

```
? CalcularScore("1001", "PIX TRANSF JOEL", "1001", "TED", "")
```
Expected: `0`  (começa em mismatch → eliminated)

```
? CalcularScore("1001", "pix transf joel", "1001", "PIX", "TRANSF")
```
Expected: `3`  (case-insensitive: lowercase SAP text matches uppercase rule)

```
? CalcularScore("1001", "PIX TRANSF JOEL", "1001", "", "")
```
Expected: `1`  (only conta matched, both text fields blank in rule)

- [ ] **Step 3: Commit**

```
git add vba/ModuloPendencias.bas
git commit -m "feat: add CalcularScore helper function"
```

---

### Task 2: Implement LimparResultados

**Files:**
- Modify: `vba/ModuloPendencias.bas`

**Interfaces:**
- Consumes: sheet `"SAP"`, columns K:M
- Produces: public macro `LimparResultados` — clears K1:M(lastRow) including header row

- [ ] **Step 1: Append LimparResultados to `vba/ModuloPendencias.bas`**

```vba
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
```

- [ ] **Step 2: Verify manually**

1. In the SAP sheet, type any value in cells K2, L2 and M2
2. Run `LimparResultados` (Alt+F8 → LimparResultados → Run)
3. Confirm K2, L2 and M2 are now empty and the success message appeared

- [ ] **Step 3: Commit**

```
git add vba/ModuloPendencias.bas
git commit -m "feat: add LimparResultados macro"
```

---

### Task 3: Implement ProcessarPendencias — validation and row deletion (Etapas 1 and 2)

**Files:**
- Modify: `vba/ModuloPendencias.bas`

**Interfaces:**
- Consumes: sheets `"SAP"` and `"BASE"`
- Produces: partial `Sub ProcessarPendencias()` that validates both sheets exist with data, then deletes every row in SAP where column F (Tipo de documento) is blank, iterating bottom-up

- [ ] **Step 1: Set up test data in the SAP sheet**

Populate rows 1–4 of the SAP sheet as follows (F3 must be blank):

| Row | A | B | C | D | E | F | G | H | I | J |
|-----|---|---|---|---|---|---|---|---|---|---|
| 1 | Empresa | Conta do Razão | Data de documento | Data de lançamento | Nº documento | Tipo de documento | Valor em moeda da empresa | Chave de lançamento | Atribuição | Texto |
| 2 | 1000 | 1001 | 01/01/2026 | 01/01/2026 | 100 | KR | 500 | 40 | A1 | PIX TRANSF JOEL |
| 3 | 1000 | 1001 | 02/01/2026 | 02/01/2026 | 101 | *(blank)* | 200 | 40 | A2 | TED TRANSF |
| 4 | 1000 | 1002 | 03/01/2026 | 03/01/2026 | 102 | SA | 300 | 50 | A3 | SISPAG FOLHA |

- [ ] **Step 2: Append ProcessarPendencias (partial) to `vba/ModuloPendencias.bas`**

```vba
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
```

- [ ] **Step 3: Run ProcessarPendencias and verify row deletion**

Run `ProcessarPendencias` (Alt+F8 → ProcessarPendencias → Run). Confirm:
- The row with "TED TRANSF" (blank Tipo de documento) was deleted
- Two rows remain: "PIX TRANSF JOEL" (now row 2) and "SISPAG FOLHA" (now row 3)
- Rows with data were untouched

- [ ] **Step 4: Commit**

```
git add vba/ModuloPendencias.bas
git commit -m "feat: add ProcessarPendencias validation and row deletion"
```

---

### Task 4: Implement ProcessarPendencias — array loading, matching and output (Etapas 3 and 4)

**Files:**
- Modify: `vba/ModuloPendencias.bas`

**Interfaces:**
- Consumes: `CalcularScore(contaSAP As String, textoSAP As String, contaBase As String, comecaEm As String, contem As String) As Integer`
- Produces: columns K (Data Base as Date), L (Dias em Aberto as Long or "—"), M (Área Responsável as String) written in one batch; headers written to K1:M1; MsgBox with linhas analisadas / atribuídos / não identificados

- [ ] **Step 1: Set up test data in the BASE sheet**

Populate the BASE sheet:

| Row | A | B | C | D |
|-----|---|---|---|---|
| 1 | Conta do Razão | Começa em | Contém | Área Responsável |
| 2 | 1001 | PIX | TRANSF | Financeiro |
| 3 | 1001 | | SISPAG | TI |
| 4 | | | FOLHA | RH |
| 5 | 1002 | | | Contabilidade |

- [ ] **Step 2: Replace the full ProcessarPendencias in `vba/ModuloPendencias.bas`**

Remove the partial stub from Task 3 and replace it with the complete implementation:

```vba
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

    ' Recalculate after deletions
    lastRowSAP = wsSAP.Cells(wsSAP.Rows.Count, "B").End(xlUp).Row
    If lastRowSAP < 2 Then
        Application.ScreenUpdating = True
        Application.Calculation    = xlCalculationAutomatic
        MsgBox "Todas as linhas foram excluídas (Tipo de documento em branco).", vbExclamation, "Sem dados"
        Exit Sub
    End If

    ' --- Etapa 3: Carregamento em arrays ---
    Dim arrSAP()  As Variant
    Dim arrBase() As Variant
    arrSAP  = wsSAP.Range("A2:J" & lastRowSAP).Value   ' (1 To n, 1 To 10)
    arrBase = wsBase.Range("A2:D" & lastRowBase).Value  ' (1 To m, 1 To 4)

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
    Dim areaBase   As String
    Dim bestScore  As Integer
    Dim bestArea   As String
    Dim scoreAtual As Integer

    dataBase   = Date
    totalRows  = lastRowSAP - 1   ' rows 2..lastRowSAP → array indices 1..totalRows
    totalAtrib = 0
    totalNaoId = 0

    Dim arrOut() As Variant
    ReDim arrOut(1 To totalRows, 1 To 3)

    For i = 1 To totalRows
        contaSAP = Trim(CStr(arrSAP(i, 2)))   ' col B — Conta do Razão
        textoSAP = Trim(CStr(arrSAP(i, 10)))  ' col J — Texto
        dataLanc = arrSAP(i, 4)               ' col D — Data de lançamento

        ' Dias em Aberto
        If IsDate(dataLanc) And CStr(dataLanc) <> "" Then
            diasAberto = CLng(dataBase) - CLng(CDate(dataLanc))
        Else
            diasAberto = "—"
        End If

        ' Find best matching rule in BASE
        bestScore = 0
        bestArea  = ""

        For j = 1 To lastRowBase - 1
            contaBase = Trim(CStr(arrBase(j, 1)))  ' col A
            comecaEm  = Trim(CStr(arrBase(j, 2)))  ' col B
            contem    = Trim(CStr(arrBase(j, 3)))  ' col C
            areaBase  = Trim(CStr(arrBase(j, 4)))  ' col D

            ' Skip rules with no Área Responsável defined
            If areaBase = "" Then GoTo ProximaRegra

            scoreAtual = CalcularScore(contaSAP, textoSAP, contaBase, comecaEm, contem)

            ' Strict greater-than preserves first-match-wins on tie
            If scoreAtual > bestScore Then
                bestScore = scoreAtual
                bestArea  = areaBase
            End If

ProximaRegra:
        Next j

        ' Fill output row
        arrOut(i, 1) = dataBase
        arrOut(i, 2) = diasAberto

        If bestScore > 0 Then
            arrOut(i, 3) = bestArea
            totalAtrib = totalAtrib + 1
        Else
            arrOut(i, 3) = Chr(9888) & " Não identificado"
            totalNaoId = totalNaoId + 1
        End If
    Next i

    ' Write headers
    wsSAP.Cells(1, 11).Value = "Data Base"
    wsSAP.Cells(1, 12).Value = "Dias em Aberto"
    wsSAP.Cells(1, 13).Value = Chr(193) & "rea Respons" & Chr(225) & "vel"

    ' Batch write output
    wsSAP.Range("K2:M" & lastRowSAP).Value = arrOut

    ' Format Data Base column as short date
    wsSAP.Range("K2:K" & lastRowSAP).NumberFormat = "dd/mm/yyyy"

    Application.ScreenUpdating = True
    Application.Calculation    = xlCalculationAutomatic

    MsgBox "Conclu" & Chr(237) & "do!" & vbCrLf & vbCrLf & _
           "Linhas analisadas:   " & totalRows & vbCrLf & _
           "Atribu" & Chr(237) & "dos:          " & totalAtrib & vbCrLf & _
           "N" & Chr(227) & "o identificados: " & totalNaoId, _
           vbInformation, "Processamento Conclu" & Chr(237) & "do"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation    = xlCalculationAutomatic
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro"
End Sub
```

- [ ] **Step 3: Restore the SAP test data (Task 3 deleted rows) and run ProcessarPendencias**

Re-populate SAP rows 2–4 as in Task 3 Step 1 (with the blank Tipo de documento in row 3). Run `ProcessarPendencias`. Expected results after row deletion and processing:

| SAP row | Conta | Texto | Área Responsável esperada | Score vencedor |
|---------|-------|-------|--------------------------|----------------|
| 2 | 1001 | PIX TRANSF JOEL | Financeiro | 3 (conta + começa + contém) |
| 3 | 1002 | SISPAG FOLHA | RH | 2 (conta branco + contém "FOLHA") |

Verify:
- K2 and K3 show today's date formatted as dd/mm/yyyy
- L2 and L3 show the correct number of days since 01/01/2026 and 03/01/2026
- M2 = `Financeiro`, M3 = `RH`
- MsgBox shows: Linhas analisadas: 2, Atribuídos: 2, Não identificados: 0

- [ ] **Step 4: Test edge case — unmatched row**

Add a row to SAP: Empresa=1000, Conta=9999, Data doc=01/01/2026, Data lanç=01/01/2026, Nº doc=999, Tipo doc=AB, Valor=100, Chave=40, Atrib=ZZ, Texto=DESCONHECIDO.

Run `ProcessarPendencias`. Verify:
- That row's column M shows `⚠ Não identificado`
- MsgBox shows Não identificados: 1

- [ ] **Step 5: Test edge case — invalid Data de lançamento**

Add a row with Tipo doc=KR and Data de lançamento blank. Run `ProcessarPendencias`. Verify column L for that row shows `—`.

- [ ] **Step 6: Commit**

```
git add vba/ModuloPendencias.bas
git commit -m "feat: complete ProcessarPendencias with array processing and batch output"
```

---

## Self-Review

### Spec coverage

| Requirement | Task |
|-------------|------|
| Aba SAP: A:J entrada, K:M saída | Task 4 |
| Aba BASE: 4 colunas (Conta, Começa em, Contém, Área) | Task 4 |
| Linha 1 cabeçalho, dados da linha 2 | Tasks 3–4 |
| Excluir linhas com Tipo de documento em branco | Task 3 |
| Deletar de baixo para cima | Task 3 (Step -1 loop) |
| Carregamento em arrays para performance | Task 4 (Etapa 3) |
| CalcularScore com pontuação 0–3 | Task 1 |
| Conta em branco na BASE = wildcard | Task 1 (CalcularScore) |
| "Começa em" case-insensitive | Task 1 (StrComp vbTextCompare) |
| "Contém" case-insensitive | Task 1 (InStr vbTextCompare) |
| Regra com maior score vence | Task 4 (scoreAtual > bestScore) |
| Empate: primeira na BASE vence | Task 4 (strict `>`, not `>=`) |
| Data Base = data atual | Task 4 (dataBase = Date) |
| Dias em Aberto = Data Base − Data lançamento | Task 4 |
| Data lançamento inválida → "—" | Task 4 (IsDate check) |
| Área Responsável preenchida | Task 4 |
| Sem regra → "⚠ Não identificado" | Task 4 |
| Sem formatação de cor | Task 4 (no .Interior.Color calls) |
| Cabeçalhos K:M escritos automaticamente | Task 4 |
| Mensagem: analisadas, atribuídos, não identificados | Task 4 |
| LimparResultados limpa K:M incluindo cabeçalho | Task 2 |

All requirements covered. ✓

### Placeholder scan

No TBD, TODO, or vague steps. All code blocks are complete. ✓

### Type consistency

- `CalcularScore` defined in Task 1: `(contaSAP As String, textoSAP As String, contaBase As String, comecaEm As String, contem As String) As Integer`
- Called in Task 4: `CalcularScore(contaSAP, textoSAP, contaBase, comecaEm, contem)` — all String variables, return assigned to `scoreAtual As Integer` ✓
