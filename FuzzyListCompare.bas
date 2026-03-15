Attribute VB_Name = "FuzzyListCompare"
Option Explicit

' =============================================================================
' FuzzyListCompare — Row-wise Fuzzy Duplicate Finder
' =============================================================================
' For each item in each list column, finds the best fuzzy match in every other
' list. Outputs pairwise confidence scores and matched text as new columns
' aligned row-by-row with the original data, making it easy to filter and sort.
' Original cells are colour-coded by their best overall match confidence.
' =============================================================================

' ---------------------------------------------------------------------------
' CleanString — strip leading numbering, punctuation, collapse spaces, lowercase
' ---------------------------------------------------------------------------
Private Function CleanString(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim result As String
    Dim startPos As Long

    ' Step 1: Strip leading sequential numbering (e.g. "001.", "22)", "(3)", "7.")
    ' Skip any leading digits, spaces, dots, opening/closing parens until we
    ' hit a letter — then take the rest from that position onward.
    startPos = 1
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z]" Then
            startPos = i
            Exit For
        ElseIf ch Like "[0-9 .()]" Then
            ' Still in the numbering prefix — keep skipping
        Else
            ' Hit a non-numbering, non-letter character — stop stripping
            startPos = i
            Exit For
        End If
    Next i
    ' If the entire string was numbering/digits, keep it all
    If i > Len(s) Then startPos = 1
    s = Mid$(s, startPos)

    ' Step 2: Remove punctuation — keep only letters, digits, and spaces
    result = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9 ]" Then
            result = result & ch
        End If
    Next i

    ' Step 3: Collapse multiple spaces into one
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop

    ' Step 4 & 5: Trim and lowercase
    CleanString = LCase$(Trim$(result))
End Function

' ---------------------------------------------------------------------------
' CharBagSimilarity — character-frequency similarity in [0, 1]
'   score = (2 * common_chars) / (len(A) + len(B))
' ---------------------------------------------------------------------------
Private Function CharBagSimilarity(ByVal a As String, ByVal b As String) As Double
    Dim freqA(0 To 255) As Long
    Dim freqB(0 To 255) As Long
    Dim i As Long
    Dim c As Long
    Dim commonChars As Long
    Dim totalLen As Long

    totalLen = Len(a) + Len(b)
    If totalLen = 0 Then
        CharBagSimilarity = 0#
        Exit Function
    End If

    For i = 1 To Len(a)
        c = Asc(Mid$(a, i, 1))
        freqA(c) = freqA(c) + 1
    Next i

    For i = 1 To Len(b)
        c = Asc(Mid$(b, i, 1))
        freqB(c) = freqB(c) + 1
    Next i

    commonChars = 0
    For i = 0 To 255
        If freqA(i) < freqB(i) Then
            commonChars = commonChars + freqA(i)
        Else
            commonChars = commonChars + freqB(i)
        End If
    Next i

    CharBagSimilarity = (2# * commonChars) / totalLen
End Function

' ---------------------------------------------------------------------------
' ApplyConfidenceColour — set cell fill based on confidence score
' ---------------------------------------------------------------------------
Private Sub ApplyConfidenceColour(ByVal cell As Range, ByVal confidence As Double)
    With cell.Interior
        Select Case True
            Case confidence >= 0.95
                .Color = RGB(0, 128, 0)       ' Dark green  — near-exact
            Case confidence >= 0.8
                .Color = RGB(0, 176, 80)       ' Green       — strong
            Case confidence >= 0.6
                .Color = RGB(146, 208, 80)     ' Yellow-green — moderate
            Case confidence >= 0.4
                .Color = RGB(255, 255, 0)       ' Yellow      — weak
            Case confidence >= 0.2
                .Color = RGB(255, 165, 0)       ' Orange      — very weak
            Case Else
                .Color = RGB(255, 0, 0)         ' Red         — no match
        End Select
    End With
End Sub

' ---------------------------------------------------------------------------
' FuzzyListCompare — main entry point
' ---------------------------------------------------------------------------
Public Sub FuzzyListCompare()

    On Error GoTo ErrHandler

    ' ---- Performance switches ----
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' ==================================================================
    ' 1. DETECT LIST COLUMNS
    ' ==================================================================
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastCol < 1 Then
        MsgBox "No data found on the active sheet.", vbExclamation
        GoTo Cleanup
    End If

    Dim numLists As Long
    Dim colIndices() As Long
    Dim colHeaders() As String
    Dim tempCount As Long

    tempCount = 0
    ReDim colIndices(1 To lastCol)
    ReDim colHeaders(1 To lastCol)

    Dim c As Long
    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(1, c).Value)) <> "" Then
            tempCount = tempCount + 1
            colIndices(tempCount) = c
            colHeaders(tempCount) = CStr(ws.Cells(1, c).Value)
        End If
    Next c
    numLists = tempCount

    If numLists < 2 Then
        MsgBox "Need at least 2 lists (columns with headers) to compare. " & _
               "Found " & numLists & ".", vbExclamation
        GoTo Cleanup
    End If

    ReDim Preserve colIndices(1 To numLists)
    ReDim Preserve colHeaders(1 To numLists)

    ' ==================================================================
    ' 2. READ DATA AND CLEANED VERSIONS
    ' ==================================================================
    Dim listSizes() As Long
    ReDim listSizes(1 To numLists)

    Dim maxRows As Long
    maxRows = 0

    Dim L As Long
    For L = 1 To numLists
        Dim lr As Long
        lr = ws.Cells(ws.Rows.Count, colIndices(L)).End(xlUp).Row
        If lr < 2 Then
            listSizes(L) = 0
        Else
            listSizes(L) = lr - 1
        End If
        If listSizes(L) > maxRows Then maxRows = listSizes(L)
    Next L

    If maxRows = 0 Then
        MsgBox "All lists are empty (headers only).", vbExclamation
        GoTo Cleanup
    End If

    Dim items() As String
    Dim cleaned() As String
    ReDim items(1 To numLists, 1 To maxRows)
    ReDim cleaned(1 To numLists, 1 To maxRows)

    Dim totalItems As Long
    totalItems = 0

    Dim r As Long
    For L = 1 To numLists
        For r = 1 To listSizes(L)
            items(L, r) = CStr(ws.Cells(r + 1, colIndices(L)).Value)
            cleaned(L, r) = CleanString(items(L, r))
            totalItems = totalItems + 1
        Next r
    Next L

    ' ==================================================================
    ' 3. COMPUTE BEST MATCHES (pairwise and overall)
    ' ==================================================================
    ' bestScore(L, r, M) — best similarity of item (L,r) vs any item in list M
    ' bestText(L, r, M)  — original text of the best-matching item from list M
    ' We can't use 3D dynamic arrays directly in VBA, so we flatten:
    '   index = ((L-1)*maxRows + (r-1)) * numLists + M
    ' But for clarity, we'll use a helper offset and two 1D arrays.

    Dim arrSize As Long
    arrSize = numLists * maxRows * numLists
    Dim bScore() As Double
    Dim bText() As String
    ReDim bScore(1 To arrSize)
    ReDim bText(1 To arrSize)

    Dim M As Long, r2 As Long
    Dim conf As Double
    Dim idx As Long
    Dim curBest As Double

    For L = 1 To numLists
        For r = 1 To listSizes(L)
            If cleaned(L, r) <> "" Then
                For M = 1 To numLists
                    If M <> L Then
                        idx = ((L - 1) * maxRows + (r - 1)) * numLists + M
                        curBest = 0#
                        For r2 = 1 To listSizes(M)
                            If cleaned(M, r2) <> "" Then
                                conf = CharBagSimilarity(cleaned(L, r), cleaned(M, r2))
                                If conf > curBest Then
                                    curBest = conf
                                    bScore(idx) = conf
                                    bText(idx) = items(M, r2)
                                    If conf >= 1# Then Exit For
                                End If
                            End If
                        Next r2
                    End If
                Next M
            End If
        Next r
    Next L

    ' Compute best overall score and list name for each (L, r)
    Dim bestOverallScore() As Double
    Dim bestOverallList() As String
    ReDim bestOverallScore(1 To numLists, 1 To maxRows)
    ReDim bestOverallList(1 To numLists, 1 To maxRows)

    For L = 1 To numLists
        For r = 1 To listSizes(L)
            Dim bestOS As Double
            Dim bestOL As String
            bestOS = 0#
            bestOL = ""
            For M = 1 To numLists
                If M <> L Then
                    idx = ((L - 1) * maxRows + (r - 1)) * numLists + M
                    If bScore(idx) > bestOS Then
                        bestOS = bScore(idx)
                        bestOL = colHeaders(M)
                    End If
                End If
            Next M
            bestOverallScore(L, r) = bestOS
            bestOverallList(L, r) = bestOL
        Next r
    Next L

    ' ==================================================================
    ' 4. OUTPUT COLUMNS (row-aligned, filterable/sortable)
    ' ==================================================================
    ' For each list L, output a block of columns:
    '   For each other list M: Conf(L->M), BestMatch(L->M)
    '   Then: BestOverallConf(L), BestOverallList(L)

    Dim outCol As Long
    outCol = lastCol + 1

    For L = 1 To numLists
        ' Pairwise columns for each other list M
        For M = 1 To numLists
            If M <> L Then
                ' Confidence column
                ws.Cells(1, outCol).Value = "Conf(" & colHeaders(L) & ChrW$(8594) & colHeaders(M) & ")"
                ws.Cells(1, outCol).Font.Bold = True
                For r = 1 To listSizes(L)
                    idx = ((L - 1) * maxRows + (r - 1)) * numLists + M
                    ws.Cells(r + 1, outCol).Value = Round(bScore(idx), 4)
                Next r
                outCol = outCol + 1

                ' Best match text column
                ws.Cells(1, outCol).Value = "BestMatch(" & colHeaders(L) & ChrW$(8594) & colHeaders(M) & ")"
                ws.Cells(1, outCol).Font.Bold = True
                For r = 1 To listSizes(L)
                    idx = ((L - 1) * maxRows + (r - 1)) * numLists + M
                    ws.Cells(r + 1, outCol).Value = bText(idx)
                Next r
                outCol = outCol + 1
            End If
        Next M

        ' Best overall confidence column for this list
        ws.Cells(1, outCol).Value = "BestOverallConf(" & colHeaders(L) & ")"
        ws.Cells(1, outCol).Font.Bold = True
        For r = 1 To listSizes(L)
            ws.Cells(r + 1, outCol).Value = Round(bestOverallScore(L, r), 4)
        Next r
        outCol = outCol + 1

        ' Best overall list column for this list
        ws.Cells(1, outCol).Value = "BestOverallList(" & colHeaders(L) & ")"
        ws.Cells(1, outCol).Font.Bold = True
        For r = 1 To listSizes(L)
            ws.Cells(r + 1, outCol).Value = bestOverallList(L, r)
        Next r
        outCol = outCol + 1
    Next L

    ' ==================================================================
    ' 5. COLOUR-CODE ORIGINAL CELLS BY BEST OVERALL CONFIDENCE
    ' ==================================================================
    For L = 1 To numLists
        For r = 1 To listSizes(L)
            ApplyConfidenceColour ws.Cells(r + 1, colIndices(L)), bestOverallScore(L, r)
        Next r
    Next L

    ' ==================================================================
    ' 6. DONE
    ' ==================================================================
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Done!" & vbCrLf & vbCrLf & _
           "Lists compared: " & numLists & vbCrLf & _
           "Total items processed: " & totalItems, _
           vbInformation, "FuzzyListCompare"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "An error occurred:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "In procedure FuzzyListCompare", vbCritical, "FuzzyListCompare Error"
    Exit Sub

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
