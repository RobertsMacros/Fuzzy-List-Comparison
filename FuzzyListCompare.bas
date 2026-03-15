Attribute VB_Name = "FuzzyListCompare"
Option Explicit

' =============================================================================
' FuzzyListCompare - Sheet-Based Pairwise Fuzzy Duplicate Comparison
' =============================================================================
' Compares two user-selected lists for fuzzy duplicates via a UserForm.
' Results go on a new dedicated sheet sorted by match confidence.
' The best-match (secondary/target) column is colour-coded.
' The original sheet is never modified.
' =============================================================================

' ---------------------------------------------------------------------------
' CleanString - strip leading numbering, punctuation, collapse spaces, lowercase
' ---------------------------------------------------------------------------
Public Function CleanString(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim result As String
    Dim startPos As Long
    Dim foundLetter As Boolean

    ' Step 1: Strip leading sequential numbering
    ' Skip digits, spaces, dots, parens, hyphens until the first letter.
    startPos = 1
    foundLetter = False
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z]" Then
            startPos = i
            foundLetter = True
            Exit For
        ElseIf ch Like "[0-9 .()-]" Then
            ' Still in numbering prefix
        Else
            startPos = i
            foundLetter = True
            Exit For
        End If
    Next i
    If Not foundLetter Then startPos = 1
    s = Mid$(s, startPos)

    ' Step 2: Keep only letters, digits, and spaces
    result = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9 ]" Then
            result = result & ch
        End If
    Next i

    ' Step 3: Collapse multiple spaces
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop

    ' Step 4 & 5: Trim and lowercase
    CleanString = LCase$(Trim$(result))
End Function

' ---------------------------------------------------------------------------
' CharBagSimilarity - character-frequency similarity in [0, 1]
' ---------------------------------------------------------------------------
Public Function CharBagSimilarity(ByVal a As String, ByVal b As String) As Double
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
' ApplyConfidenceColour - set cell fill based on confidence score
' ---------------------------------------------------------------------------
Public Sub ApplyConfidenceColour(ByVal cell As Range, ByVal confidence As Double)
    With cell.Interior
        Select Case True
            Case confidence >= 0.95
                .Color = RGB(0, 128, 0)       ' Dark green  - near-exact
            Case confidence >= 0.8
                .Color = RGB(0, 176, 80)       ' Green       - strong
            Case confidence >= 0.6
                .Color = RGB(146, 208, 80)     ' Yellow-green - moderate
            Case confidence >= 0.4
                .Color = RGB(255, 255, 0)       ' Yellow      - weak
            Case confidence >= 0.2
                .Color = RGB(255, 165, 0)       ' Orange      - very weak
            Case Else
                .Color = RGB(255, 0, 0)         ' Red         - no match
        End Select
    End With
End Sub

' ---------------------------------------------------------------------------
' FuzzyListCompare - entry point (shows the UserForm)
' ---------------------------------------------------------------------------
Public Sub FuzzyListCompare()
    frmFuzzyCompare.Show
End Sub

' ---------------------------------------------------------------------------
' RunFuzzyComparison - main comparison logic, called from the UserForm
' ---------------------------------------------------------------------------
Public Sub RunFuzzyComparison(ByVal sourceWs As Worksheet, _
                              ByVal primaryCol As Long, _
                              ByVal secondaryCol As Long, _
                              ByVal primaryHeader As String, _
                              ByVal secondaryHeader As String, _
                              ByVal matchThreshold As Double)

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ==================================================================
    ' 1. READ DATA from the source sheet
    ' ==================================================================
    Dim primaryCount As Long
    Dim secondaryCount As Long
    Dim lr As Long

    lr = sourceWs.Cells(sourceWs.Rows.Count, primaryCol).End(xlUp).Row
    If lr < 2 Then primaryCount = 0 Else primaryCount = lr - 1

    lr = sourceWs.Cells(sourceWs.Rows.Count, secondaryCol).End(xlUp).Row
    If lr < 2 Then secondaryCount = 0 Else secondaryCount = lr - 1

    If primaryCount = 0 Then
        MsgBox "The primary list (" & primaryHeader & ") is empty.", vbExclamation
        GoTo Cleanup
    End If
    If secondaryCount = 0 Then
        MsgBox "The secondary list (" & secondaryHeader & ") is empty.", vbExclamation
        GoTo Cleanup
    End If

    Dim primaryItems() As String
    Dim primaryCleaned() As String
    Dim secondaryItems() As String
    Dim secondaryCleaned() As String
    ReDim primaryItems(1 To primaryCount)
    ReDim primaryCleaned(1 To primaryCount)
    ReDim secondaryItems(1 To secondaryCount)
    ReDim secondaryCleaned(1 To secondaryCount)

    Dim i As Long
    For i = 1 To primaryCount
        primaryItems(i) = CStr(sourceWs.Cells(i + 1, primaryCol).Value)
        primaryCleaned(i) = CleanString(primaryItems(i))
    Next i

    For i = 1 To secondaryCount
        secondaryItems(i) = CStr(sourceWs.Cells(i + 1, secondaryCol).Value)
        secondaryCleaned(i) = CleanString(secondaryItems(i))
    Next i

    ' ==================================================================
    ' 2. COMPUTE MATCHES
    ' ==================================================================
    Dim bestMatchScore() As Double
    Dim bestMatchIdx() As Long
    Dim bestMatchText() As String
    Dim secondaryMatched() As Boolean

    ReDim bestMatchScore(1 To primaryCount)
    ReDim bestMatchIdx(1 To primaryCount)
    ReDim bestMatchText(1 To primaryCount)
    ReDim secondaryMatched(1 To secondaryCount)

    Dim j As Long
    Dim conf As Double

    For i = 1 To primaryCount
        bestMatchScore(i) = 0#
        bestMatchIdx(i) = 0
        bestMatchText(i) = ""

        If primaryCleaned(i) <> "" Then
            For j = 1 To secondaryCount
                If secondaryCleaned(j) <> "" Then
                    conf = CharBagSimilarity(primaryCleaned(i), secondaryCleaned(j))
                    If conf > bestMatchScore(i) Then
                        bestMatchScore(i) = conf
                        bestMatchIdx(i) = j
                        bestMatchText(i) = secondaryItems(j)
                        If conf >= 1# Then Exit For
                    End If
                End If
            Next j
        End If

        ' Mark secondary item as matched if above threshold
        If bestMatchScore(i) >= matchThreshold And bestMatchIdx(i) > 0 Then
            secondaryMatched(bestMatchIdx(i)) = True
        End If
    Next i

    ' ==================================================================
    ' 3. SORT primary items by bestMatchScore DESCENDING
    ' ==================================================================
    Dim sortOrder() As Long
    ReDim sortOrder(1 To primaryCount)
    For i = 1 To primaryCount
        sortOrder(i) = i
    Next i

    Dim swapped As Boolean
    Dim tempLong As Long
    Dim k As Long
    Do
        swapped = False
        For k = 1 To primaryCount - 1
            If bestMatchScore(sortOrder(k)) < bestMatchScore(sortOrder(k + 1)) Then
                tempLong = sortOrder(k)
                sortOrder(k) = sortOrder(k + 1)
                sortOrder(k + 1) = tempLong
                swapped = True
            End If
        Next k
    Loop While swapped

    ' ==================================================================
    ' 4. CREATE NEW RESULTS SHEET
    ' ==================================================================
    Dim sheetName As String
    sheetName = primaryHeader & " vs " & secondaryHeader
    If Len(sheetName) > 31 Then sheetName = Left$(sheetName, 31)

    ' Handle duplicate sheet names
    Dim nameBase As String
    Dim nameNum As Long
    Dim nameOk As Boolean
    Dim wsCheck As Worksheet
    nameBase = sheetName
    nameNum = 1
    Do
        nameOk = True
        For Each wsCheck In ThisWorkbook.Worksheets
            If LCase$(wsCheck.Name) = LCase$(sheetName) Then
                nameOk = False
                nameNum = nameNum + 1
                sheetName = Left$(nameBase, 31 - Len(CStr(nameNum)) - 1) & " " & nameNum
                Exit For
            End If
        Next wsCheck
    Loop Until nameOk

    Dim resultsWs As Worksheet
    Set resultsWs = ThisWorkbook.Worksheets.Add(After:=sourceWs)
    resultsWs.Name = sheetName

    ' ==================================================================
    ' 5. WRITE RESULTS
    ' ==================================================================
    Dim row As Long

    ' --- Title row (merged A1:C1) ---
    resultsWs.Range("A1").Value = "Comparison: " & primaryHeader & _
                                  " checked against " & secondaryHeader & _
                                  " | Threshold: " & Format$(matchThreshold * 100, "0") & "%"
    With resultsWs.Range("A1:C1")
        .Merge
        .Font.Bold = True
        .Font.Size = 13
    End With

    ' Row 2: blank spacer

    ' --- Column headers (row 3) ---
    resultsWs.Cells(3, 1).Value = primaryHeader
    resultsWs.Cells(3, 2).Value = "Confidence"
    resultsWs.Cells(3, 3).Value = "Best Match from " & secondaryHeader
    resultsWs.Range("A3:C3").Font.Bold = True

    ' --- Data rows (row 4 onward, sorted by confidence desc) ---
    For k = 1 To primaryCount
        i = sortOrder(k)
        row = k + 3

        resultsWs.Cells(row, 1).Value = primaryItems(i)
        resultsWs.Cells(row, 2).Value = Round(bestMatchScore(i), 4)

        If bestMatchScore(i) > 0 Then
            resultsWs.Cells(row, 3).Value = bestMatchText(i)
        End If

        ' Colour-code the BEST MATCH (target/secondary) column
        ApplyConfidenceColour resultsWs.Cells(row, 3), bestMatchScore(i)
    Next k

    ' --- "Missing from secondary" section ---
    row = primaryCount + 3 + 3  ' 2 blank rows after data

    resultsWs.Cells(row, 1).Value = "Items in " & primaryHeader & _
        " with no match in " & secondaryHeader & _
        " (below " & Format$(matchThreshold * 100, "0") & "%)"
    With resultsWs.Range(resultsWs.Cells(row, 1), resultsWs.Cells(row, 3))
        .Merge
        .Font.Bold = True
    End With
    row = row + 1

    Dim missingCount As Long
    missingCount = 0
    For k = 1 To primaryCount
        i = sortOrder(k)
        If bestMatchScore(i) < matchThreshold Then
            resultsWs.Cells(row, 1).Value = primaryItems(i)
            resultsWs.Cells(row, 1).Interior.Color = RGB(255, 0, 0)
            row = row + 1
            missingCount = missingCount + 1
        End If
    Next k
    If missingCount = 0 Then
        resultsWs.Cells(row, 1).Value = "(none)"
        resultsWs.Cells(row, 1).Font.Italic = True
        row = row + 1
    End If

    ' --- "Missing from primary" section ---
    row = row + 2  ' 2 blank rows

    resultsWs.Cells(row, 1).Value = "Items in " & secondaryHeader & _
        " with no match in " & primaryHeader & _
        " (below " & Format$(matchThreshold * 100, "0") & "%)"
    With resultsWs.Range(resultsWs.Cells(row, 1), resultsWs.Cells(row, 3))
        .Merge
        .Font.Bold = True
    End With
    row = row + 1

    missingCount = 0
    For j = 1 To secondaryCount
        If Not secondaryMatched(j) Then
            resultsWs.Cells(row, 1).Value = secondaryItems(j)
            resultsWs.Cells(row, 1).Interior.Color = RGB(255, 0, 0)
            row = row + 1
            missingCount = missingCount + 1
        End If
    Next j
    If missingCount = 0 Then
        resultsWs.Cells(row, 1).Value = "(none)"
        resultsWs.Cells(row, 1).Font.Italic = True
    End If

    ' ==================================================================
    ' 6. FORMAT THE RESULTS SHEET
    ' ==================================================================
    resultsWs.Columns("A:C").AutoFit
    resultsWs.Columns("B").NumberFormat = "0.00%"

    ' Freeze panes at row 4
    resultsWs.Activate
    resultsWs.Cells(4, 1).Select
    ActiveWindow.FreezePanes = True

    ' AutoFilter on the header row
    resultsWs.Range("A3:C3").AutoFilter

    resultsWs.Cells(4, 1).Select

    ' ==================================================================
    ' 7. DONE
    ' ==================================================================
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Comparison complete!" & vbCrLf & vbCrLf & _
           "Primary list: " & primaryHeader & " (" & primaryCount & " items)" & vbCrLf & _
           "Secondary list: " & secondaryHeader & " (" & secondaryCount & " items)" & vbCrLf & vbCrLf & _
           "Results are on sheet: " & resultsWs.Name, _
           vbInformation, "FuzzyListCompare"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "An error occurred:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "In procedure RunFuzzyComparison", vbCritical, "FuzzyListCompare Error"
    Exit Sub

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
