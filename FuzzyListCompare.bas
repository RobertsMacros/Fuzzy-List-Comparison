Attribute VB_Name = "FuzzyListCompare"
Option Explicit

' =============================================================================
' FuzzyListCompare - Sheet-Based Pairwise Fuzzy Duplicate Comparison
' =============================================================================
' Compares two user-selected lists for fuzzy duplicates. Results go on a new
' dedicated sheet sorted by match confidence, with "missing" sections at the
' bottom. The original sheet is never modified.
' =============================================================================

' ---------------------------------------------------------------------------
' CleanString - strip leading numbering, punctuation, collapse spaces, lowercase
' ---------------------------------------------------------------------------
Private Function CleanString(ByVal s As String) As String
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
            ' Still in numbering prefix - keep skipping
        Else
            ' Non-numbering, non-letter char - stop stripping
            startPos = i
            foundLetter = True
            Exit For
        End If
    Next i
    ' If no letters found at all, keep the original string
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
' ApplyConfidenceColour - set cell fill based on confidence score
' ---------------------------------------------------------------------------
Private Sub ApplyConfidenceColour(ByVal cell As Range, ByVal confidence As Double)
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
' FuzzyListCompare - main entry point
' ---------------------------------------------------------------------------
Public Sub FuzzyListCompare()

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim sourceWs As Worksheet
    Set sourceWs = ActiveSheet

    ' ==================================================================
    ' 1. DETECT LIST COLUMNS on the source sheet
    ' ==================================================================
    Dim lastCol As Long
    lastCol = sourceWs.Cells(1, sourceWs.Columns.Count).End(xlToLeft).Column

    If lastCol < 1 Then
        MsgBox "No data found on the active sheet.", vbExclamation
        GoTo Cleanup
    End If

    Dim numLists As Long
    Dim colIndices() As Long
    Dim colHeaders() As String
    Dim tempCount As Long
    Dim c As Long

    tempCount = 0
    ReDim colIndices(1 To lastCol)
    ReDim colHeaders(1 To lastCol)

    For c = 1 To lastCol
        If Trim$(CStr(sourceWs.Cells(1, c).Value)) <> "" Then
            tempCount = tempCount + 1
            colIndices(tempCount) = c
            colHeaders(tempCount) = CStr(sourceWs.Cells(1, c).Value)
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
    ' 2. USER INPUT - pick primary, secondary, and threshold
    ' ==================================================================

    ' Step 1: Show detected lists
    Dim listMsg As String
    Dim n As Long
    listMsg = "The following lists were found:" & vbCrLf & vbCrLf
    For n = 1 To numLists
        listMsg = listMsg & n & " - " & colHeaders(n) & vbCrLf
    Next n
    listMsg = listMsg & vbCrLf & "Press OK to continue."
    MsgBox listMsg, vbInformation, "FuzzyListCompare - Detected Lists"

    ' Step 2: Get primary list number
    Dim primaryIdx As Long
    Dim inputStr As String
    Do
        inputStr = InputBox("Enter the number of the PRIMARY list (the list you want to check):" & _
                            vbCrLf & vbCrLf & "Enter a number from 1 to " & numLists & ".", _
                            "FuzzyListCompare - Primary List")
        If StrPtr(inputStr) = 0 Then
            ' User pressed Cancel
            GoTo Cleanup
        End If
        If IsNumeric(inputStr) Then
            primaryIdx = CLng(inputStr)
            If primaryIdx >= 1 And primaryIdx <= numLists Then Exit Do
        End If
        MsgBox "Invalid input. Please enter a number between 1 and " & numLists & ".", vbExclamation
    Loop

    ' Step 3: Get secondary list number
    Dim secondaryIdx As Long
    Do
        inputStr = InputBox("Enter the number of the SECONDARY list (the list to check against):" & _
                            vbCrLf & vbCrLf & "Enter a number from 1 to " & numLists & _
                            " (not " & primaryIdx & ").", _
                            "FuzzyListCompare - Secondary List")
        If StrPtr(inputStr) = 0 Then GoTo Cleanup
        If IsNumeric(inputStr) Then
            secondaryIdx = CLng(inputStr)
            If secondaryIdx >= 1 And secondaryIdx <= numLists And secondaryIdx <> primaryIdx Then Exit Do
        End If
        MsgBox "Invalid input. Please enter a number between 1 and " & numLists & _
               ", different from " & primaryIdx & ".", vbExclamation
    Loop

    ' Step 4: Get confidence threshold
    Dim matchThreshold As Double
    matchThreshold = 0.7
    inputStr = InputBox("Enter the confidence threshold for a 'match' (e.g. 0.7 for 70%):" & _
                        vbCrLf & vbCrLf & "Default: 0.7", _
                        "FuzzyListCompare - Threshold", "0.7")
    If StrPtr(inputStr) <> 0 And inputStr <> "" Then
        If IsNumeric(inputStr) Then
            Dim tempThresh As Double
            tempThresh = CDbl(inputStr)
            If tempThresh >= 0 And tempThresh <= 1 Then
                matchThreshold = tempThresh
            End If
        End If
    End If

    ' ==================================================================
    ' 3. READ DATA from the source sheet
    ' ==================================================================
    Dim primaryCount As Long
    Dim secondaryCount As Long
    Dim lr As Long

    ' Primary list
    lr = sourceWs.Cells(sourceWs.Rows.Count, colIndices(primaryIdx)).End(xlUp).Row
    If lr < 2 Then
        primaryCount = 0
    Else
        primaryCount = lr - 1
    End If

    ' Secondary list
    lr = sourceWs.Cells(sourceWs.Rows.Count, colIndices(secondaryIdx)).End(xlUp).Row
    If lr < 2 Then
        secondaryCount = 0
    Else
        secondaryCount = lr - 1
    End If

    If primaryCount = 0 Then
        MsgBox "The primary list (" & colHeaders(primaryIdx) & ") is empty.", vbExclamation
        GoTo Cleanup
    End If
    If secondaryCount = 0 Then
        MsgBox "The secondary list (" & colHeaders(secondaryIdx) & ") is empty.", vbExclamation
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
        primaryItems(i) = CStr(sourceWs.Cells(i + 1, colIndices(primaryIdx)).Value)
        primaryCleaned(i) = CleanString(primaryItems(i))
    Next i

    For i = 1 To secondaryCount
        secondaryItems(i) = CStr(sourceWs.Cells(i + 1, colIndices(secondaryIdx)).Value)
        secondaryCleaned(i) = CleanString(secondaryItems(i))
    Next i

    ' ==================================================================
    ' 4. COMPUTE MATCHES
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
    ' 5. SORT primary items by bestMatchScore DESCENDING
    ' ==================================================================
    Dim sortOrder() As Long
    ReDim sortOrder(1 To primaryCount)
    For i = 1 To primaryCount
        sortOrder(i) = i
    Next i

    ' Simple bubble sort (descending by score)
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
    ' 6. CREATE NEW RESULTS SHEET
    ' ==================================================================
    Dim sheetName As String
    sheetName = colHeaders(primaryIdx) & " vs " & colHeaders(secondaryIdx)
    ' Excel sheet names max 31 chars
    If Len(sheetName) > 31 Then sheetName = Left$(sheetName, 31)

    ' If sheet name already exists, append a number
    Dim nameOk As Boolean
    Dim nameBase As String
    Dim nameNum As Long
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
    ' 7. WRITE RESULTS on the new sheet
    ' ==================================================================
    Dim row As Long

    ' --- Title row ---
    resultsWs.Range("A1").Value = "Comparison: " & colHeaders(primaryIdx) & _
                                  " checked against " & colHeaders(secondaryIdx) & _
                                  " | Threshold: " & Format$(matchThreshold * 100, "0") & "%"
    With resultsWs.Range("A1:C1")
        .Merge
        .Font.Bold = True
        .Font.Size = 13
    End With

    ' Row 2: blank spacer

    ' --- Column headers in row 3 ---
    resultsWs.Cells(3, 1).Value = colHeaders(primaryIdx)
    resultsWs.Cells(3, 2).Value = "Confidence"
    resultsWs.Cells(3, 3).Value = "Best Match from " & colHeaders(secondaryIdx)
    resultsWs.Range("A3:C3").Font.Bold = True

    ' --- Data rows (sorted by confidence descending) ---
    For k = 1 To primaryCount
        i = sortOrder(k)
        row = k + 3
        resultsWs.Cells(row, 1).Value = primaryItems(i)
        resultsWs.Cells(row, 2).Value = Round(bestMatchScore(i), 4)
        If bestMatchScore(i) > 0 Then
            resultsWs.Cells(row, 3).Value = bestMatchText(i)
        End If

        ' Colour-code the primary item cell
        ApplyConfidenceColour resultsWs.Cells(row, 1), bestMatchScore(i)
    Next k

    ' --- "Missing from secondary" section ---
    row = primaryCount + 3 + 2 + 1  ' 2 blank rows after data
    resultsWs.Cells(row, 1).Value = "Items in " & colHeaders(primaryIdx) & _
                                     " with no match in " & colHeaders(secondaryIdx) & _
                                     " (below " & Format$(matchThreshold * 100, "0") & "%)"
    With resultsWs.Range(resultsWs.Cells(row, 1), resultsWs.Cells(row, 3))
        .Merge
        .Font.Bold = True
    End With
    row = row + 1

    For k = 1 To primaryCount
        i = sortOrder(k)
        If bestMatchScore(i) < matchThreshold Then
            resultsWs.Cells(row, 1).Value = primaryItems(i)
            resultsWs.Cells(row, 1).Interior.Color = RGB(255, 0, 0)
            row = row + 1
        End If
    Next k

    ' --- "Missing from primary" section ---
    row = row + 2  ' 2 blank rows
    resultsWs.Cells(row, 1).Value = "Items in " & colHeaders(secondaryIdx) & _
                                     " with no match in " & colHeaders(primaryIdx) & _
                                     " (below " & Format$(matchThreshold * 100, "0") & "%)"
    With resultsWs.Range(resultsWs.Cells(row, 1), resultsWs.Cells(row, 3))
        .Merge
        .Font.Bold = True
    End With
    row = row + 1

    For j = 1 To secondaryCount
        If Not secondaryMatched(j) Then
            resultsWs.Cells(row, 1).Value = secondaryItems(j)
            resultsWs.Cells(row, 1).Interior.Color = RGB(255, 0, 0)
            row = row + 1
        End If
    Next j

    ' ==================================================================
    ' 8. FORMAT THE RESULTS SHEET
    ' ==================================================================
    ' Auto-fit columns A, B, C
    resultsWs.Columns("A:C").AutoFit

    ' Format confidence column as percentage
    resultsWs.Columns("B").NumberFormat = "0.00%"

    ' Freeze panes at row 4 (so row 3 headers stay visible)
    resultsWs.Activate
    resultsWs.Cells(4, 1).Select
    ActiveWindow.FreezePanes = True

    ' Add AutoFilter to the header row (row 3)
    resultsWs.Range("A3:C3").AutoFilter

    ' Select first data cell
    resultsWs.Cells(4, 1).Select

    ' ==================================================================
    ' 9. DONE
    ' ==================================================================
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Comparison complete!" & vbCrLf & vbCrLf & _
           "Primary list: " & colHeaders(primaryIdx) & " (" & primaryCount & " items)" & vbCrLf & _
           "Secondary list: " & colHeaders(secondaryIdx) & " (" & secondaryCount & " items)" & vbCrLf & vbCrLf & _
           "Results are on sheet: " & resultsWs.Name, _
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
