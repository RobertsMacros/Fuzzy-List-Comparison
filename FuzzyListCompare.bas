Attribute VB_Name = "FuzzyListCompare"
Option Explicit

' =============================================================================
' FuzzyListCompare - Sheet-Based Pairwise Fuzzy Duplicate Comparison
' =============================================================================
' Compares two user-selected lists for fuzzy duplicates via a UserForm.
' Results go on a new dedicated sheet sorted by match confidence.
' The best-match (secondary/target) column is colour-coded.
' The original sheet is never modified.
'
' Scoring: each cleaned item is broken into letter pairs within each word
' ("smith" -> sm mi it th) and two items are scored with the Dice
' coefficient: 2 x shared pairs / total pairs. Word order barely matters,
' but unlike a plain letter count, anagrams and jumbles do not score 100%.
' =============================================================================

' Common file extensions stripped before comparing, so "Report.pdf" matches
' "Report". Only these are stripped, so "J.Smith" keeps its "Smith".
Private Const KNOWN_EXTENSIONS As String = "|pdf|doc|docx|docm|dot|dotx|xls|xlsx|xlsm|xlsb|csv|txt|rtf|msg|eml|ppt|pptx|odt|ods|odp|pages|numbers|key|htm|html|xml|json|zip|jpg|jpeg|png|gif|bmp|tif|tiff|heic|mp3|mp4|m4a|mov|wav|avi|"

' ---------------------------------------------------------------------------
' StripExtension - drop a trailing known file extension
' ---------------------------------------------------------------------------
Public Function StripExtension(ByVal s As String) As String
    Dim p As Long
    s = Trim$(s)
    p = InStrRev(s, ".")
    If p > 1 And p < Len(s) Then
        If InStr(1, KNOWN_EXTENSIONS, "|" & LCase$(Mid$(s, p + 1)) & "|", vbBinaryCompare) > 0 Then
            s = Left$(s, p - 1)
        End If
    End If
    StripExtension = s
End Function

' ---------------------------------------------------------------------------
' StripNumbering - drop list numbering such as "1.", "12)", "(3)", "1.2.3",
' "01 -" or "001_". A number of four or more digits (a year, a case
' number) is part of the text and is kept, as is "3M" or "7up".
' ---------------------------------------------------------------------------
Public Function StripNumbering(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim run As Long
    Dim maxRun As Long
    Dim digits As Long
    Dim prefixEnd As Long

    s = Trim$(s)
    i = 1
    If Left$(s, 1) = "(" Then i = 2
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then
            run = run + 1
            digits = digits + 1
            If run > maxRun Then maxRun = run
        ElseIf ch = "." And run > 0 Then
            run = 0
        Else
            Exit Do
        End If
        i = i + 1
    Loop

    ' Need at least one digit, no group longer than three digits,
    ' and the number must end at a separator, not run into a word.
    If digits = 0 Or maxRun > 3 Then
        StripNumbering = s
        Exit Function
    End If
    If i <= Len(s) Then
        ch = Mid$(s, i, 1)
        If Not (ch Like "[ )_-]" Or ch = vbTab) And Mid$(s, i - 1, 1) <> "." Then
            StripNumbering = s
            Exit Function
        End If
    End If

    prefixEnd = i
    Do While prefixEnd <= Len(s)
        ch = Mid$(s, prefixEnd, 1)
        If ch Like "[ )_.-]" Or ch = vbTab Then
            prefixEnd = prefixEnd + 1
        Else
            Exit Do
        End If
    Loop

    ' Nothing left but the number: keep it so the item is not blank.
    If prefixEnd > Len(s) Then
        StripNumbering = s
    Else
        StripNumbering = Mid$(s, prefixEnd)
    End If
End Function

' ---------------------------------------------------------------------------
' IsWordChar - letter (including accented letters) or digit
' ---------------------------------------------------------------------------
Private Function IsWordChar(ByVal ch As String) As Boolean
    If ch Like "[A-Za-z0-9]" Then
        IsWordChar = True
    ElseIf AscW(ch) > 127 Or AscW(ch) < 0 Then
        IsWordChar = (LCase$(ch) <> UCase$(ch))
    End If
End Function

' ---------------------------------------------------------------------------
' CleanString - strip extension and list numbering, turn punctuation into
' spaces, collapse spaces, lowercase
' ---------------------------------------------------------------------------
Public Function CleanString(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim result As String

    s = StripNumbering(StripExtension(s))

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If IsWordChar(ch) Then
            result = result & ch
        ElseIf ch = "'" Or ch = ChrW(8217) Then
            ' drop apostrophes so "Smith's" matches "Smiths"
        Else
            result = result & " "
        End If
    Next i

    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop

    CleanString = LCase$(Trim$(result))
End Function

' ---------------------------------------------------------------------------
' LetterPairs - sorted letter pairs within each word of a cleaned string.
' A one-letter word counts as one pair so it is not lost.
' ---------------------------------------------------------------------------
Public Function LetterPairs(ByVal cleaned As String) As String()
    Dim words() As String
    Dim pairs() As String
    Dim n As Long
    Dim w As Long
    Dim i As Long

    ReDim pairs(0 To Len(cleaned))
    n = 0
    If cleaned <> "" Then
        words = Split(cleaned, " ")
        For w = LBound(words) To UBound(words)
            If Len(words(w)) = 1 Then
                pairs(n) = words(w) & " "
                n = n + 1
            Else
                For i = 1 To Len(words(w)) - 1
                    pairs(n) = Mid$(words(w), i, 2)
                    n = n + 1
                Next i
            End If
        Next w
    End If

    If n = 0 Then
        ReDim pairs(0 To 0)
        pairs(0) = ""
        LetterPairs = pairs
        Exit Function
    End If
    ReDim Preserve pairs(0 To n - 1)
    SortStrings pairs
    LetterPairs = pairs
End Function

' Insertion sort; pair lists are short.
Private Sub SortStrings(ByRef a() As String)
    Dim i As Long
    Dim j As Long
    Dim v As String
    For i = LBound(a) + 1 To UBound(a)
        v = a(i)
        j = i - 1
        Do While j >= LBound(a)
            If StrComp(a(j), v, vbBinaryCompare) <= 0 Then Exit Do
            a(j + 1) = a(j)
            j = j - 1
        Loop
        a(j + 1) = v
    Next i
End Sub

' ---------------------------------------------------------------------------
' PairSimilarity - Dice coefficient of two sorted pair lists, in [0, 1]
' ---------------------------------------------------------------------------
Public Function PairSimilarity(ByRef a() As String, ByRef b() As String) As Double
    Dim i As Long
    Dim j As Long
    Dim common As Long
    Dim c As Long
    Dim na As Long
    Dim nb As Long

    na = UBound(a) - LBound(a) + 1
    nb = UBound(b) - LBound(b) + 1
    If na = 1 And a(LBound(a)) = "" Then na = 0
    If nb = 1 And b(LBound(b)) = "" Then nb = 0
    If na = 0 Or nb = 0 Then
        PairSimilarity = 0#
        Exit Function
    End If

    i = LBound(a)
    j = LBound(b)
    Do While i <= UBound(a) And j <= UBound(b)
        c = StrComp(a(i), b(j), vbBinaryCompare)
        If c = 0 Then
            common = common + 1
            i = i + 1
            j = j + 1
        ElseIf c < 0 Then
            i = i + 1
        Else
            j = j + 1
        End If
    Loop

    PairSimilarity = (2# * common) / (na + nb)
End Function

' ---------------------------------------------------------------------------
' Similarity - score two raw items, in [0, 1]
' ---------------------------------------------------------------------------
Public Function Similarity(ByVal a As String, ByVal b As String) As Double
    Dim ca As String
    Dim cb As String
    ca = CleanString(a)
    cb = CleanString(b)
    If ca = "" Or cb = "" Then
        Similarity = 0#
    ElseIf ca = cb Then
        Similarity = 1#
    Else
        Dim pa() As String
        Dim pb() As String
        pa = LetterPairs(ca)
        pb = LetterPairs(cb)
        Similarity = PairSimilarity(pa, pb)
    End If
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
' SafeSheetName - make a valid, unused sheet name in wb
' ---------------------------------------------------------------------------
Public Function SafeSheetName(ByVal wb As Workbook, ByVal wanted As String) As String
    Dim bad As Variant
    Dim nameBase As String
    Dim sheetName As String
    Dim nameNum As Long
    Dim sh As Object
    Dim taken As Boolean

    For Each bad In Array(":", "\", "/", "?", "*", "[", "]")
        wanted = Replace(wanted, bad, "_")
    Next bad
    wanted = Trim$(wanted)
    Do While Left$(wanted, 1) = "'"
        wanted = Mid$(wanted, 2)
    Loop
    Do While Right$(wanted, 1) = "'"
        wanted = Left$(wanted, Len(wanted) - 1)
    Loop
    If wanted = "" Then wanted = "Comparison"
    If Len(wanted) > 31 Then wanted = Left$(wanted, 31)

    nameBase = wanted
    sheetName = wanted
    nameNum = 1
    Do
        taken = False
        For Each sh In wb.Sheets
            If LCase$(sh.Name) = LCase$(sheetName) Then
                taken = True
                Exit For
            End If
        Next sh
        If Not taken Then Exit Do
        nameNum = nameNum + 1
        sheetName = RTrim$(Left$(nameBase, 31 - Len(CStr(nameNum)) - 1)) & " " & nameNum
    Loop
    SafeSheetName = sheetName
End Function

' ---------------------------------------------------------------------------
' FuzzyListCompare - entry point (shows the UserForm)
' ---------------------------------------------------------------------------
Public Sub FuzzyListCompare()
    If ActiveSheet Is Nothing Then
        MsgBox "Open the workbook with your lists first.", vbExclamation, "FuzzyListCompare"
        Exit Sub
    End If
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Select the worksheet that holds your lists, then run this again.", vbExclamation, "FuzzyListCompare"
        Exit Sub
    End If
    frmFuzzyCompare.Show
End Sub

' ---------------------------------------------------------------------------
' ReadList - non-blank items from row 2 down in one column
' ---------------------------------------------------------------------------
Private Function ReadList(ByVal ws As Worksheet, ByVal col As Long, ByRef items() As String) As Long
    Dim lr As Long
    Dim r As Long
    Dim n As Long
    Dim v As String

    lr = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If lr < 2 Then
        ReadList = 0
        Exit Function
    End If
    ReDim items(1 To lr - 1)
    For r = 2 To lr
        If IsError(ws.Cells(r, col).Value) Then
            v = Trim$(ws.Cells(r, col).Text)
        Else
            v = Trim$(CStr(ws.Cells(r, col).Value))
        End If
        If v <> "" Then
            n = n + 1
            items(n) = v
        End If
    Next r
    If n > 0 Then ReDim Preserve items(1 To n)
    ReadList = n
End Function

' ---------------------------------------------------------------------------
' SortByScore - stable merge sort of indices by score, highest first
' ---------------------------------------------------------------------------
Private Sub SortByScore(ByRef idx() As Long, ByRef score() As Double, ByVal lo As Long, ByVal hi As Long, ByRef tmp() As Long)
    Dim midPt As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    If hi <= lo Then Exit Sub
    midPt = (lo + hi) \ 2
    SortByScore idx, score, lo, midPt, tmp
    SortByScore idx, score, midPt + 1, hi, tmp
    i = lo
    j = midPt + 1
    k = lo
    Do While i <= midPt And j <= hi
        If score(idx(j)) > score(idx(i)) Then
            tmp(k) = idx(j)
            j = j + 1
        Else
            tmp(k) = idx(i)
            i = i + 1
        End If
        k = k + 1
    Loop
    Do While i <= midPt
        tmp(k) = idx(i)
        i = i + 1
        k = k + 1
    Loop
    Do While j <= hi
        tmp(k) = idx(j)
        j = j + 1
        k = k + 1
    Loop
    For k = lo To hi
        idx(k) = tmp(k)
    Next k
End Sub

' ---------------------------------------------------------------------------
' RunFuzzyComparison - main comparison logic, called from the UserForm
' ---------------------------------------------------------------------------
Public Sub RunFuzzyComparison(ByVal sourceWs As Worksheet, ByVal primaryCol As Long, ByVal secondaryCol As Long, ByVal primaryHeader As String, ByVal secondaryHeader As String, ByVal matchThreshold As Double)

    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ==================================================================
    ' 1. READ DATA from the source sheet (blank cells are skipped)
    ' ==================================================================
    Dim primaryItems() As String
    Dim secondaryItems() As String
    Dim primaryCount As Long
    Dim secondaryCount As Long

    primaryCount = ReadList(sourceWs, primaryCol, primaryItems)
    secondaryCount = ReadList(sourceWs, secondaryCol, secondaryItems)

    If primaryCount = 0 Then
        MsgBox "The primary list (" & primaryHeader & ") is empty.", vbExclamation
        GoTo Cleanup
    End If
    If secondaryCount = 0 Then
        MsgBox "The secondary list (" & secondaryHeader & ") is empty.", vbExclamation
        GoTo Cleanup
    End If

    Dim primaryCleaned() As String
    Dim secondaryCleaned() As String
    Dim primaryPairs() As Variant
    Dim secondaryPairs() As Variant
    ReDim primaryCleaned(1 To primaryCount)
    ReDim secondaryCleaned(1 To secondaryCount)
    ReDim primaryPairs(1 To primaryCount)
    ReDim secondaryPairs(1 To secondaryCount)

    Dim i As Long
    For i = 1 To primaryCount
        primaryCleaned(i) = CleanString(primaryItems(i))
        primaryPairs(i) = LetterPairs(primaryCleaned(i))
    Next i
    For i = 1 To secondaryCount
        secondaryCleaned(i) = CleanString(secondaryItems(i))
        secondaryPairs(i) = LetterPairs(secondaryCleaned(i))
    Next i

    ' ==================================================================
    ' 2. COMPUTE MATCHES
    ' ==================================================================
    Dim bestMatchScore() As Double
    Dim bestMatchIdx() As Long
    Dim secondaryMatched() As Boolean
    ReDim bestMatchScore(1 To primaryCount)
    ReDim bestMatchIdx(1 To primaryCount)
    ReDim secondaryMatched(1 To secondaryCount)

    Dim j As Long
    Dim conf As Double
    Dim pa() As String
    Dim pb() As String

    For i = 1 To primaryCount
        If i Mod 25 = 1 Then
            Application.StatusBar = "Fuzzy List Compare: checking " & i & " of " & primaryCount & "..."
            DoEvents
        End If

        If primaryCleaned(i) <> "" Then
            pa = primaryPairs(i)
            For j = 1 To secondaryCount
                If secondaryCleaned(j) <> "" Then
                    If primaryCleaned(i) = secondaryCleaned(j) Then
                        conf = 1#
                    Else
                        pb = secondaryPairs(j)
                        conf = PairSimilarity(pa, pb)
                    End If
                    If conf > bestMatchScore(i) Then
                        bestMatchScore(i) = conf
                        bestMatchIdx(i) = j
                        If conf >= 1# Then Exit For
                    End If
                End If
            Next j
        End If

        If bestMatchScore(i) >= matchThreshold And bestMatchIdx(i) > 0 Then
            secondaryMatched(bestMatchIdx(i)) = True
        End If
    Next i

    ' ==================================================================
    ' 3. SORT primary items by bestMatchScore DESCENDING (stable)
    ' ==================================================================
    Dim sortOrder() As Long
    Dim tmp() As Long
    ReDim sortOrder(1 To primaryCount)
    ReDim tmp(1 To primaryCount)
    For i = 1 To primaryCount
        sortOrder(i) = i
    Next i
    SortByScore sortOrder, bestMatchScore, 1, primaryCount, tmp

    ' ==================================================================
    ' 4. CREATE NEW RESULTS SHEET in the lists' own workbook
    ' ==================================================================
    Dim wb As Workbook
    Set wb = sourceWs.Parent

    Dim resultsWs As Worksheet
    Set resultsWs = wb.Worksheets.Add(After:=sourceWs)
    resultsWs.Name = SafeSheetName(wb, primaryHeader & " vs " & secondaryHeader)

    ' ==================================================================
    ' 5. WRITE RESULTS
    ' ==================================================================
    Dim row As Long
    Dim k As Long

    ' --- Title row (merged A1:C1) ---
    Dim title As String
    title = "Comparison: " & primaryHeader & " checked against " & secondaryHeader
    title = title & " | Threshold: " & Format$(matchThreshold * 100, "0") & "%"
    resultsWs.Range("A1").Value = title
    With resultsWs.Range("A1:C1")
        .Merge
        .Font.Bold = True
        .Font.Size = 13
    End With

    ' Row 2: blank spacer

    ' --- Column headers (row 3) ---
    resultsWs.Range("A:A,C:C").NumberFormat = "@"
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
        If bestMatchIdx(i) > 0 Then
            resultsWs.Cells(row, 3).Value = secondaryItems(bestMatchIdx(i))
        End If

        ApplyConfidenceColour resultsWs.Cells(row, 3), bestMatchScore(i)
    Next k

    ' --- "Missing from secondary" section ---
    row = primaryCount + 3 + 3  ' 2 blank rows after data

    Dim missingHdr As String
    missingHdr = "Items in " & primaryHeader & " with no match in " & secondaryHeader
    missingHdr = missingHdr & " (below " & Format$(matchThreshold * 100, "0") & "%)"
    resultsWs.Cells(row, 1).Value = missingHdr
    With resultsWs.Range(resultsWs.Cells(row, 1), resultsWs.Cells(row, 3))
        .Merge
        .Font.Bold = True
    End With
    row = row + 1

    Dim missingCount As Long
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

    missingHdr = "Items in " & secondaryHeader & " with no match in " & primaryHeader
    missingHdr = missingHdr & " (below " & Format$(matchThreshold * 100, "0") & "%)"
    resultsWs.Cells(row, 1).Value = missingHdr
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
    resultsWs.Columns("B").NumberFormat = "0.00%"
    resultsWs.Columns("A:C").AutoFit
    If resultsWs.Columns("A").ColumnWidth > 80 Then resultsWs.Columns("A").ColumnWidth = 80
    If resultsWs.Columns("C").ColumnWidth > 80 Then resultsWs.Columns("C").ColumnWidth = 80

    ' Freeze panes below the header row
    resultsWs.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    resultsWs.Cells(4, 1).Select
    ActiveWindow.FreezePanes = True

    ' AutoFilter on the results table only
    resultsWs.Range(resultsWs.Cells(3, 1), resultsWs.Cells(primaryCount + 3, 3)).AutoFilter

    resultsWs.Cells(4, 1).Select

    ' ==================================================================
    ' 7. DONE
    ' ==================================================================
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = prevCalc

    Dim doneMsg As String
    doneMsg = "Comparison complete!" & vbCrLf & vbCrLf
    doneMsg = doneMsg & "Primary list: " & primaryHeader & " (" & primaryCount & " items)" & vbCrLf
    doneMsg = doneMsg & "Secondary list: " & secondaryHeader & " (" & secondaryCount & " items)" & vbCrLf & vbCrLf
    doneMsg = doneMsg & "Results are on sheet: " & resultsWs.Name
    MsgBox doneMsg, vbInformation, "FuzzyListCompare"
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = prevCalc
    Dim errMsg As String
    errMsg = "An error occurred:" & vbCrLf & vbCrLf
    errMsg = errMsg & "Error " & Err.Number & ": " & Err.Description
    MsgBox errMsg, vbCritical, "FuzzyListCompare Error"
    Exit Sub

Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = prevCalc
End Sub
