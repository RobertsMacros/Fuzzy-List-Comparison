Attribute VB_Name = "FuzzyListCompare"
Option Explicit

' =============================================================================
' FuzzyListCompare
' =============================================================================
' Compares an arbitrary number of lists (one per column) on the active sheet
' using character-bag fuzzy matching. Each cell is colour-coded by its best
' match confidence against items in every other list. A deduplicated "Unique
' Items" column and a "Present In" column are appended to the right.
' =============================================================================

' ---------------------------------------------------------------------------
' CleanString — strip punctuation, collapse spaces, lowercase
' ---------------------------------------------------------------------------
Private Function CleanString(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim result As String

    result = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        ' Keep only letters, digits, and spaces
        If ch Like "[A-Za-z0-9 ]" Then
            result = result & ch
        End If
    Next i

    ' Collapse multiple spaces into one
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop

    ' Trim and lowercase
    CleanString = LCase$(Trim$(result))
End Function

' ---------------------------------------------------------------------------
' CharBagSimilarity — character-frequency-based similarity in [0, 1]
'   confidence = (2 * common_chars) / (len(A) + len(B))
' ---------------------------------------------------------------------------
Private Function CharBagSimilarity(ByVal a As String, ByVal b As String) As Double
    Dim freqA() As Long
    Dim freqB() As Long
    Dim i As Long
    Dim c As Long
    Dim commonChars As Long
    Dim totalLen As Long

    totalLen = Len(a) + Len(b)
    If totalLen = 0 Then
        CharBagSimilarity = 0#
        Exit Function
    End If

    ' Use a 256-element array to cover all byte values
    ReDim freqA(0 To 255)
    ReDim freqB(0 To 255)

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
        If freqA(i) > 0 And freqB(i) > 0 Then
            If freqA(i) < freqB(i) Then
                commonChars = commonChars + freqA(i)
            Else
                commonChars = commonChars + freqB(i)
            End If
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
                .Color = RGB(0, 128, 0)       ' Dark green — near-exact
            Case confidence >= 0.8
                .Color = RGB(0, 176, 80)       ' Green — strong
            Case confidence >= 0.6
                .Color = RGB(146, 208, 80)     ' Yellow-green — moderate
            Case confidence >= 0.4
                .Color = RGB(255, 255, 0)       ' Yellow — weak
            Case confidence >= 0.2
                .Color = RGB(255, 165, 0)       ' Orange — very weak
            Case Else
                .Color = RGB(255, 0, 0)         ' Red — no meaningful match
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

    ' ---- Detect used columns ----
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastCol < 1 Then
        MsgBox "No data found on the active sheet.", vbExclamation
        GoTo Cleanup
    End If

    ' Determine which columns actually have a header in row 1
    Dim numLists As Long
    Dim colIndices() As Long   ' 1-based column numbers of valid lists
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
        MsgBox "At least 2 lists (columns with headers) are needed for comparison. " & _
               "Found " & numLists & ".", vbExclamation
        GoTo Cleanup
    End If

    ReDim Preserve colIndices(1 To numLists)
    ReDim Preserve colHeaders(1 To numLists)

    ' ---- Read and clean all items ----
    ' items(listIndex, rowIndex) — 1-based for both dimensions
    ' cleaned(listIndex, rowIndex) — cleaned versions
    ' listSizes(listIndex) — number of items in each list

    Dim listSizes() As Long
    ReDim listSizes(1 To numLists)

    Dim maxRows As Long
    maxRows = 0

    Dim ci As Long
    For ci = 1 To numLists
        Dim lr As Long
        lr = ws.Cells(ws.Rows.Count, colIndices(ci)).End(xlUp).Row
        If lr < 2 Then
            listSizes(ci) = 0
        Else
            listSizes(ci) = lr - 1  ' items start at row 2
        End If
        If listSizes(ci) > maxRows Then maxRows = listSizes(ci)
    Next ci

    If maxRows = 0 Then
        MsgBox "All lists are empty (headers only).", vbExclamation
        GoTo Cleanup
    End If

    ' Store original and cleaned values
    Dim items() As String
    Dim cleaned() As String
    ReDim items(1 To numLists, 1 To maxRows)
    ReDim cleaned(1 To numLists, 1 To maxRows)

    Dim totalItems As Long
    totalItems = 0

    Dim r As Long
    For ci = 1 To numLists
        For r = 1 To listSizes(ci)
            items(ci, r) = CStr(ws.Cells(r + 1, colIndices(ci)).Value)
            cleaned(ci, r) = CleanString(items(ci, r))
            totalItems = totalItems + 1
        Next r
    Next ci

    ' ---- Compute best confidence for each cell ----
    ' bestConf(listIndex, rowIndex)
    Dim bestConf() As Double
    ReDim bestConf(1 To numLists, 1 To maxRows)

    Dim ci2 As Long, r2 As Long
    Dim conf As Double

    For ci = 1 To numLists
        For r = 1 To listSizes(ci)
            Dim best As Double
            best = 0#
            For ci2 = 1 To numLists
                If ci2 <> ci Then
                    For r2 = 1 To listSizes(ci2)
                        conf = CharBagSimilarity(cleaned(ci, r), cleaned(ci2, r2))
                        If conf > best Then
                            best = conf
                            ' Early exit if perfect match found
                            If best >= 1# Then GoTo NextItem
                        End If
                    Next r2
                End If
            Next ci2
NextItem:
            bestConf(ci, r) = best
        Next r
    Next ci

    ' ---- Apply colours to original cells ----
    For ci = 1 To numLists
        For r = 1 To listSizes(ci)
            ApplyConfidenceColour ws.Cells(r + 1, colIndices(ci)), bestConf(ci, r)
        Next r
    Next ci

    ' ---- Build deduplicated "Unique Items" list ----
    ' We iterate column by column, item by item. An item is added to the
    ' unique list only if no existing unique item scores >= 0.80 against it.

    ' Dynamic arrays for unique items
    Dim uCount As Long          ' number of unique items so far
    Dim uCleaned() As String    ' cleaned text of each unique item
    Dim uOriginal() As String   ' display text (canonical)
    Dim uPresence() As String   ' comma-separated header names
    Dim uSourceCol() As Long    ' which column the canonical text came from

    ' Pre-size generously
    Dim totalMax As Long
    totalMax = 0
    For ci = 1 To numLists
        totalMax = totalMax + listSizes(ci)
    Next ci
    If totalMax = 0 Then totalMax = 1

    ReDim uCleaned(1 To totalMax)
    ReDim uOriginal(1 To totalMax)
    ReDim uPresence(1 To totalMax)
    ReDim uSourceCol(1 To totalMax)

    ' Track which lists each unique item appears in (bit-style with string)
    ' We will use a boolean 2D: uInList(uniqueIndex, listIndex)
    Dim uInList() As Boolean
    ReDim uInList(1 To totalMax, 1 To numLists)

    uCount = 0

    For ci = 1 To numLists
        For r = 1 To listSizes(ci)
            Dim matched As Boolean
            matched = False

            Dim u As Long
            For u = 1 To uCount
                conf = CharBagSimilarity(cleaned(ci, r), uCleaned(u))
                If conf >= 0.8 Then
                    ' This item matches an existing unique item
                    matched = True
                    uInList(u, ci) = True
                    Exit For
                End If
            Next u

            If Not matched Then
                ' New unique item
                uCount = uCount + 1
                uCleaned(uCount) = cleaned(ci, r)
                uOriginal(uCount) = items(ci, r)
                uSourceCol(uCount) = ci
                uInList(uCount, ci) = True
            End If
        Next r
    Next ci

    ' ---- Write "Unique Items" and "Present In" columns ----
    Dim uniqueCol As Long
    uniqueCol = lastCol + 1
    Dim presentCol As Long
    presentCol = lastCol + 2

    ws.Cells(1, uniqueCol).Value = "Unique Items"
    ws.Cells(1, presentCol).Value = "Present In"

    ' Make headers bold
    ws.Cells(1, uniqueCol).Font.Bold = True
    ws.Cells(1, presentCol).Font.Bold = True

    For u = 1 To uCount
        ' Write the canonical text
        ws.Cells(u + 1, uniqueCol).Value = uOriginal(u)

        ' Build "Present In" string and count how many lists
        Dim presenceStr As String
        Dim presenceCount As Long
        presenceStr = ""
        presenceCount = 0

        For ci = 1 To numLists
            If uInList(u, ci) Then
                presenceCount = presenceCount + 1
                If presenceStr <> "" Then presenceStr = presenceStr & ", "
                presenceStr = presenceStr & colHeaders(ci)
            End If
        Next ci

        ws.Cells(u + 1, presentCol).Value = presenceStr

        ' Colour-code the Unique Items cell by presence count
        With ws.Cells(u + 1, uniqueCol).Interior
            If presenceCount = numLists Then
                .Color = RGB(0, 128, 0)         ' Dark green — all lists
            ElseIf presenceCount > numLists / 2 Then
                .Color = RGB(146, 208, 80)       ' Yellow-green — most lists
            Else
                .Color = RGB(255, 0, 0)           ' Red — one list only
            End If
        End With
    Next u

    ' ---- Done ----
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Done!" & vbCrLf & vbCrLf & _
           "Lists compared: " & numLists & vbCrLf & _
           "Total items processed: " & totalItems & vbCrLf & _
           "Unique items found: " & uCount, vbInformation, "FuzzyListCompare"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "An error occurred:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "In procedure FuzzyListCompare", vbCritical, "FuzzyListCompare Error"

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
