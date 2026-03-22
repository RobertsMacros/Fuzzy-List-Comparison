Attribute VB_Name = "ActReferenceChecker"
Option Explicit

' =============================================================================
' ActReferenceChecker - Word VBA Macro
' =============================================================================
' Scans a Word document for legislative references (e.g. "Section 12 of the
' FIA") within each user-defined document Part. Tracks first-seen references
' per Part and adds comments to flag subsequent (duplicate) references.
'
' Usage:
'   1. Open the Word document to check.
'   2. Run the macro FindDuplicateActReferences.
'   3. Enter Act abbreviations as CSV (e.g. "FIA,FAIS,Banks Act").
'   4. Enter Part headings as CSV, in document order
'      (e.g. "PART 1: INTERPRETATION, PART 2: AUTHORISATION").
'   5. The macro scans sentences within each Part range and inserts comments
'      on duplicate references.
' =============================================================================

' ---------------------------------------------------------------------------
' FindDuplicateActReferences - main entry point
' ---------------------------------------------------------------------------
Public Sub FindDuplicateActReferences()
    Dim doc As Document
    Dim csvInput As String
    Dim Act_List() As String
    Dim Parts_List() As String
    Dim PartRanges As Collection
    Dim i As Long

    Set doc = ActiveDocument

    ' --- Step 1: Ask user for Act CSV ---
    csvInput = InputBox( _
        "Enter Act abbreviations separated by commas" & vbCrLf & _
        "(e.g. FIA, FAIS, Banks Act):", _
        "Act Reference Checker - Acts")
    If Len(Trim(csvInput)) = 0 Then Exit Sub

    Act_List = ParseCSVList(csvInput)
    If Not HasItems(Act_List) Then
        MsgBox "No valid Act names found in input.", vbExclamation
        Exit Sub
    End If

    ' --- Step 2: Ask user for Parts CSV ---
    csvInput = InputBox( _
        "Enter Part headings separated by commas, in document order" & vbCrLf & _
        "(e.g. PART 1: INTERPRETATION, PART 2: AUTHORISATION):", _
        "Act Reference Checker - Parts")
    If Len(Trim(csvInput)) = 0 Then Exit Sub

    Parts_List = ParseCSVList(csvInput)
    If Not HasItems(Parts_List) Then
        MsgBox "No valid Part headings found in input.", vbExclamation
        Exit Sub
    End If

    ' --- Step 3: Resolve Part headings to document ranges ---
    Set PartRanges = BuildPartRanges(doc, Parts_List)
    If PartRanges Is Nothing Then
        ' BuildPartRanges already showed the error message
        Exit Sub
    End If

    ' --- Step 4: Process each Part range ---
    For i = 1 To PartRanges.Count
        ProcessPartRange doc, PartRanges(i), Act_List
    Next i

    MsgBox "Reference checking complete." & vbCrLf & _
           PartRanges.Count & " Part(s) scanned." & vbCrLf & _
           "Check document comments for flagged duplicates.", _
           vbInformation, "Act Reference Checker"
End Sub

' ---------------------------------------------------------------------------
' ParseCSVList - split comma-separated string into a trimmed, ordered array
' ---------------------------------------------------------------------------
' Returns a String array with LBound 0. If there are no valid items, returns
' an unallocated array (test with HasItems before using).
' ---------------------------------------------------------------------------
Private Function ParseCSVList(ByVal csv As String) As String()
    Dim raw() As String
    Dim result() As String
    Dim i As Long
    Dim count As Long

    raw = Split(csv, ",")

    ' First pass: count non-empty items
    count = 0
    For i = LBound(raw) To UBound(raw)
        If Len(Trim(raw(i))) > 0 Then count = count + 1
    Next i

    If count = 0 Then
        ' Return unallocated array
        ParseCSVList = result
        Exit Function
    End If

    ReDim result(0 To count - 1)
    count = 0
    For i = LBound(raw) To UBound(raw)
        If Len(Trim(raw(i))) > 0 Then
            result(count) = Trim(raw(i))
            count = count + 1
        End If
    Next i

    ParseCSVList = result
End Function

' ---------------------------------------------------------------------------
' HasItems - check whether a String array has been allocated and has items
' ---------------------------------------------------------------------------
Private Function HasItems(arr() As String) As Boolean
    On Error GoTo NoItems
    If UBound(arr) >= LBound(arr) Then
        HasItems = True
    Else
        HasItems = False
    End If
    Exit Function
NoItems:
    HasItems = False
End Function

' ---------------------------------------------------------------------------
' NormaliseText - normalise a string for heading comparison
' ---------------------------------------------------------------------------
' Trims, collapses repeated internal spaces, and lowercases.
' ---------------------------------------------------------------------------
Private Function NormaliseText(ByVal s As String) As String
    Dim result As String
    result = Trim(s)

    ' Collapse repeated spaces
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop

    ' Strip trailing paragraph mark (Chr(13)) that Word appends to
    ' paragraph text - without this, comparisons would always fail
    Do While Len(result) > 0
        Dim lastCh As String
        lastCh = Right$(result, 1)
        If lastCh = vbCr Or lastCh = vbLf Or lastCh = Chr$(7) Then
            result = Left$(result, Len(result) - 1)
        Else
            Exit Do
        End If
    Loop

    NormaliseText = LCase$(result)
End Function

' ---------------------------------------------------------------------------
' BuildPartRanges - resolve user-supplied headings to document ranges
' ---------------------------------------------------------------------------
' For each heading in Parts_List, finds the matching paragraph in the
' document. Returns a Collection of Range objects, one per Part.
' Returns Nothing and shows an error if any heading is missing, duplicated,
' or if headings are not in increasing document order.
' ---------------------------------------------------------------------------
Private Function BuildPartRanges( _
        doc As Document, _
        Parts_List() As String) As Collection

    Dim partCount As Long
    Dim i As Long
    Dim j As Long

    partCount = UBound(Parts_List) - LBound(Parts_List) + 1

    ' --- Find each heading's paragraph position ---
    Dim startPositions() As Long   ' document character position
    Dim foundParas() As Paragraph  ' the matched paragraph
    ReDim startPositions(0 To partCount - 1)
    ReDim foundParas(0 To partCount - 1)

    Dim para As Paragraph
    Dim normParaText As String

    For i = 0 To partCount - 1
        Dim normHeading As String
        Dim matchCount As Long
        Dim matchedPara As Paragraph

        normHeading = NormaliseText(Parts_List(i))
        matchCount = 0
        Set matchedPara = Nothing

        For Each para In doc.Paragraphs
            normParaText = NormaliseText(para.Range.Text)
            If normParaText = normHeading Then
                matchCount = matchCount + 1
                If matchedPara Is Nothing Then
                    Set matchedPara = para
                End If
            End If
        Next para

        ' --- Validate: heading must exist ---
        If matchCount = 0 Then
            MsgBox "Part heading not found in document:" & vbCrLf & vbCrLf & _
                   """" & Parts_List(i) & """" & vbCrLf & vbCrLf & _
                   "Please check spelling and try again.", _
                   vbCritical, "Act Reference Checker"
            Set BuildPartRanges = Nothing
            Exit Function
        End If

        ' --- Validate: heading must not be duplicated ---
        If matchCount > 1 Then
            MsgBox "Part heading appears " & matchCount & " times in the " & _
                   "document:" & vbCrLf & vbCrLf & _
                   """" & Parts_List(i) & """" & vbCrLf & vbCrLf & _
                   "The heading must be unique. Please make it more specific.", _
                   vbCritical, "Act Reference Checker"
            Set BuildPartRanges = Nothing
            Exit Function
        End If

        Set foundParas(i) = matchedPara
        startPositions(i) = matchedPara.Range.Start
    Next i

    ' --- Validate: headings must be in increasing document order ---
    For i = 0 To partCount - 2
        If startPositions(i) >= startPositions(i + 1) Then
            MsgBox "Part headings are not in document order." & vbCrLf & vbCrLf & _
                   """" & Parts_List(i) & """" & " appears at or after " & _
                   """" & Parts_List(i + 1) & """" & " in the document." & _
                   vbCrLf & vbCrLf & _
                   "Please supply headings in the order they appear.", _
                   vbCritical, "Act Reference Checker"
            Set BuildPartRanges = Nothing
            Exit Function
        End If
    Next i

    ' --- Build ranges ---
    Dim ranges As New Collection
    For i = 0 To partCount - 1
        Dim startPos As Long
        Dim endPos As Long

        startPos = startPositions(i)

        If i < partCount - 1 Then
            ' End just before the next Part heading
            endPos = startPositions(i + 1)
        Else
            ' Last Part runs to end of document
            endPos = doc.Content.End
        End If

        If endPos > startPos Then
            ranges.Add doc.Range(startPos, endPos)
        End If
    Next i

    Set BuildPartRanges = ranges
End Function

' ---------------------------------------------------------------------------
' ProcessPartRange - scan one Part range for Section references and flag
'                    duplicates
' ---------------------------------------------------------------------------
Private Sub ProcessPartRange( _
        doc As Document, _
        partRange As Range, _
        Act_List() As String)

    ' --- Fresh First_Seen dictionary for this Part ---
    ' Key = "Section N|ActName", value = True
    Dim First_Seen As Object
    Set First_Seen = CreateObject("Scripting.Dictionary")
    First_Seen.CompareMode = vbTextCompare

    ' --- Regex to find "Section <number>" ---
    ' Captures digits only (stops at "(" or lowercase letter)
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "Section\s+(\d+)"
    re.IgnoreCase = True
    re.Global = True

    ' --- Loop through each sentence in this Part ---
    Dim s As Long
    Dim sent As Range
    Dim sentText As String
    Dim matches As Object
    Dim m As Long

    For s = 1 To partRange.Sentences.Count
        Set sent = partRange.Sentences(s)
        sentText = sent.Text

        ' Find all "Section [number]" occurrences in this sentence
        Set matches = re.Execute(sentText)
        If matches.Count = 0 Then GoTo NextSentence

        For m = 0 To matches.Count - 1
            Dim sectionMatch As Object
            Set sectionMatch = matches(m)

            Dim sectionNum As String
            sectionNum = sectionMatch.SubMatches(0)

            ' Find the closest Act from Act_List after this match
            Dim actName As String
            actName = FindClosestAct( _
                sentText, _
                sectionMatch.FirstIndex + sectionMatch.Length, _
                Act_List)

            If Len(actName) = 0 Then GoTo NextMatch

            ' Build the reference key
            Dim refKey As String
            refKey = "Section " & sectionNum & "|" & actName

            ' Locate the full reference range in the document
            ' (from "Section" to end of the Act name)
            Dim refRange As Range
            Set refRange = BuildRefRange( _
                doc, sent, sentText, sectionMatch, actName)

            If refRange Is Nothing Then GoTo NextMatch

            ' --- First-seen check ---
            If Not First_Seen.Exists(refKey) Then
                ' Record as first occurrence
                First_Seen.Add refKey, True
            Else
                ' Flag as additional/duplicate reference
                doc.Comments.Add refRange, _
                    "Additional reference: Section " & sectionNum & _
                    " of " & actName & _
                    " (first seen earlier in this Part)"
            End If

NextMatch:
        Next m
NextSentence:
    Next s
End Sub

' ---------------------------------------------------------------------------
' FindClosestAct - find the nearest Act name after a given position
' ---------------------------------------------------------------------------
' Searches sentText from startPos onward for the Act in Act_List that
' appears closest (leftmost) to the section reference.
' ---------------------------------------------------------------------------
Private Function FindClosestAct( _
        ByVal sentText As String, _
        ByVal startPos As Long, _
        Act_List() As String) As String

    Dim closestPos As Long
    Dim closestAct As String
    Dim i As Long
    Dim pos As Long

    closestPos = Len(sentText) + 1
    closestAct = ""

    For i = LBound(Act_List) To UBound(Act_List)
        ' InStr is 1-based; startPos from regex is 0-based, so +1
        pos = InStr(startPos + 1, sentText, Act_List(i), vbTextCompare)
        If pos > 0 And pos < closestPos Then
            closestPos = pos
            closestAct = Act_List(i)
        End If
    Next i

    FindClosestAct = closestAct
End Function

' ---------------------------------------------------------------------------
' BuildRefRange - compute the document Range covering the full reference
' ---------------------------------------------------------------------------
' Returns a Range from the start of "Section ..." to the end of the Act name
' within the sentence. Returns Nothing if the Act name position cannot be
' determined.
' ---------------------------------------------------------------------------
Private Function BuildRefRange( _
        doc As Document, _
        sent As Range, _
        ByVal sentText As String, _
        sectionMatch As Object, _
        ByVal actName As String) As Range

    Dim matchStart As Long
    Dim matchEnd As Long
    Dim actPos As Long

    ' Start of "Section ..." in the document
    matchStart = sent.Start + sectionMatch.FirstIndex

    ' Find where the Act name appears after the section number
    actPos = InStr( _
        sectionMatch.FirstIndex + sectionMatch.Length + 1, _
        sentText, actName, vbTextCompare)

    If actPos = 0 Then
        Set BuildRefRange = Nothing
        Exit Function
    End If

    ' End of the Act name in the document (InStr is 1-based, Range is 0-based)
    matchEnd = sent.Start + (actPos - 1) + Len(actName)

    Set BuildRefRange = doc.Range(matchStart, matchEnd)
End Function
