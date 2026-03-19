Attribute VB_Name = "ActReferenceChecker"
Option Explicit

' =============================================================================
' ActReferenceChecker - Word VBA Macro
' =============================================================================
' Scans a Word document for legislative references (e.g. "Section 12 of the
' FIA") within each document Part. Tracks first-seen references per Part and
' adds comments to flag subsequent (duplicate) references.
'
' Usage:
'   1. Open the Word document to check.
'   2. Run the macro FindDuplicateActReferences.
'   3. Enter Act abbreviations as CSV (e.g. "FIA,FAIS,Banks Act").
'   4. The macro detects Parts, scans sentences, and inserts comments on
'      duplicate references.
' =============================================================================

' ---------------------------------------------------------------------------
' FindDuplicateActReferences - main entry point
' ---------------------------------------------------------------------------
Public Sub FindDuplicateActReferences()
    Dim doc As Document
    Dim csvInput As String
    Dim Act_List() As String
    Dim Parts_List As Collection
    Dim i As Long

    Set doc = ActiveDocument

    ' Step 1: Ask user for CSV input of Act names
    csvInput = InputBox( _
        "Enter Act abbreviations separated by commas" & vbCrLf & _
        "(e.g. FIA,FAIS,Banks Act):", _
        "Act Reference Checker")
    If Len(Trim(csvInput)) = 0 Then Exit Sub

    ' Step 2: Parse CSV into ordered Act_List
    Act_List = ParseCSV(csvInput)
    If UBound(Act_List) < 0 Then
        MsgBox "No valid Act names found in input.", vbExclamation
        Exit Sub
    End If

    ' Step 3: Detect document Parts
    Set Parts_List = FindDocumentParts(doc)
    If Parts_List.Count = 0 Then
        MsgBox "No Parts found in the document." & vbCrLf & _
               "The macro looks for paragraphs starting with " & _
               """Part"" followed by a number (e.g. ""Part 1"").", _
               vbExclamation
        Exit Sub
    End If

    ' Step 4: Process each Part
    For i = 1 To Parts_List.Count
        ProcessPart doc, Parts_List, i, Act_List
    Next i

    MsgBox "Reference checking complete." & vbCrLf & _
           Parts_List.Count & " Part(s) scanned." & vbCrLf & _
           "Check document comments for flagged duplicates.", _
           vbInformation, "Act Reference Checker"
End Sub

' ---------------------------------------------------------------------------
' ParseCSV - split comma-separated string into trimmed array
' ---------------------------------------------------------------------------
Private Function ParseCSV(ByVal csv As String) As String()
    Dim raw() As String
    Dim cleaned() As String
    Dim i As Long
    Dim count As Long

    raw = Split(csv, ",")

    ' First pass: count non-empty items
    count = 0
    For i = LBound(raw) To UBound(raw)
        If Len(Trim(raw(i))) > 0 Then count = count + 1
    Next i

    If count = 0 Then
        cleaned = Split("", ",")  ' empty array
        ParseCSV = cleaned
        Exit Function
    End If

    ReDim cleaned(0 To count - 1)
    count = 0
    For i = LBound(raw) To UBound(raw)
        If Len(Trim(raw(i))) > 0 Then
            cleaned(count) = Trim(raw(i))
            count = count + 1
        End If
    Next i

    ParseCSV = cleaned
End Function

' ---------------------------------------------------------------------------
' FindDocumentParts - locate paragraphs that begin a "Part" division
' ---------------------------------------------------------------------------
' Looks for paragraphs whose text starts with "Part" followed by a number,
' e.g. "Part 1", "Part 2 - Definitions", "PART III".
' Returns a Collection of Paragraph objects marking each Part boundary.
' ---------------------------------------------------------------------------
Private Function FindDocumentParts(doc As Document) As Collection
    Dim parts As New Collection
    Dim para As Paragraph
    Dim txt As String
    Dim re As Object

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\s*Part\s+(\d+|[IVXLCDM]+)\b"
    re.IgnoreCase = True

    For Each para In doc.Paragraphs
        txt = para.Range.Text
        If re.Test(txt) Then
            parts.Add para
        End If
    Next para

    Set FindDocumentParts = parts
End Function

' ---------------------------------------------------------------------------
' ProcessPart - scan one Part for Section references and flag duplicates
' ---------------------------------------------------------------------------
Private Sub ProcessPart( _
        doc As Document, _
        Parts_List As Collection, _
        partIndex As Long, _
        Act_List() As String)

    ' --- Build the range covering this Part ---
    Dim startPos As Long
    Dim endPos As Long

    startPos = Parts_List(partIndex).Range.Start

    If partIndex < Parts_List.Count Then
        endPos = Parts_List(partIndex + 1).Range.Start - 1
    Else
        endPos = doc.Content.End
    End If

    ' Guard against invalid range
    If endPos <= startPos Then Exit Sub

    Dim partRange As Range
    Set partRange = doc.Range(startPos, endPos)

    ' --- First_Seen dictionary: key = "Section N|ActName", value = True ---
    Dim First_Seen As Object
    Set First_Seen = CreateObject("Scripting.Dictionary")
    First_Seen.CompareMode = vbTextCompare

    ' --- Regex to find "Section <number>" ---
    ' Captures digits only (stops at "(" or lowercase letter per spec)
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
