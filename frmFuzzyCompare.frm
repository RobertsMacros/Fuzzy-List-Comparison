VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFuzzyCompare
   Caption         =   "Fuzzy List Compare"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFuzzyCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' -------------------------------------------------------
' Module-level storage for detected lists
' -------------------------------------------------------
Private m_colIndices() As Long
Private m_colHeaders() As String
Private m_numLists As Long
Private m_sourceWs As Worksheet

' -------------------------------------------------------
' UserForm_Initialize - detect lists, populate controls
' -------------------------------------------------------
Private Sub UserForm_Initialize()
    Set m_sourceWs = ActiveSheet

    ' -- Instructions --
    Dim txt As String
    txt = "FUZZY LIST COMPARE" & vbCrLf & String(35, "-") & vbCrLf & vbCrLf
    txt = txt & "Compare two lists for fuzzy duplicate matches." & vbCrLf & vbCrLf
    txt = txt & "HOW TO USE:" & vbCrLf
    txt = txt & "1. Select your PRIMARY list - the list you want to check." & vbCrLf
    txt = txt & "2. Select the SECONDARY list - the list to compare against." & vbCrLf
    txt = txt & "3. Choose a confidence threshold (default 70%)." & vbCrLf
    txt = txt & "4. Click 'Run Comparison'." & vbCrLf & vbCrLf
    txt = txt & "WHAT IT DOES:" & vbCrLf
    txt = txt & "Creates a new results sheet sorted by match confidence." & vbCrLf
    txt = txt & "Each item shows its best match and a confidence %." & vbCrLf & vbCrLf
    txt = txt & "The best-match column is colour-coded:" & vbCrLf
    txt = txt & "  Green  = strong match (80%+)" & vbCrLf
    txt = txt & "  Yellow = moderate match (40-79%)" & vbCrLf
    txt = txt & "  Orange/Red = weak or no match" & vbCrLf & vbCrLf
    txt = txt & "Below the main table:" & vbCrLf
    txt = txt & "- Primary items with no good match" & vbCrLf
    txt = txt & "- Secondary items not matched" & vbCrLf & vbCrLf
    txt = txt & "Your original data is never modified."
    txtInstructions.Value = txt

    ' -- Detect list columns --
    Dim lastCol As Long
    lastCol = m_sourceWs.Cells(1, m_sourceWs.Columns.Count).End(xlToLeft).Column

    m_numLists = 0
    If lastCol >= 1 Then
        ReDim m_colIndices(1 To lastCol)
        ReDim m_colHeaders(1 To lastCol)

        Dim c As Long
        Dim hdr As String
        Dim colLtr As String

        For c = 1 To lastCol
            hdr = Trim$(CStr(m_sourceWs.Cells(1, c).Value))
            If hdr <> "" Then
                m_numLists = m_numLists + 1
                m_colIndices(m_numLists) = c
                m_colHeaders(m_numLists) = hdr

                ' Get column letter (works for A, B, ... AA, AB, etc.)
                colLtr = Split(m_sourceWs.Cells(1, c).Address(True, False), "$")(0)

                cmbPrimary.AddItem colLtr & " - " & hdr
                cmbSecondary.AddItem colLtr & " - " & hdr
            End If
        Next c

        If m_numLists > 0 Then
            ReDim Preserve m_colIndices(1 To m_numLists)
            ReDim Preserve m_colHeaders(1 To m_numLists)
        End If
    End If

    ' -- Threshold options --
    cmbThreshold.AddItem "50%"
    cmbThreshold.AddItem "60%"
    cmbThreshold.AddItem "70%"
    cmbThreshold.AddItem "80%"
    cmbThreshold.AddItem "90%"
    cmbThreshold.ListIndex = 2  ' Default to 70%

    ' -- Disable Run if fewer than 2 lists --
    If m_numLists < 2 Then
        btnRun.Enabled = False
        MsgBox "Need at least 2 columns with headers on the active sheet.", vbExclamation, "FuzzyListCompare"
    End If
End Sub

' -------------------------------------------------------
' btnRun_Click - validate inputs and run comparison
' -------------------------------------------------------
Private Sub btnRun_Click()
    ' Validate primary selection
    If cmbPrimary.ListIndex = -1 Then
        MsgBox "Please select a primary list.", vbExclamation
        Exit Sub
    End If

    ' Validate secondary selection
    If cmbSecondary.ListIndex = -1 Then
        MsgBox "Please select a secondary list.", vbExclamation
        Exit Sub
    End If

    ' Must be different lists
    If cmbPrimary.ListIndex = cmbSecondary.ListIndex Then
        MsgBox "Primary and secondary lists must be different.", vbExclamation
        Exit Sub
    End If

    ' Parse threshold from combo selection
    Dim threshold As Double
    Select Case cmbThreshold.ListIndex
        Case 0: threshold = 0.5
        Case 1: threshold = 0.6
        Case 2: threshold = 0.7
        Case 3: threshold = 0.8
        Case 4: threshold = 0.9
        Case Else: threshold = 0.7
    End Select

    ' Map combo indices (0-based) to our arrays (1-based)
    Dim priIdx As Long
    Dim secIdx As Long
    priIdx = cmbPrimary.ListIndex + 1
    secIdx = cmbSecondary.ListIndex + 1

    ' Hide form and run comparison
    Me.Hide

    RunFuzzyComparison m_sourceWs, m_colIndices(priIdx), m_colIndices(secIdx), m_colHeaders(priIdx), m_colHeaders(secIdx), threshold

    Unload Me
End Sub

' -------------------------------------------------------
' btnClose_Click - close without running
' -------------------------------------------------------
Private Sub btnClose_Click()
    Unload Me
End Sub
