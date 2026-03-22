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
    txt = txt & "IMPORT FILENAMES:" & vbCrLf
    txt = txt & "Click 'Import Filenames' to load file names from a Windows folder "
    txt = txt & "into a column. You can import multiple folders into different "
    txt = txt & "columns before running the comparison." & vbCrLf & vbCrLf
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

    ' -- Threshold options --
    cmbThreshold.AddItem "50%"
    cmbThreshold.AddItem "60%"
    cmbThreshold.AddItem "70%"
    cmbThreshold.AddItem "80%"
    cmbThreshold.AddItem "90%"
    cmbThreshold.ListIndex = 2  ' Default to 70%

    ' -- Detect list columns and populate combos --
    RefreshCombos
End Sub

' -------------------------------------------------------
' RefreshCombos - scan sheet headers, repopulate combos
' -------------------------------------------------------
Private Sub RefreshCombos()
    Dim priIdx As Long
    Dim secIdx As Long
    priIdx = cmbPrimary.ListIndex
    secIdx = cmbSecondary.ListIndex

    cmbPrimary.Clear
    cmbSecondary.Clear

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

    ' Restore selections if still valid
    If priIdx >= 0 And priIdx < cmbPrimary.ListCount Then cmbPrimary.ListIndex = priIdx
    If secIdx >= 0 And secIdx < cmbSecondary.ListCount Then cmbSecondary.ListIndex = secIdx

    btnRun.Enabled = (m_numLists >= 2)
End Sub

' -------------------------------------------------------
' btnImport_Click - import filenames from a folder
' -------------------------------------------------------
Private Sub btnImport_Click()
    ' Step 1: Get target column letter
    Dim colLetter As String
    colLetter = InputBox("Enter column letter to import into (e.g. A, B, C):", "Import Filenames")
    If colLetter = "" Then Exit Sub
    colLetter = UCase$(Trim$(colLetter))

    ' Validate column letter
    Dim targetCol As Long
    targetCol = 0
    On Error Resume Next
    targetCol = m_sourceWs.Range(colLetter & "1").Column
    On Error GoTo 0
    If targetCol = 0 Then
        MsgBox "Invalid column letter: " & colLetter, vbExclamation
        Exit Sub
    End If

    ' Step 2: Open folder picker
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.title = "Select folder containing files to import"
    If fd.Show = 0 Then Exit Sub

    Dim folderPath As String
    folderPath = fd.SelectedItems(1)
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ' Step 3: Read filenames using Dir
    Dim fileNames() As String
    Dim fileCount As Long
    Dim fn As String
    fileCount = 0

    fn = Dir(folderPath & "*.*", vbNormal)
    Do While fn <> ""
        fileCount = fileCount + 1
        ReDim Preserve fileNames(1 To fileCount)
        fileNames(fileCount) = fn
        fn = Dir()
    Loop

    If fileCount = 0 Then
        MsgBox "No files found in the selected folder.", vbExclamation
        Exit Sub
    End If

    ' Step 4: Confirm overwrite if column has data
    Dim existingHdr As String
    existingHdr = Trim$(CStr(m_sourceWs.Cells(1, targetCol).Value))
    If existingHdr <> "" Then
        Dim confirmMsg As String
        confirmMsg = "Column " & colLetter & " already has data (header: " & existingHdr & ")."
        confirmMsg = confirmMsg & vbCrLf & "Clear it and import filenames?"
        If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Import Filenames") <> vbYes Then Exit Sub
    End If

    ' Step 5: Clear column and write data
    Dim lastRow As Long
    lastRow = m_sourceWs.Cells(m_sourceWs.Rows.Count, targetCol).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1
    m_sourceWs.Range(m_sourceWs.Cells(1, targetCol), m_sourceWs.Cells(lastRow, targetCol)).ClearContents

    ' Folder name as header
    Dim trimmedPath As String
    trimmedPath = Left$(folderPath, Len(folderPath) - 1)
    Dim folderName As String
    folderName = Mid$(trimmedPath, InStrRev(trimmedPath, "\") + 1)
    m_sourceWs.Cells(1, targetCol).Value = folderName

    ' Filenames in rows 2+
    Dim r As Long
    For r = 1 To fileCount
        m_sourceWs.Cells(r + 1, targetCol).Value = fileNames(r)
    Next r

    ' Step 6: Refresh combos and notify
    RefreshCombos

    Dim doneMsg As String
    doneMsg = "Imported " & fileCount & " filenames into column " & colLetter & "."
    doneMsg = doneMsg & vbCrLf & "Header set to: " & folderName
    MsgBox doneMsg, vbInformation, "Import Filenames"
End Sub

' -------------------------------------------------------
' btnRun_Click - validate inputs and run comparison
' -------------------------------------------------------
Private Sub btnRun_Click()
    If cmbPrimary.ListIndex = -1 Then
        MsgBox "Please select a primary list.", vbExclamation
        Exit Sub
    End If

    If cmbSecondary.ListIndex = -1 Then
        MsgBox "Please select a secondary list.", vbExclamation
        Exit Sub
    End If

    If cmbPrimary.ListIndex = cmbSecondary.ListIndex Then
        MsgBox "Primary and secondary lists must be different.", vbExclamation
        Exit Sub
    End If

    Dim threshold As Double
    Select Case cmbThreshold.ListIndex
        Case 0: threshold = 0.5
        Case 1: threshold = 0.6
        Case 2: threshold = 0.7
        Case 3: threshold = 0.8
        Case 4: threshold = 0.9
        Case Else: threshold = 0.7
    End Select

    Dim priIdx As Long
    Dim secIdx As Long
    priIdx = cmbPrimary.ListIndex + 1
    secIdx = cmbSecondary.ListIndex + 1

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
