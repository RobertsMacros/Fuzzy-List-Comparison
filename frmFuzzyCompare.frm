VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFuzzyCompare
   Caption         =   "Fuzzy List Compare"
   ClientHeight    =   7740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFuzzyCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' -------------------------------------------------------
' The controls are created in code (BuildControls), so importing this
' .frm is all that is needed: no Toolbox work, no .frx file.
' -------------------------------------------------------
Private WithEvents mBtnImport As MSForms.CommandButton
Private WithEvents mBtnRun As MSForms.CommandButton
Private WithEvents mBtnClose As MSForms.CommandButton
Private mTxtInstructions As MSForms.TextBox
Private mCmbPrimary As MSForms.ComboBox
Private mCmbSecondary As MSForms.ComboBox
Private mCmbThreshold As MSForms.ComboBox

' -------------------------------------------------------
' Module-level storage for detected lists
' -------------------------------------------------------
Private m_colIndices() As Long
Private m_colHeaders() As String
Private m_numLists As Long
Private m_sourceWs As Worksheet

' -------------------------------------------------------
' GetOrAdd - reuse a control drawn at design time, else create it
' -------------------------------------------------------
Private Function GetOrAdd(ByVal progId As String, ByVal ctlName As String) As Object
    Dim ctl As Object
    On Error Resume Next
    Set ctl = Me.Controls(ctlName)
    On Error GoTo 0
    If ctl Is Nothing Then Set ctl = Me.Controls.Add(progId, ctlName, True)
    Set GetOrAdd = ctl
End Function

Private Sub AddLabel(ByVal ctlName As String, ByVal cap As String, ByVal y As Single)
    Dim lbl As Object
    Set lbl = GetOrAdd("Forms.Label.1", ctlName)
    lbl.Caption = cap
    lbl.Left = 12
    lbl.Top = y
    lbl.Width = 330
    lbl.Height = 14
End Sub

Private Sub PlaceCombo(ByVal cmb As MSForms.ComboBox, ByVal y As Single, ByVal w As Single)
    cmb.Left = 12
    cmb.Top = y
    cmb.Width = w
    cmb.Height = 18
    cmb.Style = fmStyleDropDownList
End Sub

Private Sub PlaceButton(ByVal btn As MSForms.CommandButton, ByVal cap As String, ByVal x As Single, ByVal w As Single)
    btn.Caption = cap
    btn.Left = x
    btn.Top = 350
    btn.Width = w
    btn.Height = 26
End Sub

' -------------------------------------------------------
' BuildControls - lay out the window
' -------------------------------------------------------
Private Sub BuildControls()
    Me.Caption = "Fuzzy List Compare"
    Me.Width = 366
    Me.Height = 410

    Set mTxtInstructions = GetOrAdd("Forms.TextBox.1", "txtInstructions")
    With mTxtInstructions
        .Left = 12
        .Top = 12
        .Width = 330
        .Height = 190
        .MultiLine = True
        .WordWrap = True
        .ScrollBars = fmScrollBarsVertical
        .Locked = True
        .BackColor = Me.BackColor
        .TabStop = False
    End With

    AddLabel "lblPrimary", "Primary list (the list to check):", 212
    Set mCmbPrimary = GetOrAdd("Forms.ComboBox.1", "cmbPrimary")
    PlaceCombo mCmbPrimary, 226, 330

    AddLabel "lblSecondary", "Secondary list (the list to compare against):", 254
    Set mCmbSecondary = GetOrAdd("Forms.ComboBox.1", "cmbSecondary")
    PlaceCombo mCmbSecondary, 268, 330

    AddLabel "lblThreshold", "Match threshold:", 296
    Set mCmbThreshold = GetOrAdd("Forms.ComboBox.1", "cmbThreshold")
    PlaceCombo mCmbThreshold, 310, 80

    Set mBtnImport = GetOrAdd("Forms.CommandButton.1", "btnImport")
    PlaceButton mBtnImport, "Import Filenames...", 12, 110
    Set mBtnRun = GetOrAdd("Forms.CommandButton.1", "btnRun")
    PlaceButton mBtnRun, "Run Comparison", 168, 96
    mBtnRun.Default = True
    Set mBtnClose = GetOrAdd("Forms.CommandButton.1", "btnClose")
    PlaceButton mBtnClose, "Close", 272, 70
    mBtnClose.Cancel = True
End Sub

' -------------------------------------------------------
' UserForm_Initialize - detect lists, populate controls
' -------------------------------------------------------
Private Sub UserForm_Initialize()
    Set m_sourceWs = ActiveSheet
    BuildControls

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
    mTxtInstructions.Value = txt

    ' -- Threshold options --
    Dim pct As Variant
    For Each pct In Array("40%", "50%", "60%", "70%", "80%", "90%")
        mCmbThreshold.AddItem pct
    Next pct
    mCmbThreshold.ListIndex = 3  ' Default to 70%

    ' -- Detect list columns and populate combos --
    RefreshCombos
End Sub

' -------------------------------------------------------
' RefreshCombos - scan sheet headers, repopulate combos
' -------------------------------------------------------
Private Sub RefreshCombos()
    Dim priIdx As Long
    Dim secIdx As Long
    priIdx = mCmbPrimary.ListIndex
    secIdx = mCmbSecondary.ListIndex

    mCmbPrimary.Clear
    mCmbSecondary.Clear

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
                mCmbPrimary.AddItem colLtr & " - " & hdr
                mCmbSecondary.AddItem colLtr & " - " & hdr
            End If
        Next c

        If m_numLists > 0 Then
            ReDim Preserve m_colIndices(1 To m_numLists)
            ReDim Preserve m_colHeaders(1 To m_numLists)
        End If
    End If

    ' Restore selections if still valid, else pick the first two lists
    If priIdx < 0 Then priIdx = 0
    If secIdx < 0 Then secIdx = 1
    If priIdx < mCmbPrimary.ListCount Then mCmbPrimary.ListIndex = priIdx
    If secIdx < mCmbSecondary.ListCount Then mCmbSecondary.ListIndex = secIdx

    mBtnRun.Enabled = (m_numLists >= 2)
End Sub

' -------------------------------------------------------
' mBtnImport_Click - import filenames from a folder
' -------------------------------------------------------
Private Sub mBtnImport_Click()
#If Mac Then
    MsgBox "Import Filenames works in Excel for Windows only. On a Mac, paste the file names into a column instead.", vbInformation, "Import Filenames"
    Exit Sub
#End If

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

    ' Filenames in rows 2+, as text so "001.pdf" or "=x" stay as typed
    m_sourceWs.Range(m_sourceWs.Cells(2, targetCol), m_sourceWs.Cells(fileCount + 1, targetCol)).NumberFormat = "@"
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
' mBtnRun_Click - validate inputs and run comparison
' -------------------------------------------------------
Private Sub mBtnRun_Click()
    If mCmbPrimary.ListIndex = -1 Then
        MsgBox "Please select a primary list.", vbExclamation
        Exit Sub
    End If

    If mCmbSecondary.ListIndex = -1 Then
        MsgBox "Please select a secondary list.", vbExclamation
        Exit Sub
    End If

    If mCmbPrimary.ListIndex = mCmbSecondary.ListIndex Then
        MsgBox "Primary and secondary lists must be different.", vbExclamation
        Exit Sub
    End If

    Dim threshold As Double
    threshold = 0.7
    If mCmbThreshold.ListIndex >= 0 Then threshold = Val(Replace(mCmbThreshold.Value, "%", "")) / 100

    Dim priIdx As Long
    Dim secIdx As Long
    priIdx = mCmbPrimary.ListIndex + 1
    secIdx = mCmbSecondary.ListIndex + 1

    Me.Hide

    RunFuzzyComparison m_sourceWs, m_colIndices(priIdx), m_colIndices(secIdx), m_colHeaders(priIdx), m_colHeaders(secIdx), threshold

    Unload Me
End Sub

' -------------------------------------------------------
' mBtnClose_Click - close without running
' -------------------------------------------------------
Private Sub mBtnClose_Click()
    Unload Me
End Sub
