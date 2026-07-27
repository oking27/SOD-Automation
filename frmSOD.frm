VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSOD 
   Caption         =   "SOD Editor"
   ClientHeight    =   7259
   ClientLeft      =   273
   ClientTop       =   973
   ClientWidth     =   8715.001
   OleObjectBlob   =   "frmSOD.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmSOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'====================================================
' SOD FORM
' Manages section data and writes it to the Excel
' table for use by PopulateSOD.
'====================================================

' Special section names in display order
Private Const SEC_TITLE As String = "Title"
Private Const SEC_GROUP As String = "Group"
Private Const SEC_PURPOSE As String = "Purpose"
Private Const SEC_SCOPE As String = "Scope"
Private Const SEC_DICTIONARY As String = "Dictionary"
Private Const SEC_ROLES As String = "Roles"
Private Const SEC_OBJECTIVES As String = "Objectives"
Private Const SEC_STEPS As String = "Process Steps"
Private Const SEC_KPIS As String = "Key Performance Indicators"
Private Const SEC_RESOURCES As String = "Additional Resources"

Private Const SEC_RESOURCES_COL1 As String = "Resource Type"
Private Const SEC_RESOURCES_COL2 As String = "List or Document Link"
Private Const SEC_RESOURCES_COL3 As String = "Document Number"

Private Const SEC_DICT_COL1 As String = "Term"
Private Const SEC_DICT_COL2 As String = "Definition"

' Section types
Private Const TYPE_PLAIN As String = "PLAIN"           ' single TextBox
Private Const TYPE_GROUP As String = "GROUP"           ' single dropdown, metadata only
Private Const TYPE_LIST As String = "LIST"             ' covers both simple and nested lists
Private Const TYPE_TABLE As String = "TABLE"           ' variable column count + rows
Private Const TYPE_RESOURCES As String = "RESOURCES"   ' fixed 3-column table
Private Const TYPE_DICTIONARY As String = "DICTIONARY" ' fixed 2-column table

' Data store: parallel arrays indexed by section
Private mSecNames() As String        ' display name
Private mSecTypes() As String        ' one of the TYPE_ constants above
Private mSecCount As Long

' Each section's data stored as pipe-delimited strings
' PLAIN/GROUP: mSecData(i) = full text
' LIST:        mSecItems(i) = Collection of "item" (no subs) or
'              "item|sub1~sub2~sub3..." (mSecHasSubItems(i) = True)
' TABLE/RESOURCES/DICTIONARY:
'              mSecItems(i) = Collection of "col1val|col2val|col3val"
'              mSecCols(i) stores the column count as a tilde-delimited
'              placeholder string (e.g. "col1~col2~col3")
Private mSecData() As String
Private mSecItems() As Collection
Private mSecCols() As String
Private mSecHasHeader() As Boolean     ' TYPE_TABLE only: is row 1 a header?
Private mSecHasSubItems() As Boolean   ' TYPE_LIST only: does this list have sub-items?

Private mCurrentSection As Long      ' 1-based index of selected section
Private mLoading As Boolean          ' suppress change events during load
Private mJustSaved As Boolean        ' skip discard-confirmation right after a save

' Edit-mode state (0/-1 = not editing)
Private mEditingItemIdx As Long      ' 0 = not editing, otherwise 1-based item index
Private mEditingSubIdx As Long       ' -1 = not editing, otherwise 0-based sub-item index
Private mEditingRowIdx As Long       ' 0 = not editing, otherwise 1-based row index
Private mEditingSectionIdx As Long   ' 0 = not editing/reordering a section
Private mAddingSection As Boolean    ' True while naming a brand-new section

' Undo (single-level, scoped to the section it happened in)
Private mUndoAvailable As Boolean
Private mUndoSecIdx As Long
Private mUndoItems As Collection
Private mUndoData As String
Private mUndoCols As String

' Pre-edit snapshots, used so Cancel can fully restore state
' (separate from Undo, which only covers the most recent committed change)
Private mEditSnapshot As Collection
Private mEditSnapshotSec As Long

Private mSubEditSnapshot As Collection
Private mSubEditSnapshotSec As Long
Private mSubEditSnapshotItemIdx As Long

Private mRowEditSnapshot As Collection
Private mRowEditSnapshotSec As Long

'====================================================
' CENTRAL ERROR HANDLER
' Every event handler and data function routes runtime
' errors here instead of letting VBA show its own dialog
' and potentially leave the form in a broken state.
'====================================================

Private Sub HandleFormError(ByVal procName As String)

    MsgBox "Something went wrong in '" & procName & "'." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Your other entries are safe. Please try again, and if this repeats, " & _
           "note down what you were doing and report it.", _
           vbExclamation, "SOD Form - Error"

End Sub

Private Sub fraContent_Click()

End Sub

Private Sub lblSubItems_Click()

End Sub

'====================================================
' FORM LIFECYCLE
'====================================================

Private Sub UserForm_Initialize()

    On Error GoTo ErrHandler

    mLoading = True
    mEditingItemIdx = 0
    mEditingSubIdx = -1
    mEditingRowIdx = 0
    mEditingSectionIdx = 0
    mAddingSection = False
    mJustSaved = False
    mUndoAvailable = False

    Set mEditSnapshot = Nothing
    Set mSubEditSnapshot = Nothing
    Set mRowEditSnapshot = Nothing

    Dim ws As Worksheet
    Set ws = ActiveSheet

    If ws.ListObjects.count > 0 Then
        If ws.ListObjects(1).ListRows.count > 0 Then
            mSecCount = 0
            LoadFromSheet
        Else
            InitSections
        End If
    Else
        InitSections
    End If

    PopulateGroupDropdown
    RefreshSectionList

    If mSecCount > 0 Then
        mCurrentSection = 1
        lstSections.ListIndex = 0
        ShowSection 1
    End If

    UpdateSectionButtons
    mLoading = False
    RefreshUndoButton
    UpdateFormCaption

    Exit Sub

ErrHandler:
    mLoading = False
    HandleFormError "UserForm_Initialize"

End Sub

Private Sub InitSections()

    On Error GoTo ErrHandler

    mSecCount = 10
    ReDim mSecNames(1 To mSecCount)
    ReDim mSecTypes(1 To mSecCount)
    ReDim mSecData(1 To mSecCount)
    ReDim mSecItems(1 To mSecCount)
    ReDim mSecCols(1 To mSecCount)
    ReDim mSecHasHeader(1 To mSecCount)
    ReDim mSecHasSubItems(1 To mSecCount)

    mSecNames(1) = SEC_TITLE
    mSecNames(2) = SEC_GROUP
    mSecNames(3) = SEC_PURPOSE
    mSecNames(4) = SEC_SCOPE
    mSecNames(5) = SEC_DICTIONARY
    mSecNames(6) = SEC_ROLES
    mSecNames(7) = SEC_OBJECTIVES
    mSecNames(8) = SEC_STEPS
    mSecNames(9) = SEC_KPIS
    mSecNames(10) = SEC_RESOURCES

    mSecTypes(1) = TYPE_PLAIN
    mSecTypes(2) = TYPE_GROUP
    mSecTypes(3) = TYPE_PLAIN
    mSecTypes(4) = TYPE_DICTIONARY
    mSecTypes(5) = TYPE_LIST
    mSecTypes(6) = TYPE_LIST
    mSecTypes(7) = TYPE_LIST
    mSecTypes(8) = TYPE_LIST
    mSecTypes(9) = TYPE_RESOURCES

    Dim i As Long
    For i = 1 To mSecCount
        mSecData(i) = ""
        mSecCols(i) = ""
        mSecHasHeader(i) = False
        mSecHasSubItems(i) = IsBuiltInSubItemSection(mSecNames(i))
        Set mSecItems(i) = New Collection
    Next i

    Exit Sub

ErrHandler:
    HandleFormError "InitSections"

End Sub

Private Sub UpdateFormCaption()

    On Error GoTo ErrHandler

    Dim titleSecIdx As Long
    titleSecIdx = FindSection(SEC_TITLE)

    Dim titleText As String
    If titleSecIdx > 0 Then
        titleText = Trim(mSecData(titleSecIdx))
    End If

    If titleText <> "" Then
        Me.Caption = "SOD Editor: " & titleText
    Else
        Me.Caption = "SOD Editor"
    End If

    Exit Sub

ErrHandler:
    HandleFormError "UpdateFormCaption"

End Sub

Private Sub PopulateGroupDropdown()

    On Error GoTo ErrHandler

    cboGroup.Clear

    ' TODO: replace with your real Group options
    cboGroup.AddItem "Group A"
    cboGroup.AddItem "Group B"
    cboGroup.AddItem "Group C"

    Dim groupIdx As Long
    groupIdx = FindSection(SEC_GROUP)
    If groupIdx > 0 Then
        cboGroup.Value = mSecData(groupIdx)
    End If

    Exit Sub

ErrHandler:
    HandleFormError "PopulateGroupDropdown"

End Sub

Public Sub OpenSODEditor()
    frmSOD.Show
End Sub

'====================================================
' UNDO (single-level, scoped to the section it happened in)
'====================================================

Private Sub PushUndo()

    On Error GoTo ErrHandler

    mJustSaved = False

    mUndoSecIdx = mCurrentSection
    Set mUndoItems = CloneCollection(mSecItems(mCurrentSection))
    mUndoData = mSecData(mCurrentSection)
    mUndoCols = mSecCols(mCurrentSection)
    mUndoAvailable = True

    RefreshUndoButton

    Exit Sub

ErrHandler:
    HandleFormError "PushUndo"

End Sub

Private Sub RefreshUndoButton()

    On Error Resume Next
    btnUndo.Enabled = mUndoAvailable
    On Error GoTo 0

End Sub

Private Function CloneCollection(ByVal src As Collection) As Collection

    Dim result As New Collection
    Dim i As Long
    For i = 1 To src.count
        result.Add src(i)
    Next i
    Set CloneCollection = result

End Function

Private Sub btnUndo_Click()

    On Error GoTo ErrHandler

    If Not mUndoAvailable Then
        MsgBox "Nothing to undo.", vbInformation
        Exit Sub
    End If

    If mUndoSecIdx <> mCurrentSection Then
        MsgBox "The last change was in '" & mSecNames(mUndoSecIdx) & _
               "'. Switch to that section to undo it.", vbExclamation
        Exit Sub
    End If

    Set mSecItems(mCurrentSection) = CloneCollection(mUndoItems)
    mSecData(mCurrentSection) = mUndoData
    mSecCols(mCurrentSection) = mUndoCols

    mUndoAvailable = False
    RefreshUndoButton

    ShowSection mCurrentSection

    Exit Sub

ErrHandler:
    HandleFormError "btnUndo_Click"

End Sub

'====================================================
' SAVE CURRENT SECTION DATA TO MEMORY
'====================================================

Private Sub SaveCurrentSection()

    On Error GoTo ErrHandler

    If mCurrentSection < 1 Or mCurrentSection > mSecCount Then Exit Sub

    Select Case mSecTypes(mCurrentSection)

        Case TYPE_PLAIN
            mSecData(mCurrentSection) = txtContent.Text
            If mCurrentSection = FindSection(SEC_TITLE) Then UpdateFormCaption

        Case TYPE_GROUP
            mSecData(mCurrentSection) = cboGroup.Value

        Case TYPE_LIST
            ' Nothing to save here - data lives in mSecItems,
            ' updated live by Add/Remove/Edit buttons

        Case TYPE_TABLE, TYPE_RESOURCES, TYPE_DICTIONARY
            SaveRowEditor

    End Select

    Exit Sub

ErrHandler:
    HandleFormError "SaveCurrentSection"

End Sub
'====================================================
' SECTION LIST
'====================================================

Private Sub RefreshSectionList()

    On Error GoTo ErrHandler

    Dim sel As Long
    sel = lstSections.ListIndex

    lstSections.Clear

    Dim i As Long
    For i = 1 To mSecCount
        lstSections.AddItem mSecNames(i)
    Next i

    If sel >= 0 And sel < lstSections.ListCount Then
        lstSections.ListIndex = sel
    End If

    UpdateSectionButtons

    Exit Sub

ErrHandler:
    HandleFormError "RefreshSectionList"

End Sub

Private Sub lstSections_Click()

    On Error GoTo ErrHandler

    If mLoading Then Exit Sub
    If lstSections.ListIndex < 0 Then Exit Sub

    CancelAnyActiveEditMode

    SaveCurrentSection

    mCurrentSection = lstSections.ListIndex + 1
    ShowSection mCurrentSection
    UpdateSectionButtons

    Exit Sub

ErrHandler:
    HandleFormError "lstSections_Click"

End Sub

Private Sub UpdateSectionButtons()

    On Error GoTo ErrHandler

    Dim selIdx As Long
    selIdx = lstSections.ListIndex

    If selIdx < 0 Then
        btnRemoveSection.Enabled = False
        btnEditSection.Enabled = False
        Exit Sub
    End If

    Dim secIdx As Long
    secIdx = selIdx + 1

    Dim isBuiltIn As Boolean
    isBuiltIn = IsBuiltInSectionName(mSecNames(secIdx))

    btnRemoveSection.Enabled = Not isBuiltIn
    btnEditSection.Enabled = True   ' rename is still allowed, with a warning

    Exit Sub

ErrHandler:
    HandleFormError "UpdateSectionButtons"

End Sub

'====================================================
' BUILT-IN SECTION PROTECTION
'====================================================

Private Function IsBuiltInSectionName(ByVal name As String) As Boolean

    Select Case LCase(Trim(name))
        Case LCase(SEC_TITLE), LCase(SEC_GROUP), LCase(SEC_PURPOSE), LCase(SEC_SCOPE), _
             LCase(SEC_DICTIONARY), LCase(SEC_ROLES), LCase(SEC_OBJECTIVES), _
             LCase(SEC_STEPS), LCase(SEC_KPIS), LCase(SEC_RESOURCES)
            IsBuiltInSectionName = True
    End Select

End Function

Private Function IsBuiltInSubItemSection(ByVal name As String) As Boolean
    ' Built-in sections that always have sub-items - checkbox never shown
    Select Case LCase(Trim(name))
        Case LCase(SEC_ROLES), LCase(SEC_STEPS), LCase(SEC_RESOURCES)
            IsBuiltInSubItemSection = True
    End Select
End Function

Private Function IsBuiltInListSection(ByVal name As String) As Boolean
    ' Built-in sections that are lists but never have sub-items -
    ' checkbox never shown, always unchecked
    Select Case LCase(Trim(name))
        Case LCase(SEC_OBJECTIVES), LCase(SEC_KPIS)
            IsBuiltInListSection = True
    End Select
End Function

'====================================================
' REORDERING SECTIONS
'====================================================

Private Sub SwapSections(ByVal idx1 As Long, ByVal idx2 As Long)

    Dim tmpName As String, tmpType As String, tmpData As String, tmpCols As String
    Dim tmpHasHeader As Boolean, tmpHasSubItems As Boolean
    Dim tmpItems As Collection

    tmpName = mSecNames(idx1): mSecNames(idx1) = mSecNames(idx2): mSecNames(idx2) = tmpName
    tmpType = mSecTypes(idx1): mSecTypes(idx1) = mSecTypes(idx2): mSecTypes(idx2) = tmpType
    tmpData = mSecData(idx1): mSecData(idx1) = mSecData(idx2): mSecData(idx2) = tmpData
    tmpCols = mSecCols(idx1): mSecCols(idx1) = mSecCols(idx2): mSecCols(idx2) = tmpCols
    tmpHasHeader = mSecHasHeader(idx1): mSecHasHeader(idx1) = mSecHasHeader(idx2): mSecHasHeader(idx2) = tmpHasHeader
    tmpHasSubItems = mSecHasSubItems(idx1): mSecHasSubItems(idx1) = mSecHasSubItems(idx2): mSecHasSubItems(idx2) = tmpHasSubItems

    Set tmpItems = mSecItems(idx1)
    Set mSecItems(idx1) = mSecItems(idx2)
    Set mSecItems(idx2) = tmpItems

    ' Keep mCurrentSection pointed at whichever slot now holds the
    ' section that was actually being displayed
    If mCurrentSection = idx1 Then
        mCurrentSection = idx2
    ElseIf mCurrentSection = idx2 Then
        mCurrentSection = idx1
    End If

End Sub

Private Sub RemoveSectionAt(ByVal secIdx As Long)

    Dim i As Long
    For i = secIdx To mSecCount - 1
        mSecNames(i) = mSecNames(i + 1)
        mSecTypes(i) = mSecTypes(i + 1)
        mSecData(i) = mSecData(i + 1)
        mSecCols(i) = mSecCols(i + 1)
        mSecHasHeader(i) = mSecHasHeader(i + 1)
        mSecHasSubItems(i) = mSecHasSubItems(i + 1)
        Set mSecItems(i) = mSecItems(i + 1)
    Next i

    mSecCount = mSecCount - 1

    ReDim Preserve mSecNames(1 To mSecCount)
    ReDim Preserve mSecTypes(1 To mSecCount)
    ReDim Preserve mSecData(1 To mSecCount)
    ReDim Preserve mSecCols(1 To mSecCount)
    ReDim Preserve mSecHasHeader(1 To mSecCount)
    ReDim Preserve mSecHasSubItems(1 To mSecCount)
    ReDim Preserve mSecItems(1 To mSecCount)

End Sub

'====================================================
' ADD / EDIT / REMOVE SECTION
' Uses the same Add->Move Up / Edit->Confirm / Remove->Move Down
' pattern as items, sub-items, and rows elsewhere in the form.
' While naming a NEW section, the three buttons temporarily become
' Text Section / List Section / Table Section instead.
'====================================================

Private Sub btnAddSection_Click()

    On Error GoTo ErrHandler

    If mAddingSection Then
        CreateNewSection TYPE_PLAIN
        Exit Sub
    End If

    If mEditingSectionIdx > 0 Then

        ' EDIT MODE: this button is "Move Up"
        If mEditingSectionIdx <= 1 Then Exit Sub

        PushUndo
        SwapSections mEditingSectionIdx, mEditingSectionIdx - 1
        mEditingSectionIdx = mEditingSectionIdx - 1

        RefreshSectionList
        lstSections.ListIndex = mEditingSectionIdx - 1

        Exit Sub

    End If

    ' Enter "naming a new section" mode
    mAddingSection = True

    txtSectionName.Text = ""
    txtSectionName.Visible = True
    txtSectionName.SetFocus

    btnAddSection.Caption = "Text Section"
    btnEditSection.Caption = "List Section"
    btnRemoveSection.Caption = "Table Section"
    btnCancel.Caption = "Cancel"

    Exit Sub

ErrHandler:
    HandleFormError "btnAddSection_Click"

End Sub

Private Sub btnEditSection_Click()

    On Error GoTo ErrHandler

    If mAddingSection Then
        CreateNewSection TYPE_LIST
        Exit Sub
    End If

    If mEditingSectionIdx > 0 Then

        ' CONFIRM the rename
        Dim newName As String
        newName = Trim(txtSectionName.Text)

        If newName = "" Then
            MsgBox "Section name cannot be blank.", vbExclamation
            Exit Sub
        End If

        Dim currentName As String
        currentName = mSecNames(mEditingSectionIdx)

        If IsBuiltInSectionName(currentName) And LCase(newName) <> LCase(currentName) Then
            Dim resp As VbMsgBoxResult
            resp = MsgBox("'" & currentName & "' is a built-in section with special Word formatting." & vbCrLf & vbCrLf & _
                          "Renaming it will cause it to lose that special formatting when generating the document." & vbCrLf & vbCrLf & _
                          "Continue renaming anyway?", vbYesNo + vbExclamation, "Built-in Section")
            If resp = vbNo Then Exit Sub
        End If

        PushUndo
        mSecNames(mEditingSectionIdx) = newName

        RefreshSectionList

        Dim savedIdx As Long
        savedIdx = mEditingSectionIdx

        mEditingSectionIdx = 0
        btnEditSection.Caption = "Edit Section"
        btnAddSection.Caption = "+ Add Section"
        btnRemoveSection.Caption = "- Remove Section"
        btnCancel.Caption = "Close Editor"
        txtSectionName.Visible = False

        lstSections.ListIndex = savedIdx - 1

        If savedIdx = mCurrentSection Then
            lblSectionTitle.Caption = newName
        End If

    Else

        ' ENTER rename/reorder mode
        Dim selIdx As Long
        selIdx = lstSections.ListIndex

        If selIdx < 0 Then
            MsgBox "Select a section to edit.", vbExclamation
            Exit Sub
        End If

        Dim secIdx As Long
        secIdx = selIdx + 1

        txtSectionName.Text = mSecNames(secIdx)
        txtSectionName.Visible = True
        txtSectionName.SetFocus

        mEditingSectionIdx = secIdx
        btnEditSection.Caption = "Confirm"
        btnAddSection.Caption = "Move Up"
        btnRemoveSection.Caption = "Move Down"
        btnCancel.Caption = "Cancel"

    End If

    Exit Sub

ErrHandler:
    HandleFormError "btnEditSection_Click"

End Sub

Private Sub btnRemoveSection_Click()

    On Error GoTo ErrHandler

    If mAddingSection Then
        CreateNewSection TYPE_TABLE
        Exit Sub
    End If

    If mEditingSectionIdx > 0 Then

        ' EDIT MODE: this button is "Move Down"
        If mEditingSectionIdx >= mSecCount Then Exit Sub

        PushUndo
        SwapSections mEditingSectionIdx, mEditingSectionIdx + 1
        mEditingSectionIdx = mEditingSectionIdx + 1

        RefreshSectionList
        lstSections.ListIndex = mEditingSectionIdx - 1

        Exit Sub

    End If

    Dim selIdx As Long
    selIdx = lstSections.ListIndex

    If selIdx < 0 Then
        MsgBox "Select a section to remove.", vbExclamation
        Exit Sub
    End If

    Dim secIdx As Long
    secIdx = selIdx + 1

    If IsBuiltInSectionName(mSecNames(secIdx)) Then
        MsgBox "'" & mSecNames(secIdx) & "' is a built-in section and cannot be removed.", vbExclamation
        Exit Sub
    End If

    If mSecCount <= 1 Then
        MsgBox "Cannot remove the only remaining section.", vbExclamation
        Exit Sub
    End If

    Dim resp As VbMsgBoxResult
    resp = MsgBox("Remove section '" & mSecNames(secIdx) & "'?" & vbCrLf & vbCrLf & _
                  "This permanently deletes all of its data and cannot be undone. Consider saving first.", _
                  vbYesNo + vbExclamation, "Remove Section")
    If resp = vbNo Then Exit Sub

    RemoveSectionAt secIdx

    RefreshSectionList

    Dim newSel As Long
    newSel = secIdx - 1
    If newSel < 0 Then newSel = 0
    If newSel > mSecCount - 1 Then newSel = mSecCount - 1

    lstSections.ListIndex = newSel
    mCurrentSection = newSel + 1
    ShowSection mCurrentSection

    Exit Sub

ErrHandler:
    HandleFormError "btnRemoveSection_Click"

End Sub

Private Sub CreateNewSection(ByVal secType As String)

    On Error GoTo ErrHandler

    Dim secName As String
    secName = Trim(txtSectionName.Text)

    If secName = "" Then
        MsgBox "Please type a section name first.", vbExclamation
        Exit Sub
    End If

    mSecCount = mSecCount + 1
    ReDim Preserve mSecNames(1 To mSecCount)
    ReDim Preserve mSecTypes(1 To mSecCount)
    ReDim Preserve mSecData(1 To mSecCount)
    ReDim Preserve mSecItems(1 To mSecCount)
    ReDim Preserve mSecCols(1 To mSecCount)
    ReDim Preserve mSecHasHeader(1 To mSecCount)
    ReDim Preserve mSecHasSubItems(1 To mSecCount)

    mSecNames(mSecCount) = secName
    mSecTypes(mSecCount) = secType
    mSecData(mSecCount) = ""
    mSecCols(mSecCount) = ""
    mSecHasHeader(mSecCount) = False
    mSecHasSubItems(mSecCount) = False
    Set mSecItems(mSecCount) = New Collection

    RefreshSectionList
    SaveCurrentSection

    mAddingSection = False
    txtSectionName.Visible = False
    btnAddSection.Caption = "+ Add Section"
    btnEditSection.Caption = "Edit Section"
    btnRemoveSection.Caption = "- Remove Section"
    btnCancel.Caption = "Close Editor"

    lstSections.ListIndex = mSecCount - 1
    mCurrentSection = mSecCount
    ShowSection mCurrentSection

    Exit Sub

ErrHandler:
    HandleFormError "CreateNewSection"

End Sub
'====================================================
' SHOW SECTION - switches the right panel content
'====================================================

Private Sub ShowSection(ByVal idx As Long)

    On Error GoTo ErrHandler

    If idx < 1 Or idx > mSecCount Then Exit Sub

    mLoading = True

    ' Reset every edit/add mode and button caption before showing
    ' whatever section is being switched to - prevents a half-finished
    ' edit in one section from leaking into another
    mEditingItemIdx = 0
    mEditingSubIdx = -1
    mEditingRowIdx = 0
    mEditingSectionIdx = 0

    btnAddItem.Caption = "+ Add"
    btnRemoveItem.Caption = "- Remove"
    btnEditItem.Caption = "Edit"

    btnAddSubItem.Caption = "+ Add"
    btnRemoveSubItem.Caption = "- Remove"
    btnEditSubItem.Caption = "Edit"

    btnAddRow.Caption = "+ Add"
    btnRemoveRow.Caption = "- Remove"
    btnEditRow.Caption = "Edit"

    btnEditSection.Caption = "Edit Section"
    btnAddSection.Caption = "+ Add Section"
    btnRemoveSection.Caption = "- Remove Section"
    txtSectionName.Visible = False

    btnCancel.Caption = "Close Editor"

    Set mEditSnapshot = Nothing
    Set mSubEditSnapshot = Nothing
    Set mRowEditSnapshot = Nothing

    ' Start every button disabled; each Show*Section sub below re-enables
    ' whatever's actually applicable once it knows the real selection state
    btnEditItem.Enabled = False
    btnRemoveItem.Enabled = False
    txtSubItem.Enabled = False
    btnAddSubItem.Enabled = False
    btnEditSubItem.Enabled = False
    btnRemoveSubItem.Enabled = False
    btnEditRow.Enabled = False
    btnRemoveRow.Enabled = False

    lblSectionTitle.Caption = mSecNames(idx)

    HideAllControls

    Select Case mSecTypes(idx)

        Case TYPE_PLAIN
            ShowPlainSection idx

        Case TYPE_GROUP
            ShowGroupSection idx

        Case TYPE_LIST
            ShowListSection idx

        Case TYPE_TABLE
            ShowTableSection idx

        Case TYPE_RESOURCES
            ShowResourcesSection idx

        Case TYPE_DICTIONARY
            ShowDictionarySection idx

    End Select

    mLoading = False

    Exit Sub

ErrHandler:
    mLoading = False
    HandleFormError "ShowSection"

End Sub

Private Sub HideAllControls()

    On Error GoTo ErrHandler

    ' Plain / Group
    txtContent.Visible = False
    cboGroup.Visible = False

    ' List (items + optional sub-items)
    lstItems.Visible = False
    btnAddItem.Visible = False
    btnRemoveItem.Visible = False
    btnEditItem.Visible = False
    txtItem.Visible = False
    chkSubItems.Visible = False
    lblSubItems.Visible = False
    lstSubItems.Visible = False
    btnAddSubItem.Visible = False
    btnRemoveSubItem.Visible = False
    btnEditSubItem.Visible = False
    txtSubItem.Visible = False

    ' Generic Table
    lblColCount.Visible = False
    spnColCount.Visible = False
    lblColCountValue.Visible = False
    chkHasHeader.Visible = False
    fraRowEditor.Visible = False

    ' Shared by Table / Resources / Dictionary
    lstRows.Visible = False
    btnAddRow.Visible = False
    btnRemoveRow.Visible = False
    btnEditRow.Visible = False
    lblRowEditor.Visible = False

    ' Resources fixed fields
    lblResCol1.Visible = False
    txtResCol1.Visible = False
    lblResCol2.Visible = False
    txtResCol2.Visible = False
    lblResCol3.Visible = False
    txtResCol3.Visible = False

    ' Dictionary fixed fields
    lblDictCol1.Visible = False
    txtDictCol1.Visible = False
    lblDictCol2.Visible = False
    txtDictCol2.Visible = False

    Exit Sub

ErrHandler:
    HandleFormError "HideAllControls"

End Sub

Private Sub ShowPlainSection(ByVal idx As Long)

    On Error GoTo ErrHandler

    txtContent.Text = mSecData(idx)
    txtContent.Visible = True

    Exit Sub

ErrHandler:
    HandleFormError "ShowPlainSection"

End Sub

Private Sub txtContent_Change()

    On Error GoTo ErrHandler

    If mLoading Then Exit Sub
    If mCurrentSection = FindSection(SEC_TITLE) Then
        Me.Caption = IIf(Trim(txtContent.Text) <> "", "SOD Editor: " & Trim(txtContent.Text), "SOD Editor")
    End If

    Exit Sub

ErrHandler:
    HandleFormError "txtContent_Change"

End Sub

Private Sub ShowGroupSection(ByVal idx As Long)

    On Error GoTo ErrHandler

    cboGroup.Visible = True
    cboGroup.Value = mSecData(idx)

    Exit Sub

ErrHandler:
    HandleFormError "ShowGroupSection"

End Sub
'====================================================
' LIST SECTIONS (covers both simple lists and lists
' with sub-items - mSecHasSubItems(idx) controls which)
'====================================================

Private Sub ShowListSection(ByVal idx As Long)

    On Error GoTo ErrHandler

    lstItems.Clear
    lblSubItems.Caption = "Sub-items:"

    Dim i As Long
    For i = 1 To mSecItems(idx).count
        Dim displayVal As String
        If mSecHasSubItems(idx) Then
            displayVal = Split(mSecItems(idx)(i), "|")(0)
        Else
            displayVal = mSecItems(idx)(i)
            If Len(displayVal) > 100 Then displayVal = Left(displayVal, 97) & "..."
        End If
        lstItems.AddItem displayVal
    Next i

    lstItems.Visible = True
    btnAddItem.Visible = True
    btnRemoveItem.Visible = True
    btnEditItem.Visible = True
    txtItem.Visible = True
    txtItem.Text = ""

    ' Checkbox: visible only for user-created sections.
    ' Hidden for built-ins that always have sub-items (Roles/Steps/Resources)
    ' and for built-ins that never have sub-items (Objectives/KPIs).
    Dim isBuiltInWithSubs As Boolean
    Dim isBuiltInWithoutSubs As Boolean
    isBuiltInWithSubs = IsBuiltInSubItemSection(mSecNames(idx))
    isBuiltInWithoutSubs = IsBuiltInListSection(mSecNames(idx))

    If isBuiltInWithSubs Or isBuiltInWithoutSubs Then
        chkSubItems.Visible = False
    Else
        chkSubItems.Visible = True
        mLoading = True
        chkSubItems.Value = mSecHasSubItems(idx)
        mLoading = False
    End If

    Dim showSubs As Boolean
    showSubs = mSecHasSubItems(idx)

    lblSubItems.Visible = showSubs
    lstSubItems.Visible = showSubs
    btnAddSubItem.Visible = showSubs
    btnRemoveSubItem.Visible = showSubs
    btnEditSubItem.Visible = showSubs
    txtSubItem.Visible = showSubs

    If showSubs Then

        lstSubItems.Clear
        txtSubItem.Text = ""

        If lstItems.ListCount > 0 Then
            lstItems.ListIndex = 0
            UpdateSubItemsLabel
            UpdateSubItemControls
            RefreshSubItems
        End If

    Else

        ' Defensive reset: controls are invisible, but keep enabled
        ' state and contents clean rather than leaving stale state
        ' from whatever section was shown previously
        lstSubItems.Clear
        txtSubItem.Enabled = False
        btnAddSubItem.Enabled = False
        btnEditSubItem.Enabled = False
        btnRemoveSubItem.Enabled = False

    End If

    UpdateItemControls

    Exit Sub

ErrHandler:
    HandleFormError "ShowListSection"

End Sub

Private Sub chkSubItems_Click()

    On Error GoTo ErrHandler

    If mLoading Then Exit Sub
    If mSecTypes(mCurrentSection) <> TYPE_LIST Then Exit Sub

    If chkSubItems.Value Then

        ' Switching ON: nothing needs to change about existing data -
        ' plain items just gain the *option* of having subs appended
        mSecHasSubItems(mCurrentSection) = True

    Else

        ' Switching OFF is destructive - warn if any item actually
        ' has sub-item data that would be discarded
        Dim hasAnySubs As Boolean
        hasAnySubs = False

        Dim k As Long
        For k = 1 To mSecItems(mCurrentSection).count
            If InStr(mSecItems(mCurrentSection)(k), "|") > 0 Then
                hasAnySubs = True
                Exit For
            End If
        Next k

        If hasAnySubs Then
            Dim resp As VbMsgBoxResult
            resp = MsgBox("Turning off sub-items will permanently remove all sub-item data for this section." & vbCrLf & vbCrLf & _
                          "Continue?", vbYesNo + vbExclamation, "Remove Sub-items")
            If resp = vbNo Then
                mLoading = True
                chkSubItems.Value = True
                mLoading = False
                Exit Sub
            End If
        End If

        ' Strip everything after the pipe from each item
        Dim newCol As New Collection
        Dim j As Long
        For j = 1 To mSecItems(mCurrentSection).count
            Dim parts() As String
            parts = Split(mSecItems(mCurrentSection)(j), "|")
            newCol.Add parts(0)
        Next j
        Set mSecItems(mCurrentSection) = newCol

        mSecHasSubItems(mCurrentSection) = False

    End If

    ShowListSection mCurrentSection

    Exit Sub

ErrHandler:
    HandleFormError "chkSubItems_Click"

End Sub

Private Sub UpdateSubItemsLabel()

    On Error GoTo ErrHandler

    Dim selIdx As Long
    selIdx = lstItems.ListIndex

    If selIdx < 0 Then
        lblSubItems.Caption = "Sub-items:"
        Exit Sub
    End If

    Dim itemIdx As Long
    itemIdx = selIdx + 1

    If itemIdx > mSecItems(mCurrentSection).count Then
        lblSubItems.Caption = "Sub-items:"
        Exit Sub
    End If

    Dim mainVal As String
    mainVal = Split(mSecItems(mCurrentSection)(itemIdx), "|")(0)

    lblSubItems.Caption = mainVal & ":"

    Exit Sub

ErrHandler:
    HandleFormError "UpdateSubItemsLabel"

End Sub

Private Sub RefreshSubItems()

    On Error GoTo ErrHandler

    lstSubItems.Clear

    Dim selIdx As Long
    selIdx = lstItems.ListIndex
    If selIdx < 0 Then
        UpdateSubItemControls
        Exit Sub
    End If

    Dim itemIdx As Long
    itemIdx = selIdx + 1

    If itemIdx > mSecItems(mCurrentSection).count Then
        UpdateSubItemControls
        Exit Sub
    End If

    Dim entry As String
    entry = mSecItems(mCurrentSection)(itemIdx)

    If InStr(entry, "|") > 0 Then

        Dim parts() As String
        parts = Split(entry, "|")

        If UBound(parts) >= 1 Then
            If Trim(parts(1)) <> "" Then

                Dim subs() As String
                subs = Split(parts(1), "~")

                Dim i As Long
                For i = 0 To UBound(subs)
                    Dim txt As String
                    txt = Trim(subs(i))
                    If txt <> "" Then
                        If Len(txt) > 90 Then txt = Left(txt, 87) & "..."
                        lstSubItems.AddItem txt
                    End If
                Next i

            End If
        End If

    End If

    UpdateSubItemControls

    Exit Sub

ErrHandler:
    HandleFormError "RefreshSubItems"

End Sub

Private Sub RefreshItemsDisplay()

    On Error GoTo ErrHandler

    lstItems.Clear
    Dim i As Long
    For i = 1 To mSecItems(mCurrentSection).count
        If mSecTypes(mCurrentSection) = TYPE_LIST And mSecHasSubItems(mCurrentSection) Then
            lstItems.AddItem Split(mSecItems(mCurrentSection)(i), "|")(0)
        Else
            Dim displayVal As String
            displayVal = mSecItems(mCurrentSection)(i)
            If Len(displayVal) > 100 Then displayVal = Left(displayVal, 97) & "..."
            lstItems.AddItem displayVal
        End If
    Next i

    Exit Sub

ErrHandler:
    HandleFormError "RefreshItemsDisplay"

End Sub

'====================================================
' BUTTON STATE (Item / Sub-item)
'====================================================

Private Sub UpdateItemControls()

    On Error GoTo ErrHandler

    If mEditingItemIdx > 0 Then Exit Sub   ' don't interfere with edit mode

    ' If currently editing a sub-item, lock out ALL item-level controls
    If mEditingSubIdx >= 0 Then
        txtItem.Enabled = False
        txtItem.BackColor = &H8000000F
        btnAddItem.Enabled = False
        btnEditItem.Enabled = False
        btnRemoveItem.Enabled = False
        Exit Sub
    End If

    Dim hasSelection As Boolean
    hasSelection = (lstItems.ListIndex >= 0)

    txtItem.Enabled = True
    txtItem.BackColor = &H80000005
    btnAddItem.Enabled = True
    btnEditItem.Enabled = hasSelection
    btnRemoveItem.Enabled = hasSelection

    Exit Sub

ErrHandler:
    HandleFormError "UpdateItemControls"

End Sub

Private Sub UpdateSubItemControls()

    On Error GoTo ErrHandler

    If mEditingSubIdx >= 0 Then Exit Sub   ' don't interfere with edit mode

    ' If currently editing a parent item, lock out sub-item controls entirely
    If mEditingItemIdx > 0 Then
        txtSubItem.Enabled = False
        btnAddSubItem.Enabled = False
        btnEditSubItem.Enabled = False
        btnRemoveSubItem.Enabled = False
        txtSubItem.BackColor = &H8000000F
        Exit Sub
    End If

    Dim hasParent As Boolean
    Dim hasSubSelection As Boolean

    hasParent = (lstItems.ListIndex >= 0)
    hasSubSelection = (lstSubItems.ListIndex >= 0)

    txtSubItem.Enabled = hasParent
    btnAddSubItem.Enabled = hasParent
    btnEditSubItem.Enabled = hasSubSelection
    btnRemoveSubItem.Enabled = hasSubSelection

    If Not hasParent Then
        txtSubItem.BackColor = &H8000000F
    Else
        txtSubItem.BackColor = &H80000005
    End If

    Exit Sub

ErrHandler:
    HandleFormError "UpdateSubItemControls"

End Sub

'====================================================
' LIST SELECTION EVENTS
'====================================================

Private Sub lstItems_Click()

    On Error GoTo ErrHandler

    If mLoading Then Exit Sub
    If lstItems.ListIndex < 0 Then Exit Sub

    CancelAnyActiveEditMode

    UpdateItemControls

    If mSecTypes(mCurrentSection) = TYPE_LIST And mSecHasSubItems(mCurrentSection) Then
        lstSubItems.Clear
        txtSubItem.Text = ""
        UpdateSubItemsLabel
        UpdateSubItemControls
        RefreshSubItems
    End If

    Exit Sub

ErrHandler:
    HandleFormError "lstItems_Click"

End Sub

Private Sub lstSubItems_Click()

    On Error GoTo ErrHandler

    If mLoading Then Exit Sub

    CancelAnyActiveEditMode

    UpdateSubItemControls

    Exit Sub

ErrHandler:
    HandleFormError "lstSubItems_Click"

End Sub

'====================================================
' ITEM ADD / EDIT / REMOVE / REORDER
'====================================================

Private Sub btnAddItem_Click()

    On Error GoTo ErrHandler

    If mEditingItemIdx > 0 Then

        ' EDIT MODE: this button is "Move Up"
        If mEditingItemIdx <= 1 Then Exit Sub

        SaveEditedItemTextInPlace
        PushUndo

        Set mSecItems(mCurrentSection) = _
            SwapCollectionItems(mSecItems(mCurrentSection), mEditingItemIdx, mEditingItemIdx - 1)

        mEditingItemIdx = mEditingItemIdx - 1

        RefreshItemsDisplay
        lstItems.ListIndex = mEditingItemIdx - 1
        txtItem.SetFocus

    Else

        ' NORMAL MODE: split pasted text on line breaks so a paste of
        ' multiple lines becomes multiple items in one click
        Dim rawLines() As String
        rawLines = Split(txtItem.Text, vbCrLf)

        Dim addedAny As Boolean
        addedAny = False

        Dim lineIdx As Long
        For lineIdx = 0 To UBound(rawLines)

            Dim cleanLine As String
            cleanLine = Trim(rawLines(lineIdx))

            If cleanLine <> "" Then
                mSecItems(mCurrentSection).Add cleanLine
                lstItems.AddItem cleanLine
                addedAny = True
            End If

        Next lineIdx

        If Not addedAny Then
            MsgBox "Please type or paste at least one item before adding.", vbExclamation
            Exit Sub
        End If

        txtItem.Text = ""

        lstItems.ListIndex = lstItems.ListCount - 1
        UpdateItemControls

        If mSecTypes(mCurrentSection) = TYPE_LIST And mSecHasSubItems(mCurrentSection) Then
            UpdateSubItemsLabel
            RefreshSubItems
        End If

    End If

    Exit Sub

ErrHandler:
    HandleFormError "btnAddItem_Click"

End Sub

Private Sub btnRemoveItem_Click()

    On Error GoTo ErrHandler

    If mEditingItemIdx > 0 Then

        ' EDIT MODE: this button is "Move Down"
        If mEditingItemIdx >= mSecItems(mCurrentSection).count Then Exit Sub

        SaveEditedItemTextInPlace
        PushUndo

        Set mSecItems(mCurrentSection) = _
            SwapCollectionItems(mSecItems(mCurrentSection), mEditingItemIdx, mEditingItemIdx + 1)

        mEditingItemIdx = mEditingItemIdx + 1

        RefreshItemsDisplay
        lstItems.ListIndex = mEditingItemIdx - 1
        txtItem.SetFocus

    Else

        Dim selIdx As Long
        selIdx = lstItems.ListIndex

        If selIdx < 0 Then
            MsgBox "Select an item to remove.", vbExclamation
            Exit Sub
        End If

        PushUndo

        Dim i As Long
        Dim newCol As New Collection
        For i = 1 To mSecItems(mCurrentSection).count
            If i <> selIdx + 1 Then
                newCol.Add mSecItems(mCurrentSection)(i)
            End If
        Next i
        Set mSecItems(mCurrentSection) = newCol

        lstItems.RemoveItem selIdx
        lstSubItems.Clear
        UpdateSubItemControls
        UpdateItemControls

    End If

    Exit Sub

ErrHandler:
    HandleFormError "btnRemoveItem_Click"

End Sub

Private Sub btnEditItem_Click()

    On Error GoTo ErrHandler

    If mEditingItemIdx > 0 Then

        ' CONFIRM
        Dim newItem As String
        newItem = Trim(txtItem.Text)

        If newItem = "" Then
            MsgBox "Please type an item before confirming.", vbExclamation
            Exit Sub
        End If

        PushUndo

        Dim newCol As New Collection
        Dim i As Long
        For i = 1 To mSecItems(mCurrentSection).count
            If i = mEditingItemIdx Then
                If mSecTypes(mCurrentSection) = TYPE_LIST And InStr(mSecItems(mCurrentSection)(i), "|") > 0 Then
                    Dim parts() As String
                    parts = Split(mSecItems(mCurrentSection)(i), "|")
                    parts(0) = newItem
                    newCol.Add Join(parts, "|")
                Else
                    newCol.Add newItem
                End If
            Else
                newCol.Add mSecItems(mCurrentSection)(i)
            End If
        Next i
        Set mSecItems(mCurrentSection) = newCol

        Dim savedIdx As Long
        savedIdx = mEditingItemIdx

        ExitItemEditMode

        RefreshItemsDisplay
        lstItems.ListIndex = savedIdx - 1
        If mSecTypes(mCurrentSection) = TYPE_LIST And mSecHasSubItems(mCurrentSection) Then RefreshSubItems

    Else

        Dim selIdx As Long
        selIdx = lstItems.ListIndex

        If selIdx < 0 Then
            MsgBox "Select an item to edit.", vbExclamation
            Exit Sub
        End If

        Dim itemIdx As Long
        itemIdx = selIdx + 1

        Dim current As String
        If mSecTypes(mCurrentSection) = TYPE_LIST And InStr(mSecItems(mCurrentSection)(itemIdx), "|") > 0 Then
            current = Split(mSecItems(mCurrentSection)(itemIdx), "|")(0)
        Else
            current = mSecItems(mCurrentSection)(itemIdx)
        End If

        txtItem.Text = current
        txtItem.SetFocus

        Set mEditSnapshot = CloneCollection(mSecItems(mCurrentSection))
        mEditSnapshotSec = mCurrentSection

        mEditingItemIdx = itemIdx
        UpdateSubItemControls   ' lock out sub-item controls while editing parent

        btnEditItem.Caption = "Confirm"
        btnAddItem.Caption = "Move Up"
        btnRemoveItem.Caption = "Move Down"
        btnCancel.Caption = "Cancel"

    End If

    Exit Sub

ErrHandler:
    HandleFormError "btnEditItem_Click"

End Sub

Private Sub ExitItemEditMode()

    mEditingItemIdx = 0
    btnEditItem.Caption = "Edit"
    btnAddItem.Caption = "+ Add"
    btnRemoveItem.Caption = "- Remove"
    btnCancel.Caption = "Close Editor"
    txtItem.Text = ""
    Set mEditSnapshot = Nothing

    UpdateItemControls
    UpdateSubItemControls   ' restore sub-item controls

End Sub

Private Sub CancelItemEditMode()

    If Not mEditSnapshot Is Nothing Then
        If mEditSnapshotSec = mCurrentSection Then
            Set mSecItems(mCurrentSection) = CloneCollection(mEditSnapshot)
            RefreshItemsDisplay
            If mSecTypes(mCurrentSection) = TYPE_LIST And mSecHasSubItems(mCurrentSection) Then
                lstSubItems.Clear
                UpdateSubItemsLabel
            End If
        End If
    End If

    mEditingItemIdx = 0
    btnEditItem.Caption = "Edit"
    btnAddItem.Caption = "+ Add"
    btnRemoveItem.Caption = "- Remove"
    btnCancel.Caption = "Close Editor"
    txtItem.Text = ""
    Set mEditSnapshot = Nothing

    UpdateItemControls
    UpdateSubItemControls

End Sub

Private Sub SaveEditedItemTextInPlace()

    Dim newItem As String
    newItem = Trim(txtItem.Text)
    If newItem = "" Then Exit Sub

    Dim newCol As New Collection
    Dim i As Long
    For i = 1 To mSecItems(mCurrentSection).count
        If i = mEditingItemIdx Then
            If mSecTypes(mCurrentSection) = TYPE_LIST And InStr(mSecItems(mCurrentSection)(i), "|") > 0 Then
                Dim parts() As String
                parts = Split(mSecItems(mCurrentSection)(i), "|")
                parts(0) = newItem
                newCol.Add Join(parts, "|")
            Else
                newCol.Add newItem
            End If
        Else
            newCol.Add mSecItems(mCurrentSection)(i)
        End If
    Next i
    Set mSecItems(mCurrentSection) = newCol

End Sub

'====================================================
' SUB-ITEM ADD / EDIT / REMOVE / REORDER
'====================================================

Private Sub btnAddSubItem_Click()

    On Error GoTo ErrHandler

    Dim selIdx As Long
    selIdx = lstItems.ListIndex

    If selIdx < 0 Then
        MsgBox "Select a parent item first.", vbExclamation
        Exit Sub
    End If

    Dim itemIdx As Long
    itemIdx = selIdx + 1

    If mEditingSubIdx >= 0 Then

        ' EDIT MODE: this button is "Move Up"
        If mEditingSubIdx <= 0 Then Exit Sub

        SaveEditedSubItemTextInPlace itemIdx
        PushUndo

        Dim parts() As String
        parts = Split(mSecItems(mCurrentSection)(itemIdx), "|")
        Dim subs() As String
        subs = Split(parts(1), "~")

        Dim tmp As String
        tmp = subs(mEditingSubIdx)
        subs(mEditingSubIdx) = subs(mEditingSubIdx - 1)
        subs(mEditingSubIdx - 1) = tmp

        Dim newStr As String
        newStr = parts(0) & "|" & Join(subs, "~")

        Dim newCol As New Collection
        Dim i As Long
        For i = 1 To mSecItems(mCurrentSection).count
            If i = itemIdx Then
                newCol.Add newStr
            Else
                newCol.Add mSecItems(mCurrentSection)(i)
            End If
        Next i
        Set mSecItems(mCurrentSection) = newCol

        mEditingSubIdx = mEditingSubIdx - 1

        RefreshSubItems
        lstSubItems.ListIndex = mEditingSubIdx
        txtSubItem.SetFocus

    Else

        ' NORMAL MODE: split pasted text on line breaks, each line
        ' becomes its own sub-item
        Dim rawLines() As String
        rawLines = Split(txtSubItem.Text, vbCrLf)

        Dim addedAny As Boolean
        addedAny = False

        Dim current As String
        current = mSecItems(mCurrentSection)(itemIdx)

        Dim parts2() As String
        parts2 = Split(current, "|")

        Dim mainPart As String
        mainPart = parts2(0)

        Dim existingSubs As String
        If UBound(parts2) >= 1 Then
            existingSubs = parts2(1)
        Else
            existingSubs = ""
        End If

        Dim lineIdx As Long
        For lineIdx = 0 To UBound(rawLines)

            Dim cleanLine As String
            cleanLine = Trim(rawLines(lineIdx))

            If cleanLine <> "" Then

                If existingSubs = "" Then
                    existingSubs = cleanLine
                Else
                    existingSubs = existingSubs & "~" & cleanLine
                End If

                Dim displayVal As String
                displayVal = cleanLine
                If Len(displayVal) > 90 Then displayVal = Left(displayVal, 87) & "..."
                lstSubItems.AddItem displayVal

                addedAny = True

            End If

        Next lineIdx

        If Not addedAny Then
            MsgBox "Please type or paste at least one sub-item before adding.", vbExclamation
            Exit Sub
        End If

        current = mainPart & "|" & existingSubs

        Dim newCol2 As New Collection
        Dim j As Long
        For j = 1 To mSecItems(mCurrentSection).count
            If j = itemIdx Then
                newCol2.Add current
            Else
                newCol2.Add mSecItems(mCurrentSection)(j)
            End If
        Next j
        Set mSecItems(mCurrentSection) = newCol2

        txtSubItem.Text = ""
        lstSubItems.ListIndex = lstSubItems.ListCount - 1
        UpdateSubItemControls

    End If

    Exit Sub

ErrHandler:
    HandleFormError "btnAddSubItem_Click"

End Sub

Private Sub btnRemoveSubItem_Click()

    On Error GoTo ErrHandler

    Dim parentIdx As Long
    parentIdx = lstItems.ListIndex
    If parentIdx < 0 Then Exit Sub

    Dim itemIdx As Long
    itemIdx = parentIdx + 1

    If mEditingSubIdx >= 0 Then

        ' EDIT MODE: this button is "Move Down"
        Dim parts() As String
        parts = Split(mSecItems(mCurrentSection)(itemIdx), "|")
        Dim subs() As String
        subs = Split(parts(1), "~")

        If mEditingSubIdx >= UBound(subs) Then Exit Sub

        SaveEditedSubItemTextInPlace itemIdx
        PushUndo

        parts = Split(mSecItems(mCurrentSection)(itemIdx), "|")
        subs = Split(parts(1), "~")

        Dim tmp As String
        tmp = subs(mEditingSubIdx)
        subs(mEditingSubIdx) = subs(mEditingSubIdx + 1)
        subs(mEditingSubIdx + 1) = tmp

        Dim newStr As String
        newStr = parts(0) & "|" & Join(subs, "~")

        Dim newCol As New Collection
        Dim i As Long
        For i = 1 To mSecItems(mCurrentSection).count
            If i = itemIdx Then
                newCol.Add newStr
            Else
                newCol.Add mSecItems(mCurrentSection)(i)
            End If
        Next i
        Set mSecItems(mCurrentSection) = newCol

        mEditingSubIdx = mEditingSubIdx + 1

        RefreshSubItems
        lstSubItems.ListIndex = mEditingSubIdx
        txtSubItem.SetFocus

    Else

        Dim subIdx As Long
        subIdx = lstSubItems.ListIndex

        If subIdx < 0 Then
            MsgBox "Select a sub-item to remove.", vbExclamation
            Exit Sub
        End If

        PushUndo

        Dim parts2() As String
        parts2 = Split(mSecItems(mCurrentSection)(itemIdx), "|")

        Dim newStr2 As String
        newStr2 = parts2(0)

        If UBound(parts2) > 0 Then

            Dim subs2() As String
            subs2 = Split(parts2(1), "~")

            Dim newSubs As String
            Dim j As Long
            For j = 0 To UBound(subs2)
                If j <> subIdx Then
                    If newSubs = "" Then
                        newSubs = subs2(j)
                    Else
                        newSubs = newSubs & "~" & subs2(j)
                    End If
                End If
            Next j

            If newSubs <> "" Then newStr2 = newStr2 & "|" & newSubs

        End If

        Dim newCol2 As New Collection
        Dim k As Long
        For k = 1 To mSecItems(mCurrentSection).count
            If k = itemIdx Then
                newCol2.Add newStr2
            Else
                newCol2.Add mSecItems(mCurrentSection)(k)
            End If
        Next k
        Set mSecItems(mCurrentSection) = newCol2

        lstSubItems.RemoveItem subIdx
        UpdateSubItemControls

    End If

    Exit Sub

ErrHandler:
    HandleFormError "btnRemoveSubItem_Click"

End Sub

Private Sub btnEditSubItem_Click()

    On Error GoTo ErrHandler

    If mEditingSubIdx >= 0 Then

        ' CONFIRM
        Dim parentIdx As Long
        parentIdx = lstItems.ListIndex
        If parentIdx < 0 Then Exit Sub

        Dim itemIdx As Long
        itemIdx = parentIdx + 1

        Dim newSub As String
        newSub = Trim(txtSubItem.Text)

        If newSub = "" Then
            MsgBox "Please type a sub-item before confirming.", vbExclamation
            Exit Sub
        End If

        PushUndo

        Dim parts() As String
        parts = Split(mSecItems(mCurrentSection)(itemIdx), "|")

        Dim subs() As String
        If UBound(parts) >= 1 Then
            subs = Split(parts(1), "~")
        Else
            ReDim subs(0)
        End If

        subs(mEditingSubIdx) = newSub

        Dim newStr As String
        newStr = parts(0) & "|" & Join(subs, "~")

        Dim newCol As New Collection
        Dim i As Long
        For i = 1 To mSecItems(mCurrentSection).count
            If i = itemIdx Then
                newCol.Add newStr
            Else
                newCol.Add mSecItems(mCurrentSection)(i)
            End If
        Next i
        Set mSecItems(mCurrentSection) = newCol

        Dim savedSubIdx As Long
        savedSubIdx = mEditingSubIdx

        ExitSubItemEditMode
        UpdateItemControls   ' restore item controls now that sub-item edit is over

        RefreshSubItems
        lstSubItems.ListIndex = savedSubIdx

    Else

        Dim parentIdx2 As Long
        parentIdx2 = lstItems.ListIndex

        If parentIdx2 < 0 Then
            MsgBox "Select a parent item first.", vbExclamation
            Exit Sub
        End If

        Dim subIdx As Long
        subIdx = lstSubItems.ListIndex

        If subIdx < 0 Then
            MsgBox "Select a sub-item to edit.", vbExclamation
            Exit Sub
        End If

        Dim itemIdx2 As Long
        itemIdx2 = parentIdx2 + 1

        Dim parts2() As String
        parts2 = Split(mSecItems(mCurrentSection)(itemIdx2), "|")

        If UBound(parts2) < 1 Then Exit Sub

        Dim subs2() As String
        subs2 = Split(parts2(1), "~")

        If subIdx > UBound(subs2) Then Exit Sub

        txtSubItem.Text = subs2(subIdx)
        txtSubItem.SetFocus

        Set mSubEditSnapshot = CloneCollection(mSecItems(mCurrentSection))
        mSubEditSnapshotSec = mCurrentSection
        mSubEditSnapshotItemIdx = lstItems.ListIndex

        mEditingSubIdx = subIdx
        UpdateItemControls      ' gray out item txt and Add btn immediately
        UpdateSubItemControls

        btnEditSubItem.Caption = "Confirm"
        btnAddSubItem.Caption = "Move Up"
        btnRemoveSubItem.Caption = "Move Down"
        btnCancel.Caption = "Cancel"

    End If

    Exit Sub

ErrHandler:
    HandleFormError "btnEditSubItem_Click"

End Sub

Private Sub ExitSubItemEditMode()

    mEditingSubIdx = -1
    btnEditSubItem.Caption = "Edit"
    btnAddSubItem.Caption = "+ Add"
    btnRemoveSubItem.Caption = "- Remove"
    btnCancel.Caption = "Close Editor"
    txtSubItem.Text = ""
    Set mSubEditSnapshot = Nothing

    UpdateSubItemControls

End Sub

Private Sub CancelSubItemEditMode()

    If Not mSubEditSnapshot Is Nothing Then
        If mSubEditSnapshotSec = mCurrentSection Then
            Set mSecItems(mCurrentSection) = CloneCollection(mSubEditSnapshot)
            lstItems.ListIndex = mSubEditSnapshotItemIdx
            RefreshSubItems
            UpdateSubItemsLabel
        End If
    End If

    mEditingSubIdx = -1
    btnEditSubItem.Caption = "Edit"
    btnAddSubItem.Caption = "+ Add"
    btnRemoveSubItem.Caption = "- Remove"
    btnCancel.Caption = "Close Editor"
    txtSubItem.Text = ""
    Set mSubEditSnapshot = Nothing

    UpdateItemControls
    UpdateSubItemControls

End Sub

Private Sub SaveEditedSubItemTextInPlace(ByVal itemIdx As Long)

    Dim newSub As String
    newSub = Trim(txtSubItem.Text)
    If newSub = "" Then Exit Sub

    Dim parts() As String
    parts = Split(mSecItems(mCurrentSection)(itemIdx), "|")

    Dim subs() As String
    If UBound(parts) >= 1 Then
        subs = Split(parts(1), "~")
    Else
        ReDim subs(0)
    End If

    subs(mEditingSubIdx) = newSub

    Dim newStr As String
    newStr = parts(0) & "|" & Join(subs, "~")

    Dim newCol As New Collection
    Dim i As Long
    For i = 1 To mSecItems(mCurrentSection).count
        If i = itemIdx Then
            newCol.Add newStr
        Else
            newCol.Add mSecItems(mCurrentSection)(i)
        End If
    Next i
    Set mSecItems(mCurrentSection) = newCol

End Sub

'====================================================
' REORDER: ITEMS AND SUB-ITEMS
' (Table row reorder lives in Partition 5)
'====================================================

Private Function SwapCollectionItems(ByVal col As Collection, _
                                      ByVal idx1 As Long, ByVal idx2 As Long) As Collection

    Dim newCol As New Collection
    Dim arr() As Variant
    Dim i As Long

    ReDim arr(1 To col.count)
    For i = 1 To col.count
        arr(i) = col(i)
    Next i

    Dim tmp As Variant
    tmp = arr(idx1)
    arr(idx1) = arr(idx2)
    arr(idx2) = tmp

    For i = 1 To col.count
        newCol.Add arr(i)
    Next i

    Set SwapCollectionItems = newCol

End Function

Private Sub btnItemUp_Click()

    On Error GoTo ErrHandler

    Dim selIdx As Long
    selIdx = lstItems.ListIndex
    If selIdx <= 0 Then Exit Sub

    PushUndo

    Dim itemIdx As Long
    itemIdx = selIdx + 1
    Set mSecItems(mCurrentSection) = SwapCollectionItems(mSecItems(mCurrentSection), itemIdx, itemIdx - 1)

    RefreshItemsDisplay
    lstItems.ListIndex = selIdx - 1
    UpdateSubItemsLabel
    If mSecTypes(mCurrentSection) = TYPE_LIST And mSecHasSubItems(mCurrentSection) Then RefreshSubItems

    Exit Sub

ErrHandler:
    HandleFormError "btnItemUp_Click"

End Sub

Private Sub btnItemDown_Click()

    On Error GoTo ErrHandler

    Dim selIdx As Long
    selIdx = lstItems.ListIndex
    If selIdx < 0 Or selIdx >= lstItems.ListCount - 1 Then Exit Sub

    PushUndo

    Dim itemIdx As Long
    itemIdx = selIdx + 1
    Set mSecItems(mCurrentSection) = SwapCollectionItems(mSecItems(mCurrentSection), itemIdx, itemIdx + 1)

    RefreshItemsDisplay
    lstItems.ListIndex = selIdx + 1
    UpdateSubItemsLabel
    If mSecTypes(mCurrentSection) = TYPE_LIST And mSecHasSubItems(mCurrentSection) Then RefreshSubItems

    Exit Sub

ErrHandler:
    HandleFormError "btnItemDown_Click"

End Sub

Private Sub btnSubItemUp_Click()

    On Error GoTo ErrHandler

    Dim parentIdx As Long
    parentIdx = lstItems.ListIndex
    If parentIdx < 0 Then Exit Sub

    Dim subIdx As Long
    subIdx = lstSubItems.ListIndex
    If subIdx <= 0 Then Exit Sub

    PushUndo

    Dim itemIdx As Long
    itemIdx = parentIdx + 1

    Dim parts() As String
    parts = Split(mSecItems(mCurrentSection)(itemIdx), "|")
    If UBound(parts) < 1 Then Exit Sub

    Dim subs() As String
    subs = Split(parts(1), "~")

    Dim tmp As String
    tmp = subs(subIdx)
    subs(subIdx) = subs(subIdx - 1)
    subs(subIdx - 1) = tmp

    Dim newStr As String
    newStr = parts(0) & "|" & Join(subs, "~")

    Dim newCol As New Collection
    Dim i As Long
    For i = 1 To mSecItems(mCurrentSection).count
        If i = itemIdx Then
            newCol.Add newStr
        Else
            newCol.Add mSecItems(mCurrentSection)(i)
        End If
    Next i
    Set mSecItems(mCurrentSection) = newCol

    RefreshSubItems
    lstSubItems.ListIndex = subIdx - 1

    Exit Sub

ErrHandler:
    HandleFormError "btnSubItemUp_Click"

End Sub

Private Sub btnSubItemDown_Click()

    On Error GoTo ErrHandler

    Dim parentIdx As Long
    parentIdx = lstItems.ListIndex
    If parentIdx < 0 Then Exit Sub

    Dim subIdx As Long
    subIdx = lstSubItems.ListIndex
    If subIdx < 0 Or subIdx >= lstSubItems.ListCount - 1 Then Exit Sub

    PushUndo

    Dim itemIdx As Long
    itemIdx = parentIdx + 1

    Dim parts() As String
    parts = Split(mSecItems(mCurrentSection)(itemIdx), "|")
    If UBound(parts) < 1 Then Exit Sub

    Dim subs() As String
    subs = Split(parts(1), "~")

    Dim tmp As String
    tmp = subs(subIdx)
    subs(subIdx) = subs(subIdx + 1)
    subs(subIdx + 1) = tmp

    Dim newStr As String
    newStr = parts(0) & "|" & Join(subs, "~")

    Dim newCol As New Collection
    Dim i As Long
    For i = 1 To mSecItems(mCurrentSection).count
        If i = itemIdx Then
            newCol.Add newStr
        Else
            newCol.Add mSecItems(mCurrentSection)(i)
        End If
    Next i
    Set mSecItems(mCurrentSection) = newCol

    RefreshSubItems
    lstSubItems.ListIndex = subIdx + 1

    Exit Sub

ErrHandler:
    HandleFormError "btnSubItemDown_Click"

End Sub

'====================================================
' GENERIC TABLE SECTIONS (variable column count)
'====================================================

Private Sub ShowTableSection(ByVal idx As Long)

    On Error GoTo ErrHandler

    lblColCount.Visible = True
    spnColCount.Visible = True
    lblColCountValue.Visible = True
    chkHasHeader.Visible = True

    Dim colCount As Long
    If mSecCols(idx) <> "" Then
        colCount = UBound(Split(mSecCols(idx), "~")) + 1
    Else
        colCount = 1
        mSecCols(idx) = "col1"
    End If

    mLoading = True
    spnColCount.Value = colCount
    lblColCountValue.Caption = CStr(colCount)
    chkHasHeader.Value = mSecHasHeader(idx)
    mLoading = False

    lstRows.Visible = True
    btnAddRow.Visible = True
    btnRemoveRow.Visible = True
    btnEditRow.Visible = True
    lblRowEditor.Visible = True
    fraRowEditor.Visible = True

    RefreshRowList idx
    BuildRowEditorFields

    Exit Sub

ErrHandler:
    mLoading = False
    HandleFormError "ShowTableSection"

End Sub

Private Sub BuildRowEditorFields()

    On Error GoTo ErrHandler

    ClearRowEditor

    If mSecCols(mCurrentSection) = "" Then Exit Sub

    Dim cols() As String
    cols = Split(mSecCols(mCurrentSection), "~")

    Dim i As Long
    Dim topPos As Long
    topPos = 6

    For i = 0 To UBound(cols)

        Dim lbl As MSForms.Label
        Set lbl = fraRowEditor.Controls.Add("Forms.Label.1", "lbl_col_" & i)
        lbl.Tag = "dynamic"
        lbl.Left = 6
        lbl.Top = topPos
        lbl.Width = 70
        lbl.Height = 20

        If mSecHasHeader(mCurrentSection) And mEditingRowIdx = 1 Then
            lbl.Caption = "Header " & (i + 1) & ":"
        ElseIf mSecHasHeader(mCurrentSection) Then
            lbl.Caption = TableColumnLabel(i) & ":"
        Else
            lbl.Caption = "Column " & (i + 1) & ":"
        End If

        Dim txt As MSForms.TextBox
        Set txt = fraRowEditor.Controls.Add("Forms.TextBox.1", "txt_col_" & i)
        txt.Tag = "dynamic"
        txt.Left = 80
        txt.Top = topPos
        txt.Width = 300
        txt.Height = 20

        topPos = topPos + 28

    Next i

    fraRowEditor.ScrollBars = 2
    fraRowEditor.ScrollHeight = topPos + 10
    If fraRowEditor.ScrollHeight < 40 Then fraRowEditor.ScrollHeight = 40

    Exit Sub

ErrHandler:
    HandleFormError "BuildRowEditorFields"

End Sub

Private Sub chkHasHeader_Click()

    On Error GoTo ErrHandler

    If mLoading Then Exit Sub
    If mSecTypes(mCurrentSection) <> TYPE_TABLE Then Exit Sub

    mSecHasHeader(mCurrentSection) = chkHasHeader.Value

    RefreshRowEditorLabels

    Exit Sub

ErrHandler:
    HandleFormError "chkHasHeader_Click"

End Sub

Private Sub RefreshRowEditorLabels()

    On Error GoTo ErrHandler

    If mSecCols(mCurrentSection) = "" Then Exit Sub

    Dim cols() As String
    cols = Split(mSecCols(mCurrentSection), "~")

    Dim editingRow1 As Boolean
    editingRow1 = (mEditingRowIdx = 1)

    Dim i As Long
    For i = 0 To UBound(cols)

        Dim lbl As MSForms.Control
        On Error Resume Next
        Set lbl = fraRowEditor.Controls("lbl_col_" & i)
        On Error GoTo ErrHandler

        If Not lbl Is Nothing Then

            If mSecHasHeader(mCurrentSection) And editingRow1 Then
                lbl.Caption = "Header " & (i + 1) & ":"
            ElseIf mSecHasHeader(mCurrentSection) Then
                lbl.Caption = TableColumnLabel(i) & ":"
            Else
                lbl.Caption = "Column " & (i + 1) & ":"
            End If

        End If

    Next i

    Exit Sub

ErrHandler:
    HandleFormError "RefreshRowEditorLabels"

End Sub

Private Function TableColumnLabel(ByVal colIdx As Long) As String

    On Error GoTo ErrHandler

    If mSecItems(mCurrentSection).count = 0 Then
        TableColumnLabel = "Column " & (colIdx + 1)
        Exit Function
    End If

    Dim headerParts() As String
    headerParts = Split(mSecItems(mCurrentSection)(1), "|")

    If colIdx <= UBound(headerParts) Then
        Dim val As String
        val = Trim(headerParts(colIdx))
        If val <> "" Then
            TableColumnLabel = val
        Else
            TableColumnLabel = "Column " & (colIdx + 1)
        End If
    Else
        TableColumnLabel = "Column " & (colIdx + 1)
    End If

    Exit Function

ErrHandler:
    TableColumnLabel = "Column " & (colIdx + 1)

End Function

Private Sub spnColCount_Change()

    On Error GoTo ErrHandler

    If mLoading Then Exit Sub

    Dim oldCount As Long
    If mSecCols(mCurrentSection) <> "" Then
        oldCount = UBound(Split(mSecCols(mCurrentSection), "~")) + 1
    Else
        oldCount = 0
    End If

    Dim newCount As Long
    newCount = spnColCount.Value

    ' Only warn (and only trim) when actually shrinking, and only if
    ' there's real data that could be affected
    If newCount < oldCount And mSecItems(mCurrentSection).count > 0 Then

        Dim resp As VbMsgBoxResult
        resp = MsgBox("Reducing the column count will permanently remove the data " & _
                      "in the extra column(s) for all existing rows." & vbCrLf & vbCrLf & _
                      "Continue?", vbYesNo + vbExclamation, "Reduce Columns")

        If resp = vbNo Then
            mLoading = True
            spnColCount.Value = oldCount
            mLoading = False
            Exit Sub
        End If

    End If

    PushUndo

    lblColCountValue.Caption = CStr(newCount)

    Dim i As Long
    Dim newCols As String
    For i = 1 To newCount
        If i = 1 Then
            newCols = "col1"
        Else
            newCols = newCols & "~col" & i
        End If
    Next i

    mSecCols(mCurrentSection) = newCols

    If newCount < oldCount Then
        TrimRowsToColumnCount mCurrentSection, newCount
    End If

    RefreshRowList mCurrentSection

    mEditingRowIdx = 0
    btnAddRow.Caption = "+ Add"
    btnRemoveRow.Caption = "- Remove"
    btnEditRow.Caption = "Edit"

    BuildRowEditorFields

    Exit Sub

ErrHandler:
    HandleFormError "spnColCount_Change"

End Sub

Private Sub TrimRowsToColumnCount(ByVal secIdx As Long, ByVal newCount As Long)

    On Error GoTo ErrHandler

    If mSecItems(secIdx).count = 0 Then Exit Sub

    Dim newCol As New Collection
    Dim i As Long

    For i = 1 To mSecItems(secIdx).count

        Dim parts() As String
        parts = Split(mSecItems(secIdx)(i), "|")

        Dim trimmed As String
        Dim c As Long
        For c = 0 To newCount - 1
            If c <= UBound(parts) Then
                If c = 0 Then
                    trimmed = parts(c)
                Else
                    trimmed = trimmed & "|" & parts(c)
                End If
            Else
                If c = 0 Then
                    trimmed = ""
                Else
                    trimmed = trimmed & "|"
                End If
            End If
        Next c

        newCol.Add trimmed

    Next i

    Set mSecItems(secIdx) = newCol

    Exit Sub

ErrHandler:
    HandleFormError "TrimRowsToColumnCount"

End Sub

Private Sub ClearRowEditor()

    On Error GoTo ErrHandler

    Dim ctrl As MSForms.Control
    Dim toDelete() As String
    Dim count As Long
    count = 0

    For Each ctrl In fraRowEditor.Controls
        If ctrl.Tag = "dynamic" Then
            count = count + 1
            ReDim Preserve toDelete(1 To count)
            toDelete(count) = ctrl.name
        End If
    Next ctrl

    Dim j As Long
    For j = 1 To count
        fraRowEditor.Controls.Remove toDelete(j)
    Next j

    fraRowEditor.Tag = ""

    Exit Sub

ErrHandler:
    HandleFormError "ClearRowEditor"

End Sub

Private Function ReadTableRowFields() As String

    On Error GoTo ErrHandler

    Dim cols() As String
    cols = Split(mSecCols(mCurrentSection), "~")

    Dim result As String
    Dim i As Long
    For i = 0 To UBound(cols)

        Dim txt As MSForms.Control
        On Error Resume Next
        Set txt = fraRowEditor.Controls("txt_col_" & i)
        On Error GoTo ErrHandler

        Dim val As String
        If Not txt Is Nothing Then
            val = Trim(txt.Text)
        Else
            val = ""
        End If

        If i = 0 Then
            result = val
        Else
            result = result & "|" & val
        End If

    Next i

    ReadTableRowFields = result
    Exit Function

ErrHandler:
    HandleFormError "ReadTableRowFields"
    ReadTableRowFields = ""

End Function

Private Sub ClearTableRowFields()

    On Error GoTo ErrHandler

    Dim cols() As String
    cols = Split(mSecCols(mCurrentSection), "~")

    Dim i As Long
    For i = 0 To UBound(cols)

        Dim txt As MSForms.Control
        On Error Resume Next
        Set txt = fraRowEditor.Controls("txt_col_" & i)
        On Error GoTo ErrHandler

        If Not txt Is Nothing Then txt.Text = ""

    Next i

    Exit Sub

ErrHandler:
    HandleFormError "ClearTableRowFields"

End Sub

Private Sub LoadTableRowIntoFields(ByVal itemIdx As Long)

    On Error GoTo ErrHandler

    If itemIdx > mSecItems(mCurrentSection).count Then Exit Sub

    Dim parts() As String
    parts = Split(mSecItems(mCurrentSection)(itemIdx), "|")

    Dim cols() As String
    cols = Split(mSecCols(mCurrentSection), "~")

    Dim i As Long
    For i = 0 To UBound(cols)

        Dim txt As MSForms.Control
        On Error Resume Next
        Set txt = fraRowEditor.Controls("txt_col_" & i)
        On Error GoTo ErrHandler

        If Not txt Is Nothing Then
            If i <= UBound(parts) Then
                txt.Text = parts(i)
            Else
                txt.Text = ""
            End If
        End If

    Next i

    Exit Sub

ErrHandler:
    HandleFormError "LoadTableRowIntoFields"

End Sub

Private Sub SaveTableRowFieldsInPlace(ByVal itemIdx As Long)

    On Error GoTo ErrHandler

    Dim newRow As String
    newRow = ReadTableRowFields()

    Dim newCol As New Collection
    Dim i As Long
    For i = 1 To mSecItems(mCurrentSection).count
        If i = itemIdx Then
            newCol.Add newRow
        Else
            newCol.Add mSecItems(mCurrentSection)(i)
        End If
    Next i
    Set mSecItems(mCurrentSection) = newCol

    Exit Sub

ErrHandler:
    HandleFormError "SaveTableRowFieldsInPlace"

End Sub

Private Sub SaveRowEditor()

    On Error GoTo ErrHandler

    If mSecTypes(mCurrentSection) = TYPE_TABLE Then

        If fraRowEditor.Tag = "" Then Exit Sub
        Dim rowIdx As Long
        rowIdx = CLng(fraRowEditor.Tag) + 1
        If rowIdx > mSecItems(mCurrentSection).count Then Exit Sub
        SaveTableRowFieldsInPlace rowIdx

    ElseIf mSecTypes(mCurrentSection) = TYPE_DICTIONARY Then

        If mEditingRowIdx > 0 Then SaveEditedRowTextInPlace

    ElseIf mSecTypes(mCurrentSection) = TYPE_RESOURCES Then

        If mEditingRowIdx > 0 Then SaveEditedRowTextInPlace

    End If

    Exit Sub

ErrHandler:
    HandleFormError "SaveRowEditor"

End Sub

'====================================================
' RESOURCES / DICTIONARY (fixed-field tables)
'====================================================

Private Sub ShowResourcesSection(ByVal idx As Long)

    On Error GoTo ErrHandler

    If mSecCols(idx) = "" Then mSecCols(idx) = "col1~col2~col3"

    lstRows.Visible = True
    btnAddRow.Visible = True
    btnRemoveRow.Visible = True
    btnEditRow.Visible = True

    lblResCol1.Visible = True
    txtResCol1.Visible = True
    lblResCol2.Visible = True
    txtResCol2.Visible = True
    lblResCol3.Visible = True
    txtResCol3.Visible = True

    txtResCol1.Text = ""
    txtResCol2.Text = ""
    txtResCol3.Text = ""

    mEditingRowIdx = 0
    btnAddRow.Caption = "+ Add"
    btnRemoveRow.Caption = "- Remove"

    RefreshRowList idx
    UpdateRowControls

    Exit Sub

ErrHandler:
    HandleFormError "ShowResourcesSection"

End Sub

Private Sub ShowDictionarySection(ByVal idx As Long)

    On Error GoTo ErrHandler

    If mSecCols(idx) = "" Then mSecCols(idx) = "col1~col2"

    lstRows.Visible = True
    btnAddRow.Visible = True
    btnRemoveRow.Visible = True
    btnEditRow.Visible = True

    lblDictCol1.Visible = True
    txtDictCol1.Visible = True
    lblDictCol2.Visible = True
    txtDictCol2.Visible = True

    txtDictCol1.Text = ""
    txtDictCol2.Text = ""

    mEditingRowIdx = 0
    btnAddRow.Caption = "+ Add"
    btnRemoveRow.Caption = "- Remove"

    RefreshRowList idx
    UpdateRowControls

    Exit Sub

ErrHandler:
    HandleFormError "ShowDictionarySection"

End Sub

Private Function ResourceColumnLabel(ByVal colIdx As Long) As String

    Select Case colIdx
        Case 0: ResourceColumnLabel = SEC_RESOURCES_COL1
        Case 1: ResourceColumnLabel = SEC_RESOURCES_COL2
        Case 2: ResourceColumnLabel = SEC_RESOURCES_COL3
        Case Else: ResourceColumnLabel = "Column " & (colIdx + 1)
    End Select

End Function

Private Sub SaveEditedRowTextInPlace()

    On Error GoTo ErrHandler

    Dim newRow As String

    If mSecTypes(mCurrentSection) = TYPE_DICTIONARY Then
        newRow = Trim(txtDictCol1.Text) & "|" & Trim(txtDictCol2.Text)
    Else
        newRow = Trim(txtResCol1.Text) & "|" & Trim(txtResCol2.Text) & "|" & Trim(txtResCol3.Text)
    End If

    Dim newCol As New Collection
    Dim i As Long
    For i = 1 To mSecItems(mCurrentSection).count
        If i = mEditingRowIdx Then
            newCol.Add newRow
        Else
            newCol.Add mSecItems(mCurrentSection)(i)
        End If
    Next i
    Set mSecItems(mCurrentSection) = newCol

    Exit Sub

ErrHandler:
    HandleFormError "SaveEditedRowTextInPlace"

End Sub

Private Sub ExitRowEditMode()

    mEditingRowIdx = 0
    btnAddRow.Caption = "+ Add"
    btnRemoveRow.Caption = "- Remove"
    btnEditRow.Caption = "Edit"
    btnCancel.Caption = "Close Editor"
    txtResCol1.Text = ""
    txtResCol2.Text = ""
    txtResCol3.Text = ""
    txtDictCol1.Text = ""
    txtDictCol2.Text = ""

    UpdateRowControls

End Sub

Private Sub CancelRowEditMode()

    If Not mRowEditSnapshot Is Nothing Then
        If mRowEditSnapshotSec = mCurrentSection Then
            Set mSecItems(mCurrentSection) = CloneCollection(mRowEditSnapshot)
            RefreshRowList mCurrentSection
        End If
    End If

    mEditingRowIdx = 0
    btnEditRow.Caption = "Edit"
    btnAddRow.Caption = "+ Add"
    btnRemoveRow.Caption = "- Remove"
    btnCancel.Caption = "Close Editor"

    If mSecTypes(mCurrentSection) = TYPE_TABLE Then
        ClearTableRowFields
        RefreshRowEditorLabels
    Else
        txtResCol1.Text = ""
        txtResCol2.Text = ""
        txtResCol3.Text = ""
        txtDictCol1.Text = ""
        txtDictCol2.Text = ""
    End If

    Set mRowEditSnapshot = Nothing
    UpdateRowControls

End Sub

'====================================================
' ROW LIST / SELECTION
'====================================================

Private Sub RefreshRowList(ByVal idx As Long)

    On Error GoTo ErrHandler

    lstRows.Clear

    Dim i As Long
    For i = 1 To mSecItems(idx).count

        Dim parts() As String
        parts = Split(mSecItems(idx)(i), "|")

        Dim display As String
        display = ""

        Dim j As Long
        For j = 0 To UBound(parts)
            Dim val As String
            val = Trim(parts(j))
            If val <> "" Then
                If display = "" Then
                    display = val
                Else
                    display = display & "  |  " & val
                End If
            End If
        Next j

        If display = "" Then display = "(empty row " & i & ")"

        lstRows.AddItem display

    Next i

    If mSecTypes(idx) = TYPE_TABLE Then

        chkHasHeader.Enabled = (mSecItems(idx).count >= 1)

        If Not chkHasHeader.Enabled Then
            chkHasHeader.Value = False
            mSecHasHeader(idx) = False
        End If

    End If

    UpdateRowControls

    Exit Sub

ErrHandler:
    HandleFormError "RefreshRowList"

End Sub

Private Sub lstRows_Click()

    On Error GoTo ErrHandler

    If mLoading Then Exit Sub
    If lstRows.ListIndex < 0 Then Exit Sub

    CancelAnyActiveEditMode

    UpdateRowControls

    Exit Sub

ErrHandler:
    HandleFormError "lstRows_Click"

End Sub

Private Sub UpdateRowControls()

    On Error GoTo ErrHandler

    If mEditingRowIdx > 0 Then Exit Sub   ' don't interfere with edit mode

    Dim hasSelection As Boolean
    hasSelection = (lstRows.ListIndex >= 0)

    btnEditRow.Enabled = hasSelection
    btnRemoveRow.Enabled = hasSelection

    Exit Sub

ErrHandler:
    HandleFormError "UpdateRowControls"

End Sub

'====================================================
' ROW ADD / EDIT / REMOVE / REORDER
' (shared by TYPE_TABLE, TYPE_RESOURCES, TYPE_DICTIONARY)
'====================================================

Private Sub btnAddRow_Click()

    On Error GoTo ErrHandler

    Dim isFixed As Boolean
    isFixed = (mSecTypes(mCurrentSection) = TYPE_RESOURCES Or mSecTypes(mCurrentSection) = TYPE_DICTIONARY)

    If isFixed Then

        If mEditingRowIdx > 0 Then

            ' EDIT MODE: this button is "Move Up"
            If mEditingRowIdx <= 1 Then Exit Sub

            SaveEditedRowTextInPlace
            PushUndo

            Set mSecItems(mCurrentSection) = _
                SwapCollectionItems(mSecItems(mCurrentSection), mEditingRowIdx, mEditingRowIdx - 1)

            mEditingRowIdx = mEditingRowIdx - 1

            RefreshRowList mCurrentSection
            lstRows.ListIndex = mEditingRowIdx - 1

        Else

            Dim newRowStr As String

            If mSecTypes(mCurrentSection) = TYPE_DICTIONARY Then

                Dim dCol1 As String, dCol2 As String
                dCol1 = Trim(txtDictCol1.Text)
                dCol2 = Trim(txtDictCol2.Text)

                If dCol1 = "" And dCol2 = "" Then
                    MsgBox "Enter at least one field before adding.", vbExclamation
                    Exit Sub
                End If

                newRowStr = dCol1 & "|" & dCol2
                txtDictCol1.Text = ""
                txtDictCol2.Text = ""

            Else

                Dim col1 As String, col2 As String, col3 As String
                col1 = Trim(txtResCol1.Text)
                col2 = Trim(txtResCol2.Text)
                col3 = Trim(txtResCol3.Text)

                If col1 = "" And col2 = "" And col3 = "" Then
                    MsgBox "Enter at least one field before adding.", vbExclamation
                    Exit Sub
                End If

                newRowStr = col1 & "|" & col2 & "|" & col3
                txtResCol1.Text = ""
                txtResCol2.Text = ""
                txtResCol3.Text = ""

            End If

            mSecItems(mCurrentSection).Add newRowStr

            RefreshRowList mCurrentSection
            lstRows.ListIndex = lstRows.ListCount - 1

        End If

    ElseIf mSecTypes(mCurrentSection) = TYPE_TABLE Then

        If mEditingRowIdx > 0 Then

            ' EDIT MODE: this button is "Move Up"
            If mEditingRowIdx <= 1 Then Exit Sub

            SaveTableRowFieldsInPlace mEditingRowIdx
            PushUndo

            Set mSecItems(mCurrentSection) = _
                SwapCollectionItems(mSecItems(mCurrentSection), mEditingRowIdx, mEditingRowIdx - 1)

            mEditingRowIdx = mEditingRowIdx - 1

            RefreshRowList mCurrentSection
            lstRows.ListIndex = mEditingRowIdx - 1
            LoadTableRowIntoFields mEditingRowIdx

        Else

            If mSecCols(mCurrentSection) = "" Then
                MsgBox "Set a column count before adding rows.", vbExclamation
                Exit Sub
            End If

            Dim newRow As String
            newRow = ReadTableRowFields()

            If Replace(newRow, "|", "") = "" Then
                MsgBox "Enter at least one field before adding.", vbExclamation
                Exit Sub
            End If

            mSecItems(mCurrentSection).Add newRow
            RefreshRowList mCurrentSection

            ClearTableRowFields

        End If

    End If

    Exit Sub

ErrHandler:
    HandleFormError "btnAddRow_Click"

End Sub

Private Sub btnRemoveRow_Click()

    On Error GoTo ErrHandler

    Dim isFixed As Boolean
    Dim selIdx As Long
    Dim newCol As New Collection
    Dim i As Long

    isFixed = (mSecTypes(mCurrentSection) = TYPE_RESOURCES Or mSecTypes(mCurrentSection) = TYPE_DICTIONARY)

    If isFixed Then

        If mEditingRowIdx > 0 Then

            ' EDIT MODE: this button is "Move Down"
            If mEditingRowIdx >= mSecItems(mCurrentSection).count Then Exit Sub

            SaveEditedRowTextInPlace
            PushUndo

            Set mSecItems(mCurrentSection) = _
                SwapCollectionItems(mSecItems(mCurrentSection), mEditingRowIdx, mEditingRowIdx + 1)

            mEditingRowIdx = mEditingRowIdx + 1

            RefreshRowList mCurrentSection
            lstRows.ListIndex = mEditingRowIdx - 1

        Else

            selIdx = lstRows.ListIndex

            If selIdx < 0 Then
                MsgBox "Select a row to remove.", vbExclamation
                Exit Sub
            End If

            PushUndo

            For i = 1 To mSecItems(mCurrentSection).count
                If i <> selIdx + 1 Then newCol.Add mSecItems(mCurrentSection)(i)
            Next i
            Set mSecItems(mCurrentSection) = newCol

            RefreshRowList mCurrentSection

        End If

    ElseIf mSecTypes(mCurrentSection) = TYPE_TABLE Then

        If mEditingRowIdx > 0 Then

            ' EDIT MODE: this button is "Move Down"
            If mEditingRowIdx >= mSecItems(mCurrentSection).count Then Exit Sub

            SaveTableRowFieldsInPlace mEditingRowIdx
            PushUndo

            Set mSecItems(mCurrentSection) = _
                SwapCollectionItems(mSecItems(mCurrentSection), mEditingRowIdx, mEditingRowIdx + 1)

            mEditingRowIdx = mEditingRowIdx + 1

            RefreshRowList mCurrentSection
            lstRows.ListIndex = mEditingRowIdx - 1
            LoadTableRowIntoFields mEditingRowIdx

        Else

            selIdx = lstRows.ListIndex

            If selIdx < 0 Then
                MsgBox "Select a row to remove.", vbExclamation
                Exit Sub
            End If

            PushUndo

            For i = 1 To mSecItems(mCurrentSection).count
                If i <> selIdx + 1 Then newCol.Add mSecItems(mCurrentSection)(i)
            Next i
            Set mSecItems(mCurrentSection) = newCol

            RefreshRowList mCurrentSection

        End If

    End If

    Exit Sub

ErrHandler:
    HandleFormError "btnRemoveRow_Click"

End Sub

Private Sub btnEditRow_Click()

    On Error GoTo ErrHandler

    Dim isFixed As Boolean
    isFixed = (mSecTypes(mCurrentSection) = TYPE_RESOURCES Or mSecTypes(mCurrentSection) = TYPE_DICTIONARY)

    If isFixed Then

        If mEditingRowIdx > 0 Then

            PushUndo
            SaveEditedRowTextInPlace

            Dim savedIdx As Long
            savedIdx = mEditingRowIdx

            ExitRowEditMode

            RefreshRowList mCurrentSection
            lstRows.ListIndex = savedIdx - 1

        Else

            Dim selIdx As Long
            selIdx = lstRows.ListIndex

            If selIdx < 0 Then
                MsgBox "Select a row to edit.", vbExclamation
                Exit Sub
            End If

            Dim rowIdx As Long
            rowIdx = selIdx + 1

            Dim parts() As String
            parts = Split(mSecItems(mCurrentSection)(rowIdx), "|")

            If mSecTypes(mCurrentSection) = TYPE_DICTIONARY Then
                txtDictCol1.Text = IIf(UBound(parts) >= 0, parts(0), "")
                txtDictCol2.Text = IIf(UBound(parts) >= 1, parts(1), "")
                txtDictCol1.SetFocus
            Else
                txtResCol1.Text = IIf(UBound(parts) >= 0, parts(0), "")
                txtResCol2.Text = IIf(UBound(parts) >= 1, parts(1), "")
                txtResCol3.Text = IIf(UBound(parts) >= 2, parts(2), "")
                txtResCol1.SetFocus
            End If

            Set mRowEditSnapshot = CloneCollection(mSecItems(mCurrentSection))
            mRowEditSnapshotSec = mCurrentSection

            mEditingRowIdx = rowIdx
            btnEditRow.Caption = "Confirm"
            btnAddRow.Caption = "Move Up"
            btnRemoveRow.Caption = "Move Down"
            btnCancel.Caption = "Cancel"

        End If

    ElseIf mSecTypes(mCurrentSection) = TYPE_TABLE Then

        If mEditingRowIdx > 0 Then

            ' CONFIRM
            PushUndo
            SaveTableRowFieldsInPlace mEditingRowIdx

            Dim wasEditingRow1 As Boolean
            wasEditingRow1 = (mEditingRowIdx = 1)

            mEditingRowIdx = 0
            btnEditRow.Caption = "Edit"
            btnAddRow.Caption = "+ Add"
            btnRemoveRow.Caption = "- Remove"
            btnCancel.Caption = "Close Editor"

            RefreshRowList mCurrentSection
            ClearTableRowFields

            If wasEditingRow1 Then RefreshRowEditorLabels

        Else

            Dim selIdx2 As Long
            selIdx2 = lstRows.ListIndex

            If selIdx2 < 0 Then
                MsgBox "Select a row to edit.", vbExclamation
                Exit Sub
            End If

            Dim itemIdx As Long
            itemIdx = selIdx2 + 1

            Set mRowEditSnapshot = CloneCollection(mSecItems(mCurrentSection))
            mRowEditSnapshotSec = mCurrentSection

            LoadTableRowIntoFields itemIdx

            mEditingRowIdx = itemIdx
            btnEditRow.Caption = "Confirm"
            btnAddRow.Caption = "Move Up"
            btnRemoveRow.Caption = "Move Down"
            btnCancel.Caption = "Cancel"

            RefreshRowEditorLabels

        End If

    End If

    Exit Sub

ErrHandler:
    HandleFormError "btnEditRow_Click"

End Sub

Private Sub btnRowUp_Click()

    On Error GoTo ErrHandler

    Dim selIdx As Long
    selIdx = lstRows.ListIndex
    If selIdx <= 0 Then Exit Sub

    PushUndo

    Dim rowIdx As Long
    rowIdx = selIdx + 1
    Set mSecItems(mCurrentSection) = SwapCollectionItems(mSecItems(mCurrentSection), rowIdx, rowIdx - 1)

    RefreshRowList mCurrentSection
    lstRows.ListIndex = selIdx - 1

    Exit Sub

ErrHandler:
    HandleFormError "btnRowUp_Click"

End Sub

Private Sub btnRowDown_Click()

    On Error GoTo ErrHandler

    Dim selIdx As Long
    selIdx = lstRows.ListIndex
    If selIdx < 0 Or selIdx >= lstRows.ListCount - 1 Then Exit Sub

    PushUndo

    Dim rowIdx As Long
    rowIdx = selIdx + 1
    Set mSecItems(mCurrentSection) = SwapCollectionItems(mSecItems(mCurrentSection), rowIdx, rowIdx + 1)

    RefreshRowList mCurrentSection
    lstRows.ListIndex = selIdx + 1

    Exit Sub

ErrHandler:
    HandleFormError "btnRowDown_Click"

End Sub

'====================================================
' LOAD FROM SHEET
'
' The big idea: rather than assuming a fixed list of
' sections always exists (Title, Purpose, Scope, ...),
' this reads the sheet's actual columns, left to right,
' and rebuilds mSecNames/mSecTypes/mSecItems from scratch
' to match whatever's really there. This means a
' user-created section (like a custom Table or List) that
' was saved to the sheet will correctly come back as that
' same type when the form is reopened later.
'====================================================

Private Sub LoadFromSheet()

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = ActiveSheet
    If ws.ListObjects.count = 0 Then Exit Sub

    Set tbl = ws.ListObjects(1)
    If tbl.ListRows.count = 0 Then Exit Sub

    ' mSecCount is reset to 0 by the caller (UserForm_Initialize)
    ' before this runs, so AddDynamicSection below builds the
    ' section arrays completely fresh, one section per iteration.
    mSecCount = 0

    Dim col As Long
    col = 1

    ' Walk every column left to right. Each "real" column becomes
    ' one section; a Table*/Bullet* helper column is NOT its own
    ' section - it's consumed as part of whichever section came
    ' before it (that's what the Case blocks' inner loops do,
    ' by advancing "col" past however many helper columns belong
    ' to the section they just loaded).
    Do While col <= tbl.ListColumns.count

        Dim hdr As String
        hdr = Trim(tbl.ListColumns(col).name)

        ' If we somehow land directly on a helper column (e.g. sheet
        ' was hand-edited and a helper got separated from its parent),
        ' just skip it rather than treating it as a bogus new section.
        If IsTableColumn(hdr) Or IsBulletColumn(hdr) Then
            col = col + 1

        Else

            ' Figure out what kind of section this column represents
            ' (plain text, list, table, etc.) by checking its name
            ' against the known built-ins first, then by peeking at
            ' the NEXT column's name for a TableN/BulletN pattern.
            Dim secType As String
            secType = DetectSectionType(hdr, tbl, col)

            ' Create the new section slot (grows all the mSec* arrays
            ' by one and gives them sensible blank defaults).
            AddDynamicSection hdr, secType

            ' AddDynamicSection just incremented mSecCount, so this
            ' new section's index is simply the current mSecCount.
            Dim secIdx As Long
            secIdx = mSecCount

            Select Case secType

                Case TYPE_PLAIN, TYPE_GROUP
                    ' Simplest case: one column, no helpers to skip.
                    LoadPlainColumn tbl, col, secIdx
                    col = col + 1

                Case TYPE_LIST
                    ' A list section MIGHT have a "Bullet1" helper
                    ' column right after it (meaning it has sub-items)
                    ' or it might not (a plain list with no sub-items).
                    ' We peek at the next column to find out which.
                    Dim hasSubs As Boolean
                    hasSubs = False

                    If col < tbl.ListColumns.count Then
                        Dim peekHdr As String
                        peekHdr = LCase(Trim(tbl.ListColumns(col + 1).name))
                        If Left(peekHdr, 6) = "bullet" And InStr(peekHdr, " ") = 0 Then
                            hasSubs = True
                        End If
                    End If

                    mSecHasSubItems(secIdx) = hasSubs

                    If hasSubs Then

                        LoadNestedColumns tbl, col, secIdx

                        ' Skip past every contiguous BulletN column
                        ' that belongs to this same section (Bullet1,
                        ' Bullet2, etc. - as many levels as exist).
                        Do While col + 1 <= tbl.ListColumns.count
                            Dim nextL As String
                            nextL = LCase(Trim(tbl.ListColumns(col + 1).name))
                            If Left(nextL, 6) = "bullet" And InStr(nextL, " ") = 0 Then
                                col = col + 1
                            Else
                                Exit Do
                            End If
                        Loop

                    Else
                        LoadListColumn tbl, col, secIdx
                    End If

                    col = col + 1

                Case TYPE_TABLE
                    ' A generic table has one main column plus however
                    ' many TableN helper columns follow it - could be
                    ' any number, since the user picks the column count.
                    LoadTableColumns tbl, col, secIdx

                    Do While col + 1 <= tbl.ListColumns.count
                        Dim nextTable As String
                        nextTable = LCase(Trim(tbl.ListColumns(col + 1).name))
                        If Left(nextTable, 5) = "table" And InStr(nextTable, " ") = 0 Then
                            col = col + 1
                        Else
                            Exit Do
                        End If
                    Loop

                    col = col + 1

                Case TYPE_DICTIONARY
                    ' Dictionary is ALWAYS exactly 2 columns (Term,
                    ' Definition), so we just manually skip one extra
                    ' column rather than looping - there's nothing to
                    ' "detect", the shape is fixed.
                    LoadTableColumns tbl, col, secIdx
                    If col + 1 <= tbl.ListColumns.count Then col = col + 1
                    col = col + 1

                Case TYPE_RESOURCES
                    ' Same idea as Dictionary, but ALWAYS exactly 3
                    ' columns (Resource Type, Link, Document Number).
                    LoadTableColumns tbl, col, secIdx
                    If col + 1 <= tbl.ListColumns.count Then col = col + 1
                    If col + 1 <= tbl.ListColumns.count Then col = col + 1
                    col = col + 1

                Case Else
                    ' Shouldn't normally happen, but guarantees the
                    ' loop always makes forward progress even if
                    ' DetectSectionType ever returns something unexpected.
                    col = col + 1

            End Select

        End If

    Loop

    Exit Sub

ErrHandler:
    HandleFormError "LoadFromSheet"

End Sub

'====================================================
' FindSection
' Looks up a section's 1-based index by its display name.
' Returns 0 if not found. Used all over the form whenever
' code needs to check "does a Title/Group/etc. section
' currently exist, and if so, which index is it at."
'====================================================

Private Function FindSection(ByVal name As String) As Long

    On Error GoTo ErrHandler

    Dim i As Long
    For i = 1 To mSecCount
        If LCase(Trim(mSecNames(i))) = LCase(Trim(name)) Then
            FindSection = i
            Exit Function
        End If
    Next i

    FindSection = 0
    Exit Function

ErrHandler:
    HandleFormError "FindSection"
    FindSection = 0

End Function

'====================================================
' DetectSectionType
' Decides what TYPE_ a column represents, in two passes:
'   1. Check if the column's NAME exactly matches one of
'      our known built-in section names (Title, Roles,
'      Additional Resources, etc.) - these always get
'      their fixed, known type regardless of what helper
'      columns follow them.
'   2. If it's not a recognized name (i.e. it's a section
'      the USER created), figure out its type by peeking
'      at whether the NEXT column looks like a TableN or
'      BulletN helper column.
' If neither check matches anything, it defaults to
' TYPE_PLAIN (just a single column of text, like a
' paragraph) since that's the safest fallback.
'====================================================

Private Function DetectSectionType(ByVal hdr As String, _
                                    ByVal tbl As ListObject, _
                                    ByVal col As Long) As String

    Select Case LCase(Trim(hdr))
        Case LCase(SEC_TITLE):      DetectSectionType = TYPE_PLAIN:      Exit Function
        Case LCase(SEC_GROUP):      DetectSectionType = TYPE_GROUP:      Exit Function
        Case LCase(SEC_PURPOSE):    DetectSectionType = TYPE_PLAIN:      Exit Function
        Case LCase(SEC_SCOPE):      DetectSectionType = TYPE_PLAIN:      Exit Function
        Case LCase(SEC_DICTIONARY): DetectSectionType = TYPE_DICTIONARY: Exit Function
        Case LCase(SEC_ROLES):      DetectSectionType = TYPE_LIST:       Exit Function
        Case LCase(SEC_OBJECTIVES): DetectSectionType = TYPE_LIST:       Exit Function
        Case LCase(SEC_STEPS):      DetectSectionType = TYPE_LIST:       Exit Function
        Case LCase(SEC_KPIS):       DetectSectionType = TYPE_LIST:       Exit Function
        Case LCase(SEC_RESOURCES):  DetectSectionType = TYPE_RESOURCES:  Exit Function
    End Select

    ' Not a recognized built-in name - this is a user-created section.
    ' First check for a helper column (Table*/Bullet*) next to it.
    If col < tbl.ListColumns.count Then

        Dim nextHdr As String
        nextHdr = LCase(Trim(tbl.ListColumns(col + 1).name))

        If Left(nextHdr, 5) = "table" And InStr(nextHdr, " ") = 0 Then
            DetectSectionType = TYPE_TABLE
            Exit Function
        End If

        If Left(nextHdr, 6) = "bullet" And InStr(nextHdr, " ") = 0 Then
            DetectSectionType = TYPE_LIST
            Exit Function
        End If

    End If

    ' No helper column - fall back to counting how many rows actually
    ' have data. One filled row = a single paragraph (Plain).
    ' More than one filled row = separate list entries (List), since
    ' a real paragraph would normally live in row 1 only.
    Dim filledRows As Long
    filledRows = 0

    Dim r As Long
    For r = 1 To tbl.ListRows.count
        If Trim(tbl.DataBodyRange(r, col).Value) <> "" Then
            filledRows = filledRows + 1
            If filledRows > 1 Then Exit For   ' no need to keep counting
        End If
    Next r

    If filledRows > 1 Then
        DetectSectionType = TYPE_LIST
    Else
        DetectSectionType = TYPE_PLAIN
    End If

End Function

'====================================================
' AddDynamicSection
' Grows every mSec* array by exactly one slot and fills
' that new slot with sensible blank defaults. This is
' called once per section while LoadFromSheet walks the
' sheet's columns.
'
' Note the special-case for mSecCount = 1: VBA's
' "ReDim Preserve" throws an error if you try to use it
' on an array that has never been ReDim'd at all (i.e.
' one that's completely empty/uninitialized). So the very
' first section has to use a plain ReDim instead; every
' section after that uses ReDim Preserve to grow the
' array while keeping what's already in it.
'====================================================

Private Sub AddDynamicSection(ByVal secName As String, ByVal secType As String)

    On Error GoTo ErrHandler

    mSecCount = mSecCount + 1

    If mSecCount = 1 Then
        ReDim mSecNames(1 To 1)
        ReDim mSecTypes(1 To 1)
        ReDim mSecData(1 To 1)
        ReDim mSecItems(1 To 1)
        ReDim mSecCols(1 To 1)
        ReDim mSecHasHeader(1 To 1)
        ReDim mSecHasSubItems(1 To 1)
    Else
        ReDim Preserve mSecNames(1 To mSecCount)
        ReDim Preserve mSecTypes(1 To mSecCount)
        ReDim Preserve mSecData(1 To mSecCount)
        ReDim Preserve mSecItems(1 To mSecCount)
        ReDim Preserve mSecCols(1 To mSecCount)
        ReDim Preserve mSecHasHeader(1 To mSecCount)
        ReDim Preserve mSecHasSubItems(1 To mSecCount)
    End If

    mSecNames(mSecCount) = secName
    mSecTypes(mSecCount) = secType
    mSecData(mSecCount) = ""
    mSecCols(mSecCount) = ""
    mSecHasHeader(mSecCount) = False
    mSecHasSubItems(mSecCount) = False
    Set mSecItems(mSecCount) = New Collection

    Exit Sub

ErrHandler:
    HandleFormError "AddDynamicSection"

End Sub

'====================================================
' LoadPlainColumn
' Reads every non-empty cell in one column and joins
' them together with line breaks into a single block of
' text. Used for Title/Purpose/Scope/Group - anything
' that's just a paragraph of text in the sheet, possibly
' spread across multiple rows.
'====================================================

Private Sub LoadPlainColumn(ByVal tbl As ListObject, _
                             ByVal col As Long, _
                             ByVal secIdx As Long)

    On Error GoTo ErrHandler

    Dim r As Long
    Dim txt As String

    For r = 1 To tbl.ListRows.count
        txt = Trim(tbl.DataBodyRange(r, col).Value)
        If txt <> "" Then
            If mSecData(secIdx) <> "" Then
                mSecData(secIdx) = mSecData(secIdx) & vbCrLf & txt
            Else
                mSecData(secIdx) = txt
            End If
        End If
    Next r

    Exit Sub

ErrHandler:
    HandleFormError "LoadPlainColumn"

End Sub

'====================================================
' LoadListColumn
' Reads every non-empty cell in one column as a SEPARATE
' item (unlike LoadPlainColumn, which merges them into one
' block of text). Used for TYPE_LIST sections that have
' NO sub-items - each row becomes one bullet/list entry.
'====================================================

Private Sub LoadListColumn(ByVal tbl As ListObject, _
                            ByVal col As Long, _
                            ByVal secIdx As Long)

    On Error GoTo ErrHandler

    Dim r As Long
    Dim txt As String

    Set mSecItems(secIdx) = New Collection

    For r = 1 To tbl.ListRows.count
        txt = Trim(tbl.DataBodyRange(r, col).Value)
        If txt <> "" Then mSecItems(secIdx).Add txt
    Next r

    Exit Sub

ErrHandler:
    HandleFormError "LoadListColumn"

End Sub

'====================================================
' LoadNestedColumns
' Reads a TYPE_LIST section that DOES have sub-items -
' i.e. it has a main column (like "Roles") plus one or
' more BulletN helper columns sitting next to it.
'
' The tricky part: on the sheet, a row with something in
' the main column starts a NEW item, and any BulletN value
' on THAT ROW becomes its first sub-item. But if the next
' row has the main column BLANK and only a BulletN value
' filled in, that's still a sub-item of the PREVIOUS item,
' not a new item of its own. So we have to accumulate
' sub-items across consecutive rows until we hit the next
' row that actually has something in the main column.
'
' Each finished item gets "flushed" (written into the
' output Collection) via FlushNestedEntry, either when we
' encounter the next new item, or at the very end of the
' loop for whichever item was accumulating last.
'====================================================

Private Sub LoadNestedColumns(ByVal tbl As ListObject, _
                               ByVal parentCol As Long, _
                               ByVal secIdx As Long)

    On Error GoTo ErrHandler

    ' Find how many contiguous BulletN columns follow the main column.
    Dim lastBulletCol As Long
    lastBulletCol = parentCol

    Dim c As Long
    For c = parentCol + 1 To tbl.ListColumns.count
        Dim colName As String
        colName = LCase(Trim(tbl.ListColumns(c).name))
        If Left(colName, 6) = "bullet" And InStr(colName, " ") = 0 Then
            lastBulletCol = c
        Else
            Exit For
        End If
    Next c

    Set mSecItems(secIdx) = New Collection

    Dim r As Long
    Dim mainVal As String
    Dim currentMain As String     ' the main-column text of the item we're currently building
    Dim currentSubs As String     ' its sub-items so far, joined with "~"
    Dim haveEntry As Boolean      ' True once we've started building an item

    haveEntry = False
    currentMain = ""
    currentSubs = ""

    For r = 1 To tbl.ListRows.count

        mainVal = Trim(tbl.DataBodyRange(r, parentCol).Value)

        If mainVal <> "" Then

            ' This row starts a brand-new item. Before we start
            ' building it, flush whatever item we were previously
            ' accumulating (if any) so its data doesn't get lost.
            If haveEntry Then
                FlushNestedEntry mSecItems(secIdx), currentMain, currentSubs
            End If

            currentMain = mainVal
            currentSubs = ""
            haveEntry = True

        End If

        ' Whether this row started a new item or is just a
        ' continuation row (blank main column), grab whatever
        ' BulletN values are on THIS row and add them as more
        ' sub-items of whichever item we're currently building.
        If haveEntry Then
            For c = parentCol + 1 To lastBulletCol
                Dim subVal As String
                subVal = Trim(tbl.DataBodyRange(r, c).Value)
                If subVal <> "" Then
                    If currentSubs = "" Then
                        currentSubs = subVal
                    Else
                        currentSubs = currentSubs & "~" & subVal
                    End If
                End If
            Next c
        End If

    Next r

    ' The loop ends without ever flushing the LAST item we were
    ' building, since flushing only happens when a NEW item starts.
    ' So we flush it manually here, once, after the loop.
    If haveEntry Then
        FlushNestedEntry mSecItems(secIdx), currentMain, currentSubs
    End If

    Exit Sub

ErrHandler:
    HandleFormError "LoadNestedColumns"

End Sub

'====================================================
' FlushNestedEntry
' Small helper used only by LoadNestedColumns. Combines
' an item's main text and its accumulated sub-items (if
' any) into the "item|sub1~sub2~sub3" format that
' mSecItems expects, and adds it to the output Collection.
'====================================================

Private Sub FlushNestedEntry(ByVal col As Collection, ByVal mainVal As String, ByVal subs As String)

    On Error GoTo ErrHandler

    Dim entry As String
    entry = mainVal
    If subs <> "" Then entry = entry & "|" & subs

    col.Add entry

    Exit Sub

ErrHandler:
    HandleFormError "FlushNestedEntry"

End Sub

'====================================================
' LoadTableColumns
' Shared loader for TYPE_TABLE, TYPE_RESOURCES, and
' TYPE_DICTIONARY - anything that's a multi-column table
' on the sheet. Reads the main column plus however many
' contiguous TableN helper columns follow it, and builds
' one pipe-delimited "col1|col2|col3" string per row.
'
' Special handling for TYPE_TABLE's row 1: for generic
' tables, row 1 is a permanent reserved slot that's either
' the header text (if mSecHasHeader is True) or literally
' blank (if False) - it's ALWAYS kept as a real row, never
' skipped, even when blank. That's different from every
' other row, where a completely blank row just gets
' ignored entirely.
'
' IsHeaderLookalikeRow is used to filter OUT rows that are
' just the literal label text for Resources/Dictionary
' (e.g. "Resource Type | List or Document Link | Document
' Number") - those exist on the sheet purely so the
' generated Word table has a proper header row, but they
' shouldn't show up as an actual editable row in the form.
'====================================================

Private Sub LoadTableColumns(ByVal tbl As ListObject, _
                              ByVal col As Long, _
                              ByVal secIdx As Long)

    On Error GoTo ErrHandler

    Dim lastTableCol As Long
    lastTableCol = col

    Dim c As Long
    For c = col + 1 To tbl.ListColumns.count
        Dim colName As String
        colName = LCase(Trim(tbl.ListColumns(c).name))
        If Left(colName, 5) = "table" And InStr(colName, " ") = 0 Then
            lastTableCol = c
        Else
            Exit For
        End If
    Next c

    Dim colCount As Long
    colCount = lastTableCol - col + 1

    ' mSecCols doesn't store real column names, just a placeholder
    ' string like "col1~col2~col3" - all that matters is HOW MANY
    ' tildes are in it, which tells the rest of the form how many
    ' columns this table has.
    Dim i As Long
    Dim colStr As String
    For i = 1 To colCount
        If i = 1 Then
            colStr = "col1"
        Else
            colStr = colStr & "~col" & i
        End If
    Next i
    mSecCols(secIdx) = colStr

    Set mSecItems(secIdx) = New Collection

    Dim r As Long
    For r = 1 To tbl.ListRows.count

        Dim rowStr As String
        rowStr = ""

        For c = col To lastTableCol
            Dim val As String
            val = Trim(tbl.DataBodyRange(r, c).Value)
            If c = col Then
                rowStr = val
            Else
                rowStr = rowStr & "|" & val
            End If
        Next c

        Dim isBlank As Boolean
        isBlank = (Replace(rowStr, "|", "") = "")

        If mSecTypes(secIdx) = TYPE_TABLE And r = 1 Then

            ' Generic Table's row 1 is always kept no matter what -
            ' it's the reserved header/no-header slot.
            mSecItems(secIdx).Add rowStr

        ElseIf Not isBlank Then

            ' For every other row (and for Resources/Dictionary's
            ' row 1 too), skip it if it's blank, and also skip it
            ' if it's just the literal column-label row that
            ' Resources/Dictionary always write for Word's benefit.
            If Not IsHeaderLookalikeRow(secIdx, rowStr) Then
                mSecItems(secIdx).Add rowStr
            End If

        End If

    Next r

    ' For generic Table sections specifically: figure out whether
    ' the checkbox should start checked or unchecked, based on
    ' whether row 1 (which we just loaded above) actually has any
    ' text in it or is completely blank.
    If mSecTypes(secIdx) = TYPE_TABLE Then
        If mSecItems(secIdx).count >= 1 Then
            Dim firstRowBlank As Boolean
            firstRowBlank = (Replace(mSecItems(secIdx)(1), "|", "") = "")
            mSecHasHeader(secIdx) = Not firstRowBlank
        Else
            mSecHasHeader(secIdx) = False
        End If
    End If

    Exit Sub

ErrHandler:
    HandleFormError "LoadTableColumns"

End Sub

'====================================================
' IsHeaderLookalikeRow
' Returns True if a row's values exactly match the fixed
' label text that Resources/Dictionary always write into
' the sheet as their "header row for Word's benefit" (see
' WriteResourcesToSheet / WriteDictionaryToSheet further
' down). This lets LoadTableColumns filter that row out so
' it doesn't show up as a fake editable row in the form.
'====================================================

Private Function IsHeaderLookalikeRow(ByVal secIdx As Long, ByVal rowStr As String) As Boolean

    Dim parts() As String
    parts = Split(rowStr, "|")

    If mSecTypes(secIdx) = TYPE_RESOURCES Then
        If UBound(parts) < 2 Then Exit Function
        If Trim(parts(0)) = SEC_RESOURCES_COL1 And _
           Trim(parts(1)) = SEC_RESOURCES_COL2 And _
           Trim(parts(2)) = SEC_RESOURCES_COL3 Then
            IsHeaderLookalikeRow = True
        End If
    ElseIf mSecTypes(secIdx) = TYPE_DICTIONARY Then
        If UBound(parts) < 1 Then Exit Function
        If Trim(parts(0)) = SEC_DICT_COL1 And _
           Trim(parts(1)) = SEC_DICT_COL2 Then
            IsHeaderLookalikeRow = True
        End If
    End If

End Function


'====================================================
' SAVE TO SHEET
' The reverse direction: takes everything currently in
' mSecNames/mSecItems/etc. and writes it out as actual
' cells on the worksheet, then wraps the whole thing in
' a proper Excel Table (ListObject) so PopulateSOD and
' a future LoadFromSheet can both read it back correctly.
'====================================================

Private Sub btnSave_Click()

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    If ws.ListObjects.count > 0 Then
        If ws.ListObjects(1).ListRows.count > 0 Then
            Dim resp As VbMsgBoxResult
            resp = MsgBox("Saving will overwrite the previous data on the sheet with the new data from this form." & vbCrLf & vbCrLf & _
                          "Continue?", vbYesNo + vbExclamation, "Overwrite Warning")
            If resp = vbNo Then Exit Sub
        End If
    End If

    ' Make sure whatever's currently being typed/edited on screen
    ' gets committed to mSecItems/mSecData BEFORE we write to the
    ' sheet, otherwise a not-yet-confirmed edit would be lost.
    SaveCurrentSection

    ClearSheet ws
    WriteToSheet ws

    MsgBox "Saved to sheet successfully.", vbInformation, "SOD Form"

    mJustSaved = True

    Exit Sub

ErrHandler:
    HandleFormError "btnSave_Click"

End Sub

Private Sub ClearSheet(ByVal ws As Worksheet)

    On Error GoTo ErrHandler

    ws.Cells.Clear

    Exit Sub

ErrHandler:
    HandleFormError "ClearSheet"

End Sub

'====================================================
' WriteToSheet
' Walks every section in order and writes it to the
' sheet starting at column 1, moving "currentCol" forward
' by however many columns each section actually used
' (1 column for plain text, several for a table, etc.).
' After all sections are written, it finds the true last
' used row/column and wraps that whole range in a
' proper Excel Table so the sheet is immediately usable
' by PopulateSOD or a future reload.
'====================================================

Private Sub WriteToSheet(ByVal ws As Worksheet)

    On Error GoTo ErrHandler

    Dim currentCol As Long
    currentCol = 1

    Dim i As Long
    For i = 1 To mSecCount

        ' Row 1, column currentCol always gets the section's display
        ' name - this becomes the Excel column header.
        ws.Cells(1, currentCol).Value = mSecNames(i)

        Select Case mSecTypes(i)

            Case TYPE_PLAIN, TYPE_GROUP
                WritePlainToSheet ws, i, currentCol
                currentCol = currentCol + 1

            Case TYPE_LIST
                If mSecHasSubItems(i) Then
                    Dim depth As Long
                    depth = MaxSubDepth(i)
                    WriteNestedToSheet ws, i, currentCol, depth
                    currentCol = currentCol + 1 + depth
                Else
                    WriteListToSheet ws, i, currentCol
                    currentCol = currentCol + 1
                End If

            Case TYPE_TABLE
                Dim tblCols As Long
                If mSecCols(i) <> "" Then
                    tblCols = UBound(Split(mSecCols(i), "~")) + 1
                Else
                    tblCols = 1
                End If
                WriteTableToSheet ws, i, currentCol
                currentCol = currentCol + tblCols

            Case TYPE_RESOURCES
                WriteResourcesToSheet ws, i, currentCol
                currentCol = currentCol + 3   ' always exactly 3 columns

            Case TYPE_DICTIONARY
                WriteDictionaryToSheet ws, i, currentCol
                currentCol = currentCol + 2   ' always exactly 2 columns

        End Select

    Next i

    Dim lastCol As Long
    lastCol = currentCol - 1

    ' Find the true last used row by checking EVERY column's own
    ' last-used-row, not just column 1 - a section whose first
    ' column happens to be short (e.g. Additional Resources, where
    ' data starts a couple rows down) would otherwise cause the
    ' Table's range to be too small and cut off real data.
    Dim lastRow As Long
    Dim maxRow As Long
    maxRow = 1

    Dim c As Long
    For c = 1 To lastCol
        Dim colLastRow As Long
        colLastRow = ws.Cells(ws.Rows.count, c).End(-4162).Row   ' -4162 = xlUp
        If colLastRow > maxRow Then maxRow = colLastRow
    Next c

    lastRow = maxRow

    If lastRow >= 2 Then
        Dim tblRange As Range
        Set tblRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        ws.ListObjects.Add(1, tblRange, , 1).name = "SODTable"  ' 1 = xlSrcRange, last arg 1 = xlYes (has headers)
    End If

    Exit Sub

ErrHandler:
    HandleFormError "WriteToSheet"

End Sub

'====================================================
' MaxSubDepth
' For a TYPE_LIST section with sub-items, this checks
' whether ANY item actually has at least one sub-item
' filled in. Returns 1 if so, 0 if every item is empty
' of sub-items. Used by WriteToSheet to decide whether
' a "Bullet1" helper column needs to be written at all.
'====================================================

Private Function MaxSubDepth(ByVal secIdx As Long) As Long

    On Error GoTo ErrHandler

    Dim i As Long
    For i = 1 To mSecItems(secIdx).count

        Dim parts() As String
        parts = Split(mSecItems(secIdx)(i), "|")

        ' Careful: VBA's "And" does NOT short-circuit, so we can't
        ' write "If UBound(parts) > 0 And parts(1) <> ''" on one line -
        ' that would still try to read parts(1) even when UBound(parts)
        ' is 0 (no pipe at all), which throws a Subscript out of range
        ' error. Nesting the Ifs avoids that.
        If UBound(parts) > 0 Then
            If parts(1) <> "" Then
                MaxSubDepth = 1
                Exit Function
            End If
        End If

    Next i

    MaxSubDepth = 0
    Exit Function

ErrHandler:
    HandleFormError "MaxSubDepth"
    MaxSubDepth = 0

End Function

'====================================================
' WritePlainToSheet
' Writes a block of text (possibly containing line breaks)
' back out as one cell per line, starting at row 2 (row 1
' is the column header). Used for Title/Purpose/Scope/Group.
'====================================================

Private Sub WritePlainToSheet(ByVal ws As Worksheet, _
                               ByVal secIdx As Long, _
                               ByVal col As Long)

    On Error GoTo ErrHandler

    Dim lines() As String
    lines = Split(mSecData(secIdx), vbCrLf)

    Dim r As Long
    For r = 0 To UBound(lines)
        If Trim(lines(r)) <> "" Then
            ws.Cells(r + 2, col).Value = Trim(lines(r))
        End If
    Next r

    Exit Sub

ErrHandler:
    HandleFormError "WritePlainToSheet"

End Sub

'====================================================
' WriteListToSheet
' Writes a TYPE_LIST section that has NO sub-items - one
' item per row, no helper column needed.
'====================================================

Private Sub WriteListToSheet(ByVal ws As Worksheet, _
                              ByVal secIdx As Long, _
                              ByVal col As Long)

    On Error GoTo ErrHandler

    Dim i As Long
    For i = 1 To mSecItems(secIdx).count
        ws.Cells(i + 1, col).Value = mSecItems(secIdx)(i)
    Next i

    Exit Sub

ErrHandler:
    HandleFormError "WriteListToSheet"

End Sub

'====================================================
' WriteNestedToSheet
' Writes a TYPE_LIST section that DOES have sub-items.
' The main item goes in its own row/column; each of its
' sub-items goes on its OWN separate row in the "Bullet1"
' helper column, with the main column left blank on those
' continuation rows. This is the exact layout that
' LoadNestedColumns (further up) knows how to read back in.
'====================================================

Private Sub WriteNestedToSheet(ByVal ws As Worksheet, _
                                ByVal secIdx As Long, _
                                ByVal startCol As Long, _
                                ByVal depth As Long)

    On Error GoTo ErrHandler

    ws.Cells(1, startCol + 1).Value = "Bullet1"

    Dim r As Long
    r = 2

    Dim i As Long
    For i = 1 To mSecItems(secIdx).count

        Dim parts() As String
        parts = Split(mSecItems(secIdx)(i), "|")

        ws.Cells(r, startCol).Value = parts(0)

        Dim hasSubs As Boolean
        hasSubs = False

        If UBound(parts) > 0 Then
            If parts(1) <> "" Then hasSubs = True
        End If

        If hasSubs Then

            Dim subs() As String
            subs = Split(parts(1), "~")

            Dim j As Long
            For j = 0 To UBound(subs)
                If subs(j) <> "" Then
                    ' Each sub-item after the first pushes to a NEW row,
                    ' with the main column left blank on that row.
                    ws.Cells(r, startCol + 1).Value = subs(j)
                    r = r + 1
                End If
            Next j

        Else
            r = r + 1
        End If

    Next i

    Exit Sub

ErrHandler:
    HandleFormError "WriteNestedToSheet"

End Sub

'====================================================
' WriteTableToSheet
' Writes a generic TYPE_TABLE section. Row 1 gets written
' unconditionally, whatever it currently contains - if the
' header checkbox was unchecked, mSecItems(1) is simply
' blank text already (that's enforced elsewhere in the
' form), so this doesn't need any special-case logic here;
' it just writes whatever's really in the data.
'====================================================

Private Sub WriteTableToSheet(ByVal ws As Worksheet, _
                               ByVal secIdx As Long, _
                               ByVal startCol As Long)

    On Error GoTo ErrHandler

    If mSecCols(secIdx) = "" Then Exit Sub

    Dim colCount As Long
    colCount = UBound(Split(mSecCols(secIdx), "~")) + 1

    ws.Cells(1, startCol).Value = mSecNames(secIdx)

    Dim c As Long
    For c = 1 To colCount - 1
        ws.Cells(1, startCol + c).Value = "Table" & (c + 1)
    Next c

    Dim r As Long
    r = 2

    Dim i As Long
    For i = 1 To mSecItems(secIdx).count

        Dim parts() As String
        parts = Split(mSecItems(secIdx)(i), "|")

        For c = 0 To colCount - 1
            If c <= UBound(parts) Then
                ws.Cells(r, startCol + c).Value = Trim(parts(c))
            Else
                ws.Cells(r, startCol + c).Value = ""
            End If
        Next c

        r = r + 1

    Next i

    Exit Sub

ErrHandler:
    HandleFormError "WriteTableToSheet"

End Sub

'====================================================
' WriteDictionaryToSheet
' Dictionary is always exactly Term/Definition. Row 2 of
' the sheet ALWAYS gets the literal words "Term" and
' "Definition" written in - not because that's real data,
' but because PopulateSOD builds the Word table's header
' row from whatever's in the sheet's first data row, and
' Dictionary is designed to always show a header in the
' final Word document. (IsHeaderLookalikeRow, further up,
' is what keeps this row from being treated as a fake
' editable row when the sheet is loaded back into the form.)
'====================================================

Private Sub WriteDictionaryToSheet(ByVal ws As Worksheet, _
                                    ByVal secIdx As Long, _
                                    ByVal startCol As Long)

    On Error GoTo ErrHandler

    ws.Cells(1, startCol).Value = mSecNames(secIdx)       ' "Dictionary"
    ws.Cells(1, startCol + 1).Value = "Table2"

    ws.Cells(2, startCol).Value = SEC_DICT_COL1           ' "Term"
    ws.Cells(2, startCol + 1).Value = SEC_DICT_COL2       ' "Definition"

    Dim r As Long
    r = 3   ' actual data starts here, since row 2 is the header label row

    Dim i As Long
    For i = 1 To mSecItems(secIdx).count

        Dim parts() As String
        parts = Split(mSecItems(secIdx)(i), "|")

        Dim c As Long
        For c = 0 To 1
            If c <= UBound(parts) Then
                ws.Cells(r, startCol + c).Value = Trim(parts(c))
            End If
        Next c

        r = r + 1

    Next i

    Exit Sub

ErrHandler:
    HandleFormError "WriteDictionaryToSheet"

End Sub

'====================================================
' WriteResourcesToSheet
' Same idea as WriteDictionaryToSheet, but always exactly
' 3 columns: Resource Type, List or Document Link,
' Document Number.
'====================================================

Private Sub WriteResourcesToSheet(ByVal ws As Worksheet, _
                                    ByVal secIdx As Long, _
                                    ByVal startCol As Long)

    On Error GoTo ErrHandler

    ws.Cells(1, startCol).Value = mSecNames(secIdx)       ' "Additional Resources"
    ws.Cells(1, startCol + 1).Value = "Table2"
    ws.Cells(1, startCol + 2).Value = "Table3"

    ws.Cells(2, startCol).Value = SEC_RESOURCES_COL1
    ws.Cells(2, startCol + 1).Value = SEC_RESOURCES_COL2
    ws.Cells(2, startCol + 2).Value = SEC_RESOURCES_COL3

    Dim r As Long
    r = 3   ' actual data starts here, since row 2 is the header label row

    Dim i As Long
    For i = 1 To mSecItems(secIdx).count

        Dim parts() As String
        parts = Split(mSecItems(secIdx)(i), "|")

        Dim c As Long
        For c = 0 To 2
            If c <= UBound(parts) Then
                ws.Cells(r, startCol + c).Value = Trim(parts(c))
            End If
        Next c

        r = r + 1

    Next i

    Exit Sub

ErrHandler:
    HandleFormError "WriteResourcesToSheet"

End Sub

'====================================================
' CANCEL / CLOSE
' btnCancel doubles as "back out of whatever edit mode
' is active" and, when nothing is active, "close the form."
'====================================================

Private Sub btnCancel_Click()

    On Error GoTo ErrHandler

    If mAddingSection Then
        mAddingSection = False
        txtSectionName.Visible = False
        txtSectionName.Text = ""
        btnAddSection.Caption = "+ Add Section"
        btnEditSection.Caption = "Edit Section"
        btnRemoveSection.Caption = "- Remove Section"
        btnCancel.Caption = "Close Editor"
        Exit Sub
    End If

    If mEditingSectionIdx > 0 Then
        mEditingSectionIdx = 0
        txtSectionName.Visible = False
        txtSectionName.Text = ""
        btnEditSection.Caption = "Edit Section"
        btnAddSection.Caption = "+ Add Section"
        btnRemoveSection.Caption = "- Remove Section"
        btnCancel.Caption = "Close Editor"
        Exit Sub
    End If

    If mEditingItemIdx > 0 Then
        CancelItemEditMode
        RefreshItemsDisplay
        Exit Sub
    End If

    If mEditingSubIdx >= 0 Then
        CancelSubItemEditMode
        Exit Sub
    End If

    If mEditingRowIdx > 0 Then
        CancelRowEditMode
        Exit Sub
    End If

    ' Nothing is mid-edit anywhere - normal close behavior
    If ConfirmDiscard() Then
        Unload Me
    End If

    Exit Sub

ErrHandler:
    HandleFormError "btnCancel_Click"

End Sub

Private Function ConfirmDiscard() As Boolean

    On Error GoTo ErrHandler

    If mJustSaved Then
        ConfirmDiscard = True
        Exit Function
    End If

    Dim resp As VbMsgBoxResult
    resp = MsgBox("Any unsaved changes will be lost." & vbCrLf & vbCrLf & _
                  "Are you sure you want to close without saving?", _
                  vbYesNo + vbExclamation, "Discard Changes?")

    ConfirmDiscard = (resp = vbYes)
    Exit Function

ErrHandler:
    HandleFormError "ConfirmDiscard"
    ConfirmDiscard = False

End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    On Error GoTo ErrHandler

    ' CloseMode 0 = user clicked the form's own X button.
    ' Other modes are triggered by our own code (btnCancel_Click's
    ' Unload Me) which has already confirmed via ConfirmDiscard,
    ' so we only need to intercept the X button specifically.
    If CloseMode = 0 Then
        If Not ConfirmDiscard() Then
            Cancel = True
        End If
    End If

    Exit Sub

ErrHandler:
    HandleFormError "UserForm_QueryClose"

End Sub

Private Sub CancelAnyActiveEditMode()

    On Error GoTo ErrHandler

    If mEditingSubIdx >= 0 Then
        CancelSubItemEditMode
        Exit Sub
    End If

    If mEditingItemIdx > 0 Then
        CancelItemEditMode
        RefreshItemsDisplay
        Exit Sub
    End If

    If mEditingRowIdx > 0 Then
        CancelRowEditMode
        Exit Sub
    End If

    Exit Sub

ErrHandler:
    HandleFormError "CancelAnyActiveEditMode"

End Sub

