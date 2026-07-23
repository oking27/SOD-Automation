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
Private Const SEC_PURPOSE As String = "Purpose"
Private Const SEC_SCOPE As String = "Scope"
Private Const SEC_ROLES As String = "Roles"
Private Const SEC_OBJECTIVES As String = "Objectives"
Private Const SEC_STEPS As String = "Process Steps"
Private Const SEC_KPIS As String = "Key Performance Indicators"
Private Const SEC_RESOURCES As String = "Additional Resources"
Private Const SEC_RESOURCES_COL1 As String = "Resource Type"
Private Const SEC_RESOURCES_COL2 As String = "List or Document Link"
Private Const SEC_RESOURCES_COL3 As String = "Document Number"

' Dictionary
Private Const SEC_DICTIONARY As String = "Dictionary"
Private Const SEC_DICT_COL1 As String = "Term"
Private Const SEC_DICT_COL2 As String = "Definition"
Private Const TYPE_DICTIONARY As String = "DICTIONARY"   ' fixed 2-column table

' Section types
Private Const TYPE_PLAIN As String = "PLAIN"           ' single TextBox
Private Const TYPE_NESTED As String = "NESTED"         ' items + sub-items
Private Const TYPE_LIST As String = "LIST"             ' items only, no sub-items
Private Const TYPE_TABLE As String = "TABLE"           ' column count + rows
Private Const TYPE_RESOURCES As String = "RESOURCES"   ' fixed 3-column table

' Data store: parallel arrays indexed by section
Private mSecNames() As String      ' display name
Private mSecTypes() As String      ' TYPE_PLAIN / TYPE_NESTED / TYPE_LIST / TYPE_TABLE
Private mSecCount As Long

' Each section's data stored as pipe-delimited strings
' PLAIN:  mSecData(i) = full text
' LIST:   mSecItems(i) = Collection of strings
' NESTED: mSecItems(i) = Collection of "item|sub1~sub2~sub3..."
' TABLE:  mSecItems(i) = Collection of "col1val|col2val|col3val"
'         mSecCols(i) stores the column count as a tilde-delimited placeholder string
Private mSecData() As String
Private mSecItems() As Collection
Private mSecCols() As String

Private mCurrentSection As Long      ' 1-based index of selected section
Private mLoading As Boolean          ' suppress change events during load

Private mEditingItemIdx As Long      ' 0 = not editing, otherwise 1-based item index
Private mEditingSubIdx As Long       ' 0 = not editing, otherwise 0-based sub-item index
Private mEditingRowIdx As Long       ' 0 = not editing, otherwise 1-based row index

Private mJustSaved As Boolean

Private mSecHasHeader() As Boolean  ' module-level variable for tables
Private mSecHeaderRow() As String

Private mEditingSectionIdx As Long   ' 0 = not editing/reordering
Private mAddingSection As Boolean

Private Const SEC_GROUP As String = "Group"
Private Const TYPE_GROUP As String = "GROUP"   ' single dropdown, metadata only

'====================================================
' UNDO (single-level, scoped to the section it happened in)
'====================================================

Private mUndoAvailable As Boolean
Private mUndoSecIdx As Long
Private mUndoItems As Collection
Private mUndoData As String
Private mUndoCols As String

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
' REORDERING
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

Private Sub RefreshItemsDisplay()

    On Error GoTo ErrHandler

    lstItems.Clear
    Dim i As Long
    For i = 1 To mSecItems(mCurrentSection).count
        If mSecTypes(mCurrentSection) = TYPE_NESTED Then
            lstItems.AddItem Split(mSecItems(mCurrentSection)(i), "|")(0)
        Else
            lstItems.AddItem mSecItems(mCurrentSection)(i)
        End If
    Next i

    Exit Sub

ErrHandler:
    HandleFormError "RefreshItemsDisplay"

End Sub

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
    If mSecTypes(mCurrentSection) = TYPE_NESTED Then RefreshSubItems

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
    If mSecTypes(mCurrentSection) = TYPE_NESTED Then RefreshSubItems

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

Private Sub btnRowUp_Click()

    On Error GoTo ErrHandler

    Dim selIdx As Long
    selIdx = lstRows.ListIndex
    If selIdx <= 0 Then Exit Sub

    SaveRowEditor
    PushUndo

    Dim rowIdx As Long
    rowIdx = selIdx + RowIndexOffset(mCurrentSection)
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

    SaveRowEditor
    PushUndo

    Dim rowIdx As Long
    rowIdx = selIdx + RowIndexOffset(mCurrentSection)
    Set mSecItems(mCurrentSection) = SwapCollectionItems(mSecItems(mCurrentSection), rowIdx, rowIdx + 1)

    RefreshRowList mCurrentSection
    lstRows.ListIndex = selIdx + 1

    Exit Sub

ErrHandler:
    HandleFormError "btnRowDown_Click"

End Sub

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

Private Sub fraRowEditor_Click()

End Sub

Private Sub lblSectionTitle_Click()

End Sub

Private Sub lstSubItems_Click()

End Sub

Private Sub RefreshCancelButtonCaption()

    On Error Resume Next

    If mAddingSection Or mEditingSectionIdx > 0 Or mEditingItemIdx > 0 _
       Or mEditingSubIdx >= 0 Or mEditingRowIdx > 0 Then
        btnCancel.Caption = "Cancel"
    Else
        btnCancel.Caption = "Close Editor"
    End If

    On Error GoTo 0

End Sub

'---helper to protect built-in sections---
Private Function IsBuiltInSectionName(ByVal name As String) As Boolean

    Select Case LCase(Trim(name))
        Case LCase(SEC_TITLE), LCase(SEC_PURPOSE), LCase(SEC_SCOPE), LCase(SEC_ROLES), _
             LCase(SEC_OBJECTIVES), LCase(SEC_STEPS), LCase(SEC_KPIS), _
             LCase(SEC_RESOURCES), LCase(SEC_DICTIONARY)
            IsBuiltInSectionName = True
    End Select

End Function

Private Sub SwapSections(ByVal idx1 As Long, ByVal idx2 As Long)

    Dim tmpName As String, tmpType As String, tmpData As String, tmpCols As String
    Dim tmpHasHeader As Boolean
    Dim tmpItems As Collection

    tmpName = mSecNames(idx1): mSecNames(idx1) = mSecNames(idx2): mSecNames(idx2) = tmpName
    tmpType = mSecTypes(idx1): mSecTypes(idx1) = mSecTypes(idx2): mSecTypes(idx2) = tmpType
    tmpData = mSecData(idx1): mSecData(idx1) = mSecData(idx2): mSecData(idx2) = tmpData
    tmpCols = mSecCols(idx1): mSecCols(idx1) = mSecCols(idx2): mSecCols(idx2) = tmpCols
    tmpHasHeader = mSecHasHeader(idx1): mSecHasHeader(idx1) = mSecHasHeader(idx2): mSecHasHeader(idx2) = tmpHasHeader

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
        Set mSecItems(i) = mSecItems(i + 1)
    Next i

    mSecCount = mSecCount - 1

    ReDim Preserve mSecNames(1 To mSecCount)
    ReDim Preserve mSecTypes(1 To mSecCount)
    ReDim Preserve mSecData(1 To mSecCount)
    ReDim Preserve mSecCols(1 To mSecCount)
    ReDim Preserve mSecHasHeader(1 To mSecCount)
    ReDim Preserve mSecItems(1 To mSecCount)

End Sub

'====================================================
' FORM INITIALIZE
'====================================================

Private Sub UserForm_Initialize()

    On Error GoTo ErrHandler

    mLoading = True
    mEditingItemIdx = 0
    mEditingSubIdx = -1
    
    mJustSaved = False

    InitSections
    RefreshSectionList
    LoadFromSheet

    mCurrentSection = 1
    lstSections.ListIndex = 0
    ShowSection 1

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
    mSecTypes(4) = TYPE_PLAIN
    mSecTypes(5) = TYPE_DICTIONARY
    mSecTypes(6) = TYPE_NESTED
    mSecTypes(7) = TYPE_LIST
    mSecTypes(8) = TYPE_NESTED
    mSecTypes(9) = TYPE_LIST
    mSecTypes(10) = TYPE_RESOURCES

    Dim i As Long
    For i = 1 To mSecCount
        mSecData(i) = ""
        mSecCols(i) = ""
        mSecHasHeader(i) = False
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

Public Sub OpenSODEditor()
    frmSOD.Show
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

    Exit Sub

ErrHandler:
    HandleFormError "RefreshSectionList"

End Sub

Private Sub lstSections_Click()

    On Error GoTo ErrHandler

    If mLoading Then Exit Sub
    If lstSections.ListIndex < 0 Then Exit Sub

    SaveCurrentSection

    mCurrentSection = lstSections.ListIndex + 1
    ShowSection mCurrentSection

    Exit Sub

ErrHandler:
    HandleFormError "lstSections_Click"

End Sub

'====================================================
' SHOW SECTION - switches the right panel content
'====================================================

Private Sub ShowSection(ByVal idx As Long)

    On Error GoTo ErrHandler

    If idx < 1 Or idx > mSecCount Then Exit Sub

    mLoading = True
    mEditingItemIdx = 0
    mEditingSubIdx = -1
    btnAddItem.Caption = "+ Add"
    btnRemoveItem.Caption = "– Remove"
    btnEditItem.Caption = "Edit"
    btnAddSubItem.Caption = "+ Add"
    btnRemoveSubItem.Caption = "– Remove"
    btnEditSubItem.Caption = "Edit"
    mEditingRowIdx = 0
    btnAddRow.Caption = "+ Add"
    btnRemoveRow.Caption = "– Remove"
    btnEditRow.Caption = "Edit"
    
    mEditingSectionIdx = 0
    btnEditSection.Caption = "Edit"
    btnAddSection.Caption = "+ Add Section"
    btnRemoveSection.Caption = "– Remove"
    txtSectionName.Visible = False
    btnCancel.Caption = "Close Editor"

    lblSectionTitle.Caption = mSecNames(idx)

    HideAllControls

    Select Case mSecTypes(idx)

        Case TYPE_PLAIN
            ShowPlainSection idx

        Case TYPE_LIST
            ShowListSection idx

        Case TYPE_NESTED
            ShowNestedSection idx

        Case TYPE_TABLE
            ShowTableSection idx
            
        Case TYPE_RESOURCES
            ShowResourcesSection idx
            
        Case TYPE_DICTIONARY
            ShowDictionarySection idx
            
        Case TYPE_GROUP
            ShowGroupSection idx
            
    End Select

    mLoading = False

    Exit Sub

ErrHandler:
    mLoading = False
    HandleFormError "ShowSection"

End Sub

Private Sub HideAllControls()

    On Error GoTo ErrHandler

    txtContent.Visible = False
    lstItems.Visible = False
    btnAddItem.Visible = False
    btnRemoveItem.Visible = False
    btnEditItem.Visible = False
    txtItem.Visible = False
    lblSubItems.Visible = False
    lstSubItems.Visible = False
    btnAddSubItem.Visible = False
    btnRemoveSubItem.Visible = False
    btnEditSubItem.Visible = False
    txtSubItem.Visible = False
    cboGroup.Visible = False
    
    chkHasHeader.Visible = False
    lblColCount.Visible = False
    spnColCount.Visible = False
    lstRows.Visible = False
    btnAddRow.Visible = False
    btnRemoveRow.Visible = False
    lblRowEditor.Visible = False
    
    lblResCol1.Visible = False
    txtResCol1.Visible = False
    lblResCol2.Visible = False
    txtResCol2.Visible = False
    lblResCol3.Visible = False
    txtResCol3.Visible = False
    btnEditRow.Visible = False
    
    fraRowEditor.Visible = False
    lblColCountValue.Visible = False
    lblDictCol1.Visible = False
    txtDictCol1.Visible = False
    lblDictCol2.Visible = False
    txtDictCol2.Visible = False

    Exit Sub

ErrHandler:
    HandleFormError "HideAllControls"

End Sub

Private Sub ShowResourcesSection(ByVal idx As Long)

    On Error GoTo ErrHandler

    If mSecCols(idx) = "" Then
        mSecCols(idx) = "col1~col2~col3"
    End If

    lstRows.Visible = True
    btnAddRow.Visible = True
    btnRemoveRow.Visible = True

    lblResCol1.Visible = True
    txtResCol1.Visible = True
    lblResCol2.Visible = True
    txtResCol2.Visible = True
    lblResCol3.Visible = True
    txtResCol3.Visible = True
    btnEditRow.Visible = True

    txtResCol1.Text = ""
    txtResCol2.Text = ""
    txtResCol3.Text = ""

    mEditingRowIdx = 0
    btnAddRow.Caption = "+ Add"
    btnRemoveRow.Caption = "– Remove"

    RefreshRowList idx

    Exit Sub

ErrHandler:
    HandleFormError "ShowResourcesSection"

End Sub

Private Sub ShowPlainSection(ByVal idx As Long)

    On Error GoTo ErrHandler

    txtContent.Text = mSecData(idx)
    txtContent.Visible = True

    Exit Sub

ErrHandler:
    HandleFormError "ShowPlainSection"

End Sub

Private Sub ShowListSection(ByVal idx As Long)

    On Error GoTo ErrHandler

    lstItems.Clear
    Dim i As Long
    For i = 1 To mSecItems(idx).count
        Dim displayVal As String
        displayVal = mSecItems(idx)(i)
        If Len(displayVal) > 90 Then displayVal = Left(displayVal, 87) & "..."
        lstItems.AddItem displayVal
    Next i

    lstItems.Visible = True
    btnAddItem.Visible = True
    btnRemoveItem.Visible = True
    btnEditItem.Visible = True
    txtItem.Visible = True
    txtItem.Text = ""

    Exit Sub

ErrHandler:
    HandleFormError "ShowListSection"

End Sub

Private Sub ShowNestedSection(ByVal idx As Long)

    On Error GoTo ErrHandler

    lstItems.Clear
    Dim i As Long
    For i = 1 To mSecItems(idx).count
        lstItems.AddItem Split(mSecItems(idx)(i), "|")(0)
    Next i

    lstItems.Visible = True
    btnAddItem.Visible = True
    btnRemoveItem.Visible = True
    btnEditItem.Visible = True
    txtItem.Visible = True
    txtItem.Text = ""

    lblSubItems.Visible = True
    lstSubItems.Visible = True
    btnAddSubItem.Visible = True
    btnRemoveSubItem.Visible = True
    btnEditSubItem.Visible = True
    txtSubItem.Visible = True
    txtSubItem.Text = ""
    lstSubItems.Clear

    If lstItems.ListCount > 0 Then
        lstItems.ListIndex = 0
        RefreshSubItems
        UpdateSubItemsLabel
    End If

    Exit Sub

ErrHandler:
    HandleFormError "ShowNestedSection"

End Sub

Private Sub ShowGroupSection(ByVal idx As Long)

    On Error GoTo ErrHandler

    cboGroup.Clear

    ' TODO: replace this placeholder list with your real Group options
    cboGroup.AddItem "Group A"
    cboGroup.AddItem "Group B"
    cboGroup.AddItem "Group C"

    cboGroup.Visible = True
    cboGroup.Value = mSecData(idx)

    Exit Sub

ErrHandler:
    HandleFormError "ShowGroupSection"

End Sub

Private Sub ShowDictionarySection(ByVal idx As Long)

    On Error GoTo ErrHandler

    If mSecCols(idx) = "" Then
        mSecCols(idx) = "col1~col2"
    End If

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
    btnRemoveRow.Caption = "– Remove"

    RefreshRowList idx

    Exit Sub

ErrHandler:
    HandleFormError "ShowDictionarySection"

End Sub

'====================================================
' TABLE SECTIONS
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

    mEditingRowIdx = 0
    btnAddRow.Caption = "+ Add"
    btnRemoveRow.Caption = "– Remove"
    btnEditRow.Caption = "Edit"

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

Private Function RowIndexOffset(ByVal secIdx As Long) As Long
    RowIndexOffset = 1
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
            ' Revert the spinner back to the old count without applying any change
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
    btnRemoveRow.Caption = "– Remove"
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

        If display = "" Then
            display = "(empty row " & i & ")"
        End If

        lstRows.AddItem display

    Next i

    If mSecTypes(idx) = TYPE_TABLE Then

        chkHasHeader.Enabled = (mSecItems(idx).count >= 1)

        If Not chkHasHeader.Enabled Then
            chkHasHeader.Value = False
            mSecHasHeader(idx) = False
        End If

    End If

    Exit Sub

ErrHandler:
    HandleFormError "RefreshRowList"

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

Private Sub btnAddRow_Click()

    On Error GoTo ErrHandler

    Dim isFixed As Boolean
    isFixed = (mSecTypes(mCurrentSection) = TYPE_RESOURCES Or mSecTypes(mCurrentSection) = TYPE_DICTIONARY)

    If isFixed Then

        If mEditingRowIdx > 0 Then

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
            Dim minIdx As Long
            minIdx = RowIndexOffset(mCurrentSection)

            If mEditingRowIdx <= minIdx Then Exit Sub

            SaveTableRowFieldsInPlace mEditingRowIdx
            PushUndo

            Set mSecItems(mCurrentSection) = _
                SwapCollectionItems(mSecItems(mCurrentSection), mEditingRowIdx, mEditingRowIdx - 1)

            mEditingRowIdx = mEditingRowIdx - 1

            RefreshRowList mCurrentSection
            Dim newSelIdx As Long
            newSelIdx = mEditingRowIdx - RowIndexOffset(mCurrentSection)
            lstRows.ListIndex = newSelIdx
            LoadTableRowIntoFields mEditingRowIdx

        Else

            ' NORMAL MODE: read whatever's currently typed in the
            ' (always-visible) blank fields, add it as a new row,
            ' then clear the fields back to blank for the next entry.
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

Private Sub SaveEditedRowTextInPlace()

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
            btnRemoveRow.Caption = "– Remove"
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
            itemIdx = selIdx2 + RowIndexOffset(mCurrentSection)

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

Private Sub btnRemoveRow_Click()

    On Error GoTo ErrHandler

    Dim isFixed As Boolean
    Dim selIdx As Long
    Dim newCol As New Collection
    Dim i As Long
    Dim newSelIdx As Long
    Dim itemIdx As Long

    isFixed = (mSecTypes(mCurrentSection) = TYPE_RESOURCES Or mSecTypes(mCurrentSection) = TYPE_DICTIONARY)

    If isFixed Then

        If mEditingRowIdx > 0 Then

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
                If i <> selIdx + 1 Then
                    newCol.Add mSecItems(mCurrentSection)(i)
                End If
            Next i
            Set mSecItems(mCurrentSection) = newCol

            RefreshRowList mCurrentSection

        End If

    ElseIf mSecTypes(mCurrentSection) = TYPE_TABLE Then

        If mEditingRowIdx > 0 Then

            If mEditingRowIdx >= mSecItems(mCurrentSection).count Then Exit Sub

            SaveTableRowFieldsInPlace mEditingRowIdx
            PushUndo

            Set mSecItems(mCurrentSection) = _
                SwapCollectionItems(mSecItems(mCurrentSection), mEditingRowIdx, mEditingRowIdx + 1)

            mEditingRowIdx = mEditingRowIdx + 1

            RefreshRowList mCurrentSection
            newSelIdx = mEditingRowIdx - RowIndexOffset(mCurrentSection)
            lstRows.ListIndex = newSelIdx
            LoadTableRowIntoFields mEditingRowIdx

        Else

            selIdx = lstRows.ListIndex

            If selIdx < 0 Then
                MsgBox "Select a row to remove.", vbExclamation
                Exit Sub
            End If

            PushUndo

            itemIdx = selIdx + RowIndexOffset(mCurrentSection)

            For i = 1 To mSecItems(mCurrentSection).count
                If i <> itemIdx Then
                    newCol.Add mSecItems(mCurrentSection)(i)
                End If
            Next i
            Set mSecItems(mCurrentSection) = newCol

            RefreshRowList mCurrentSection

        End If

    End If

    Exit Sub

ErrHandler:
    HandleFormError "btnRemoveRow_Click"

End Sub

Private Sub lstRows_Click()

    On Error GoTo ErrHandler

    If mLoading Then Exit Sub
    If lstRows.ListIndex < 0 Then Exit Sub

    ' Every section type now uses click-to-select, explicit Edit to open -
    ' consistent with Resources/Dictionary. Nothing auto-saves or
    ' auto-opens on a mere click.

    Exit Sub

ErrHandler:
    HandleFormError "lstRows_Click"

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

Private Function ResourceColumnLabel(ByVal colIdx As Long) As String

    Select Case colIdx
        Case 0: ResourceColumnLabel = SEC_RESOURCES_COL1
        Case 1: ResourceColumnLabel = SEC_RESOURCES_COL2
        Case 2: ResourceColumnLabel = SEC_RESOURCES_COL3
        Case Else: ResourceColumnLabel = "Column " & (colIdx + 1)
    End Select

End Function

Private Sub SaveRowEditor()

    On Error GoTo ErrHandler

    If fraRowEditor.Tag = "" Then Exit Sub

Dim rowIdx As Long
    rowIdx = CLng(fraRowEditor.Tag)

    Dim itemIdx As Long
    itemIdx = rowIdx + RowIndexOffset(mCurrentSection)

    If itemIdx > mSecItems(mCurrentSection).count Then Exit Sub
    If mSecCols(mCurrentSection) = "" Then Exit Sub
    If mSecCols(mCurrentSection) = "" Then Exit Sub

    Dim cols() As String
    cols = Split(mSecCols(mCurrentSection), "~")

    Dim newRow As String
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
            newRow = val
        Else
            newRow = newRow & "|" & val
        End If

    Next i

    Dim newCol As New Collection
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
    HandleFormError "SaveRowEditor"

End Sub

'====================================================
' EDIT MAIN ITEM / SUB-ITEM (List / Nested sections)
'====================================================

Private Sub btnEditItem_Click()

    On Error GoTo ErrHandler

    If mEditingItemIdx > 0 Then
    
        btnCancel.Caption = "Close Editor"

        ' Already editing - this click means CONFIRM
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
                If mSecTypes(mCurrentSection) = TYPE_NESTED Then
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
        If mSecTypes(mCurrentSection) = TYPE_NESTED Then RefreshSubItems

    Else

        ' Not editing - ENTER edit mode
        Dim selIdx As Long
        selIdx = lstItems.ListIndex

        If selIdx < 0 Then
            MsgBox "Select an item to edit.", vbExclamation
            Exit Sub
        End If

        Dim itemIdx As Long
        itemIdx = selIdx + 1

        Dim current As String
        If mSecTypes(mCurrentSection) = TYPE_NESTED Then
            current = Split(mSecItems(mCurrentSection)(itemIdx), "|")(0)
        Else
            current = mSecItems(mCurrentSection)(itemIdx)
        End If

        txtItem.Text = current
        txtItem.SetFocus

        mEditingItemIdx = itemIdx
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
    btnRemoveItem.Caption = "– Remove"
    btnCancel.Caption = "Close Editor"
    txtItem.Text = ""

End Sub

Private Sub SaveEditedItemTextInPlace()

    Dim newItem As String
    newItem = Trim(txtItem.Text)
    If newItem = "" Then Exit Sub

    Dim newCol As New Collection
    Dim i As Long
    For i = 1 To mSecItems(mCurrentSection).count
        If i = mEditingItemIdx Then
            If mSecTypes(mCurrentSection) = TYPE_NESTED Then
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

Private Sub btnEditSubItem_Click()

    On Error GoTo ErrHandler

    If mEditingSubIdx >= 0 Then

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

        mEditingSubIdx = subIdx
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
    btnRemoveSubItem.Caption = "– Remove"
    btnCancel.Caption = "Close Editor"
    txtSubItem.Text = ""

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

        Case TYPE_TABLE
            SaveRowEditor

        Case TYPE_RESOURCES
            SaveRowEditor
            
        Case TYPE_DICTIONARY
            SaveRowEditor

    End Select

    Exit Sub

ErrHandler:
    HandleFormError "SaveCurrentSection"

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

'====================================================
' SUB-ITEMS (for NESTED sections)
'====================================================

Private Sub lstItems_Click()

    On Error GoTo ErrHandler

    If mLoading Then Exit Sub
    If mSecTypes(mCurrentSection) <> TYPE_NESTED Then Exit Sub
    If lstItems.ListIndex < 0 Then Exit Sub

    lstSubItems.Clear
    txtSubItem.Text = ""
    UpdateSubItemsLabel
    RefreshSubItems

    Exit Sub

ErrHandler:
    HandleFormError "lstItems_Click"

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
    If selIdx < 0 Then Exit Sub

    Dim itemIdx As Long
    itemIdx = selIdx + 1

    If itemIdx > mSecItems(mCurrentSection).count Then Exit Sub

    Dim entry As String
    entry = mSecItems(mCurrentSection)(itemIdx)

    If InStr(entry, "|") = 0 Then Exit Sub

    Dim parts() As String
    parts = Split(entry, "|")

    If UBound(parts) < 1 Then Exit Sub
    If Trim(parts(1)) = "" Then Exit Sub

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

    Exit Sub

ErrHandler:
    HandleFormError "RefreshSubItems"

End Sub

'====================================================
' ADD / REMOVE MAIN ITEMS
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

        ' NORMAL MODE: standard Add - split on line breaks so a paste
        ' of multiple lines becomes multiple items in one click
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

        ' Select the last item added (handles both single-add and
        ' multi-line paste-splitting) so sub-items can be added immediately
        lstItems.ListIndex = lstItems.ListCount - 1

        If mSecTypes(mCurrentSection) = TYPE_NESTED Then
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

        ' NORMAL MODE: standard Remove
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

    End If

    Exit Sub

ErrHandler:
    HandleFormError "btnRemoveItem_Click"

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

'====================================================
' ADD / REMOVE SUB-ITEMS
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

        ' ADD mode - split pasted text on line breaks, each line becomes
        ' its own sub-item
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
        ' Select the last sub-item added
        lstSubItems.ListIndex = lstSubItems.ListCount - 1

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

    End If

    Exit Sub

ErrHandler:
    HandleFormError "btnRemoveSubItem_Click"

End Sub

'====================================================
' ADD NEW SECTIONS
'====================================================
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

    mSecNames(mSecCount) = secName
    mSecTypes(mSecCount) = secType
    mSecData(mSecCount) = ""
    mSecCols(mSecCount) = ""
    mSecHasHeader(mSecCount) = False
    Set mSecItems(mSecCount) = New Collection

    RefreshSectionList

    SaveCurrentSection

    mAddingSection = False
    txtSectionName.Visible = False
    btnAddSection.Caption = "+ Add Section"
    btnEditSection.Caption = "Edit"
    btnRemoveSection.Caption = "– Remove"
    btnCancel.Caption = "Close Editor"

    lstSections.ListIndex = mSecCount - 1
    mCurrentSection = mSecCount
    ShowSection mCurrentSection

    Exit Sub

ErrHandler:
    HandleFormError "CreateNewSection"

End Sub

Private Sub btnAddSection_Click()

    On Error GoTo ErrHandler

    If mAddingSection Then
        CreateNewSection TYPE_PLAIN
        Exit Sub
    End If

    If mEditingSectionIdx > 0 Then

        If mEditingSectionIdx <= 1 Then Exit Sub

        PushUndo
        SwapSections mEditingSectionIdx, mEditingSectionIdx - 1
        mEditingSectionIdx = mEditingSectionIdx - 1

        RefreshSectionList
        lstSections.ListIndex = mEditingSectionIdx - 1

        Exit Sub

    End If

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
        btnEditSection.Caption = "Edit"
        btnAddSection.Caption = "+ Add Section"
        btnRemoveSection.Caption = "– Remove"
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

        ' Currently in rename/reorder mode: this button is "Move Down"
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
    resp = MsgBox("Remove section '" & mSecNames(secIdx) & "'? This deletes all of its data permanently.", _
                  vbYesNo + vbExclamation, "Remove Section")
    If resp = vbNo Then Exit Sub

    PushUndo
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

'====================================================
' LOAD FROM SHEET
'====================================================

Private Sub LoadFromSheet()

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = ActiveSheet
    If ws.ListObjects.count = 0 Then Exit Sub

    Set tbl = ws.ListObjects(1)
    If tbl.ListRows.count = 0 Then Exit Sub

    Dim col As Long
    Dim hdr As String

    For col = 1 To tbl.ListColumns.count

        hdr = Trim(tbl.ListColumns(col).name)

        Dim secIdx As Long
        secIdx = FindSection(hdr)

        If secIdx > 0 Then

            Select Case mSecTypes(secIdx)

                Case TYPE_PLAIN
                    LoadPlainColumn tbl, col, secIdx

                Case TYPE_LIST
                    LoadListColumn tbl, col, secIdx

                Case TYPE_NESTED
                    LoadNestedColumns tbl, col, secIdx

                Case TYPE_TABLE
                    LoadTableColumns tbl, col, secIdx
                    
                Case TYPE_RESOURCES
                    LoadTableColumns tbl, col, secIdx
                    
                Case TYPE_DICTIONARY
                    LoadTableColumns tbl, col, secIdx
                    
                Case TYPE_GROUP
                    LoadPlainColumn tbl, col, secIdx

            End Select

        End If

    Next col

    Exit Sub

ErrHandler:
    HandleFormError "LoadFromSheet"

End Sub

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

Private Sub LoadNestedColumns(ByVal tbl As ListObject, _
                               ByVal parentCol As Long, _
                               ByVal secIdx As Long)

    On Error GoTo ErrHandler

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
    Dim currentMain As String
    Dim currentSubs As String
    Dim haveEntry As Boolean

    haveEntry = False
    currentMain = ""
    currentSubs = ""

    For r = 1 To tbl.ListRows.count

        mainVal = Trim(tbl.DataBodyRange(r, parentCol).Value)

        If mainVal <> "" Then

            ' A new role/parent starts here - flush whatever we were
            ' accumulating for the previous one first
            If haveEntry Then
                FlushNestedEntry mSecItems(secIdx), currentMain, currentSubs
            End If

            currentMain = mainVal
            currentSubs = ""
            haveEntry = True

        End If

        ' Collect this row's helper-column value(s) into the CURRENT entry -
        ' whether this row started a new item or is a continuation row
        ' where the parent cell was left blank
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

    If haveEntry Then
        FlushNestedEntry mSecItems(secIdx), currentMain, currentSubs
    End If

    Exit Sub

ErrHandler:
    HandleFormError "LoadNestedColumns"

End Sub

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

            ' Row 1 for Table sections is always kept, blank or not -
            ' it's the reserved header/no-header slot.
            mSecItems(secIdx).Add rowStr

        ElseIf Not isBlank Then

            If Not IsHeaderLookalikeRow(secIdx, rowStr) Then
                mSecItems(secIdx).Add rowStr
            End If

        End If

    Next r

    ' Auto-check the header box if the sheet's first row had content
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
'====================================================

Private Sub btnSave_Click()

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    If ws.ListObjects.count > 0 Then
        If ws.ListObjects(1).ListRows.count > 0 Then
            Dim resp As VbMsgBoxResult
            resp = MsgBox("This sheet already has data. Saving will overwrite it." & vbCrLf & vbCrLf & _
                          "Continue?", vbYesNo + vbExclamation, "Overwrite Warning")
            If resp = vbNo Then Exit Sub
        End If
    End If

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

Private Sub WriteToSheet(ByVal ws As Worksheet)

    On Error GoTo ErrHandler

    Dim currentCol As Long
    currentCol = 1

    Dim i As Long
    For i = 1 To mSecCount

        ws.Cells(1, currentCol).Value = mSecNames(i)

        Select Case mSecTypes(i)

            Case TYPE_PLAIN
                WritePlainToSheet ws, i, currentCol
                currentCol = currentCol + 1

            Case TYPE_LIST
                WriteListToSheet ws, i, currentCol
                currentCol = currentCol + 1

            Case TYPE_NESTED
                Dim depth As Long
                depth = MaxSubDepth(i)
                WriteNestedToSheet ws, i, currentCol, depth
                currentCol = currentCol + 1 + depth

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
                currentCol = currentCol + 3
                
            Case TYPE_DICTIONARY
                WriteDictionaryToSheet ws, i, currentCol
                currentCol = currentCol + 2
                
            Case TYPE_GROUP
                WritePlainToSheet ws, i, currentCol
                currentCol = currentCol + 1

        End Select

    Next i

    Dim lastCol As Long
    lastCol = currentCol - 1

    ' Find the true last used row by checking EVERY column, not just column 1
    Dim lastRow As Long
    Dim maxRow As Long
    maxRow = 1

    Dim c As Long
    For c = 1 To lastCol
        Dim colLastRow As Long
        colLastRow = ws.Cells(ws.Rows.count, c).End(-4162).Row   ' xlUp
        If colLastRow > maxRow Then maxRow = colLastRow
    Next c

    lastRow = maxRow

    If lastRow >= 2 Then
        Dim tblRange As Range
        Set tblRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        ws.ListObjects.Add(1, tblRange, , 1).name = "SODTable"  ' xlSrcRange, xlYes
    End If

    Exit Sub

ErrHandler:
    HandleFormError "WriteToSheet"

End Sub

Private Function MaxSubDepth(ByVal secIdx As Long) As Long

    On Error GoTo ErrHandler

    Dim i As Long
    For i = 1 To mSecItems(secIdx).count

        Dim parts() As String
        parts = Split(mSecItems(secIdx)(i), "|")

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

    If mSecHasHeader(secIdx) And mSecHeaderRow(secIdx) <> "" Then

        Dim headerParts() As String
        headerParts = Split(mSecHeaderRow(secIdx), "|")

        For c = 0 To colCount - 1
            If c <= UBound(headerParts) Then
                ws.Cells(r, startCol + c).Value = Trim(headerParts(c))
            End If
        Next c

        r = r + 1

    End If

    Dim i As Long
    For i = 1 To mSecItems(secIdx).count

        Dim parts() As String
        parts = Split(mSecItems(secIdx)(i), "|")

        For c = 0 To colCount - 1
            If c <= UBound(parts) Then
                ws.Cells(r, startCol + c).Value = Trim(parts(c))
            End If
        Next c

        r = r + 1

    Next i

    Exit Sub

ErrHandler:
    HandleFormError "WriteTableToSheet"

End Sub

Private Sub WriteDictionaryToSheet(ByVal ws As Worksheet, _
                                    ByVal secIdx As Long, _
                                    ByVal startCol As Long)

    On Error GoTo ErrHandler

    ws.Cells(1, startCol).Value = mSecNames(secIdx)     ' "Dictionary"
    ws.Cells(1, startCol + 1).Value = "Table2"

    ws.Cells(2, startCol).Value = SEC_DICT_COL1
    ws.Cells(2, startCol + 1).Value = SEC_DICT_COL2

    Dim r As Long
    r = 3

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

Private Sub WriteResourcesToSheet(ByVal ws As Worksheet, _
                                    ByVal secIdx As Long, _
                                    ByVal startCol As Long)

    On Error GoTo ErrHandler

    ws.Cells(1, startCol).Value = mSecNames(secIdx)     ' "Additional Resources"
    ws.Cells(1, startCol + 1).Value = "Table2"
    ws.Cells(1, startCol + 2).Value = "Table3"

    ' Row 2 = the Word table's header row - written as literal data,
    ' since PopulateSOD builds the Word table from data rows, not
    ' from these Excel column names.
    ws.Cells(2, startCol).Value = SEC_RESOURCES_COL1
    ws.Cells(2, startCol + 1).Value = SEC_RESOURCES_COL2
    ws.Cells(2, startCol + 2).Value = SEC_RESOURCES_COL3

    Dim r As Long
    r = 3   ' actual data starts one row lower now

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
' CANCEL/CLOSE
'====================================================

Private Sub btnCancel_Click()

    On Error GoTo ErrHandler

    ' If any edit/add mode is active anywhere on the form, this button
    ' backs out of THAT instead of closing the whole form.
    If mAddingSection Then

        mAddingSection = False
        txtSectionName.Visible = False
        txtSectionName.Text = ""
        btnAddSection.Caption = "+ Add Section"
        btnEditSection.Caption = "Edit"
        btnRemoveSection.Caption = "– Remove"
        btnCancel.Caption = "Close Editor"
        Exit Sub

    End If

    If mEditingSectionIdx > 0 Then

        mEditingSectionIdx = 0
        txtSectionName.Visible = False
        txtSectionName.Text = ""
        btnEditSection.Caption = "Edit"
        btnAddSection.Caption = "+ Add Section"
        btnRemoveSection.Caption = "– Remove"
        btnCancel.Caption = "Close Editor"
        Exit Sub

    End If

    If mEditingItemIdx > 0 Then
        ExitItemEditMode
        RefreshItemsDisplay
        Exit Sub
    End If

    If mEditingSubIdx >= 0 Then
        ExitSubItemEditMode
        RefreshSubItems
        Exit Sub
    End If

    If mEditingRowIdx > 0 Then

        mEditingRowIdx = 0
        btnEditRow.Caption = "Edit"
        btnAddRow.Caption = "+ Add"
        btnRemoveRow.Caption = "– Remove"
        btnCancel.Caption = "Close Editor"

        If mSecTypes(mCurrentSection) = TYPE_TABLE Then
            ClearTableRowFields
            RefreshRowEditorLabels
        Else
            ExitRowEditMode
        End If

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

    If CloseMode = 0 Then
        If Not ConfirmDiscard() Then
            Cancel = True
        End If
    End If

    Exit Sub

ErrHandler:
    HandleFormError "UserForm_QueryClose"

End Sub
