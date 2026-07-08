VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSOD 
   Caption         =   "SOD Form"
   ClientHeight    =   9420.001
   ClientLeft      =   240
   ClientTop       =   930
   ClientWidth     =   12165
   OleObjectBlob   =   "frmSOD.frx":0000
   StartUpPosition =   1  'CenterOwner
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
Private Const SEC_STEPS As String = "Process Steps"
Private Const SEC_KPIS As String = "Key Performance Indicators"
Private Const SEC_RESOURCES As String = "Additional Resources"

' Section types
Private Const TYPE_PLAIN As String = "PLAIN"       ' single TextBox
Private Const TYPE_NESTED As String = "NESTED"     ' items + sub-items
Private Const TYPE_LIST As String = "LIST"         ' items only, no sub-items

' Data store: parallel arrays indexed by section
Private mSecNames() As String      ' display name
Private mSecTypes() As String      ' TYPE_PLAIN / TYPE_NESTED / TYPE_LIST
Private mSecCount As Long

' Each section's data stored as pipe-delimited strings
' PLAIN: mSecData(i) = full text
' LIST:  mSecItems(i) = Collection of strings
' NESTED: mSecItems(i) = Collection of "item|sub1|sub2|sub3..."
Private mSecData() As String
Private mSecItems() As Collection

Private mCurrentSection As Long    ' 1-based index of selected section
Private mLoading As Boolean        ' suppress change events during load

' Table row data stored as pipe-delimited strings, one entry per row
' mSecItems(i) is reused: each item = "col1val|col2val|col3val"
' mSecCols(i) stores column header names as tilde-delimited string
Private mSecCols() As String

Private Const TYPE_TABLE As String = "TABLE"

Private Sub fraContent_Click()

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub fraRowEditor_Click()

End Sub

'====================================================
' FORM INITIALIZE
'====================================================

Private Sub UserForm_Initialize()

    mLoading = True

    ' Build default section list
    InitSections

    ' Populate the ListBox
    RefreshSectionList

    ' Load existing sheet data if present
    LoadFromSheet

    ' Select first section
    If mSecCount > 0 Then
        lstSections.ListIndex = 0
        mCurrentSection = 1
        ShowSection 1
    End If

    mLoading = False

End Sub

Private Sub InitSections()

    mSecCount = 7
    ReDim mSecNames(1 To mSecCount)
    ReDim mSecTypes(1 To mSecCount)
    ReDim mSecData(1 To mSecCount)
    ReDim mSecItems(1 To mSecCount)
    ReDim mSecCols(1 To mSecCount)

    mSecNames(1) = SEC_TITLE
    mSecNames(2) = SEC_PURPOSE
    mSecNames(3) = SEC_SCOPE
    mSecNames(4) = SEC_ROLES
    mSecNames(5) = SEC_STEPS
    mSecNames(6) = SEC_KPIS
    mSecNames(7) = SEC_RESOURCES

    mSecTypes(1) = TYPE_PLAIN
    mSecTypes(2) = TYPE_PLAIN
    mSecTypes(3) = TYPE_PLAIN
    mSecTypes(4) = TYPE_NESTED
    mSecTypes(5) = TYPE_NESTED
    mSecTypes(6) = TYPE_LIST
    mSecTypes(7) = TYPE_NESTED

    Dim i As Long
    For i = 1 To mSecCount
        mSecData(i) = ""
        mSecCols(i) = ""
        Set mSecItems(i) = New Collection
    Next i

End Sub

'====================================================
' SECTION LIST
'====================================================

Private Sub RefreshSectionList()

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

End Sub

Private Sub lstSections_Click()

    If mLoading Then Exit Sub

    ' Save current section before switching
    If mCurrentSection > 0 Then
        SaveCurrentSection
    End If

    mCurrentSection = lstSections.ListIndex + 1
    ShowSection mCurrentSection

End Sub

'====================================================
' SHOW SECTION — switches the right panel content
'====================================================

Private Sub ShowSection(ByVal idx As Long)

    If idx < 1 Or idx > mSecCount Then Exit Sub

    mLoading = True

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

    End Select

    mLoading = False

End Sub

Private Sub ShowTableSection(ByVal idx As Long)

    ' Show column header controls
    lblColumns.Visible = True
    lstColumns.Visible = True
    btnAddCol.Visible = True
    btnRemoveCol.Visible = True

    ' Show row controls
    lblRows.Visible = True
    lstRows.Visible = True
    btnAddRow.Visible = True
    btnRemoveRow.Visible = True
    btnEditRow.Visible = True
    lblRowEditor.Visible = True
    fraRowEditor.Visible = True

    ' Load column headers
    lstColumns.Clear
    If mSecCols(idx) <> "" Then
        Dim cols() As String
        cols = Split(mSecCols(idx), "~")
        Dim i As Long
        For i = 0 To UBound(cols)
            If cols(i) <> "" Then lstColumns.AddItem cols(i)
        Next i
    End If

    ' Load rows
    RefreshRowList idx

    ' Clear editor
    ClearRowEditor

End Sub

Private Sub RefreshRowList(ByVal idx As Long)

    lstRows.Clear

    Dim i As Long
    For i = 1 To mSecItems(idx).count
        ' Display row as pipe-separated values
        Dim parts() As String
        parts = Split(mSecItems(idx)(i), "|")
        lstRows.AddItem Join(parts, "  |  ")
    Next i

End Sub

Private Sub ClearRowEditor()

    ' Remove any dynamically created controls from fraRowEditor
    Dim ctrl As Control
    Dim toDelete() As String
    Dim count As Long
    count = 0

    For Each ctrl In fraRowEditor.Controls
        count = count + 1
        ReDim Preserve toDelete(1 To count)
        toDelete(count) = ctrl.name
    Next ctrl

    Dim j As Long
    For j = 1 To count
        fraRowEditor.Controls.Remove toDelete(j)
    Next j

End Sub

Private Sub btnAddCol_Click()

    Dim colName As String
    colName = Trim(InputBox("Enter column header name:", "Add Column"))

    If colName = "" Then Exit Sub

    ' Append to tilde-delimited column string
    If mSecCols(mCurrentSection) = "" Then
        mSecCols(mCurrentSection) = colName
    Else
        mSecCols(mCurrentSection) = mSecCols(mCurrentSection) & "~" & colName
    End If

    lstColumns.AddItem colName

End Sub

Private Sub btnRemoveCol_Click()

    Dim selIdx As Long
    selIdx = lstColumns.ListIndex

    If selIdx < 0 Then
        MsgBox "Select a column to remove.", vbExclamation
        Exit Sub
    End If

    If mSecItems(mCurrentSection).count > 0 Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("Removing a column will delete that column's data from all rows. Continue?", _
                      vbYesNo + vbExclamation, "Remove Column")
        If resp = vbNo Then Exit Sub
    End If

    ' Rebuild column string without removed column
    Dim cols() As String
    cols = Split(mSecCols(mCurrentSection), "~")

    Dim newCols As String
    Dim i As Long
    For i = 0 To UBound(cols)
        If i <> selIdx Then
            If newCols = "" Then
                newCols = cols(i)
            Else
                newCols = newCols & "~" & cols(i)
            End If
        End If
    Next i
    mSecCols(mCurrentSection) = newCols

    ' Rebuild all rows removing that column's value
    Dim newCol As New Collection
    For i = 1 To mSecItems(mCurrentSection).count
        Dim parts() As String
        parts = Split(mSecItems(mCurrentSection)(i), "|")
        Dim newRow As String
        Dim j As Long
        For j = 0 To UBound(parts)
            If j <> selIdx Then
                If newRow = "" Then
                    newRow = parts(j)
                Else
                    newRow = newRow & "|" & parts(j)
                End If
            End If
        Next j
        newCol.Add newRow
    Next i
    Set mSecItems(mCurrentSection) = newCol

    lstColumns.RemoveItem selIdx
    RefreshRowList mCurrentSection

End Sub

Private Sub btnAddRow_Click()

    If mSecCols(mCurrentSection) = "" Then
        MsgBox "Add at least one column before adding rows.", vbExclamation
        Exit Sub
    End If

    ' Build an empty row with one slot per column
    Dim cols() As String
    cols = Split(mSecCols(mCurrentSection), "~")

    Dim emptyRow As String
    Dim i As Long
    For i = 0 To UBound(cols)
        If i = 0 Then
            emptyRow = ""
        Else
            emptyRow = emptyRow & "|"
        End If
    Next i

    mSecItems(mCurrentSection).Add emptyRow
    RefreshRowList mCurrentSection

    ' Auto-select and open editor for new row
    lstRows.ListIndex = lstRows.ListCount - 1
    OpenRowEditor lstRows.ListCount - 1

End Sub

Private Sub btnRemoveRow_Click()

    Dim selIdx As Long
    selIdx = lstRows.ListIndex

    If selIdx < 0 Then
        MsgBox "Select a row to remove.", vbExclamation
        Exit Sub
    End If

    Dim newCol As New Collection
    Dim i As Long
    For i = 1 To mSecItems(mCurrentSection).count
        If i <> selIdx + 1 Then
            newCol.Add mSecItems(mCurrentSection)(i)
        End If
    Next i
    Set mSecItems(mCurrentSection) = newCol

    RefreshRowList mCurrentSection
    ClearRowEditor

End Sub

Private Sub btnEditRow_Click()

    Dim selIdx As Long
    selIdx = lstRows.ListIndex

    If selIdx < 0 Then
        MsgBox "Select a row to edit.", vbExclamation
        Exit Sub
    End If

    OpenRowEditor selIdx

End Sub

Private Sub lstRows_Click()

    If mLoading Then Exit Sub
    If lstRows.ListIndex < 0 Then Exit Sub
    OpenRowEditor lstRows.ListIndex

End Sub

Private Sub OpenRowEditor(ByVal rowIdx As Long)

    ClearRowEditor

    If mSecCols(mCurrentSection) = "" Then Exit Sub

    Dim cols() As String
    cols = Split(mSecCols(mCurrentSection), "~")

    Dim parts() As String
    Dim rowData As String
    rowData = mSecItems(mCurrentSection)(rowIdx + 1)
    parts = Split(rowData, "|")

    ' Dynamically create one Label + TextBox per column inside fraRowEditor
    Dim i As Long
    Dim topPos As Long
    topPos = 6

    For i = 0 To UBound(cols)

        ' Label
        Dim lbl As MSForms.Label
        Set lbl = fraRowEditor.Controls.Add("Forms.Label.1", "lbl_col_" & i)
        lbl.Caption = cols(i) & ":"
        lbl.Left = 6
        lbl.Top = topPos
        lbl.Width = 100
        lbl.Height = 16

        ' TextBox
        Dim txt As MSForms.TextBox
        Set txt = fraRowEditor.Controls.Add("Forms.TextBox.1", "txt_col_" & i)
        txt.Left = 110
        txt.Top = topPos
        txt.Width = 330
        txt.Height = 20

        If i <= UBound(parts) Then
            txt.Text = parts(i)
        End If

        topPos = topPos + 28

    Next i

    ' Store which row is being edited
    fraRowEditor.Tag = CStr(rowIdx)

    ' Resize frame to fit
    fraRowEditor.Height = topPos + 10
    If fraRowEditor.Height < 40 Then fraRowEditor.Height = 40

End Sub

Private Sub SaveRowEditor()

    If fraRowEditor.Tag = "" Then Exit Sub

    Dim rowIdx As Long
    rowIdx = CLng(fraRowEditor.Tag)

    If rowIdx + 1 > mSecItems(mCurrentSection).count Then Exit Sub

    If mSecCols(mCurrentSection) = "" Then Exit Sub

    Dim cols() As String
    cols = Split(mSecCols(mCurrentSection), "~")

    ' Read values from dynamically created TextBoxes
    Dim newRow As String
    Dim i As Long
    For i = 0 To UBound(cols)
        Dim txt As MSForms.Control
        On Error Resume Next
        Set txt = fraRowEditor.Controls("txt_col_" & i)
        On Error GoTo 0

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

    ' Replace in collection
    Dim newCol As New Collection
    For i = 1 To mSecItems(mCurrentSection).count
        If i = rowIdx + 1 Then
            newCol.Add newRow
        Else
            newCol.Add mSecItems(mCurrentSection)(i)
        End If
    Next i
    Set mSecItems(mCurrentSection) = newCol

    RefreshRowList mCurrentSection

End Sub

Private Sub HideAllControls()

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

    ' Table controls
    lblColumns.Visible = False
    lstColumns.Visible = False
    btnAddCol.Visible = False
    btnRemoveCol.Visible = False
    lblRows.Visible = False
    lstRows.Visible = False
    btnAddRow.Visible = False
    btnRemoveRow.Visible = False
    btnEditRow.Visible = False
    lblRowEditor.Visible = False
    fraRowEditor.Visible = False

End Sub

Private Sub ShowPlainSection(ByVal idx As Long)

    txtContent.Text = mSecData(idx)
    txtContent.Visible = True

End Sub

Private Sub ShowListSection(ByVal idx As Long)

    lstItems.Clear
    Dim i As Long
    For i = 1 To mSecItems(idx).count
        lstItems.AddItem mSecItems(idx)(i)
    Next i

    lstItems.Visible = True
    btnAddItem.Visible = True
    btnRemoveItem.Visible = True
    btnEditItem.Visible = True          ' new
    txtItem.Visible = True
    txtItem.Text = ""

End Sub

Private Sub ShowNestedSection(ByVal idx As Long)

    lstItems.Clear
    Dim i As Long
    For i = 1 To mSecItems(idx).count
        lstItems.AddItem Split(mSecItems(idx)(i), "|")(0)
    Next i

    lstItems.Visible = True
    btnAddItem.Visible = True
    btnRemoveItem.Visible = True
    btnEditItem.Visible = True          ' new
    txtItem.Visible = True
    txtItem.Text = ""

    lblSubItems.Visible = True
    lstSubItems.Visible = True
    btnAddSubItem.Visible = True
    btnRemoveSubItem.Visible = True
    btnEditSubItem.Visible = True       ' new
    txtSubItem.Visible = True
    txtSubItem.Text = ""
    lstSubItems.Clear

    If lstItems.ListCount > 0 Then
        lstItems.ListIndex = 0
        RefreshSubItems
    End If

End Sub

Private Sub btnEditItem_Click()

    Dim selIdx As Long
    selIdx = lstItems.ListIndex

    If selIdx < 0 Then
        MsgBox "Select an item to edit.", vbExclamation
        Exit Sub
    End If

    Dim itemIdx As Long
    itemIdx = selIdx + 1

    ' For NESTED, show only the main part (before first pipe)
    Dim current As String
    If mSecTypes(mCurrentSection) = TYPE_NESTED Then
        current = Split(mSecItems(mCurrentSection)(itemIdx), "|")(0)
    Else
        current = mSecItems(mCurrentSection)(itemIdx)
    End If

    Dim newVal As String
    newVal = Trim(InputBox("Edit item:", "Edit", current))

    If newVal = "" Then Exit Sub

    ' Rebuild collection with updated value
    Dim newCol As New Collection
    Dim i As Long
    For i = 1 To mSecItems(mCurrentSection).count
        If i = itemIdx Then
            If mSecTypes(mCurrentSection) = TYPE_NESTED Then
                ' Preserve existing sub-items after the pipe
                Dim parts() As String
                parts = Split(mSecItems(mCurrentSection)(i), "|")
                parts(0) = newVal
                newCol.Add Join(parts, "|")
            Else
                newCol.Add newVal
            End If
        Else
            newCol.Add mSecItems(mCurrentSection)(i)
        End If
    Next i
    Set mSecItems(mCurrentSection) = newCol

    ' Refresh display
    lstItems.List(selIdx) = newVal

End Sub

Private Sub btnEditSubItem_Click()

    Dim parentIdx As Long
    parentIdx = lstItems.ListIndex

    If parentIdx < 0 Then
        MsgBox "Select a parent item first.", vbExclamation
        Exit Sub
    End If

    Dim subIdx As Long
    subIdx = lstSubItems.ListIndex

    If subIdx < 0 Then
        MsgBox "Select a sub-item to edit.", vbExclamation
        Exit Sub
    End If

    Dim itemIdx As Long
    itemIdx = parentIdx + 1

    Dim parts() As String
    parts = Split(mSecItems(mCurrentSection)(itemIdx), "|")

    Dim subs() As String
    subs = Split(parts(1), "~")

    Dim newVal As String
    newVal = Trim(InputBox("Edit sub-item:", "Edit", subs(subIdx)))

    If newVal = "" Then Exit Sub

    subs(subIdx) = newVal

    ' Rebuild
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

    lstSubItems.List(subIdx) = newVal

End Sub

'====================================================
' SAVE CURRENT SECTION DATA TO MEMORY
'====================================================

Private Sub SaveCurrentSection()

    If mCurrentSection < 1 Or mCurrentSection > mSecCount Then Exit Sub

    Select Case mSecTypes(mCurrentSection)

        Case TYPE_PLAIN
            mSecData(mCurrentSection) = txtContent.Text

        Case TYPE_TABLE
            SaveRowEditor

    End Select

End Sub

'====================================================
' SUB-ITEMS (for NESTED sections)
'====================================================

Private Sub lstItems_Click()

    If mLoading Then Exit Sub
    If mSecTypes(mCurrentSection) <> TYPE_NESTED Then Exit Sub

    RefreshSubItems

End Sub

Private Sub RefreshSubItems()

    lstSubItems.Clear

    Dim selIdx As Long
    selIdx = lstItems.ListIndex
    If selIdx < 0 Then Exit Sub

    Dim itemIdx As Long
    itemIdx = selIdx + 1

    If itemIdx > mSecItems(mCurrentSection).count Then Exit Sub

    Dim parts() As String
    parts = Split(mSecItems(mCurrentSection)(itemIdx), "|")

    ' parts(0) = main item; parts(1) = tilde-delimited sub-items
    If UBound(parts) = 0 Then Exit Sub
    If parts(1) = "" Then Exit Sub

    Dim subs() As String
    subs = Split(parts(1), "~")

    Dim i As Long
    For i = 0 To UBound(subs)
        If subs(i) <> "" Then lstSubItems.AddItem subs(i)
    Next i

End Sub

'====================================================
' ADD / REMOVE MAIN ITEMS
'====================================================

Private Sub btnAddItem_Click()

    Dim newItem As String
    newItem = Trim(txtItem.Text)

    If newItem = "" Then
        MsgBox "Please type an item before adding.", vbExclamation
        Exit Sub
    End If

    mSecItems(mCurrentSection).Add newItem
    lstItems.AddItem newItem
    txtItem.Text = ""

End Sub

Private Sub btnRemoveItem_Click()

    Dim selIdx As Long
    selIdx = lstItems.ListIndex

    If selIdx < 0 Then
        MsgBox "Select an item to remove.", vbExclamation
        Exit Sub
    End If

    ' Remove from collection (1-based)
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

End Sub

'====================================================
' ADD / REMOVE SUB-ITEMS
'====================================================

Private Sub btnAddSubItem_Click()

    Dim selIdx As Long
    selIdx = lstItems.ListIndex

    If selIdx < 0 Then
        MsgBox "Select a parent item first.", vbExclamation
        Exit Sub
    End If

    Dim newSub As String
    newSub = Trim(txtSubItem.Text)

    If newSub = "" Then
        MsgBox "Please type a sub-item before adding.", vbExclamation
        Exit Sub
    End If

    Dim itemIdx As Long
    itemIdx = selIdx + 1

    Dim current As String
    current = mSecItems(mCurrentSection)(itemIdx)

    ' Split on pipe: parts(0) = main item, parts(1) = tilde-delimited sub-items
    Dim parts() As String
    parts = Split(current, "|")

    If UBound(parts) = 0 Then
        ' No sub-items yet — create the sub-items segment
        current = parts(0) & "|" & newSub
    Else
        ' Append to existing sub-items segment with tilde delimiter
        current = parts(0) & "|" & parts(1) & "~" & newSub
    End If

    ' Replace in collection
    Dim newCol As New Collection
    Dim i As Long
    For i = 1 To mSecItems(mCurrentSection).count
        If i = itemIdx Then
            newCol.Add current
        Else
            newCol.Add mSecItems(mCurrentSection)(i)
        End If
    Next i
    Set mSecItems(mCurrentSection) = newCol

    lstSubItems.AddItem newSub
    txtSubItem.Text = ""

End Sub

Private Sub btnRemoveSubItem_Click()

    Dim parentIdx As Long
    parentIdx = lstItems.ListIndex

    If parentIdx < 0 Then Exit Sub

    Dim subIdx As Long
    subIdx = lstSubItems.ListIndex

    If subIdx < 0 Then
        MsgBox "Select a sub-item to remove.", vbExclamation
        Exit Sub
    End If

    Dim itemIdx As Long
    itemIdx = parentIdx + 1

    Dim parts() As String
    parts = Split(mSecItems(mCurrentSection)(itemIdx), "|")

    Dim newStr As String
    newStr = parts(0)   ' always keep the main item name

    If UBound(parts) > 0 Then
        ' Rebuild tilde-delimited sub-items without the removed one
        Dim subs() As String
        subs = Split(parts(1), "~")

        Dim newSubs As String
        Dim i As Long
        For i = 0 To UBound(subs)
            If i <> subIdx Then
                If newSubs = "" Then
                    newSubs = subs(i)
                Else
                    newSubs = newSubs & "~" & subs(i)
                End If
            End If
        Next i

        If newSubs <> "" Then newStr = newStr & "|" & newSubs
    End If

    ' Replace in collection
    Dim newCol As New Collection
    For i = 1 To mSecItems(mCurrentSection).count
        If i = itemIdx Then
            newCol.Add newStr
        Else
            newCol.Add mSecItems(mCurrentSection)(i)
        End If
    Next i
    Set mSecItems(mCurrentSection) = newCol

    lstSubItems.RemoveItem subIdx

End Sub

'====================================================
' ADD NEW GENERIC SECTION
'====================================================

Private Sub btnAddSection_Click()

    Dim secName As String
    secName = Trim(InputBox("Enter section name:", "Add Section"))

    If secName = "" Then Exit Sub

    Dim secType As String

    Dim choice As Long
    choice = MsgBox("What type of section is this?" & vbCrLf & vbCrLf & _
                    "Yes = Table" & vbCrLf & _
                    "No = Bullets or Paragraph", _
                    vbYesNoCancel + vbQuestion, "Section Type")

    If choice = vbCancel Then Exit Sub

    If choice = vbYes Then
        secType = TYPE_TABLE
    Else
        Dim choice2 As Long
        choice2 = MsgBox("Bullet list or plain paragraph?" & vbCrLf & vbCrLf & _
                         "Yes = Bullets (items with optional sub-items)" & vbCrLf & _
                         "No = Paragraph (plain text)", _
                         vbYesNoCancel + vbQuestion, "Section Type")

        If choice2 = vbCancel Then Exit Sub

        If choice2 = vbYes Then
            secType = TYPE_NESTED
        Else
            secType = TYPE_PLAIN
        End If
    End If

    mSecCount = mSecCount + 1
    ReDim Preserve mSecNames(1 To mSecCount)
    ReDim Preserve mSecTypes(1 To mSecCount)
    ReDim Preserve mSecData(1 To mSecCount)
    ReDim Preserve mSecItems(1 To mSecCount)
    ReDim Preserve mSecCols(1 To mSecCount)

    mSecNames(mSecCount) = secName
    mSecTypes(mSecCount) = secType
    mSecData(mSecCount) = ""
    mSecCols(mSecCount) = ""
    Set mSecItems(mSecCount) = New Collection

    RefreshSectionList

    SaveCurrentSection
    lstSections.ListIndex = mSecCount - 1
    mCurrentSection = mSecCount
    ShowSection mCurrentSection

End Sub

'====================================================
' LOAD FROM SHEET
'====================================================

Private Sub LoadFromSheet()

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

            End Select

        End If

    Next col

End Sub

Private Sub LoadTableColumns(ByVal tbl As ListObject, _
                              ByVal col As Long, _
                              ByVal secIdx As Long)

    ' Find contiguous Table# columns
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

    ' Build column header string from Table# column names
    ' (use row 1 values if present, otherwise use column names)
    Dim colHeaders As String
    For c = col To lastTableCol
        Dim hdr As String
        hdr = Trim(tbl.ListColumns(c).name)
        If colHeaders = "" Then
            colHeaders = hdr
        Else
            colHeaders = colHeaders & "~" & hdr
        End If
    Next c
    mSecCols(secIdx) = colHeaders

    ' Load rows
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

        ' Only add non-empty rows
        If Replace(rowStr, "|", "") <> "" Then
            mSecItems(secIdx).Add rowStr
        End If

    Next r

End Sub

Private Function FindSection(ByVal name As String) As Long

    Dim i As Long
    For i = 1 To mSecCount
        If LCase(Trim(mSecNames(i))) = LCase(Trim(name)) Then
            FindSection = i
            Exit Function
        End If
    Next i

    FindSection = 0

End Function

Private Sub LoadPlainColumn(ByVal tbl As ListObject, _
                             ByVal col As Long, _
                             ByVal secIdx As Long)

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

End Sub

Private Sub LoadListColumn(ByVal tbl As ListObject, _
                            ByVal col As Long, _
                            ByVal secIdx As Long)

    Dim r As Long
    Dim txt As String

    Set mSecItems(secIdx) = New Collection

    For r = 1 To tbl.ListRows.count
        txt = Trim(tbl.DataBodyRange(r, col).Value)
        If txt <> "" Then mSecItems(secIdx).Add txt
    Next r

End Sub

Private Sub LoadNestedColumns(ByVal tbl As ListObject, _
                               ByVal parentCol As Long, _
                               ByVal secIdx As Long)

    ' Find contiguous Bullet columns after parentCol
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
    For r = 1 To tbl.ListRows.count

        Dim mainVal As String
        mainVal = Trim(tbl.DataBodyRange(r, parentCol).Value)

        If mainVal <> "" Then

            Dim entry As String
            entry = mainVal

            For c = parentCol + 1 To lastBulletCol
                Dim subVal As String
                subVal = Trim(tbl.DataBodyRange(r, c).Value)
                If subVal <> "" Then entry = entry & "|" & subVal
            Next c

            mSecItems(secIdx).Add entry

        End If

    Next r

End Sub

'====================================================
' SAVE TO SHEET
'====================================================

Private Sub btnSave_Click()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Warn if sheet already has data
    If ws.ListObjects.count > 0 Then
        If ws.ListObjects(1).ListRows.count > 0 Then
            Dim resp As VbMsgBoxResult
            resp = MsgBox("This sheet already has data. Saving will overwrite it." & vbCrLf & vbCrLf & _
                          "Continue?", vbYesNo + vbExclamation, "Overwrite Warning")
            If resp = vbNo Then Exit Sub
        End If
    End If

    ' Save whatever section is currently visible
    SaveCurrentSection

    ' Clear existing table or create new one
    ClearSheet ws

    ' Build column headers and data
    WriteToSheet ws

    MsgBox "Saved to sheet successfully.", vbInformation, "SOD Form"

End Sub

Private Sub ClearSheet(ByVal ws As Worksheet)

    ws.Cells.Clear

End Sub

Private Sub WriteToSheet(ByVal ws As Worksheet)

    Dim headers() As String
    Dim colCount As Long
    colCount = 0

    ' First pass: count total columns needed
    Dim i As Long
    For i = 1 To mSecCount

        Select Case mSecTypes(i)

            Case TYPE_PLAIN
                colCount = colCount + 1

            Case TYPE_LIST
                colCount = colCount + 1

            Case TYPE_NESTED
                ' Parent column + max sub-item depth across all items
                Dim maxDepth As Long
                maxDepth = MaxSubDepth(i)
                colCount = colCount + 1 + maxDepth
                
            Case TYPE_TABLE
                WriteTableToSheet ws, i, currentCol
                ' Count columns used: 1 header col + Table# helper cols
                Dim tblColCount As Long
                If mSecCols(i) <> "" Then
                    tblColCount = UBound(Split(mSecCols(i), "~")) + 1
                Else
                    tblColCount = 1
                End If
                currentCol = currentCol + tblColCount

        End Select

    Next i

    ' Write headers and data column by column
    Dim currentCol As Long
    currentCol = 1

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

        End Select

    Next i

    ' Create Excel table from the written data
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(-4162).Row  ' xlUp
    lastCol = ws.Cells(1, ws.Columns.count).End(-4159).Column  ' xlToLeft

    If lastRow >= 2 Then
        Dim tblRange As Range
        Set tblRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        ws.ListObjects.Add(1, tblRange, , 1).name = "SODTable"  ' xlSrcRange, xlYes header
    End If

End Sub

Private Sub WriteTableToSheet(ByVal ws As Worksheet, _
                               ByVal secIdx As Long, _
                               ByVal startCol As Long)

    If mSecCols(secIdx) = "" Then Exit Sub

    Dim cols() As String
    cols = Split(mSecCols(secIdx), "~")

    ' Write section name in first column
    ws.Cells(1, startCol).Value = mSecNames(secIdx)

    ' Write Table# helper column headers
    Dim c As Long
    For c = 0 To UBound(cols)
        ws.Cells(1, startCol + c).Value = "Table" & (c + 1)
    Next c

    ' Override first column with section name (Table1 becomes the heading col)
    ws.Cells(1, startCol).Value = mSecNames(secIdx)

    ' Write row data
    Dim r As Long
    r = 2

    Dim i As Long
    For i = 1 To mSecItems(secIdx).count

        Dim parts() As String
        parts = Split(mSecItems(secIdx)(i), "|")

        For c = 0 To UBound(cols)
            If c <= UBound(parts) Then
                ws.Cells(r, startCol + c).Value = Trim(parts(c))
            End If
        Next c

        r = r + 1

    Next i

End Sub

Private Function MaxSubDepth(ByVal secIdx As Long) As Long

    ' Always exactly 1 bullet column for nested sections
    Dim i As Long
    For i = 1 To mSecItems(secIdx).count
        Dim parts() As String
        parts = Split(mSecItems(secIdx)(i), "|")
        If UBound(parts) > 0 And parts(1) <> "" Then
            MaxSubDepth = 1
            Exit Function
        End If
    Next i

    MaxSubDepth = 0

End Function

Private Sub WritePlainToSheet(ByVal ws As Worksheet, _
                               ByVal secIdx As Long, _
                               ByVal col As Long)

    Dim lines() As String
    lines = Split(mSecData(secIdx), vbCrLf)

    Dim r As Long
    For r = 0 To UBound(lines)
        If Trim(lines(r)) <> "" Then
            ws.Cells(r + 2, col).Value = Trim(lines(r))
        End If
    Next r

End Sub

Private Sub WriteListToSheet(ByVal ws As Worksheet, _
                              ByVal secIdx As Long, _
                              ByVal col As Long)

    Dim i As Long
    For i = 1 To mSecItems(secIdx).count
        ws.Cells(i + 1, col).Value = mSecItems(secIdx)(i)
    Next i

End Sub

Private Sub WriteNestedToSheet(ByVal ws As Worksheet, _
                                ByVal secIdx As Long, _
                                ByVal startCol As Long, _
                                ByVal depth As Long)

    ' Always exactly one Bullet1 helper column — no deeper nesting
    ws.Cells(1, startCol + 1).Value = "Bullet1"

    Dim r As Long
    r = 2

    Dim i As Long
    For i = 1 To mSecItems(secIdx).count

        Dim parts() As String
        parts = Split(mSecItems(secIdx)(i), "|")

        ' Main item in parent column
        ws.Cells(r, startCol).Value = parts(0)

        ' Sub-items each go on their own row in Bullet1 column
        If UBound(parts) > 0 And parts(1) <> "" Then

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

End Sub

'====================================================
' CANCEL
'====================================================

Private Sub btnCancel_Click()

    Unload Me

End Sub

