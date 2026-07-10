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
Private Const TYPE_TABLE As String = "TABLE"       ' column count + rows

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

Private mCurrentSection As Long    ' 1-based index of selected section
Private mLoading As Boolean        ' suppress change events during load

Private Sub fraRowEditor_Click()

End Sub

Private Sub lstSubItems_Click()

End Sub

'====================================================
' FORM INITIALIZE
'====================================================

Private Sub UserForm_Initialize()

    mLoading = True

    InitSections
    RefreshSectionList
    LoadFromSheet

    mCurrentSection = 1
    lstSections.ListIndex = 0
    ShowSection 1

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

    If mCurrentSection > 0 Then
        SaveCurrentSection
    End If

    mCurrentSection = lstSections.ListIndex + 1
    ShowSection mCurrentSection

End Sub

'====================================================
' SHOW SECTION - switches the right panel content
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
    lblColCount.Visible = False
    spnColCount.Visible = False
    lblRows.Visible = False
    lstRows.Visible = False
    btnAddRow.Visible = False
    btnRemoveRow.Visible = False
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
    btnEditItem.Visible = True
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
    End If

End Sub

'====================================================
' TABLE SECTIONS
'====================================================

Private Sub ShowTableSection(ByVal idx As Long)

    lblColCount.Visible = True
    spnColCount.Visible = True

    Dim colCount As Long
    If mSecCols(idx) <> "" Then
        colCount = UBound(Split(mSecCols(idx), "~")) + 1
    Else
        colCount = 1
    End If

    mLoading = True
    spnColCount.Value = colCount
    mLoading = False

    lblRows.Visible = True
    lstRows.Visible = True
    btnAddRow.Visible = True
    btnRemoveRow.Visible = True
    lblRowEditor.Visible = True
    fraRowEditor.Visible = True

    RefreshRowList idx
    ClearRowEditor

End Sub

Private Sub spnColCount_Change()

    If mLoading Then Exit Sub

    Dim newCount As Long
    newCount = spnColCount.Value

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

    If mSecItems(mCurrentSection).count > 0 Then
        MsgBox "Column count changed. Existing row data may be misaligned. " & _
               "Review your rows after changing column count.", vbExclamation
    End If

    RefreshRowList mCurrentSection
    ClearRowEditor

End Sub

Private Sub RefreshRowList(ByVal idx As Long)

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

End Sub

Private Sub ClearRowEditor()

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

End Sub

Private Sub btnAddRow_Click()

    If mSecCols(mCurrentSection) = "" Then
        MsgBox "Set a column count before adding rows.", vbExclamation
        Exit Sub
    End If

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

Private Sub lstRows_Click()

    If mLoading Then Exit Sub
    If lstRows.ListIndex < 0 Then Exit Sub

    ' Save whatever was in the editor before switching rows,
    ' then refresh the list so the display updates immediately
    SaveRowEditor
    RefreshRowList mCurrentSection

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

    Dim i As Long
    Dim topPos As Long
    topPos = 6

    For i = 0 To UBound(cols)

        Dim lbl As MSForms.Label
        Set lbl = fraRowEditor.Controls.Add("Forms.Label.1", "lbl_col_" & i)
        lbl.Tag = "dynamic"
        lbl.Caption = "Column " & (i + 1) & ":"
        lbl.Left = 6
        lbl.Top = topPos
        lbl.Width = 70
        lbl.Height = 16

        Dim txt As MSForms.TextBox
        Set txt = fraRowEditor.Controls.Add("Forms.TextBox.1", "txt_col_" & i)
        txt.Tag = "dynamic"
        txt.Left = 80
        txt.Top = topPos
        txt.Width = 360
        txt.Height = 20

        If i <= UBound(parts) Then
            txt.Text = parts(i)
        End If

        topPos = topPos + 28

    Next i

    fraRowEditor.Tag = CStr(rowIdx)

    ' Keep the frame's visible height fixed; scroll if content overflows
    fraRowEditor.ScrollBars = 2   ' fmScrollBarsVertical
    fraRowEditor.ScrollHeight = topPos + 10
    If fraRowEditor.ScrollHeight < 40 Then fraRowEditor.ScrollHeight = 40

End Sub

Private Sub SaveRowEditor()

    If fraRowEditor.Tag = "" Then Exit Sub

    Dim rowIdx As Long
    rowIdx = CLng(fraRowEditor.Tag)

    If rowIdx + 1 > mSecItems(mCurrentSection).count Then Exit Sub
    If mSecCols(mCurrentSection) = "" Then Exit Sub

    Dim cols() As String
    cols = Split(mSecCols(mCurrentSection), "~")

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

    Dim newCol As New Collection
    For i = 1 To mSecItems(mCurrentSection).count
        If i = rowIdx + 1 Then
            newCol.Add newRow
        Else
            newCol.Add mSecItems(mCurrentSection)(i)
        End If
    Next i
    Set mSecItems(mCurrentSection) = newCol

End Sub

'====================================================
' EDIT MAIN ITEM / SUB-ITEM (List / Nested sections)
'====================================================

Private Sub btnEditItem_Click()

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

    Dim newVal As String
    newVal = Trim(InputBox("Edit item:", "Edit", current))

    If newVal = "" Then Exit Sub

    Dim newCol As New Collection
    Dim i As Long
    For i = 1 To mSecItems(mCurrentSection).count
        If i = itemIdx Then
            If mSecTypes(mCurrentSection) = TYPE_NESTED Then
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

    If UBound(parts) < 1 Then Exit Sub

    Dim subs() As String
    subs = Split(parts(1), "~")

    If subIdx > UBound(subs) Then Exit Sub

    Dim newVal As String
    newVal = Trim(InputBox("Edit sub-item:", "Edit", subs(subIdx)))

    If newVal = "" Then Exit Sub

    subs(subIdx) = newVal

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
    If lstItems.ListIndex < 0 Then Exit Sub

    lstSubItems.Clear
    txtSubItem.Text = ""
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
            If Len(txt) > 50 Then txt = Left(txt, 47) & "..."
            lstSubItems.AddItem txt
        End If
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

    Dim parts() As String
    parts = Split(current, "|")

    If UBound(parts) = 0 Then
        current = parts(0) & "|" & newSub
    Else
        current = parts(0) & "|" & parts(1) & "~" & newSub
    End If

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

    lstSubItems.AddItem IIf(Len(newSub) > 50, Left(newSub, 47) & "...", newSub)
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
    newStr = parts(0)

    If UBound(parts) > 0 Then

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

    Dim choice1 As Long
    choice1 = MsgBox("Is this new section a paragraph description?" & vbCrLf & vbCrLf & _
                     "Yes = Simple paragraph text" & vbCrLf & _
                     "No  = More advanced options (list or table)", _
                     vbYesNo + vbQuestion, "Section Type")

    If choice1 = vbYes Then
        secType = TYPE_PLAIN
    Else
        Dim choice2 As Long
        choice2 = MsgBox("Is this new section a bulleted list?" & vbCrLf & vbCrLf & _
                         "Yes = Nested text" & vbCrLf & _
                         "No  = Table with rows and columns", _
                         vbYesNo + vbQuestion, "Advanced Section Type")

        If choice2 = vbYes Then
            secType = TYPE_NESTED
        Else
            secType = TYPE_TABLE
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

            Dim subsFound As String
            subsFound = ""

            For c = parentCol + 1 To lastBulletCol
                Dim subVal As String
                subVal = Trim(tbl.DataBodyRange(r, c).Value)
                If subVal <> "" Then
                    If subsFound = "" Then
                        subsFound = subVal
                    Else
                        subsFound = subsFound & "~" & subVal
                    End If
                End If
            Next c

            If subsFound <> "" Then entry = entry & "|" & subsFound

            mSecItems(secIdx).Add entry

        End If

    Next r

End Sub

Private Sub LoadTableColumns(ByVal tbl As ListObject, _
                              ByVal col As Long, _
                              ByVal secIdx As Long)

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

    ' Column count only — names aren't preserved on the sheet
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

        If Replace(rowStr, "|", "") <> "" Then
            mSecItems(secIdx).Add rowStr
        End If

    Next r

End Sub

'====================================================
' SAVE TO SHEET
'====================================================

Private Sub btnSave_Click()

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

End Sub

Private Sub ClearSheet(ByVal ws As Worksheet)

    ws.Cells.Clear

End Sub

Private Sub WriteToSheet(ByVal ws As Worksheet)

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

        End Select

    Next i

    Dim LastRow As Long
    Dim lastCol As Long
    LastRow = ws.Cells(ws.Rows.count, 1).End(-4162).Row    ' xlUp
    lastCol = ws.Cells(1, ws.Columns.count).End(-4159).Column  ' xlToLeft

    If LastRow >= 2 Then
        Dim tblRange As Range
        Set tblRange = ws.Range(ws.Cells(1, 1), ws.Cells(LastRow, lastCol))
        ws.ListObjects.Add(1, tblRange, , 1).name = "SODTable"  ' xlSrcRange, xlYes
    End If

End Sub

Private Function MaxSubDepth(ByVal secIdx As Long) As Long

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

    ws.Cells(1, startCol + 1).Value = "Bullet1"

    Dim r As Long
    r = 2

    Dim i As Long
    For i = 1 To mSecItems(secIdx).count

        Dim parts() As String
        parts = Split(mSecItems(secIdx)(i), "|")

        ws.Cells(r, startCol).Value = parts(0)

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

Private Sub WriteTableToSheet(ByVal ws As Worksheet, _
                               ByVal secIdx As Long, _
                               ByVal startCol As Long)

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
            End If
        Next c

        r = r + 1

    Next i

End Sub

'====================================================
' CANCEL
'====================================================

Private Sub btnCancel_Click()

    Unload Me

End Sub

