VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSOD 
   Caption         =   "SOD Editor"
   ClientHeight    =   9156.001
   ClientLeft      =   987
   ClientTop       =   3766
   ClientWidth     =   11060
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

' Internal field/list delimiters. Non-printable characters, chosen
' specifically because they cannot be typed or pasted by a user
' (e.g. RACI entries legitimately use the "|" character), so no real content can
' ever collide with them.
' NOTE: VBA's Const only accepts literals, not function calls like Chr(31),
' so these are plain module-level variables instead, set once at the very
' top of UserForm_Initialize before anything else can use them.
Private TABLE_SEP As String  ' = Chr(31), replaces the old literal pipe character - separates fields/columns within a row
Private LIST_SEP As String   ' = Chr(30), replaces the old literal tilde character - separates sub-items / column names within a field
Private SUBSUB_SEP As String   ' = Chr(29) - separates a sub-item's own text from its sub-sub-items blob
Private SUBSUB_ITEM_SEP As String   ' = Chr(28) - joins multiple sub-sub-items within that blob

Private Const LIST_DISPLAY_MAX As Long = 88

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

Private mDeveloperMode As Boolean      ' when True, removes guardrails and enables debug output
Private mSkipConfirmPrompts As Boolean   ' suppresses nested duplicate prompts when the caller already confirmed
Private mShowingUnconfirmedPrompt As Boolean

' -1 = not drilled into any sub-item, otherwise 0-based index of
' the sub-item (within its parent Item) currently drilled into.
' Mirrors mEditingSubIdx's convention but one level deeper.
Private mDrilledSubIdx As Long       ' -1 = not drilled in, else 0-based index of the sub-item (within current Item) whose children we're viewing
Private mEditingSubSubIdx As Long    ' -1 = not editing, else 0-based index of the sub-sub-item being edited within the drilled-into sub-item
                                      
Private mSubSubEditSnapshotSec As Long
Private mSubSubEditSnapshotSubIdx As Long   ' which sub-item (lstSubItems index) we were drilled into when the edit snapshot was taken

Private mSecHasSubSubItems() As Boolean   ' TYPE_LIST only, and only meaningful when mSecHasSubItems(i) = True

Private mDrilledSubText As String   ' the parent sub-item's own display text, captured at drill-in time, used only for prompt messages while drilled in
                                              
Private mSubSubEditSnapshot As Collection

Private mUndoHasSubSubItems As Boolean
Private mUndoHasSubItems As Boolean

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

Private Sub Workbook_Open()

    If InStr(1, ThisWorkbook.path, "SharePoint Site Name", vbTextCompare) > 0 Or _
       InStr(1, ThisWorkbook.FullName, "https://", vbTextCompare) > 0 Then

        MsgBox "You're viewing the master copy of this file." & vbCrLf & vbCrLf & _
               "Please download your own copy before making changes - edits here won't be saved back correctly.", _
               vbExclamation, "Master File - Read Only"
    End If

End Sub

Private Sub btnDevMode_Click()
    On Error GoTo ErrHandler
    
    If Not mDeveloperMode Then
        ' Entering Dev Mode - require confirmation
        Dim resp As VbMsgBoxResult
        resp = MsgBox("Enable Developer Mode?" & vbCrLf & vbCrLf & _
                      "- Checkboxes unlocked" & vbCrLf & _
                      "- Can edit/delete built-in sections" & vbCrLf & _
                      "- No save prompt on close" & vbCrLf & _
                      "- Debug output enabled", _
                      vbYesNo + vbExclamation, "Developer Mode")
        If resp = vbNo Then Exit Sub
        
        mDeveloperMode = True
        btnDevMode.Visible = True
    Else
        ' Turning off Dev Mode - just confirm
        Dim resp2 As VbMsgBoxResult
        resp2 = MsgBox("Disable Developer Mode?" & vbCrLf & vbCrLf & _
                       "Guardrails will be restored.", _
                       vbYesNo + vbExclamation, "Developer Mode")
        If resp2 = vbNo Then Exit Sub
        
        mDeveloperMode = False
        btnDevMode.Visible = False
    End If

    ' Refresh to re-apply all restrictions
    UpdateSectionButtons
    ShowSection mCurrentSection 'Now runs on BOTH enable and disable
    Exit Sub
    
ErrHandler:
    HandleFormError "btnDevMode_Click"
    
End Sub

Private Sub btnDrillIntoSub_Click()
    On Error GoTo ErrHandler

    If Not mDeveloperMode And HasUnconfirmedContent() Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox(BuildUnconfirmedPromptText(), vbYesNoCancel + vbQuestion, "Unsaved Text")

        Select Case resp
            Case vbCancel
                Exit Sub

            Case vbYes
                If Not ConfirmPendingChanges() Then Exit Sub

            Case vbNo
                ' fall through - discard and proceed
        End Select
    End If

    If mDrilledSubIdx >= 0 Then
        DrillOutOfSub
    Else
        If lstSubItems.ListIndex < 0 Then Exit Sub
        DrillIntoSub lstSubItems.ListIndex
    End If

    Exit Sub

ErrHandler:
    HandleFormError "btnDrillIntoSub_Click"

End Sub

Private Sub DrillIntoSub(ByVal subIdx As Long)
    On Error GoTo ErrHandler

    Debug.Print "DrillIntoSub called. subIdx=" & subIdx & " lstItems.ListIndex before=" & lstItems.ListIndex

    mDrilledSubIdx = subIdx
    mDrilledSubText = lstSubItems.List(subIdx)

    lblSubItems.Caption = "Bullets under: " & TruncateForDisplay(mDrilledSubText)
    btnDrillIntoSub.Caption = "• Bullets"

    lstItems.Enabled = False
    btnAddItem.Enabled = False
    btnEditItem.Enabled = False
    btnRemoveItem.Enabled = False
    txtItem.Enabled = False
    txtItem.BackColor = &H8000000F

    txtSubItem.Text = ""
    RefreshSubSubItems
    UpdateSubSubItemButtonsBasedOnFocus

    Debug.Print "DrillIntoSub finished. lstItems.ListIndex after=" & lstItems.ListIndex

    Exit Sub

ErrHandler:
    HandleFormError "DrillIntoSub"

End Sub

Private Sub DrillOutOfSub()
    On Error GoTo ErrHandler

    If Len(Trim(txtSubItem.Text)) > 0 And mEditingSubSubIdx < 0 Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("Do you want to add """ & Trim(txtSubItem.Text) & """ to """ & mDrilledSubText & """?" & vbCrLf & vbCrLf & _
              "Otherwise it will be discarded.", _
              vbYesNoCancel + vbQuestion, "Unsaved Text")

        If resp = vbCancel Then
            Exit Sub
        ElseIf resp = vbYes Then
            AddSubSubItem
        End If
    End If

    If mEditingSubSubIdx >= 0 Then
        CancelSubSubItemEditMode
    End If

    mDrilledSubIdx = -1
    mDrilledSubText = ""
    lblSubItems.Caption = "Sub-items:"
    btnDrillIntoSub.Caption = "• Bullets"

    lstItems.Enabled = True
    txtItem.Enabled = True
    txtItem.BackColor = &H80000005
    UpdateItemButtonsBasedOnFocus

    txtSubItem.Text = ""
    RefreshSubItems
    UpdateSubItemButtonsBasedOnFocus

    Exit Sub

ErrHandler:
    HandleFormError "DrillOutOfSub"

End Sub

Private Sub chkSubSubItems_Click()
    On Error GoTo ErrHandler

    If mLoading Then Exit Sub

    If mSecTypes(mCurrentSection) <> TYPE_LIST Then Exit Sub
    
    If Not mSecHasSubItems(mCurrentSection) Then Exit Sub   ' shouldn't be reachable if enabled state is correct, but guard anyway

    ' Guard: don't let this checkbox blow past an active edit
    ' (item / sub-item / sub-sub-item / row) without confirming.
    If Not mDeveloperMode And HasUnconfirmedContent() Then
        Dim priorValue As Boolean
        priorValue = Not chkSubSubItems.Value   ' .Value already reflects the click; this recovers the pre-click state

        Dim guardResp As VbMsgBoxResult
        guardResp = MsgBox(BuildUnconfirmedPromptText(), vbYesNoCancel + vbQuestion, "Unsaved Text")

        Select Case guardResp
            Case vbCancel
                mLoading = True
                chkSubSubItems.Value = priorValue
                DoEvents
                mLoading = False
                Exit Sub

            Case vbYes
                If Not ConfirmPendingChanges() Then
                    mLoading = True
                    chkSubSubItems.Value = priorValue
                    DoEvents
                    mLoading = False
                    Exit Sub
                End If

            Case vbNo
                ' fall through - discard and proceed with the toggle
        End Select
    End If

    If chkSubSubItems.Value Then
        ' Switching ON: nothing needs to change about existing data -
        ' sub-items just gain the *option* of having their own children
        mSecHasSubSubItems(mCurrentSection) = True
    Else
        ' Switching OFF is destructive - warn if any sub-item actually
        ' has sub-sub-item data that would be discarded
        Dim hasAnySubSubs As Boolean
        hasAnySubSubs = False

        Dim i As Long, k As Long
        For i = 1 To mSecItems(mCurrentSection).count
            Dim parts() As String
            parts = Split(mSecItems(mCurrentSection)(i), TABLE_SEP)
            If UBound(parts) >= 1 Then
                Dim subs() As String
                subs = Split(parts(1), LIST_SEP)
                For k = 0 To UBound(subs)
                    If InStr(subs(k), SUBSUB_SEP) > 0 Then
                        hasAnySubSubs = True
                        Exit For
                    End If
                Next k
            End If
            If hasAnySubSubs Then Exit For
        Next i

        If hasAnySubSubs Then
            Dim resp As VbMsgBoxResult
            resp = MsgBox("Turning off sub-details will permanently remove all sub-detail data for this section." & vbCrLf & vbCrLf & _
                          "Continue?", vbYesNo + vbExclamation, "Remove Sub-details")

            If resp = vbNo Then
                mLoading = True
                chkSubSubItems.Value = True
                mLoading = False
                Exit Sub
            End If
        End If

        ' Strip everything after SUBSUB_SEP from every sub-item, in every item
        Dim newCol As New Collection
        Dim j As Long
        For j = 1 To mSecItems(mCurrentSection).count
            Dim parts2() As String
            parts2 = Split(mSecItems(mCurrentSection)(j), TABLE_SEP)

            If UBound(parts2) >= 1 Then
                Dim subs2() As String
                subs2 = Split(parts2(1), LIST_SEP)
                Dim m As Long
                For m = 0 To UBound(subs2)
                    If InStr(subs2(m), SUBSUB_SEP) > 0 Then
                        subs2(m) = Split(subs2(m), SUBSUB_SEP)(0)
                    End If
                Next m
                newCol.Add parts2(0) & TABLE_SEP & Join(subs2, LIST_SEP)
            Else
                newCol.Add mSecItems(mCurrentSection)(j)
            End If
        Next j

        Set mSecItems(mCurrentSection) = newCol
        mSecHasSubSubItems(mCurrentSection) = False

        ' If we were drilled in when this got switched off, back out
        If mDrilledSubIdx >= 0 Then DrillOutOfSub
    End If

    ShowListSection mCurrentSection
    Exit Sub

ErrHandler:
    HandleFormError "chkSubSubItems_Click"

End Sub

Private Sub fraContent_Click()

End Sub

Private Sub fraRowEditor_Click()

End Sub

Private Sub lblColCount_Click()

End Sub

Private Function TruncateForDisplay(ByVal txt As String) As String

    If Len(txt) > LIST_DISPLAY_MAX Then
        TruncateForDisplay = Left(txt, LIST_DISPLAY_MAX) & "..."
    Else
        TruncateForDisplay = txt
    End If
    
End Function

Private Sub lblSectionHelp_Click()
    On Error GoTo ErrHandler
    
    Dim helpText As String
    
    If mCurrentSection > 0 And mCurrentSection <= mSecCount Then
    
        Dim secName As String
        secName = mSecNames(mCurrentSection)
        If IsBuiltInSectionName(secName) Then
            ' Built-in specific instructions
            Select Case LCase(Trim(secName))
                Case LCase(SEC_TITLE):
                    helpText = "Title Section" & vbCrLf & vbCrLf & _
                              "Enter the title/name of the procedure or process." & vbCrLf & _
                              "This should be a clear, concise summary."
                
                Case LCase(SEC_GROUP):
                    helpText = "Group Section" & vbCrLf & vbCrLf & _
                              "Select the department or functional group responsible for this procedure." & vbCrLf & _
                              "Type or choose from the dropdown list."
                
                Case LCase(SEC_PURPOSE):
                    helpText = "Purpose Section" & vbCrLf & vbCrLf & _
                              "Explain WHY this procedure exists." & vbCrLf & _
                              "What problem does it solve?"
                
                Case LCase(SEC_SCOPE):
                    helpText = "Scope Section" & vbCrLf & vbCrLf & _
                              "Define what IS and IS NOT covered by this procedure." & vbCrLf & _
                              "Specify boundaries and exceptions."
                
                Case LCase(SEC_DICTIONARY):
                    helpText = "Dictionary Section" & vbCrLf & vbCrLf & _
                              "Define key terms, acronyms, and abbreviations used." & vbCrLf & _
                              "Each row: Term | Definition"
                
                Case LCase(SEC_ROLES):
                    helpText = "Roles Section" & vbCrLf & vbCrLf & _
                              "List roles involved in this process." & vbCrLf & _
                              "Add sub-items to describe each role's responsibilities."
                
                Case LCase(SEC_OBJECTIVES):
                    helpText = "Objectives Section" & vbCrLf & vbCrLf & _
                              "List the desired outcomes of this procedure." & vbCrLf & _
                              "What should be accomplished?"
                
                Case LCase(SEC_STEPS):
                    helpText = "Process Steps Section" & vbCrLf & vbCrLf & _
                              "Enter step-by-step instructions." & vbCrLf & _
                              "Parent item = major step" & vbCrLf & _
                              "Sub-items = details, alternatives, decision points"
                
                Case LCase(SEC_KPIS):
                    helpText = "Key Performance Indicators" & vbCrLf & vbCrLf & _
                              "List metrics used to measure process effectiveness." & vbCrLf & _
                              "What quantifies success?"
                
                Case LCase(SEC_RESOURCES):
                    helpText = "Resources Section" & vbCrLf & vbCrLf & _
                              "List tools, templates, forms, and references needed." & vbCrLf & _
                              "Each row: Resource Name | Type | Location"
                
                Case Else:
                    helpText = "Built-in Section" & vbCrLf & vbCrLf & _
                              "This is a default SOD section."
            End Select
        Else
            ' Custom section generic instructions
            helpText = "Custom Section: " & secName & vbCrLf & vbCrLf & _
                      "Add and manage entries:" & vbCrLf & vbCrLf & _
                      "1. Type in the text field" & vbCrLf & _
                      "2. Click '+ Add' to save" & vbCrLf & _
                      "3. Select and click 'Edit' to modify" & vbCrLf & _
                      "4. Click '– Remove' to delete" & vbCrLf & _
                      "5. Use Move Up/Down to reorder"
        End If
    Else
        helpText = "Help" & vbCrLf & vbCrLf & _
                  "Select a section to view specific instructions."
    End If
    
    MsgBox helpText, vbInformation, "Help: " & mSecNames(mCurrentSection)
    Exit Sub
    
ErrHandler:
    HandleFormError "lblSectionHelp_Click"
    
End Sub

Private Sub lblSectionTitle_Click()
    
End Sub

Private Sub DebugLog(ByVal message As String)

    If mDeveloperMode Then
        Debug.Print "[DEV] " & Now & ": " & message
    End If
    
End Sub

Private Sub lblUpdate_Click()
'
'    Dim uDate As String
'
'    uDate = "Title Section" & vbCrLf & vbCrLf & _
'                              "Enter the title/name of the procedure or process." & vbCrLf & _
'                              "This should be a clear, concise summary."

End Sub

'====================================================
' FORM LIFECYCLE
'====================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler

    TABLE_SEP = Chr(31)
    LIST_SEP = Chr(30)
    SUBSUB_SEP = Chr(29)
    SUBSUB_ITEM_SEP = Chr(28)

    mDrilledSubIdx = -1
    mEditingSubSubIdx = -1
    
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

Private Sub UpdateRowButtonsBasedOnContent()
    On Error GoTo ErrHandler
    
    Dim isTableLike As Boolean
    isTableLike = (mSecTypes(mCurrentSection) = TYPE_TABLE Or _
                   mSecTypes(mCurrentSection) = TYPE_RESOURCES Or _
                   mSecTypes(mCurrentSection) = TYPE_DICTIONARY)

    If Not isTableLike Then Exit Sub
    If mEditingRowIdx > 0 Then Exit Sub
    
    ' Since dynamic textbox Change events don't bubble to the frame,
    ' just always enable Add for table-like sections - validation happens on click instead
    btnAddRow.Enabled = True
    Exit Sub
    
ErrHandler:
    HandleFormError "UpdateRowButtonsBasedOnContent"
    
End Sub

Private Sub UpdateSingleRowDisplay(ByVal rowNum As Long, ByVal rowStr As String, ByVal isHeaderRow As Boolean)
    ' Updates just one row's cells in lstRows, without reassigning
    ' the whole .List array - avoids the ColumnCount/.List-triggered
    ' auto-resize quirk entirely, since we're not touching .List at all
    
    On Error GoTo ErrHandler

    Dim parts() As String
    parts = Split(rowStr, TABLE_SEP)

    Dim colCount As Long
    colCount = lstRows.ColumnCount
Debug.Print "UpdateSingleRowDisplay: rowNum=" & rowNum & " colCount=" & colCount & " rowStr=" & rowStr
    Dim j As Long
    For j = 0 To colCount - 1
        Dim val As String
        If j <= UBound(parts) Then
            val = Trim(parts(j))
        Else
            val = ""
        End If

        If isHeaderRow Then val = "• " & val

        ' lstRows.List(rowIndex, columnIndex) - rowIndex is 0-based here
        lstRows.List(rowNum - 1, j) = val
    Next j

    Exit Sub

ErrHandler:
    HandleFormError "UpdateSingleRowDisplay"

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
    ReDim mSecHasSubSubItems(1 To mSecCount)

    mSecNames(1) = SEC_TITLE
    mSecNames(2) = SEC_GROUP
    mSecNames(3) = SEC_PURPOSE
    mSecNames(4) = SEC_SCOPE
    mSecNames(5) = SEC_DICTIONARY
    mSecNames(6) = SEC_OBJECTIVES
    mSecNames(7) = SEC_ROLES
    mSecNames(8) = SEC_STEPS
    mSecNames(9) = SEC_KPIS
    mSecNames(10) = SEC_RESOURCES

    mSecTypes(1) = TYPE_PLAIN
    mSecTypes(2) = TYPE_GROUP
    mSecTypes(3) = TYPE_PLAIN
    mSecTypes(4) = TYPE_PLAIN
    mSecTypes(5) = TYPE_DICTIONARY
    mSecTypes(6) = TYPE_LIST
    mSecTypes(7) = TYPE_LIST
    mSecTypes(8) = TYPE_LIST
    mSecTypes(9) = TYPE_LIST
    mSecTypes(10) = TYPE_RESOURCES

    Dim i As Long
    For i = 1 To mSecCount
        mSecData(i) = ""
        mSecCols(i) = ""
        mSecHasHeader(i) = False
        mSecHasSubItems(i) = IsBuiltInSubItemSection(mSecNames(i))
        mSecHasSubSubItems(i) = False          ' No built-in section defaults to 3 levels
        Set mSecItems(i) = New Collection
    Next i

    Exit Sub

ErrHandler:
    HandleFormError "InitSections"

End Sub

Private Function BuildUnconfirmedPromptText() As String
    On Error GoTo ErrHandler

    If mEditingItemIdx > 0 Then
        BuildUnconfirmedPromptText = "You're still editing an item in """ & mSecNames(mCurrentSection) & """." & vbCrLf & vbCrLf & _
            "Do you want to confirm your changes? Otherwise they will be discarded."
        Exit Function
    End If

    If mEditingSubIdx >= 0 Then
        BuildUnconfirmedPromptText = "You're still editing a sub-item in """ & mSecNames(mCurrentSection) & """." & vbCrLf & vbCrLf & _
            "Do you want to confirm your changes? Otherwise they will be discarded."
        Exit Function
    End If

    If mEditingSubSubIdx >= 0 Then
        BuildUnconfirmedPromptText = "You're still editing a sub-bullet in """ & mSecNames(mCurrentSection) & """." & vbCrLf & vbCrLf & _
            "Do you want to confirm your changes? Otherwise they will be discarded."
        Exit Function
    End If

    If mEditingRowIdx > 0 Then
        BuildUnconfirmedPromptText = "You're still editing a row in """ & mSecNames(mCurrentSection) & """." & vbCrLf & vbCrLf & _
            "Do you want to confirm your changes? Otherwise they will be discarded."
        Exit Function
    End If

    Select Case mSecTypes(mCurrentSection)

        Case TYPE_LIST
            If mDrilledSubIdx >= 0 Then
                If Len(Trim(txtSubItem.Text)) > 0 Then
                    BuildUnconfirmedPromptText = "Do you want to add """ & Trim(txtSubItem.Text) & """ to """ & TruncateForDisplay(mDrilledSubText) & """?" & vbCrLf & vbCrLf & _
                        "Otherwise it will be discarded."
                    Exit Function
                End If
            Else
                If Len(Trim(txtItem.Text)) > 0 Then
                    BuildUnconfirmedPromptText = "Do you want to add """ & Trim(txtItem.Text) & """ to " & mSecNames(mCurrentSection) & "?" & vbCrLf & vbCrLf & _
                        "Otherwise it will be discarded."
                    Exit Function

                ElseIf mSecHasSubItems(mCurrentSection) And Len(Trim(txtSubItem.Text)) > 0 Then
                    Dim parentDisplay As String
                    If lstItems.ListIndex >= 0 Then
                        parentDisplay = lstItems.List(lstItems.ListIndex)
                    Else
                        parentDisplay = "(no item selected)"
                    End If
                    BuildUnconfirmedPromptText = "Do you want to add """ & Trim(txtSubItem.Text) & """ to """ & parentDisplay & """?" & vbCrLf & vbCrLf & _
                        "Otherwise it will be discarded."
                    Exit Function
                End If
            End If

        Case TYPE_TABLE, TYPE_RESOURCES, TYPE_DICTIONARY
            BuildUnconfirmedPromptText = "Do you want to add this row to """ & mSecNames(mCurrentSection) & """?" & vbCrLf & vbCrLf & _
                "Otherwise it will be discarded."
            Exit Function

    End Select

    ' Fallback - shouldn't normally be reached if HasUnconfirmedContent() said True
    BuildUnconfirmedPromptText = "You have unconfirmed changes in this section." & vbCrLf & vbCrLf & _
        "Do you want to confirm and add them before continuing? Otherwise they will be discarded."

    Exit Function

ErrHandler:
    HandleFormError "BuildUnconfirmedPromptText"
    BuildUnconfirmedPromptText = "You have unconfirmed changes. Confirm before continuing?"

End Function

Private Function GetSectionHelpText(ByVal secName As String, ByVal secType As String) As String

    Select Case LCase(Trim(secName))
        Case LCase(SEC_TITLE)
            GetSectionHelpText = "The name of this Standard Operating Document. Keep it short and specific."

        Case LCase(SEC_GROUP)
            GetSectionHelpText = "Which department or team this SOD belongs to."

        Case LCase(SEC_PURPOSE)
            GetSectionHelpText = "One or two sentences explaining why this process exists."

        Case LCase(SEC_SCOPE)
            GetSectionHelpText = "What this process covers - and, if helpful, what it does NOT cover."

        Case LCase(SEC_DICTIONARY)
            GetSectionHelpText = "Define any terms or abbreviations someone unfamiliar with this process might not know."

        Case LCase(SEC_ROLES)
            GetSectionHelpText = "List each role involved (e.g. 'Shop Assistant'), then add their duties underneath as sub-items."

        Case LCase(SEC_OBJECTIVES)
            GetSectionHelpText = "The measurable goals this process is meant to achieve."

        Case LCase(SEC_STEPS)
            GetSectionHelpText = "The steps someone would follow, in order. Add details for each step as sub-items underneath it."

        Case LCase(SEC_KPIS)
            GetSectionHelpText = "How success for this process is measured."

        Case LCase(SEC_RESOURCES)
            GetSectionHelpText = "Links, documents, or tools someone would need to complete this process."

        Case Else
            ' User-created sections get generic help based on their type
            Select Case secType
                Case TYPE_PLAIN
                    GetSectionHelpText = "Enter a paragraph of text for this section."
                    
                Case TYPE_LIST
                    GetSectionHelpText = "Add one item at a time. Check ""List has sub-items"" if each item needs its own indented details underneath."
                
                Case TYPE_TABLE
                    GetSectionHelpText = "Set how many columns this table needs, then add rows. Check ""First row is headers"" if the first row should be labels rather than data."
                
                Case Else
                    GetSectionHelpText = ""
            End Select
    End Select

End Function

Private Function IsReservedSectionName(ByVal name As String) As Boolean
    ' Check if this name matches a built-in section or reserved feature
    Select Case LCase(Trim(name))
        Case LCase(SEC_TITLE), LCase(SEC_GROUP), LCase(SEC_PURPOSE), LCase(SEC_SCOPE), _
             LCase(SEC_DICTIONARY), LCase(SEC_ROLES), LCase(SEC_OBJECTIVES), _
             LCase(SEC_STEPS), LCase(SEC_KPIS), LCase(SEC_RESOURCES), "developer mode"
            IsReservedSectionName = True
    End Select
    
End Function

Private Function GetBuiltInType(ByVal name As String) As String
    ' Return the TYPE_ for a built-in section name
    Select Case LCase(Trim(name))
        Case LCase(SEC_TITLE):      GetBuiltInType = TYPE_PLAIN
        Case LCase(SEC_GROUP):      GetBuiltInType = TYPE_GROUP
        Case LCase(SEC_PURPOSE):    GetBuiltInType = TYPE_PLAIN
        Case LCase(SEC_SCOPE):      GetBuiltInType = TYPE_PLAIN
        Case LCase(SEC_DICTIONARY): GetBuiltInType = TYPE_DICTIONARY
        Case LCase(SEC_ROLES):      GetBuiltInType = TYPE_LIST
        Case LCase(SEC_OBJECTIVES): GetBuiltInType = TYPE_LIST
        Case LCase(SEC_STEPS):      GetBuiltInType = TYPE_LIST
        Case LCase(SEC_KPIS):       GetBuiltInType = TYPE_LIST
        Case LCase(SEC_RESOURCES):  GetBuiltInType = TYPE_RESOURCES
        Case Else:                  GetBuiltInType = TYPE_PLAIN
    End Select
    
End Function

Private Function GetBuiltInHasSubItems(ByVal name As String) As Boolean
    ' Return whether a built-in section has sub-items
    Select Case LCase(Trim(name))
        Case LCase(SEC_ROLES), LCase(SEC_STEPS), LCase(SEC_RESOURCES)
            GetBuiltInHasSubItems = True
    End Select
    
End Function

Private Sub UpdateFormCaption()
    On Error GoTo ErrHandler
    
    Dim titleSecIdx As Long
    Dim titleText As String
    titleSecIdx = FindSection(SEC_TITLE)
    If titleSecIdx > 0 Then
        titleText = Trim(mSecData(titleSecIdx))
    End If
    If titleText <> "" Then
        Me.Caption = "SOD Editor: " & titleText
    Else
        Me.Caption = "SOD Editor"
    End If

    If mDeveloperMode Then
        Me.Caption = Me.Caption & " [Developer Mode]"
    End If

    Exit Sub
ErrHandler:
    HandleFormError "UpdateFormCaption"
    
End Sub

Private Sub PopulateGroupDropdown()
    On Error GoTo ErrHandler

    cboGroup.Clear
    ' TODO: replace with your real Group options
    cboGroup.AddItem "Project Controls & Processes"
    cboGroup.AddItem "Safety"
    cboGroup.AddItem "Estimating"
    
    Dim groupIdx As Long
    groupIdx = FindSection(SEC_GROUP)
    If groupIdx > 0 Then
        cboGroup.Value = mSecData(groupIdx)
    End If

    Exit Sub

ErrHandler:
    HandleFormError "PopulateGroupDropdown"

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
    mUndoHasSubItems = mSecHasSubItems(mCurrentSection)
    mUndoHasSubSubItems = mSecHasSubSubItems(mCurrentSection)
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
    mSecHasSubItems(mCurrentSection) = mUndoHasSubItems
    mSecHasSubSubItems(mCurrentSection) = mUndoHasSubSubItems
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

'====================================================
' Update button states during edit mode based on item position
'====================================================

Private Sub UpdateItemReorderButtonStates()
    On Error GoTo ErrHandler
    
    btnAddItem.Enabled = (Len(Trim(txtItem.Text)) > 0)
    
    ' During edit mode:
    ' - btnAddItem is labeled "Move Up" ? disable if first item
    ' - btnRemoveItem is labeled "Move Down" ? disable if last item
    
    If mEditingItemIdx <= 0 Then
        ' Not in edit mode - shouldn't happen, but be safe
        btnAddItem.Enabled = True
        btnRemoveItem.Enabled = True
        Exit Sub
    End If
    
    Dim itemCount As Long
    itemCount = mSecItems(mCurrentSection).count
    
    ' Disable "Move Up" if on first item
    btnAddItem.Enabled = (mEditingItemIdx > 1)
    
    ' Disable "Move Down" if on last item
    btnRemoveItem.Enabled = (mEditingItemIdx < itemCount)

    Exit Sub

ErrHandler:
    HandleFormError "UpdateItemReorderButtonStates"

End Sub

Private Sub UpdateSubItemReorderButtonStates()
    On Error GoTo ErrHandler
    
    If mEditingSubIdx < 0 Then
        btnAddSubItem.Enabled = True
        btnRemoveSubItem.Enabled = True
        Exit Sub
    End If
    
    Dim parts() As String
    parts = Split(mSecItems(mCurrentSection)(lstItems.ListIndex + 1), TABLE_SEP)
    Dim subs() As String
    
    If UBound(parts) < 1 Then
        btnAddSubItem.Enabled = False
        btnRemoveSubItem.Enabled = False
        Exit Sub
    End If
    
    subs = Split(parts(1), LIST_SEP)
    
    ' Disable "Move Up" if first sub-item
    btnAddSubItem.Enabled = (mEditingSubIdx > 0)
    
    ' Disable "Move Down" if last sub-item
    btnRemoveSubItem.Enabled = (mEditingSubIdx < UBound(subs))

    Exit Sub

ErrHandler:
    HandleFormError "UpdateSubItemReorderButtonStates"

End Sub

Private Sub UpdateRowReorderButtonStates()
    On Error GoTo ErrHandler
    
    If mEditingRowIdx <= 0 Then
        btnAddRow.Enabled = True
        btnRemoveRow.Enabled = True
        Exit Sub
    End If
    
    Dim rowCount As Long
    rowCount = mSecItems(mCurrentSection).count
    
    ' Disable "Move Up" if on first row, OR moving into a locked header
    Dim cannotMoveUp As Boolean
    cannotMoveUp = (mEditingRowIdx <= 1)
    If mEditingRowIdx = 2 And mSecHasHeader(mCurrentSection) Then cannotMoveUp = True
    
    btnAddRow.Enabled = Not cannotMoveUp
    
    ' Disable "Move Down" if on last row, OR moving header down
    Dim cannotMoveDown As Boolean
    cannotMoveDown = (mEditingRowIdx >= rowCount)
    If mEditingRowIdx = 1 And mSecHasHeader(mCurrentSection) Then cannotMoveDown = True
    
    btnRemoveRow.Enabled = Not cannotMoveDown
    Exit Sub

ErrHandler:
    HandleFormError "UpdateRowReorderButtonStates"

End Sub

'====================================================
' HELPER: Update button states based on focus and edit mode
'====================================================

Private Sub UpdateItemButtonsBasedOnFocus()
    On Error GoTo ErrHandler
    
    ' BLOCK: If SubItem is being edited, disable all item buttons
    If mEditingSubIdx >= 0 Then
        btnAddItem.Enabled = False
        btnEditItem.Enabled = False
        btnRemoveItem.Enabled = False
        Exit Sub
    End If
    
    ' EDIT MODE: Position-based enabling
    If mEditingItemIdx > 0 Then
        btnAddItem.Enabled = (mEditingItemIdx > 1)
        btnEditItem.Enabled = True
        btnRemoveItem.Enabled = (mEditingItemIdx < mSecItems(mCurrentSection).count)
        Exit Sub
    End If
    
    ' NORMAL MODE: Check if txtItem has content (not just focus)
    ' If it has content, "+ Add" is enabled
    btnAddItem.Enabled = (Len(Trim(txtItem.Text)) > 0)
    
    ' "Edit" and "– Remove" enabled only if lstItems has selection
    Dim hasSelection As Boolean
    hasSelection = (lstItems.ListIndex >= 0)
    btnEditItem.Enabled = hasSelection
    btnRemoveItem.Enabled = hasSelection
    Exit Sub

ErrHandler:
    HandleFormError "UpdateItemButtonsBasedOnFocus"

End Sub

Private Sub UpdateSubItemButtonsBasedOnFocus()
    On Error GoTo ErrHandler

    ' Drilled-in mode takes priority and short-circuits into
    ' the sub-sub-item logic, which mirrors the block below almost exactly
    If mDrilledSubIdx >= 0 Then
        UpdateSubSubItemButtonsBasedOnFocus   ' sibling function, same pattern
        Exit Sub
    End If

    ' ... existing body, completely unchanged below this point ...
    If mEditingItemIdx > 0 Then
        btnAddSubItem.Enabled = False
        btnEditSubItem.Enabled = False
        btnRemoveSubItem.Enabled = False
        Exit Sub
    End If

    If lstItems.ListIndex < 0 Then
        btnAddSubItem.Enabled = False
        btnEditSubItem.Enabled = False
        btnRemoveSubItem.Enabled = False
        Exit Sub
    End If

    If mEditingSubIdx >= 0 Then
        Dim parentIdx As Long
        parentIdx = lstItems.ListIndex + 1
        Dim parts() As String
        parts = Split(mSecItems(mCurrentSection)(parentIdx), TABLE_SEP)

        If UBound(parts) < 1 Then
            btnAddSubItem.Enabled = False
            btnEditSubItem.Enabled = False
            btnRemoveSubItem.Enabled = False
            btnDrillIntoSub.Enabled = False
            Exit Sub
        End If

        Dim subs() As String
        subs = Split(parts(1), LIST_SEP)
        btnAddSubItem.Enabled = (mEditingSubIdx > 0)
        btnEditSubItem.Enabled = True
        btnRemoveSubItem.Enabled = (mEditingSubIdx < UBound(subs))
        Exit Sub
    End If

    ' NORMAL MODE: Check if txtSubItem has content
    btnAddSubItem.Enabled = (Len(Trim(txtSubItem.Text)) > 0)
    Dim hasSelection As Boolean
    hasSelection = (lstSubItems.ListIndex >= 0)
    btnEditSubItem.Enabled = hasSelection
    btnRemoveSubItem.Enabled = hasSelection
    btnDrillIntoSub.Enabled = hasSelection And chkSubSubItems.Value
    Exit Sub

ErrHandler:
    HandleFormError "UpdateSubItemButtonsBasedOnFocus"

End Sub

Private Sub UpdateSubSubItemButtonsBasedOnFocus()
    On Error GoTo ErrHandler

    If lstItems.ListIndex < 0 Then
        btnAddSubItem.Enabled = False
        btnEditSubItem.Enabled = False
        btnRemoveSubItem.Enabled = False
        Exit Sub
    End If

    Dim parentIdx As Long
    parentIdx = lstItems.ListIndex + 1
    Dim parts() As String
    parts = Split(mSecItems(mCurrentSection)(parentIdx), TABLE_SEP)

    Dim subs() As String
    subs = Split(parts(1), LIST_SEP)

    Dim subParts() As String
    subParts = Split(subs(mDrilledSubIdx), SUBSUB_SEP)

    If mEditingSubSubIdx >= 0 Then
        If UBound(subParts) < 1 Then
            btnAddSubItem.Enabled = False
            btnEditSubItem.Enabled = False
            btnRemoveSubItem.Enabled = False
            Exit Sub
        End If

        Dim subsubs() As String
        subsubs = Split(subParts(1), SUBSUB_ITEM_SEP)   ' CHANGED
        btnAddSubItem.Enabled = (mEditingSubSubIdx > 0)
        btnEditSubItem.Enabled = True
        btnRemoveSubItem.Enabled = (mEditingSubSubIdx < UBound(subsubs))
        Exit Sub
    End If

    btnAddSubItem.Enabled = (Len(Trim(txtSubItem.Text)) > 0)
    Dim hasSelection As Boolean
    hasSelection = (lstSubItems.ListIndex >= 0)
    btnEditSubItem.Enabled = hasSelection
    btnRemoveSubItem.Enabled = hasSelection
    Exit Sub

ErrHandler:
    HandleFormError "UpdateSubSubItemButtonsBasedOnFocus"

End Sub

Private Function ConfirmPendingChanges() As Boolean
    On Error GoTo ErrHandler

    If mEditingItemIdx > 0 Then
        Call btnEditItem_Click       ' runs its CONFIRM branch
        ConfirmPendingChanges = (mEditingItemIdx = 0)
        Exit Function
    End If

    If mEditingSubIdx >= 0 Then
        Call btnEditSubItem_Click    ' runs its CONFIRM branch
        ConfirmPendingChanges = (mEditingSubIdx = -1)
        Exit Function
    End If

    If mEditingSubSubIdx >= 0 Then
        Call btnEditSubItem_Click    ' mDrilledSubIdx >= 0 routes this into EditSubSubItem's CONFIRM branch
        ConfirmPendingChanges = (mEditingSubSubIdx = -1)
        Exit Function
    End If

    If mEditingRowIdx > 0 Then
        Call btnEditRow_Click        ' runs its CONFIRM branch
        ConfirmPendingChanges = (mEditingRowIdx = 0)
        Exit Function
    End If

    Select Case mSecTypes(mCurrentSection)
        Case TYPE_LIST
            If mDrilledSubIdx >= 0 Then
                If Len(Trim(txtSubItem.Text)) > 0 Then
                    Call btnAddSubItem_Click     ' assumes this routes into AddSubSubItem via a top-of-function mDrilledSubIdx gate, same pattern as btnEditSubItem_Click
                    ConfirmPendingChanges = (Len(Trim(txtSubItem.Text)) = 0)
                    Exit Function
                End If
            Else
                If Len(Trim(txtItem.Text)) > 0 Or _
                   (mSecHasSubItems(mCurrentSection) And Len(Trim(txtSubItem.Text)) > 0) Then
                    Call btnAddItem_Click
                    Dim itemAdded As Boolean
                    itemAdded = (Len(Trim(txtItem.Text)) = 0)
                    ConfirmPendingChanges = itemAdded
                    Exit Function
                End If
            End If

        Case TYPE_TABLE, TYPE_RESOURCES, TYPE_DICTIONARY
            If Replace(ReadTableRowFields(), TABLE_SEP, "") <> "" Then
                Call btnAddRow_Click     ' runs its "add new row" branch
                ConfirmPendingChanges = (Replace(ReadTableRowFields(), TABLE_SEP, "") = "")
                Exit Function
            End If
    End Select

    ConfirmPendingChanges = True   ' nothing was actually pending
    Exit Function

ErrHandler:
    HandleFormError "ConfirmPendingChanges"
    ConfirmPendingChanges = False

End Function

Private Sub lstSections_Click()
    On Error GoTo ErrHandler

    If mLoading Then Exit Sub
    
    If lstSections.ListIndex < 0 Then Exit Sub
    
    If mEditingSectionIdx > 0 Then Exit Sub
    
    If lstSections.ListIndex = mCurrentSection - 1 Then Exit Sub

    If Not mDeveloperMode And HasUnconfirmedContent() Then

        Dim resp As VbMsgBoxResult
        resp = MsgBox(BuildUnconfirmedPromptText(), vbYesNoCancel + vbQuestion, "Unsaved Text")

        Select Case resp
            Case vbCancel
                mLoading = True
                lstSections.ListIndex = mCurrentSection - 1
                ' Deliberately do NOT set mLoading = False here -
                ' MSForms queues a Click event for the snap-back and
                ' releases it the moment mLoading goes False. Instead
                ' we let the sub exit with mLoading = True and reset
                ' it in a DoEvents+timer pattern via a helper sub.
                DoEvents
                mLoading = False
                Exit Sub

            Case vbYes
                If Not ConfirmPendingChanges() Then
                    mLoading = True
                    lstSections.ListIndex = mCurrentSection - 1
                    DoEvents
                    mLoading = False
                    Exit Sub
                End If

            Case vbNo
                ' fall through
        End Select

    End If

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
    
    If mDeveloperMode Then
        btnRemoveSection.Enabled = True
        btnEditSection.Enabled = True
    Else
        ' Process Steps can be deleted, so it's always removable
        ' Process Steps and Dictionary can be deleted
        Dim isDeletable As Boolean
        isDeletable = (LCase(Trim(mSecNames(secIdx))) = LCase(SEC_STEPS) Or _
                       LCase(Trim(mSecNames(secIdx))) = LCase(SEC_DICTIONARY))
        
        btnRemoveSection.Enabled = (Not isBuiltIn Or isDeletable)
        btnEditSection.Enabled = Not isBuiltIn
    End If
    
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

Private Function HasUnconfirmedContent() As Boolean
    On Error GoTo ErrHandler

    If mEditingItemIdx > 0 Then
        HasUnconfirmedContent = True
        Exit Function
    End If

    If mEditingSubIdx >= 0 Then
        HasUnconfirmedContent = True
        Exit Function
    End If

    If mEditingSubSubIdx >= 0 Then
        HasUnconfirmedContent = True
        Exit Function
    End If

    If mEditingRowIdx > 0 Then
        HasUnconfirmedContent = True
        Exit Function
    End If

    Select Case mSecTypes(mCurrentSection)
        Case TYPE_LIST
            If mDrilledSubIdx >= 0 Then
                ' Drilled into a sub-item: txtItem is disabled/inert and
                ' must NOT be considered. txtSubItem now holds a pending
                ' NEW sub-sub-item (the existing-edit case is already
                ' caught above via mEditingSubSubIdx).
                If Len(Trim(txtSubItem.Text)) > 0 Then
                    HasUnconfirmedContent = True
                    Exit Function
                End If
            Else
                If Len(Trim(txtItem.Text)) > 0 Then
                    HasUnconfirmedContent = True
                    Exit Function
                End If
                If mSecHasSubItems(mCurrentSection) And Len(Trim(txtSubItem.Text)) > 0 Then
                    HasUnconfirmedContent = True
                    Exit Function
                End If
            End If

        Case TYPE_TABLE, TYPE_RESOURCES, TYPE_DICTIONARY
            If Replace(ReadTableRowFields(), TABLE_SEP, "") <> "" Then
                HasUnconfirmedContent = True
                Exit Function
            End If
    End Select

    HasUnconfirmedContent = False
    Exit Function

ErrHandler:
    HandleFormError "HasUnconfirmedContent"
    HasUnconfirmedContent = False

End Function

'====================================================
' REORDERING SECTIONS
'====================================================

Private Sub SwapSections(ByVal idx1 As Long, ByVal idx2 As Long)

    Dim tmpName As String, tmpType As String, tmpData As String, tmpCols As String
    Dim tmpHasHeader As Boolean, tmpHasSubItems As Boolean, tmpHasSubSubItems As Boolean
    Dim tmpItems As Collection

    tmpName = mSecNames(idx1): mSecNames(idx1) = mSecNames(idx2): mSecNames(idx2) = tmpName
    tmpType = mSecTypes(idx1): mSecTypes(idx1) = mSecTypes(idx2): mSecTypes(idx2) = tmpType
    tmpData = mSecData(idx1): mSecData(idx1) = mSecData(idx2): mSecData(idx2) = tmpData
    tmpCols = mSecCols(idx1): mSecCols(idx1) = mSecCols(idx2): mSecCols(idx2) = tmpCols
    tmpHasHeader = mSecHasHeader(idx1): mSecHasHeader(idx1) = mSecHasHeader(idx2): mSecHasHeader(idx2) = tmpHasHeader
    tmpHasSubItems = mSecHasSubItems(idx1): mSecHasSubItems(idx1) = mSecHasSubItems(idx2): mSecHasSubItems(idx2) = tmpHasSubItems
    tmpHasSubSubItems = mSecHasSubSubItems(idx1): mSecHasSubSubItems(idx1) = mSecHasSubSubItems(idx2): mSecHasSubSubItems(idx2) = tmpHasSubSubItems

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
        mSecHasSubSubItems(i) = mSecHasSubSubItems(i + 1)
        Set mSecItems(i) = mSecItems(i + 1)
    Next i

    mSecCount = mSecCount - 1
    ReDim Preserve mSecNames(1 To mSecCount)
    ReDim Preserve mSecTypes(1 To mSecCount)
    ReDim Preserve mSecData(1 To mSecCount)
    ReDim Preserve mSecCols(1 To mSecCount)
    ReDim Preserve mSecHasHeader(1 To mSecCount)
    ReDim Preserve mSecHasSubItems(1 To mSecCount)
    ReDim Preserve mSecHasSubSubItems(1 To mSecCount)
    ReDim Preserve mSecItems(1 To mSecCount)

End Sub

'====================================================
' ADD / EDIT / REMOVE SECTION
' Uses the same Add->Move Up / Edit->Confirm / Remove->Move Down
' pattern as items, sub-items, and rows elsewhere in the form.
' While naming a NEW section, the three buttons temporarily become
' Paragraph Section / List Section / Table Section instead.
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
        
        mLoading = True
        lstSections.ListIndex = lstSections.ListIndex - 1
        mLoading = False
        
        mCurrentSection = mEditingSectionIdx
        ShowSection mCurrentSection
        Exit Sub
    End If
    
    ' Enter "naming a new section" mode
    mAddingSection = True
    txtSectionName.Text = ""
    txtSectionName.Visible = True
    txtSectionName.SetFocus
    
    btnAddSection.Caption = "Paragraph Section"
    btnEditSection.Caption = "List Section"
    btnRemoveSection.Caption = "Table Section"
    btnCancel.Caption = "Cancel"
    
    btnAddSection.Enabled = True
    btnEditSection.Enabled = True
    btnRemoveSection.Enabled = True
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
        PushUndo
        mSecNames(mEditingSectionIdx) = newName
        RefreshSectionList
        Dim savedIdx As Long
        savedIdx = mEditingSectionIdx
        mEditingSectionIdx = 0
        btnEditSection.Caption = "Edit Section"
        btnAddSection.Caption = "+ Add Section"
        btnRemoveSection.Caption = "– Remove Section"
        btnCancel.Caption = "Close Editor"
        txtSectionName.Visible = False
        lstSections.SpecialEffect = 3   ' Etched (back to normal)
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
        
        btnAddSection.Enabled = True
        btnEditSection.Enabled = True
        btnRemoveSection.Enabled = True

        lstSections.SpecialEffect = 2   ' Sunken (edit mode indicator)
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
        
        mLoading = True
        lstSections.ListIndex = lstSections.ListIndex + 1
        mLoading = False
        
        mCurrentSection = mEditingSectionIdx
        ShowSection mCurrentSection
        Exit Sub
    End If

    Dim selIdx As Long
    selIdx = lstSections.ListIndex

    If selIdx < 0 Then
        MsgBox "Select a section to remove.", vbExclamation
        Exit Sub
    End If

    Dim resp As VbMsgBoxResult
    Dim secIdx As Long
    secIdx = selIdx + 1

    If IsBuiltInSectionName(mSecNames(secIdx)) Then
    ' Special case: Process Steps and Dictionary CAN be deleted, but with confirmation
            If LCase(Trim(mSecNames(secIdx))) = LCase(SEC_STEPS) Or LCase(Trim(mSecNames(secIdx))) = LCase(SEC_DICTIONARY) Then
                resp = MsgBox("Delete '" & mSecNames(secIdx) & "'?" & vbCrLf & vbCrLf & _
                              "This action cannot be undone. Are you SURE?", _
                              vbYesNo + vbExclamation, "Delete " & mSecNames(secIdx) & "?")
                If resp = vbNo Then Exit Sub
                ' Allow it to continue (confirmation happens below)
            ElseIf Not mDeveloperMode Then
            ' Other built-in sections cannot be deleted (unless dev mode)
            MsgBox "'" & mSecNames(secIdx) & "' is a built-in section and cannot be removed.", vbExclamation
            Exit Sub
        End If
    End If

    If mSecCount <= 1 Then
        MsgBox "Cannot remove the only remaining section.", vbExclamation
        Exit Sub
    End If

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
    Dim existingIdx As Long  ' DECLARE ONCE HERE
    secName = Trim(txtSectionName.Text)
    If secName = "" Then
        MsgBox "Please type a section name first.", vbExclamation
        Exit Sub
    End If

    ' Cannot create a section with a reserved name
    If IsReservedSectionName(secName) Then
        ' Special case: can only recreate built-in sections, not "Developer Mode"
        If LCase(secName) = "developer mode" Then
            MsgBox "'" & secName & "' is reserved and cannot be used.", vbExclamation
            Exit Sub
        End If
        ' Built-in sections can be recreated if deleted
        existingIdx = FindSection(secName)
        If existingIdx > 0 Then
            MsgBox "A section named '" & secName & "' already exists.", vbExclamation
            Exit Sub
        End If
    Else
        ' Check if custom name already in use
        existingIdx = FindSection(secName)
        If existingIdx > 0 Then
            MsgBox "A section named '" & secName & "' already exists.", vbExclamation
            Exit Sub
        End If
    End If

    ' Determine type and position
    Dim finalType As String
    Dim isBuiltIn As Boolean
    Dim insertIdx As Long
    
    finalType = secType
    isBuiltIn = False
    insertIdx = mSecCount + 1  ' Default: append to end
    
    If IsReservedSectionName(secName) Then
        finalType = GetBuiltInType(secName)
        isBuiltIn = True
        insertIdx = GetCanonicalIndex(secName)  ' Use canonical position
    End If

    ' Insert the section
    If isBuiltIn Then
        InsertSectionAt insertIdx, secName, finalType
    Else
        ' Custom section: append to end
        mSecCount = mSecCount + 1
        ReDim Preserve mSecNames(1 To mSecCount)
        ReDim Preserve mSecTypes(1 To mSecCount)
        ReDim Preserve mSecData(1 To mSecCount)
        ReDim Preserve mSecItems(1 To mSecCount)
        ReDim Preserve mSecCols(1 To mSecCount)
        ReDim Preserve mSecHasHeader(1 To mSecCount)
        ReDim Preserve mSecHasSubItems(1 To mSecCount)
        ReDim Preserve mSecHasSubSubItems(1 To mSecCount)

        mSecNames(mSecCount) = secName
        mSecTypes(mSecCount) = finalType
        mSecData(mSecCount) = ""
        mSecCols(mSecCount) = ""
        mSecHasHeader(mSecCount) = False
        mSecHasSubItems(mSecCount) = False
        mSecHasSubSubItems(mSecCount) = False
        Set mSecItems(mSecCount) = New Collection
        
        insertIdx = mSecCount
    End If

    RefreshSectionList
    SaveCurrentSection

    mAddingSection = False
    txtSectionName.Visible = False
    btnAddSection.Caption = "+ Add Section"
    btnEditSection.Caption = "Edit Section"
    btnRemoveSection.Caption = "– Remove Section"
    btnCancel.Caption = "Close Editor"

    ' Select the recreated section
    lstSections.ListIndex = insertIdx - 1
    mCurrentSection = insertIdx
    ShowSection mCurrentSection

    ' Inform user if standard format was applied
    If isBuiltIn Then
        MsgBox "'" & secName & "' recreated with standard format applied." & vbCrLf & _
               "(Restored to original position.)", vbInformation
    End If

    Exit Sub

ErrHandler:
    HandleFormError "CreateNewSection"

End Sub

Private Sub txtSectionName_Change()
    On Error GoTo ErrHandler
    
    Dim name As String
    name = Trim(txtSectionName.Text)
    ' Easter egg: typing "Developer Mode" unhides the dev button
    If LCase(name) = "developer mode" Then
        btnDevMode.Visible = True
        ' Fall through to mark it as reserved (red - can't create)
    End If
    
    If mAddingSection And name <> "" Then
    
        Dim existingIdx As Long
        existingIdx = FindSection(name)
        ' Check if name already exists (built-in OR custom)
        If existingIdx > 0 Then
            ' Name conflict - can't use
            txtSectionName.BackColor = &HFF0000  ' Red: can't use
            lstSections.ListIndex = existingIdx - 1  ' Highlight the conflicting section
        ElseIf IsReservedSectionName(name) Then
            ' Built-in name or reserved - can't use
            txtSectionName.BackColor = &HFF0000  ' Red: can't use
        Else
            txtSectionName.BackColor = &H80000005   ' Normal
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    HandleFormError "txtSectionName_Change"
    
End Sub

Private Function GetCanonicalIndex(ByVal name As String) As Long
    ' Return the original built-in position for a section name
    Select Case LCase(Trim(name))
        Case LCase(SEC_TITLE):      GetCanonicalIndex = 1
        Case LCase(SEC_GROUP):      GetCanonicalIndex = 2
        Case LCase(SEC_PURPOSE):    GetCanonicalIndex = 3
        Case LCase(SEC_SCOPE):      GetCanonicalIndex = 4
        Case LCase(SEC_DICTIONARY): GetCanonicalIndex = 5
        Case LCase(SEC_OBJECTIVES): GetCanonicalIndex = 6  ' (swapped with Roles)
        Case LCase(SEC_ROLES):      GetCanonicalIndex = 7  ' (swapped with Objectives)
        Case LCase(SEC_STEPS):      GetCanonicalIndex = 8
        Case LCase(SEC_KPIS):       GetCanonicalIndex = 9
        Case LCase(SEC_RESOURCES):  GetCanonicalIndex = 10
        Case Else:                  GetCanonicalIndex = 0
    End Select
    
End Function

Private Sub InsertSectionAt(ByVal insertIdx As Long, ByVal secName As String, ByVal secType As String)
    ' Insert a section at a specific index, shifting others down
    On Error GoTo ErrHandler
    
    ' Expand arrays
    mSecCount = mSecCount + 1
    ReDim Preserve mSecNames(1 To mSecCount)
    ReDim Preserve mSecTypes(1 To mSecCount)
    ReDim Preserve mSecData(1 To mSecCount)
    ReDim Preserve mSecItems(1 To mSecCount)
    ReDim Preserve mSecCols(1 To mSecCount)
    ReDim Preserve mSecHasHeader(1 To mSecCount)
    ReDim Preserve mSecHasSubItems(1 To mSecCount)
    ReDim Preserve mSecHasSubSubItems(1 To mSecCount)
    
    ' Shift everything at insertIdx and beyond down one position
    Dim i As Long
    For i = mSecCount - 1 To insertIdx Step -1
        mSecNames(i + 1) = mSecNames(i)
        mSecTypes(i + 1) = mSecTypes(i)
        mSecData(i + 1) = mSecData(i)
        Set mSecItems(i + 1) = mSecItems(i)
        mSecCols(i + 1) = mSecCols(i)
        mSecHasHeader(i + 1) = mSecHasHeader(i)
        mSecHasSubItems(i + 1) = mSecHasSubItems(i)
        mSecHasSubSubItems(i + 1) = mSecHasSubSubItems(i)
    Next i
    
    ' Insert new section at insertIdx
    mSecNames(insertIdx) = secName
    mSecTypes(insertIdx) = secType
    mSecData(insertIdx) = ""
    mSecCols(insertIdx) = ""
    mSecHasHeader(insertIdx) = False
    mSecHasSubItems(insertIdx) = GetBuiltInHasSubItems(secName)
    mSecHasSubSubItems(insertIdx) = False   ' No built-in ever defaults to 3 levels
    Set mSecItems(insertIdx) = New Collection
    
    Exit Sub

ErrHandler:
    HandleFormError "InsertSectionAt"

End Sub

'====================================================
' SHOW SECTION - switches the right panel content
'====================================================

Private Sub ShowSection(ByVal idx As Long)
    On Error GoTo ErrHandler
    
    lstSections.ListIndex = idx - 1

    If idx < 1 Or idx > mSecCount Then Exit Sub

    ' Reset every edit/add mode and button caption before showing
    ' whatever section is being switched to - prevents a half-finished
    ' edit in one section from leaking into another
    mLoading = True
    mEditingItemIdx = 0
    mEditingSubIdx = -1
    mEditingRowIdx = 0
    mEditingSectionIdx = 0
    mAddingSection = False

    ' Sub-sub-item drill/edit state also needs a clean reset on every
    ' section switch. By this point, any unconfirmed sub-sub text has
    ' already been resolved by the caller's HasUnconfirmedContent /
    ' ConfirmPendingChanges guard (Yes = committed, No = user chose to
    ' discard, Cancel = we never reach ShowSection at all) - so this is
    ' a silent cleanup, not a second prompt.
    If mEditingSubSubIdx >= 0 Then
        CancelSubSubItemEditMode   ' reverts to snapshot, clears mEditingSubSubIdx and txtSubItem
    End If
    mDrilledSubIdx = -1
    mDrilledSubText = ""
    txtSubItem.Text = ""
    lstItems.Enabled = True
    txtItem.Enabled = True
    txtItem.BackColor = &H80000005

    btnAddItem.Caption = "+ Add"
    btnRemoveItem.Caption = "– Remove"
    btnEditItem.Caption = "Edit"

    btnAddSubItem.Caption = "+ Add"
    btnRemoveSubItem.Caption = "– Remove"
    btnEditSubItem.Caption = "Edit"

    btnAddRow.Caption = "+ Add"
    btnRemoveRow.Caption = "– Remove"
    btnEditRow.Caption = "Edit"

    btnEditSection.Caption = "Edit Section"
    btnAddSection.Caption = "+ Add Section"
    btnRemoveSection.Caption = "– Remove Section"
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
    lblSectionHelp.Caption = "s"
    lblSectionHelp.ControlTipText = GetSectionHelpText(mSecNames(idx), mSecTypes(idx))

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
    chkSubSubItems.Visible = False
    btnDrillIntoSub.Visible = False

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
    mCurrentSection = idx

    RefreshItemsDisplay

    If lstItems.ListCount > 0 Then
        lstItems.ListIndex = 0
    End If

    lstItems.Visible = True
    btnAddItem.Visible = True
    btnRemoveItem.Visible = True
    btnEditItem.Visible = True
    txtItem.Visible = True
    txtItem.Text = ""

    Dim isBuiltInWithSubs As Boolean
    Dim isBuiltInWithoutSubs As Boolean
    isBuiltInWithSubs = IsBuiltInSubItemSection(mSecNames(idx))
    isBuiltInWithoutSubs = IsBuiltInListSection(mSecNames(idx))

    chkSubItems.Visible = True
    mLoading = True
    chkSubItems.Value = mSecHasSubItems(idx)
    mLoading = False

    If mDeveloperMode Then
        chkSubItems.Enabled = True
        chkSubItems.ControlTipText = "DEV: Unlocked for testing"
    Else
        chkSubItems.Enabled = Not (isBuiltInWithSubs Or isBuiltInWithoutSubs)

        If Not chkSubItems.Enabled Then
            chkSubItems.ControlTipText = "Cannot change - this is a built-in section"
        Else
            chkSubItems.ControlTipText = "Check if items should have sub-items"
        End If
    End If

    Dim showSubs As Boolean
    showSubs = mSecHasSubItems(idx)
    lblSubItems.Visible = showSubs
    lstSubItems.Visible = showSubs
    btnAddSubItem.Visible = showSubs
    btnRemoveSubItem.Visible = showSubs
    btnEditSubItem.Visible = showSubs
    txtSubItem.Visible = showSubs

    ' chkSubSubItems visibility mirrors showSubs (no point showing
    ' "does a sub-item have children" when there are no sub-items at all).
    ' Enabled state additionally requires chkSubItems to actually be checked.
    chkSubSubItems.Visible = showSubs
    mLoading = True
    chkSubSubItems.Value = mSecHasSubSubItems(idx)
    mLoading = False
    chkSubSubItems.Enabled = showSubs
    If showSubs Then
        chkSubSubItems.ControlTipText = "Check if sub-items should have their own sub-bullets"
    End If

    If showSubs Then
        lstSubItems.Clear
        txtSubItem.Text = ""

        If lstItems.ListCount > 0 Then
            mLoading = True
            lstItems.ListIndex = 0
            mLoading = False
            UpdateSubItemsLabel
            UpdateSubItemControls
            RefreshSubItems
        Else
            UpdateSubItemsLabel
            UpdateSubItemControls
        End If
    Else
        lstSubItems.Clear
        txtSubItem.Enabled = False
        btnAddSubItem.Enabled = False
        btnEditSubItem.Enabled = False
        btnRemoveSubItem.Enabled = False
    End If

    ' Drill-toggle visibility - only ever shown alongside sub-items, and its Enabled state (gated on chkSubSubItems AND a selection) is handled inside UpdateSubItemButtonsBasedOnFocus, not here
    ' Visible only when sub-sub-items are actually enabled for this section
    btnDrillIntoSub.Visible = showSubs And mSecHasSubSubItems(idx)
    UpdateItemControls
    UpdateItemButtonsBasedOnFocus
    UpdateSubItemButtonsBasedOnFocus
    Exit Sub

ErrHandler:
    HandleFormError "ShowListSection"

End Sub

Private Sub chkSubItems_Click()
    On Error GoTo ErrHandler

    If mLoading Then Exit Sub
    
    If mSecTypes(mCurrentSection) <> TYPE_LIST Then Exit Sub

    ' Guard: don't let this checkbox blow past an active edit
    ' (item / sub-item / sub-sub-item / row) without confirming.
    If Not mDeveloperMode And HasUnconfirmedContent() Then
        Dim priorValue As Boolean
        priorValue = Not chkSubItems.Value

        Dim guardResp As VbMsgBoxResult
        guardResp = MsgBox(BuildUnconfirmedPromptText(), vbYesNoCancel + vbQuestion, "Unsaved Text")

        Select Case guardResp
            Case vbCancel
                mLoading = True
                chkSubItems.Value = priorValue
                DoEvents
                mLoading = False
                Exit Sub

            Case vbYes
                If Not ConfirmPendingChanges() Then
                    mLoading = True
                    chkSubItems.Value = priorValue
                    DoEvents
                    mLoading = False
                    Exit Sub
                End If

            Case vbNo
                ' fall through - discard and proceed with the toggle
        End Select
    End If

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
            If InStr(mSecItems(mCurrentSection)(k), TABLE_SEP) > 0 Then
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
            parts = Split(mSecItems(mCurrentSection)(j), TABLE_SEP)
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
    mainVal = Split(mSecItems(mCurrentSection)(itemIdx), TABLE_SEP)(0)
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
    If InStr(entry, TABLE_SEP) > 0 Then
    
        Dim parts() As String
        parts = Split(entry, TABLE_SEP)
        If UBound(parts) >= 1 Then
            If Trim(parts(1)) <> "" Then
            
                Dim subs() As String
                subs = Split(parts(1), LIST_SEP)
                Dim i As Long
                For i = 0 To UBound(subs)
                
                    Dim txt As String
                    txt = Trim(subs(i))
                    If txt <> "" Then
                        Dim displayTxt As String
                        If InStr(txt, SUBSUB_SEP) > 0 Then
                            displayTxt = "•  " & Split(txt, SUBSUB_SEP)(0)
                        Else
                            displayTxt = txt
                        End If
                        lstSubItems.AddItem TruncateForDisplay(displayTxt)
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

Private Sub RefreshSubSubItems()
    On Error GoTo ErrHandler

    lstSubItems.Clear   ' same physical control, now showing sub-sub-items

    If mDrilledSubIdx < 0 Then
        UpdateSubSubItemControls
        Exit Sub
    End If

    Dim parentIdx As Long
    parentIdx = lstItems.ListIndex + 1
    If lstItems.ListIndex < 0 Or parentIdx > mSecItems(mCurrentSection).count Then
        UpdateSubSubItemControls
        Exit Sub
    End If

    Dim parts() As String
    parts = Split(mSecItems(mCurrentSection)(parentIdx), TABLE_SEP)
    If UBound(parts) < 1 Then
        UpdateSubSubItemControls
        Exit Sub
    End If

    Dim subs() As String
    subs = Split(parts(1), LIST_SEP)
    If mDrilledSubIdx > UBound(subs) Then
        UpdateSubSubItemControls
        Exit Sub
    End If

    Dim entry As String
    entry = subs(mDrilledSubIdx)
    If InStr(entry, SUBSUB_SEP) > 0 Then

        Dim subParts() As String
        subParts = Split(entry, SUBSUB_SEP)
        If UBound(subParts) >= 1 Then
            If Trim(subParts(1)) <> "" Then

                Dim subsubs() As String
                subsubs = Split(subParts(1), SUBSUB_ITEM_SEP)
                Dim i As Long
                For i = 0 To UBound(subsubs)

                    Dim txt As String
                    txt = Trim(subsubs(i))
                    If txt <> "" Then
                        lstSubItems.AddItem TruncateForDisplay(txt)
                    End If
                Next i
            End If
        End If
    End If

    UpdateSubSubItemControls
    Exit Sub

ErrHandler:
    HandleFormError "RefreshSubSubItems"

End Sub

Private Sub RefreshItemsDisplay()
    On Error GoTo ErrHandler
    
    lstItems.Clear
    
    Dim i As Long
    For i = 1 To mSecItems(mCurrentSection).count
    
        Dim rawVal As String
        rawVal = mSecItems(mCurrentSection)(i)
        
        If mSecHasSubItems(mCurrentSection) Or InStr(rawVal, TABLE_SEP) > 0 Then
            lstItems.AddItem TruncateForDisplay(Split(rawVal, TABLE_SEP)(0))
        Else
            lstItems.AddItem TruncateForDisplay(rawVal)
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

    ' This function now only handles:
    ' - Textbox background color (visual state)
    ' - Disabling controls when SubItem is editing
    ' Button enabling/disabling is now handled by UpdateItemButtonsBasedOnFocus()

    If mEditingItemIdx > 0 Then Exit Sub   ' don't interfere with edit mode

    ' If currently editing a sub-item, lock out ALL item-level controls
    If mEditingSubIdx >= 0 Then
        txtItem.Enabled = False
        txtItem.BackColor = &H8000000F
        ' Don't touch button states here - let UpdateItemButtonsBasedOnFocus handle it
        Exit Sub
    End If

    ' Normal mode: textbox is enabled with normal color
    txtItem.Enabled = True
    txtItem.BackColor = &H80000005
    ' Don't touch button states here - let UpdateItemButtonsBasedOnFocus handle it
    Exit Sub

ErrHandler:
    HandleFormError "UpdateItemControls"

End Sub

Private Sub UpdateSubItemControls()
    On Error GoTo ErrHandler

    ' This function now only handles:
    ' - Textbox background color (visual state)
    ' - Disabling controls when Item is editing
    ' Button enabling/disabling is now handled by UpdateSubItemButtonsBasedOnFocus()

    If mEditingSubIdx >= 0 Then Exit Sub   ' don't interfere with edit mode

    ' If currently editing a parent item, lock out sub-item controls entirely
    If mEditingItemIdx > 0 Then
        txtSubItem.Enabled = False
        txtSubItem.BackColor = &H8000000F
        txtSubItem.ControlTipText = "Finish editing the item first"
        ' Don't touch button states here - let UpdateSubItemButtonsBasedOnFocus handle it
        Exit Sub
    End If

    ' Normal mode: check if parent is selected
    Dim hasParent As Boolean
    hasParent = (lstItems.ListIndex >= 0)

    txtSubItem.Enabled = hasParent
    ' Don't touch button states here - let UpdateSubItemButtonsBasedOnFocus handle it

    If Not hasParent Then
        txtSubItem.BackColor = &H8000000F
        txtSubItem.ControlTipText = "Add an item first"
    Else
        txtSubItem.BackColor = &H80000005
        txtSubItem.ControlTipText = ""
    End If
    
    Exit Sub

ErrHandler:
    HandleFormError "UpdateSubItemControls"

End Sub

Private Sub UpdateSubSubItemControls()
    On Error GoTo ErrHandler

    ' Same role as UpdateSubItemControls: textbox visual state only.
    ' Button enabling handled by UpdateSubSubItemButtonsBasedOnFocus()

    If mEditingSubSubIdx >= 0 Then Exit Sub   ' don't interfere with edit mode

    ' Normal drilled-in mode: txtSubItem is always enabled here, since you
    ' can't be drilled in without a valid parent sub-item already selected
    txtSubItem.Enabled = True
    txtSubItem.BackColor = &H80000005
    txtSubItem.ControlTipText = ""

    Exit Sub

ErrHandler:
    HandleFormError "UpdateSubSubItemControls"

End Sub

'====================================================
' LIST SELECTION EVENTS
'====================================================

Private Sub lstItems_Click()
    On Error GoTo ErrHandler
    Debug.Print "lstItems_Click fired. New ListIndex=" & lstItems.ListIndex
    
    If mLoading Then Exit Sub
    
    If mDrilledSubIdx >= 0 Then Exit Sub    ' Items list is locked while drilled in
    
    If lstItems.ListIndex < 0 Then Exit Sub
    
    ' If in Item Edit mode, switching items is allowed
    If mEditingItemIdx > 0 Then
    
        ' Capture the newly clicked index BEFORE any refresh
        Dim newIdx As Long
        newIdx = lstItems.ListIndex + 1
        ' Always save current text when switching items (no prompt)
        If Trim(txtItem.Text) <> "" Then
            SaveEditedItemTextInPlace
            mLoading = True
            RefreshItemsDisplay
            mLoading = False
        End If
        
        ' Switch to newly selected item
        mEditingItemIdx = newIdx
        ' Guard against out of bounds
        If mEditingItemIdx < 1 Or mEditingItemIdx > mSecItems(mCurrentSection).count Then
            mEditingItemIdx = 1
        End If
        
        ' Load new item text
        Dim newText As String
        If mSecHasSubItems(mCurrentSection) And InStr(mSecItems(mCurrentSection)(mEditingItemIdx), TABLE_SEP) > 0 Then
            newText = Split(mSecItems(mCurrentSection)(mEditingItemIdx), TABLE_SEP)(0)
        Else
            newText = mSecItems(mCurrentSection)(mEditingItemIdx)
        End If
        
        txtItem.Text = newText
        mLoading = True
        lstItems.ListIndex = mEditingItemIdx - 1
        mLoading = False
        
        UpdateItemButtonsBasedOnFocus
        Exit Sub
    End If

    CancelAnyActiveEditMode
    UpdateItemControls
    UpdateItemButtonsBasedOnFocus
    UpdateSubItemButtonsBasedOnFocus
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
    
    If lstSubItems.ListIndex < 0 Then Exit Sub

    ' If in SubSubItem Edit mode (drilled in), switching sub-sub-items
    ' is allowed - mirrors the mEditingSubIdx auto-save-on-switch pattern
    If mEditingSubSubIdx >= 0 Then

        Dim newSubSubIdx As Long
        newSubSubIdx = lstSubItems.ListIndex

        Dim parentIdx3 As Long
        parentIdx3 = lstItems.ListIndex + 1

        ' Always save current text when switching (no prompt)
        If Trim(txtSubItem.Text) <> "" Then
            SaveEditedSubSubItemTextInPlace parentIdx3, mDrilledSubIdx
            mLoading = True
            RefreshSubSubItems
            mLoading = False
        End If

        ' Switch to newly selected sub-sub-item
        mEditingSubSubIdx = newSubSubIdx

        ' Guard against out of bounds, and load the new text
        Dim parts3() As String
        parts3 = Split(mSecItems(mCurrentSection)(parentIdx3), TABLE_SEP)
        If UBound(parts3) >= 1 Then
            Dim subs3() As String
            subs3 = Split(parts3(1), LIST_SEP)
            If mDrilledSubIdx <= UBound(subs3) Then
                Dim subParts3() As String
                If InStr(subs3(mDrilledSubIdx), SUBSUB_SEP) > 0 Then
                    subParts3 = Split(subs3(mDrilledSubIdx), SUBSUB_SEP)
                    If UBound(subParts3) >= 1 Then
                        Dim subsubs3() As String
                        subsubs3 = Split(subParts3(1), SUBSUB_ITEM_SEP)
                        If mEditingSubSubIdx > UBound(subsubs3) Then mEditingSubSubIdx = UBound(subsubs3)
                        txtSubItem.Text = subsubs3(mEditingSubSubIdx)
                    End If
                End If
            End If
        End If

        mLoading = True
        lstSubItems.ListIndex = mEditingSubSubIdx
        mLoading = False

        UpdateSubSubItemButtonsBasedOnFocus
        Exit Sub
    End If

    ' If in SubItem Edit mode, switching sub-items is allowed
    If mEditingSubIdx >= 0 Then

        ' Capture newly clicked index BEFORE any refresh
        Dim newIdx As Long
        newIdx = lstSubItems.ListIndex

        Dim parentIdx As Long
        parentIdx = lstItems.ListIndex + 1

        ' Always save current text when switching (no prompt)
        If Trim(txtSubItem.Text) <> "" Then
            SaveEditedSubItemTextInPlace parentIdx
            mLoading = True
            RefreshSubItems
            mLoading = False
        End If

        ' Switch to newly selected sub-item
        mEditingSubIdx = newIdx

        ' Guard against out of bounds
        Dim parts() As String
        parts = Split(mSecItems(mCurrentSection)(parentIdx), TABLE_SEP)
        If UBound(parts) >= 1 Then
            Dim subs() As String
            subs = Split(parts(1), LIST_SEP)
            If mEditingSubIdx > UBound(subs) Then mEditingSubIdx = UBound(subs)
            txtSubItem.Text = subs(mEditingSubIdx)
        End If

        mLoading = True
        lstSubItems.ListIndex = mEditingSubIdx
        mLoading = False

        UpdateSubItemButtonsBasedOnFocus
        Exit Sub
    End If

    CancelAnyActiveEditMode
    UpdateItemButtonsBasedOnFocus
    UpdateSubItemButtonsBasedOnFocus
    Exit Sub

ErrHandler:
    HandleFormError "lstSubItems_Click"

End Sub

Private Sub txtItem_Change()
    On Error GoTo ErrHandler
    
    UpdateItemButtonsBasedOnFocus
    Exit Sub
    
ErrHandler:
    HandleFormError "txtItem_Change"
    
End Sub

Private Sub txtSubItem_Change()
    On Error GoTo ErrHandler
    
    UpdateSubItemButtonsBasedOnFocus
    Exit Sub
    
ErrHandler:
    HandleFormError "txtSubItem_Change"
    
End Sub

'====================================================
' ITEM ADD / EDIT / REMOVE / REORDER
'====================================================

Private Sub btnAddItem_Click()
    On Error GoTo ErrHandler
    
    DebugLog "btnAddItem_Click: mode=" & IIf(mEditingItemIdx > 0, "EDIT", "ADD")
    If mEditingItemIdx > 0 Then
        ' EDIT MODE: Move Up
        ReorderItem mEditingItemIdx, -1, mSecItems(mCurrentSection).count
        RefreshItemsAfterReorder
    Else
        ' If there's unsaved text in the Sub-item field, ask how to proceed
        ' If there's unsaved text in the Sub-item field, ask how to proceed
        If mSecHasSubItems(mCurrentSection) And Len(Trim(txtSubItem.Text)) > 0 Then

            Dim parentDisplay As String
            If lstItems.ListIndex >= 0 Then
                parentDisplay = lstItems.List(lstItems.ListIndex)
            Else
                parentDisplay = "(no item selected)"
            End If

            Dim subResp As VbMsgBoxResult
            If mSkipConfirmPrompts Then
                subResp = vbYes
            Else
                subResp = MsgBox("Do you want to add """ & Trim(txtSubItem.Text) & """ to """ & parentDisplay & """?" & vbCrLf & vbCrLf & _
                     "Otherwise it will be discarded.", _
                     vbYesNoCancel + vbQuestion, "Unsaved Text")
            End If
            
            If subResp = vbCancel Then
                Exit Sub
            ElseIf subResp = vbYes Then
                ' Add the Sub-item only, under whichever parent is currently
                ' selected - Item text stays untouched in its box
                Dim selIdx As Long
                selIdx = lstItems.ListIndex

                If selIdx < 0 Then
                    MsgBox "Select a parent item first.", vbExclamation
                    Exit Sub
                End If

                Dim itemIdx As Long
                itemIdx = selIdx + 1

                Dim rawSubLines() As String
                rawSubLines = Split(txtSubItem.Text, vbCrLf)

                Dim subAddedAny As Boolean
                subAddedAny = False

                Dim current As String
                current = mSecItems(mCurrentSection)(itemIdx)

                Dim parts2() As String
                parts2 = Split(current, TABLE_SEP)

                Dim mainPart As String
                mainPart = parts2(0)

                Dim existingSubs As String
                If UBound(parts2) >= 1 Then
                    existingSubs = parts2(1)
                Else
                    existingSubs = ""
                End If

                Dim subLineIdx As Long
                For subLineIdx = 0 To UBound(rawSubLines)
                    Dim cleanSubLine As String
                    cleanSubLine = Trim(rawSubLines(subLineIdx))
                    If cleanSubLine <> "" Then
                        If existingSubs = "" Then
                            existingSubs = cleanSubLine
                        Else
                            existingSubs = existingSubs & LIST_SEP & cleanSubLine
                        End If

                        lstSubItems.AddItem TruncateForDisplay(cleanSubLine)
                        subAddedAny = True
                    End If
                Next subLineIdx

                If Not subAddedAny Then
                    MsgBox "Please type or paste at least one sub-item before adding.", vbExclamation
                    Exit Sub
                End If

                current = mainPart & TABLE_SEP & existingSubs

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

                ' Capture item text before touching lstSubItems, in case
                ' any side-effect handler clears it
                Dim savedItemText As String
                savedItemText = txtItem.Text

                txtSubItem.Text = ""
                lstSubItems.ListIndex = lstSubItems.ListCount - 1
                UpdateSubItemControls

                txtItem.Text = savedItemText

                Exit Sub
            ElseIf subResp = vbNo Then
                ' Add the Item now. Sub-item text stays untouched in its box -
                ' fall through to normal add logic below.
            End If
        End If

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
        ' EDIT MODE: Move Down
        ReorderItem mEditingItemIdx, 1, mSecItems(mCurrentSection).count
        RefreshItemsAfterReorder
    Else
    
        Dim selIdx As Long
        selIdx = lstItems.ListIndex
        If selIdx < 0 Then
            MsgBox "Select an item to remove.", vbExclamation
            Exit Sub
        End If

        Dim itemIdx As Long
        itemIdx = selIdx + 1

        ' Check if this item has sub-items, for a stronger warning
        Dim hasSubs As Boolean
        hasSubs = False

        If mSecHasSubItems(mCurrentSection) Then
            Dim parts() As String
            parts = Split(mSecItems(mCurrentSection)(itemIdx), TABLE_SEP)
            If UBound(parts) >= 1 Then
                If Trim(parts(1)) <> "" Then hasSubs = True
            End If
        End If

        Dim resp As VbMsgBoxResult
        If hasSubs Then
            resp = MsgBox("This item has sub-items attached to it." & vbCrLf & vbCrLf & _
                          "Removing it will permanently delete the item AND all of its sub-items." & vbCrLf & vbCrLf & _
                          "Are you sure?", vbYesNo + vbExclamation, "Remove Item and Sub-items?")
            If resp = vbNo Then Exit Sub
        End If

        If resp = vbNo Then Exit Sub

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
        
        ' Re-select the item that's now at this position (or the last item,
        ' if we removed the final row), and refresh its sub-items display
        If lstItems.ListCount > 0 Then
            Dim newSelIdx As Long
            newSelIdx = selIdx
            If newSelIdx >= lstItems.ListCount Then newSelIdx = lstItems.ListCount - 1
        
            lstItems.ListIndex = newSelIdx
            UpdateSubItemsLabel
            RefreshSubItems
        Else
            lstSubItems.Clear
            UpdateSubItemControls
        End If
        
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

        ' Save any pending rename for current item
        If newItem <> "" Then
            SaveEditedItemTextInPlace
        End If
        
        PushUndo
        
        Dim savedIdx As Long
        savedIdx = mEditingItemIdx
        
        ExitItemEditMode
        RefreshItemsDisplay
        
        mLoading = True
        lstItems.ListIndex = savedIdx - 1
        mLoading = False
        
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
    
        ' Take snapshot ONCE when entering edit mode
        Set mEditSnapshot = CloneCollection(mSecItems(mCurrentSection))
        mEditSnapshotSec = mCurrentSection
    
        mEditingItemIdx = itemIdx
    
        ' Load current item text
        Dim current As String
        
        If mSecHasSubItems(mCurrentSection) And InStr(mSecItems(mCurrentSection)(itemIdx), TABLE_SEP) > 0 Then
            current = Split(mSecItems(mCurrentSection)(itemIdx), TABLE_SEP)(0)
        Else
            current = mSecItems(mCurrentSection)(itemIdx)
        End If
    
        txtItem.Text = current
        txtItem.SetFocus
    
        btnEditItem.Caption = "Confirm"
        btnAddItem.Caption = "Move Up"
        btnRemoveItem.Caption = "Move Down"
        btnCancel.Caption = "Cancel"
    
        ' Visual indicator: Raised border
        lstItems.SpecialEffect = 2  ' Sunken
    
        UpdateSubItemControls
        UpdateItemButtonsBasedOnFocus
        UpdateSubItemButtonsBasedOnFocus
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
    
    Set mEditSnapshot = Nothing
    lstItems.SpecialEffect = 3

    UpdateItemButtonsBasedOnFocus
    UpdateSubItemButtonsBasedOnFocus   ' restore sub-item controls

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
    btnRemoveItem.Caption = "– Remove"
    btnCancel.Caption = "Close Editor"
    txtItem.Text = ""
    
    Set mEditSnapshot = Nothing
    lstItems.SpecialEffect = 3

    UpdateItemButtonsBasedOnFocus
    UpdateSubItemButtonsBasedOnFocus

End Sub


Private Sub ReorderSubItem(ByRef subIndex As Long, _
                           ByVal direction As Long, _
                           ByVal itemIdx As Long)
    On Error GoTo ErrHandler

    Dim parts() As String
    parts = Split(mSecItems(mCurrentSection)(itemIdx), TABLE_SEP)
    Dim subs() As String
    subs = Split(parts(1), LIST_SEP)

    If direction = -1 And subIndex <= 0 Then Exit Sub
    If direction = 1 And subIndex >= UBound(subs) Then Exit Sub

    SaveEditedSubItemTextInPlace itemIdx
    PushUndo

    parts = Split(mSecItems(mCurrentSection)(itemIdx), TABLE_SEP)
    subs = Split(parts(1), LIST_SEP)

    Dim tmp As String
    tmp = subs(subIndex)
    subs(subIndex) = subs(subIndex + direction)
    subs(subIndex + direction) = tmp
    Dim newStr As String
    newStr = parts(0) & TABLE_SEP & Join(subs, LIST_SEP)
    
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
    subIndex = subIndex + direction
    Exit Sub

ErrHandler:
    HandleFormError "ReorderSubItem"

End Sub

Private Sub ReorderSubSubItem(ByRef subSubIndex As Long, _
                               ByVal direction As Long, _
                               ByVal itemIdx As Long, _
                               ByVal subIdx As Long)
    On Error GoTo ErrHandler

    Dim parts() As String
    parts = Split(mSecItems(mCurrentSection)(itemIdx), TABLE_SEP)
    Dim subs() As String
    subs = Split(parts(1), LIST_SEP)

    Dim subParts() As String
    If InStr(subs(subIdx), SUBSUB_SEP) = 0 Then Exit Sub
    subParts = Split(subs(subIdx), SUBSUB_SEP)

    Dim subsubs() As String
    subsubs = Split(subParts(1), SUBSUB_ITEM_SEP)

    If direction = -1 And subSubIndex <= 0 Then Exit Sub
    If direction = 1 And subSubIndex >= UBound(subsubs) Then Exit Sub

    SaveEditedSubSubItemTextInPlace itemIdx, subIdx
    PushUndo

    ' Re-split after the save, same defensive re-read pattern as ReorderSubItem
    parts = Split(mSecItems(mCurrentSection)(itemIdx), TABLE_SEP)
    subs = Split(parts(1), LIST_SEP)
    subParts = Split(subs(subIdx), SUBSUB_SEP)
    subsubs = Split(subParts(1), SUBSUB_ITEM_SEP)

    Dim tmp As String
    tmp = subsubs(subSubIndex)
    subsubs(subSubIndex) = subsubs(subSubIndex + direction)
    subsubs(subSubIndex + direction) = tmp

    subs(subIdx) = subParts(0) & SUBSUB_SEP & Join(subsubs, SUBSUB_ITEM_SEP)

    Dim newItemStr As String
    newItemStr = parts(0) & TABLE_SEP & Join(subs, LIST_SEP)

    Dim newCol As New Collection
    Dim i As Long
    For i = 1 To mSecItems(mCurrentSection).count
        If i = itemIdx Then
            newCol.Add newItemStr
        Else
            newCol.Add mSecItems(mCurrentSection)(i)
        End If
    Next i

    Set mSecItems(mCurrentSection) = newCol
    subSubIndex = subSubIndex + direction
    Exit Sub

ErrHandler:
    HandleFormError "ReorderSubSubItem"

End Sub

Private Sub SaveEditedItemTextInPlace()

    Dim newItem As String
    newItem = Trim(txtItem.Text)
    
    If newItem = "" Then Exit Sub

    Dim newCol As New Collection
    Dim i As Long
    For i = 1 To mSecItems(mCurrentSection).count
        If i = mEditingItemIdx Then
            If mSecTypes(mCurrentSection) = TYPE_LIST And InStr(mSecItems(mCurrentSection)(i), TABLE_SEP) > 0 Then
                Dim parts() As String
                parts = Split(mSecItems(mCurrentSection)(i), TABLE_SEP)
                parts(0) = newItem
                newCol.Add Join(parts, TABLE_SEP)
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
    
    If mDrilledSubIdx >= 0 Then
        AddSubSubItem
        Exit Sub
    End If
    
    Dim selIdx As Long
    selIdx = lstItems.ListIndex
    If selIdx < 0 Then
        MsgBox "Select a parent item first.", vbExclamation
        Exit Sub
    End If
    
    Dim itemIdx As Long
    itemIdx = selIdx + 1
    
    If mEditingSubIdx >= 0 Then
        ' EDIT MODE: Move Up
        ReorderSubItem mEditingSubIdx, -1, itemIdx
        RefreshSubItemsAfterReorder
    Else
        ' If there's unsaved text in the Item field, ask how to proceed
        If Len(Trim(txtItem.Text)) > 0 Then
            Dim itemResp As VbMsgBoxResult
            itemResp = MsgBox("Do you want to add """ & Trim(txtItem.Text) & """ to """ & mSecNames(mCurrentSection) & """?" & vbCrLf & vbCrLf & _
                  "Otherwise it will be discarded.", _
                  vbYesNoCancel + vbQuestion, "Unsaved Text")
            If itemResp = vbCancel Then
                Exit Sub
            ElseIf itemResp = vbYes Then
                ' Add the Item only - Sub-item text stays untouched in its box
                Dim pendingLines() As String
                pendingLines = Split(txtItem.Text, vbCrLf)
    
                Dim pendingAdded As Boolean
                pendingAdded = False
    
                Dim pendingIdx As Long
                For pendingIdx = 0 To UBound(pendingLines)
                    Dim pendingClean As String
                    pendingClean = Trim(pendingLines(pendingIdx))
                    If pendingClean <> "" Then
                        mSecItems(mCurrentSection).Add pendingClean
                        lstItems.AddItem pendingClean
                        pendingAdded = True
                    End If
                Next pendingIdx
    
                If Not pendingAdded Then
                    MsgBox "Please type or paste at least one item before adding.", vbExclamation
                    Exit Sub
                End If
    
                ' Capture sub-item text before selecting the new item, since
                ' lstItems_Click may clear txtSubItem as a side effect
                Dim savedSubText As String
                savedSubText = txtSubItem.Text
    
                txtItem.Text = ""
                lstItems.ListIndex = lstItems.ListCount - 1
                UpdateItemControls
    
                txtSubItem.Text = savedSubText
    
                Exit Sub
            ElseIf itemResp = vbNo Then
                ' Add the Sub-item now, using whichever parent is currently selected.
                ' Item text stays untouched in its box - fall through to normal add logic below.
            End If
        End If
    
        ' NORMAL MODE: split pasted text on line breaks, each line
        ' becomes its own sub-item
        Dim rawLines() As String
        rawLines = Split(txtSubItem.Text, vbCrLf)
        Dim addedAny As Boolean
        addedAny = False
        Dim current As String
        current = mSecItems(mCurrentSection)(itemIdx)
        Dim parts2() As String
        parts2 = Split(current, TABLE_SEP)
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
                    existingSubs = existingSubs & LIST_SEP & cleanLine
                End If
                
                lstSubItems.AddItem TruncateForDisplay(cleanLine)
                addedAny = True
            End If
        Next lineIdx
        
        If Not addedAny Then
            MsgBox "Please type or paste at least one sub-item before adding.", vbExclamation
            Exit Sub
        End If
        
        current = mainPart & TABLE_SEP & existingSubs
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
    
    If mDrilledSubIdx >= 0 Then
        RemoveSubSubItem
        Exit Sub
    End If

    Dim parentIdx As Long
    parentIdx = lstItems.ListIndex
    If parentIdx < 0 Then Exit Sub

    Dim itemIdx As Long
    itemIdx = parentIdx + 1

    If mEditingSubIdx >= 0 Then
        ' EDIT MODE: Move Up
        ReorderSubItem mEditingSubIdx, 1, itemIdx
        RefreshSubItemsAfterReorder
    Else
        Dim subIdx As Long
        subIdx = lstSubItems.ListIndex
        If subIdx < 0 Then
            MsgBox "Select a sub-item to remove.", vbExclamation
            Exit Sub
        End If

        ' Warn if this sub-item has its own sub-sub-items that would be lost
        Dim parentIdxCheck As Long
        parentIdxCheck = itemIdx
        Dim partsCheck() As String
        partsCheck = Split(mSecItems(mCurrentSection)(parentIdxCheck), TABLE_SEP)
        If UBound(partsCheck) >= 1 Then
            Dim subsCheck() As String
            subsCheck = Split(partsCheck(1), LIST_SEP)
            If subIdx <= UBound(subsCheck) Then
                If InStr(subsCheck(subIdx), SUBSUB_SEP) > 0 Then
                    Dim removeResp As VbMsgBoxResult
                    removeResp = MsgBox("This sub-item has bullet points, which will also be permanently removed." & vbCrLf & vbCrLf & _
                                        "Continue?", vbYesNo + vbExclamation, "Remove Sub-item and bullet points?")
                    If removeResp = vbNo Then Exit Sub
                End If
            End If
        End If
        
        PushUndo

        Dim parts2() As String
        parts2 = Split(mSecItems(mCurrentSection)(itemIdx), TABLE_SEP)

        Dim newStr2 As String
        newStr2 = parts2(0)

        If UBound(parts2) > 0 Then
            Dim subs2() As String
            subs2 = Split(parts2(1), LIST_SEP)

            Dim newSubs As String
            Dim j As Long
            For j = 0 To UBound(subs2)
                If j <> subIdx Then
                    If newSubs = "" Then
                        newSubs = subs2(j)
                    Else
                        newSubs = newSubs & LIST_SEP & subs2(j)
                    End If
                End If
            Next j

            If newSubs <> "" Then newStr2 = newStr2 & TABLE_SEP & newSubs
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
    
    If mDrilledSubIdx >= 0 Then
        EditSubSubItem
        Exit Sub
    End If

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
        parts = Split(mSecItems(mCurrentSection)(itemIdx), TABLE_SEP)

        Dim subs() As String
        If UBound(parts) >= 1 Then
            subs = Split(parts(1), LIST_SEP)
        Else
            ReDim subs(0)
        End If

        subs(mEditingSubIdx) = newSub

        Dim newStr As String
        newStr = parts(0) & TABLE_SEP & Join(subs, LIST_SEP)

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
        parts2 = Split(mSecItems(mCurrentSection)(itemIdx2), TABLE_SEP)

        If UBound(parts2) < 1 Then Exit Sub

        Dim subs2() As String
        subs2 = Split(parts2(1), LIST_SEP)

        If subIdx > UBound(subs2) Then Exit Sub

        txtSubItem.Text = subs2(subIdx)
        txtSubItem.SetFocus

        Set mSubEditSnapshot = CloneCollection(mSecItems(mCurrentSection))
        mSubEditSnapshotSec = mCurrentSection
        mSubEditSnapshotItemIdx = lstItems.ListIndex

        mEditingSubIdx = subIdx
        
        UpdateItemControls
        
        btnEditSubItem.Caption = "Confirm"
        btnAddSubItem.Caption = "Move Up"
        btnRemoveSubItem.Caption = "Move Down"
        btnCancel.Caption = "Cancel"
        
        lstSubItems.SpecialEffect = 2
        
        UpdateItemButtonsBasedOnFocus
        UpdateSubItemButtonsBasedOnFocus
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
    
    Set mSubEditSnapshot = Nothing
    lstSubItems.SpecialEffect = 3

    UpdateItemControls
    UpdateItemButtonsBasedOnFocus
    UpdateSubItemButtonsBasedOnFocus

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
    btnRemoveSubItem.Caption = "– Remove"
    btnCancel.Caption = "Close Editor"
    txtSubItem.Text = ""
    
    Set mSubEditSnapshot = Nothing
    lstSubItems.SpecialEffect = 3

    UpdateItemControls
    UpdateSubItemControls

End Sub

Private Sub CancelSubSubItemEditMode()

    If Not mSubSubEditSnapshot Is Nothing Then
        If mSubSubEditSnapshotSec = mCurrentSection Then
            Set mSecItems(mCurrentSection) = CloneCollection(mSubSubEditSnapshot)
            RefreshSubSubItems
            mLoading = True
            If mSubSubEditSnapshotSubIdx <= lstSubItems.ListCount - 1 Then
                lstSubItems.ListIndex = mSubSubEditSnapshotSubIdx
            End If
            mLoading = False
        End If
    End If

    mEditingSubSubIdx = -1
    btnEditSubItem.Caption = "Edit"
    btnAddSubItem.Caption = "+ Add"
    btnRemoveSubItem.Caption = "– Remove"
    btnCancel.Caption = "Close Editor"
    txtSubItem.Text = ""

    Set mSubSubEditSnapshot = Nothing
    lstSubItems.SpecialEffect = 3

    UpdateSubSubItemControls

End Sub

Private Sub SaveEditedSubItemTextInPlace(ByVal itemIdx As Long)

    Dim newSub As String
    newSub = Trim(txtSubItem.Text)
    If newSub = "" Then Exit Sub

    Dim parts() As String
    parts = Split(mSecItems(mCurrentSection)(itemIdx), TABLE_SEP)

    Dim subs() As String
    If UBound(parts) >= 1 Then
        subs = Split(parts(1), LIST_SEP)
    Else
        ReDim subs(0)
    End If

    subs(mEditingSubIdx) = newSub

    Dim newStr As String
    newStr = parts(0) & TABLE_SEP & Join(subs, LIST_SEP)

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

Private Sub SaveEditedSubSubItemTextInPlace(ByVal itemIdx As Long, ByVal subIdx As Long)

    Dim newSubSub As String
    newSubSub = Trim(txtSubItem.Text)
    If newSubSub = "" Then Exit Sub

    Dim parts() As String
    parts = Split(mSecItems(mCurrentSection)(itemIdx), TABLE_SEP)

    Dim subs() As String
    subs = Split(parts(1), LIST_SEP)

    Dim subParts() As String
    If InStr(subs(subIdx), SUBSUB_SEP) > 0 Then
        subParts = Split(subs(subIdx), SUBSUB_SEP)
    Else
        ReDim subParts(0)
        subParts(0) = subs(subIdx)
    End If

    Dim subsubs() As String
    If UBound(subParts) >= 1 Then
        subsubs = Split(subParts(1), SUBSUB_ITEM_SEP)   ' CHANGED
    Else
        ReDim subsubs(0)
    End If

    subsubs(mEditingSubSubIdx) = newSubSub

    ' Rebuild bottom-up: subsub array -> sub-item string -> subs array -> item string
    Dim newSubItemStr As String
    newSubItemStr = subParts(0) & SUBSUB_SEP & Join(subsubs, SUBSUB_ITEM_SEP)   ' CHANGED

    subs(subIdx) = newSubItemStr

    Dim newStr As String
    newStr = parts(0) & TABLE_SEP & Join(subs, LIST_SEP)

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

Private Sub AddSubSubItem()
    On Error GoTo ErrHandler

    Dim parentIdx As Long
    parentIdx = lstItems.ListIndex + 1

    ' Move Up branch:
    If mEditingSubSubIdx >= 0 Then
        ReorderSubSubItem mEditingSubSubIdx, -1, parentIdx, mDrilledSubIdx
        RefreshSubSubItemsAfterReorder
        Exit Sub
    End If

    ' NORMAL MODE: split pasted text on line breaks, each line becomes
    ' its own sub-sub-item. Mirrors btnAddSubItem_Click's normal-mode block.
    Dim rawLines() As String
    rawLines = Split(txtSubItem.Text, vbCrLf)
    Dim addedAny As Boolean
    addedAny = False

    Dim parts() As String
    parts = Split(mSecItems(mCurrentSection)(parentIdx), TABLE_SEP)
    Dim mainPart As String
    mainPart = parts(0)

    Dim subs() As String
    subs = Split(parts(1), LIST_SEP)

    Dim subParts() As String
    If InStr(subs(mDrilledSubIdx), SUBSUB_SEP) > 0 Then
        subParts = Split(subs(mDrilledSubIdx), SUBSUB_SEP)
    Else
        ReDim subParts(0)
        subParts(0) = subs(mDrilledSubIdx)
    End If

    Dim subText As String
    subText = subParts(0)

    Dim existingSubSubs As String
    If UBound(subParts) >= 1 Then
        existingSubSubs = subParts(1)
    Else
        existingSubSubs = ""
    End If

    Dim lineIdx As Long
    For lineIdx = 0 To UBound(rawLines)
        Dim cleanLine As String
        cleanLine = Trim(rawLines(lineIdx))
        If cleanLine <> "" Then
            If existingSubSubs = "" Then
                existingSubSubs = cleanLine
            Else
                existingSubSubs = existingSubSubs & SUBSUB_ITEM_SEP & cleanLine   ' CHANGED
            End If
            lstSubItems.AddItem TruncateForDisplay(cleanLine)
            addedAny = True
        End If
    Next lineIdx

    If Not addedAny Then
        MsgBox "Please type or paste at least one bullet point before adding.", vbExclamation
        Exit Sub
    End If

    subs(mDrilledSubIdx) = subText & SUBSUB_SEP & existingSubSubs

    Dim newItemStr As String
    newItemStr = mainPart & TABLE_SEP & Join(subs, LIST_SEP)

    Dim newCol As New Collection
    Dim j As Long
    For j = 1 To mSecItems(mCurrentSection).count
        If j = parentIdx Then
            newCol.Add newItemStr
        Else
            newCol.Add mSecItems(mCurrentSection)(j)
        End If
    Next j

    Set mSecItems(mCurrentSection) = newCol
    txtSubItem.Text = ""
    mLoading = True
    lstSubItems.ListIndex = lstSubItems.ListCount - 1
    mLoading = False
    UpdateSubSubItemButtonsBasedOnFocus

    Exit Sub

ErrHandler:
    HandleFormError "AddSubSubItem"

End Sub

Private Sub RemoveSubSubItem()
    On Error GoTo ErrHandler

    Dim parentIdx As Long
    parentIdx = lstItems.ListIndex + 1

    ' Move Down branch:
    If mEditingSubSubIdx >= 0 Then
        ReorderSubSubItem mEditingSubSubIdx, 1, parentIdx, mDrilledSubIdx
        RefreshSubSubItemsAfterReorder
        Exit Sub
    End If

    Dim subSubIdx As Long
    subSubIdx = lstSubItems.ListIndex

    If subSubIdx < 0 Then
        MsgBox "Select a bullet point to remove.", vbExclamation
        Exit Sub
    End If

    PushUndo

    Dim parts() As String
    parts = Split(mSecItems(mCurrentSection)(parentIdx), TABLE_SEP)

    Dim subs() As String
    subs = Split(parts(1), LIST_SEP)

    Dim subParts() As String
    subParts = Split(subs(mDrilledSubIdx), SUBSUB_SEP)
    Dim subText As String
    subText = subParts(0)

    Dim newSubSubs As String
    If UBound(subParts) >= 1 Then
        Dim subsubs() As String
        subsubs = Split(subParts(1), SUBSUB_ITEM_SEP)

        Dim j As Long
        For j = 0 To UBound(subsubs)
            If j <> subSubIdx Then
                If newSubSubs = "" Then
                    newSubSubs = subsubs(j)
                Else
                    newSubSubs = newSubSubs & SUBSUB_ITEM_SEP & subsubs(j)
                End If
            End If
        Next j
    End If

    If newSubSubs <> "" Then
        subs(mDrilledSubIdx) = subText & SUBSUB_SEP & newSubSubs
    Else
        subs(mDrilledSubIdx) = subText   ' no children left - drop the marker entirely
    End If

    Dim newItemStr As String
    newItemStr = parts(0) & TABLE_SEP & Join(subs, LIST_SEP)

    Dim newCol As New Collection
    Dim k As Long
    For k = 1 To mSecItems(mCurrentSection).count
        If k = parentIdx Then
            newCol.Add newItemStr
        Else
            newCol.Add mSecItems(mCurrentSection)(k)
        End If
    Next k
    Set mSecItems(mCurrentSection) = newCol

    lstSubItems.RemoveItem subSubIdx
    UpdateSubSubItemControls

    Exit Sub

ErrHandler:
    HandleFormError "RemoveSubSubItem"

End Sub

Private Sub EditSubSubItem()
    On Error GoTo ErrHandler

    Dim parentIdx As Long
    parentIdx = lstItems.ListIndex + 1

    If mEditingSubSubIdx >= 0 Then

        ' CONFIRM
        Dim newSubSub As String
        newSubSub = Trim(txtSubItem.Text)

        If newSubSub = "" Then
            MsgBox "Please type a bullet point before confirming.", vbExclamation
            Exit Sub
        End If

        PushUndo
        SaveEditedSubSubItemTextInPlace parentIdx, mDrilledSubIdx

        Dim savedIdx As Long
        savedIdx = mEditingSubSubIdx

        ExitSubSubItemEditMode
        RefreshSubSubItems
        mLoading = True
        lstSubItems.ListIndex = savedIdx
        mLoading = False

    Else
        Dim subSubIdx As Long
        subSubIdx = lstSubItems.ListIndex

        If subSubIdx < 0 Then
            MsgBox "Select a bullet point to edit.", vbExclamation
            Exit Sub
        End If

        Dim parts() As String
        parts = Split(mSecItems(mCurrentSection)(parentIdx), TABLE_SEP)
        Dim subs() As String
        subs = Split(parts(1), LIST_SEP)

        Dim subParts() As String
        If InStr(subs(mDrilledSubIdx), SUBSUB_SEP) = 0 Then Exit Sub
        subParts = Split(subs(mDrilledSubIdx), SUBSUB_SEP)

        Dim subsubs() As String
        subsubs = Split(subParts(1), SUBSUB_ITEM_SEP)   ' CHANGED

        If subSubIdx > UBound(subsubs) Then Exit Sub

        txtSubItem.Text = subsubs(subSubIdx)
        txtSubItem.SetFocus

        Set mSubSubEditSnapshot = CloneCollection(mSecItems(mCurrentSection))
        mSubSubEditSnapshotSec = mCurrentSection
        mSubSubEditSnapshotSubIdx = subSubIdx   ' FIXED - was mDrilledSubIdx

        mEditingSubSubIdx = subSubIdx

        btnEditSubItem.Caption = "Confirm"
        btnAddSubItem.Caption = "Move Up"
        btnRemoveSubItem.Caption = "Move Down"
        btnCancel.Caption = "Cancel"

        lstSubItems.SpecialEffect = 2

        UpdateSubSubItemButtonsBasedOnFocus
    End If

    Exit Sub

ErrHandler:
    HandleFormError "EditSubSubItem"

End Sub

Private Sub ExitSubSubItemEditMode()

    mEditingSubSubIdx = -1
    btnEditSubItem.Caption = "Edit"
    btnAddSubItem.Caption = "+ Add"
    btnRemoveSubItem.Caption = "– Remove"
    btnCancel.Caption = "Close Editor"
    txtSubItem.Text = ""

    Set mSubSubEditSnapshot = Nothing
    lstSubItems.SpecialEffect = 3

    UpdateSubSubItemButtonsBasedOnFocus

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

'====================================================
' HELPER: Reusable Item Reorder Logic
'====================================================

Private Sub ReorderItem(ByRef itemIndex As Long, _
                        ByVal direction As Long, _
                        ByVal maxCount As Long)
    ' This function moves an item up (-1) or down (+1) in a collection
    '
    ' Parameters:
    '   itemIndex - the 1-based index of the item to move (gets updated by this function)
    '   direction - pass -1 to move UP, pass +1 to move DOWN
    '   maxCount  - the total number of items in the collection
    '
    ' Example: ReorderItem mEditingItemIdx, -1, mSecItems(mCurrentSection).count
    '          This moves the current item UP one position
    On Error GoTo ErrHandler

    ' Bounds check: can't move up past item 1, can't move down past last item
    If direction = -1 And itemIndex <= 1 Then Exit Sub
    If direction = 1 And itemIndex >= maxCount Then Exit Sub

    ' Do the swap
    SaveEditedItemTextInPlace
    PushUndo
    
    Set mSecItems(mCurrentSection) = _
        SwapCollectionItems(mSecItems(mCurrentSection), itemIndex, itemIndex + direction)

    ' Update the index to point to the new location
    itemIndex = itemIndex + direction
    Exit Sub

ErrHandler:
    HandleFormError "ReorderItem"

End Sub

Private Sub RefreshItemsAfterReorder()
    On Error GoTo ErrHandler

    RefreshItemsDisplay
    
    mLoading = True
    lstItems.ListIndex = mEditingItemIdx - 1
    mLoading = False
    
    UpdateItemButtonsBasedOnFocus  ' ADD THIS
    txtItem.SetFocus
    Exit Sub

ErrHandler:
    HandleFormError "RefreshItemsAfterReorder"

End Sub

Private Sub RefreshSubItemsAfterReorder()
    On Error GoTo ErrHandler

    RefreshSubItems
    
    mLoading = True
    lstSubItems.ListIndex = mEditingSubIdx
    mLoading = False
    
    UpdateSubItemButtonsBasedOnFocus  ' ADD THIS
    txtSubItem.SetFocus
    Exit Sub

ErrHandler:
    HandleFormError "RefreshSubItemsAfterReorder"

End Sub

Private Sub RefreshSubSubItemsAfterReorder()
    On Error GoTo ErrHandler

    RefreshSubSubItems

    mLoading = True
    lstSubItems.ListIndex = mEditingSubSubIdx
    mLoading = False

    UpdateSubSubItemButtonsBasedOnFocus
    txtSubItem.SetFocus
    Exit Sub

ErrHandler:
    HandleFormError "RefreshSubSubItemsAfterReorder"

End Sub

Private Sub RefreshRowsAfterReorder()
    On Error GoTo ErrHandler

    RefreshRowList mCurrentSection
    
    mLoading = True
    lstRows.ListIndex = mEditingRowIdx - 1
    mLoading = False
    Exit Sub

ErrHandler:
    HandleFormError "RefreshRowsAfterReorder"

End Sub

Private Sub SetListIndexWithoutTriggeringClick(ByVal listCtrl As Object, ByVal idx As Long)

    mLoading = True
    listCtrl.ListIndex = idx
    mLoading = False
    
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
        colCount = UBound(Split(mSecCols(idx), LIST_SEP)) + 1
    Else
        colCount = 1
        mSecCols(idx) = "col1"
    End If
    
    mLoading = True
    spnColCount.Value = colCount
    lblColCountValue.Caption = CStr(colCount)
    chkHasHeader.Value = mSecHasHeader(idx)
    mLoading = False
    ' Genuine custom Tables: spinner and header checkbox always fully editable
    spnColCount.Enabled = True
    chkHasHeader.Enabled = True
    lstRows.Visible = True
    btnAddRow.Visible = True
    btnRemoveRow.Visible = True
    btnEditRow.Visible = True
    fraRowEditor.Visible = True
    BuildRowEditorFields
    RefreshRowList idx
    UpdateRowButtonsBasedOnContent
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
    cols = Split(mSecCols(mCurrentSection), LIST_SEP)
    Dim topPos As Long
    topPos = 15
    
    Dim i As Long
    For i = 0 To UBound(cols)
        Dim lbl As MSForms.Label
        Set lbl = fraRowEditor.Controls.Add("Forms.Label.1", "lbl_col_" & i)
        lbl.Tag = "dynamic"
        lbl.Left = 6
        lbl.Top = topPos
        lbl.Width = 70
        lbl.Height = 20
        lbl.TextAlign = 3   ' fmTextAlignRight
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
    Dim isTableLike As Boolean
    isTableLike = (mSecTypes(mCurrentSection) = TYPE_TABLE Or _
                   mSecTypes(mCurrentSection) = TYPE_RESOURCES Or _
                   mSecTypes(mCurrentSection) = TYPE_DICTIONARY)
    If Not isTableLike Then Exit Sub
    mSecHasHeader(mCurrentSection) = chkHasHeader.Value

    BuildRowEditorFields                ' ? ADD: re-assert ColumnWidths BEFORE RefreshRowList touches .List
    RefreshRowList mCurrentSection      ' updates « » indicator
    RefreshRowEditorLabels
    If mEditingRowIdx > 0 Then
        UpdateRowReorderButtonStates    ' re-evaluate Move Up/Down for the row being edited
    End If
    Exit Sub
    
ErrHandler:
    HandleFormError "chkHasHeader_Click"
    
End Sub

Private Sub RefreshRowEditorLabels()
    On Error GoTo ErrHandler

    If mSecCols(mCurrentSection) = "" Then Exit Sub

    Dim cols() As String
    cols = Split(mSecCols(mCurrentSection), LIST_SEP)

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
    headerParts = Split(mSecItems(mCurrentSection)(1), TABLE_SEP)

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
    
    ' Only allow column changes for Table, or Dict/AR in Dev Mode
    Dim allowChange As Boolean
    allowChange = (mSecTypes(mCurrentSection) = TYPE_TABLE) Or _
                  ((mSecTypes(mCurrentSection) = TYPE_RESOURCES Or mSecTypes(mCurrentSection) = TYPE_DICTIONARY) And mDeveloperMode)

    If Not allowChange Then Exit Sub

    Dim oldCount As Long
    If mSecCols(mCurrentSection) <> "" Then
        oldCount = UBound(Split(mSecCols(mCurrentSection), LIST_SEP)) + 1
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
            newCols = newCols & LIST_SEP & "col" & i
        End If
    Next i

    mSecCols(mCurrentSection) = newCols

    If newCount < oldCount Then
        TrimRowsToColumnCount mCurrentSection, newCount
    End If

    mEditingRowIdx = 0
    btnAddRow.Caption = "+ Add"
    btnRemoveRow.Caption = "– Remove"
    btnEditRow.Caption = "Edit"

    BuildRowEditorFields
    RefreshRowList mCurrentSection
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
        parts = Split(mSecItems(secIdx)(i), TABLE_SEP)

        Dim trimmed As String
        Dim c As Long
        For c = 0 To newCount - 1
            If c <= UBound(parts) Then
                If c = 0 Then
                    trimmed = parts(c)
                Else
                    trimmed = trimmed & TABLE_SEP & parts(c)
                End If
            Else
                If c = 0 Then
                    trimmed = ""
                Else
                    trimmed = trimmed & TABLE_SEP
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
    cols = Split(mSecCols(mCurrentSection), LIST_SEP)

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
            result = result & TABLE_SEP & val
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
    cols = Split(mSecCols(mCurrentSection), LIST_SEP)

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
    parts = Split(mSecItems(mCurrentSection)(itemIdx), TABLE_SEP)

    Dim cols() As String
    cols = Split(mSecCols(mCurrentSection), LIST_SEP)

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

    Dim isTableLike As Boolean
    isTableLike = (mSecTypes(mCurrentSection) = TYPE_TABLE Or _
                   mSecTypes(mCurrentSection) = TYPE_RESOURCES Or _
                   mSecTypes(mCurrentSection) = TYPE_DICTIONARY)

    If Not isTableLike Then Exit Sub
    If mEditingRowIdx <= 0 Then Exit Sub
    If mEditingRowIdx > mSecItems(mCurrentSection).count Then Exit Sub

    SaveTableRowFieldsInPlace mEditingRowIdx
    Exit Sub
    
ErrHandler:
    HandleFormError "SaveRowEditor"
    
End Sub

'====================================================
' RESOURCES / DICTIONARY (fixed-field tables)
'====================================================

Private Sub ShowResourcesSection(ByVal idx As Long)
    On Error GoTo ErrHandler
    
    ' Resources is a restricted Table: 3 columns, header always on
    If mSecCols(idx) = "" Then mSecCols(idx) = "col1" & LIST_SEP & "col2" & LIST_SEP & "col3"
    lblColCount.Visible = True
    spnColCount.Visible = True
    lblColCountValue.Visible = True
    chkHasHeader.Visible = True
    
    Dim colCount As Long
    colCount = 3
    mLoading = True
    spnColCount.Value = colCount
    lblColCountValue.Caption = CStr(colCount)
    mSecHasHeader(idx) = True   ' Resources always has header
    chkHasHeader.Value = True
    mLoading = False
    ' Restrict spinner and checkbox unless Dev Mode
    If mDeveloperMode Then
        spnColCount.Enabled = True
        chkHasHeader.Enabled = True
    Else
        spnColCount.Enabled = False
        chkHasHeader.Enabled = False
    End If
    
    lstRows.Visible = True
    btnAddRow.Visible = True
    btnRemoveRow.Visible = True
    btnEditRow.Visible = True
    fraRowEditor.Visible = True
    
    ' Ensure header row exists with correct labels
    If mSecItems(idx).count = 0 Then
        mSecItems(idx).Add SEC_RESOURCES_COL1 & TABLE_SEP & SEC_RESOURCES_COL2 & TABLE_SEP & SEC_RESOURCES_COL3
    End If
    
    BuildRowEditorFields
    RefreshRowList idx
    Exit Sub
    
ErrHandler:
    mLoading = False
    HandleFormError "ShowResourcesSection"
    
End Sub

Private Sub ShowDictionarySection(ByVal idx As Long)
    On Error GoTo ErrHandler
    
    ' Dictionary is a restricted Table: 2 columns, header always on
    If mSecCols(idx) = "" Then mSecCols(idx) = "col1" & LIST_SEP & "col2"
    lblColCount.Visible = True
    spnColCount.Visible = True
    lblColCountValue.Visible = True
    chkHasHeader.Visible = True
    
    Dim colCount As Long
    colCount = 2
    mLoading = True
    spnColCount.Value = colCount
    lblColCountValue.Caption = CStr(colCount)
    mSecHasHeader(idx) = True   ' Dictionary always has header
    chkHasHeader.Value = True
    mLoading = False
    ' Restrict spinner and checkbox unless Dev Mode
    If mDeveloperMode Then
        spnColCount.Enabled = True
        chkHasHeader.Enabled = True
    Else
        spnColCount.Enabled = False
        chkHasHeader.Enabled = False
    End If
    
    lstRows.Visible = True
    btnAddRow.Visible = True
    btnRemoveRow.Visible = True
    btnEditRow.Visible = True
    fraRowEditor.Visible = True
    
    ' Ensure header row exists with correct labels
    If mSecItems(idx).count = 0 Then
        mSecItems(idx).Add SEC_DICT_COL1 & TABLE_SEP & SEC_DICT_COL2
    End If
    BuildRowEditorFields
    RefreshRowList idx
    Exit Sub
    
ErrHandler:
    mLoading = False
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

Private Sub ExitRowEditMode()

    mEditingRowIdx = 0
    btnAddRow.Caption = "+ Add"
    btnRemoveRow.Caption = "– Remove"
    btnEditRow.Caption = "Edit"
    btnCancel.Caption = "Close Editor"
    lstRows.SpecialEffect = 3
    
    UpdateRowControls
    UpdateRowButtonsBasedOnContent

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
    btnRemoveRow.Caption = "– Remove"
    btnCancel.Caption = "Close Editor"
    lstRows.SpecialEffect = 3

    ClearTableRowFields
    RefreshRowEditorLabels

    Set mRowEditSnapshot = Nothing
    UpdateRowControls
    UpdateRowButtonsBasedOnContent   ' <-- ADD THIS LINE

End Sub

'====================================================
' ROW LIST / SELECTION
'====================================================

Private Sub RefreshRowList(ByVal idx As Long)
    On Error GoTo ErrHandler
    
    lstRows.Clear
    lstRows.ColumnCount = 1

    Dim i As Long
    For i = 1 To mSecItems(idx).count
        Dim parts() As String
        parts = Split(mSecItems(idx)(i), TABLE_SEP)
        Dim display As String
        display = ""
        
        Dim j As Long
        For j = 0 To UBound(parts)
            Dim val As String
            val = Trim(parts(j))
            If j = 0 Then
                display = val
            Else
                display = display & "  |  " & val
            End If
        Next j
        
        If Replace(display, "|", "") = "" Then display = Replace(display, "|", "") & "(empty row " & i & ")"
        
        ' Add • prefix to every column-ish segment, if this is row 1 and header checkbox is checked
        If i = 1 And mSecHasHeader(idx) Then
            display = "« " & display & " »"
        End If
        
        lstRows.AddItem TruncateForDisplay(display)
    Next i
    
    If mSecTypes(idx) = TYPE_TABLE Then
        If mDeveloperMode Then
            chkHasHeader.Enabled = True
            chkHasHeader.ControlTipText = "DEV: Unlocked for testing"
        Else
            chkHasHeader.Enabled = (mSecItems(idx).count >= 1)
            If Not chkHasHeader.Enabled Then
                chkHasHeader.Value = False
                mSecHasHeader(idx) = False
                chkHasHeader.ControlTipText = "Cannot change — add at least one row first"
            Else
                chkHasHeader.ControlTipText = "Check if the first row contains headers"
            End If
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

    Dim isTableLike As Boolean
    isTableLike = (mSecTypes(mCurrentSection) = TYPE_TABLE Or _
                   mSecTypes(mCurrentSection) = TYPE_RESOURCES Or _
                   mSecTypes(mCurrentSection) = TYPE_DICTIONARY)

    If mEditingRowIdx > 0 Then
        Dim newIdx As Long
        newIdx = lstRows.ListIndex + 1
    
        If isTableLike Then
            SaveTableRowFieldsInPlace mEditingRowIdx
            ' Update just the row we're leaving, so its display reflects
            ' any edits - no full rebuild needed
            UpdateSingleRowDisplay mEditingRowIdx, mSecItems(mCurrentSection)(mEditingRowIdx), _
                                   (mEditingRowIdx = 1 And mSecHasHeader(mCurrentSection))
        End If
    
        mEditingRowIdx = newIdx
    
        If mEditingRowIdx < 1 Or mEditingRowIdx > mSecItems(mCurrentSection).count Then
            mEditingRowIdx = 1
        End If
    
        If isTableLike Then
            LoadTableRowIntoFields mEditingRowIdx
            RefreshRowEditorLabels
        End If
    
        UpdateRowReorderButtonStates
        Exit Sub
    End If

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

    Dim isHeaderRow As Boolean
    isHeaderRow = (lstRows.ListIndex = 0 And mSecHasHeader(mCurrentSection))

    Dim isRestrictedType As Boolean
    isRestrictedType = (mSecTypes(mCurrentSection) = TYPE_RESOURCES Or mSecTypes(mCurrentSection) = TYPE_DICTIONARY)

    If isHeaderRow And isRestrictedType And Not mDeveloperMode Then
        ' Dict/AR: Edit & Remove grayed for header row
        btnEditRow.Enabled = False
        btnRemoveRow.Enabled = False
    Else
        ' Regular Table (or Dev Mode): Edit allowed, Remove allowed (with warning)
        btnEditRow.Enabled = hasSelection
        btnRemoveRow.Enabled = hasSelection
    End If

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

    Dim isTableLike As Boolean
    isTableLike = (mSecTypes(mCurrentSection) = TYPE_TABLE Or _
                   mSecTypes(mCurrentSection) = TYPE_RESOURCES Or _
                   mSecTypes(mCurrentSection) = TYPE_DICTIONARY)
    If Not isTableLike Then Exit Sub
    If mEditingRowIdx > 0 Then
        If mEditingRowIdx <= 1 Then Exit Sub
        If mEditingRowIdx = 2 And mSecHasHeader(mCurrentSection) Then Exit Sub
        SaveTableRowFieldsInPlace mEditingRowIdx
        PushUndo
        Set mSecItems(mCurrentSection) = _
            SwapCollectionItems(mSecItems(mCurrentSection), mEditingRowIdx, mEditingRowIdx - 1)
        mEditingRowIdx = mEditingRowIdx - 1
        mLoading = True
        RefreshRowList mCurrentSection
        lstRows.ListIndex = mEditingRowIdx - 1
        mLoading = False
        LoadTableRowIntoFields mEditingRowIdx
        UpdateRowReorderButtonStates
    Else
        If mSecCols(mCurrentSection) = "" Then
            MsgBox "Set a column count before adding rows.", vbExclamation
            Exit Sub
        End If
        Dim newRow As String
        newRow = ReadTableRowFields()
        If Replace(newRow, TABLE_SEP, "") = "" Then
            MsgBox "Enter at least one field before adding.", vbExclamation
            Exit Sub
        End If
        mSecItems(mCurrentSection).Add newRow
        RefreshRowList mCurrentSection
        ClearTableRowFields
    End If
    Exit Sub
ErrHandler:
    HandleFormError "btnAddRow_Click"
End Sub

Private Sub btnRemoveRow_Click()
    On Error GoTo ErrHandler

    Dim isTableLike As Boolean
    isTableLike = (mSecTypes(mCurrentSection) = TYPE_TABLE Or _
                   mSecTypes(mCurrentSection) = TYPE_RESOURCES Or _
                   mSecTypes(mCurrentSection) = TYPE_DICTIONARY)

    If Not isTableLike Then Exit Sub

    Dim newCol As New Collection
    Dim i As Long

    If mEditingRowIdx > 0 Then
        ' EDIT MODE: this button is "Move Down"
        If mEditingRowIdx >= mSecItems(mCurrentSection).count Then Exit Sub
        If mEditingRowIdx = 1 And mSecHasHeader(mCurrentSection) Then Exit Sub  ' can't move header down

        SaveTableRowFieldsInPlace mEditingRowIdx
        PushUndo

        Set mSecItems(mCurrentSection) = _
            SwapCollectionItems(mSecItems(mCurrentSection), mEditingRowIdx, mEditingRowIdx + 1)

        mEditingRowIdx = mEditingRowIdx + 1

        mLoading = True
        RefreshRowList mCurrentSection
        lstRows.ListIndex = mEditingRowIdx - 1
        mLoading = False

        LoadTableRowIntoFields mEditingRowIdx
        UpdateRowReorderButtonStates
    Else
        Dim selIdx As Long
        selIdx = lstRows.ListIndex

        If selIdx < 0 Then
            MsgBox "Select a row to remove.", vbExclamation
            Exit Sub
        End If

        Dim rowIdx As Long
        rowIdx = selIdx + 1

        ' Warn if trying to remove the header row
        If rowIdx = 1 And mSecHasHeader(mCurrentSection) Then
            ' Dict/AR: header removal is blocked entirely unless Dev Mode
            If (mSecTypes(mCurrentSection) = TYPE_RESOURCES Or mSecTypes(mCurrentSection) = TYPE_DICTIONARY) And Not mDeveloperMode Then
                MsgBox "This header row is required and cannot be removed.", vbExclamation
                Exit Sub
            End If

            ' Regular Table: warn before allowing removal
            Dim resp As VbMsgBoxResult
            resp = MsgBox("This is your header row." & vbCrLf & vbCrLf & _
                          "Delete it anyway?", vbYesNo + vbExclamation, "Delete Header Row?")
            If resp = vbNo Then Exit Sub
        End If

        PushUndo

        For i = 1 To mSecItems(mCurrentSection).count
            If i <> rowIdx Then newCol.Add mSecItems(mCurrentSection)(i)
        Next i
        Set mSecItems(mCurrentSection) = newCol

        RefreshRowList mCurrentSection
    End If

    Exit Sub

ErrHandler:
    HandleFormError "btnRemoveRow_Click"

End Sub

Private Sub btnEditRow_Click()
    On Error GoTo ErrHandler
    
    Dim isTableLike As Boolean
    isTableLike = (mSecTypes(mCurrentSection) = TYPE_TABLE Or _
                   mSecTypes(mCurrentSection) = TYPE_RESOURCES Or _
                   mSecTypes(mCurrentSection) = TYPE_DICTIONARY)

    If Not isTableLike Then Exit Sub

    If mEditingRowIdx > 0 Then
        ' CONFIRM
        PushUndo
        SaveTableRowFieldsInPlace mEditingRowIdx

        Dim wasEditingRow1 As Boolean
        wasEditingRow1 = (mEditingRowIdx = 1)

        Dim savedIdx As Long
        savedIdx = mEditingRowIdx

        ExitRowEditMode
        RefreshRowList mCurrentSection
        ClearTableRowFields

        If wasEditingRow1 Then RefreshRowEditorLabels
    Else
        Dim selIdx As Long
        selIdx = lstRows.ListIndex

        If selIdx < 0 Then
            MsgBox "Select a row to edit.", vbExclamation
            Exit Sub
        End If

        Dim itemIdx As Long
        itemIdx = selIdx + 1

        ' Dict/AR: block editing header row unless Dev Mode
        If itemIdx = 1 And mSecHasHeader(mCurrentSection) Then
            If (mSecTypes(mCurrentSection) = TYPE_RESOURCES Or mSecTypes(mCurrentSection) = TYPE_DICTIONARY) And Not mDeveloperMode Then
                MsgBox "This header row is locked and cannot be edited.", vbExclamation
                Exit Sub
            End If
        End If

        Set mRowEditSnapshot = CloneCollection(mSecItems(mCurrentSection))
        mRowEditSnapshotSec = mCurrentSection

        LoadTableRowIntoFields itemIdx

        mEditingRowIdx = itemIdx
        btnEditRow.Caption = "Confirm"
        btnAddRow.Caption = "Move Up"
        btnRemoveRow.Caption = "Move Down"
        btnCancel.Caption = "Cancel"

        ' Visual indicator: Raised border
        lstRows.SpecialEffect = 2  ' Sunken

        UpdateRowReorderButtonStates
        RefreshRowEditorLabels
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

                    If IsBuiltInSubItemSection(mSecNames(secIdx)) Then
                        mSecHasSubItems(secIdx) = True
                    Else
                        mSecHasSubItems(secIdx) = hasSubs
                    End If
                    
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
        ReDim mSecHasSubSubItems(1 To 1)
    Else
        ReDim Preserve mSecNames(1 To mSecCount)
        ReDim Preserve mSecTypes(1 To mSecCount)
        ReDim Preserve mSecData(1 To mSecCount)
        ReDim Preserve mSecItems(1 To mSecCount)
        ReDim Preserve mSecCols(1 To mSecCount)
        ReDim Preserve mSecHasHeader(1 To mSecCount)
        ReDim Preserve mSecHasSubItems(1 To mSecCount)
        ReDim Preserve mSecHasSubSubItems(1 To mSecCount)
    End If
    
    mSecNames(mSecCount) = secName
    mSecTypes(mSecCount) = secType
    mSecData(mSecCount) = ""
    mSecCols(mSecCount) = ""
    mSecHasHeader(mSecCount) = False
    mSecHasSubItems(mSecCount) = False
    mSecHasSubSubItems(mSecCount) = False
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
    Dim currentMain As String        ' Item text
    Dim currentSub As String         ' most recent Bullet1 (level-1) text
    Dim currentSubSubs As String     ' LIST_SEP-joined children of currentSub
    Dim currentSubs As String        ' LIST_SEP-joined level-1 entries, each possibly
                                      ' carrying its own SUBSUB_SEP-attached children
    Dim haveEntry As Boolean
    Dim haveSub As Boolean           ' True once currentSub has been started

    haveEntry = False
    haveSub = False
    currentMain = ""
    currentSub = ""
    currentSubSubs = ""
    currentSubs = ""

    For r = 1 To tbl.ListRows.count
        Dim mainVal As String
        mainVal = Trim(tbl.DataBodyRange(r, parentCol).Value)

        Dim bullet1Val As String, bullet2Val As String
        bullet1Val = ""
        bullet2Val = ""
        If lastBulletCol >= parentCol + 1 Then
            bullet1Val = Trim(tbl.DataBodyRange(r, parentCol + 1).Value)
        End If
        If lastBulletCol >= parentCol + 2 Then
            bullet2Val = Trim(tbl.DataBodyRange(r, parentCol + 2).Value)
        End If

        If mainVal <> "" Then
            ' New Item starts. Flush whatever was being built - first
            ' close out any pending level-1 bullet, then the item itself.
            If haveEntry Then
                FlushPendingSub currentSubs, currentSub, currentSubSubs, haveSub
                FlushNestedEntry mSecItems(secIdx), currentMain, currentSubs
            End If

            currentMain = mainVal
            currentSubs = ""
            currentSub = ""
            currentSubSubs = ""
            haveSub = False
            haveEntry = True
        End If

        If haveEntry Then
            If bullet1Val <> "" Then
                ' New level-1 bullet starts. Flush whatever level-1 bullet
                ' (and its accumulated children) was previously being built.
                FlushPendingSub currentSubs, currentSub, currentSubSubs, haveSub

                currentSub = bullet1Val
                currentSubSubs = ""
                haveSub = True

                ' Same row can also carry its own level-2 child inline
                If bullet2Val <> "" Then
                    currentSubSubs = bullet2Val
                End If
            ElseIf bullet2Val <> "" Then
                ' Continuation row: level-2 value with no new level-1 on
                ' this row - attaches as another child of currentSub
                If haveSub Then
                    If currentSubSubs = "" Then
                        currentSubSubs = bullet2Val
                    Else
                        currentSubSubs = currentSubSubs & SUBSUB_ITEM_SEP & bullet2Val
                    End If
                End If
            End If
        End If
    Next r

    If haveEntry Then
        FlushPendingSub currentSubs, currentSub, currentSubSubs, haveSub
        FlushNestedEntry mSecItems(secIdx), currentMain, currentSubs
    End If

    Exit Sub

ErrHandler:
    HandleFormError "LoadNestedColumns"

End Sub

' Closes out one level-1 bullet (currentSub + its currentSubSubs children,
' if any) and appends it onto the running currentSubs LIST_SEP chain.
' No-op if no level-1 bullet is currently pending.
Private Sub FlushPendingSub(ByRef currentSubs As String, _
                            ByRef currentSub As String, _
                            ByRef currentSubSubs As String, _
                            ByRef haveSub As Boolean)

    If Not haveSub Then Exit Sub

    Dim entry As String
    entry = currentSub
    If currentSubSubs <> "" Then entry = entry & SUBSUB_SEP & currentSubSubs

    If currentSubs = "" Then
        currentSubs = entry
    Else
        currentSubs = currentSubs & LIST_SEP & entry
    End If

    currentSub = ""
    currentSubSubs = ""
    haveSub = False

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
    If subs <> "" Then entry = entry & TABLE_SEP & subs

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
            colStr = colStr & LIST_SEP & "col" & i
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
                rowStr = rowStr & TABLE_SEP & val
            End If
        Next c

        Dim isBlank As Boolean
        isBlank = (Replace(rowStr, TABLE_SEP, "") = "")

        Dim isRestrictedTable As Boolean
        isRestrictedTable = (mSecTypes(secIdx) = TYPE_RESOURCES Or mSecTypes(secIdx) = TYPE_DICTIONARY)
        
        If (mSecTypes(secIdx) = TYPE_TABLE Or isRestrictedTable) And r = 1 Then
            ' Row 1 is always kept as the header row (Table's reserved
            ' slot, or Resources/Dictionary's fixed label row)
            mSecItems(secIdx).Add rowStr
        
        ElseIf Not isBlank Then
        
            mSecItems(secIdx).Add rowStr
        End If
    Next r

    ' For generic Table sections specifically: figure out whether
    ' the checkbox should start checked or unchecked, based on
    ' whether row 1 (which we just loaded above) actually has any
    ' text in it or is completely blank.
    If mSecTypes(secIdx) = TYPE_TABLE Then
        If mSecItems(secIdx).count >= 1 Then
            Dim firstRowBlank As Boolean
            firstRowBlank = (Replace(mSecItems(secIdx)(1), TABLE_SEP, "") = "")
            mSecHasHeader(secIdx) = Not firstRowBlank
        Else
            mSecHasHeader(secIdx) = False
        End If
    ElseIf isRestrictedTable Then
        ' Dictionary/Resources ALWAYS have a header
        mSecHasHeader(secIdx) = True
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
    parts = Split(rowStr, TABLE_SEP)

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

    ' Check for unconfirmed text before anything else - same
    ' guardrail as section-switching, skipped in Dev Mode
    If Not mDeveloperMode And HasUnconfirmedContent() Then

        Dim saveResp As VbMsgBoxResult
        saveResp = MsgBox(BuildUnconfirmedPromptText(), vbYesNoCancel + vbQuestion, "Unsaved Text")

        Select Case saveResp
            Case vbCancel
                Exit Sub
            Case vbYes
                If Not ConfirmPendingChanges() Then Exit Sub
                ' fall through to save
            Case vbNo
                ' discard the pending text and continue to save
                CancelAnyActiveEditMode
        End Select

    End If

    If ws.ListObjects.count > 0 Then
        If ws.ListObjects(1).ListRows.count > 0 Then

            Dim resp As VbMsgBoxResult
            resp = MsgBox("Saving will overwrite the previous data on the sheet with the new data from this form. You may close the editor after saving." & vbCrLf & vbCrLf & _
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
                    ' Always advance by at least 2 (main col + Bullet1)
                    ' even when no sub-items exist yet - WriteNestedToSheet
                    ' always writes the Bullet1 header so it must be skipped
                    currentCol = currentCol + 1 + IIf(depth > 0, depth, 1)
                Else
                    WriteListToSheet ws, i, currentCol
                    currentCol = currentCol + 1
                End If

            Case TYPE_TABLE
                Dim tblCols As Long
                If mSecCols(i) <> "" Then
                    tblCols = UBound(Split(mSecCols(i), LIST_SEP)) + 1
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
        parts = Split(mSecItems(secIdx)(i), TABLE_SEP)

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
        parts = Split(mSecItems(secIdx)(i), TABLE_SEP)

        ws.Cells(r, startCol).Value = parts(0)

        Dim hasSubs As Boolean
        hasSubs = False

        If UBound(parts) > 0 Then
            If parts(1) <> "" Then hasSubs = True
        End If

        If hasSubs Then

            Dim subs() As String
            subs = Split(parts(1), LIST_SEP)

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
    colCount = UBound(Split(mSecCols(secIdx), LIST_SEP)) + 1

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
        parts = Split(mSecItems(secIdx)(i), TABLE_SEP)
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

    Dim r As Long
    r = 2   ' data (including header row) starts here now

    Dim i As Long
    For i = 1 To mSecItems(secIdx).count
        Dim parts() As String
        parts = Split(mSecItems(secIdx)(i), TABLE_SEP)
        
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
    
    ws.Cells(1, startCol).Value = mSecNames(secIdx)       ' "Additional Resources"
    ws.Cells(1, startCol + 1).Value = "Table2"
    ws.Cells(1, startCol + 2).Value = "Table3"

    Dim r As Long
    r = 2   ' data (including header row) starts here now

    Dim i As Long
    For i = 1 To mSecItems(secIdx).count
        Dim parts() As String
        parts = Split(mSecItems(secIdx)(i), TABLE_SEP)
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
        btnRemoveSection.Caption = "– Remove Section"
        btnCancel.Caption = "Close Editor"
        Exit Sub
    End If

    If mEditingSectionIdx > 0 Then
        mEditingSectionIdx = 0
        txtSectionName.Visible = False
        txtSectionName.Text = ""
        btnEditSection.Caption = "Edit Section"
        btnAddSection.Caption = "+ Add Section"
        btnRemoveSection.Caption = "– Remove Section"
        btnCancel.Caption = "Close Editor"
        Exit Sub
    End If

    If mEditingItemIdx > 0 Then
        CancelItemEditMode
        RefreshItemsDisplay
        Exit Sub
    End If

    If mEditingSubSubIdx >= 0 Then
        CancelSubSubItemEditMode
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

    If mDeveloperMode Then
        ConfirmDiscard = True
        Exit Function
    End If

    Dim resp As VbMsgBoxResult
    resp = MsgBox("Closing SOD Editor form." & vbCrLf & vbCrLf & _
                  "Do you want to save changes to sheet before exiting?" & vbCrLf & _
                  "Otherwise any unsaved changes will be discarded.", _
                  vbYesNoCancel + vbExclamation, "Save Before Closing?")

    Select Case resp
        Case vbYes
            Call btnSave_Click
            ' Only close if the save actually completed - if the user
            ' cancelled out of the overwrite confirmation inside
            ' btnSave_Click, mJustSaved will still be False
            ConfirmDiscard = mJustSaved
        Case vbNo
            ConfirmDiscard = True
        Case vbCancel
            ConfirmDiscard = False
    End Select

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

    If mEditingSubSubIdx >= 0 Then
        CancelSubSubItemEditMode
        Exit Sub
    End If

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



