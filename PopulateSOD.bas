Attribute VB_Name = "Module7"
Option Explicit

'====================================================
' SOD GENERATOR
'
' Requires Template.docx built per Template_Spec.md:
'   Bookmarks : DocTitle, ContentStart
'   Styles    : "SOD Title", "SOD Subtitle", "SOD Heading", "SOD Body"
'   Table Sty : "SOD Table"
'====================================================

' --- registry keys (per-user, persists across sessions) ---
Private Const REGISTRY_APP As String = "SOPGenerator"
Private Const REGISTRY_SECTION As String = "Config"
Private Const REGISTRY_KEY_TEMPLATE As String = "TemplatePath"

' --- validation severities ---
Private Const SEV_FATAL As String = "FATAL"
Private Const SEV_WARN As String = "WARNING"

' --- Word constants (late-bound, so mirrored locally) ---
Private Const wdBulletGallery As Long = 2
Private Const wdListApplyToWholeList As Long = 0
Private Const wdWord10ListBehavior As Long = 2


'====================================================
' PUBLIC ENTRY POINTS
'====================================================

Public Sub ValidateSOD()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim issues As Collection

    On Error GoTo ErrHandler

    Set ws = ActiveSheet

    If ws.ListObjects.Count = 0 Then
        MsgBox "No table found on active sheet.", vbCritical
        Exit Sub
    End If

    Set tbl = ws.ListObjects(1)
    Set issues = ValidateSheet(tbl)

    If issues.Count = 0 Then
        MsgBox "No issues found. Sheet is ready to generate.", vbInformation, "Validate SOD"
    Else
        MsgBox FormatIssueReport(issues), vbExclamation, "Validate SOD"
    End If

    Exit Sub

ErrHandler:
    MsgBox "Validation error: " & Err.Description, vbCritical

End Sub


Public Sub PopulateSOD()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim issues As Collection

    Dim wdApp As Object
    Dim wdDoc As Object
    Dim weCreatedApp As Boolean

    Dim templatePath As String
    Dim tempPath As String
    Dim col As Long
    Dim hdr As String

    Dim TitleCol As Long
    Dim DocTitle As String

    Dim Response As VbMsgBoxResult

    On Error GoTo ErrHandler

    Set ws = ActiveSheet

    If ws.ListObjects.Count = 0 Then
        MsgBox "No table found on active sheet.", vbCritical
        Exit Sub
    End If

    Set tbl = ws.ListObjects(1)

    ' ---- Validation gate ----
    Set issues = ValidateSheet(tbl)

    If HasSeverity(issues, SEV_FATAL) Then
        MsgBox "Cannot generate:" & vbCrLf & vbCrLf & FormatIssueReport(issues), _
               vbCritical, "SOD Generator"
        Exit Sub
    End If

    If issues.Count > 0 Then
        Response = MsgBox(FormatIssueReport(issues) & vbCrLf & vbCrLf & _
            "Fix these now instead of generating?", vbYesNo + vbExclamation, "SOD Generator")
        If Response = vbYes Then Exit Sub
        ' else: continue best-effort; orphaned/unrecognized columns are skipped, as before
    End If

    ' ---- Locate template ----
    templatePath = GetTemplatePath()

    If templatePath = "" Or Dir(templatePath) = "" Then
        MsgBox "Could not locate Template.docx. Generation cancelled.", vbCritical
        Exit Sub
    End If

    ' ---- Determine document title (also used for the output filename) ----
    For TitleCol = 1 To tbl.ListColumns.Count
        If IsTitleColumn(tbl.ListColumns(TitleCol).Name) Then
            DocTitle = GetDocumentTitle(tbl, TitleCol)
            Exit For
        End If
    Next TitleCol

    If DocTitle = "" Then DocTitle = "Populated SOD"

    ' ---- Fresh working copy of the template ----
    tempPath = Environ$("TEMP") & "\SOD_" & Format(Now, "yyyymmdd_hhnnss") & ".docx"
    FileCopy templatePath, tempPath

    Set wdApp = GetWordApp(weCreatedApp)
    Set wdDoc = wdApp.Documents.Open(tempPath)

    ValidateTemplate wdDoc

    ' ---- Title ----
    WriteDocumentTitle wdDoc, DocTitle

    ' ---- Body ----
    col = 1

    Do While col <= tbl.ListColumns.Count

        hdr = Trim(tbl.ListColumns(col).Name)

        If Not IsTableColumn(hdr) And Not IsBulletColumn(hdr) And Not IsTitleColumn(hdr) Then

            WriteHeading wdDoc, hdr

            If HasTableColumns(tbl, col) Then
                CreateTableSection wdDoc, tbl, col
                col = GetLastTableColumn(tbl, col) + 1

            ElseIf HasBulletColumns(tbl, col) Then
                CreateNestedBulletSection wdDoc, tbl, col
                col = GetLastBulletColumn(tbl, col) + 1

            Else
                CreateContentSection wdDoc, tbl, col
                col = col + 1
            End If

        Else
            col = col + 1
        End If

    Loop

    ' ---- Finish ----
    Response = MsgBox( _
        "SOD generated successfully." & vbCrLf & vbCrLf & _
        "Would you like to view the document?", _
        vbYesNo + vbQuestion, "SOD Generator")

    wdApp.Visible = True

    If Response = vbYes Then
        ' leave open for review; user saves manually when ready

    Else
        Dim defaultName As String
        Dim saveTarget As Variant

        defaultName = SanitizeFilename("SOD " & DocTitle) & ".docx"

        saveTarget = Application.GetSaveAsFilename( _
            InitialFileName:=defaultName, _
            FileFilter:="Word Document (*.docx), *.docx")

        If Not saveTarget = False Then
            wdDoc.SaveAs2 FileName:=CStr(saveTarget)
        End If

        wdDoc.Close SaveChanges:=False

        If weCreatedApp Then
            If wdApp.Documents.Count = 0 Then wdApp.Quit
        End If
    End If

    On Error Resume Next
    Kill tempPath
    On Error GoTo 0

    Exit Sub

ErrHandler:

    Dim errMsg As String
    errMsg = Err.Description

    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If weCreatedApp Then
        If Not wdApp Is Nothing Then
            If wdApp.Documents.Count = 0 Then wdApp.Quit
        End If
    End If
    If tempPath <> "" Then Kill tempPath
    On Error GoTo 0

    MsgBox errMsg, vbCritical

End Sub


'====================================================
' VALIDATION
'====================================================

Private Function ValidateSheet(ByVal tbl As ListObject) As Collection

    Dim issues As New Collection
    Dim col As Long
    Dim hdr As String
    Dim titleCount As Long
    Dim expectingGroup As String  ' "", "PENDING", "TABLE", or "BULLET"

    If tbl.ListColumns.Count = 0 Then
        issues.Add SEV_FATAL & "|(sheet)|Table has no columns."
        Set ValidateSheet = issues
        Exit Function
    End If

    If tbl.ListRows.Count = 0 Then
        issues.Add SEV_FATAL & "|(sheet)|Table has no data rows."
        Set ValidateSheet = issues
        Exit Function
    End If

    expectingGroup = ""

    For col = 1 To tbl.ListColumns.Count

        hdr = Trim(tbl.ListColumns(col).Name)

        If IsTitleColumn(hdr) Then
            titleCount = titleCount + 1
            expectingGroup = ""

        ElseIf IsTableColumn(hdr) Then
            If expectingGroup <> "PENDING" And expectingGroup <> "TABLE" Then
                issues.Add SEV_WARN & "|" & hdr & _
                    "|Orphaned table column: not preceded by a heading column. Its content will not appear in the document."
            End If
            expectingGroup = "TABLE"

        ElseIf IsBulletColumn(hdr) Then
            If expectingGroup <> "PENDING" And expectingGroup <> "BULLET" Then
                issues.Add SEV_WARN & "|" & hdr & _
                    "|Orphaned bullet column: not preceded by a heading column. Its content will not appear in the document."
            End If
            expectingGroup = "BULLET"

        Else
            If IsLikelyTypo(hdr) Then
                issues.Add SEV_WARN & "|" & hdr & _
                    "|Header looks like a mistyped Table/Bullet column (space or extra characters) and will be treated as a section heading instead."
            End If
            expectingGroup = "PENDING"
        End If

    Next col

    If titleCount = 0 Then
        issues.Add SEV_WARN & "|(sheet)|No 'Title' column found. The document title will fall back to a generic name."
    ElseIf titleCount > 1 Then
        issues.Add SEV_WARN & "|(sheet)|Multiple 'Title' columns found. Only the first is used."
    End If

    Set ValidateSheet = issues

End Function

Private Function IsLikelyTypo(ByVal hdr As String) As Boolean

    Dim lower As String
    lower = LCase(Trim(hdr))

    If IsTableColumn(hdr) Or IsBulletColumn(hdr) Or IsTitleColumn(hdr) Then Exit Function

    If InStr(lower, "table") > 0 Or InStr(lower, "bullet") > 0 Then
        IsLikelyTypo = True
    End If

End Function

Private Function HasSeverity(ByVal issues As Collection, ByVal sev As String) As Boolean

    Dim itm As Variant

    For Each itm In issues
        If Split(itm, "|")(0) = sev Then
            HasSeverity = True
            Exit Function
        End If
    Next itm

End Function

Private Function FormatIssueReport(ByVal issues As Collection) As String

    Dim itm As Variant
    Dim parts() As String
    Dim s As String

    For Each itm In issues
        parts = Split(itm, "|", 3)
        s = s & "[" & parts(0) & "] " & parts(1) & ": " & parts(2) & vbCrLf
    Next itm

    FormatIssueReport = s

End Function


'====================================================
' WORD APPLICATION / TEMPLATE ACCESS
'====================================================

Private Function GetWordApp(ByRef weCreatedApp As Boolean) As Object

    Dim wdApp As Object

    weCreatedApp = False

    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    On Error GoTo 0

    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
        weCreatedApp = True
        wdApp.Visible = False   ' hide only while we build; restored to True before showing/saving
    End If

    Set GetWordApp = wdApp

End Function

Private Function DefaultTemplatePath() As String

    ' TODO: point this at your SharePoint-synced templates folder once set up
    DefaultTemplatePath = Environ$("USERPROFILE") & _
        "\CompanyName\SOP Site - Templates\Template.docx"

End Function

Private Function GetTemplatePath() As String

    Dim savedPath As String
    Dim picked As Variant

    savedPath = GetSetting(REGISTRY_APP, REGISTRY_SECTION, REGISTRY_KEY_TEMPLATE, "")

    If savedPath = "" Or Dir(savedPath) = "" Then

        If Dir(DefaultTemplatePath()) <> "" Then
            savedPath = DefaultTemplatePath()
        Else
            picked = Application.GetOpenFilename( _
                "Word Files (*.docx), *.docx", _
                Title:="Locate Template.docx")

            If picked = False Then
                GetTemplatePath = ""
                Exit Function
            End If

            savedPath = CStr(picked)
        End If

        SaveSetting REGISTRY_APP, REGISTRY_SECTION, REGISTRY_KEY_TEMPLATE, savedPath

    End If

    GetTemplatePath = savedPath

End Function

Private Sub ValidateTemplate(ByVal wdDoc As Object)

    Dim missing As String
    Dim names As Variant
    Dim i As Long
    Dim s As Object

    names = Array("DocTitle", "ContentStart")
    For i = LBound(names) To UBound(names)
        If Not wdDoc.Bookmarks.Exists(names(i)) Then
            missing = missing & "- Bookmark: " & names(i) & vbCrLf
        End If
    Next i

    names = Array("SOD Title", "SOD Subtitle", "SOD Heading", "SOD Body")
    For i = LBound(names) To UBound(names)
        Set s = Nothing
        On Error Resume Next
        Set s = wdDoc.Styles(names(i))
        On Error GoTo 0
        If s Is Nothing Then
            missing = missing & "- Style: " & names(i) & vbCrLf
        End If
    Next i

    If missing <> "" Then
        Err.Raise vbObjectError + 2, , _
            "Template.docx is missing required items:" & vbCrLf & missing
    End If

End Sub

Private Function SanitizeFilename(ByVal txt As String) As String

    Dim badChars As String
    Dim i As Long
    Dim ch As String

    badChars = "\/:*?""<>|"

    For i = 1 To Len(txt)
        ch = Mid(txt, i, 1)
        If InStr(badChars, ch) > 0 Then ch = "-"
        SanitizeFilename = SanitizeFilename & ch
    Next i

End Function


'====================================================
' TITLE / HEADINGS / PARAGRAPHS
'====================================================

Private Sub WriteDocumentTitle(ByVal wdDoc As Object, ByVal TitleText As String)

    Dim rng As Object

    Set rng = wdDoc.Bookmarks("DocTitle").Range
    rng.Text = TitleText
    wdDoc.Bookmarks.Add "DocTitle", rng
    rng.Style = "SOD Title"

End Sub

Private Sub WriteHeading(ByVal wdDoc As Object, ByVal txt As String)

    wdDoc.Content.InsertAfter txt & vbCr
    wdDoc.Paragraphs(wdDoc.Paragraphs.Count - 1).Range.Style = "SOD Heading"

End Sub

Private Sub WriteParagraph(ByVal wdDoc As Object, ByVal txt As String)

    wdDoc.Content.InsertAfter txt & vbCr
    wdDoc.Paragraphs(wdDoc.Paragraphs.Count - 1).Range.Style = "SOD Body"

End Sub


'====================================================
' CONTENT SECTIONS
'====================================================

Private Sub CreateContentSection(ByVal wdDoc As Object, _
                                 ByVal tbl As ListObject, _
                                 ByVal colNum As Long)

    Dim r As Long
    Dim lastRow As Long
    Dim block As Collection

    Set block = New Collection
    lastRow = tbl.ListRows.Count

    For r = 1 To lastRow + 1

        If r <= lastRow Then

            If Trim(tbl.DataBodyRange(r, colNum).Value) <> "" Then
                block.Add tbl.DataBodyRange(r, colNum).Value
            Else
                RenderBlock wdDoc, block
                Set block = New Collection
            End If

        Else
            RenderBlock wdDoc, block
        End If

    Next r

End Sub

Private Sub RenderBlock(ByVal wdDoc As Object, ByVal block As Collection)

    If block.Count = 0 Then Exit Sub

    If block.Count = 1 Then
        WriteParagraph wdDoc, block(1)
    Else
        WriteBulletList wdDoc, block
    End If

End Sub


'====================================================
' BULLETS (native Word multilevel list formatting)
'====================================================

Private Sub ApplyBulletLevel(ByVal rng As Object, ByVal level As Long)

    With rng.ListFormat
        .ApplyListTemplateWithLevel _
            ListTemplate:=rng.Application.ListGalleries(wdBulletGallery).ListTemplates(1), _
            ContinuePreviousList:=True, _
            ApplyTo:=wdListApplyToWholeList, _
            DefaultListBehavior:=wdWord10ListBehavior
        .ListLevelNumber = level
    End With

End Sub

Private Sub WriteBulletList(ByVal wdDoc As Object, ByVal items As Collection)

    Dim i As Long
    Dim startPara As Long
    Dim rng As Object

    startPara = wdDoc.Paragraphs.Count

    For i = 1 To items.Count
        wdDoc.Content.InsertAfter items(i) & vbCr
    Next i

    Set rng = wdDoc.Range( _
        Start:=wdDoc.Paragraphs(startPara).Range.Start, _
        End:=wdDoc.Paragraphs(wdDoc.Paragraphs.Count - 1).Range.End)

    rng.Style = "SOD Body"
    ApplyBulletLevel rng, 1

End Sub

Private Sub CreateNestedBulletSection(ByVal wdDoc As Object, _
                                      ByVal tbl As ListObject, _
                                      ByVal parentCol As Long)

    Dim r As Long, c As Long
    Dim lastRow As Long
    Dim bulletLastCol As Long
    Dim txt As String
    Dim level As Long
    Dim rng As Object

    bulletLastCol = GetLastBulletColumn(tbl, parentCol)
    lastRow = tbl.ListRows.Count

    For r = 1 To lastRow
        For c = parentCol To bulletLastCol

            txt = Trim(tbl.DataBodyRange(r, c).Value)

            If txt <> "" Then
                wdDoc.Content.InsertAfter txt & vbCr
                level = c - parentCol + 1
                Set rng = wdDoc.Paragraphs(wdDoc.Paragraphs.Count - 1).Range
                rng.Style = "SOD Body"
                ApplyBulletLevel rng, level
            End If

        Next c
    Next r

End Sub


'====================================================
' TABLES
'====================================================

Private Sub CreateTableSection(ByVal wdDoc As Object, _
                               ByVal tbl As ListObject, _
                               ByVal parentCol As Long)

    Dim firstCol As Long, lastCol As Long
    Dim rowCount As Long, colCount As Long
    Dim wdTable As Object
    Dim r As Long, c As Long

    firstCol = parentCol
    lastCol = GetLastTableColumn(tbl, parentCol)
    rowCount = LastUsedTableRow(tbl, firstCol, lastCol)

    If rowCount = 0 Then Exit Sub

    colCount = lastCol - firstCol + 1

    Set wdTable = wdDoc.Tables.Add(wdDoc.Content.Characters.Last, rowCount, colCount)

    For r = 1 To rowCount
        For c = 1 To colCount
            ' .Text (not .Value) preserves displayed formatting: dates,
            ' leading zeros on text-formatted numbers, etc.
            wdTable.Cell(r, c).Range.Text = tbl.DataBodyRange(r, firstCol + c - 1).Text
        Next c
    Next r

    On Error Resume Next
    wdTable.Style = "SOD Table"
    If Err.Number <> 0 Then
        Err.Clear
        wdTable.Style = "Table Grid"
    End If
    On Error GoTo 0

    wdTable.Rows(1).Range.Bold = True
    wdTable.Rows(1).HeadingFormat = True

    wdDoc.Content.InsertAfter vbCr

End Sub

Private Function LastUsedTableRow(ByVal tbl As ListObject, _
                                  ByVal firstCol As Long, _
                                  ByVal lastCol As Long) As Long

    Dim r As Long, c As Long

    For r = tbl.ListRows.Count To 1 Step -1
        For c = firstCol To lastCol
            If Trim(tbl.DataBodyRange(r, c).Value) <> "" Then
                LastUsedTableRow = r
                Exit Function
            End If
        Next c
    Next r

End Function


'====================================================
' HEADER DETECTION HELPERS
'====================================================

Private Function HasTableColumns(ByVal tbl As ListObject, ByVal colNum As Long) As Boolean
    If colNum >= tbl.ListColumns.Count Then Exit Function
    HasTableColumns = IsTableColumn(tbl.ListColumns(colNum + 1).Name)
End Function

Private Function HasBulletColumns(ByVal tbl As ListObject, ByVal colNum As Long) As Boolean
    If colNum >= tbl.ListColumns.Count Then Exit Function
    HasBulletColumns = IsBulletColumn(tbl.ListColumns(colNum + 1).Name)
End Function

Private Function GetLastTableColumn(ByVal tbl As ListObject, ByVal colNum As Long) As Long
    Dim c As Long
    c = colNum + 1
    Do While c <= tbl.ListColumns.Count
        If Not IsTableColumn(tbl.ListColumns(c).Name) Then Exit Do
        c = c + 1
    Loop
    GetLastTableColumn = c - 1
End Function

Private Function GetLastBulletColumn(ByVal tbl As ListObject, ByVal colNum As Long) As Long
    Dim c As Long
    c = colNum + 1
    Do While c <= tbl.ListColumns.Count
        If Not IsBulletColumn(tbl.ListColumns(c).Name) Then Exit Do
        c = c + 1
    Loop
    GetLastBulletColumn = c - 1
End Function

Private Function IsTableColumn(ByVal txt As String) As Boolean
    txt = Trim(txt)
    If InStr(txt, " ") > 0 Then Exit Function
    IsTableColumn = (LCase(Left(txt, 5)) = "table")
End Function

Private Function IsBulletColumn(ByVal txt As String) As Boolean
    txt = Trim(txt)
    If InStr(txt, " ") > 0 Then Exit Function
    IsBulletColumn = (LCase(Left(txt, 6)) = "bullet")
End Function

Private Function IsTitleColumn(ByVal txt As String) As Boolean
    IsTitleColumn = (LCase(Trim(txt)) = "title")
End Function

Private Function GetDocumentTitle(ByVal tbl As ListObject, ByVal TitleCol As Long) As String

    Dim r As Long

    For r = 1 To tbl.ListRows.Count
        If Trim(tbl.DataBodyRange(r, TitleCol).Value) <> "" Then
            GetDocumentTitle = Trim(tbl.DataBodyRange(r, TitleCol).Value)
            Exit Function
        End If
    Next r

    GetDocumentTitle = "Generated SOD"

End Function


