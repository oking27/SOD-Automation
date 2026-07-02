Attribute VB_Name = "Module7"
Option Explicit

'====================================================
' SOD POPULATION
'
' No external template file. All formatting (styles, fonts, colors) is
' defined in the FORMATTING section below and built into each generated
' document at runtime via BuildStyles. Edit that section to change the
' document's appearance.
'====================================================

' --- validation severities ---
Private Const SEV_FATAL As String = "FATAL"
Private Const SEV_WARN As String = "WARNING"

' --- Word constants (late-bound, so mirrored locally) ---
Private Const wdBulletGallery As Long = 2
Private Const wdListApplyToWholeList As Long = 0
Private Const wdWord10ListBehavior As Long = 2
Private Const wdStyleTypeParagraph As Long = 1

'====================================================
' FORMATTING — edit these to change document appearance
'====================================================
Private Const DOCUMENT_SUBTITLE As String = "Standard Operating Document"

Private Const GREG_BLUE As Long = &H88431B

Private Const TITLE_FONT As String = "Bebas Neue"
Private Const TITLE_SIZE As Long = 20
Private Const TITLE_COLOR As Long = GREG_BLUE

Private Const SUBTITLE_FONT As String = "Lato"
Private Const SUBTITLE_SIZE As Long = 12

Private Const HEADING_FONT As String = "Bebas Neue"
Private Const HEADING_SIZE As Long = 18

Private Const BODY_FONT As String = "Aptos"
Private Const BODY_SIZE As Long = 12


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
        MsgBox "No issues found. Sheet is ready to populate.", vbInformation, "Validate SOD"
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
        MsgBox "Cannot populate:" & vbCrLf & vbCrLf & FormatIssueReport(issues), _
               vbCritical, "SOD Populator"
        Exit Sub
    End If

    If issues.Count > 0 Then
        Response = MsgBox(FormatIssueReport(issues) & vbCrLf & vbCrLf & _
            "Fix these now instead of continuing?", vbYesNo + vbExclamation, "SOD Population")
        If Response = vbYes Then Exit Sub
        ' else: continue best-effort; orphaned/unrecognized columns are skipped, as before
    End If

    ' ---- Determine document title (also used for the output filename) ----
    For TitleCol = 1 To tbl.ListColumns.Count
        If IsTitleColumn(tbl.ListColumns(TitleCol).Name) Then
            DocTitle = GetDocumentTitle(tbl, TitleCol)
            Exit For
        End If
    Next TitleCol

    If DocTitle = "" Then DocTitle = "Populated SOD"

    ' ---- Fresh blank document with our styles built in ----
    Set wdApp = GetWordApp(weCreatedApp)
    Set wdDoc = wdApp.Documents.Add

    BuildStyles wdDoc

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
        "SOD populated successfully." & vbCrLf & vbCrLf & _
        "Would you like to view the document?", _
        vbYesNo + vbQuestion, "SOD Populator")

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
' WORD APPLICATION
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

Private Sub BuildStyles(ByVal wdDoc As Object)

    Dim s As Object

    Set s = AddOrGetStyle(wdDoc, "SOD Title")
    With s.Font
        .Name = TITLE_FONT
        .Size = TITLE_SIZE
        .Color = GREG_BLUE
        .Spacing = 2
    End With

    Set s = AddOrGetStyle(wdDoc, "SOD Subtitle")
    With s.Font
        .Name = SUBTITLE_FONT
        .Size = SUBTITLE_SIZE
        .Allcaps = True
        .Spacing = 1
    End With

    Set s = AddOrGetStyle(wdDoc, "SOD Heading")
    With s.Font
        .Name = HEADING_FONT
        .Size = HEADING_SIZE
        .Color = GREG_BLUE
        .Spacing = 2
    End With
    s.ParagraphFormat.SpaceBefore = 12
    s.ParagraphFormat.SpaceAfter = 6

    Set s = AddOrGetStyle(wdDoc, "SOD Body")
    With s.Font
        .Name = BODY_FONT
        .Size = BODY_SIZE
    End With
    s.ParagraphFormat.SpaceAfter = 6

End Sub

Private Function AddOrGetStyle(ByVal wdDoc As Object, ByVal styleName As String) As Object

    Dim s As Object

    On Error Resume Next
    Set s = wdDoc.Styles(styleName)
    On Error GoTo 0

    If s Is Nothing Then
        Set s = wdDoc.Styles.Add(Name:=styleName, Type:=wdStyleTypeParagraph)
    End If

    Set AddOrGetStyle = s

End Function

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

    ' Documents.Add starts with a single empty paragraph; insert before it.
    Set rng = wdDoc.Content
    rng.InsertAfter TitleText & vbCr
    wdDoc.Paragraphs(1).Range.Style = "SOD Title"

    rng.Collapse 0
    rng.InsertAfter DOCUMENT_SUBTITLE & vbCr
    wdDoc.Paragraphs(2).Range.Style = "SOD Subtitle"

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

    wdTable.Style = "Table Grid"
    wdTable.Rows(1).Range.Bold = True
    wdTable.Rows(1).HeadingFormat = True
    wdTable.Rows(1).Shading.BackgroundPatternColor = RGB(230, 230, 230)   ' light gray header

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

    GetDocumentTitle = "Populated SOD"

End Function

