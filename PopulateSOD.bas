Attribute VB_Name = "PopSOD"
Option Explicit

'====================================================
' SOD POPULATION
'
' No external template file. All formatting (styles, fonts, colors) is
' defined in the FORMATTING section below and built into each generated
' document at runtime via BuildStyles. Edit that section to change the
' document's appearance.
'====================================================

' tracks whether the "Executive Overview" header has been written
Private mShortFormHeaderWritten As Boolean

' --- validation severities ---
Private Const SEV_FATAL As String = "FATAL"
Private Const SEV_WARN As String = "WARNING"

' --- Word constants (late-bound, so mirrored locally) ---
Private Const wdBulletGallery As Long = 1
Private Const wdListApplyToWholeList As Long = 0
Private Const wdWord10ListBehavior As Long = 2
Private Const wdStyleTypeParagraph As Long = 1
Private Const wdSectionBreakContinuous As Long = 3
Private Const wdHeaderFooterPrimary As Long = 1
Private Const wdHeaderFooterFirstPage As Long = 2
Private Const wdFieldPage As Long = 33
Private Const wdAlignParagraphRight As Long = 2
Private Const wdAlignParagraphLeft As Long = 0
Private Const wdAlignParagraphCenter As Long = 1

'====================================================
' FORMATTING - edit these to change document appearance
'====================================================
Private Const DOCUMENT_SUBTITLE As String = "Standard Operating Document"

Private Const GREG_BLUE As Long = &H953500
Private Const GREG_GRAY As Long = &H999896
Private Const GREG_YELLOW As Long = &H2AC6FF

Private Const TITLE_FONT As String = "Bebas Neue"
Private Const TITLE_SIZE As Long = 20
Private Const TITLE_COLOR As Long = GREG_BLUE

Private Const SUBTITLE_FONT As String = "Lato"
Private Const SUBTITLE_SIZE As Long = 12

Private Const HEADING_FONT As String = "Bebas Neue"
Private Const HEADING_SIZE As Long = 18

Private Const BODY_FONT As String = "Lato"
Private Const BODY_SIZE As Long = 12

Private Const PROCESS_INDENT As Long = 36    ' Process Steps helper-column indent (points)
Private Const ROLE_INDENT As Long = 24       ' Roles helper-column indent (points)

' --- headers/footers ---
Private Const HEADER_TEXT As String = "Concrete Results. Civil Solutions."
Private Const LOGO_PATH As String = "C:\Users\OliviaKing\OneDrive - Gregory Construction Services\Desktop\greg_logo.png"  ' update this path
Private Const LOGO_WIDTH_FIRST_IN As Single = 2.4
Private Const LOGO_HEIGHT_FIRST_IN As Single = 0.8
Private Const LOGO_WIDTH_OTHER_IN As Single = 1.5
Private Const LOGO_HEIGHT_OTHER_IN As Single = 0.5

'====================================================
' PUBLIC ENTRY POINTS
'====================================================

Public Sub ValidateSOD()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim issues As Collection

    On Error GoTo ErrHandler

    Set ws = ActiveSheet

    If ws.ListObjects.count = 0 Then
        MsgBox "No table found on active sheet.", vbCritical
        Exit Sub
    End If

    Set tbl = ws.ListObjects(1)
    Set issues = ValidateSheet(tbl)

    If issues.count = 0 Then
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
    
    Dim GroupVal As String

    Dim Response As VbMsgBoxResult

    On Error GoTo ErrHandler

    Set ws = ActiveSheet

    If ws.ListObjects.count = 0 Then
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

    If issues.count > 0 Then
        Response = MsgBox(FormatIssueReport(issues) & vbCrLf & vbCrLf & _
            "Fix these now instead of continuing?", vbYesNo + vbExclamation, "SOD Population")
        If Response = vbYes Then Exit Sub
        ' else: continue best-effort; orphaned/unrecognized columns are skipped, as before
    End If

    ' ---- Determine document title (also used for the output filename) ----
    For TitleCol = 1 To tbl.ListColumns.count
        If IsTitleColumn(tbl.ListColumns(TitleCol).name) Then
            DocTitle = GetDocumentTitle(tbl, TitleCol)
            Exit For
        End If
    Next TitleCol

    If DocTitle = "" Then DocTitle = "Populated SOD"

    ' ---- Fresh blank document with our styles built in ----
    Set wdApp = GetWordApp(weCreatedApp)
    Set wdDoc = wdApp.Documents.Add

    BuildStyles wdDoc
    mShortFormHeaderWritten = False

    ' ---- Title ----
    WriteDocumentTitle wdDoc, DocTitle

    ' ---- Body ----
    col = 1

    Do While col <= tbl.ListColumns.count

        hdr = Trim(tbl.ListColumns(col).name)
        
        Select Case LCase(hdr)
            Case "group"
                GroupVal = GetCombinedText(tbl, col)
        End Select

        If Not IsTableColumn(hdr) And Not IsBulletColumn(hdr) And Not IsTitleColumn(hdr) And Not IsMetaColumn(hdr) Then

            If RenderSpecialSection(wdDoc, tbl, col) Then

                If HasTableColumns(tbl, col) Then
                    col = GetLastTableColumn(tbl, col) + 1

                ElseIf HasBulletColumns(tbl, col) Then
                    col = GetLastBulletColumn(tbl, col) + 1

                Else
                    col = col + 1
                End If

            Else

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

            End If

        Else
            col = col + 1
        End If

    Loop
    
    ' ---- Headers / Footers / Logo ----
    BuildHeadersFooters wdDoc, DocTitle, GroupVal

    ' ---- Finish ----
    wdApp.Visible = True


    Exit Sub

ErrHandler:

    Dim errMsg As String
    errMsg = Err.Description

    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close SaveChanges:=False
    If weCreatedApp Then
        If Not wdApp Is Nothing Then
            If wdApp.Documents.count = 0 Then wdApp.Quit
        End If
    End If
    On Error GoTo 0

    MsgBox errMsg, vbCritical

End Sub

Private Function RenderSpecialSection( _
    ByVal wdDoc As Object, _
    ByVal tbl As ListObject, _
    ByVal colNum As Long) As Boolean

    Select Case LCase(Trim(tbl.ListColumns(colNum).name))

        Case "purpose"
            RenderPurpose wdDoc, tbl, colNum
            RenderSpecialSection = True

        Case "scope"
            RenderScope wdDoc, tbl, colNum
            RenderSpecialSection = True
            
        Case "dictionary"
            RenderDictionary wdDoc, tbl, colNum
            RenderSpecialSection = True

        Case "roles"
            RenderRoles wdDoc, tbl, colNum
            RenderSpecialSection = True

        Case "process steps"
            RenderProcessSteps wdDoc, tbl, colNum
            RenderSpecialSection = True

        Case "key performance indicators"
            RenderKPIs wdDoc, tbl, colNum
            RenderSpecialSection = True

        Case "additional resources"
            RenderAdditionalResources wdDoc, tbl, colNum
            RenderSpecialSection = True

        Case Else
            RenderSpecialSection = False

    End Select

End Function

Private Sub RenderPurpose( _
    ByVal wdDoc As Object, _
    ByVal tbl As ListObject, _
    ByVal colNum As Long)

    If Not mShortFormHeaderWritten Then
        WriteHeading wdDoc, "Executive Overview"
        mShortFormHeaderWritten = True
    End If

    WriteBoldInlineLabel wdDoc, "Purpose", GetCombinedText(tbl, colNum)

End Sub

Private Sub RenderScope( _
    ByVal wdDoc As Object, _
    ByVal tbl As ListObject, _
    ByVal colNum As Long)

    If Not mShortFormHeaderWritten Then
        WriteHeading wdDoc, "Executive Overview"
        mShortFormHeaderWritten = True
    End If

    WriteBoldInlineLabel wdDoc, "Scope", GetCombinedText(tbl, colNum)

End Sub

Private Sub RenderDictionary( _
    ByVal wdDoc As Object, _
    ByVal tbl As ListObject, _
    ByVal colNum As Long)

    Dim firstCol As Long, lastCol As Long
    Dim rowCount As Long, colCount As Long
    Dim wdTable As Object
    Dim r As Long, c As Long

    WriteHeading wdDoc, Trim(tbl.ListColumns(colNum).name)

    firstCol = colNum
    lastCol = firstCol + 1  ' always 2 columns: Term, Definition

    rowCount = LastUsedTableRow(tbl, firstCol, lastCol)
    If rowCount = 0 Then Exit Sub

    colCount = 2

    Set wdTable = wdDoc.Tables.Add(wdDoc.Content.Characters.Last, rowCount, colCount)
    wdTable.Style = "Table Grid"

    For r = 1 To rowCount
        For c = 1 To colCount
            wdTable.cell(r, c).Range.Text = Trim(tbl.DataBodyRange(r, firstCol + c - 1).Text)
            wdTable.cell(r, c).Range.Font.name = "Lato"
            If r = 1 Then
                wdTable.cell(r, c).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                wdTable.cell(r, c).Range.Bold = True
                wdTable.cell(r, c).Range.Font.Color = GREG_YELLOW
                wdTable.cell(r, c).Shading.Texture = 0
                wdTable.cell(r, c).Shading.BackgroundPatternColor = GREG_BLUE
            Else
                wdTable.cell(r, c).Range.Font.Size = 11
                If c = 1 Then
                    wdTable.cell(r, c).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
                End If
            End If
        Next c
    Next r

    ApplyColumnWidths wdDoc, wdTable

End Sub

Private Sub WriteBoldInlineLabel( _
    ByVal wdDoc As Object, _
    ByVal lbl As String, _
    ByVal txt As String)

    Dim p As Object

    wdDoc.Content.InsertAfter lbl & ": " & txt & vbCr

    Set p = wdDoc.Paragraphs(wdDoc.Paragraphs.count - 1).Range

    p.Style = "SOD Body"

    p.Words(1).Bold = True

End Sub

Private Sub RenderRoles( _
    ByVal wdDoc As Object, _
    ByVal tbl As ListObject, _
    ByVal colNum As Long)

    Dim r As Long, c As Long
    Dim role As String
    Dim txt As String
    Dim para As Object
    Dim firstHelperCol As Long, lastHelperCol As Long

    WriteHeading wdDoc, Trim(tbl.ListColumns(colNum).name)

    firstHelperCol = colNum + 1
    lastHelperCol = GetLastBulletColumn(tbl, colNum)

    For r = 1 To tbl.ListRows.count

        role = Trim(tbl.DataBodyRange(r, colNum).Value)

        If role <> "" Then
            wdDoc.Content.InsertAfter role & vbCr
            Set para = wdDoc.Paragraphs(wdDoc.Paragraphs.count - 1).Range
            para.Style = "SOD Body"
            para.ParagraphFormat.SpaceBefore = 8
            para.ParagraphFormat.SpaceAfter = 0
            para.Bold = True
        End If

        ' Render all helper columns as indented lines
        For c = firstHelperCol To lastHelperCol
            If c <= tbl.ListColumns.count Then
                txt = Trim(tbl.DataBodyRange(r, c).Value)
                If txt <> "" Then
                    wdDoc.Content.InsertAfter txt & vbCr
                    Set para = wdDoc.Paragraphs(wdDoc.Paragraphs.count - 1).Range
                    para.Style = "SOD Body"
                    para.ParagraphFormat.LeftIndent = ROLE_INDENT
                End If
            End If
        Next c

    Next r

End Sub

Private Sub RenderProcessSteps( _
    ByVal wdDoc As Object, _
    ByVal tbl As ListObject, _
    ByVal colNum As Long)

    Dim r As Long, c As Long
    Dim stepText As String
    Dim txt As String
    Dim rng As Object
    Dim firstHelperCol As Long, lastHelperCol As Long

    WriteHeading wdDoc, Trim(tbl.ListColumns(colNum).name)

    firstHelperCol = colNum + 1
    lastHelperCol = GetLastBulletColumn(tbl, colNum)

    For r = 1 To tbl.ListRows.count

        stepText = Trim(tbl.DataBodyRange(r, colNum).Value)

        If stepText <> "" Then
            ' Add colon to the step text before inserting; results in "1. Project setup:"
            wdDoc.Content.InsertAfter stepText & ":" & vbCr
            Set rng = wdDoc.Paragraphs(wdDoc.Paragraphs.count - 1).Range
            rng.Style = "SOD Body"
            rng.Bold = True
            rng.ParagraphFormat.SpaceBefore = 8
            rng.ParagraphFormat.SpaceAfter = 0
            rng.ListFormat.ApplyNumberDefault
        End If

        ' Render all helper columns as indented lines (no bullet/number)
        For c = firstHelperCol To lastHelperCol
            If c <= tbl.ListColumns.count Then
                txt = Trim(tbl.DataBodyRange(r, c).Value)
                If txt <> "" Then
                    wdDoc.Content.InsertAfter txt & vbCr
                    Set rng = wdDoc.Paragraphs(wdDoc.Paragraphs.count - 1).Range
                    rng.Style = "SOD Body"
                    rng.ParagraphFormat.LeftIndent = PROCESS_INDENT
                End If
            End If
        Next c

    Next r

End Sub

Private Sub RenderKPIs( _
    ByVal wdDoc As Object, _
    ByVal tbl As ListObject, _
    ByVal colNum As Long)

    Dim r As Long
    Dim txt As String
    Dim items As Collection
    Dim rng As Object

    WriteHeading wdDoc, Trim(tbl.ListColumns(colNum).name)

    ' Collect all non-empty items
    Set items = New Collection
    For r = 1 To tbl.ListRows.count
        txt = Trim(tbl.DataBodyRange(r, colNum).Value)
        If txt <> "" Then
            items.Add txt
        End If
    Next r

    If items.count = 0 Then Exit Sub

    ' Insert opening section break for 2-column layout
    Set rng = wdDoc.Content
    rng.Collapse 0
    rng.InsertBreak Type:=wdSectionBreakContinuous

    ' Set the new section to 2 columns
    With wdDoc.Sections(wdDoc.Sections.count).PageSetup.TextColumns
        .SetCount 2
        .EvenlySpaced = True
        .LineBetween = False
    End With

    ' Write items as bullets
    WriteBulletList wdDoc, items

    ' Insert closing section break to return to 1 column
    Set rng = wdDoc.Content
    rng.Collapse 0
    rng.InsertBreak Type:=wdSectionBreakContinuous
    wdDoc.Sections(wdDoc.Sections.count).PageSetup.TextColumns.SetCount 1

End Sub

Private Sub RenderAdditionalResources( _
    ByVal wdDoc As Object, _
    ByVal tbl As ListObject, _
    ByVal colNum As Long)

    Dim firstCol As Long, lastCol As Long
    Dim rowCount As Long, colCount As Long
    Dim wdTable As Object
    Dim r As Long, c As Long
    Dim rng As Object

    WriteHeading wdDoc, Trim(tbl.ListColumns(colNum).name)

    ' Determine table structure (treat as a TableN column group)
    firstCol = colNum
    lastCol = GetLastTableColumn(tbl, colNum)

    ' If no explicit Table columns follow, treat this column alone as the table
    If lastCol < colNum Then
        lastCol = colNum
    End If

    rowCount = LastUsedTableRow(tbl, firstCol, lastCol)

    If rowCount = 0 Then Exit Sub

    colCount = lastCol - firstCol + 1

    ' Create the table
    Set wdTable = wdDoc.Tables.Add(wdDoc.Content.Characters.Last, rowCount, colCount)
    wdTable.Style = "Table Grid"
    
    For r = 1 To rowCount
        For c = 1 To colCount
            wdTable.cell(r, c).Range.Text = Trim(tbl.DataBodyRange(r, firstCol + c - 1).Text)
            wdTable.cell(r, c).Range.Font.name = "Lato"
            If r = 1 Then
                wdTable.cell(r, c).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                wdTable.cell(r, c).Range.Bold = True
                wdTable.cell(r, c).Range.Font.Color = GREG_YELLOW
                wdTable.cell(r, c).Shading.Texture = 0
                wdTable.cell(r, c).Shading.BackgroundPatternColor = GREG_BLUE
            Else
                wdTable.cell(r, c).Range.Font.Size = 11
                If c = 1 Then
                    Set rng = wdTable.cell(r, 1).Range
                    rng.ListFormat.ApplyListTemplateWithLevel _
                        ListTemplate:=rng.Application.ListGalleries(2).ListTemplates(1), _
                        ContinuePreviousList:=(r > 2), _
                        ApplyTo:=wdListApplyToWholeList, _
                        DefaultListBehavior:=wdWord10ListBehavior
                    rng.ParagraphFormat.LeftIndent = 16
                End If
            End If
        Next c
    Next r

    wdDoc.Content.InsertAfter vbCr

End Sub

Private Function GetCombinedText( _
    ByVal tbl As ListObject, _
    ByVal colNum As Long) As String

    Dim r As Long

    For r = 1 To tbl.ListRows.count

        If Trim(tbl.DataBodyRange(r, colNum).Value) <> "" Then

            If Len(GetCombinedText) > 0 Then

                GetCombinedText = _
                    GetCombinedText & " "

            End If

            GetCombinedText = _
                GetCombinedText & _
                Trim(tbl.DataBodyRange(r, colNum).Value)

        End If

    Next r

End Function

'====================================================
' VALIDATION
'====================================================

Private Function ValidateSheet(ByVal tbl As ListObject) As Collection

    Dim issues As New Collection
    Dim col As Long
    Dim hdr As String
    Dim titleCount As Long
    Dim expectingGroup As String  ' "", "PENDING", "TABLE", or "BULLET"

    If tbl.ListColumns.count = 0 Then
        issues.Add SEV_FATAL & "|(sheet)|Table has no columns."
        Set ValidateSheet = issues
        Exit Function
    End If

    If tbl.ListRows.count = 0 Then
        issues.Add SEV_FATAL & "|(sheet)|Table has no data rows."
        Set ValidateSheet = issues
        Exit Function
    End If

    expectingGroup = ""

    For col = 1 To tbl.ListColumns.count

        hdr = Trim(tbl.ListColumns(col).name)

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
    End If

    wdApp.Visible = True

    Set GetWordApp = wdApp

End Function

Private Sub BuildStyles(ByVal wdDoc As Object)

    Dim s As Object

    Set s = AddOrGetStyle(wdDoc, "SOD Title")
    With s.Font
        .name = TITLE_FONT
        .Size = TITLE_SIZE
        .Color = GREG_BLUE
        .Spacing = 1.5
    End With
    s.ParagraphFormat.SpaceBefore = 10
    s.ParagraphFormat.SpaceAfter = 0
    s.ParagraphFormat.Alignment = wdAlignParagraphCenter
    s.ParagraphFormat.LineSpacingRule = 0  ' wdLineSpaceSingle

    Set s = AddOrGetStyle(wdDoc, "SOD Subtitle")
    With s.Font
        .name = SUBTITLE_FONT
        .Size = SUBTITLE_SIZE
        .Bold = True
        .Allcaps = True
        .Spacing = 1.5  ' Extended letter spacing
    End With
    s.ParagraphFormat.SpaceBefore = 0
    s.ParagraphFormat.SpaceAfter = 10
    s.ParagraphFormat.Alignment = wdAlignParagraphCenter
    s.ParagraphFormat.LineSpacingRule = 0  ' wdLineSpaceSingle

    Set s = AddOrGetStyle(wdDoc, "SOD Heading")
    With s.Font
        .name = HEADING_FONT
        .Size = HEADING_SIZE
        .Color = GREG_BLUE
        .Spacing = 1.5
        .Underline = 1
        .UnderlineColor = GREG_YELLOW
    End With
    s.ParagraphFormat.SpaceBefore = 18
    s.ParagraphFormat.SpaceAfter = 0
    s.ParagraphFormat.Alignment = wdAlignParagraphCenter
    s.ParagraphFormat.KeepWithNext = True

    Set s = AddOrGetStyle(wdDoc, "SOD Body")
    With s.Font
        .name = BODY_FONT
        .Size = BODY_SIZE
    End With
    s.ParagraphFormat.SpaceBefore = 0
    s.ParagraphFormat.SpaceAfter = 6

End Sub

Private Function AddOrGetStyle(ByVal wdDoc As Object, ByVal styleName As String) As Object

    Dim s As Object

    On Error Resume Next
    Set s = wdDoc.Styles(styleName)
    On Error GoTo 0

    If s Is Nothing Then
        Set s = wdDoc.Styles.Add(name:=styleName, Type:=wdStyleTypeParagraph)
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

Private Sub WriteDocumentTitle(ByVal wdDoc As Object, ByVal titleText As String)

    Dim tbl As Object
    Dim cellRng As Object
    Dim titleRange As Object, subtitleRange As Object

    ' Create a 2-row, 1-column table for title banner
    Set tbl = wdDoc.Tables.Add(wdDoc.Content.Characters.First, 2, 1)

    ' Set table width to full page
    tbl.PreferredWidthType = 2  ' wdPreferredWidthPercent
    tbl.PreferredWidth = 100

    ' Remove all borders
    With tbl.Borders
        .InsideLineStyle = 0  ' wdLineStyleNone
        .OutsideLineStyle = 0  ' wdLineStyleNone
    End With

    ' Add only top and bottom borders
    tbl.Borders(1).LineStyle = 1   ' wdLineStyleSingle (top)
    tbl.Borders(3).LineStyle = 1   ' wdLineStyleSingle (bottom)
    tbl.Borders(1).LineWidth = 8   ' wdLineWidth100pt
    tbl.Borders(3).LineWidth = 8   ' wdLineWidth100pt

    ' Row 1: Title (no trailing newline)
    Set cellRng = tbl.cell(1, 1).Range
    cellRng.InsertAfter Trim(titleText)
    Set titleRange = cellRng
    titleRange.Style = "SOD Title"

    ' Row 2: Subtitle (no trailing newline)
    Set cellRng = tbl.cell(2, 1).Range
    cellRng.InsertAfter DOCUMENT_SUBTITLE
    Set subtitleRange = cellRng
    subtitleRange.Style = "SOD Subtitle"

End Sub

Private Sub WriteHeading(ByVal wdDoc As Object, ByVal txt As String)

    wdDoc.Content.InsertAfter txt & vbCr
    wdDoc.Paragraphs(wdDoc.Paragraphs.count - 1).Range.Style = "SOD Heading"

End Sub

Private Sub WriteParagraph(ByVal wdDoc As Object, ByVal txt As String)

    wdDoc.Content.InsertAfter txt & vbCr
    wdDoc.Paragraphs(wdDoc.Paragraphs.count - 1).Range.Style = "SOD Body"

End Sub

'====================================================
' HEADERS / FOOTERS / LOGO
'====================================================

Private Sub BuildHeadersFooters(ByVal wdDoc As Object, ByVal DocTitle As String, _
                                ByVal GroupVal As String)

    Dim sec As Object
    Dim hdr As Object
    Dim ftr As Object
    Dim rng As Object
    Dim tbl As Object
    Dim rightRng As Object
    Dim todayStr As String
    Dim usableWidth As Single
    Dim logoShape As Object

    todayStr = Format(Now, "mm/dd/yyyy")

    usableWidth = wdDoc.PageSetup.PageWidth _
                - wdDoc.PageSetup.LeftMargin _
                - wdDoc.PageSetup.RightMargin

    Set sec = wdDoc.Sections(1)
    sec.PageSetup.DifferentFirstPageHeaderFooter = True

    '======================
    ' FIRST PAGE HEADER
    '======================
    Set hdr = sec.headers(wdHeaderFooterFirstPage)
    Set rng = hdr.Range
    rng.Text = ""

    If LOGO_PATH <> "" And Dir(LOGO_PATH) <> "" Then

        Set logoShape = hdr.Range.InlineShapes.AddPicture( _
            fileName:=LOGO_PATH, _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Range:=hdr.Range)

        logoShape.Width = wdDoc.Application.InchesToPoints(LOGO_WIDTH_FIRST_IN)
        logoShape.Height = wdDoc.Application.InchesToPoints(LOGO_HEIGHT_FIRST_IN)
        logoShape.LockAspectRatio = False

    End If

    ' Tagline always renders, on its own line after the logo (or alone if no logo)
    hdr.Range.InsertAfter vbCr & HEADER_TEXT

    ' Logo paragraph: centered, no spacing
    With hdr.Range.Paragraphs(1).Range.ParagraphFormat
        .Alignment = wdAlignParagraphCenter
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = 0
    End With

    ' Tagline paragraph: centered, 6pt after
    With hdr.Range.Paragraphs.Last.Range
        .Font.name = "Bebas Neue"
        .Font.Size = 12.5
        .Font.Spacing = 1
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 6
        .ParagraphFormat.LineSpacingRule = 0
    End With

    '======================
    ' FIRST PAGE FOOTER - 2-row, 2-col borderless table (65/35)
    '======================
    Set ftr = sec.Footers(wdHeaderFooterFirstPage)
    Set rng = ftr.Range
    rng.Text = ""

    Set tbl = ftr.Range.Tables.Add(ftr.Range, 2, 2)

    With tbl.Borders
        .InsideLineStyle = 0
        .OutsideLineStyle = 0
    End With

    tbl.PreferredWidthType = 2
    tbl.PreferredWidth = 100
    tbl.Columns(1).SetWidth usableWidth * 0.75, 0
    tbl.Columns(2).SetWidth usableWidth * 0.25, 0

    With tbl.cell(1, 1).Range
        .Text = "Group: " & GroupVal
        .Font.name = "Lato"
        .Font.Color = GREG_GRAY
        .Font.Size = 10
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
    End With

    With tbl.cell(2, 2).Range
        .Text = "Page 1"
        .Font.name = "Lato"
        .Font.Color = GREG_GRAY
        .Font.Size = 10
        .ParagraphFormat.Alignment = wdAlignParagraphRight
    End With

    With tbl.cell(2, 1).Range
        .Text = DocTitle 'Later might be Process Number
        .Font.name = "Lato"
        .Font.Color = GREG_GRAY
        .Font.Size = 10
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
    End With

    With tbl.cell(1, 2).Range
        .Text = "Updated: " & todayStr
        .Font.name = "Lato"
        .Font.Color = GREG_GRAY
        .Font.Size = 10
        .ParagraphFormat.Alignment = wdAlignParagraphRight
    End With

    '======================
    ' ALL OTHER PAGES HEADER
    '======================
    Set hdr = sec.headers(wdHeaderFooterPrimary)
    Set rng = hdr.Range
    rng.Text = ""

    If LOGO_PATH <> "" And Dir(LOGO_PATH) <> "" Then

        Set logoShape = hdr.Range.InlineShapes.AddPicture( _
            fileName:=LOGO_PATH, _
            LinkToFile:=False, _
            SaveWithDocument:=True, _
            Range:=hdr.Range)

        logoShape.Width = wdDoc.Application.InchesToPoints(LOGO_WIDTH_OTHER_IN)
        logoShape.Height = wdDoc.Application.InchesToPoints(LOGO_HEIGHT_OTHER_IN)
        logoShape.LockAspectRatio = False

    End If

    With hdr.Range.ParagraphFormat
        .Alignment = wdAlignParagraphCenter
        .SpaceBefore = 0
        .SpaceAfter = 6
        .LineSpacingRule = 0
    End With

    '======================
    ' ALL OTHER PAGES FOOTER - 1-row, 2-col borderless table (90/10)
    '======================
    Set ftr = sec.Footers(wdHeaderFooterPrimary)
    Set rng = ftr.Range
    rng.Text = ""

    Set tbl = ftr.Range.Tables.Add(ftr.Range, 1, 2)

    With tbl.Borders
        .InsideLineStyle = 0
        .OutsideLineStyle = 0
    End With

    tbl.PreferredWidthType = 2
    tbl.PreferredWidth = 100
    tbl.Columns(1).SetWidth usableWidth * 0.9, 0
    tbl.Columns(2).SetWidth usableWidth * 0.1, 0

    With tbl.cell(1, 1).Range
        .Text = DocTitle
        .Font.name = "Lato"
        .Font.Color = GREG_GRAY
        .Font.Size = 10
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
    End With

    Set rightRng = tbl.cell(1, 2).Range
    rightRng.Collapse 1   ' wdCollapseStart
    rightRng.InsertAfter "Page "
    rightRng.Collapse 0
    rightRng.Fields.Add rightRng, wdFieldPage

    tbl.cell(1, 2).Range.Font.name = "Lato"
    tbl.cell(1, 2).Range.Font.Size = 10
    tbl.cell(1, 2).Range.Font.Color = GREG_GRAY
    tbl.cell(1, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphRight

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
    lastRow = tbl.ListRows.count

    For r = 1 To lastRow + 1

        If r <= lastRow Then

            If Trim(tbl.DataBodyRange(r, colNum).Value) <> "" Then
                block.Add Trim(tbl.DataBodyRange(r, colNum).Value)
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

    If block.count = 0 Then Exit Sub

    If block.count = 1 Then
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
            ListTemplate:=rng.Application.ListGalleries(wdBulletGallery).ListTemplates(3), _
            ContinuePreviousList:=False, _
            ApplyTo:=wdListApplyToWholeList, _
            DefaultListBehavior:=wdWord10ListBehavior
        .ListLevelNumber = level
    End With

    ' Set bullet color via the list level's font
    rng.ListFormat.ListTemplate.ListLevels(level).Font.Color = GREG_YELLOW

End Sub

Private Sub WriteBulletList(ByVal wdDoc As Object, ByVal items As Collection)

    Dim i As Long
    Dim startPara As Long
    Dim rng As Object

    startPara = wdDoc.Paragraphs.count

    For i = 1 To items.count
        wdDoc.Content.InsertAfter items(i) & vbCr
    Next i

    Set rng = wdDoc.Range( _
        Start:=wdDoc.Paragraphs(startPara).Range.Start, _
        End:=wdDoc.Paragraphs(wdDoc.Paragraphs.count - 1).Range.End)

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
    lastRow = tbl.ListRows.count

    For r = 1 To lastRow
        For c = parentCol To bulletLastCol

            txt = Trim(tbl.DataBodyRange(r, c).Value)

            If txt <> "" Then
                wdDoc.Content.InsertAfter txt & vbCr
                level = c - parentCol + 1
                Set rng = wdDoc.Paragraphs(wdDoc.Paragraphs.count - 1).Range
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
    Dim startRow As Long
    Dim hasHeader As Boolean

    firstCol = parentCol
    lastCol = GetLastTableColumn(tbl, parentCol)
    rowCount = LastUsedTableRow(tbl, firstCol, lastCol)

    If rowCount = 0 Then Exit Sub

    colCount = lastCol - firstCol + 1

    ' A blank row 1 means "no header" - skip it entirely rather than
    ' inserting an empty formatted header row into the Word table.
    hasHeader = False
    For c = firstCol To lastCol
        If Trim(tbl.DataBodyRange(1, c).Value) <> "" Then
            hasHeader = True
            Exit For
        End If
    Next c

    If hasHeader Then
        startRow = 1
    Else
        startRow = 2
    End If

    If startRow > rowCount Then Exit Sub

    Dim wordRowCount As Long
    wordRowCount = rowCount - startRow + 1

    Set wdTable = wdDoc.Tables.Add(wdDoc.Content.Characters.Last, wordRowCount, colCount)

    Dim wordR As Long
    wordR = 0

    For r = startRow To rowCount

        wordR = wordR + 1

        For c = 1 To colCount
            wdTable.cell(wordR, c).Range.Text = Trim(tbl.DataBodyRange(r, firstCol + c - 1).Text)
            wdTable.cell(wordR, c).Range.Font.name = "Lato"

            If hasHeader And wordR = 1 Then
                wdTable.cell(wordR, c).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                wdTable.cell(wordR, c).Range.Bold = True
            End If
        Next c

    Next r

    wdTable.Style = "Table Grid"
    ApplyColumnWidths wdDoc, wdTable

    wdDoc.Content.InsertAfter vbCr

End Sub

Private Sub ApplyColumnWidths(ByVal wdDoc As Object, ByVal wdTable As Object)

    Dim totalWidth As Single
    Dim firstColWidth As Single
    Dim remainingWidth As Single
    Dim colCount As Long
    Dim c As Long

    colCount = wdTable.Columns.count

    If colCount = 0 Then Exit Sub

    totalWidth = wdDoc.PageSetup.PageWidth _
               - wdDoc.PageSetup.LeftMargin _
               - wdDoc.PageSetup.RightMargin

    ' Fit first column to content, cap at 40% of total width
    wdTable.Columns(1).AutoFit
    firstColWidth = wdTable.Columns(1).Width
    If firstColWidth > totalWidth * 0.4 Then
        firstColWidth = totalWidth * 0.4
        wdTable.Columns(1).SetWidth firstColWidth, 0
    End If

    ' Distribute remaining width evenly across other columns
    If colCount > 1 Then
        remainingWidth = (totalWidth - firstColWidth) / (colCount - 1)
        For c = 2 To colCount
            wdTable.Columns(c).SetWidth remainingWidth, 0
        Next c
    End If

End Sub

Private Function LastUsedTableRow(ByVal tbl As ListObject, _
                                  ByVal firstCol As Long, _
                                  ByVal lastCol As Long) As Long

    Dim r As Long, c As Long

    For r = tbl.ListRows.count To 1 Step -1
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
    If colNum >= tbl.ListColumns.count Then Exit Function
    HasTableColumns = IsTableColumn(tbl.ListColumns(colNum + 1).name)
End Function

Private Function HasBulletColumns(ByVal tbl As ListObject, ByVal colNum As Long) As Boolean
    If colNum >= tbl.ListColumns.count Then Exit Function
    HasBulletColumns = IsBulletColumn(tbl.ListColumns(colNum + 1).name)
End Function

Private Function GetLastTableColumn(ByVal tbl As ListObject, ByVal colNum As Long) As Long
    Dim c As Long
    c = colNum + 1
    Do While c <= tbl.ListColumns.count
        If Not IsTableColumn(tbl.ListColumns(c).name) Then Exit Do
        c = c + 1
    Loop
    GetLastTableColumn = c - 1
End Function

Private Function GetLastBulletColumn(ByVal tbl As ListObject, ByVal colNum As Long) As Long
    Dim c As Long
    c = colNum + 1
    Do While c <= tbl.ListColumns.count
        If Not IsBulletColumn(tbl.ListColumns(c).name) Then Exit Do
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

Private Function IsMetaColumn(ByVal txt As String) As Boolean
    Select Case LCase(Trim(txt))
        Case "title", "group"
            IsMetaColumn = True
    End Select
End Function

Private Function GetDocumentTitle(ByVal tbl As ListObject, ByVal TitleCol As Long) As String

    Dim r As Long

    For r = 1 To tbl.ListRows.count
        If Trim(tbl.DataBodyRange(r, TitleCol).Value) <> "" Then
            GetDocumentTitle = Trim(tbl.DataBodyRange(r, TitleCol).Value)
            Exit Function
        End If
    Next r

    GetDocumentTitle = "Populated SOD"

End Function

