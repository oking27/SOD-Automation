Attribute VB_Name = "Module1"
Sub GenerateSOP_Final_Clean()

    Dim wdApp As Object
    Dim wdDoc As Object
    Dim ws As Worksheet
    Dim templatePath As String
    Dim outputFolder As String
    Dim outputPath As String
    Dim fileName As String
    Dim cc As Object
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, col As Long
    Dim headerRow As Long
    Dim tbl As Object
    Dim defIndex As Long
    Dim para As Object

    Set ws = ActiveSheet

    ' =======================
    ' BUILD HEADER MAP ?
    ' =======================
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For col = 1 To lastCol
        If ws.Cells(1, col).Value <> "" Then
            colMap(Trim(ws.Cells(1, col).Value)) = col
        End If
    Next col

    ' =======================
    ' PATHS
    ' =======================
    templatePath = "C:\Users\OliviaKing\OneDrive - Gregory Construction Services\Desktop\SOD_Desktop.docx"
    outputFolder = "C:\Users\OliviaKing\OneDrive - Gregory Construction Services\Desktop\SOP_Output\"

    If Dir(outputFolder, vbDirectory) = "" Then MkDir outputFolder

    ' =======================
    ' FIND HEADER ROW
    ' =======================
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    headerRow = 0

    For i = 2 To lastRow
        If ws.Cells(i, colMap("title")).Value <> "" Then
            headerRow = i
            Exit For
        End If
    Next i

    If headerRow = 0 Then
        MsgBox "No SOP title found"
        Exit Sub
    End If

    ' =======================
    ' FILE NAME
    ' =======================
    fileName = ws.Cells(headerRow, colMap("title")).Value
    fileName = Replace(fileName, "/", "-")
    fileName = Replace(fileName, "\", "-")
    fileName = Replace(fileName, ":", "-")

    outputPath = outputFolder & fileName & ".docx"

    ' =======================
    ' OPEN WORD
    ' =======================
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open(templatePath, ReadOnly:=True, AddToRecentFiles:=False)

    ' =======================
    ' VERSION
    ' =======================
    Dim verRange As Object
    On Error Resume Next
    Set verRange = wdDoc.Bookmarks("SOD_Version").Range
    On Error GoTo 0

    If Not verRange Is Nothing Then
        verRange.Text = "Version: 1.0"
    End If

    ' =======================
    ' CONTENT CONTROLS
    ' =======================
    Dim storyRange As Object

    For Each storyRange In wdDoc.StoryRanges
        Do
            For Each cc In storyRange.ContentControls
    
                Select Case cc.Title
    
                    Case "sod_title"
                        cc.Range.Text = ""
                        cc.Range.InsertAfter ws.Cells(headerRow, colMap("title")).Value
    
                    Case "sod_policy_statement"
                        cc.Range.Text = ""
                        cc.Range.InsertAfter ws.Cells(headerRow, colMap("policy statement")).Value
    
                    Case "sod_purpose"
                        cc.Range.Text = ""
                        cc.Range.InsertAfter ws.Cells(headerRow, colMap("purpose")).Value
    
                    Case "sod_scope"
                        cc.Range.Text = ""
                        cc.Range.InsertAfter ws.Cells(headerRow, colMap("scope")).Value
    
                    Case "sod_resource"
                        cc.Range.Text = ""
                        cc.Range.InsertAfter ws.Cells(headerRow, colMap("resource")).Value
                        
                    Case "sod_version"
                        cc.Range.Text = ""
                        cc.Range.InsertAfter ws.Cells(headerRow, colMap("version")).Value
    
                End Select
    
            Next cc
    
            Set storyRange = storyRange.NextStoryRange
        Loop Until storyRange Is Nothing
    Next storyRange


    ' =======================
    ' DEFINITIONS TABLE
    ' =======================
    On Error Resume Next
    Set tbl = wdDoc.Bookmarks("DefinitionsTable").Range.Tables(1)
    On Error GoTo 0

    If tbl Is Nothing Then Set tbl = wdDoc.Tables(1)

    defIndex = 0

    For i = 2 To lastRow

        If ws.Cells(i, colMap("term")).Value <> "" Then

            defIndex = defIndex + 1

            If defIndex = 1 Then
                If tbl.Rows.Count < 2 Then tbl.Rows.Add
                With tbl.Rows(2)
                    .Cells(1).Range.Text = ws.Cells(i, colMap("term")).Value
                    .Cells(2).Range.Text = ws.Cells(i, colMap("definition")).Value
                End With
            Else
                tbl.Rows.Add
                With tbl.Rows(tbl.Rows.Count)
                    .Cells(1).Range.Text = ws.Cells(i, colMap("term")).Value
                    .Cells(2).Range.Text = ws.Cells(i, colMap("definition")).Value
                End With
            End If

        End If

    Next i

    ' =======================
    ' ROLES
    ' =======================
    Dim rrRange As Object
    Dim startPos As Long
    Dim boldRange As Object

    On Error Resume Next
    Set rrRange = wdDoc.Bookmarks("RolesResponsibilities").Range
    On Error GoTo 0

    If Not rrRange Is Nothing Then

        rrRange.Text = ""

        For i = 2 To lastRow

            If ws.Cells(i, colMap("roles")).Value <> "" Then

                startPos = rrRange.End

                rrRange.InsertAfter ws.Cells(i, colMap("roles")).Value & ": " & _
                                    ws.Cells(i, colMap("responsibilities")).Value & vbCr

                Set boldRange = wdDoc.Range(startPos, startPos + Len(ws.Cells(i, colMap("roles")).Value) + 1)
                boldRange.Bold = True

            End If

        Next i

    End If

    ' =======================
    ' PROCESS OBJECTIVES
    ' =======================
    Dim objRange As Object
    On Error Resume Next
    Set objRange = wdDoc.Bookmarks("ProcessObjectives").Range
    On Error GoTo 0

    If Not objRange Is Nothing Then

        objRange.Text = ""

        For i = 2 To lastRow
            If Trim(ws.Cells(i, colMap("process objectives")).Value) <> "" Then
                objRange.InsertAfter ws.Cells(i, colMap("process objectives")).Value & vbCr
            End If
        Next i

        For Each para In objRange.Paragraphs
            para.Range.ListFormat.ApplyBulletDefault
            para.SpaceBefore = 2
            para.SpaceAfter = 2
            para.LineSpacingRule = 1
            para.LineSpacing = 13.8
        Next para

    End If

    ' =======================
    ' PROCESS STEPS
    ' =======================
    Dim stepRange As Object
    Dim stepStart As Long
    Dim mainRange As Object

    On Error Resume Next
    Set stepRange = wdDoc.Bookmarks("ProcessSteps").Range
    On Error GoTo 0

    If Not stepRange Is Nothing Then

        stepRange.Text = ""

        For i = 2 To lastRow

            If ws.Cells(i, colMap("process steps")).Value <> "" Then

                stepStart = stepRange.End
                stepRange.InsertAfter ws.Cells(i, colMap("process steps")).Value & vbCr

                Set para = stepRange.Paragraphs(stepRange.Paragraphs.Count)
                para.Range.ListFormat.ApplyNumberDefault

                para.SpaceBefore = 8
                para.SpaceAfter = 2
                para.LineSpacingRule = 1
                para.LineSpacing = 13.8

                Set mainRange = wdDoc.Range(stepStart, para.Range.End)
                mainRange.Bold = True
                mainRange.Underline = True

            End If

            If ws.Cells(i, colMap("step steps")).Value <> "" Then

                stepRange.InsertAfter ws.Cells(i, colMap("step steps")).Value & vbCr

                Set para = stepRange.Paragraphs(stepRange.Paragraphs.Count)
                para.Range.ListFormat.ApplyBulletDefault

                para.LeftIndent = 54
                para.FirstLineIndent = -18

                para.SpaceBefore = 2
                para.SpaceAfter = 2
                para.LineSpacingRule = 1
                para.LineSpacing = 13.8

            End If

        Next i

    End If

    ' =======================
    ' PERFORMANCE OUTPUT
    ' =======================
    Dim outRange As Object
    On Error Resume Next
    Set outRange = wdDoc.Bookmarks("PerformanceOutput").Range
    On Error GoTo 0

    If Not outRange Is Nothing Then

        outRange.Text = ""

        For i = 2 To lastRow
            If Trim(ws.Cells(i, colMap("performance output")).Value) <> "" Then
                outRange.InsertAfter ws.Cells(i, colMap("performance output")).Value & vbCr
            End If
        Next i

        For Each para In outRange.Paragraphs
            para.Range.ListFormat.ApplyBulletDefault
            para.SpaceBefore = 2
            para.SpaceAfter = 2
            para.LineSpacingRule = 1
            para.LineSpacing = 13.8
        Next para

    End If

    ' =======================
    ' KPIs
    ' =======================
    Dim kpiRange As Object
    On Error Resume Next
    Set kpiRange = wdDoc.Bookmarks("KPIs").Range
    On Error GoTo 0

    If Not kpiRange Is Nothing Then

        kpiRange.Text = ""

        For i = 2 To lastRow
            If Trim(ws.Cells(i, colMap("kpis")).Value) <> "" Then
                kpiRange.InsertAfter ws.Cells(i, colMap("kpis")).Value & vbCr
            End If
        Next i

        For Each para In kpiRange.Paragraphs
            para.Range.ListFormat.ApplyBulletDefault
            para.SpaceBefore = 2
            para.SpaceAfter = 2
            para.LineSpacingRule = 1
            para.LineSpacing = 13.8
        Next para

    End If

    ' =======================
    ' SAVE
    ' =======================
    wdDoc.SaveAs2 outputPath

    wdDoc.Close
    wdApp.Quit

    MsgBox "SOP generated successfully:" & vbCrLf & outputPath

End Sub


