Attribute VB_Name = "Module5"
Option Explicit

'====================================================
' SOD GENERATOR
'====================================================
Private Const DOCUMENT_SUBTITLE As String = _
    "Standard Operating Document"

Public Sub GenerateSODs()

    Dim Response As VbMsgBoxResult
    
    Dim ws As Worksheet
    Dim tbl As ListObject

    Dim wdApp As Object
    Dim wdDoc As Object

    Dim col As Long
    Dim hdr As String

    On Error GoTo ErrHandler

    Set ws = ActiveSheet

    If ws.ListObjects.Count = 0 Then
        MsgBox "No table found on active sheet.", vbCritical
        Exit Sub
    End If

    Set tbl = ws.ListObjects(1)

    Set wdApp = GetWordApp()
    Set wdDoc = wdApp.Documents.Add
    
    Dim TitleCol As Long
    Dim DocTitle As String
    
    For TitleCol = 1 To tbl.ListColumns.Count
    
        If IsTitleColumn(tbl.ListColumns(TitleCol).Name) Then
    
            DocTitle = GetDocumentTitle(tbl, TitleCol)
    
            WriteDocumentTitle wdDoc, DocTitle
    
            Exit For
    
        End If
    
    Next TitleCol

    col = 1

    Do While col <= tbl.ListColumns.Count

        hdr = Trim(tbl.ListColumns(col).Name)

        If Not IsTableColumn(hdr) _
        And Not IsBulletColumn(hdr) _
        And Not IsTitleColumn(hdr) Then

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
    
        Response = MsgBox( _
            "SOD generated successfully." & vbCrLf & vbCrLf & _
            "Would you like to view the document?", _
            vbYesNo + vbQuestion, _
            "SOD Generator")
        
        If Response = vbYes Then
        
            wdApp.Visible = True
        
        Else
        
            wdApp.Visible = True
        
            wdDoc.Close SaveChanges:=True
        
            If wdApp.Documents.Count = 0 Then
                wdApp.Quit
            End If
        
        End If
    
    Exit Sub

ErrHandler:

    MsgBox Err.Description, vbCritical

End Sub


'====================================================
' WORD
'====================================================

Private Function GetWordApp() As Object

    Dim wdApp As Object

    On Error Resume Next

    Set wdApp = GetObject(, "Word.Application")

    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
    End If

    On Error GoTo 0

    wdApp.Visible = False

    Set GetWordApp = wdApp

End Function

'====================================================
' SECTION HEADINGS
'====================================================

Private Sub WriteHeading(ByVal wdDoc As Object, _
                         ByVal txt As String)

    Dim rng As Object

    Set rng = wdDoc.Content
    rng.Collapse 0

    rng.InsertAfter txt & vbCrLf

    rng.SetRange rng.End - Len(txt) - 1, rng.End - 1

    FormatHeading rng

    wdDoc.Content.Collapse 0
    wdDoc.Content.InsertAfter vbCrLf

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

Private Sub RenderBlock(ByVal wdDoc As Object, _
                        ByVal block As Collection)

    Dim i As Long

    If block.Count = 0 Then Exit Sub

    If block.Count = 1 Then

        WriteParagraph wdDoc, block(1)

    Else

        WriteBulletList wdDoc, block

    End If

End Sub

'====================================================
' PARAGRAPHS
'====================================================

Private Sub WriteParagraph(ByVal wdDoc As Object, _
                           ByVal txt As String)

    wdDoc.Content.Collapse 0
    wdDoc.Content.InsertAfter txt & vbCrLf & vbCrLf

End Sub

'====================================================
' BULLETS
'====================================================

Private Sub WriteBulletList(ByVal wdDoc As Object, _
                            ByVal items As Collection)

    Dim i As Long

    For i = 1 To items.Count

        wdDoc.Content.Collapse 0
        wdDoc.Content.InsertAfter "• " & items(i) & vbCrLf

    Next i

    wdDoc.Content.InsertAfter vbCrLf

End Sub

'====================================================
' TABLES
'====================================================

Private Sub CreateTableSection(ByVal wdDoc As Object, _
                               ByVal tbl As ListObject, _
                               ByVal parentCol As Long)

    Dim firstCol As Long
    Dim lastCol As Long

    Dim rowCount As Long
    Dim colCount As Long

    Dim wdTable As Object

    Dim r As Long
    Dim c As Long

    firstCol = parentCol
    lastCol = GetLastTableColumn(tbl, parentCol)

    rowCount = LastUsedTableRow(tbl, firstCol, lastCol)

    If rowCount = 0 Then Exit Sub

    colCount = lastCol - firstCol + 1

    Set wdTable = wdDoc.Tables.Add( _
        wdDoc.Content.Characters.Last, _
        rowCount, _
        colCount)

    For r = 1 To rowCount

        For c = 1 To colCount

            wdTable.Cell(r, c).Range.Text = _
                tbl.DataBodyRange(r, firstCol + c - 1).Value

        Next c
    Next r

    FormatTable wdTable

    wdDoc.Content.InsertAfter vbCrLf & vbCrLf

End Sub

Private Function LastUsedTableRow(ByVal tbl As ListObject, _
                                  ByVal firstCol As Long, _
                                  ByVal lastCol As Long) As Long

    Dim r As Long
    Dim c As Long

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
' TABLE HEADER DETECTION
'====================================================

Public Function HasTableColumns(ByVal tbl As ListObject, _
                                 ByVal colNum As Long) As Boolean

    If colNum >= tbl.ListColumns.Count Then Exit Function

    HasTableColumns = _
        IsTableColumn(tbl.ListColumns(colNum + 1).Name)

End Function

Private Sub WriteDocumentTitle(ByVal wdDoc As Object, _
                               ByVal TitleText As String)

    Dim rng As Object

    Set rng = wdDoc.Content
    rng.Collapse 0

    rng.InsertAfter TitleText & vbCrLf
    rng.InsertAfter "Standard Operating Document" & vbCrLf & vbCrLf

    Set rng = wdDoc.Paragraphs(1).Range

    With rng.Font
        .Bold = True
        .Size = 22
        .Name = "Aptos"
    End With

End Sub

Public Function GetLastTableColumn(ByVal tbl As ListObject, _
                                    ByVal colNum As Long) As Long

    Dim c As Long

    c = colNum + 1

    Do While c <= tbl.ListColumns.Count

        If Not IsTableColumn(tbl.ListColumns(c).Name) Then Exit Do

        c = c + 1

    Loop

    GetLastTableColumn = c - 1

End Function

Private Function HasBulletColumns(ByVal tbl As ListObject, _
                                  ByVal colNum As Long) As Boolean

    If colNum >= tbl.ListColumns.Count Then Exit Function

    HasBulletColumns = _
        IsBulletColumn(tbl.ListColumns(colNum + 1).Name)

End Function

Private Function GetLastBulletColumn(ByVal tbl As ListObject, _
                                     ByVal colNum As Long) As Long

    Dim c As Long

    c = colNum + 1

    Do While c <= tbl.ListColumns.Count

        If Not IsBulletColumn(tbl.ListColumns(c).Name) Then Exit Do

        c = c + 1

    Loop

    GetLastBulletColumn = c - 1

End Function

Public Function IsTableColumn(ByVal txt As String) As Boolean

    txt = Trim(txt)

    If InStr(txt, " ") > 0 Then Exit Function

    IsTableColumn = (LCase(Left(txt, 5)) = "table")

End Function

Public Function IsBulletColumn(ByVal txt As String) As Boolean

    txt = Trim(txt)

    If InStr(txt, " ") > 0 Then Exit Function

    IsBulletColumn = (LCase(Left(txt, 6)) = "bullet")

End Function

Public Function IsTitleColumn(ByVal txt As String) As Boolean

    IsTitleColumn = (LCase(Trim(txt)) = "title")

End Function

Private Function GetDocumentTitle(ByVal tbl As ListObject, _
                                  ByVal TitleCol As Long) As String

    Dim r As Long

    For r = 1 To tbl.ListRows.Count

        If Trim(tbl.DataBodyRange(r, TitleCol).Value) <> "" Then

            GetDocumentTitle = _
                Trim(tbl.DataBodyRange(r, TitleCol).Value)

            Exit Function

        End If

    Next r

    GetDocumentTitle = "Generated SOD"

End Function

Private Sub CreateNestedBulletSection(ByVal wdDoc As Object, _
                                      ByVal tbl As ListObject, _
                                      ByVal parentCol As Long)

    Dim r As Long
    Dim lastRow As Long

    Dim bulletLastCol As Long
    Dim parentText As String

    bulletLastCol = GetLastBulletColumn(tbl, parentCol)

    lastRow = tbl.ListRows.Count

    For r = 1 To lastRow

        parentText = Trim(tbl.DataBodyRange(r, parentCol).Value)

        If parentText <> "" Then

            wdDoc.Content.InsertAfter "• " & parentText & vbCrLf

        End If

        WriteNestedLevels wdDoc, tbl, r, parentCol + 1, bulletLastCol

    Next r

    wdDoc.Content.InsertAfter vbCrLf

End Sub

Private Sub WriteNestedLevels(ByVal wdDoc As Object, _
                              ByVal tbl As ListObject, _
                              ByVal rowNum As Long, _
                              ByVal firstBulletCol As Long, _
                              ByVal lastBulletCol As Long)

    Dim c As Long
    Dim txt As String
    Dim indent As String

    For c = firstBulletCol To lastBulletCol

        txt = Trim(tbl.DataBodyRange(rowNum, c).Value)

        If txt <> "" Then

            indent = String((c - firstBulletCol + 1) * 4, " ")

            wdDoc.Content.InsertAfter _
                indent & "• " & txt & vbCrLf

        End If

    Next c

End Sub

'====================================================
' FORMATTING
'====================================================
' MODIFY THIS SECTION LATER
'====================================================

Private Sub FormatHeading(ByVal rng As Object)

    With rng.Font
        .Bold = True
        .Size = 16
        .Name = "Aptos"
    End With

End Sub

Private Sub FormatTable(ByVal tbl As Object)

    On Error Resume Next

    tbl.Style = "Table Grid"

    tbl.Rows(1).Range.Bold = True

    On Error GoTo 0

End Sub
