Attribute VB_Name = "Module3"
Option Explicit

Sub BuildSOD()

    Dim wdApp As Object
    Dim wdDoc As Object
    Dim ws As Worksheet
    Dim templatePath As String
    Dim outputFolder As String
    Dim outputPath As String
    Dim fileName As String
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, col As Long
    Dim headerRow As Long
    Dim colMap As Object

    Set ws = ActiveSheet

    ' =======================
    ' BUILD HEADER MAP
    ' =======================
    Set colMap = CreateObject("Scripting.Dictionary")

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For col = 1 To lastCol
        If ws.Cells(1, col).Value <> "" Then
            colMap(Trim(ws.Cells(1, col).Value)) = col
        End If
    Next col

    ' =======================
    ' FIND HEADER ROW
    ' =======================
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If ws.Cells(i, colMap("SOD Title")).Value <> "" Then
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
    fileName = ws.Cells(headerRow, colMap("SOD Title")).Value

    If fileName = "" Then fileName = "SOD_Output"

    fileName = Replace(fileName, "/", "-")
    fileName = Replace(fileName, "\", "-")
    fileName = Replace(fileName, ":", "-")

    ' ? SAVE LOCALLY FIRST (FIXES DUPES)
    outputFolder = "C:\Users\OliviaKing\OneDrive - Gregory Construction Services\Desktop\SOP_Output\"

    If Dir(outputFolder, vbDirectory) = "" Then MkDir outputFolder

    outputPath = outputFolder & fileName & ".docx"

    ' =======================
    ' OPEN WORD + TEMPLATE
    ' =======================
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False

    wdApp.Documents.Open "C:\Users\OliviaKing\OneDrive - Gregory Construction Services\Desktop\SOD_Template.docx" 'Document used!!
    Set wdDoc = wdApp.ActiveDocument 'Current document
    
    
    '====================
    'Content Control
    '====================
    Dim cc As Object
    Dim tbl As Object
    

    For Each cc In wdDoc.ContentControls

        Select Case cc.Title
            Case "cc_title"
            cc.Range.Text = ws.Cells(headerRow, colMap("SOD Title")).Value
                
            Case "cc_policy_statement"
                cc.Range.Text = ws.Cells(headerRow, colMap("Policy Statement")).Value
                
            Case "cc_purpose"
                cc.Range.Text = ws.Cells(headerRow, colMap("Purpose")).Value
                
            Case "cc_scope"
                cc.Range.Text = ws.Cells(headerRow, colMap("Scope")).Value
                
            Case "cc_resource"
                cc.Range.Text = ws.Cells(headerRow, colMap("SOD Resource")).Value
                
            Case "cc_version"
                cc.Range.Text = ws.Cells(headerRow, colMap("SOD Version")).Value
                
            Case "cc_definitions"
                Set tbl = wdDoc.Tables.Add(cc.Range, 1, 2)

                tbl.cell(1, 1).Range.Text = "Term"
                tbl.cell(1, 2).Range.Text = "Definition"

                For i = headerRow To lastRow

                    If ws.Cells(i, colMap("Dictionary")).Value <> "" Then
                        tbl.Rows.Add

                        tbl.cell(tbl.Rows.Count, 1).Range.Text = _
                        ws.Cells(i, colMap("Term")).Value

                        tbl.cell(tbl.Rows.Count, 2).Range.Text = _
                        ws.Cells(i, colMap("Definition")).Value

                    End If

                Next i

        End Select

    Next cc


    ' =======================
    ' SAVE NEW FILE
    ' =======================
    wdDoc.SaveAs2 outputPath

    wdApp.Visible = True

    MsgBox "SOP generated successfully:" & vbCrLf & outputPath

End Sub
