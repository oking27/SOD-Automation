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
            colMap(LCase(Trim(ws.Cells(1, col).Value))) = col
        End If
    Next col

    ' =======================
    ' FIND HEADER ROW
    ' =======================
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

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

    wdApp.Documents.Open "C:\Users\OliviaKing\OneDrive - Gregory Construction Services\Desktop\SOD_Desktop.docx"
    Set wdDoc = wdApp.ActiveDocument

    ' =======================
    ' SAVE NEW FILE
    ' =======================
    wdDoc.SaveAs2 outputPath

    wdApp.Visible = True

    MsgBox "SOP generated successfully:" & vbCrLf & outputPath

End Sub
