Attribute VB_Name = "Module4"
Option Explicit

Sub HideUncheckedRowsInSelectedTableColumn()

    Dim lo As ListObject
    Dim selectedCell As Range
    Dim targetColumn As Range
    Dim dataCell As Range
    Dim hasBoolean As Boolean
    Dim colIndex As Long
    Dim i As Long

    ' Must have exactly one cell selected
    If Selection.Cells.Count <> 1 Then
        MsgBox "Please select a single cell in the table column containing checkboxes or TRUE/FALSE values.", vbCritical
        Exit Sub
    End If

    Set selectedCell = Selection.Cells(1)

    ' Must be inside a table
    On Error Resume Next
    Set lo = selectedCell.ListObject
    On Error GoTo 0

    If lo Is Nothing Then
        MsgBox "The selected cell is not inside an Excel Table.", vbCritical
        Exit Sub
    End If

    ' Determine column index (header or data cell)
    If Not lo.HeaderRowRange Is Nothing _
       And Not Intersect(selectedCell, lo.HeaderRowRange) Is Nothing Then

        ' Header cell selected
        colIndex = selectedCell.Column - lo.HeaderRowRange.Cells(1, 1).Column + 1

    ElseIf Not Intersect(selectedCell, lo.DataBodyRange) Is Nothing Then

        ' Data cell selected
        colIndex = selectedCell.Column - lo.DataBodyRange.Cells(1, 1).Column + 1

    Else
        MsgBox "Please select either the header or a data cell of the checkbox column.", vbCritical
        Exit Sub
    End If

    ' Get the target column in the data body
    Set targetColumn = lo.ListColumns(colIndex).DataBodyRange

    ' Validate column contains TRUE/FALSE values
    hasBoolean = False
    For Each dataCell In targetColumn.Cells
        If VarType(dataCell.Value) = vbBoolean Then
            hasBoolean = True
            Exit For
        End If
    Next dataCell

    If Not hasBoolean Then
        MsgBox "The selected column does not appear to contain TRUE/FALSE values or linked checkboxes.", vbCritical
        Exit Sub
    End If

    ' Unhide all rows in the table first
    lo.DataBodyRange.EntireRow.Hidden = False

    ' Hide rows where value is FALSE
    For i = 1 To targetColumn.Rows.Count
        If targetColumn.Cells(i, 1).Value = False Then
            targetColumn.Cells(i, 1).EntireRow.Hidden = True
        End If
    Next i

End Sub

