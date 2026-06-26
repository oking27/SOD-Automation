Attribute VB_Name = "Module2"
Sub FlattenToPlainText()
    Dim cell As Range
    
    For Each cell In Selection
        If Not IsEmpty(cell.Value) Then
            cell.Value = CleanText(cell.Value)
        End If
    Next cell
End Sub

Function CleanText(ByVal txt As String) As String
    ' Normalize quotes
    txt = Replace(txt, Chr(147), """")
    txt = Replace(txt, Chr(148), """")
    
    ' Remove line breaks
    txt = Replace(txt, vbCr, " ")
    txt = Replace(txt, vbLf, " ")
    
    ' Remove non-breaking spaces
    txt = Replace(txt, Chr(160), " ")
    
    ' Remove bullet characters
    txt = Replace(txt, Chr(149), "")
    txt = Replace(txt, "•", "")
    
    ' Strip extra spaces
    txt = Application.WorksheetFunction.Trim(txt)
    
    CleanText = txt
End Function
