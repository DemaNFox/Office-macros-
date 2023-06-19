Function FindCellLeftValue(searchValue As Variant, searchRange As Range) As Variant
    Dim foundCell As Range
    Dim leftCell As Range
    
    Set foundCell = searchRange.Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        If foundCell.Column > 1 Then
            Set leftCell = foundCell.Offset(0, -1)
            FindCellLeftValue = leftCell.Value
        Else
            FindCellLeftValue = "No left cell"
        End If
    Else
        FindCellLeftValue = "Value not found"
    End If
End Function
