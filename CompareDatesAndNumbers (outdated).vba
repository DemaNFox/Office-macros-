Function CompareDatesAndNumbers(DateRange As Range, NumberRange As Range, TargetNumber As String) As String
    Dim Result As String
    Dim DatesArray As Variant
    Dim NumbersArray As Variant
    Dim i As Long
    
    Result = ""
    
    ' Преобразование диапазонов в массивы для улучшения производительности
    DatesArray = DateRange.Value
    NumbersArray = NumberRange.Value
    
    ' Loop through the arrays to compare numbers
    For i = LBound(NumbersArray, 1) To UBound(NumbersArray, 1)
        If NumbersArray(i, 1) = TargetNumber Then
            ' If the numbers match, add the corresponding date to the result
            If Result = "" Then
                Result = DatesArray(i, 1)
            Else
                Result = Result & ", " & DatesArray(i, 1)
            End If
        End If
    Next i
    
    CompareDatesAndNumbers = Result
End Function
