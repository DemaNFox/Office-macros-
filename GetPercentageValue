Function GetPercentageValue(cell As Range) As String
    Dim text As String
    Dim percentPos As Integer
    Dim startPos As Integer
    Dim length As Integer

    ' Получаем текст из ячейки
    text = cell.Value

    ' Ищем символ процента
    percentPos = InStr(text, "%")
    
    ' Проверяем, нашли ли процент
    If percentPos > 0 Then
        startPos = percentPos - 1
        Do While startPos > 0 And (Mid(text, startPos, 1) Like "[0-9]" Or Mid(text, startPos, 1) = ".")
            startPos = startPos - 1
        Loop
        startPos = startPos + 1

        length = percentPos - startPos
        GetPercentageValue = Mid(text, startPos, length)
    Else
        GetPercentageValue = ""
    End If
End Function
