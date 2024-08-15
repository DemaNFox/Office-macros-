' Форматирует номер телефона к чистому, убирая все возможные знаки между цифрами.

Function FormatPhoneNumber(ByVal inputString As String) As String
    Dim result As String
    result = Replace(inputString, "+7 (", "7")
    result = Replace(result, ")", "")
    result = Replace(result, " ", "")
    result = Replace(result, "-", "")
    
    ' Проверка на первую цифру числа
    If Len(result) > 0 And Mid(result, 1, 1) = "8" Then
        result = "7" & Mid(result, 2)
    End If
    
    FormatPhoneNumber = result
End Function
