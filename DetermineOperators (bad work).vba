' Определяет операторов по столбцу с номерами телефонов. Низкая точность по сравнению с python программой.

Sub DetermineOperators()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim phoneNumber As String
    Dim operatorName As String
    Dim phoneCol As Range
    Dim operatorCol As Range
    Dim phoneNumbers As Variant
    Dim operators() As String

    ' Используем активный лист
    Set ws = ActiveSheet

    ' Просим пользователя выбрать колонку с номерами телефонов
    On Error Resume Next
    Set phoneCol = Application.InputBox("Выберите колонку с номерами телефонов:", Type:=8)
    On Error GoTo 0

    If phoneCol Is Nothing Then
        MsgBox "Колонка не выбрана!", vbExclamation
        Exit Sub
    End If

    ' Определяем последнюю заполненную строку в выбранной колонке
    lastRow = ws.Cells(ws.Rows.Count, phoneCol.Column).End(xlUp).Row

    ' Вставляем новую колонку справа от выбранной
    ws.Columns(phoneCol.Column + 1).Insert Shift:=xlToRight
    Set operatorCol = ws.Cells(1, phoneCol.Column + 1)
    operatorCol.Value = "Оператор" ' Заголовок новой колонки

    ' Считываем все номера телефонов в массив
    phoneNumbers = ws.Range(ws.Cells(2, phoneCol.Column), ws.Cells(lastRow, phoneCol.Column)).Value
    ReDim operators(1 To UBound(phoneNumbers, 1), 1 To 1)

    ' Проход по каждому номеру телефона в массиве
    For i = 1 To UBound(phoneNumbers, 1)
        phoneNumber = phoneNumbers(i, 1)
        operators(i, 1) = GetOperatorName(phoneNumber)
    Next i

    ' Записываем результаты обратно в Excel
    ws.Range(ws.Cells(2, operatorCol.Column), ws.Cells(lastRow, operatorCol.Column)).Value = operators

    MsgBox "Готово! Операторы добавлены в колонку " & operatorCol.Address(0, 0), vbInformation
End Sub

Function GetOperatorName(phoneNumber As String) As String
    Dim cleanedNumber As String
    cleanedNumber = Replace(phoneNumber, "+", "")
    cleanedNumber = Replace(cleanedNumber, "-", "")
    cleanedNumber = Replace(cleanedNumber, " ", "")
    cleanedNumber = Replace(cleanedNumber, "(", "")
    cleanedNumber = Replace(cleanedNumber, ")","")

    If Len(cleanedNumber) = 11 And Left(cleanedNumber, 1) = "8" Then
        cleanedNumber = "7" & Mid(cleanedNumber, 2)
    ElseIf Len(cleanedNumber) = 10 And Left(cleanedNumber, 1) = "9" Then
        cleanedNumber = "7" & cleanedNumber
    End If

    Select Case Left(cleanedNumber, 4)
        Case "7910", "7911", "7912", "7913", "7914", "7915", "7916", "7917", "7918", "7919", _
             "7980", "7981", "7982", "7983", "7984", "7985", "7986", "7987", "7988", "7989"
            GetOperatorName = "MTS"
        Case "7920", "7921", "7922", "7923", "7924", "7925", "7926", "7927", "7928", "7929", _
             "7930", "7931", "7932", "7933", "7934", "7937", "7938", "7939"
            GetOperatorName = "Megafon"
        Case "7960", "7961", "7962", "7963", "7964", "7965", "7966", "7967", "7968", "7969", _
             "7902", "7903", "7904", "7905", "7906", "7907", "7908", "7909"
            GetOperatorName = "Beeline"
        Case "7950", "7951", "7952", "7953", "7954", "7955", "7956", "7957", "7958", "7959"
            GetOperatorName = "Tele2"
        Case "7999"
            GetOperatorName = "Yota"
        Case "7940", "7970", "7990"
            GetOperatorName = "Rostelecom"
        Case Else
            GetOperatorName = ""
    End Select
End Function
