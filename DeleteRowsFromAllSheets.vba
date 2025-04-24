' Удаляет указанное количество первых строк на всех листах

Sub DeleteRowsFromAllSheets()
    Dim ws As Worksheet
    Dim rowCount As Long
    Dim response As String
    
    ' Запрос количества строк для удаления
    response = InputBox("Введите количество строк для удаления:", "Удаление строк")
    
    ' Проверка ввода
    If IsNumeric(response) Then
        rowCount = CLng(response)
        
        ' Проход по всем листам
        For Each ws In ThisWorkbook.Worksheets
            ' Удаление строк
            ws.Rows("1:" & rowCount).Delete
        Next ws
        
        MsgBox "Удалено " & rowCount & " строк(и) с каждого листа."
    Else
        MsgBox "Пожалуйста, введите правильное числовое значение."
    End If
End Sub
