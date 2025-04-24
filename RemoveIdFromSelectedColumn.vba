Sub RemoveIdFromSelectedColumn()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim cleanedText As String
    Dim regex As Object
    Dim selectedColumn As Range

    ' Открываем окно для выбора столбца
    On Error Resume Next
    Set selectedColumn = Application.InputBox("Выберите столбец для обработки:", Type:=8)
    On Error GoTo 0

    ' Проверка на корректность выбранного диапазона
    If selectedColumn Is Nothing Then
        MsgBox "Столбец не был выбран. Операция отменена."
        Exit Sub
    End If

    ' Создадим объект RegExp для удаления id и чисел
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "id:\d+" ' Регулярное выражение для удаления "id:число"
    
    ' Проход по всем ячейкам в выбранном столбце
    For Each cell In selectedColumn
        If Not IsEmpty(cell.Value) Then
            ' Удаляем id и числовое значение
            cleanedText = regex.Replace(cell.Value, "")
            ' Обрезаем пробелы
            cell.Value = Trim(cleanedText)
        End If
    Next cell
End Sub
