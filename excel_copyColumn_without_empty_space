Sub CopySelectedColumnsWithoutTrailingBlanks()
    Dim selectedColumns As Range
    Dim lastRow As Long
    Dim copyRange As Range
    Dim ws As Worksheet
    Dim col As Range
    Dim tempRange As Range

    ' Работа с активным листом и выделенными столбцами
    Set ws = Application.ActiveSheet
    Set selectedColumns = Selection

    ' Проверка, что выделены только столбцы
    If selectedColumns.Areas.Count > 1 Or selectedColumns.Cells.Count = 0 Then
        MsgBox "Пожалуйста, выделите только столбцы."
        Exit Sub
    End If

    ' Перебор всех выделенных столбцов
    For Each col In selectedColumns.Columns
        ' Поиск последней строки с данными в текущем столбце
        lastRow = ws.Cells(ws.Rows.Count, col.Column).End(xlUp).Row
        
        ' Установка диапазона для копирования, исключая пустые строки
        Set tempRange = col.Resize(lastRow)
        
        ' Объединение диапазонов
        If copyRange Is Nothing Then
            Set copyRange = tempRange
        Else
            Set copyRange = Union(copyRange, tempRange)
        End If
    Next col

    ' Копирование объединенного диапазона
    If Not copyRange Is Nothing Then
        copyRange.Copy
    End If
End Sub
