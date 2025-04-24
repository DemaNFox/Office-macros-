' Удаляет дубликаты по строкам. Учитывает выбранный столбец или диапазон но удаляет все строки с дублирующимися значениями. Отличается втроенного инструмента тем что сносит всю строку без сдвига данных.

Sub RemoveDuplicate()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim phoneColumn As Range
    Dim i As Long, j As Long
    Dim dict As Object
    Dim removedCount As Long
    Dim data As Variant
    Dim result As Variant
    Dim resultIndex As Long
    Dim totalColumns As Long

    ' Отключаем обновление экрана и вычисления для ускорения работы
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Устанавливаем активный лист
    Set ws = ActiveSheet

    ' Запрашиваем у пользователя выбор столбца с телефонами
    On Error Resume Next
    Set phoneColumn = Application.InputBox("Выберите столбец с номерами телефонов:", "Выбор столбца", Type:=8)
    On Error GoTo 0

    If phoneColumn Is Nothing Then
        MsgBox "Выбор столбца отменен.", vbExclamation
        Exit Sub
    End If

    ' Определяем последнюю заполненную строку и количество столбцов в листе
    lastRow = ws.Cells(ws.Rows.Count, phoneColumn.Column).End(xlUp).Row
    totalColumns = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Создаём словарь для хранения уникальных номеров телефонов
    Set dict = CreateObject("Scripting.Dictionary")

    ' Инициализируем счётчик удалённых строк
    removedCount = 0

    ' Читаем все данные в диапазоне A1 до последнего столбца и последней строки
    data = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, totalColumns)).Value
    ReDim result(1 To UBound(data, 1), 1 To UBound(data, 2))
    resultIndex = 1

    ' Проходим по каждой строке данных
    For i = 1 To UBound(data, 1)
        Dim phoneNumber As String
        phoneNumber = data(i, phoneColumn.Column)

        ' Проверяем, есть ли этот номер телефона в словаре
        If dict.exists(phoneNumber) Then
            ' Если дубликат найден, увеличиваем счётчик удалённых строк
            removedCount = removedCount + 1
        Else
            ' Если номер уникален, добавляем его в словарь и копируем всю строку в итоговый массив
            dict.Add phoneNumber, Nothing
            For j = 1 To UBound(data, 2)
                result(resultIndex, j) = data(i, j)
            Next j
            resultIndex = resultIndex + 1
        End If
    Next i

    ' Очищаем исходные данные на листе
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, totalColumns)).ClearContents

    ' Записываем обратно на лист данные без дубликатов
    ws.Range(ws.Cells(1, 1), ws.Cells(resultIndex - 1, UBound(result, 2))).Value = result

    ' Очистка словаря
    Set dict = Nothing

    ' Восстанавливаем настройки экрана и вычислений
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' Выводим сообщение о завершении
    MsgBox "Удаление дубликатов завершено! Удалено " & removedCount & " строк."
End Sub

