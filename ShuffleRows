' Перемешивает строки

Public Sub ShuffleRowsOptimized()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Проверка, что выбран лист, который не является листом в книге макросов
    If ws.Parent.Name = ThisWorkbook.Name Then
        MsgBox "Пожалуйста, выберите лист в другой книге."
        Exit Sub
    End If

    ' Определяем последнюю заполненную строку на листе
    Dim lastRow As Long
    lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

    If lastRow < 2 Then
        MsgBox "Недостаточно строк для перемешивания."
        Exit Sub
    End If

    Dim data As Variant
    Dim indexArray() As Long
    ReDim indexArray(1 To lastRow)

    ' Заполняем массив индексами
    Dim i As Long
    For i = 1 To lastRow
        indexArray(i) = i
    Next i

    ' Перемешиваем индексы
    Dim j As Long, temp As Long
    Randomize
    For i = lastRow To 2 Step -1
        j = Int((i - 1) * Rnd + 1)
        temp = indexArray(i)
        indexArray(i) = indexArray(j)
        indexArray(j) = temp
    Next i

    ' Считываем данные в массив
    data = ws.Range("A1:Z" & lastRow).Value ' Примерно предполагаем, что данные заполнены до столбца Z

    ' Создаем временный массив для перемешанных данных
    Dim shuffledData() As Variant
    ReDim shuffledData(1 To UBound(data, 1), 1 To UBound(data, 2))

    ' Заполняем перемешанный массив
    For i = 1 To lastRow
        For j = 1 To UBound(data, 2)
            shuffledData(i, j) = data(indexArray(i), j)
        Next j
    Next i

    ' Отключаем обновление экрана
    Application.ScreenUpdating = False

    ' Записываем перемешанные данные обратно в лист
    ws.Range("A1:Z" & lastRow).Value = shuffledData

    ' Включаем обновление экрана
    Application.ScreenUpdating = True
End Sub
