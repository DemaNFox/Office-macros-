Sub SaveEachSheetAsNewFile()
    Dim ws As Worksheet
    Dim i As Integer
    Dim fileName As String
    Dim filePath As String
    
    ' Запрашиваем у пользователя название файла
    fileName = InputBox("Введите название файла:", "Сохранить как")
    
    ' Проверяем, чтобы пользователь ввел название файла
    If fileName = "" Then
        MsgBox "Вы не ввели название файла. Операция отменена."
        Exit Sub
    End If
    
    ' Получаем путь к рабочему столу
    filePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
    
    i = 1 ' Инициализируем счетчик
    
    ' Перебираем все листы в книге
    For Each ws In ThisWorkbook.Sheets
        ' Сохраняем текущий лист в новый файл на рабочем столе
        ws.Copy
        ActiveWorkbook.SaveAs fileName:=filePath & fileName & "_" & i & ".xlsx"
        ActiveWorkbook.Close False
        i = i + 1 ' Увеличиваем счетчик для следующего файла
    Next ws
    
    MsgBox "Файлы успешно сохранены на рабочем столе."
End Sub

Sub SplitRowsIntoSheets()
    Dim wsSource As Worksheet
    Dim wsNew As Worksheet
    Dim lastRow As Long
    Dim totalRows As Long
    Dim rowsPerSheet As Long
    Dim sheetCount As Long
    Dim i As Long
    Dim j As Long
    Dim dataRange As Range
    
    ' Запрос пользователя о количестве строк на каждом листе
    rowsPerSheet = InputBox("Введите количество строк на каждом листе:", "Разделение строк на листы", 1000)
    If rowsPerSheet <= 0 Then Exit Sub
    
    ' Указываем первый лист
    Set wsSource = ThisWorkbook.Sheets(1)
    
    ' Определяем количество строк в исходном листе
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' Определяем общее количество строк
    totalRows = lastRow ' Предполагаем, что есть заголовок
    
    ' Определяем количество необходимых листов
    sheetCount = Application.WorksheetFunction.Ceiling(totalRows / rowsPerSheet, 1)
    
    ' Копируем заголовок
    Set dataRange = wsSource.Rows(1)
    
    ' Создаем новые листы и копируем данные
    For i = 1 To sheetCount
        Set wsNew = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsNew.Name = "Sheet" & i
        
        ' Копируем заголовок
        dataRange.Copy wsNew.Rows(1)
        
        ' Копируем строки данных
        If i < sheetCount Then
            wsSource.Rows((i - 1) * rowsPerSheet + 2 & ":" & i * rowsPerSheet + 1).EntireRow.Copy wsNew.Rows(2)
        Else
            wsSource.Rows((i - 1) * rowsPerSheet + 2 & ":" & totalRows).EntireRow.Copy wsNew.Rows(2)
        End If
        
        ' Удаляем пустую первую строку
        wsNew.Rows(1).Delete
    Next i
    
    MsgBox "Разделение завершено."
End Sub
