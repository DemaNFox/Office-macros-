' Сохранения всех листов excel в одтельные файлы с указаниеи общего имени и приставкой номера файла.

Sub SaveEachSheetAsNewFile()
    Dim ws As Worksheet
    Dim i As Integer
    Dim fileName As String
    Dim filePath As String
    
    ' Запрашиваем у пользователя название файла
    fileName = InputBox("Введите название файла:", "Сохранение файлов")
    
    ' Проверяем, было ли введено название файла
    If fileName = "" Then
        MsgBox "Вы не ввели название файла. Операция отменена."
        Exit Sub
    End If
    
    ' Путь сохранения на рабочий стол
    filePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
    
    i = 1 ' Счетчик файлов
    
    ' Проходим по всем листам в активной книге
    For Each ws In ActiveWorkbook.Sheets
        ' Копируем лист в новую книгу
        ws.Copy
        ' Сохраняем новую книгу
        ActiveWorkbook.SaveAs fileName:=filePath & fileName & "_" & i & ".xlsx"
        ' Закрываем новую книгу
        ActiveWorkbook.Close False
        i = i + 1 ' Увеличиваем счетчик для следующего файла
    Next ws
    
    MsgBox "Файлы успешно сохранены на рабочем столе."
End Sub
