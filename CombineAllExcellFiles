'Кобмбинирует все файлы ексель в указанной папке на одном листе

Sub CombineAllExcellFiles()
    ' Выбор папки с файлами
    Dim sFolder As String, sFiles As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = False Then Exit Sub
        sFolder = .SelectedItems(1)
    End With
    sFolder = sFolder & IIf(Right(sFolder, 1) = Application.PathSeparator, "", Application.PathSeparator)
    sFiles = Dir(sFolder & "*.xls*")
    Application.ScreenUpdating = False

    ' Создание нового отчета
    Dim wbReport As Workbook, wsReport As Worksheet
    Set wbReport = Workbooks.Add
    Set wsReport = wbReport.Sheets(1)
    Dim n As Long
    n = 1 ' Начальная строка для копирования данных

    ' Открытие и обработка каждого файла в выбранной папке
    Do While sFiles <> ""
        Dim wbCurrent As Workbook, wsCurrent As Worksheet
        Set wbCurrent = Workbooks.Open(sFolder & sFiles)
        
        ' Копирование данных с каждого листа
        For Each wsCurrent In wbCurrent.Worksheets
            ' Определение последней строки и последнего столбца
            Dim lastRow As Long, lastCol As Long
            Dim lastCell As Range
            
            ' Поиск последней заполненной ячейки по строкам
            Set lastCell = wsCurrent.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
            If Not lastCell Is Nothing Then
                lastRow = lastCell.Row
            Else
                lastRow = 0 ' Если данных нет, устанавливаем lastRow в 0
            End If

            ' Поиск последней заполненной ячейки по столбцам
            Set lastCell = wsCurrent.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
            If Not lastCell Is Nothing Then
                lastCol = lastCell.Column
            Else
                lastCol = 0 ' Если данных нет, устанавливаем lastCol в 0
            End If
            
            ' Копирование данных, начиная со второй строки
            If lastRow > 1 Then
                wsCurrent.Range(wsCurrent.Cells(2, 1), wsCurrent.Cells(lastRow, lastCol)).Copy _
                    Destination:=wsReport.Cells(n, 1)
                n = n + lastRow - 1
            End If
        Next wsCurrent
        
        wbCurrent.Close savechanges:=False
        sFiles = Dir
    Loop

    ' Удаление первого листа (если нужно)
    Application.DisplayAlerts = False
    If wbReport.Sheets.Count > 1 Then
        wbReport.Sheets(1).Delete
    End If
    Application.DisplayAlerts = True

    Application.ScreenUpdating = True
    MsgBox "Merge complete!"
End Sub
