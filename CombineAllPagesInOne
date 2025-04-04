Sub CombineAllPagesInOne()
    Dim ws As Worksheet
    Dim wsОбщий As Worksheet
    Dim ПоследняяСтрока As Long
    Dim ПоследняяСтрокаОбщий As Long
    Dim ПервыйЛист As Boolean: ПервыйЛист = True
    Dim wbTarget As Workbook

    Set wbTarget = Application.ActiveWorkbook

    ' Защита от запуска на Personal.xlsb
    If wbTarget.Name = "PERSONAL.XLSB" Then
        MsgBox "Пожалуйста, откройте файл, листы которого нужно объединить.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' Проверка: существует ли лист "Объединено"
    On Error Resume Next
    Set wsОбщий = wbTarget.Worksheets("Объединено")
    On Error GoTo 0

    ' Если нет — создаём
    If wsОбщий Is Nothing Then
        Set wsОбщий = wbTarget.Worksheets.Add
        wsОбщий.Name = "Объединено"
    Else
        wsОбщий.Cells.Clear
    End If

    ПоследняяСтрокаОбщий = 1

    ' Проходим по всем листам книги
    For Each ws In wbTarget.Worksheets
        If ws.Name <> wsОбщий.Name Then
            ПоследняяСтрока = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

            If ПервыйЛист Then
                ws.Rows("1:" & ПоследняяСтрока).Copy Destination:=wsОбщий.Cells(ПоследняяСтрокаОбщий, 1)
                ПервыйЛист = False
            Else
                ws.Rows("2:" & ПоследняяСтрока).Copy Destination:=wsОбщий.Cells(ПоследняяСтрокаОбщий, 1)
            End If

            ПоследняяСтрокаОбщий = wsОбщий.Cells(wsОбщий.Rows.Count, 1).End(xlUp).Row + 1
        End If
    Next ws

    Application.ScreenUpdating = True

    MsgBox "Объединение завершено!", vbInformation
End Sub
