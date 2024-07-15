Sub isBeelineYota ()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim phoneNumber As String
    Dim operatorName As String
    Dim i As Long
    Dim beelineCodes As Variant
    Dim codeDict As Object
    Dim phoneColumn As Range
    Dim colIndex As Long
    Dim colLetter As String
    
    ' Установите ссылку на активный рабочий лист
    Set ws = ActiveSheet
    
    ' Запрос у пользователя выбора столбца с номерами телефонов
    On Error Resume Next
    Set phoneColumn = Application.InputBox("Выберите столбец с номерами телефонов", Type:=8)
    On Error GoTo 0
    
    ' Проверка, был ли выбран столбец
    If phoneColumn Is Nothing Then
        MsgBox "Столбец не выбран. Прерывание выполнения макроса."
        Exit Sub
    End If
    
    ' Определение индекса и буквы выбранного столбца
    colIndex = phoneColumn.Cells(1, 1).Column
    colLetter = Split(phoneColumn.Address, "$")(2)
    
    ' Определите последнюю строку в выбранном столбце
    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row

    ' Список кодов Билайн
    beelineCodes = Array("900335", "900336", "337", "338", "339", "340", "341", "342", "343", "344", _
                         "90205", "90206", "90207", "90252", "902553", "902554", "902555", "902556", _
                         "902557", "902559", "902710", "902717", "903", "90462", "904700", "904701", _
                         "904702", "904703", "904704", "904705", "904706", "904707", "904726", "904727", _
                         "905", "906", "908364", "908375", "908376", "908377", "908378", "908379", _
                         "908435", "908436", "908437", "908438", "908439", "90844", "90845", "908460", _
                         "908461", "908462", "908463", "908464", "90896", "90897", "90898", "90899", _
                         "909", "930095", "933001", "933002", "933003", "933004", "933005", "933006", _
                         "933007", "933008", "933009", "933010", "933011", "933012", "933013", _
                         "933034", "933035", "933036", "950668", "950880", "95100", "95101", "95102", _
                         "95320", "95321", "95322", "960", "961", "962", "963", "964", "965", "966", _
                         "967", "968", "969", "983888", "983999", "986667", "986668", "986669", _
                         "986670", "986671", "986672", "986673", "986666", "996", "999")
    
				' Создание словаря для быстрого поиска кодов
				Set codeDict = CreateObject("Scripting.Dictionary")
				For Each code In beelineCodes
						codeDict(code) = True
				Next code

    
    ' Вставка нового столбца для оператора
    ws.Columns(colLetter & ":" & colLetter).Insert Shift:=xlToRight
    ws.Cells(1, colIndex).Value = "Оператор"
    
    ' Проверка номеров телефонов
    For i = 2 To lastRow
        phoneNumber = CStr(ws.Cells(i, colIndex + 1).Value)
        operatorName = ""
        
        ' Удаление кода страны, если присутствует
        If Left(phoneNumber, 2) = "+7" Then
            phoneNumber = Mid(phoneNumber, 3)
        ElseIf Left(phoneNumber, 1) = "7" Then
            phoneNumber = Mid(phoneNumber, 2)
        End If
        
        ' Проверка на коды Билайн
        For Each code In codeDict.Keys
            If Left(phoneNumber, Len(code)) = code Then
                operatorName = "Билайн"
                Exit For
            End If
        Next code
        
        ws.Cells(i, colIndex).Value = operatorName
    Next i
    
    ' Освобождение ресурсов
    Set codeDict = Nothing
End Sub
