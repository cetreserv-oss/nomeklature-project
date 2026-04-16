Option Explicit

Public Const FSM_REQUEST_SHEET_NAME As String = "Отправить запрос по ФСМ"
Public Const NOMENCLATURE_REQUEST_SHEET_NAME As String = "Отправка марок (номенклатура)"
Public Const IMPORT_INFO_REQUEST_SHEET_NAME As String = "Сведения о ввозе (номенклатура)"
Public Const IMPORT_INFO_BASELINE_SHEET_NAME As String = "__Ввоз_База"
Public Const ALCO_REPORT_SHEET_NAME As String = "Алкоотчет"
Public Const SETTINGS_SHEET_NAME As String = "Настройка"
Private Const SHEET_PROTECTION_PASSWORD As String = ""

Public Sub Main()
    Call RunRefreshPipeline(True, False)
End Sub

Public Sub ValidateNomenclatureRequest()
    PrepareNomenclatureRequest
End Sub

Public Function RunRefreshPipeline(Optional ByVal showSuccessMessage As Boolean = True, Optional ByVal leaveFsmSheetWritable As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    Dim currentStep As String

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    currentStep = "подготовка листов"
    PrepareSheetsForRefresh

    currentStep = "сбор заказов"
    CollectIznachalnieZakazi
    currentStep = "загрузка Алкоотчета"
    UpdateDataSheets
    currentStep = "импорт Контроль марок"
    ImportKontrolMarokData
    currentStep = "обработка Алкоотчета"
    ObrabotkaAlkoReport
    currentStep = "перенос данных Алкоотчета"
    TransferAlcoReportToRabochiy
    currentStep = "очистка неиспользованных строк"
    DeleteNeispolzovannieStrokiizRabochego
    currentStep = "поиск изменений"
    NaitiIVidelitIzmenenia
    currentStep = "заполнение действия"
    AddDeystvieToZapros
    currentStep = "подсветка излишка ФСМ"
    HighlightIzhlishekFSM
    currentStep = "форматирование листа"
    CenterAndWrapColumnsAtoM

    GetAlcoReportSheet.Visible = xlSheetHidden

    If leaveFsmSheetWritable Then
        ProtectStagingRequestSheets
    Else
        EnsureInteractiveSheetProtection
    End If

    GetFsmRequestSheet.Activate

    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    RunRefreshPipeline = True

    If showSuccessMessage Then MsgBox "Готово", vbInformation
    Exit Function

ErrHandler:
    Dim userMessage As String
    Dim originalErrorDescription As String
    Dim failedStep As String

    originalErrorDescription = Err.Description
    failedStep = currentStep

    On Error Resume Next
    GetAlcoReportSheet.Visible = xlSheetHidden
    EnsureInteractiveSheetProtection
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0

    RunRefreshPipeline = False
    userMessage = BuildRefreshErrorMessage(originalErrorDescription, failedStep)

    If AutomationModeEnabled Then
        SetAutomationResult "error", userMessage
        Exit Function
    End If

    MsgBox userMessage, vbCritical, "Не удалось обновить данные"
End Function

Public Sub EnsureInteractiveSheetProtection()
    ProtectFsmRequestSheet
    ProtectStagingRequestSheets
End Sub

Public Sub PrepareSheetsForRefresh()
    UnprotectAutomationSheets
End Sub

Public Sub UnprotectAutomationSheets()
    On Error Resume Next
    GetFsmRequestSheet.Unprotect Password:=SHEET_PROTECTION_PASSWORD
    GetNomenclatureRequestSheet.Unprotect Password:=SHEET_PROTECTION_PASSWORD
    GetImportInfoRequestSheet.Unprotect Password:=SHEET_PROTECTION_PASSWORD
    On Error GoTo 0
End Sub

Private Sub ProtectStagingRequestSheets()
    ProtectNomenclatureRequestSheet
    ProtectImportInfoRequestSheet
    EnsureImportInfoBaselineHidden
End Sub

Public Sub ProtectFsmRequestSheet()
    Dim ws As Worksheet
    Dim lastCol As Long

    Set ws = GetFsmRequestSheet
    ws.Unprotect Password:=SHEET_PROTECTION_PASSWORD

    lastCol = LastUsedColumn(ws)
    If lastCol < 1 Then lastCol = 1

    ws.Cells.Locked = True
    ws.Range(ws.Cells(2, 1), ws.Cells(ws.Rows.Count, lastCol)).Locked = False

    ws.Protect Password:=SHEET_PROTECTION_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
               UserInterfaceOnly:=False, AllowFormattingCells:=False, AllowFormattingColumns:=True, _
               AllowFormattingRows:=False, AllowInsertingColumns:=False, AllowInsertingRows:=True, _
               AllowInsertingHyperlinks:=False, AllowDeletingColumns:=False, AllowDeletingRows:=False, _
               AllowSorting:=False, AllowFiltering:=True, AllowUsingPivotTables:=False
    ws.EnableSelection = xlUnlockedCells
End Sub

Public Sub ProtectNomenclatureRequestSheet()
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim lastRow As Long

    Set ws = GetNomenclatureRequestSheet
    ws.Unprotect Password:=SHEET_PROTECTION_PASSWORD

    lastCol = LastUsedColumn(ws)
    If lastCol < 9 Then lastCol = 9
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then lastRow = 2

    ws.Cells.Locked = False
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Locked = True

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 9)).AutoFilter

    ws.Protect Password:=SHEET_PROTECTION_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
               UserInterfaceOnly:=False, AllowFormattingCells:=False, AllowFormattingColumns:=True, _
               AllowFormattingRows:=False, AllowInsertingColumns:=False, AllowInsertingRows:=True, _
               AllowInsertingHyperlinks:=False, AllowDeletingColumns:=False, AllowDeletingRows:=False, _
               AllowSorting:=False, AllowFiltering:=True, AllowUsingPivotTables:=False
    ws.EnableSelection = xlUnlockedCells
End Sub

Public Sub ProtectImportInfoRequestSheet()
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim lastRow As Long

    Set ws = GetImportInfoRequestSheet
    ws.Unprotect Password:=SHEET_PROTECTION_PASSWORD

    lastCol = LastUsedColumn(ws)
    If lastCol < 14 Then lastCol = 14
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then lastRow = 2

    ws.Cells.Locked = False
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Locked = True

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 14)).AutoFilter

    ws.Protect Password:=SHEET_PROTECTION_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
               UserInterfaceOnly:=False, AllowFormattingCells:=False, AllowFormattingColumns:=True, _
               AllowFormattingRows:=False, AllowInsertingColumns:=False, AllowInsertingRows:=True, _
               AllowInsertingHyperlinks:=False, AllowDeletingColumns:=False, AllowDeletingRows:=True, _
               AllowSorting:=False, AllowFiltering:=True, AllowUsingPivotTables:=False
    ws.EnableSelection = xlUnlockedCells
End Sub

Private Sub EnsureImportInfoBaselineHidden()
    On Error Resume Next
    GetImportInfoBaselineSheet.Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub

Public Function BuildRefreshErrorMessage(ByVal errorDescription As String, Optional ByVal failedStep As String = "") As String
    Dim defaultDownloadsPath As String
    Dim failedStepText As String

    defaultDownloadsPath = GetDefaultDownloadsPath()
    If Len(Trim$(failedStep)) > 0 Then
        failedStepText = vbCrLf & "Шаг, на котором остановился макрос: " & failedStep
    End If

    If InStr(1, errorDescription, "среди данных файла 'Контроль марок.xlsx'", vbTextCompare) > 0 Then
        BuildRefreshErrorMessage = "Введенный заказ не найден в источнике 'Контроль марок.xlsx'." & vbCrLf & vbCrLf & _
                                   "Что нужно сделать:" & vbCrLf & _
                                   "1. Проверьте номер заказа." & vbCrLf & _
                                   "2. Если вводите 'ТК' кириллицей, это допустимо: макрос сам приводит его к 'TK'." & vbCrLf & _
                                   "3. Если номер верный, проверьте, что заказ уже появился в источнике 'Контроль марок.xlsx'." & failedStepText & vbCrLf & vbCrLf & _
                                   "Техническая причина: " & errorDescription
        Exit Function
    End If

    If InStr(1, errorDescription, "Папка загрузки не найдена:", vbTextCompare) > 0 Then
        BuildRefreshErrorMessage = "Не удалось найти папку загрузки, указанную на листе '" & SETTINGS_SHEET_NAME & "'." & vbCrLf & vbCrLf & _
                                   "Что нужно сделать:" & vbCrLf & _
                                   "1. Откройте лист '" & SETTINGS_SHEET_NAME & "'." & vbCrLf & _
                                   "2. Проверьте строку 'Папка загрузки'." & vbCrLf & _
                                   "3. Укажите существующую папку, куда выгружается Алкоотчет." & vbCrLf & _
                                   IIf(Len(defaultDownloadsPath) > 0, "Рекомендуемый путь для этого ПК: " & defaultDownloadsPath & vbCrLf, "") & failedStepText & vbCrLf & vbCrLf & _
                                   "Техническая причина: " & errorDescription
        Exit Function
    End If

    If InStr(1, errorDescription, "Файл Алкоотчета не найден", vbTextCompare) > 0 Then
        BuildRefreshErrorMessage = "Не найден актуальный файл Алкоотчета в папке загрузки." & vbCrLf & vbCrLf & _
                                   "Что нужно сделать:" & vbCrLf & _
                                   "1. Выгрузите свежий Алкоотчет." & vbCrLf & _
                                   "2. Поместите файл вида '*_ALCOHOL_REPORT.xlsx' в папку загрузки из листа '" & SETTINGS_SHEET_NAME & "'." & failedStepText & vbCrLf & vbCrLf & _
                                   "Техническая причина: " & errorDescription
        Exit Function
    End If

    If InStr(1, errorDescription, "Контроль марок.xlsx", vbTextCompare) > 0 Then
        BuildRefreshErrorMessage = "Не удалось получить данные из файла 'Контроль марок.xlsx'." & vbCrLf & vbCrLf & _
                                   "Что нужно сделать:" & vbCrLf & _
                                   "1. Проверьте путь в строке 'Контроль марок' на листе '" & SETTINGS_SHEET_NAME & "'." & vbCrLf & _
                                   "2. Убедитесь, что файл существует и не заблокирован для чтения." & vbCrLf & _
                                   "3. Если файл открыт, сохраните его вручную или закройте и повторите попытку." & failedStepText & vbCrLf & vbCrLf & _
                                   "Техническая причина: " & errorDescription
        Exit Function
    End If

    If InStr(1, errorDescription, "нет в выгрузке Алкоотчета", vbTextCompare) > 0 Then
        BuildRefreshErrorMessage = "Выгрузка Алкоотчета не содержит один или несколько введенных заказов." & vbCrLf & vbCrLf & _
                                   "Что нужно сделать:" & vbCrLf & _
                                   "1. Проверьте номер заказа." & vbCrLf & _
                                   "2. Если вводите 'ТК' кириллицей, это допустимо: макрос сам приводит его к 'TK'." & vbCrLf & _
                                   "3. Обновите выгрузку Алкоотчета и убедитесь, что заказ в нее попал." & failedStepText & vbCrLf & vbCrLf & _
                                   "Техническая причина: " & errorDescription
        Exit Function
    End If

    If InStr(1, errorDescription, "Столбец '", vbTextCompare) > 0 Then
        BuildRefreshErrorMessage = "Не удалось обработать выгрузку, потому что в файле не хватает ожидаемых столбцов." & vbCrLf & vbCrLf & _
                                   "Что нужно сделать:" & vbCrLf & _
                                   "1. Убедитесь, что выгружен правильный файл нужного формата." & vbCrLf & _
                                   "2. Если формат источника изменился, передайте файл на доработку макроса." & failedStepText & vbCrLf & vbCrLf & _
                                   "Техническая причина: " & errorDescription
        Exit Function
    End If

    If InStr(1, errorDescription, "Invalid procedure call or argument", vbTextCompare) > 0 Then
        BuildRefreshErrorMessage = "Не удалось обработать один или несколько введенных заказов." & vbCrLf & vbCrLf & _
                                   "Что нужно сделать:" & vbCrLf & _
                                   "1. Проверьте номер заказа." & vbCrLf & _
                                   "2. Если вводите 'ТК' кириллицей, это допустимо: макрос сам приводит его к 'TK'." & vbCrLf & _
                                   "3. Убедитесь, что заказ существует в источниках 'Контроль марок.xlsx' и '_ALCOHOL_REPORT'." & failedStepText & vbCrLf & vbCrLf & _
                                   "Техническая причина: " & errorDescription
        Exit Function
    End If

    BuildRefreshErrorMessage = "Не удалось обновить данные." & vbCrLf & vbCrLf & _
                               "Проверьте номер заказа, настройки путей и актуальность исходных файлов." & failedStepText & vbCrLf & vbCrLf & _
                               "Техническая причина: " & errorDescription
End Function

Public Function GetFsmRequestSheet() As Worksheet
    Set GetFsmRequestSheet = Sheet1
End Function

Public Function GetSettingsSheet() As Worksheet
    Set GetSettingsSheet = Sheet5
End Function

Public Function GetAlcoReportSheet() As Worksheet
    Set GetAlcoReportSheet = Sheet2
End Function

Public Function GetNomenclatureRequestSheet() As Worksheet
    On Error Resume Next
    Set GetNomenclatureRequestSheet = ThisWorkbook.Worksheets(NOMENCLATURE_REQUEST_SHEET_NAME)
    On Error GoTo 0

    If GetNomenclatureRequestSheet Is Nothing Then
        Err.Raise vbObjectError + 910, , "Не найден лист '" & NOMENCLATURE_REQUEST_SHEET_NAME & "'."
    End If
End Function

Public Function GetImportInfoRequestSheet() As Worksheet
    On Error Resume Next
    Set GetImportInfoRequestSheet = ThisWorkbook.Worksheets(IMPORT_INFO_REQUEST_SHEET_NAME)
    On Error GoTo 0

    If GetImportInfoRequestSheet Is Nothing Then
        Err.Raise vbObjectError + 911, , "Не найден лист '" & IMPORT_INFO_REQUEST_SHEET_NAME & "'."
    End If
End Function

Public Function GetImportInfoBaselineSheet() As Worksheet
    On Error Resume Next
    Set GetImportInfoBaselineSheet = ThisWorkbook.Worksheets(IMPORT_INFO_BASELINE_SHEET_NAME)
    On Error GoTo 0

    If GetImportInfoBaselineSheet Is Nothing Then
        Err.Raise vbObjectError + 912, , "Не найден лист '" & IMPORT_INFO_BASELINE_SHEET_NAME & "'."
    End If
End Function

Public Function GetSettingsValue(ByVal settingName As String, Optional ByVal defaultValue As String = "") As String
    Dim ws As Worksheet
    Dim foundCell As Range

    Set ws = GetSettingsSheet
    Set foundCell = ws.Columns(1).Find(What:=settingName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If foundCell Is Nothing Then
        GetSettingsValue = defaultValue
    Else
        GetSettingsValue = Trim(CStr(ws.Cells(foundCell.Row, 2).Value))
    End If
End Function

Public Function EnsureTrailingSlash(ByVal pathValue As String) As String
    pathValue = Trim(pathValue)
    If Len(pathValue) = 0 Then
        EnsureTrailingSlash = ""
    ElseIf Right$(pathValue, 1) = "\" Then
        EnsureTrailingSlash = pathValue
    Else
        EnsureTrailingSlash = pathValue & "\"
    End If
End Function

Private Function GetDefaultDownloadsPath() As String
    Dim userProfilePath As String

    userProfilePath = Trim(Environ$("USERPROFILE"))
    If Len(userProfilePath) = 0 Then Exit Function

    If Right$(userProfilePath, 1) <> "\" Then userProfilePath = userProfilePath & "\"
    GetDefaultDownloadsPath = userProfilePath & "Downloads"
End Function

Public Function FindColumn(ws As Worksheet, colName As String) As Long
    Dim lastCol As Long
    Dim c As Long

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If StrComp(Trim(CStr(ws.Cells(1, c).Value)), colName, vbTextCompare) = 0 Then
            FindColumn = c
            Exit Function
        End If
    Next c

    Err.Raise vbObjectError + 700, , "Не найден столбец '" & colName & "' на листе '" & ws.Name & "'."
End Function

Public Function LastUsedColumn(ws As Worksheet) As Long
    Dim lastCell As Range

    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, LookAt:=xlPart, _
                                 SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)

    If lastCell Is Nothing Then
        LastUsedColumn = 1
    Else
        LastUsedColumn = lastCell.Column
    End If
End Function

Public Function NormalizeCyrLat(ByVal s As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String

    If Len(Trim(s)) = 0 Then
        NormalizeCyrLat = ""
        Exit Function
    End If

    result = UCase(Trim(s))

    For i = 1 To Len(result)
        ch = Mid(result, i, 1)

        Select Case ch
            Case "А": ch = "A"
            Case "В": ch = "B"
            Case "С": ch = "C"
            Case "Е": ch = "E"
            Case "Н": ch = "H"
            Case "К": ch = "K"
            Case "М": ch = "M"
            Case "О": ch = "O"
            Case "Р": ch = "P"
            Case "Т": ch = "T"
            Case "Х": ch = "X"
            Case "У": ch = "Y"
        End Select

        Mid(result, i, 1) = ch
    Next i

    NormalizeCyrLat = result
End Function

Public Function ValidateFsmData() As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim r As Long
    Dim hasErrors As Boolean
    Dim colPairs As Variant
    Dim colLeft As Long
    Dim colRight As Long
    Dim leftValue As Variant
    Dim rightValue As Variant

    Set ws = GetFsmRequestSheet
    ClearFsmValidationFormatting ws

    lastRow = ws.Cells(ws.Rows.Count, FindColumn(ws, "Заказ")).End(xlUp).Row
    If lastRow < 2 Then
        ValidateFsmData = True
        Exit Function
    End If

    colPairs = Array( _
        Array("Заявление (КМ)", "Заявление (новое)"), _
        Array("Поставщик (КМ)", "Поставщик (новый)"), _
        Array("Код (КМ)", "Код (новый)"), _
        Array("Позиция (КМ)", "Позиция (новая)"), _
        Array("Кол-во (КМ)", "Кол-во (новое)") _
    )

    For i = LBound(colPairs) To UBound(colPairs)
        colLeft = FindColumn(ws, CStr(colPairs(i)(0)))
        colRight = FindColumn(ws, CStr(colPairs(i)(1)))

        For r = 2 To lastRow
            leftValue = ws.Cells(r, colLeft).Value
            rightValue = ws.Cells(r, colRight).Value

            If HasValidationIssue(leftValue, rightValue) Then
                MarkValidationCell ws.Cells(r, colLeft)
                MarkValidationCell ws.Cells(r, colRight)
                hasErrors = True
            End If
        Next r
    Next i

    colLeft = FindColumn(ws, "Заявление (КМ)")
    colRight = FindColumn(ws, "Заявление (новое)")

    For r = 2 To lastRow
        If ContainsReserveWord(ws.Cells(r, colLeft).Value) Then
            MarkValidationCell ws.Cells(r, colLeft)
            hasErrors = True
        End If

        If ContainsReserveWord(ws.Cells(r, colRight).Value) Then
            MarkValidationCell ws.Cells(r, colRight)
            hasErrors = True
        End If
    Next r

    ValidateFsmData = Not hasErrors
End Function

Private Sub ClearFsmValidationFormatting(ws As Worksheet)
    Dim lastRow As Long
    Dim lastCol As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = LastUsedColumn(ws)

    If lastRow >= 2 And lastCol >= 1 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Interior.ColorIndex = xlNone
    End If
End Sub

Private Sub MarkValidationCell(targetCell As Range)
    targetCell.Interior.Color = RGB(255, 199, 206)
End Sub

Private Function ContainsReserveWord(ByVal value As Variant) As Boolean
    ContainsReserveWord = InStr(1, NormalizeCellText(value), "резерв", vbTextCompare) > 0
End Function

Private Function HasValidationIssue(ByVal leftValue As Variant, ByVal rightValue As Variant) As Boolean
    If NormalizeCellText(leftValue) = "" Or NormalizeCellText(rightValue) = "" Then
        HasValidationIssue = True
        Exit Function
    End If

    HasValidationIssue = Not AreValidationValuesEqual(leftValue, rightValue)
End Function

Private Function AreValidationValuesEqual(ByVal leftValue As Variant, ByVal rightValue As Variant) As Boolean
    If IsNumeric(leftValue) And IsNumeric(rightValue) Then
        AreValidationValuesEqual = CDbl(leftValue) = CDbl(rightValue)
    Else
        AreValidationValuesEqual = StrComp(NormalizeCellText(leftValue), NormalizeCellText(rightValue), vbTextCompare) = 0
    End If
End Function

Private Function NormalizeCellText(ByVal value As Variant) As String
    Dim result As String

    If IsError(value) Or IsEmpty(value) Then
        NormalizeCellText = ""
        Exit Function
    End If

    result = CStr(value)
    result = Replace(result, Chr(160), " ")
    result = Replace(result, vbCrLf, " ")
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")

    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop

    NormalizeCellText = Trim(result)
End Function
