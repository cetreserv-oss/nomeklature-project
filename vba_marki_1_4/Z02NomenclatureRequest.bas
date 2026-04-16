Option Explicit

Private Const NOMENCLATURE_HEADER_ORDER As String = "Номер заказа"
Private Const NOMENCLATURE_HEADER_CODE As String = "Код УТ"
Private Const NOMENCLATURE_HEADER_NAME As String = "Номенклатура"
Private Const NOMENCLATURE_HEADER_STATEMENT As String = "Заявление на выдачу ФСМ"
Private Const NOMENCLATURE_HEADER_QTY As String = "Кол-во"
Private Const NOMENCLATURE_HEADER_STICKER As String = "Оклейщик"
Private Const NOMENCLATURE_HEADER_SUPPLIER As String = "Поставщик"
Private Const NOMENCLATURE_HEADER_MIL_COMMENT As String = "Комментарий МИЛ"
Private Const NOMENCLATURE_HEADER_MIL As String = "МИЛ"
Private Const IMPORT_INFO_HEADER_ORDER As String = "Номер заказа"
Private Const IMPORT_INFO_HEADER_SUPPLIER As String = "Поставщик"
Private Const IMPORT_INFO_HEADER_CODE As String = "Код УТ"
Private Const IMPORT_INFO_HEADER_NAME As String = "Номенклатура"
Private Const IMPORT_INFO_HEADER_STATEMENT As String = "Заявление на выдачу ФСМ"
Private Const IMPORT_INFO_HEADER_QTY As String = "Количество"
Private Const IMPORT_INFO_HEADER_STICKER As String = "Оклейщик"
Private Const IMPORT_INFO_HEADER_SHIPMENT_INVOICE As String = "Накладная отправки"
Private Const IMPORT_INFO_HEADER_SHIPMENT_DATE As String = "Дата отправки"
Private Const IMPORT_INFO_HEADER_ALCOHOL As String = "Градус алкоголя"
Private Const IMPORT_INFO_HEADER_VOLUME As String = "Объем бутылки"
Private Const IMPORT_INFO_HEADER_VINTAGE As String = "Год урожая"
Private Const IMPORT_INFO_HEADER_MIL_COMMENT As String = "Комментарий МИЛ"
Private Const IMPORT_INFO_HEADER_MIL As String = "МИЛ"

Private Const SETTING_NOMENCLATURE_PATH As String = "Номенклатура"
Private Const SETTING_ARCHIVE_PATH As String = "Архив исходных запросов"
Private Const SETTING_NOMENCLATURE_PASSWORD As String = "Пароль номенклатуры"
Private Const SETTING_NOMENCLATURE_PROTECTED_SHEETS_PASSWORD As String = "Пароль защищенных вкладок номенклатуры"

Private Const ARCHIVE_SHEET_NAME As String = "Исх. запросы (отпр ФСМ)"
Private Const CORRECTION_ARCHIVE_SHEET_NAME As String = "Корр. (отпр ФСМ)"
Private Const IMPORT_INFO_ARCHIVE_SHEET_NAME As String = "Исх. запросы (ввоз)"
Private Const IMPORT_INFO_CORRECTION_ARCHIVE_SHEET_NAME As String = "Корр. (ввоз)"
Private Const ARCHIVE_PASSWORD As String = "7777"
Private Const ARCHIVE_HEADER_DATE As String = "Дата внесения строки"
Private Const ARCHIVE_HEADER_CONFIRMATION As String = "Подтверждение повторной отправки"
Private Const ARCHIVE_HEADER_QTY_MISMATCH_CONFIRMATION As String = "Подтверждение несоответствия количества"
Private Const CORRECTION_STATUS_HEADER As String = "Статус"
Private Const CORRECTION_STATUS_PENDING As String = "Ожидает корректировки"

Private Const TARGET_TK_SHEET_NAME As String = "TK отправка марки"
Private Const TARGET_LA_SHEET_NAME As String = "LA отправка марки"
Private Const CORRECTION_TARGET_TK_SHEET_NAME As String = "TK КОРРЕКТИРОВКА отправки марки"
Private Const CORRECTION_TARGET_LA_SHEET_NAME As String = "LA КОРРЕКТИРОВКА отправки марки"
Private Const IMPORT_INFO_TARGET_TK_SHEET_NAME As String = "TK фиксация сведений о ввозе"
Private Const IMPORT_INFO_TARGET_LA_SHEET_NAME As String = "LA фиксация сведений о ввозе"
Private Const IMPORT_INFO_CORRECTION_TARGET_TK_SHEET_NAME As String = "TK КОРР. сведений о ввозе"
Private Const IMPORT_INFO_CORRECTION_TARGET_LA_SHEET_NAME As String = "LA КОРР. сведений о ввозе"
Private Const TARGET_HEADER_ARTICLE As String = "Артикул"
Private Const TARGET_HEADER_NAME As String = "Номенклатура"
Private Const TARGET_HEADER_ORDER As String = "Номер Заказа"
Private Const TARGET_HEADER_STATEMENT As String = "Заявление на выдачу ФСМ"
Private Const TARGET_HEADER_QTY As String = "Кол-во"
Private Const TARGET_HEADER_STICKER As String = "Оклейщик"
Private Const TARGET_HEADER_SHIPMENT_INVOICE As String = "Накладная отправки"
Private Const TARGET_HEADER_SHIPMENT_DATE As String = "Дата отправки"
Private Const TARGET_HEADER_DATE As String = "Дата внесения строки"
Private Const TARGET_HEADER_SUPPLIER As String = "Поставщик"
Private Const TARGET_HEADER_MIL_COMMENT As String = "Комментарий МИЛ"
Private Const TARGET_HEADER_MIL As String = "МИЛ"
Private Const TARGET_HEADER_ALCOHOL As String = "Градус алкоголя"
Private Const TARGET_HEADER_VOLUME As String = "Объем бутылки,л"
Private Const TARGET_HEADER_VINTAGE As String = "Год урожая"
Private Const MISSING_DATA_TEXT As String = "данные не подтянулись, вбейте информацию руками"
Private Const MANUAL_FILL_TEXT As String = "Заполните данные"
Private Const MULTIPLE_MATCHES_TEXT As String = "найдено несколько совпадений, проверьте вручную"
Private Const IMPORT_INFO_QTY_CONFIRMATION_KEY As String = "__IMPORT_INFO_QTY_MISMATCH__"

Private Const LOCK_FILE_SUFFIX As String = ".macro.lock"
Private Const LOCK_RETRY_COUNT As Long = 20
Private Const LOCK_RETRY_SECONDS As Double = 3# / 86400#
Private Const CONFIRMATION_TEXT As String = "подтверждаю"
Private Const CONFIRMATION_PROMPT_TEXT As String = "ПОДТВЕРЖДАЮ"
Private Const AUTOMATION_TRACE_FILE_NAME As String = "send_trace.log"

Private Const AUTOMATION_ERROR_NUMBER As Long = vbObjectError + 970
Private Const SEND_ERROR_NUMBER As Long = vbObjectError + 971
Private Const SEND_CONFLICT_ERROR_NUMBER As Long = vbObjectError + 972

Public AutomationModeEnabled As Boolean
Private AutomationLastStatus As String
Private AutomationLastMessage As String
Private AutomationConfirmations As Object

Public Sub EnableAutomationMode()
    AutomationModeEnabled = True
    ResetAutomationContext
End Sub

Public Sub DisableAutomationMode()
    AutomationModeEnabled = False
    ResetAutomationContext
End Sub

Public Function GetAutomationStatus() As String
    GetAutomationStatus = AutomationLastStatus
End Function

Public Function GetAutomationMessage() As String
    GetAutomationMessage = AutomationLastMessage
End Function

Public Sub SetAutomationConfirmation(ByVal orderValue As String, ByVal responseValue As String)
    If AutomationConfirmations Is Nothing Then
        Set AutomationConfirmations = CreateObject("Scripting.Dictionary")
        AutomationConfirmations.CompareMode = vbTextCompare
    End If

    AutomationConfirmations(NormalizeOrderValue(orderValue)) = Trim$(CStr(responseValue))
End Sub

Public Sub PrepareNomenclatureRequest()
    On Error GoTo ErrHandler

    Dim orderList As Collection
    Dim orderMap As Object
    Dim wasAutomationMode As Boolean

    Set orderList = New Collection
    Set orderMap = CreateObject("Scripting.Dictionary")
    orderMap.CompareMode = vbTextCompare
    wasAutomationMode = AutomationModeEnabled
    ResetAutomationResult

    PrepareSheetsForRefresh
    NormalizeNomenclatureOrderInput orderList, orderMap

    If orderList.Count = 0 Then
        ShowNomenclatureRequestError "На листе '" & NOMENCLATURE_REQUEST_SHEET_NAME & "' нет заказов для подготовки.", False, vbExclamation
    End If

    LoadOrdersIntoFsmSheet orderList

    If Not RunRefreshPipeline(False, True) Then
        If wasAutomationMode Then
            EndAutomationRun
            Exit Sub
        End If
        End
    End If

    PrepareSheetsForRefresh
    Application.ScreenUpdating = False

    If Not ValidateFsmData() Then
        ShowNomenclatureRequestError "Найдены разночтения, которые нужно устранить до подготовки запроса. Они подсвечены на листе '" & FSM_REQUEST_SHEET_NAME & "'.", True, vbExclamation
    End If

    PopulateNomenclatureRequestSheet orderMap

    EnsureInteractiveSheetProtection
    GetNomenclatureRequestSheet.Activate
    SetAutomationResult "success", "Готово"
    Application.ScreenUpdating = True
    If Not wasAutomationMode Then MsgBox "Готово", vbInformation
    If wasAutomationMode Then EndAutomationRun
    Exit Sub

ErrHandler:
    On Error Resume Next
    EnsureInteractiveSheetProtection
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    If wasAutomationMode Then
        If Err.Number <> AUTOMATION_ERROR_NUMBER And Len(AutomationLastStatus) = 0 Then
            SetAutomationResult "runtime_error", Err.Description
        End If
        EndAutomationRun
        Exit Sub
    End If
    MsgBox "Ошибка при подготовке запроса: " & Err.Description, vbCritical
    End
End Sub

Public Sub PrepareImportInfoRequest()
    On Error GoTo ErrHandler

    Dim orderList As Collection
    Dim orderMap As Object
    Dim shipmentLookup As Object
    Dim wasAutomationMode As Boolean

    Set orderList = New Collection
    Set orderMap = CreateObject("Scripting.Dictionary")
    orderMap.CompareMode = vbTextCompare
    wasAutomationMode = AutomationModeEnabled
    ResetAutomationResult

    PrepareSheetsForRefresh
    NormalizeImportInfoOrderInput orderList, orderMap

    If orderList.Count = 0 Then
        ShowImportInfoRequestError "На листе '" & IMPORT_INFO_REQUEST_SHEET_NAME & "' нет заказов для подготовки.", False, vbExclamation
    End If

    LoadOrdersIntoFsmSheet orderList

    If Not RunRefreshPipeline(False, True) Then
        If wasAutomationMode Then
            EndAutomationRun
            Exit Sub
        End If
        End
    End If

    PrepareSheetsForRefresh
    Application.ScreenUpdating = False

    If Not ValidateFsmData() Then
        ShowImportInfoRequestError "Найдены разночтения, которые нужно устранить до подготовки запроса. Они подсвечены на листе '" & FSM_REQUEST_SHEET_NAME & "'.", True, vbExclamation
    End If

    Set shipmentLookup = BuildImportInfoShipmentLookup(orderMap)
    PopulateImportInfoRequestSheet orderMap, shipmentLookup

    EnsureInteractiveSheetProtection
    GetImportInfoRequestSheet.Activate
    SetAutomationResult "success", "Готово"
    Application.ScreenUpdating = True
    If Not wasAutomationMode Then MsgBox "Готово", vbInformation
    If wasAutomationMode Then EndAutomationRun
    Exit Sub

ErrHandler:
    On Error Resume Next
    EnsureInteractiveSheetProtection
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    If wasAutomationMode Then
        If Err.Number <> AUTOMATION_ERROR_NUMBER And Len(AutomationLastStatus) = 0 Then
            SetAutomationResult "runtime_error", Err.Description
        End If
        EndAutomationRun
        Exit Sub
    End If
    MsgBox "Ошибка при подготовке запроса: " & Err.Description, vbCritical
    End
End Sub

Public Sub SendImportInfoRequest()
    On Error GoTo ErrHandler

    Dim preparedRows As Collection
    Dim uniqueOrders As Collection
    Dim orderSnapshot As Object
    Dim confirmationMap As Object
    Dim nomenclaturePath As String
    Dim archivePath As String
    Dim nomenclaturePassword As String
    Dim lockPath As String
    Dim lockHandle As Integer
    Dim lockAcquired As Boolean
    Dim wbNomenclature As Workbook
    Dim wbArchive As Workbook
    Dim archiveStartRow As Long
    Dim archiveRowCount As Long
    Dim targetProtectionStates As Object
    Dim archiveSaved As Boolean
    Dim successMessage As String
    Dim errorMessage As String
    Dim rollbackMessage As String
    Dim quantityMismatchConfirmation As String
    Dim wasAutomationMode As Boolean

    wasAutomationMode = AutomationModeEnabled
    ResetAutomationResult
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    TraceAutomationStep "import_send:start"
    Set preparedRows = ReadPreparedImportInfoRows(False)
    TraceAutomationStep "import_send:preparedRows=" & preparedRows.Count
    quantityMismatchConfirmation = ConfirmImportInfoQuantityMismatches(preparedRows)
    TraceAutomationStep "import_send:qtyMismatchConfirmation=" & quantityMismatchConfirmation
    Set uniqueOrders = CollectUniqueOrders(preparedRows)
    TraceAutomationStep "import_send:uniqueOrders=" & uniqueOrders.Count

    nomenclaturePath = Trim$(GetSettingsValue(SETTING_NOMENCLATURE_PATH, ""))
    archivePath = Trim$(GetSettingsValue(SETTING_ARCHIVE_PATH, ""))
    nomenclaturePassword = Trim$(GetSettingsValue(SETTING_NOMENCLATURE_PASSWORD, ""))
    TraceAutomationStep "import_send:pathsLoaded"

    ValidateExternalWorkbookPath nomenclaturePath, "файла номенклатуры"
    ValidateExternalWorkbookPath archivePath, "архивного файла"
    TraceAutomationStep "import_send:pathsValidated"

    lockPath = BuildLockFilePath(nomenclaturePath)
    EnsureWorkbookClosedInCurrentInstance nomenclaturePath, "Файл номенклатуры"
    EnsureWorkbookClosedInCurrentInstance archivePath, "Архивный файл"
    EnsureLockFolderWritable lockPath
    TraceAutomationStep "import_send:preflightClosedAndWritable"

    Set orderSnapshot = ReadExistingImportInfoOrderSnapshot(nomenclaturePath, nomenclaturePassword, uniqueOrders)
    TraceAutomationStep "import_send:orderSnapshotLoaded"
    EnsureImportInfoArchiveAccessible archivePath
    TraceAutomationStep "import_send:archiveAccessible"
    Set confirmationMap = CreateObject("Scripting.Dictionary")
    confirmationMap.CompareMode = vbTextCompare
    PopulateRepeatConfirmations uniqueOrders, orderSnapshot, confirmationMap
    TraceAutomationStep "import_send:confirmationsCollected"

    If Not AcquireMacroLock(lockPath, lockHandle) Then
        RaiseSendRequestError "Другой пользователь сейчас вносит строки в shared-номенклатуру. Повторите попытку позже."
    End If
    lockAcquired = True
    TraceAutomationStep "import_send:lockAcquired"

    Set wbNomenclature = OpenNomenclatureWorkbook(nomenclaturePath, nomenclaturePassword, False)
    TraceAutomationStep "import_send:nomenclatureOpenedForWrite"
    Set wbArchive = OpenArchiveWorkbook(archivePath, False)
    TraceAutomationStep "import_send:archiveOpenedForWrite"

    ValidateImportInfoOrdersStillWritable wbNomenclature, uniqueOrders, orderSnapshot
    TraceAutomationStep "import_send:ordersStillWritable"
    Set targetProtectionStates = CreateObject("Scripting.Dictionary")
    targetProtectionStates.CompareMode = vbTextCompare
    WriteImportInfoRowsToExternalBooks wbNomenclature, wbArchive, preparedRows, confirmationMap, quantityMismatchConfirmation, targetProtectionStates, archiveStartRow, archiveRowCount, nomenclaturePassword
    RestoreTargetSheetProtection wbNomenclature, targetProtectionStates, nomenclaturePassword
    TraceAutomationStep "import_send:targetProtectionRestored"

    wbArchive.Save
    archiveSaved = True
    TraceAutomationStep "import_send:archiveSaved"
    wbNomenclature.Save
    TraceAutomationStep "import_send:nomenclatureSaved"

    successMessage = "Готово. Строчки сведений о ввозе внесены в номенклатуру и архив."

CleanExit:
    On Error Resume Next
    If Not wbNomenclature Is Nothing Then
        If Not targetProtectionStates Is Nothing Then
            RestoreTargetSheetProtection wbNomenclature, targetProtectionStates, nomenclaturePassword
        End If
    End If
    If Not wbArchive Is Nothing Then wbArchive.Close SaveChanges:=False
    If Not wbNomenclature Is Nothing Then wbNomenclature.Close SaveChanges:=False
    If lockAcquired Then ReleaseMacroLock lockHandle, lockPath
    EnsureInteractiveSheetProtection
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    GetImportInfoRequestSheet.Activate
    On Error GoTo 0

    If Len(successMessage) > 0 Then
        SetAutomationResult "success", successMessage
        If Not wasAutomationMode Then MsgBox successMessage, vbInformation
    ElseIf Len(errorMessage) > 0 Then
        SetAutomationResult "error", errorMessage
        If Not wasAutomationMode Then MsgBox errorMessage, vbCritical
    End If
    If wasAutomationMode Then EndAutomationRun
    Exit Sub

ErrHandler:
    Dim handledErrorNumber As Long
    Dim handledErrorDescription As String

    handledErrorNumber = Err.Number
    handledErrorDescription = Err.Description
    TraceAutomationStep "import_send:error:" & handledErrorNumber & ":" & handledErrorDescription
    errorMessage = BuildSendErrorMessage(handledErrorNumber, handledErrorDescription)

    If archiveSaved And archiveRowCount > 0 Then
        rollbackMessage = RollbackArchiveRows(wbArchive, archiveStartRow, archiveRowCount, IMPORT_INFO_ARCHIVE_SHEET_NAME)
        If Len(rollbackMessage) > 0 Then
            errorMessage = errorMessage & vbCrLf & vbCrLf & rollbackMessage
        End If
    End If

    Resume CleanExit
End Sub

Public Sub SendImportInfoCorrectionRequest()
    On Error GoTo ErrHandler

    Dim preparedRows As Collection
    Dim duplicateSnapshot As Object
    Dim confirmationMap As Object
    Dim nomenclaturePath As String
    Dim archivePath As String
    Dim nomenclaturePassword As String
    Dim protectedSheetsPassword As String
    Dim lockPath As String
    Dim lockHandle As Integer
    Dim lockAcquired As Boolean
    Dim wbNomenclature As Workbook
    Dim wbArchive As Workbook
    Dim archiveStartRow As Long
    Dim archiveRowCount As Long
    Dim targetProtectionStates As Object
    Dim archiveSaved As Boolean
    Dim successMessage As String
    Dim errorMessage As String
    Dim rollbackMessage As String
    Dim quantityMismatchConfirmation As String
    Dim wasAutomationMode As Boolean

    wasAutomationMode = AutomationModeEnabled
    ResetAutomationResult
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    TraceAutomationStep "import_correction:start"
    Set preparedRows = ReadPreparedImportInfoRows(True)
    TraceAutomationStep "import_correction:preparedRows=" & preparedRows.Count
    quantityMismatchConfirmation = ConfirmImportInfoQuantityMismatches(preparedRows)
    TraceAutomationStep "import_correction:qtyMismatchConfirmation=" & quantityMismatchConfirmation

    nomenclaturePath = Trim$(GetSettingsValue(SETTING_NOMENCLATURE_PATH, ""))
    archivePath = Trim$(GetSettingsValue(SETTING_ARCHIVE_PATH, ""))
    nomenclaturePassword = Trim$(GetSettingsValue(SETTING_NOMENCLATURE_PASSWORD, ""))
    protectedSheetsPassword = Trim$(GetSettingsValue(SETTING_NOMENCLATURE_PROTECTED_SHEETS_PASSWORD, ""))
    TraceAutomationStep "import_correction:pathsLoaded"

    ValidateExternalWorkbookPath nomenclaturePath, "файла номенклатуры"
    ValidateExternalWorkbookPath archivePath, "архивного файла"
    lockPath = BuildLockFilePath(nomenclaturePath)
    EnsureWorkbookClosedInCurrentInstance nomenclaturePath, "Файл номенклатуры"
    EnsureWorkbookClosedInCurrentInstance archivePath, "Архивный файл"
    EnsureLockFolderWritable lockPath
    TraceAutomationStep "import_correction:preflightClosedAndWritable"

    Set duplicateSnapshot = ReadExistingImportInfoCorrectionDuplicateSnapshot(nomenclaturePath, nomenclaturePassword, preparedRows)
    TraceAutomationStep "import_correction:duplicateSnapshotLoaded"
    EnsureImportInfoCorrectionArchiveAccessible archivePath
    TraceAutomationStep "import_correction:archiveAccessible"
    Set confirmationMap = CreateObject("Scripting.Dictionary")
    confirmationMap.CompareMode = vbTextCompare
    PopulateImportInfoCorrectionDuplicateConfirmations preparedRows, duplicateSnapshot, confirmationMap
    TraceAutomationStep "import_correction:confirmationsCollected"

    If Not AcquireMacroLock(lockPath, lockHandle) Then
        RaiseSendRequestError "Другой пользователь сейчас вносит строки в shared-номенклатуру. Повторите попытку позже."
    End If
    lockAcquired = True
    TraceAutomationStep "import_correction:lockAcquired"

    Set wbNomenclature = OpenNomenclatureWorkbook(nomenclaturePath, nomenclaturePassword, False)
    TraceAutomationStep "import_correction:nomenclatureOpenedForWrite"
    Set wbArchive = OpenArchiveWorkbook(archivePath, False)
    TraceAutomationStep "import_correction:archiveOpenedForWrite"

    ValidateImportInfoCorrectionRowsStillWritable wbNomenclature, preparedRows, duplicateSnapshot
    TraceAutomationStep "import_correction:rowsStillWritable"
    Set targetProtectionStates = CreateObject("Scripting.Dictionary")
    targetProtectionStates.CompareMode = vbTextCompare
    WriteImportInfoCorrectionRowsToExternalBooks wbNomenclature, wbArchive, preparedRows, quantityMismatchConfirmation, targetProtectionStates, archiveStartRow, archiveRowCount, protectedSheetsPassword
    RestoreTargetSheetProtection wbNomenclature, targetProtectionStates, protectedSheetsPassword
    TraceAutomationStep "import_correction:targetProtectionRestored"

    wbArchive.Save
    archiveSaved = True
    TraceAutomationStep "import_correction:archiveSaved"
    wbNomenclature.Save
    TraceAutomationStep "import_correction:nomenclatureSaved"

    successMessage = "Готово. Строчки внесены в корректировку сведений о ввозе и архив."

CleanExit:
    On Error Resume Next
    If Not wbNomenclature Is Nothing Then
        If Not targetProtectionStates Is Nothing Then
            RestoreTargetSheetProtection wbNomenclature, targetProtectionStates, protectedSheetsPassword
        End If
    End If
    If Not wbArchive Is Nothing Then wbArchive.Close SaveChanges:=False
    If Not wbNomenclature Is Nothing Then wbNomenclature.Close SaveChanges:=False
    If lockAcquired Then ReleaseMacroLock lockHandle, lockPath
    EnsureInteractiveSheetProtection
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    GetImportInfoRequestSheet.Activate
    On Error GoTo 0

    If Len(successMessage) > 0 Then
        SetAutomationResult "success", successMessage
        If Not wasAutomationMode Then MsgBox successMessage, vbInformation
    ElseIf Len(errorMessage) > 0 Then
        SetAutomationResult "error", errorMessage
        If Not wasAutomationMode Then MsgBox errorMessage, vbCritical
    End If
    If wasAutomationMode Then EndAutomationRun
    Exit Sub

ErrHandler:
    Dim handledErrorNumber As Long
    Dim handledErrorDescription As String

    handledErrorNumber = Err.Number
    handledErrorDescription = Err.Description
    TraceAutomationStep "import_correction:error:" & handledErrorNumber & ":" & handledErrorDescription
    errorMessage = BuildCorrectionSendErrorMessage(handledErrorNumber, handledErrorDescription)

    If archiveSaved And archiveRowCount > 0 Then
        rollbackMessage = RollbackArchiveRows(wbArchive, archiveStartRow, archiveRowCount, IMPORT_INFO_CORRECTION_ARCHIVE_SHEET_NAME)
        If Len(rollbackMessage) > 0 Then
            errorMessage = errorMessage & vbCrLf & vbCrLf & rollbackMessage
        End If
    End If

    Resume CleanExit
End Sub

Public Sub SendNomenclatureRequest()
    On Error GoTo ErrHandler

    Dim preparedRows As Collection
    Dim uniqueOrders As Collection
    Dim orderSnapshot As Object
    Dim confirmationMap As Object
    Dim nomenclaturePath As String
    Dim archivePath As String
    Dim nomenclaturePassword As String
    Dim lockPath As String
    Dim lockHandle As Integer
    Dim lockAcquired As Boolean
    Dim wbNomenclature As Workbook
    Dim wbArchive As Workbook
    Dim archiveStartRow As Long
    Dim archiveRowCount As Long
    Dim targetProtectionStates As Object
    Dim archiveSaved As Boolean
    Dim successMessage As String
    Dim errorMessage As String
    Dim rollbackMessage As String
    Dim wasAutomationMode As Boolean

    wasAutomationMode = AutomationModeEnabled
    ResetAutomationResult
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    TraceAutomationStep "send:start"
    Set preparedRows = ReadPreparedNomenclatureRows()
    TraceAutomationStep "send:preparedRows=" & preparedRows.Count
    Set uniqueOrders = CollectUniqueOrders(preparedRows)
    TraceAutomationStep "send:uniqueOrders=" & uniqueOrders.Count

    nomenclaturePath = Trim$(GetSettingsValue(SETTING_NOMENCLATURE_PATH, ""))
    archivePath = Trim$(GetSettingsValue(SETTING_ARCHIVE_PATH, ""))
    nomenclaturePassword = Trim$(GetSettingsValue(SETTING_NOMENCLATURE_PASSWORD, ""))
    TraceAutomationStep "send:pathsLoaded"

    ValidateExternalWorkbookPath nomenclaturePath, "файла номенклатуры"
    ValidateExternalWorkbookPath archivePath, "архивного файла"
    TraceAutomationStep "send:pathsValidated"

    lockPath = BuildLockFilePath(nomenclaturePath)
    TraceAutomationStep "send:lockPathReady"

    EnsureWorkbookClosedInCurrentInstance nomenclaturePath, "Файл номенклатуры"
    EnsureWorkbookClosedInCurrentInstance archivePath, "Архивный файл"
    EnsureLockFolderWritable lockPath
    TraceAutomationStep "send:preflightClosedAndWritable"

    Set orderSnapshot = ReadExistingOrderSnapshot(nomenclaturePath, nomenclaturePassword, uniqueOrders)
    TraceAutomationStep "send:orderSnapshotLoaded"
    EnsureArchiveAccessible archivePath
    TraceAutomationStep "send:archiveAccessible"
    Set confirmationMap = CreateObject("Scripting.Dictionary")
    confirmationMap.CompareMode = vbTextCompare
    PopulateRepeatConfirmations uniqueOrders, orderSnapshot, confirmationMap
    TraceAutomationStep "send:confirmationsCollected"

    If Not AcquireMacroLock(lockPath, lockHandle) Then
        RaiseSendRequestError "Другой пользователь сейчас вносит строки в shared-номенклатуру. Повторите попытку позже."
    End If
    lockAcquired = True
    TraceAutomationStep "send:lockAcquired"

    Set wbNomenclature = OpenNomenclatureWorkbook(nomenclaturePath, nomenclaturePassword, False)
    TraceAutomationStep "send:nomenclatureOpenedForWrite"
    Set wbArchive = OpenArchiveWorkbook(archivePath, False)
    TraceAutomationStep "send:archiveOpenedForWrite"

    ValidateOrdersStillWritable wbNomenclature, uniqueOrders, orderSnapshot
    TraceAutomationStep "send:ordersStillWritable"
    Set targetProtectionStates = CreateObject("Scripting.Dictionary")
    targetProtectionStates.CompareMode = vbTextCompare
    WritePreparedRowsToExternalBooks wbNomenclature, wbArchive, preparedRows, confirmationMap, targetProtectionStates, archiveStartRow, archiveRowCount, nomenclaturePassword
    RestoreTargetSheetProtection wbNomenclature, targetProtectionStates, nomenclaturePassword
    TraceAutomationStep "send:targetProtectionRestored"
    TraceAutomationStep "send:rowsWritten"

    wbArchive.Save
    archiveSaved = True
    TraceAutomationStep "send:archiveSaved"
    wbNomenclature.Save
    TraceAutomationStep "send:nomenclatureSaved"

    successMessage = "Готово. Строчки внесены в номенклатуру и архив."

CleanExit:
    On Error Resume Next
    If Not wbNomenclature Is Nothing Then
        If Not targetProtectionStates Is Nothing Then
            RestoreTargetSheetProtection wbNomenclature, targetProtectionStates, nomenclaturePassword
        End If
    End If
    If Not wbArchive Is Nothing Then wbArchive.Close SaveChanges:=False
    If Not wbNomenclature Is Nothing Then wbNomenclature.Close SaveChanges:=False
    If lockAcquired Then ReleaseMacroLock lockHandle, lockPath
    EnsureInteractiveSheetProtection
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    GetNomenclatureRequestSheet.Activate
    On Error GoTo 0

    If Len(successMessage) > 0 Then
        SetAutomationResult "success", successMessage
        If Not wasAutomationMode Then MsgBox successMessage, vbInformation
    ElseIf Len(errorMessage) > 0 Then
        SetAutomationResult "error", errorMessage
        If Not wasAutomationMode Then MsgBox errorMessage, vbCritical
    End If
    If wasAutomationMode Then EndAutomationRun
    Exit Sub

ErrHandler:
    Dim handledErrorNumber As Long
    Dim handledErrorDescription As String

    handledErrorNumber = Err.Number
    handledErrorDescription = Err.Description
    TraceAutomationStep "send:error:" & handledErrorNumber & ":" & handledErrorDescription
    errorMessage = BuildSendErrorMessage(handledErrorNumber, handledErrorDescription)

    If archiveSaved And archiveRowCount > 0 Then
        rollbackMessage = RollbackArchiveRows(wbArchive, archiveStartRow, archiveRowCount)
        If Len(rollbackMessage) > 0 Then
            errorMessage = errorMessage & vbCrLf & vbCrLf & rollbackMessage
        End If
    End If

    Resume CleanExit
End Sub

Public Sub SendNomenclatureCorrectionRequest()
    On Error GoTo ErrHandler

    Dim preparedRows As Collection
    Dim duplicateSnapshot As Object
    Dim confirmationMap As Object
    Dim nomenclaturePath As String
    Dim archivePath As String
    Dim nomenclaturePassword As String
    Dim protectedSheetsPassword As String
    Dim lockPath As String
    Dim lockHandle As Integer
    Dim lockAcquired As Boolean
    Dim wbNomenclature As Workbook
    Dim wbArchive As Workbook
    Dim archiveStartRow As Long
    Dim archiveRowCount As Long
    Dim targetProtectionStates As Object
    Dim archiveSaved As Boolean
    Dim successMessage As String
    Dim errorMessage As String
    Dim rollbackMessage As String
    Dim wasAutomationMode As Boolean

    wasAutomationMode = AutomationModeEnabled
    ResetAutomationResult
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    TraceAutomationStep "correction:start"
    Set preparedRows = ReadPreparedNomenclatureRows(True)
    TraceAutomationStep "correction:preparedRows=" & preparedRows.Count

    nomenclaturePath = Trim$(GetSettingsValue(SETTING_NOMENCLATURE_PATH, ""))
    archivePath = Trim$(GetSettingsValue(SETTING_ARCHIVE_PATH, ""))
    nomenclaturePassword = Trim$(GetSettingsValue(SETTING_NOMENCLATURE_PASSWORD, ""))
    protectedSheetsPassword = Trim$(GetSettingsValue(SETTING_NOMENCLATURE_PROTECTED_SHEETS_PASSWORD, ""))
    TraceAutomationStep "correction:pathsLoaded"

    ValidateExternalWorkbookPath nomenclaturePath, "файла номенклатуры"
    ValidateExternalWorkbookPath archivePath, "архивного файла"
    TraceAutomationStep "correction:pathsValidated"

    lockPath = BuildLockFilePath(nomenclaturePath)
    TraceAutomationStep "correction:lockPathReady"

    EnsureWorkbookClosedInCurrentInstance nomenclaturePath, "Файл номенклатуры"
    EnsureWorkbookClosedInCurrentInstance archivePath, "Архивный файл"
    EnsureLockFolderWritable lockPath
    TraceAutomationStep "correction:preflightClosedAndWritable"

    Set duplicateSnapshot = ReadExistingCorrectionDuplicateSnapshot(nomenclaturePath, nomenclaturePassword, preparedRows)
    TraceAutomationStep "correction:duplicateSnapshotLoaded"
    EnsureCorrectionArchiveAccessible archivePath
    TraceAutomationStep "correction:archiveAccessible"
    Set confirmationMap = CreateObject("Scripting.Dictionary")
    confirmationMap.CompareMode = vbTextCompare
    PopulateCorrectionDuplicateConfirmations preparedRows, duplicateSnapshot, confirmationMap
    TraceAutomationStep "correction:confirmationsCollected"

    If Not AcquireMacroLock(lockPath, lockHandle) Then
        RaiseSendRequestError "Другой пользователь сейчас вносит строки в shared-номенклатуру. Повторите попытку позже."
    End If
    lockAcquired = True
    TraceAutomationStep "correction:lockAcquired"

    Set wbNomenclature = OpenNomenclatureWorkbook(nomenclaturePath, nomenclaturePassword, False)
    TraceAutomationStep "correction:nomenclatureOpenedForWrite"
    Set wbArchive = OpenArchiveWorkbook(archivePath, False)
    TraceAutomationStep "correction:archiveOpenedForWrite"

    ValidateCorrectionRowsStillWritable wbNomenclature, preparedRows, duplicateSnapshot
    TraceAutomationStep "correction:rowsStillWritable"
    Set targetProtectionStates = CreateObject("Scripting.Dictionary")
    targetProtectionStates.CompareMode = vbTextCompare
    WritePreparedCorrectionRowsToExternalBooks wbNomenclature, wbArchive, preparedRows, targetProtectionStates, archiveStartRow, archiveRowCount, protectedSheetsPassword
    RestoreTargetSheetProtection wbNomenclature, targetProtectionStates, protectedSheetsPassword
    TraceAutomationStep "correction:targetProtectionRestored"
    TraceAutomationStep "correction:rowsWritten"

    wbArchive.Save
    archiveSaved = True
    TraceAutomationStep "correction:archiveSaved"
    wbNomenclature.Save
    TraceAutomationStep "correction:nomenclatureSaved"

    successMessage = "Готово. Строчки внесены в корректировку номенклатуры и архив."

CleanExit:
    On Error Resume Next
    If Not wbNomenclature Is Nothing Then
        If Not targetProtectionStates Is Nothing Then
            RestoreTargetSheetProtection wbNomenclature, targetProtectionStates, protectedSheetsPassword
        End If
    End If
    If Not wbArchive Is Nothing Then wbArchive.Close SaveChanges:=False
    If Not wbNomenclature Is Nothing Then wbNomenclature.Close SaveChanges:=False
    If lockAcquired Then ReleaseMacroLock lockHandle, lockPath
    EnsureInteractiveSheetProtection
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    GetNomenclatureRequestSheet.Activate
    On Error GoTo 0

    If Len(successMessage) > 0 Then
        SetAutomationResult "success", successMessage
        If Not wasAutomationMode Then MsgBox successMessage, vbInformation
    ElseIf Len(errorMessage) > 0 Then
        SetAutomationResult "error", errorMessage
        If Not wasAutomationMode Then MsgBox errorMessage, vbCritical
    End If
    If wasAutomationMode Then EndAutomationRun
    Exit Sub

ErrHandler:
    Dim handledErrorNumber As Long
    Dim handledErrorDescription As String

    handledErrorNumber = Err.Number
    handledErrorDescription = Err.Description
    TraceAutomationStep "correction:error:" & handledErrorNumber & ":" & handledErrorDescription
    errorMessage = BuildCorrectionSendErrorMessage(handledErrorNumber, handledErrorDescription)

    If archiveSaved And archiveRowCount > 0 Then
        rollbackMessage = RollbackArchiveRows(wbArchive, archiveStartRow, archiveRowCount, CORRECTION_ARCHIVE_SHEET_NAME)
        If Len(rollbackMessage) > 0 Then
            errorMessage = errorMessage & vbCrLf & vbCrLf & rollbackMessage
        End If
    End If

    Resume CleanExit
End Sub

Private Function ReadPreparedNomenclatureRows(Optional ByVal useCorrectionTargets As Boolean = False) As Collection
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim rowData As Object
    Dim orderValue As String
    Dim codeValue As String
    Dim nameValue As String
    Dim statementValue As String
    Dim stickerValue As String
    Dim supplierValue As String
    Dim milCommentValue As String
    Dim milValue As String
    Dim qtyValue As Double
    Dim rowHasAnyData As Boolean
    Dim orderColumn As Long
    Dim codeColumn As Long
    Dim nameColumn As Long
    Dim statementColumn As Long
    Dim qtyColumn As Long
    Dim stickerColumn As Long
    Dim supplierColumn As Long
    Dim milColumn As Long
    Dim milCommentColumn As Long

    Set ws = GetNomenclatureRequestSheet
    Set ReadPreparedNomenclatureRows = New Collection
    lastRow = FindSheetLastRow(ws)
    orderColumn = FindColumn(ws, NOMENCLATURE_HEADER_ORDER)
    supplierColumn = FindColumn(ws, NOMENCLATURE_HEADER_SUPPLIER)
    codeColumn = FindColumn(ws, NOMENCLATURE_HEADER_CODE)
    nameColumn = FindColumn(ws, NOMENCLATURE_HEADER_NAME)
    statementColumn = FindColumn(ws, NOMENCLATURE_HEADER_STATEMENT)
    qtyColumn = FindColumn(ws, NOMENCLATURE_HEADER_QTY)
    stickerColumn = FindColumn(ws, NOMENCLATURE_HEADER_STICKER)
    milColumn = FindColumn(ws, NOMENCLATURE_HEADER_MIL)
    milCommentColumn = FindColumn(ws, NOMENCLATURE_HEADER_MIL_COMMENT)

    If lastRow < 2 Then
        RaiseSendRequestError "На листе '" & NOMENCLATURE_REQUEST_SHEET_NAME & "' нет подготовленных строк. Сначала нажмите 'Подготовить строки к внесению в номенклатуру'."
    End If

    For rowIndex = 2 To lastRow
        rowHasAnyData = HasAnyStagingData(ws, rowIndex)
        If Not rowHasAnyData Then GoTo NextPreparedRow

        orderValue = NormalizeOrderValue(ws.Cells(rowIndex, orderColumn).Value)
        supplierValue = Trim$(CStr(ws.Cells(rowIndex, supplierColumn).Value))
        codeValue = Trim$(CStr(ws.Cells(rowIndex, codeColumn).Value))
        nameValue = Trim$(CStr(ws.Cells(rowIndex, nameColumn).Value))
        statementValue = Trim$(CStr(ws.Cells(rowIndex, statementColumn).Value))
        stickerValue = Trim$(CStr(ws.Cells(rowIndex, stickerColumn).Value))
        milValue = Trim$(CStr(ws.Cells(rowIndex, milColumn).Value))
        milCommentValue = Trim$(CStr(ws.Cells(rowIndex, milCommentColumn).Value))

        If Len(orderValue) = 0 Then
            RaiseSendRequestError "В строке " & rowIndex & " листа '" & NOMENCLATURE_REQUEST_SHEET_NAME & "' не заполнен столбец '" & NOMENCLATURE_HEADER_ORDER & "'."
        End If
        If Len(NormalizeStageText(codeValue)) = 0 Then
            RaiseSendRequestError "В строке " & rowIndex & " листа '" & NOMENCLATURE_REQUEST_SHEET_NAME & "' не заполнен столбец '" & NOMENCLATURE_HEADER_CODE & "'."
        End If
        If Len(NormalizeStageText(nameValue)) = 0 Then
            RaiseSendRequestError "В строке " & rowIndex & " листа '" & NOMENCLATURE_REQUEST_SHEET_NAME & "' не заполнен столбец '" & NOMENCLATURE_HEADER_NAME & "'."
        End If
        If Len(NormalizeStageText(statementValue)) = 0 Then
            RaiseSendRequestError "В строке " & rowIndex & " листа '" & NOMENCLATURE_REQUEST_SHEET_NAME & "' не заполнен столбец '" & NOMENCLATURE_HEADER_STATEMENT & "'."
        End If
        If Len(NormalizeStageText(supplierValue)) = 0 Then
            RaiseSendRequestError "В строке " & rowIndex & " листа '" & NOMENCLATURE_REQUEST_SHEET_NAME & "' не заполнен столбец '" & NOMENCLATURE_HEADER_SUPPLIER & "'."
        End If
        If Len(NormalizeStageText(milValue)) = 0 Then
            RaiseSendRequestError "В строке " & rowIndex & " листа '" & NOMENCLATURE_REQUEST_SHEET_NAME & "' не заполнен столбец '" & NOMENCLATURE_HEADER_MIL & "'."
        End If

        qtyValue = ParsePreparedQuantity(ws.Cells(rowIndex, qtyColumn).Value, orderValue, rowIndex)

        Set rowData = CreateObject("Scripting.Dictionary")
        rowData.CompareMode = vbTextCompare
        rowData("Order") = orderValue
        rowData("Code") = codeValue
        rowData("Name") = nameValue
        rowData("Statement") = statementValue
        rowData("Qty") = qtyValue
        rowData("Sticker") = stickerValue
        rowData("Supplier") = supplierValue
        rowData("MilComment") = milCommentValue
        rowData("Mil") = milValue
        If useCorrectionTargets Then
            rowData("TargetSheet") = ResolveCorrectionTargetSheetName(orderValue)
        Else
            rowData("TargetSheet") = ResolveTargetSheetName(orderValue)
        End If
        ReadPreparedNomenclatureRows.Add rowData
NextPreparedRow:
    Next rowIndex

    If ReadPreparedNomenclatureRows.Count = 0 Then
        RaiseSendRequestError "На листе '" & NOMENCLATURE_REQUEST_SHEET_NAME & "' нет подготовленных строк. Сначала нажмите 'Подготовить строки к внесению в номенклатуру'."
    End If
End Function

Private Function HasAnyStagingData(ByVal ws As Worksheet, ByVal rowIndex As Long) As Boolean
    Dim columnIndex As Long

    For columnIndex = 1 To 9
        If Len(NormalizeStageText(ws.Cells(rowIndex, columnIndex).Value)) > 0 Then
            HasAnyStagingData = True
            Exit Function
        End If
    Next columnIndex
End Function

Private Function ReadPreparedImportInfoRows(Optional ByVal useCorrectionTargets As Boolean = False) As Collection
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim rowData As Object
    Dim orderColumn As Long
    Dim supplierColumn As Long
    Dim codeColumn As Long
    Dim nameColumn As Long
    Dim statementColumn As Long
    Dim qtyColumn As Long
    Dim stickerColumn As Long
    Dim invoiceColumn As Long
    Dim shipmentDateColumn As Long
    Dim alcoholColumn As Long
    Dim volumeColumn As Long
    Dim vintageColumn As Long
    Dim milColumn As Long
    Dim milCommentColumn As Long
    Dim orderValue As String
    Dim qtyValue As Double

    Set ws = GetImportInfoRequestSheet
    Set ReadPreparedImportInfoRows = New Collection
    lastRow = FindSheetLastRow(ws)

    orderColumn = FindColumn(ws, IMPORT_INFO_HEADER_ORDER)
    supplierColumn = FindColumn(ws, IMPORT_INFO_HEADER_SUPPLIER)
    codeColumn = FindColumn(ws, IMPORT_INFO_HEADER_CODE)
    nameColumn = FindColumn(ws, IMPORT_INFO_HEADER_NAME)
    statementColumn = FindColumn(ws, IMPORT_INFO_HEADER_STATEMENT)
    qtyColumn = FindColumn(ws, IMPORT_INFO_HEADER_QTY)
    stickerColumn = FindColumn(ws, IMPORT_INFO_HEADER_STICKER)
    invoiceColumn = FindColumn(ws, IMPORT_INFO_HEADER_SHIPMENT_INVOICE)
    shipmentDateColumn = FindColumn(ws, IMPORT_INFO_HEADER_SHIPMENT_DATE)
    alcoholColumn = FindColumn(ws, IMPORT_INFO_HEADER_ALCOHOL)
    volumeColumn = FindColumn(ws, IMPORT_INFO_HEADER_VOLUME)
    vintageColumn = FindColumn(ws, IMPORT_INFO_HEADER_VINTAGE)
    milColumn = FindColumn(ws, IMPORT_INFO_HEADER_MIL)
    milCommentColumn = FindColumn(ws, IMPORT_INFO_HEADER_MIL_COMMENT)

    If lastRow < 2 Then
        RaiseSendRequestError "На листе '" & IMPORT_INFO_REQUEST_SHEET_NAME & "' нет подготовленных строк. Сначала нажмите 'Подготовить строки к внесению в номенклатуру'."
    End If

    For rowIndex = 2 To lastRow
        If Not HasAnyImportInfoStagingData(ws, rowIndex) Then GoTo NextPreparedRow

        orderValue = NormalizeOrderValue(ws.Cells(rowIndex, orderColumn).Value)
        ValidateImportInfoRequiredCell rowIndex, IMPORT_INFO_HEADER_ORDER, orderValue
        ValidateImportInfoRequiredCell rowIndex, IMPORT_INFO_HEADER_SUPPLIER, ws.Cells(rowIndex, supplierColumn).Value
        ValidateImportInfoRequiredCell rowIndex, IMPORT_INFO_HEADER_CODE, ws.Cells(rowIndex, codeColumn).Value
        ValidateImportInfoRequiredCell rowIndex, IMPORT_INFO_HEADER_NAME, ws.Cells(rowIndex, nameColumn).Value
        ValidateImportInfoRequiredCell rowIndex, IMPORT_INFO_HEADER_STATEMENT, ws.Cells(rowIndex, statementColumn).Value
        ValidateImportInfoRequiredCell rowIndex, IMPORT_INFO_HEADER_QTY, ws.Cells(rowIndex, qtyColumn).Value
        ValidateImportInfoRequiredCell rowIndex, IMPORT_INFO_HEADER_STICKER, ws.Cells(rowIndex, stickerColumn).Value
        ValidateImportInfoRequiredCell rowIndex, IMPORT_INFO_HEADER_SHIPMENT_INVOICE, ws.Cells(rowIndex, invoiceColumn).Value
        ValidateImportInfoRequiredCell rowIndex, IMPORT_INFO_HEADER_SHIPMENT_DATE, ws.Cells(rowIndex, shipmentDateColumn).Value
        ValidateImportInfoRequiredCell rowIndex, IMPORT_INFO_HEADER_ALCOHOL, ws.Cells(rowIndex, alcoholColumn).Value
        ValidateImportInfoRequiredCell rowIndex, IMPORT_INFO_HEADER_VOLUME, ws.Cells(rowIndex, volumeColumn).Value
        ValidateImportInfoRequiredCell rowIndex, IMPORT_INFO_HEADER_VINTAGE, ws.Cells(rowIndex, vintageColumn).Value
        ValidateImportInfoRequiredCell rowIndex, IMPORT_INFO_HEADER_MIL, ws.Cells(rowIndex, milColumn).Value

        qtyValue = ParsePreparedImportInfoQuantity(ws.Cells(rowIndex, qtyColumn).Value, orderValue, rowIndex)

        Set rowData = CreateObject("Scripting.Dictionary")
        rowData.CompareMode = vbTextCompare
        rowData("Order") = orderValue
        rowData("Supplier") = Trim$(CStr(ws.Cells(rowIndex, supplierColumn).Value))
        rowData("Code") = Trim$(CStr(ws.Cells(rowIndex, codeColumn).Value))
        rowData("Name") = Trim$(CStr(ws.Cells(rowIndex, nameColumn).Value))
        rowData("Statement") = Trim$(CStr(ws.Cells(rowIndex, statementColumn).Value))
        rowData("Qty") = qtyValue
        rowData("Sticker") = ws.Cells(rowIndex, stickerColumn).Value
        rowData("ShipmentInvoice") = ws.Cells(rowIndex, invoiceColumn).Value
        rowData("ShipmentDate") = ws.Cells(rowIndex, shipmentDateColumn).Value
        rowData("Alcohol") = ws.Cells(rowIndex, alcoholColumn).Value
        rowData("Volume") = ws.Cells(rowIndex, volumeColumn).Value
        rowData("Vintage") = ws.Cells(rowIndex, vintageColumn).Value
        rowData("Mil") = Trim$(CStr(ws.Cells(rowIndex, milColumn).Value))
        rowData("MilComment") = Trim$(CStr(ws.Cells(rowIndex, milCommentColumn).Value))
        If useCorrectionTargets Then
            rowData("TargetSheet") = ResolveImportInfoCorrectionTargetSheetName(orderValue)
        Else
            rowData("TargetSheet") = ResolveImportInfoTargetSheetName(orderValue)
        End If
        ReadPreparedImportInfoRows.Add rowData
NextPreparedRow:
    Next rowIndex

    If ReadPreparedImportInfoRows.Count = 0 Then
        RaiseSendRequestError "На листе '" & IMPORT_INFO_REQUEST_SHEET_NAME & "' нет подготовленных строк. Сначала нажмите 'Подготовить строки к внесению в номенклатуру'."
    End If
End Function

Private Function HasAnyImportInfoStagingData(ByVal ws As Worksheet, ByVal rowIndex As Long) As Boolean
    Dim columnIndex As Long

    For columnIndex = 1 To 14
        If Len(NormalizeStageText(ws.Cells(rowIndex, columnIndex).Value)) > 0 Then
            HasAnyImportInfoStagingData = True
            Exit Function
        End If
    Next columnIndex
End Function

Private Sub ValidateImportInfoRequiredCell(ByVal rowIndex As Long, ByVal headerName As String, ByVal cellValue As Variant)
    If Len(NormalizeStageText(cellValue)) = 0 Or IsWarningText(cellValue) Then
        RaiseSendRequestError "Не заполнены данные в столбце '" & headerName & "', строка " & rowIndex & " листа '" & IMPORT_INFO_REQUEST_SHEET_NAME & "'. Заполните для продолжения."
    End If
End Sub

Private Function CollectUniqueOrders(ByVal preparedRows As Collection) As Collection
    Dim seenOrders As Object
    Dim rowData As Object
    Dim orderValue As String

    Set CollectUniqueOrders = New Collection
    Set seenOrders = CreateObject("Scripting.Dictionary")
    seenOrders.CompareMode = vbTextCompare

    For Each rowData In preparedRows
        orderValue = CStr(rowData("Order"))
        If Not seenOrders.Exists(orderValue) Then
            seenOrders.Add orderValue, True
            CollectUniqueOrders.Add orderValue
        End If
    Next rowData
End Function

Private Sub ValidateExternalWorkbookPath(ByVal workbookPath As String, ByVal fileLabel As String)
    If Len(Trim$(workbookPath)) = 0 Then
        RaiseSendRequestError "На листе '" & SETTINGS_SHEET_NAME & "' не указан путь к " & fileLabel & "."
    End If

    If InStr(1, workbookPath, "://", vbTextCompare) > 0 Then
        RaiseSendRequestError "Для " & fileLabel & " указан веб-адрес. Для shared-режима нужен файловый путь к синхронизированной папке SharePoint, чтобы макрос мог создать технический lock-файл."
    End If
End Sub

Private Sub EnsureWorkbookClosedInCurrentInstance(ByVal workbookPath As String, ByVal fileLabel As String)
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, workbookPath, vbTextCompare) = 0 Then
            RaiseSendRequestError fileLabel & " уже открыт в текущем Excel. Закройте его и повторите попытку."
        End If
    Next wb
End Sub

Private Sub EnsureLockFolderWritable(ByVal lockPath As String)
    Dim probePath As String
    Dim fileNumber As Integer

    If Len(GetFolderPathFromFilePath(lockPath)) = 0 Then
        RaiseSendRequestError "Не удалось определить папку для технического lock-файла. Проверьте путь к файлу номенклатуры на листе '" & SETTINGS_SHEET_NAME & "'."
    End If

    probePath = lockPath & ".probe"
    fileNumber = FreeFile

    On Error GoTo ErrHandler
    Open probePath For Output Access Write Lock Read Write As #fileNumber
    Print #fileNumber, "probe"
    Close #fileNumber
    Kill probePath
    Exit Sub

ErrHandler:
    On Error Resume Next
    Close #fileNumber
    If Len(Dir$(probePath)) > 0 Then Kill probePath
    On Error GoTo 0
    RaiseSendRequestError "Нет доступа к папке, где должен лежать технический lock-файл для номенклатуры. Проверьте путь к файлу номенклатуры и доступ через Forti/VPN."
End Sub

Private Function ReadExistingOrderSnapshot(ByVal nomenclaturePath As String, ByVal nomenclaturePassword As String, ByVal uniqueOrders As Collection) As Object
    Dim wb As Workbook
    Dim snapshot As Object

    TraceAutomationStep "send:readSnapshot:openReadOnly"
    Set wb = OpenNomenclatureWorkbook(nomenclaturePath, nomenclaturePassword, True)
    TraceAutomationStep "send:readSnapshot:opened"
    Set snapshot = CreateObject("Scripting.Dictionary")
    snapshot.CompareMode = vbTextCompare
    PopulateRequestedOrderSnapshot wb, uniqueOrders, snapshot
    Set ReadExistingOrderSnapshot = snapshot
    TraceAutomationStep "send:readSnapshot:built"
    wb.Close SaveChanges:=False
    TraceAutomationStep "send:readSnapshot:closed"
End Function

Private Sub PopulateRequestedOrderSnapshot(ByVal wb As Workbook, ByVal uniqueOrders As Collection, ByRef snapshot As Object)
    Dim existingOrders As Object
    Dim orderValue As Variant

    TraceAutomationStep "send:populateSnapshot:start"
    Set existingOrders = CreateObject("Scripting.Dictionary")
    existingOrders.CompareMode = vbTextCompare
    PopulateExistingOrdersFromWorkbook wb, existingOrders
    TraceAutomationStep "send:populateSnapshot:existingOrders=" & existingOrders.Count

    For Each orderValue In uniqueOrders
        snapshot(CStr(orderValue)) = existingOrders.Exists(CStr(orderValue))
    Next orderValue
    TraceAutomationStep "send:populateSnapshot:done"
End Sub

Private Sub PopulateExistingOrdersFromWorkbook(ByVal wb As Workbook, ByRef result As Object)
    TraceAutomationStep "send:buildExisting:start"
    TraceAutomationStep "send:buildExisting:loadTK"
    Call LoadOrdersFromTargetSheet(GetWorkbookSheet(wb, TARGET_TK_SHEET_NAME), result)
    TraceAutomationStep "send:buildExisting:loadLA"
    Call LoadOrdersFromTargetSheet(GetWorkbookSheet(wb, TARGET_LA_SHEET_NAME), result)
    TraceAutomationStep "send:buildExisting:done count=" & result.Count
End Sub

Private Sub LoadOrdersFromTargetSheet(ByVal ws As Worksheet, ByVal result As Object)
    Dim orderColumn As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim orderValue As String

    TraceAutomationStep "send:loadOrders:start:" & ws.Name
    orderColumn = FindColumn(ws, TARGET_HEADER_ORDER)
    lastRow = ws.Cells(ws.Rows.Count, orderColumn).End(xlUp).Row
    TraceAutomationStep "send:loadOrders:meta:" & ws.Name & ":col=" & orderColumn & ":lastRow=" & lastRow

    For rowIndex = 2 To lastRow
        orderValue = NormalizeOrderValue(ws.Cells(rowIndex, orderColumn).Value)
        If Len(orderValue) > 0 Then result(orderValue) = True
        If AutomationModeEnabled Then
            If rowIndex Mod 500 = 0 Then
                TraceAutomationStep "send:loadOrders:progress:" & ws.Name & ":" & rowIndex
            End If
        End If
    Next rowIndex
    TraceAutomationStep "send:loadOrders:done:" & ws.Name & ":count=" & result.Count
End Sub

Private Function ReadExistingImportInfoOrderSnapshot(ByVal nomenclaturePath As String, ByVal nomenclaturePassword As String, ByVal uniqueOrders As Collection) As Object
    Dim wb As Workbook
    Dim snapshot As Object

    TraceAutomationStep "import_send:readSnapshot:openReadOnly"
    Set wb = OpenNomenclatureWorkbook(nomenclaturePath, nomenclaturePassword, True)
    TraceAutomationStep "import_send:readSnapshot:opened"
    Set snapshot = CreateObject("Scripting.Dictionary")
    snapshot.CompareMode = vbTextCompare
    PopulateRequestedImportInfoOrderSnapshot wb, uniqueOrders, snapshot
    Set ReadExistingImportInfoOrderSnapshot = snapshot
    TraceAutomationStep "import_send:readSnapshot:built"
    wb.Close SaveChanges:=False
    TraceAutomationStep "import_send:readSnapshot:closed"
End Function

Private Sub PopulateRequestedImportInfoOrderSnapshot(ByVal wb As Workbook, ByVal uniqueOrders As Collection, ByRef snapshot As Object)
    Dim existingOrders As Object
    Dim orderValue As Variant

    Set existingOrders = CreateObject("Scripting.Dictionary")
    existingOrders.CompareMode = vbTextCompare
    LoadOrdersFromTargetSheet GetWorkbookSheet(wb, IMPORT_INFO_TARGET_TK_SHEET_NAME), existingOrders
    LoadOrdersFromTargetSheet GetWorkbookSheet(wb, IMPORT_INFO_TARGET_LA_SHEET_NAME), existingOrders

    For Each orderValue In uniqueOrders
        snapshot(CStr(orderValue)) = existingOrders.Exists(CStr(orderValue))
    Next orderValue
End Sub

Private Sub EnsureArchiveAccessible(ByVal archivePath As String)
    Dim wb As Workbook

    TraceAutomationStep "send:archiveAccessible:openReadOnly"
    Set wb = OpenArchiveWorkbook(archivePath, True)
    TraceAutomationStep "send:archiveAccessible:opened"
    GetArchiveSheet wb
    wb.Close SaveChanges:=False
    TraceAutomationStep "send:archiveAccessible:closed"
End Sub

Private Sub EnsureCorrectionArchiveAccessible(ByVal archivePath As String)
    Dim wb As Workbook

    TraceAutomationStep "correction:archiveAccessible:openReadOnly"
    Set wb = OpenArchiveWorkbook(archivePath, True)
    TraceAutomationStep "correction:archiveAccessible:opened"
    GetCorrectionArchiveSheet wb
    wb.Close SaveChanges:=False
    TraceAutomationStep "correction:archiveAccessible:closed"
End Sub

Private Sub EnsureImportInfoArchiveAccessible(ByVal archivePath As String)
    Dim wb As Workbook

    TraceAutomationStep "import_send:archiveAccessible:openReadOnly"
    Set wb = OpenArchiveWorkbook(archivePath, True)
    TraceAutomationStep "import_send:archiveAccessible:opened"
    GetImportInfoArchiveSheet wb
    wb.Close SaveChanges:=False
    TraceAutomationStep "import_send:archiveAccessible:closed"
End Sub

Private Sub EnsureImportInfoCorrectionArchiveAccessible(ByVal archivePath As String)
    Dim wb As Workbook

    TraceAutomationStep "import_correction:archiveAccessible:openReadOnly"
    Set wb = OpenArchiveWorkbook(archivePath, True)
    TraceAutomationStep "import_correction:archiveAccessible:opened"
    GetImportInfoCorrectionArchiveSheet wb
    wb.Close SaveChanges:=False
    TraceAutomationStep "import_correction:archiveAccessible:closed"
End Sub

Private Function ReadExistingCorrectionDuplicateSnapshot(ByVal nomenclaturePath As String, ByVal nomenclaturePassword As String, ByVal preparedRows As Collection) As Object
    Dim wb As Workbook
    Dim snapshot As Object

    TraceAutomationStep "correction:readSnapshot:openReadOnly"
    Set wb = OpenNomenclatureWorkbook(nomenclaturePath, nomenclaturePassword, True)
    TraceAutomationStep "correction:readSnapshot:opened"
    Set snapshot = CreateObject("Scripting.Dictionary")
    snapshot.CompareMode = vbTextCompare
    PopulateRequestedCorrectionDuplicateSnapshot wb, preparedRows, snapshot
    Set ReadExistingCorrectionDuplicateSnapshot = snapshot
    TraceAutomationStep "correction:readSnapshot:built"
    wb.Close SaveChanges:=False
    TraceAutomationStep "correction:readSnapshot:closed"
End Function

Private Sub PopulateRequestedCorrectionDuplicateSnapshot(ByVal wb As Workbook, ByVal preparedRows As Collection, ByRef snapshot As Object)
    Dim requestedSignatures As Object

    Set requestedSignatures = CreateObject("Scripting.Dictionary")
    requestedSignatures.CompareMode = vbTextCompare
    PopulateRequestedCorrectionSignatures preparedRows, requestedSignatures

    LoadRequestedCorrectionDuplicatesFromSheet GetWorkbookSheet(wb, CORRECTION_TARGET_TK_SHEET_NAME), requestedSignatures, snapshot
    LoadRequestedCorrectionDuplicatesFromSheet GetWorkbookSheet(wb, CORRECTION_TARGET_LA_SHEET_NAME), requestedSignatures, snapshot
End Sub

Private Sub PopulateRequestedCorrectionSignatures(ByVal preparedRows As Collection, ByRef requestedSignatures As Object)
    Dim rowData As Object
    Dim signature As String

    For Each rowData In preparedRows
        signature = BuildCorrectionRowSignature(rowData)
        If Not requestedSignatures.Exists(signature) Then
            requestedSignatures.Add signature, True
        End If
    Next rowData
End Sub

Private Sub LoadRequestedCorrectionDuplicatesFromSheet(ByVal ws As Worksheet, ByVal requestedSignatures As Object, ByRef snapshot As Object)
    Dim targetColumns As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim signature As String

    Set targetColumns = CreateObject("Scripting.Dictionary")
    targetColumns.CompareMode = vbTextCompare
    PopulateCorrectionTargetColumnMap ws, targetColumns

    lastRow = ws.Cells(ws.Rows.Count, CLng(targetColumns(NOMENCLATURE_HEADER_ORDER))).End(xlUp).Row
    For rowIndex = 2 To lastRow
        signature = BuildCorrectionSheetRowSignature(ws, targetColumns, rowIndex)
        If Len(signature) = 0 Then GoTo NextRow
        If requestedSignatures.Exists(signature) Then
            If Not snapshot.Exists(signature) Then
                snapshot.Add signature, rowIndex
            End If
        End If
NextRow:
    Next rowIndex
End Sub

Private Sub PopulateCorrectionDuplicateConfirmations(ByVal preparedRows As Collection, ByVal duplicateSnapshot As Object, ByRef confirmationMap As Object)
    Dim promptedSignatures As Object
    Dim rowData As Object
    Dim signature As String
    Dim responseValue As String
    Dim existingRow As Long

    Set promptedSignatures = CreateObject("Scripting.Dictionary")
    promptedSignatures.CompareMode = vbTextCompare

    For Each rowData In preparedRows
        signature = BuildCorrectionRowSignature(rowData)
        If duplicateSnapshot.Exists(signature) Then
            existingRow = CLng(duplicateSnapshot(signature))
            If Not promptedSignatures.Exists(signature) Then
                responseValue = RequestCorrectionDuplicateConfirmation(rowData, existingRow)
                TraceAutomationStep "correction:confirmations:row=" & existingRow & ":response=" & responseValue
                If StrComp(NormalizeStageText(responseValue), CONFIRMATION_TEXT, vbTextCompare) <> 0 Then
                    RaiseSendRequestError "На листе '" & CStr(rowData("TargetSheet")) & "' уже есть полностью совпадающая строка №" & existingRow & _
                                         ". Не получено подтверждение '" & CONFIRMATION_PROMPT_TEXT & "'. Внесение строк остановлено."
                End If
                promptedSignatures.Add signature, True
            End If
            confirmationMap(signature) = CONFIRMATION_TEXT
        ElseIf Not confirmationMap.Exists(signature) Then
            confirmationMap(signature) = ""
        End If
    Next rowData
End Sub

Private Function ReadExistingImportInfoCorrectionDuplicateSnapshot(ByVal nomenclaturePath As String, ByVal nomenclaturePassword As String, ByVal preparedRows As Collection) As Object
    Dim wb As Workbook
    Dim snapshot As Object

    TraceAutomationStep "import_correction:readSnapshot:openReadOnly"
    Set wb = OpenNomenclatureWorkbook(nomenclaturePath, nomenclaturePassword, True)
    TraceAutomationStep "import_correction:readSnapshot:opened"
    Set snapshot = CreateObject("Scripting.Dictionary")
    snapshot.CompareMode = vbTextCompare
    PopulateRequestedImportInfoCorrectionDuplicateSnapshot wb, preparedRows, snapshot
    Set ReadExistingImportInfoCorrectionDuplicateSnapshot = snapshot
    TraceAutomationStep "import_correction:readSnapshot:built"
    wb.Close SaveChanges:=False
    TraceAutomationStep "import_correction:readSnapshot:closed"
End Function

Private Sub PopulateRequestedImportInfoCorrectionDuplicateSnapshot(ByVal wb As Workbook, ByVal preparedRows As Collection, ByRef snapshot As Object)
    Dim requestedSignatures As Object

    Set requestedSignatures = CreateObject("Scripting.Dictionary")
    requestedSignatures.CompareMode = vbTextCompare
    PopulateRequestedImportInfoCorrectionSignatures preparedRows, requestedSignatures

    LoadRequestedImportInfoCorrectionDuplicatesFromSheet GetWorkbookSheet(wb, IMPORT_INFO_CORRECTION_TARGET_TK_SHEET_NAME), requestedSignatures, snapshot
    LoadRequestedImportInfoCorrectionDuplicatesFromSheet GetWorkbookSheet(wb, IMPORT_INFO_CORRECTION_TARGET_LA_SHEET_NAME), requestedSignatures, snapshot
End Sub

Private Sub PopulateRequestedImportInfoCorrectionSignatures(ByVal preparedRows As Collection, ByRef requestedSignatures As Object)
    Dim rowData As Object
    Dim signature As String

    For Each rowData In preparedRows
        signature = BuildImportInfoCorrectionRowSignature(rowData)
        If Not requestedSignatures.Exists(signature) Then requestedSignatures.Add signature, True
    Next rowData
End Sub

Private Sub LoadRequestedImportInfoCorrectionDuplicatesFromSheet(ByVal ws As Worksheet, ByVal requestedSignatures As Object, ByRef snapshot As Object)
    Dim targetColumns As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim signature As String

    Set targetColumns = CreateObject("Scripting.Dictionary")
    targetColumns.CompareMode = vbTextCompare
    PopulateImportInfoCorrectionTargetColumnMap ws, targetColumns

    lastRow = ws.Cells(ws.Rows.Count, CLng(targetColumns(IMPORT_INFO_HEADER_ORDER))).End(xlUp).Row
    For rowIndex = 2 To lastRow
        signature = BuildImportInfoCorrectionSheetRowSignature(ws, targetColumns, rowIndex)
        If Len(signature) = 0 Then GoTo NextRow
        If requestedSignatures.Exists(signature) Then
            If Not snapshot.Exists(signature) Then snapshot.Add signature, rowIndex
        End If
NextRow:
    Next rowIndex
End Sub

Private Sub PopulateImportInfoCorrectionDuplicateConfirmations(ByVal preparedRows As Collection, ByVal duplicateSnapshot As Object, ByRef confirmationMap As Object)
    Dim rowData As Object
    Dim signature As String
    Dim existingRow As Long
    Dim responseValue As String

    For Each rowData In preparedRows
        signature = BuildImportInfoCorrectionRowSignature(rowData)
        If duplicateSnapshot.Exists(signature) Then
            existingRow = CLng(duplicateSnapshot(signature))
            responseValue = RequestCorrectionDuplicateConfirmation(rowData, existingRow)
            If StrComp(NormalizeStageText(responseValue), CONFIRMATION_TEXT, vbTextCompare) <> 0 Then
                RaiseSendRequestError "На листе '" & CStr(rowData("TargetSheet")) & "' уже есть полностью совпадающая строка для заказа '" & CStr(rowData("Order")) & _
                                      "' в строке " & existingRow & ". Не получено подтверждение '" & CONFIRMATION_PROMPT_TEXT & "'. Внесение строк остановлено."
            End If
            confirmationMap(signature) = CONFIRMATION_TEXT
        ElseIf Not confirmationMap.Exists(signature) Then
            confirmationMap(signature) = ""
        End If
    Next rowData
End Sub

Private Function RequestCorrectionDuplicateConfirmation(ByVal rowData As Object, ByVal existingRow As Long) As String
    Dim promptText As String

    If AutomationModeEnabled Then
        RequestCorrectionDuplicateConfirmation = GetAutomationConfirmationResponse(CStr(rowData("Order")))
        Exit Function
    End If

    promptText = "На листе '" & CStr(rowData("TargetSheet")) & "' уже есть полностью совпадающая строка №" & existingRow & _
                 " для заказа '" & CStr(rowData("Order")) & "'." & vbCrLf & vbCrLf & _
                 "Если всё равно нужно внести такой же дубликат, введите слово '" & CONFIRMATION_PROMPT_TEXT & "'."

    RequestCorrectionDuplicateConfirmation = InputBox(promptText, "Подтверждение корректировки")
End Function

Private Sub PopulateRepeatConfirmations(ByVal uniqueOrders As Collection, ByVal orderSnapshot As Object, ByRef confirmationMap As Object)
    Dim orderValue As Variant
    Dim deliveryCount As Long
    Dim responseValue As String

    For Each orderValue In uniqueOrders
        TraceAutomationStep "send:confirmations:order=" & CStr(orderValue) & ":exists=" & CStr(CBool(orderSnapshot(CStr(orderValue))))
        If CBool(orderSnapshot(CStr(orderValue))) Then
            deliveryCount = CountDistinctDeliveries(CStr(orderValue))
            TraceAutomationStep "send:confirmations:deliveryCount=" & deliveryCount

            responseValue = RequestRepeatConfirmation(CStr(orderValue), deliveryCount)
            TraceAutomationStep "send:confirmations:response=" & responseValue
            If StrComp(NormalizeStageText(responseValue), CONFIRMATION_TEXT, vbTextCompare) <> 0 Then
                RaiseSendRequestError "По заказу '" & CStr(orderValue) & "' не получено подтверждение '" & CONFIRMATION_PROMPT_TEXT & "'. Внесение строк остановлено."
            End If

            confirmationMap(CStr(orderValue)) = CONFIRMATION_TEXT
        Else
            confirmationMap(CStr(orderValue)) = ""
        End If
    Next orderValue
    TraceAutomationStep "send:confirmations:done"
End Sub

Private Function RequestRepeatConfirmation(ByVal orderValue As String, ByVal deliveryCount As Long) As String
    Dim promptText As String

    If AutomationModeEnabled Then
        RequestRepeatConfirmation = GetAutomationConfirmationResponse(orderValue)
        Exit Function
    End If

    promptText = "По заказу '" & orderValue & "' в номенклатуре уже есть существующие строки." & vbCrLf & vbCrLf & _
                 "Если нужно изменить существующий запрос, позже используйте кнопку 'Внести корректировку'." & vbCrLf & _
                 "Если всё равно нужно внести новые строки, введите слово '" & CONFIRMATION_PROMPT_TEXT & "'." 

    If deliveryCount > 0 Then
        promptText = promptText & vbCrLf & "По обработанному Алкоотчету найдено поставок: " & deliveryCount & "."
    Else
        promptText = promptText & vbCrLf & "Количество поставок в обработанном Алкоотчете определить не удалось."
    End If

    RequestRepeatConfirmation = InputBox(promptText, "Подтверждение новой отправки")
End Function

Private Function GetAutomationConfirmationResponse(ByVal confirmationKey As String) As String
    confirmationKey = NormalizeOrderValue(confirmationKey)

    If AutomationConfirmations Is Nothing Then Exit Function
    If AutomationConfirmations.Exists(confirmationKey) Then
        GetAutomationConfirmationResponse = CStr(AutomationConfirmations(confirmationKey))
    End If
End Function

Private Function CountDistinctDeliveries(ByVal orderValue As String) As Long
    Dim ws As Worksheet
    Dim orderColumn As Long
    Dim deliveryColumn As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim currentOrder As String
    Dim deliveryValue As String
    Dim deliveries As Object

    Set ws = GetAlcoReportSheet
    orderColumn = FindColumn(ws, "Номер заказа")
    deliveryColumn = FindColumn(ws, "Номер поставки")
    lastRow = ws.Cells(ws.Rows.Count, orderColumn).End(xlUp).Row
    Set deliveries = CreateObject("Scripting.Dictionary")
    deliveries.CompareMode = vbTextCompare

    For rowIndex = 2 To lastRow
        currentOrder = NormalizeOrderValue(ws.Cells(rowIndex, orderColumn).Value)
        If StrComp(currentOrder, orderValue, vbTextCompare) = 0 Then
            deliveryValue = NormalizeStageText(ws.Cells(rowIndex, deliveryColumn).Value)
            If Len(deliveryValue) > 0 Then deliveries(deliveryValue) = True
        End If
    Next rowIndex

    CountDistinctDeliveries = deliveries.Count
End Function

Private Function ConfirmImportInfoQuantityMismatches(ByVal preparedRows As Collection) As String
    Dim baselineTotals As Object
    Dim currentTotals As Object
    Dim keyLabels As Object
    Dim rowData As Object
    Dim key As String
    Dim mismatchText As String
    Dim keyValue As Variant
    Dim baseQty As Double
    Dim currentQty As Double
    Dim promptText As String
    Dim responseValue As String

    Set baselineTotals = CreateObject("Scripting.Dictionary")
    baselineTotals.CompareMode = vbTextCompare
    Set currentTotals = CreateObject("Scripting.Dictionary")
    currentTotals.CompareMode = vbTextCompare
    Set keyLabels = CreateObject("Scripting.Dictionary")
    keyLabels.CompareMode = vbTextCompare

    LoadImportInfoBaselineTotals baselineTotals, keyLabels

    For Each rowData In preparedRows
        key = BuildImportInfoQuantityKey(rowData("Order"), rowData("Statement"), rowData("Code"))
        If Not currentTotals.Exists(key) Then currentTotals.Add key, 0#
        currentTotals(key) = CDbl(currentTotals(key)) + CDbl(rowData("Qty"))
        If Not keyLabels.Exists(key) Then
            keyLabels.Add key, BuildImportInfoQuantityLabel(rowData("Order"), rowData("Statement"), rowData("Code"))
        End If
    Next rowData

    For Each keyValue In baselineTotals.Keys
        baseQty = CDbl(baselineTotals(CStr(keyValue)))
        If currentTotals.Exists(CStr(keyValue)) Then
            currentQty = CDbl(currentTotals(CStr(keyValue)))
        Else
            currentQty = 0#
        End If
        If Abs(baseQty - currentQty) > 0.000001 Then
            mismatchText = mismatchText & vbCrLf & "- " & CStr(keyLabels(CStr(keyValue))) & ": было " & FormatQuantityForMessage(baseQty) & ", стало " & FormatQuantityForMessage(currentQty)
        End If
    Next keyValue

    For Each keyValue In currentTotals.Keys
        If Not baselineTotals.Exists(CStr(keyValue)) Then
            currentQty = CDbl(currentTotals(CStr(keyValue)))
            mismatchText = mismatchText & vbCrLf & "- " & CStr(keyLabels(CStr(keyValue))) & ": было 0, стало " & FormatQuantityForMessage(currentQty)
        End If
    Next keyValue

    If Len(mismatchText) = 0 Then
        ConfirmImportInfoQuantityMismatches = ""
        Exit Function
    End If

    promptText = "Найдено несоответствие количества после ручной разбивки строк:" & vbCrLf & mismatchText & vbCrLf & vbCrLf & _
                 "Если это осознанное изменение и нужно продолжить, введите слово '" & CONFIRMATION_PROMPT_TEXT & "'."

    If AutomationModeEnabled Then
        responseValue = GetAutomationConfirmationResponse(IMPORT_INFO_QTY_CONFIRMATION_KEY)
    Else
        responseValue = InputBox(promptText, "Подтверждение несоответствия количества")
    End If

    If StrComp(NormalizeStageText(responseValue), CONFIRMATION_TEXT, vbTextCompare) <> 0 Then
        RaiseSendRequestError "Найдено несоответствие количества после ручной разбивки строк:" & vbCrLf & mismatchText & vbCrLf & vbCrLf & _
                              "Не получено подтверждение '" & CONFIRMATION_PROMPT_TEXT & "'. Внесение строк остановлено."
    End If

    ConfirmImportInfoQuantityMismatches = CONFIRMATION_TEXT
End Function

Private Sub LoadImportInfoBaselineTotals(ByRef baselineTotals As Object, ByRef keyLabels As Object)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim orderColumn As Long
    Dim statementColumn As Long
    Dim codeColumn As Long
    Dim qtyColumn As Long
    Dim orderValue As String
    Dim statementValue As String
    Dim codeValue As String
    Dim key As String
    Dim qtyValue As Double
    Dim foundRows As Boolean

    Set ws = GetImportInfoBaselineSheet
    lastRow = FindSheetLastRow(ws)
    If lastRow < 2 Then
        RaiseSendRequestError "Не найден служебный snapshot на листе '" & IMPORT_INFO_BASELINE_SHEET_NAME & "'. Сначала нажмите 'Подготовить строки к внесению в номенклатуру'."
    End If

    orderColumn = FindColumn(ws, IMPORT_INFO_HEADER_ORDER)
    statementColumn = FindColumn(ws, IMPORT_INFO_HEADER_STATEMENT)
    codeColumn = FindColumn(ws, IMPORT_INFO_HEADER_CODE)
    qtyColumn = FindColumn(ws, IMPORT_INFO_HEADER_QTY)

    For rowIndex = 2 To lastRow
        orderValue = NormalizeOrderValue(ws.Cells(rowIndex, orderColumn).Value)
        statementValue = NormalizeStageText(ws.Cells(rowIndex, statementColumn).Value)
        codeValue = NormalizeStageText(ws.Cells(rowIndex, codeColumn).Value)
        If Len(orderValue) = 0 And Len(statementValue) = 0 And Len(codeValue) = 0 Then GoTo NextRow
        foundRows = True
        qtyValue = ParsePreparedImportInfoQuantity(ws.Cells(rowIndex, qtyColumn).Value, orderValue, rowIndex)
        key = BuildImportInfoQuantityKey(orderValue, statementValue, codeValue)
        If Not baselineTotals.Exists(key) Then baselineTotals.Add key, 0#
        baselineTotals(key) = CDbl(baselineTotals(key)) + qtyValue
        If Not keyLabels.Exists(key) Then keyLabels.Add key, BuildImportInfoQuantityLabel(orderValue, statementValue, codeValue)
NextRow:
    Next rowIndex

    If Not foundRows Then
        RaiseSendRequestError "Служебный snapshot на листе '" & IMPORT_INFO_BASELINE_SHEET_NAME & "' пуст. Сначала нажмите 'Подготовить строки к внесению в номенклатуру'."
    End If
End Sub

Private Function BuildImportInfoQuantityKey(ByVal orderValue As Variant, ByVal statementValue As Variant, ByVal codeValue As Variant) As String
    BuildImportInfoQuantityKey = NormalizeOrderValue(orderValue) & "|" & _
                                 NormalizeStageText(statementValue) & "|" & _
                                 NormalizeStageText(codeValue)
End Function

Private Function BuildImportInfoQuantityLabel(ByVal orderValue As Variant, ByVal statementValue As Variant, ByVal codeValue As Variant) As String
    BuildImportInfoQuantityLabel = "заказ '" & NormalizeOrderValue(orderValue) & "', заявление '" & NormalizeStageText(statementValue) & "', код УТ '" & NormalizeStageText(codeValue) & "'"
End Function

Private Function FormatQuantityForMessage(ByVal value As Double) As String
    FormatQuantityForMessage = Format$(value, "0.###############")
End Function

Private Function AcquireMacroLock(ByVal lockPath As String, ByRef lockHandle As Integer) As Boolean
    Dim attempt As Long

    For attempt = 1 To LOCK_RETRY_COUNT
        TraceAutomationStep "send:lockAttempt:" & attempt
        lockHandle = FreeFile
        On Error Resume Next
        Open lockPath For Output Access Write Lock Read Write As #lockHandle
        If Err.Number = 0 Then
            Print #lockHandle, "User=" & Application.UserName
            Print #lockHandle, "Workbook=" & ThisWorkbook.Name
            Print #lockHandle, "Time=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
            On Error GoTo 0
            AcquireMacroLock = True
            Exit Function
        End If

        Err.Clear
        On Error Resume Next
        Close #lockHandle
        On Error GoTo 0

        If attempt < LOCK_RETRY_COUNT Then
            Application.Wait Now + LOCK_RETRY_SECONDS
        End If
    Next attempt
End Function

Private Sub ReleaseMacroLock(ByVal lockHandle As Integer, ByVal lockPath As String)
    On Error Resume Next
    If lockHandle > 0 Then Close #lockHandle
    If Len(lockPath) > 0 Then
        If Len(Dir$(lockPath)) > 0 Then Kill lockPath
    End If
    On Error GoTo 0
End Sub

Private Sub ValidateOrdersStillWritable(ByVal wbNomenclature As Workbook, ByVal uniqueOrders As Collection, ByVal previousSnapshot As Object)
    Dim currentSnapshot As Object
    Dim orderValue As Variant

    Set currentSnapshot = CreateObject("Scripting.Dictionary")
    currentSnapshot.CompareMode = vbTextCompare
    PopulateRequestedOrderSnapshot wbNomenclature, uniqueOrders, currentSnapshot

    For Each orderValue In uniqueOrders
        If Not CBool(previousSnapshot(CStr(orderValue))) And CBool(currentSnapshot(CStr(orderValue))) Then
            Err.Raise SEND_CONFLICT_ERROR_NUMBER, "SendNomenclatureRequest", _
                      "Заказ '" & CStr(orderValue) & "' появился в номенклатуре, пока макрос ожидал доступ к shared-файлу. Повторите подготовку и внесение строк заново."
        End If
    Next orderValue
End Sub

Private Sub ValidateCorrectionRowsStillWritable(ByVal wbNomenclature As Workbook, ByVal preparedRows As Collection, ByVal previousSnapshot As Object)
    Dim currentSnapshot As Object
    Dim seenSignatures As Object
    Dim rowData As Object
    Dim signature As String

    Set currentSnapshot = CreateObject("Scripting.Dictionary")
    currentSnapshot.CompareMode = vbTextCompare
    PopulateRequestedCorrectionDuplicateSnapshot wbNomenclature, preparedRows, currentSnapshot

    Set seenSignatures = CreateObject("Scripting.Dictionary")
    seenSignatures.CompareMode = vbTextCompare

    For Each rowData In preparedRows
        signature = BuildCorrectionRowSignature(rowData)
        If seenSignatures.Exists(signature) Then GoTo NextRow
        seenSignatures.Add signature, True

        If Not previousSnapshot.Exists(signature) And currentSnapshot.Exists(signature) Then
            Err.Raise SEND_CONFLICT_ERROR_NUMBER, "SendNomenclatureCorrectionRequest", _
                      "Точная строка для заказа '" & CStr(rowData("Order")) & "' уже появилась на листе '" & CStr(rowData("TargetSheet")) & _
                      "' пока макрос ожидал доступ к shared-файлу. Повторите попытку ещё раз."
        End If
NextRow:
    Next rowData
End Sub

Private Sub ValidateImportInfoOrdersStillWritable(ByVal wbNomenclature As Workbook, ByVal uniqueOrders As Collection, ByVal previousSnapshot As Object)
    Dim currentSnapshot As Object
    Dim orderValue As Variant

    Set currentSnapshot = CreateObject("Scripting.Dictionary")
    currentSnapshot.CompareMode = vbTextCompare
    PopulateRequestedImportInfoOrderSnapshot wbNomenclature, uniqueOrders, currentSnapshot

    For Each orderValue In uniqueOrders
        If Not CBool(previousSnapshot(CStr(orderValue))) And CBool(currentSnapshot(CStr(orderValue))) Then
            Err.Raise SEND_CONFLICT_ERROR_NUMBER, "SendImportInfoRequest", _
                      "Заказ '" & CStr(orderValue) & "' появился в листах фиксации сведений о ввозе, пока макрос ожидал доступ к shared-файлу. Повторите подготовку и внесение строк заново."
        End If
    Next orderValue
End Sub

Private Sub ValidateImportInfoCorrectionRowsStillWritable(ByVal wbNomenclature As Workbook, ByVal preparedRows As Collection, ByVal previousSnapshot As Object)
    Dim currentSnapshot As Object
    Dim seenSignatures As Object
    Dim rowData As Object
    Dim signature As String

    Set currentSnapshot = CreateObject("Scripting.Dictionary")
    currentSnapshot.CompareMode = vbTextCompare
    PopulateRequestedImportInfoCorrectionDuplicateSnapshot wbNomenclature, preparedRows, currentSnapshot

    Set seenSignatures = CreateObject("Scripting.Dictionary")
    seenSignatures.CompareMode = vbTextCompare

    For Each rowData In preparedRows
        signature = BuildImportInfoCorrectionRowSignature(rowData)
        If seenSignatures.Exists(signature) Then GoTo NextRow
        seenSignatures.Add signature, True

        If Not previousSnapshot.Exists(signature) And currentSnapshot.Exists(signature) Then
            Err.Raise SEND_CONFLICT_ERROR_NUMBER, "SendImportInfoCorrectionRequest", _
                      "Точная строка для заказа '" & CStr(rowData("Order")) & "' уже появилась на листе '" & CStr(rowData("TargetSheet")) & _
                      "' пока макрос ожидал доступ к shared-файлу. Повторите попытку ещё раз."
        End If
NextRow:
    Next rowData
End Sub

Private Sub WritePreparedRowsToExternalBooks(ByVal wbNomenclature As Workbook, ByVal wbArchive As Workbook, ByVal preparedRows As Collection, ByVal confirmationMap As Object, ByVal targetProtectionStates As Object, ByRef archiveStartRow As Long, ByRef archiveRowCount As Long, ByVal nomenclaturePassword As String)
    Dim archiveSheet As Worksheet
    Dim archiveColumns As Object
    Dim targetSheets As Object
    Dim targetColumns As Object
    Dim rowData As Object
    Dim targetSheetName As String
    Dim targetSheet As Worksheet
    Dim targetRow As Long
    Dim archiveRow As Long
    Dim currentDate As Date

    TraceAutomationStep "send:writeRows:start"
    Set archiveSheet = GetArchiveSheet(wbArchive)
    TraceAutomationStep "send:writeRows:archiveSheet"
    Set archiveColumns = CreateObject("Scripting.Dictionary")
    archiveColumns.CompareMode = vbTextCompare
    PopulateArchiveColumnMap archiveSheet, archiveColumns
    TraceAutomationStep "send:writeRows:archiveColumnsReady"
    Set targetSheets = CreateObject("Scripting.Dictionary")
    targetSheets.CompareMode = vbTextCompare

    archiveStartRow = GetNextAppendRow(archiveSheet, 1)
    archiveRowCount = preparedRows.Count
    archiveRow = archiveStartRow
    currentDate = Date
    TraceAutomationStep "send:writeRows:meta:startRow=" & archiveStartRow & ":count=" & archiveRowCount

    For Each rowData In preparedRows
        targetSheetName = CStr(rowData("TargetSheet"))
        Set targetSheet = GetWorkbookSheet(wbNomenclature, targetSheetName)
        TraceAutomationStep "send:writeRows:targetSheet=" & targetSheetName

        If Not targetSheets.Exists(targetSheetName) Then
            Set targetColumns = CreateObject("Scripting.Dictionary")
            targetColumns.CompareMode = vbTextCompare
            PopulateTargetColumnMap targetSheet, targetColumns
            targetSheets.Add targetSheetName, targetColumns
            TraceAutomationStep "send:writeRows:targetColumnsReady=" & targetSheetName
        Else
            Set targetColumns = targetSheets(targetSheetName)
        End If

        EnsureTargetSheetWritable targetSheet, targetProtectionStates, nomenclaturePassword
        targetRow = AppendTargetRow(targetSheet)
        TraceAutomationStep "send:writeRows:targetRow=" & targetRow
        FillTargetRow targetSheet, targetColumns, targetRow, rowData, currentDate
        FillArchiveRow archiveSheet, archiveColumns, archiveRow, rowData, currentDate, CStr(confirmationMap(CStr(rowData("Order"))))
        TraceAutomationStep "send:writeRows:archiveRow=" & archiveRow
        archiveRow = archiveRow + 1
    Next rowData
    TraceAutomationStep "send:writeRows:done"
End Sub

Private Sub WritePreparedCorrectionRowsToExternalBooks(ByVal wbNomenclature As Workbook, ByVal wbArchive As Workbook, ByVal preparedRows As Collection, ByVal targetProtectionStates As Object, ByRef archiveStartRow As Long, ByRef archiveRowCount As Long, ByVal protectedSheetsPassword As String)
    Dim archiveSheet As Worksheet
    Dim archiveColumns As Object
    Dim targetSheets As Object
    Dim targetColumns As Object
    Dim rowData As Object
    Dim targetSheetName As String
    Dim targetSheet As Worksheet
    Dim targetRow As Long
    Dim archiveRow As Long
    Dim currentDate As Date

    TraceAutomationStep "correction:writeRows:start"
    Set archiveSheet = GetCorrectionArchiveSheet(wbArchive)
    TraceAutomationStep "correction:writeRows:archiveSheet"
    Set archiveColumns = CreateObject("Scripting.Dictionary")
    archiveColumns.CompareMode = vbTextCompare
    PopulateCorrectionArchiveColumnMap archiveSheet, archiveColumns
    TraceAutomationStep "correction:writeRows:archiveColumnsReady"
    Set targetSheets = CreateObject("Scripting.Dictionary")
    targetSheets.CompareMode = vbTextCompare

    archiveStartRow = GetNextAppendRow(archiveSheet, 1)
    archiveRowCount = preparedRows.Count
    archiveRow = archiveStartRow
    currentDate = Date
    TraceAutomationStep "correction:writeRows:meta:startRow=" & archiveStartRow & ":count=" & archiveRowCount

    For Each rowData In preparedRows
        targetSheetName = CStr(rowData("TargetSheet"))
        Set targetSheet = GetWorkbookSheet(wbNomenclature, targetSheetName)
        TraceAutomationStep "correction:writeRows:targetSheet=" & targetSheetName

        If Not targetSheets.Exists(targetSheetName) Then
            Set targetColumns = CreateObject("Scripting.Dictionary")
            targetColumns.CompareMode = vbTextCompare
            PopulateCorrectionTargetColumnMap targetSheet, targetColumns
            targetSheets.Add targetSheetName, targetColumns
            TraceAutomationStep "correction:writeRows:targetColumnsReady=" & targetSheetName
        Else
            Set targetColumns = targetSheets(targetSheetName)
        End If

        EnsureTargetSheetWritable targetSheet, targetProtectionStates, protectedSheetsPassword
        targetRow = AppendTargetRow(targetSheet)
        TraceAutomationStep "correction:writeRows:targetRow=" & targetRow
        FillCorrectionTargetRow targetSheet, targetColumns, targetRow, rowData
        FillCorrectionArchiveRow archiveSheet, archiveColumns, archiveRow, rowData, currentDate
        TraceAutomationStep "correction:writeRows:archiveRow=" & archiveRow
        archiveRow = archiveRow + 1
    Next rowData
    TraceAutomationStep "correction:writeRows:done"
End Sub

Private Sub WriteImportInfoRowsToExternalBooks(ByVal wbNomenclature As Workbook, ByVal wbArchive As Workbook, ByVal preparedRows As Collection, ByVal confirmationMap As Object, ByVal quantityMismatchConfirmation As String, ByVal targetProtectionStates As Object, ByRef archiveStartRow As Long, ByRef archiveRowCount As Long, ByVal nomenclaturePassword As String)
    Dim archiveSheet As Worksheet
    Dim archiveColumns As Object
    Dim targetSheets As Object
    Dim targetColumns As Object
    Dim rowData As Object
    Dim targetSheetName As String
    Dim targetSheet As Worksheet
    Dim targetRow As Long
    Dim archiveRow As Long
    Dim currentDate As Date

    Set archiveSheet = GetImportInfoArchiveSheet(wbArchive)
    Set archiveColumns = CreateObject("Scripting.Dictionary")
    archiveColumns.CompareMode = vbTextCompare
    PopulateImportInfoArchiveColumnMap archiveSheet, archiveColumns
    Set targetSheets = CreateObject("Scripting.Dictionary")
    targetSheets.CompareMode = vbTextCompare

    archiveStartRow = GetNextAppendRow(archiveSheet, 1)
    archiveRowCount = preparedRows.Count
    archiveRow = archiveStartRow
    currentDate = Date

    For Each rowData In preparedRows
        targetSheetName = CStr(rowData("TargetSheet"))
        Set targetSheet = GetWorkbookSheet(wbNomenclature, targetSheetName)

        If Not targetSheets.Exists(targetSheetName) Then
            Set targetColumns = CreateObject("Scripting.Dictionary")
            targetColumns.CompareMode = vbTextCompare
            PopulateImportInfoTargetColumnMap targetSheet, targetColumns
            targetSheets.Add targetSheetName, targetColumns
        Else
            Set targetColumns = targetSheets(targetSheetName)
        End If

        EnsureTargetSheetWritable targetSheet, targetProtectionStates, nomenclaturePassword
        targetRow = AppendTargetRow(targetSheet)
        FillImportInfoTargetRow targetSheet, targetColumns, targetRow, rowData, currentDate
        FillImportInfoArchiveRow archiveSheet, archiveColumns, archiveRow, rowData, currentDate, CStr(confirmationMap(CStr(rowData("Order")))), quantityMismatchConfirmation
        archiveRow = archiveRow + 1
    Next rowData
End Sub

Private Sub WriteImportInfoCorrectionRowsToExternalBooks(ByVal wbNomenclature As Workbook, ByVal wbArchive As Workbook, ByVal preparedRows As Collection, ByVal quantityMismatchConfirmation As String, ByVal targetProtectionStates As Object, ByRef archiveStartRow As Long, ByRef archiveRowCount As Long, ByVal protectedSheetsPassword As String)
    Dim archiveSheet As Worksheet
    Dim archiveColumns As Object
    Dim targetSheets As Object
    Dim targetColumns As Object
    Dim rowData As Object
    Dim targetSheetName As String
    Dim targetSheet As Worksheet
    Dim targetRow As Long
    Dim archiveRow As Long
    Dim currentDate As Date

    Set archiveSheet = GetImportInfoCorrectionArchiveSheet(wbArchive)
    Set archiveColumns = CreateObject("Scripting.Dictionary")
    archiveColumns.CompareMode = vbTextCompare
    PopulateImportInfoCorrectionArchiveColumnMap archiveSheet, archiveColumns
    Set targetSheets = CreateObject("Scripting.Dictionary")
    targetSheets.CompareMode = vbTextCompare

    archiveStartRow = GetNextAppendRow(archiveSheet, 1)
    archiveRowCount = preparedRows.Count
    archiveRow = archiveStartRow
    currentDate = Date

    For Each rowData In preparedRows
        targetSheetName = CStr(rowData("TargetSheet"))
        Set targetSheet = GetWorkbookSheet(wbNomenclature, targetSheetName)

        If Not targetSheets.Exists(targetSheetName) Then
            Set targetColumns = CreateObject("Scripting.Dictionary")
            targetColumns.CompareMode = vbTextCompare
            PopulateImportInfoCorrectionTargetColumnMap targetSheet, targetColumns
            targetSheets.Add targetSheetName, targetColumns
        Else
            Set targetColumns = targetSheets(targetSheetName)
        End If

        EnsureTargetSheetWritable targetSheet, targetProtectionStates, protectedSheetsPassword
        targetRow = AppendTargetRow(targetSheet)
        FillImportInfoCorrectionTargetRow targetSheet, targetColumns, targetRow, rowData
        FillImportInfoCorrectionArchiveRow archiveSheet, archiveColumns, archiveRow, rowData, currentDate, quantityMismatchConfirmation
        archiveRow = archiveRow + 1
    Next rowData
End Sub

Private Sub EnsureTargetSheetWritable(ByVal ws As Worksheet, ByVal protectionStates As Object, ByVal nomenclaturePassword As String)
    Dim state As Object

    If protectionStates.Exists(ws.Name) Then Exit Sub

    Set state = CreateObject("Scripting.Dictionary")
    state.CompareMode = vbTextCompare
    state("WasProtected") = CBool(ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios)
    state("EnableSelection") = CLng(ws.EnableSelection)

    If CBool(state("WasProtected")) Then
        With ws.Protection
            state("AllowFormattingCells") = CBool(.AllowFormattingCells)
            state("AllowFormattingColumns") = CBool(.AllowFormattingColumns)
            state("AllowFormattingRows") = CBool(.AllowFormattingRows)
            state("AllowInsertingColumns") = CBool(.AllowInsertingColumns)
            state("AllowInsertingRows") = CBool(.AllowInsertingRows)
            state("AllowDeletingColumns") = CBool(.AllowDeletingColumns)
            state("AllowDeletingRows") = CBool(.AllowDeletingRows)
            state("AllowSorting") = CBool(.AllowSorting)
            state("AllowFiltering") = CBool(.AllowFiltering)
            state("AllowUsingPivotTables") = CBool(.AllowUsingPivotTables)
        End With
        ws.Unprotect Password:=nomenclaturePassword
        TraceAutomationStep "send:sheetUnprotected:" & ws.Name
    End If

    protectionStates.Add ws.Name, state
End Sub

Private Sub RestoreTargetSheetProtection(ByVal wbNomenclature As Workbook, ByVal protectionStates As Object, ByVal nomenclaturePassword As String)
    Dim sheetName As Variant
    Dim state As Object
    Dim ws As Worksheet

    If protectionStates Is Nothing Then Exit Sub

    For Each sheetName In protectionStates.Keys
        Set state = protectionStates(CStr(sheetName))
        If CBool(state("WasProtected")) Then
            Set ws = GetWorkbookSheet(wbNomenclature, CStr(sheetName))
            ws.Protect Password:=nomenclaturePassword, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                       UserInterfaceOnly:=False, AllowFormattingCells:=CBool(state("AllowFormattingCells")), _
                       AllowFormattingColumns:=CBool(state("AllowFormattingColumns")), AllowFormattingRows:=CBool(state("AllowFormattingRows")), _
                       AllowInsertingColumns:=CBool(state("AllowInsertingColumns")), AllowInsertingRows:=CBool(state("AllowInsertingRows")), _
                       AllowInsertingHyperlinks:=False, AllowDeletingColumns:=CBool(state("AllowDeletingColumns")), _
                       AllowDeletingRows:=CBool(state("AllowDeletingRows")), AllowSorting:=CBool(state("AllowSorting")), _
                       AllowFiltering:=CBool(state("AllowFiltering")), AllowUsingPivotTables:=CBool(state("AllowUsingPivotTables"))
            ws.EnableSelection = CLng(state("EnableSelection"))
            TraceAutomationStep "send:sheetProtected:" & ws.Name
        End If
    Next sheetName
End Sub

Private Sub PopulateTargetColumnMap(ByVal ws As Worksheet, ByRef targetColumns As Object)
    TraceAutomationStep "send:buildTargetMap:start:" & ws.Name
    targetColumns(TARGET_HEADER_ARTICLE) = FindColumn(ws, TARGET_HEADER_ARTICLE)
    targetColumns(TARGET_HEADER_NAME) = FindColumn(ws, TARGET_HEADER_NAME)
    targetColumns(TARGET_HEADER_ORDER) = FindColumn(ws, TARGET_HEADER_ORDER)
    targetColumns(TARGET_HEADER_STATEMENT) = FindColumn(ws, TARGET_HEADER_STATEMENT)
    targetColumns(TARGET_HEADER_QTY) = FindColumn(ws, TARGET_HEADER_QTY)
    targetColumns(TARGET_HEADER_STICKER) = FindColumn(ws, TARGET_HEADER_STICKER)
    targetColumns(TARGET_HEADER_DATE) = FindColumn(ws, TARGET_HEADER_DATE)
    targetColumns(TARGET_HEADER_SUPPLIER) = FindColumn(ws, TARGET_HEADER_SUPPLIER)
    targetColumns(TARGET_HEADER_MIL_COMMENT) = FindColumn(ws, TARGET_HEADER_MIL_COMMENT)
    targetColumns(TARGET_HEADER_MIL) = FindColumn(ws, TARGET_HEADER_MIL)
    TraceAutomationStep "send:buildTargetMap:done:" & ws.Name
End Sub

Private Sub PopulateCorrectionTargetColumnMap(ByVal ws As Worksheet, ByRef targetColumns As Object)
    TraceAutomationStep "correction:buildTargetMap:start:" & ws.Name
    targetColumns(NOMENCLATURE_HEADER_ORDER) = FindColumn(ws, NOMENCLATURE_HEADER_ORDER)
    targetColumns(NOMENCLATURE_HEADER_CODE) = FindColumn(ws, NOMENCLATURE_HEADER_CODE)
    targetColumns(NOMENCLATURE_HEADER_NAME) = FindColumn(ws, NOMENCLATURE_HEADER_NAME)
    targetColumns(NOMENCLATURE_HEADER_STATEMENT) = FindColumn(ws, NOMENCLATURE_HEADER_STATEMENT)
    targetColumns(NOMENCLATURE_HEADER_QTY) = FindColumn(ws, NOMENCLATURE_HEADER_QTY)
    targetColumns(NOMENCLATURE_HEADER_STICKER) = FindColumn(ws, NOMENCLATURE_HEADER_STICKER)
    targetColumns(NOMENCLATURE_HEADER_SUPPLIER) = FindColumn(ws, NOMENCLATURE_HEADER_SUPPLIER)
    targetColumns(NOMENCLATURE_HEADER_MIL_COMMENT) = FindColumn(ws, NOMENCLATURE_HEADER_MIL_COMMENT)
    targetColumns(NOMENCLATURE_HEADER_MIL) = FindColumn(ws, NOMENCLATURE_HEADER_MIL)
    targetColumns(CORRECTION_STATUS_HEADER) = FindColumn(ws, CORRECTION_STATUS_HEADER)
    TraceAutomationStep "correction:buildTargetMap:done:" & ws.Name
End Sub

Private Sub PopulateArchiveColumnMap(ByVal ws As Worksheet, ByRef archiveColumns As Object)
    TraceAutomationStep "send:buildArchiveMap:start"
    archiveColumns(NOMENCLATURE_HEADER_ORDER) = FindColumn(ws, NOMENCLATURE_HEADER_ORDER)
    archiveColumns(NOMENCLATURE_HEADER_CODE) = FindColumn(ws, NOMENCLATURE_HEADER_CODE)
    archiveColumns(NOMENCLATURE_HEADER_NAME) = FindColumn(ws, NOMENCLATURE_HEADER_NAME)
    archiveColumns(NOMENCLATURE_HEADER_STATEMENT) = FindColumn(ws, NOMENCLATURE_HEADER_STATEMENT)
    archiveColumns(NOMENCLATURE_HEADER_QTY) = FindColumn(ws, NOMENCLATURE_HEADER_QTY)
    archiveColumns(NOMENCLATURE_HEADER_STICKER) = FindColumn(ws, NOMENCLATURE_HEADER_STICKER)
    archiveColumns(NOMENCLATURE_HEADER_SUPPLIER) = FindColumn(ws, NOMENCLATURE_HEADER_SUPPLIER)
    archiveColumns(NOMENCLATURE_HEADER_MIL_COMMENT) = FindColumn(ws, NOMENCLATURE_HEADER_MIL_COMMENT)
    archiveColumns(NOMENCLATURE_HEADER_MIL) = FindColumn(ws, NOMENCLATURE_HEADER_MIL)
    archiveColumns(ARCHIVE_HEADER_DATE) = FindColumn(ws, ARCHIVE_HEADER_DATE)
    archiveColumns(ARCHIVE_HEADER_CONFIRMATION) = FindColumn(ws, ARCHIVE_HEADER_CONFIRMATION)
    TraceAutomationStep "send:buildArchiveMap:done"
End Sub

Private Sub PopulateCorrectionArchiveColumnMap(ByVal ws As Worksheet, ByRef archiveColumns As Object)
    TraceAutomationStep "correction:buildArchiveMap:start"
    archiveColumns(NOMENCLATURE_HEADER_ORDER) = FindColumn(ws, NOMENCLATURE_HEADER_ORDER)
    archiveColumns(NOMENCLATURE_HEADER_CODE) = FindColumn(ws, NOMENCLATURE_HEADER_CODE)
    archiveColumns(NOMENCLATURE_HEADER_NAME) = FindColumn(ws, NOMENCLATURE_HEADER_NAME)
    archiveColumns(NOMENCLATURE_HEADER_STATEMENT) = FindColumn(ws, NOMENCLATURE_HEADER_STATEMENT)
    archiveColumns(NOMENCLATURE_HEADER_QTY) = FindColumn(ws, NOMENCLATURE_HEADER_QTY)
    archiveColumns(NOMENCLATURE_HEADER_STICKER) = FindColumn(ws, NOMENCLATURE_HEADER_STICKER)
    archiveColumns(NOMENCLATURE_HEADER_SUPPLIER) = FindColumn(ws, NOMENCLATURE_HEADER_SUPPLIER)
    archiveColumns(NOMENCLATURE_HEADER_MIL_COMMENT) = FindColumn(ws, NOMENCLATURE_HEADER_MIL_COMMENT)
    archiveColumns(NOMENCLATURE_HEADER_MIL) = FindColumn(ws, NOMENCLATURE_HEADER_MIL)
    archiveColumns(ARCHIVE_HEADER_DATE) = FindColumn(ws, ARCHIVE_HEADER_DATE)
    TraceAutomationStep "correction:buildArchiveMap:done"
End Sub

Private Sub PopulateImportInfoTargetColumnMap(ByVal ws As Worksheet, ByRef targetColumns As Object)
    targetColumns(TARGET_HEADER_ARTICLE) = FindColumn(ws, TARGET_HEADER_ARTICLE)
    targetColumns(TARGET_HEADER_NAME) = FindColumn(ws, TARGET_HEADER_NAME)
    targetColumns(TARGET_HEADER_ORDER) = FindColumn(ws, TARGET_HEADER_ORDER)
    targetColumns(TARGET_HEADER_STATEMENT) = FindColumn(ws, TARGET_HEADER_STATEMENT)
    targetColumns(TARGET_HEADER_QTY) = FindColumn(ws, TARGET_HEADER_QTY)
    targetColumns(TARGET_HEADER_STICKER) = FindColumn(ws, TARGET_HEADER_STICKER)
    targetColumns(TARGET_HEADER_SHIPMENT_INVOICE) = FindColumn(ws, TARGET_HEADER_SHIPMENT_INVOICE)
    targetColumns(TARGET_HEADER_SHIPMENT_DATE) = FindColumn(ws, TARGET_HEADER_SHIPMENT_DATE)
    targetColumns(TARGET_HEADER_ALCOHOL) = FindColumn(ws, TARGET_HEADER_ALCOHOL)
    targetColumns(TARGET_HEADER_VOLUME) = FindColumn(ws, TARGET_HEADER_VOLUME)
    targetColumns(TARGET_HEADER_VINTAGE) = FindColumn(ws, TARGET_HEADER_VINTAGE)
    targetColumns(TARGET_HEADER_DATE) = FindColumn(ws, TARGET_HEADER_DATE)
    targetColumns(TARGET_HEADER_SUPPLIER) = FindColumn(ws, TARGET_HEADER_SUPPLIER)
    targetColumns(TARGET_HEADER_MIL_COMMENT) = FindColumn(ws, TARGET_HEADER_MIL_COMMENT)
    targetColumns(TARGET_HEADER_MIL) = FindColumn(ws, TARGET_HEADER_MIL)
End Sub

Private Sub PopulateImportInfoCorrectionTargetColumnMap(ByVal ws As Worksheet, ByRef targetColumns As Object)
    targetColumns(IMPORT_INFO_HEADER_ORDER) = FindColumn(ws, IMPORT_INFO_HEADER_ORDER)
    targetColumns(IMPORT_INFO_HEADER_SUPPLIER) = FindColumn(ws, IMPORT_INFO_HEADER_SUPPLIER)
    targetColumns(IMPORT_INFO_HEADER_CODE) = FindColumn(ws, IMPORT_INFO_HEADER_CODE)
    targetColumns(IMPORT_INFO_HEADER_NAME) = FindColumn(ws, IMPORT_INFO_HEADER_NAME)
    targetColumns(IMPORT_INFO_HEADER_STATEMENT) = FindColumn(ws, IMPORT_INFO_HEADER_STATEMENT)
    targetColumns(IMPORT_INFO_HEADER_QTY) = FindColumn(ws, IMPORT_INFO_HEADER_QTY)
    targetColumns(IMPORT_INFO_HEADER_STICKER) = FindColumn(ws, IMPORT_INFO_HEADER_STICKER)
    targetColumns(IMPORT_INFO_HEADER_SHIPMENT_INVOICE) = FindColumn(ws, IMPORT_INFO_HEADER_SHIPMENT_INVOICE)
    targetColumns(IMPORT_INFO_HEADER_SHIPMENT_DATE) = FindColumn(ws, IMPORT_INFO_HEADER_SHIPMENT_DATE)
    targetColumns(IMPORT_INFO_HEADER_ALCOHOL) = FindColumn(ws, IMPORT_INFO_HEADER_ALCOHOL)
    targetColumns(IMPORT_INFO_HEADER_VOLUME) = FindColumn(ws, IMPORT_INFO_HEADER_VOLUME)
    targetColumns(IMPORT_INFO_HEADER_VINTAGE) = FindColumn(ws, IMPORT_INFO_HEADER_VINTAGE)
    targetColumns(IMPORT_INFO_HEADER_MIL) = FindColumn(ws, IMPORT_INFO_HEADER_MIL)
    targetColumns(IMPORT_INFO_HEADER_MIL_COMMENT) = FindColumn(ws, IMPORT_INFO_HEADER_MIL_COMMENT)
    targetColumns(CORRECTION_STATUS_HEADER) = FindColumn(ws, CORRECTION_STATUS_HEADER)
End Sub

Private Sub PopulateImportInfoArchiveColumnMap(ByVal ws As Worksheet, ByRef archiveColumns As Object)
    PopulateImportInfoBaseArchiveColumnMap ws, archiveColumns
    archiveColumns(ARCHIVE_HEADER_DATE) = FindColumn(ws, ARCHIVE_HEADER_DATE)
    archiveColumns(ARCHIVE_HEADER_CONFIRMATION) = FindColumn(ws, ARCHIVE_HEADER_CONFIRMATION)
    archiveColumns(ARCHIVE_HEADER_QTY_MISMATCH_CONFIRMATION) = FindColumn(ws, ARCHIVE_HEADER_QTY_MISMATCH_CONFIRMATION)
End Sub

Private Sub PopulateImportInfoCorrectionArchiveColumnMap(ByVal ws As Worksheet, ByRef archiveColumns As Object)
    PopulateImportInfoBaseArchiveColumnMap ws, archiveColumns
    archiveColumns(ARCHIVE_HEADER_DATE) = FindColumn(ws, ARCHIVE_HEADER_DATE)
    archiveColumns(ARCHIVE_HEADER_QTY_MISMATCH_CONFIRMATION) = FindColumn(ws, ARCHIVE_HEADER_QTY_MISMATCH_CONFIRMATION)
End Sub

Private Sub PopulateImportInfoBaseArchiveColumnMap(ByVal ws As Worksheet, ByRef archiveColumns As Object)
    archiveColumns(IMPORT_INFO_HEADER_ORDER) = FindColumn(ws, IMPORT_INFO_HEADER_ORDER)
    archiveColumns(IMPORT_INFO_HEADER_SUPPLIER) = FindColumn(ws, IMPORT_INFO_HEADER_SUPPLIER)
    archiveColumns(IMPORT_INFO_HEADER_CODE) = FindColumn(ws, IMPORT_INFO_HEADER_CODE)
    archiveColumns(IMPORT_INFO_HEADER_NAME) = FindColumn(ws, IMPORT_INFO_HEADER_NAME)
    archiveColumns(IMPORT_INFO_HEADER_STATEMENT) = FindColumn(ws, IMPORT_INFO_HEADER_STATEMENT)
    archiveColumns(IMPORT_INFO_HEADER_QTY) = FindColumn(ws, IMPORT_INFO_HEADER_QTY)
    archiveColumns(IMPORT_INFO_HEADER_STICKER) = FindColumn(ws, IMPORT_INFO_HEADER_STICKER)
    archiveColumns(IMPORT_INFO_HEADER_SHIPMENT_INVOICE) = FindColumn(ws, IMPORT_INFO_HEADER_SHIPMENT_INVOICE)
    archiveColumns(IMPORT_INFO_HEADER_SHIPMENT_DATE) = FindColumn(ws, IMPORT_INFO_HEADER_SHIPMENT_DATE)
    archiveColumns(IMPORT_INFO_HEADER_ALCOHOL) = FindColumn(ws, IMPORT_INFO_HEADER_ALCOHOL)
    archiveColumns(IMPORT_INFO_HEADER_VOLUME) = FindColumn(ws, IMPORT_INFO_HEADER_VOLUME)
    archiveColumns(IMPORT_INFO_HEADER_VINTAGE) = FindColumn(ws, IMPORT_INFO_HEADER_VINTAGE)
    archiveColumns(IMPORT_INFO_HEADER_MIL) = FindColumn(ws, IMPORT_INFO_HEADER_MIL)
    archiveColumns(IMPORT_INFO_HEADER_MIL_COMMENT) = FindColumn(ws, IMPORT_INFO_HEADER_MIL_COMMENT)
End Sub

Private Function AppendTargetRow(ByVal ws As Worksheet) As Long
    Dim listRow As ListRow
    Dim targetRow As Long

    If ws.ListObjects.Count > 0 Then
        Set listRow = ws.ListObjects(1).ListRows.Add
        AppendTargetRow = listRow.Range.Row
        Exit Function
    End If

    targetRow = GetNextAppendRow(ws, 4)
    If targetRow > 2 Then
        ws.Rows(targetRow - 1).Copy Destination:=ws.Rows(targetRow)
    End If
    AppendTargetRow = targetRow
End Function

Private Sub FillTargetRow(ByVal ws As Worksheet, ByVal targetColumns As Object, ByVal targetRow As Long, ByVal rowData As Object, ByVal currentDate As Date)
    ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow, 16)).ClearContents
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_ARTICLE))).Value = CStr(rowData("Code"))
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_NAME))).Value = CStr(rowData("Name"))
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_ORDER))).Value = CStr(rowData("Order"))
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_STATEMENT))).Value = CStr(rowData("Statement"))
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_QTY))).Value = CDbl(rowData("Qty"))
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_STICKER))).Value = CStr(rowData("Sticker"))
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_DATE))).Value = currentDate
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_SUPPLIER))).Value = CStr(rowData("Supplier"))
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_MIL_COMMENT))).Value = CStr(rowData("MilComment"))
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_MIL))).Value = CStr(rowData("Mil"))
End Sub

Private Sub FillCorrectionTargetRow(ByVal ws As Worksheet, ByVal targetColumns As Object, ByVal targetRow As Long, ByVal rowData As Object)
    ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow, 9)).ClearContents
    ws.Cells(targetRow, CLng(targetColumns(NOMENCLATURE_HEADER_ORDER))).Value = CStr(rowData("Order"))
    ws.Cells(targetRow, CLng(targetColumns(NOMENCLATURE_HEADER_CODE))).Value = CStr(rowData("Code"))
    ws.Cells(targetRow, CLng(targetColumns(NOMENCLATURE_HEADER_NAME))).Value = CStr(rowData("Name"))
    ws.Cells(targetRow, CLng(targetColumns(NOMENCLATURE_HEADER_STATEMENT))).Value = CStr(rowData("Statement"))
    ws.Cells(targetRow, CLng(targetColumns(NOMENCLATURE_HEADER_QTY))).Value = CDbl(rowData("Qty"))
    ws.Cells(targetRow, CLng(targetColumns(NOMENCLATURE_HEADER_STICKER))).Value = CStr(rowData("Sticker"))
    ws.Cells(targetRow, CLng(targetColumns(NOMENCLATURE_HEADER_SUPPLIER))).Value = CStr(rowData("Supplier"))
    ws.Cells(targetRow, CLng(targetColumns(NOMENCLATURE_HEADER_MIL_COMMENT))).Value = CStr(rowData("MilComment"))
    ws.Cells(targetRow, CLng(targetColumns(NOMENCLATURE_HEADER_MIL))).Value = CStr(rowData("Mil"))
    ws.Cells(targetRow, CLng(targetColumns(CORRECTION_STATUS_HEADER))).Value = CORRECTION_STATUS_PENDING
End Sub

Private Sub FillArchiveRow(ByVal ws As Worksheet, ByVal archiveColumns As Object, ByVal archiveRow As Long, ByVal rowData As Object, ByVal currentDate As Date, ByVal confirmationValue As String)
    ws.Cells(archiveRow, CLng(archiveColumns(NOMENCLATURE_HEADER_ORDER))).Value = CStr(rowData("Order"))
    ws.Cells(archiveRow, CLng(archiveColumns(NOMENCLATURE_HEADER_CODE))).Value = CStr(rowData("Code"))
    ws.Cells(archiveRow, CLng(archiveColumns(NOMENCLATURE_HEADER_NAME))).Value = CStr(rowData("Name"))
    ws.Cells(archiveRow, CLng(archiveColumns(NOMENCLATURE_HEADER_STATEMENT))).Value = CStr(rowData("Statement"))
    ws.Cells(archiveRow, CLng(archiveColumns(NOMENCLATURE_HEADER_QTY))).Value = CDbl(rowData("Qty"))
    ws.Cells(archiveRow, CLng(archiveColumns(NOMENCLATURE_HEADER_STICKER))).Value = CStr(rowData("Sticker"))
    ws.Cells(archiveRow, CLng(archiveColumns(NOMENCLATURE_HEADER_SUPPLIER))).Value = CStr(rowData("Supplier"))
    ws.Cells(archiveRow, CLng(archiveColumns(NOMENCLATURE_HEADER_MIL_COMMENT))).Value = CStr(rowData("MilComment"))
    ws.Cells(archiveRow, CLng(archiveColumns(NOMENCLATURE_HEADER_MIL))).Value = CStr(rowData("Mil"))
    ws.Cells(archiveRow, CLng(archiveColumns(ARCHIVE_HEADER_DATE))).Value = currentDate
    ws.Cells(archiveRow, CLng(archiveColumns(ARCHIVE_HEADER_CONFIRMATION))).Value = confirmationValue
End Sub

Private Sub FillCorrectionArchiveRow(ByVal ws As Worksheet, ByVal archiveColumns As Object, ByVal archiveRow As Long, ByVal rowData As Object, ByVal currentDate As Date)
    ws.Cells(archiveRow, CLng(archiveColumns(NOMENCLATURE_HEADER_ORDER))).Value = CStr(rowData("Order"))
    ws.Cells(archiveRow, CLng(archiveColumns(NOMENCLATURE_HEADER_CODE))).Value = CStr(rowData("Code"))
    ws.Cells(archiveRow, CLng(archiveColumns(NOMENCLATURE_HEADER_NAME))).Value = CStr(rowData("Name"))
    ws.Cells(archiveRow, CLng(archiveColumns(NOMENCLATURE_HEADER_STATEMENT))).Value = CStr(rowData("Statement"))
    ws.Cells(archiveRow, CLng(archiveColumns(NOMENCLATURE_HEADER_QTY))).Value = CDbl(rowData("Qty"))
    ws.Cells(archiveRow, CLng(archiveColumns(NOMENCLATURE_HEADER_STICKER))).Value = CStr(rowData("Sticker"))
    ws.Cells(archiveRow, CLng(archiveColumns(NOMENCLATURE_HEADER_SUPPLIER))).Value = CStr(rowData("Supplier"))
    ws.Cells(archiveRow, CLng(archiveColumns(NOMENCLATURE_HEADER_MIL_COMMENT))).Value = CStr(rowData("MilComment"))
    ws.Cells(archiveRow, CLng(archiveColumns(NOMENCLATURE_HEADER_MIL))).Value = CStr(rowData("Mil"))
    ws.Cells(archiveRow, CLng(archiveColumns(ARCHIVE_HEADER_DATE))).Value = currentDate
End Sub

Private Sub FillImportInfoTargetRow(ByVal ws As Worksheet, ByVal targetColumns As Object, ByVal targetRow As Long, ByVal rowData As Object, ByVal currentDate As Date)
    Dim lastColumn As Long

    lastColumn = LastUsedColumn(ws)
    If lastColumn < 19 Then lastColumn = 19
    ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow, lastColumn)).ClearContents
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_ARTICLE))).Value = CStr(rowData("Code"))
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_NAME))).Value = CStr(rowData("Name"))
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_ORDER))).Value = CStr(rowData("Order"))
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_STATEMENT))).Value = CStr(rowData("Statement"))
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_QTY))).Value = CDbl(rowData("Qty"))
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_STICKER))).Value = rowData("Sticker")
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_SHIPMENT_INVOICE))).Value = rowData("ShipmentInvoice")
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_SHIPMENT_DATE))).Value = rowData("ShipmentDate")
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_ALCOHOL))).Value = rowData("Alcohol")
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_VOLUME))).Value = rowData("Volume")
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_VINTAGE))).Value = rowData("Vintage")
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_DATE))).Value = currentDate
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_SUPPLIER))).Value = CStr(rowData("Supplier"))
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_MIL_COMMENT))).Value = CStr(rowData("MilComment"))
    ws.Cells(targetRow, CLng(targetColumns(TARGET_HEADER_MIL))).Value = CStr(rowData("Mil"))
End Sub

Private Sub FillImportInfoCorrectionTargetRow(ByVal ws As Worksheet, ByVal targetColumns As Object, ByVal targetRow As Long, ByVal rowData As Object)
    ws.Range(ws.Cells(targetRow, 1), ws.Cells(targetRow, 15)).ClearContents
    FillImportInfoStagingLikeRow ws, targetColumns, targetRow, rowData
    ws.Cells(targetRow, CLng(targetColumns(CORRECTION_STATUS_HEADER))).Value = CORRECTION_STATUS_PENDING
End Sub

Private Sub FillImportInfoArchiveRow(ByVal ws As Worksheet, ByVal archiveColumns As Object, ByVal archiveRow As Long, ByVal rowData As Object, ByVal currentDate As Date, ByVal repeatConfirmation As String, ByVal quantityMismatchConfirmation As String)
    FillImportInfoStagingLikeRow ws, archiveColumns, archiveRow, rowData
    ws.Cells(archiveRow, CLng(archiveColumns(ARCHIVE_HEADER_DATE))).Value = currentDate
    ws.Cells(archiveRow, CLng(archiveColumns(ARCHIVE_HEADER_CONFIRMATION))).Value = repeatConfirmation
    ws.Cells(archiveRow, CLng(archiveColumns(ARCHIVE_HEADER_QTY_MISMATCH_CONFIRMATION))).Value = quantityMismatchConfirmation
End Sub

Private Sub FillImportInfoCorrectionArchiveRow(ByVal ws As Worksheet, ByVal archiveColumns As Object, ByVal archiveRow As Long, ByVal rowData As Object, ByVal currentDate As Date, ByVal quantityMismatchConfirmation As String)
    FillImportInfoStagingLikeRow ws, archiveColumns, archiveRow, rowData
    ws.Cells(archiveRow, CLng(archiveColumns(ARCHIVE_HEADER_DATE))).Value = currentDate
    ws.Cells(archiveRow, CLng(archiveColumns(ARCHIVE_HEADER_QTY_MISMATCH_CONFIRMATION))).Value = quantityMismatchConfirmation
End Sub

Private Sub FillImportInfoStagingLikeRow(ByVal ws As Worksheet, ByVal columns As Object, ByVal rowIndex As Long, ByVal rowData As Object)
    ws.Cells(rowIndex, CLng(columns(IMPORT_INFO_HEADER_ORDER))).Value = CStr(rowData("Order"))
    ws.Cells(rowIndex, CLng(columns(IMPORT_INFO_HEADER_SUPPLIER))).Value = CStr(rowData("Supplier"))
    ws.Cells(rowIndex, CLng(columns(IMPORT_INFO_HEADER_CODE))).Value = CStr(rowData("Code"))
    ws.Cells(rowIndex, CLng(columns(IMPORT_INFO_HEADER_NAME))).Value = CStr(rowData("Name"))
    ws.Cells(rowIndex, CLng(columns(IMPORT_INFO_HEADER_STATEMENT))).Value = CStr(rowData("Statement"))
    ws.Cells(rowIndex, CLng(columns(IMPORT_INFO_HEADER_QTY))).Value = CDbl(rowData("Qty"))
    ws.Cells(rowIndex, CLng(columns(IMPORT_INFO_HEADER_STICKER))).Value = rowData("Sticker")
    ws.Cells(rowIndex, CLng(columns(IMPORT_INFO_HEADER_SHIPMENT_INVOICE))).Value = rowData("ShipmentInvoice")
    ws.Cells(rowIndex, CLng(columns(IMPORT_INFO_HEADER_SHIPMENT_DATE))).Value = rowData("ShipmentDate")
    ws.Cells(rowIndex, CLng(columns(IMPORT_INFO_HEADER_ALCOHOL))).Value = rowData("Alcohol")
    ws.Cells(rowIndex, CLng(columns(IMPORT_INFO_HEADER_VOLUME))).Value = rowData("Volume")
    ws.Cells(rowIndex, CLng(columns(IMPORT_INFO_HEADER_VINTAGE))).Value = rowData("Vintage")
    ws.Cells(rowIndex, CLng(columns(IMPORT_INFO_HEADER_MIL))).Value = CStr(rowData("Mil"))
    ws.Cells(rowIndex, CLng(columns(IMPORT_INFO_HEADER_MIL_COMMENT))).Value = CStr(rowData("MilComment"))
End Sub

Private Function GetNextAppendRow(ByVal ws As Worksheet, ByVal keyColumn As Long) As Long
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, keyColumn).End(xlUp).Row
    If lastRow < 2 Then
        GetNextAppendRow = 2
    Else
        GetNextAppendRow = lastRow + 1
    End If
End Function

Private Function RollbackArchiveRows(ByVal wbArchive As Workbook, ByVal archiveStartRow As Long, ByVal archiveRowCount As Long, Optional ByVal archiveSheetName As String = ARCHIVE_SHEET_NAME) As String
    Dim ws As Worksheet

    On Error GoTo RollbackErr
    Set ws = GetWorkbookSheet(wbArchive, archiveSheetName)
    ws.Rows(CStr(archiveStartRow) & ":" & CStr(archiveStartRow + archiveRowCount - 1)).Delete
    wbArchive.Save
    Exit Function

RollbackErr:
    RollbackArchiveRows = "Не удалось откатить уже сохраненные строки в архиве. Проверьте файл архива вручную."
End Function

Private Function OpenNomenclatureWorkbook(ByVal workbookPath As String, ByVal workbookPassword As String, ByVal readOnlyMode As Boolean) As Workbook
    On Error GoTo ErrHandler

    TraceAutomationStep "send:openNomenclature:readOnly=" & CStr(readOnlyMode)
    If readOnlyMode Then
        Set OpenNomenclatureWorkbook = Workbooks.Open(FileName:=workbookPath, UpdateLinks:=False, ReadOnly:=True, Notify:=False)
    ElseIf Len(workbookPassword) > 0 Then
        Set OpenNomenclatureWorkbook = Workbooks.Open(FileName:=workbookPath, UpdateLinks:=False, ReadOnly:=False, WriteResPassword:=workbookPassword, Notify:=False)
    Else
        Set OpenNomenclatureWorkbook = Workbooks.Open(FileName:=workbookPath, UpdateLinks:=False, ReadOnly:=False, Notify:=False)
    End If
    TraceAutomationStep "send:openNomenclature:done"
    Exit Function

ErrHandler:
    RaiseSendRequestError "Нет доступа к файлу номенклатуры. Проверьте путь, включение Forti/VPN и пароль номенклатуры на листе '" & SETTINGS_SHEET_NAME & "'." & vbCrLf & vbCrLf & _
                         "Техническая причина: " & Err.Description
End Function

Private Function OpenArchiveWorkbook(ByVal workbookPath As String, ByVal readOnlyMode As Boolean) As Workbook
    On Error GoTo ErrHandler

    TraceAutomationStep "send:openArchive:readOnly=" & CStr(readOnlyMode)
    Set OpenArchiveWorkbook = Workbooks.Open(FileName:=workbookPath, UpdateLinks:=False, ReadOnly:=readOnlyMode, Password:=ARCHIVE_PASSWORD, Notify:=False)
    TraceAutomationStep "send:openArchive:done"
    Exit Function

ErrHandler:
    RaiseSendRequestError "Нет доступа к архивному файлу. Проверьте путь, включение Forti/VPN и пароль архивного файла." & vbCrLf & vbCrLf & _
                         "Техническая причина: " & Err.Description
End Function

Private Function GetWorkbookSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorkbookSheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    If GetWorkbookSheet Is Nothing Then
        RaiseSendRequestError "Во внешнем файле отсутствует лист '" & sheetName & "'."
    End If
End Function

Private Function GetArchiveSheet(ByVal wb As Workbook) As Worksheet
    Set GetArchiveSheet = GetWorkbookSheet(wb, ARCHIVE_SHEET_NAME)
End Function

Private Function GetCorrectionArchiveSheet(ByVal wb As Workbook) As Worksheet
    Set GetCorrectionArchiveSheet = GetWorkbookSheet(wb, CORRECTION_ARCHIVE_SHEET_NAME)
End Function

Private Function GetImportInfoArchiveSheet(ByVal wb As Workbook) As Worksheet
    Set GetImportInfoArchiveSheet = GetWorkbookSheet(wb, IMPORT_INFO_ARCHIVE_SHEET_NAME)
End Function

Private Function GetImportInfoCorrectionArchiveSheet(ByVal wb As Workbook) As Worksheet
    Set GetImportInfoCorrectionArchiveSheet = GetWorkbookSheet(wb, IMPORT_INFO_CORRECTION_ARCHIVE_SHEET_NAME)
End Function

Private Function BuildLockFilePath(ByVal nomenclaturePath As String) As String
    BuildLockFilePath = nomenclaturePath & LOCK_FILE_SUFFIX
End Function

Private Function GetFolderPathFromFilePath(ByVal filePath As String) As String
    Dim lastSlash As Long

    lastSlash = InStrRev(filePath, "\")
    If lastSlash = 0 Then lastSlash = InStrRev(filePath, "/")
    If lastSlash > 0 Then
        GetFolderPathFromFilePath = Left$(filePath, lastSlash - 1)
    End If
End Function

Private Function BuildSendErrorMessage(ByVal errorNumber As Long, ByVal errorDescription As String) As String
    If errorNumber = SEND_ERROR_NUMBER Or errorNumber = SEND_CONFLICT_ERROR_NUMBER Or errorNumber = AUTOMATION_ERROR_NUMBER Then
        BuildSendErrorMessage = errorDescription
    Else
        BuildSendErrorMessage = "Не удалось внести строчки в номенклатуру." & vbCrLf & vbCrLf & _
                                "Техническая причина: " & errorDescription
    End If
End Function

Private Function BuildCorrectionSendErrorMessage(ByVal errorNumber As Long, ByVal errorDescription As String) As String
    If errorNumber = SEND_ERROR_NUMBER Or errorNumber = SEND_CONFLICT_ERROR_NUMBER Or errorNumber = AUTOMATION_ERROR_NUMBER Then
        BuildCorrectionSendErrorMessage = errorDescription
    Else
        BuildCorrectionSendErrorMessage = "Не удалось внести корректировку в номенклатуру." & vbCrLf & vbCrLf & _
                                          "Техническая причина: " & errorDescription
    End If
End Function

Private Sub RaiseSendRequestError(ByVal message As String)
    Err.Raise SEND_ERROR_NUMBER, "SendNomenclatureRequest", message
End Sub

Private Function ParsePreparedQuantity(ByVal cellValue As Variant, ByVal orderValue As String, ByVal rowIndex As Long) As Double
    If Not IsNumeric(cellValue) Then
        RaiseSendRequestError "В строке " & rowIndex & " листа '" & NOMENCLATURE_REQUEST_SHEET_NAME & "' значение в столбце '" & NOMENCLATURE_HEADER_QTY & "' не является числом для заказа '" & orderValue & "'."
    End If

    ParsePreparedQuantity = CDbl(cellValue)
End Function

Private Function ParsePreparedImportInfoQuantity(ByVal cellValue As Variant, ByVal orderValue As String, ByVal rowIndex As Long) As Double
    If Not IsNumeric(cellValue) Then
        RaiseSendRequestError "В строке " & rowIndex & " листа '" & IMPORT_INFO_REQUEST_SHEET_NAME & "' значение в столбце '" & IMPORT_INFO_HEADER_QTY & "' не является числом для заказа '" & orderValue & "'."
    End If

    ParsePreparedImportInfoQuantity = CDbl(cellValue)
End Function

Private Function ResolveTargetSheetName(ByVal orderValue As String) As String
    orderValue = NormalizeOrderValue(orderValue)

    If Left$(orderValue, 3) = "TK-" Then
        ResolveTargetSheetName = TARGET_TK_SHEET_NAME
    ElseIf Left$(orderValue, 4) = "GKF-" Then
        ResolveTargetSheetName = TARGET_LA_SHEET_NAME
    Else
        RaiseSendRequestError "Заказ '" & orderValue & "' не поддерживается для внесения в номенклатуру. Допустимы только заказы с префиксом TK или GKF."
    End If
End Function

Private Function ResolveImportInfoTargetSheetName(ByVal orderValue As String) As String
    orderValue = NormalizeOrderValue(orderValue)

    If Left$(orderValue, 3) = "TK-" Then
        ResolveImportInfoTargetSheetName = IMPORT_INFO_TARGET_TK_SHEET_NAME
    ElseIf Left$(orderValue, 4) = "GKF-" Then
        ResolveImportInfoTargetSheetName = IMPORT_INFO_TARGET_LA_SHEET_NAME
    Else
        RaiseSendRequestError "Заказ '" & orderValue & "' не поддерживается для внесения сведений о ввозе. Допустимы только заказы с префиксом TK или GKF."
    End If
End Function

Private Function ResolveImportInfoCorrectionTargetSheetName(ByVal orderValue As String) As String
    orderValue = NormalizeOrderValue(orderValue)

    If Left$(orderValue, 3) = "TK-" Then
        ResolveImportInfoCorrectionTargetSheetName = IMPORT_INFO_CORRECTION_TARGET_TK_SHEET_NAME
    ElseIf Left$(orderValue, 4) = "GKF-" Then
        ResolveImportInfoCorrectionTargetSheetName = IMPORT_INFO_CORRECTION_TARGET_LA_SHEET_NAME
    Else
        RaiseSendRequestError "Заказ '" & orderValue & "' не поддерживается для внесения корректировки сведений о ввозе. Допустимы только заказы с префиксом TK или GKF."
    End If
End Function

Private Function ResolveCorrectionTargetSheetName(ByVal orderValue As String) As String
    orderValue = NormalizeOrderValue(orderValue)

    If Left$(orderValue, 3) = "TK-" Then
        ResolveCorrectionTargetSheetName = CORRECTION_TARGET_TK_SHEET_NAME
    ElseIf Left$(orderValue, 4) = "GKF-" Then
        ResolveCorrectionTargetSheetName = CORRECTION_TARGET_LA_SHEET_NAME
    Else
        RaiseSendRequestError "Заказ '" & orderValue & "' не поддерживается для внесения корректировки. Допустимы только заказы с префиксом TK или GKF."
    End If
End Function

Private Sub NormalizeNomenclatureOrderInput(ByRef orderList As Collection, ByRef orderMap As Object)
    Dim ws As Worksheet
    Dim maxRow As Long
    Dim lastRow As Long
    Dim i As Long
    Dim value As String

    Set ws = GetNomenclatureRequestSheet

    ClearPreparedColumns ws

    If ws.FilterMode Then
        On Error Resume Next
        ws.ShowAllData
        On Error GoTo 0
    End If

    maxRow = FindInputLastRow(ws)
    If maxRow < 2 Then Exit Sub

    For i = maxRow To 2 Step -1
        If NormalizeStageText(ws.Cells(i, 1).Value) = "" Then
            ws.Rows(i).Delete
        End If
    Next i

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    For i = 2 To lastRow
        value = NormalizeOrderValue(ws.Cells(i, 1).Value)
        ws.Cells(i, 1).Value = value
    Next i

    ws.Range("A1:A" & lastRow).RemoveDuplicates Columns:=1, Header:=xlYes

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        value = NormalizeOrderValue(ws.Cells(i, 1).Value)
        If Len(value) > 0 Then
            ws.Cells(i, 1).Value = value
            If Not orderMap.Exists(value) Then
                orderMap.Add value, True
                orderList.Add value
            End If
        End If
    Next i
End Sub

Private Sub NormalizeImportInfoOrderInput(ByRef orderList As Collection, ByRef orderMap As Object)
    Dim ws As Worksheet
    Dim maxRow As Long
    Dim lastRow As Long
    Dim i As Long
    Dim value As String

    Set ws = GetImportInfoRequestSheet

    ClearPreparedColumns ws, 2, 14
    ClearImportInfoBaselineSheetContents

    If ws.FilterMode Then
        On Error Resume Next
        ws.ShowAllData
        On Error GoTo 0
    End If

    maxRow = FindInputLastRow(ws)
    If maxRow < 2 Then Exit Sub

    For i = maxRow To 2 Step -1
        If NormalizeStageText(ws.Cells(i, 1).Value) = "" Then
            ws.Rows(i).Delete
        End If
    Next i

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    For i = 2 To lastRow
        value = NormalizeOrderValue(ws.Cells(i, 1).Value)
        ws.Cells(i, 1).Value = value
    Next i

    ws.Range("A1:A" & lastRow).RemoveDuplicates Columns:=1, Header:=xlYes

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        value = NormalizeOrderValue(ws.Cells(i, 1).Value)
        If Len(value) > 0 Then
            ws.Cells(i, 1).Value = value
            If Not orderMap.Exists(value) Then
                orderMap.Add value, True
                orderList.Add value
            End If
        End If
    Next i
End Sub

Private Sub LoadOrdersIntoFsmSheet(ByVal orderList As Collection)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long

    Set ws = GetFsmRequestSheet
    lastRow = FindSheetLastRow(ws)
    lastCol = LastUsedColumn(ws)
    If lastCol < 13 Then lastCol = 13

    If lastRow >= 2 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).ClearContents
    End If

    For i = 1 To orderList.Count
        ws.Cells(i + 1, 1).Value = orderList(i)
    Next i
End Sub

Private Sub PopulateNomenclatureRequestSheet(ByVal orderMap As Object)
    Dim wsFsm As Worksheet
    Dim wsNom As Worksheet
    Dim groupedRows As Object
    Dim outputKeys As Collection
    Dim foundOrders As Object
    Dim stickerMap As Object
    Dim milValue As String
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim key As String
    Dim orderValue As String
    Dim statementValue As String
    Dim codeValue As String
    Dim nameValue As String
    Dim supplierValue As String
    Dim qtyValue As Double
    Dim groupEntry As Object
    Dim dataKey As Variant
    Dim outputData() As Variant
    Dim i As Long

    Set wsFsm = GetFsmRequestSheet
    Set wsNom = GetNomenclatureRequestSheet
    Set groupedRows = CreateObject("Scripting.Dictionary")
    groupedRows.CompareMode = vbTextCompare
    Set outputKeys = New Collection
    Set foundOrders = CreateObject("Scripting.Dictionary")
    foundOrders.CompareMode = vbTextCompare
    Set stickerMap = BuildStickerMap(orderMap)

    milValue = Trim$(GetSettingsValue("МИЛ", GetSettingsSheet.Range("B2").Value))
    lastRow = wsFsm.Cells(wsFsm.Rows.Count, FindColumn(wsFsm, "Заказ")).End(xlUp).Row

    For rowIndex = 2 To lastRow
        orderValue = NormalizeOrderValue(wsFsm.Cells(rowIndex, FindColumn(wsFsm, "Заказ")).Value)
        If Len(orderValue) = 0 Then GoTo NextRow
        If Not orderMap.Exists(orderValue) Then GoTo NextRow

        statementValue = NormalizeStageText(wsFsm.Cells(rowIndex, FindColumn(wsFsm, "Заявление (КМ)")).Value)
        codeValue = NormalizeStageText(wsFsm.Cells(rowIndex, FindColumn(wsFsm, "Код (КМ)")).Value)
        nameValue = Trim$(CStr(wsFsm.Cells(rowIndex, FindColumn(wsFsm, "Позиция (КМ)")).Value))
        supplierValue = Trim$(CStr(wsFsm.Cells(rowIndex, FindColumn(wsFsm, "Поставщик (КМ)")).Value))

        If Len(statementValue) = 0 Then
            ShowNomenclatureRequestError "Для заказа '" & orderValue & "' не заполнено значение в столбце 'Заявление (КМ)' на листе '" & FSM_REQUEST_SHEET_NAME & "'.", True
        End If
        If Len(codeValue) = 0 Then
            ShowNomenclatureRequestError "Для заказа '" & orderValue & "' не заполнено значение в столбце 'Код (КМ)' на листе '" & FSM_REQUEST_SHEET_NAME & "'.", True
        End If
        If Len(NormalizeStageText(nameValue)) = 0 Then
            ShowNomenclatureRequestError "Для заказа '" & orderValue & "' не заполнено значение в столбце 'Позиция (КМ)' на листе '" & FSM_REQUEST_SHEET_NAME & "'.", True
        End If
        If Len(NormalizeStageText(supplierValue)) = 0 Then
            ShowNomenclatureRequestError "Для заказа '" & orderValue & "' не заполнено значение в столбце 'Поставщик (КМ)' на листе '" & FSM_REQUEST_SHEET_NAME & "'.", True
        End If

        qtyValue = ParseQuantityValue(wsFsm.Cells(rowIndex, FindColumn(wsFsm, "Кол-во (КМ)")).Value, orderValue, rowIndex)
        foundOrders(orderValue) = True

        key = BuildGroupKey(orderValue, statementValue, codeValue)
        If Not groupedRows.Exists(key) Then
            Set groupEntry = CreateObject("Scripting.Dictionary")
            groupEntry.CompareMode = vbTextCompare
            groupEntry("Order") = orderValue
            groupEntry("Code") = Trim$(CStr(wsFsm.Cells(rowIndex, FindColumn(wsFsm, "Код (КМ)")).Value))
            groupEntry("Name") = nameValue
            groupEntry("Statement") = Trim$(CStr(wsFsm.Cells(rowIndex, FindColumn(wsFsm, "Заявление (КМ)")).Value))
            groupEntry("Supplier") = supplierValue
            groupEntry("Qty") = qtyValue
            If stickerMap.Exists(BuildOrderCodeKey(orderValue, codeValue)) Then
                groupEntry("Sticker") = stickerMap(BuildOrderCodeKey(orderValue, codeValue))
            Else
                groupEntry("Sticker") = ""
            End If
            groupedRows.Add key, groupEntry
            outputKeys.Add key
        Else
            Set groupEntry = groupedRows(key)
            EnsureGroupConsistency groupEntry, nameValue, supplierValue, orderValue
            groupEntry("Qty") = CDbl(groupEntry("Qty")) + qtyValue
        End If
NextRow:
    Next rowIndex

    For Each dataKey In orderMap.Keys
        If Not foundOrders.Exists(CStr(dataKey)) Then
            ShowNomenclatureRequestError "После обновления на листе '" & FSM_REQUEST_SHEET_NAME & "' не найдено ни одной строки для заказа '" & CStr(dataKey) & "'.", True
        End If
    Next dataKey

    If outputKeys.Count = 0 Then
        ShowNomenclatureRequestError "Не удалось подготовить ни одной строки для листа '" & NOMENCLATURE_REQUEST_SHEET_NAME & "'.", False
    End If

    ClearNomenclatureOutput wsNom

    ReDim outputData(1 To outputKeys.Count, 1 To 9)
    For i = 1 To outputKeys.Count
        Set groupEntry = groupedRows(outputKeys(i))
        outputData(i, 1) = groupEntry("Order")
        outputData(i, 2) = groupEntry("Supplier")
        outputData(i, 3) = groupEntry("Code")
        outputData(i, 4) = groupEntry("Name")
        outputData(i, 5) = groupEntry("Statement")
        outputData(i, 6) = groupEntry("Qty")
        outputData(i, 7) = groupEntry("Sticker")
        outputData(i, 8) = milValue
        outputData(i, 9) = ""
    Next i

    wsNom.Range("A2").Resize(outputKeys.Count, 9).Value = outputData
    ApplyNomenclatureWarningPlaceholders wsNom, outputKeys.Count + 1
    ApplyNomenclatureFilter wsNom, outputKeys.Count + 1
End Sub

Private Sub PopulateImportInfoRequestSheet(ByVal orderMap As Object, ByVal shipmentLookup As Object)
    Dim wsFsm As Worksheet
    Dim wsImport As Worksheet
    Dim wsBaseline As Worksheet
    Dim groupedRows As Object
    Dim outputKeys As Collection
    Dim foundOrders As Object
    Dim milValue As String
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim key As String
    Dim orderValue As String
    Dim statementValue As String
    Dim codeValue As String
    Dim nameValue As String
    Dim supplierValue As String
    Dim qtyValue As Double
    Dim groupEntry As Object
    Dim dataKey As Variant
    Dim outputData() As Variant
    Dim baselineData() As Variant
    Dim i As Long

    Set wsFsm = GetFsmRequestSheet
    Set wsImport = GetImportInfoRequestSheet
    Set wsBaseline = GetImportInfoBaselineSheet
    Set groupedRows = CreateObject("Scripting.Dictionary")
    groupedRows.CompareMode = vbTextCompare
    Set outputKeys = New Collection
    Set foundOrders = CreateObject("Scripting.Dictionary")
    foundOrders.CompareMode = vbTextCompare

    milValue = Trim$(GetSettingsValue("МИЛ", GetSettingsSheet.Range("B2").Value))
    lastRow = wsFsm.Cells(wsFsm.Rows.Count, FindColumn(wsFsm, "Заказ")).End(xlUp).Row

    For rowIndex = 2 To lastRow
        orderValue = NormalizeOrderValue(wsFsm.Cells(rowIndex, FindColumn(wsFsm, "Заказ")).Value)
        If Len(orderValue) = 0 Then GoTo NextRow
        If Not orderMap.Exists(orderValue) Then GoTo NextRow

        statementValue = NormalizeStageText(wsFsm.Cells(rowIndex, FindColumn(wsFsm, "Заявление (КМ)")).Value)
        codeValue = NormalizeStageText(wsFsm.Cells(rowIndex, FindColumn(wsFsm, "Код (КМ)")).Value)
        nameValue = Trim$(CStr(wsFsm.Cells(rowIndex, FindColumn(wsFsm, "Позиция (КМ)")).Value))
        supplierValue = Trim$(CStr(wsFsm.Cells(rowIndex, FindColumn(wsFsm, "Поставщик (КМ)")).Value))

        If Len(statementValue) = 0 Then
            ShowImportInfoRequestError "Для заказа '" & orderValue & "' не заполнено значение в столбце 'Заявление (КМ)' на листе '" & FSM_REQUEST_SHEET_NAME & "'.", True
        End If
        If Len(codeValue) = 0 Then
            ShowImportInfoRequestError "Для заказа '" & orderValue & "' не заполнено значение в столбце 'Код (КМ)' на листе '" & FSM_REQUEST_SHEET_NAME & "'.", True
        End If
        If Len(NormalizeStageText(nameValue)) = 0 Then
            ShowImportInfoRequestError "Для заказа '" & orderValue & "' не заполнено значение в столбце 'Позиция (КМ)' на листе '" & FSM_REQUEST_SHEET_NAME & "'.", True
        End If
        If Len(NormalizeStageText(supplierValue)) = 0 Then
            ShowImportInfoRequestError "Для заказа '" & orderValue & "' не заполнено значение в столбце 'Поставщик (КМ)' на листе '" & FSM_REQUEST_SHEET_NAME & "'.", True
        End If

        qtyValue = ParseQuantityValue(wsFsm.Cells(rowIndex, FindColumn(wsFsm, "Кол-во (КМ)")).Value, orderValue, rowIndex)
        foundOrders(orderValue) = True

        key = BuildGroupKey(orderValue, statementValue, codeValue)
        If Not groupedRows.Exists(key) Then
            Set groupEntry = CreateObject("Scripting.Dictionary")
            groupEntry.CompareMode = vbTextCompare
            groupEntry("Order") = orderValue
            groupEntry("Supplier") = supplierValue
            groupEntry("Code") = Trim$(CStr(wsFsm.Cells(rowIndex, FindColumn(wsFsm, "Код (КМ)")).Value))
            groupEntry("Name") = nameValue
            groupEntry("Statement") = Trim$(CStr(wsFsm.Cells(rowIndex, FindColumn(wsFsm, "Заявление (КМ)")).Value))
            groupEntry("Qty") = qtyValue
            groupEntry("Mil") = milValue
            groupedRows.Add key, groupEntry
            outputKeys.Add key
        Else
            Set groupEntry = groupedRows(key)
            EnsureGroupConsistency groupEntry, nameValue, supplierValue, orderValue
            groupEntry("Qty") = CDbl(groupEntry("Qty")) + qtyValue
        End If
NextRow:
    Next rowIndex

    For Each dataKey In orderMap.Keys
        If Not foundOrders.Exists(CStr(dataKey)) Then
            ShowImportInfoRequestError "После обновления на листе '" & FSM_REQUEST_SHEET_NAME & "' не найдено ни одной строки для заказа '" & CStr(dataKey) & "'.", True
        End If
    Next dataKey

    If outputKeys.Count = 0 Then
        ShowImportInfoRequestError "Не удалось подготовить ни одной строки для листа '" & IMPORT_INFO_REQUEST_SHEET_NAME & "'.", False
    End If

    ClearImportInfoOutput wsImport
    ClearImportInfoBaselineSheetContents

    ReDim outputData(1 To outputKeys.Count, 1 To 14)
    ReDim baselineData(1 To outputKeys.Count, 1 To 10)
    For i = 1 To outputKeys.Count
        Set groupEntry = groupedRows(outputKeys(i))
        ApplyImportInfoShipmentLookup groupEntry, shipmentLookup

        outputData(i, 1) = groupEntry("Order")
        outputData(i, 2) = groupEntry("Supplier")
        outputData(i, 3) = groupEntry("Code")
        outputData(i, 4) = groupEntry("Name")
        outputData(i, 5) = groupEntry("Statement")
        outputData(i, 6) = groupEntry("Qty")
        outputData(i, 7) = groupEntry("ShipmentSticker")
        outputData(i, 8) = groupEntry("ShipmentInvoice")
        outputData(i, 9) = groupEntry("ShipmentDate")
        outputData(i, 10) = MANUAL_FILL_TEXT
        outputData(i, 11) = MANUAL_FILL_TEXT
        outputData(i, 12) = MANUAL_FILL_TEXT
        outputData(i, 13) = groupEntry("Mil")
        outputData(i, 14) = ""

        baselineData(i, 1) = groupEntry("Order")
        baselineData(i, 2) = groupEntry("Supplier")
        baselineData(i, 3) = groupEntry("Code")
        baselineData(i, 4) = groupEntry("Name")
        baselineData(i, 5) = groupEntry("Statement")
        baselineData(i, 6) = groupEntry("Qty")
        baselineData(i, 7) = groupEntry("ShipmentSticker")
        baselineData(i, 8) = groupEntry("ShipmentInvoice")
        baselineData(i, 9) = groupEntry("ShipmentDate")
        baselineData(i, 10) = groupEntry("Mil")
    Next i

    wsImport.Range("A2").Resize(outputKeys.Count, 14).Value = outputData
    wsBaseline.Range("A2").Resize(outputKeys.Count, 10).Value = baselineData
    wsBaseline.Visible = xlSheetVeryHidden
    ApplyImportInfoWarningPlaceholders wsImport, outputKeys.Count + 1
    ApplyImportInfoFilter wsImport, outputKeys.Count + 1
End Sub

Private Function BuildImportInfoShipmentLookup(ByVal orderMap As Object) As Object
    Dim result As Object
    Dim nomenclaturePath As String
    Dim nomenclaturePassword As String
    Dim wb As Workbook

    On Error GoTo ErrHandler

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    nomenclaturePath = Trim$(GetSettingsValue(SETTING_NOMENCLATURE_PATH, ""))
    nomenclaturePassword = Trim$(GetSettingsValue(SETTING_NOMENCLATURE_PASSWORD, ""))

    If Len(nomenclaturePath) = 0 Then
        ShowImportInfoRequestError "На листе '" & SETTINGS_SHEET_NAME & "' не указан путь к номенклатуре.", False
    End If
    If InStr(1, nomenclaturePath, "://", vbTextCompare) > 0 Then
        ShowImportInfoRequestError "Для чтения данных отправки нужен файловый путь к номенклатуре, а не веб-адрес.", False
    End If
    If Len(Dir$(nomenclaturePath)) = 0 Then
        ShowImportInfoRequestError "Файл номенклатуры не найден: " & nomenclaturePath, False
    End If

    Set wb = OpenNomenclatureWorkbook(nomenclaturePath, nomenclaturePassword, True)
    LoadImportInfoShipmentLookupFromSheet GetWorkbookSheet(wb, TARGET_TK_SHEET_NAME), orderMap, result
    LoadImportInfoShipmentLookupFromSheet GetWorkbookSheet(wb, TARGET_LA_SHEET_NAME), orderMap, result

    Set BuildImportInfoShipmentLookup = result

CleanExit:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function

ErrHandler:
    Dim errorText As String

    errorText = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    ShowImportInfoRequestError "Не удалось прочитать данные отправки из файла номенклатуры." & vbCrLf & vbCrLf & _
                               "Техническая причина: " & errorText, False
End Function

Private Sub LoadImportInfoShipmentLookupFromSheet(ByVal ws As Worksheet, ByVal orderMap As Object, ByRef result As Object)
    Dim orderColumn As Long
    Dim statementColumn As Long
    Dim articleColumn As Long
    Dim qtyColumn As Long
    Dim stickerColumn As Long
    Dim invoiceColumn As Long
    Dim shipmentDateColumn As Long
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim orderValue As String
    Dim lookupKey As String
    Dim entry As Object

    PrepareReadOnlyLookupSheetView ws

    orderColumn = FindColumn(ws, TARGET_HEADER_ORDER)
    statementColumn = FindColumn(ws, TARGET_HEADER_STATEMENT)
    articleColumn = FindColumn(ws, TARGET_HEADER_ARTICLE)
    qtyColumn = FindColumn(ws, TARGET_HEADER_QTY)
    stickerColumn = FindColumn(ws, TARGET_HEADER_STICKER)
    invoiceColumn = FindColumn(ws, TARGET_HEADER_SHIPMENT_INVOICE)
    shipmentDateColumn = FindColumn(ws, TARGET_HEADER_SHIPMENT_DATE)

    lastRow = ws.Cells(ws.Rows.Count, orderColumn).End(xlUp).Row
    For rowIndex = 2 To lastRow
        orderValue = NormalizeOrderValue(ws.Cells(rowIndex, orderColumn).Value)
        If Len(orderValue) = 0 Then GoTo NextShipmentRow
        If Not orderMap.Exists(orderValue) Then GoTo NextShipmentRow

        lookupKey = BuildImportInfoShipmentKey( _
            ws.Name, _
            orderValue, _
            ws.Cells(rowIndex, statementColumn).Value, _
            ws.Cells(rowIndex, articleColumn).Value, _
            ws.Cells(rowIndex, qtyColumn).Value)

        If result.Exists(lookupKey) Then
            Set entry = result(lookupKey)
            entry("MatchCount") = CLng(entry("MatchCount")) + 1
        Else
            Set entry = CreateObject("Scripting.Dictionary")
            entry.CompareMode = vbTextCompare
            entry("MatchCount") = 1
            entry("Sticker") = ws.Cells(rowIndex, stickerColumn).Value
            entry("Invoice") = ws.Cells(rowIndex, invoiceColumn).Value
            entry("ShipmentDate") = ws.Cells(rowIndex, shipmentDateColumn).Value
            result.Add lookupKey, entry
        End If
NextShipmentRow:
    Next rowIndex
End Sub

Private Sub PrepareReadOnlyLookupSheetView(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Rows.Hidden = False
    ws.Columns.Hidden = False
    If ws.FilterMode Then ws.ShowAllData
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo 0
End Sub

Private Sub ApplyImportInfoShipmentLookup(ByVal groupEntry As Object, ByVal shipmentLookup As Object)
    Dim targetSheetName As String
    Dim lookupKey As String
    Dim entry As Object

    targetSheetName = ResolveImportInfoShipmentTargetSheetName(CStr(groupEntry("Order")))
    If Len(targetSheetName) = 0 Then
        AssignImportInfoShipmentWarning groupEntry, MISSING_DATA_TEXT
        Exit Sub
    End If

    lookupKey = BuildImportInfoShipmentKey( _
        targetSheetName, _
        groupEntry("Order"), _
        groupEntry("Statement"), _
        groupEntry("Code"), _
        groupEntry("Qty"))

    If Not shipmentLookup.Exists(lookupKey) Then
        AssignImportInfoShipmentWarning groupEntry, MISSING_DATA_TEXT
        Exit Sub
    End If

    Set entry = shipmentLookup(lookupKey)
    If CLng(entry("MatchCount")) > 1 Then
        AssignImportInfoShipmentWarning groupEntry, MULTIPLE_MATCHES_TEXT
        Exit Sub
    End If

    groupEntry("ShipmentSticker") = ValueOrMissingWarning(entry("Sticker"))
    groupEntry("ShipmentInvoice") = ValueOrMissingWarning(entry("Invoice"))
    groupEntry("ShipmentDate") = ValueOrMissingWarning(entry("ShipmentDate"))
End Sub

Private Sub AssignImportInfoShipmentWarning(ByVal groupEntry As Object, ByVal warningText As String)
    groupEntry("ShipmentSticker") = warningText
    groupEntry("ShipmentInvoice") = warningText
    groupEntry("ShipmentDate") = warningText
End Sub

Private Function ResolveImportInfoShipmentTargetSheetName(ByVal orderValue As String) As String
    orderValue = NormalizeOrderValue(orderValue)

    If Left$(orderValue, 3) = "TK-" Then
        ResolveImportInfoShipmentTargetSheetName = TARGET_TK_SHEET_NAME
    ElseIf Left$(orderValue, 4) = "GKF-" Then
        ResolveImportInfoShipmentTargetSheetName = TARGET_LA_SHEET_NAME
    Else
        ResolveImportInfoShipmentTargetSheetName = ""
    End If
End Function

Private Function BuildImportInfoShipmentKey(ByVal targetSheetName As String, ByVal orderValue As Variant, ByVal statementValue As Variant, ByVal codeValue As Variant, ByVal qtyValue As Variant) As String
    BuildImportInfoShipmentKey = targetSheetName & "|" & _
                                 NormalizeOrderValue(orderValue) & "|" & _
                                 NormalizeStageText(statementValue) & "|" & _
                                 NormalizeStageText(codeValue) & "|" & _
                                 NormalizeQuantitySignature(qtyValue)
End Function

Private Function ValueOrMissingWarning(ByVal value As Variant) As Variant
    If Len(NormalizeStageText(value)) = 0 Then
        ValueOrMissingWarning = MISSING_DATA_TEXT
    Else
        ValueOrMissingWarning = value
    End If
End Function

Private Function BuildStickerMap(ByVal orderMap As Object) As Object
    Dim ws As Worksheet
    Dim result As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim orderValue As String
    Dim codeValue As String
    Dim stickerValue As String
    Dim dataKey As String

    Set ws = GetAlcoReportSheet
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    lastRow = ws.Cells(ws.Rows.Count, FindColumn(ws, "Номер заказа")).End(xlUp).Row
    For rowIndex = 2 To lastRow
        orderValue = NormalizeOrderValue(ws.Cells(rowIndex, FindColumn(ws, "Номер заказа")).Value)
        If Len(orderValue) = 0 Then GoTo NextAlcoRow
        If Not orderMap.Exists(orderValue) Then GoTo NextAlcoRow

        codeValue = NormalizeStageText(ws.Cells(rowIndex, FindColumn(ws, "Код номенклатуры")).Value)
        If Len(codeValue) = 0 Then GoTo NextAlcoRow

        stickerValue = ResolveStickerValue( _
            ws.Cells(rowIndex, FindColumn(ws, "Место оклейки ФСМ/ЧЗ")).Value, _
            ws.Cells(rowIndex, FindColumn(ws, "Поставщик")).Value, _
            ws.Cells(rowIndex, FindColumn(ws, "Транзитный склад оклейки")).Value)

        If Len(NormalizeStageText(stickerValue)) > 0 Then
            dataKey = BuildOrderCodeKey(orderValue, codeValue)
            If result.Exists(dataKey) Then
                If NormalizeStageText(result(dataKey)) <> NormalizeStageText(stickerValue) Then
                    ShowNomenclatureRequestError "Для заказа '" & orderValue & "' и кода '" & Trim$(CStr(ws.Cells(rowIndex, FindColumn(ws, "Код номенклатуры")).Value)) & "' найдено несколько разных значений 'Оклейщик' в обработанном Алкоотчете.", True
                End If
            Else
                result.Add dataKey, stickerValue
            End If
        End If
NextAlcoRow:
    Next rowIndex

    Set BuildStickerMap = result
End Function

Private Sub EnsureGroupConsistency(ByVal groupEntry As Object, ByVal nameValue As String, ByVal supplierValue As String, ByVal orderValue As String)
    If NormalizeStageText(groupEntry("Name")) <> NormalizeStageText(nameValue) Then
        ShowNomenclatureRequestError "Внутри одной агрегированной группы для заказа '" & orderValue & "' обнаружены разные значения 'Позиция (КМ)'.", True
    End If

    If NormalizeStageText(groupEntry("Supplier")) <> NormalizeStageText(supplierValue) Then
        ShowNomenclatureRequestError "Внутри одной агрегированной группы для заказа '" & orderValue & "' обнаружены разные значения 'Поставщик (КМ)'.", True
    End If
End Sub

Private Function ResolveStickerValue(ByVal placeValue As Variant, ByVal supplierValue As Variant, ByVal transitValue As Variant) As String
    Dim normalizedPlace As String
    Dim supplierText As String
    Dim transitText As String

    normalizedPlace = UCase$(NormalizeStageText(placeValue))
    supplierText = Trim$(CStr(supplierValue))
    transitText = Trim$(CStr(transitValue))

    If normalizedPlace = "ЗАВОД" Then
        ResolveStickerValue = supplierText
    ElseIf normalizedPlace = "ТС" Then
        If Len(NormalizeStageText(transitText)) > 0 And IsSingleStickerValue(transitText) Then
            ResolveStickerValue = transitText
        End If
    End If
End Function

Private Function IsSingleStickerValue(ByVal value As String) As Boolean
    IsSingleStickerValue = InStr(value, ",") = 0 And InStr(value, ";") = 0 And InStr(value, "/") = 0
End Function

Private Function ParseQuantityValue(ByVal cellValue As Variant, ByVal orderValue As String, ByVal rowIndex As Long) As Double
    If Not IsNumeric(cellValue) Then
        ShowNomenclatureRequestError "Для заказа '" & orderValue & "' в строке " & rowIndex & " листа '" & FSM_REQUEST_SHEET_NAME & "' значение в столбце 'Кол-во (КМ)' не является числом.", True
    End If

    ParseQuantityValue = CDbl(cellValue)
End Function

Private Function BuildGroupKey(ByVal orderValue As String, ByVal statementValue As String, ByVal codeValue As String) As String
    BuildGroupKey = NormalizeOrderValue(orderValue) & "|" & NormalizeStageText(statementValue) & "|" & NormalizeStageText(codeValue)
End Function

Private Function BuildOrderCodeKey(ByVal orderValue As String, ByVal codeValue As String) As String
    BuildOrderCodeKey = NormalizeOrderValue(orderValue) & "|" & NormalizeStageText(codeValue)
End Function

Private Function BuildCorrectionRowSignature(ByVal rowData As Object) As String
    BuildCorrectionRowSignature = CStr(rowData("TargetSheet")) & "|" & _
                                  NormalizeOrderValue(rowData("Order")) & "|" & _
                                  NormalizeStageText(rowData("Code")) & "|" & _
                                  NormalizeStageText(rowData("Name")) & "|" & _
                                  NormalizeStageText(rowData("Statement")) & "|" & _
                                  NormalizeQuantitySignature(rowData("Qty")) & "|" & _
                                  NormalizeStageText(rowData("Sticker")) & "|" & _
                                  NormalizeStageText(rowData("Supplier")) & "|" & _
                                  NormalizeStageText(rowData("MilComment")) & "|" & _
                                  NormalizeStageText(rowData("Mil"))
End Function

Private Function BuildCorrectionSheetRowSignature(ByVal ws As Worksheet, ByVal targetColumns As Object, ByVal rowIndex As Long) As String
    Dim orderValue As String

    orderValue = NormalizeOrderValue(ws.Cells(rowIndex, CLng(targetColumns(NOMENCLATURE_HEADER_ORDER))).Value)
    If Len(orderValue) = 0 Then Exit Function

    BuildCorrectionSheetRowSignature = ws.Name & "|" & _
                                       orderValue & "|" & _
                                       NormalizeStageText(ws.Cells(rowIndex, CLng(targetColumns(NOMENCLATURE_HEADER_CODE))).Value) & "|" & _
                                       NormalizeStageText(ws.Cells(rowIndex, CLng(targetColumns(NOMENCLATURE_HEADER_NAME))).Value) & "|" & _
                                       NormalizeStageText(ws.Cells(rowIndex, CLng(targetColumns(NOMENCLATURE_HEADER_STATEMENT))).Value) & "|" & _
                                       NormalizeQuantitySignature(ws.Cells(rowIndex, CLng(targetColumns(NOMENCLATURE_HEADER_QTY))).Value) & "|" & _
                                       NormalizeStageText(ws.Cells(rowIndex, CLng(targetColumns(NOMENCLATURE_HEADER_STICKER))).Value) & "|" & _
                                       NormalizeStageText(ws.Cells(rowIndex, CLng(targetColumns(NOMENCLATURE_HEADER_SUPPLIER))).Value) & "|" & _
                                       NormalizeStageText(ws.Cells(rowIndex, CLng(targetColumns(NOMENCLATURE_HEADER_MIL_COMMENT))).Value) & "|" & _
                                       NormalizeStageText(ws.Cells(rowIndex, CLng(targetColumns(NOMENCLATURE_HEADER_MIL))).Value)
End Function

Private Function BuildImportInfoCorrectionRowSignature(ByVal rowData As Object) As String
    BuildImportInfoCorrectionRowSignature = CStr(rowData("TargetSheet")) & "|" & _
                                            NormalizeOrderValue(rowData("Order")) & "|" & _
                                            NormalizeStageText(rowData("Supplier")) & "|" & _
                                            NormalizeStageText(rowData("Code")) & "|" & _
                                            NormalizeStageText(rowData("Name")) & "|" & _
                                            NormalizeStageText(rowData("Statement")) & "|" & _
                                            NormalizeQuantitySignature(rowData("Qty")) & "|" & _
                                            NormalizeStageText(rowData("Sticker")) & "|" & _
                                            NormalizeStageText(rowData("ShipmentInvoice")) & "|" & _
                                            NormalizeStageText(rowData("ShipmentDate")) & "|" & _
                                            NormalizeStageText(rowData("Alcohol")) & "|" & _
                                            NormalizeStageText(rowData("Volume")) & "|" & _
                                            NormalizeStageText(rowData("Vintage")) & "|" & _
                                            NormalizeStageText(rowData("Mil")) & "|" & _
                                            NormalizeStageText(rowData("MilComment"))
End Function

Private Function BuildImportInfoCorrectionSheetRowSignature(ByVal ws As Worksheet, ByVal targetColumns As Object, ByVal rowIndex As Long) As String
    Dim orderValue As String

    orderValue = NormalizeOrderValue(ws.Cells(rowIndex, CLng(targetColumns(IMPORT_INFO_HEADER_ORDER))).Value)
    If Len(orderValue) = 0 Then Exit Function

    BuildImportInfoCorrectionSheetRowSignature = ws.Name & "|" & _
                                                 orderValue & "|" & _
                                                 NormalizeStageText(ws.Cells(rowIndex, CLng(targetColumns(IMPORT_INFO_HEADER_SUPPLIER))).Value) & "|" & _
                                                 NormalizeStageText(ws.Cells(rowIndex, CLng(targetColumns(IMPORT_INFO_HEADER_CODE))).Value) & "|" & _
                                                 NormalizeStageText(ws.Cells(rowIndex, CLng(targetColumns(IMPORT_INFO_HEADER_NAME))).Value) & "|" & _
                                                 NormalizeStageText(ws.Cells(rowIndex, CLng(targetColumns(IMPORT_INFO_HEADER_STATEMENT))).Value) & "|" & _
                                                 NormalizeQuantitySignature(ws.Cells(rowIndex, CLng(targetColumns(IMPORT_INFO_HEADER_QTY))).Value) & "|" & _
                                                 NormalizeStageText(ws.Cells(rowIndex, CLng(targetColumns(IMPORT_INFO_HEADER_STICKER))).Value) & "|" & _
                                                 NormalizeStageText(ws.Cells(rowIndex, CLng(targetColumns(IMPORT_INFO_HEADER_SHIPMENT_INVOICE))).Value) & "|" & _
                                                 NormalizeStageText(ws.Cells(rowIndex, CLng(targetColumns(IMPORT_INFO_HEADER_SHIPMENT_DATE))).Value) & "|" & _
                                                 NormalizeStageText(ws.Cells(rowIndex, CLng(targetColumns(IMPORT_INFO_HEADER_ALCOHOL))).Value) & "|" & _
                                                 NormalizeStageText(ws.Cells(rowIndex, CLng(targetColumns(IMPORT_INFO_HEADER_VOLUME))).Value) & "|" & _
                                                 NormalizeStageText(ws.Cells(rowIndex, CLng(targetColumns(IMPORT_INFO_HEADER_VINTAGE))).Value) & "|" & _
                                                 NormalizeStageText(ws.Cells(rowIndex, CLng(targetColumns(IMPORT_INFO_HEADER_MIL))).Value) & "|" & _
                                                 NormalizeStageText(ws.Cells(rowIndex, CLng(targetColumns(IMPORT_INFO_HEADER_MIL_COMMENT))).Value)
End Function

Private Function NormalizeOrderValue(ByVal value As Variant) As String
    NormalizeOrderValue = NormalizeCyrLat(NormalizeStageText(value))
End Function

Private Function NormalizeStageText(ByVal value As Variant) As String
    Dim result As String

    If IsError(value) Or IsEmpty(value) Then
        NormalizeStageText = ""
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

    NormalizeStageText = Trim$(result)
End Function

Private Function NormalizeQuantitySignature(ByVal value As Variant) As String
    If IsNumeric(value) Then
        NormalizeQuantitySignature = Format$(CDbl(value), "0.###############")
    Else
        NormalizeQuantitySignature = NormalizeStageText(value)
    End If
End Function

Private Function FindInputLastRow(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    Dim offset As Long
    Dim foundData As Boolean

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        FindInputLastRow = lastRow
        Exit Function
    End If

    Do
        foundData = False
        For offset = 1 To 30
            If NormalizeStageText(ws.Cells(lastRow + offset, 1).Value) <> "" Then
                lastRow = lastRow + offset
                foundData = True
                Exit For
            End If
        Next offset
    Loop While foundData

    FindInputLastRow = lastRow
End Function

Private Function FindSheetLastRow(ByVal ws As Worksheet) As Long
    Dim lastCell As Range

    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, LookAt:=xlPart, _
                                 SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)

    If lastCell Is Nothing Then
        FindSheetLastRow = 1
    Else
        FindSheetLastRow = lastCell.Row
    End If
End Function

Private Sub ClearPreparedColumns(ByVal ws As Worksheet, Optional ByVal firstColumn As Long = 2, Optional ByVal lastColumn As Long = 9)
    Dim lastRow As Long

    lastRow = FindSheetLastRow(ws)
    If lastRow >= 2 Then
        ws.Range(ws.Cells(2, firstColumn), ws.Cells(lastRow, lastColumn)).ClearContents
        ws.Range(ws.Cells(2, firstColumn), ws.Cells(lastRow, lastColumn)).Font.ColorIndex = xlAutomatic
    End If
End Sub

Private Sub ClearNomenclatureOutput(ByVal ws As Worksheet)
    Dim lastRow As Long

    lastRow = FindSheetLastRow(ws)
    If lastRow >= 2 Then
        ws.Range("A2:I" & lastRow).ClearContents
        ws.Range("A2:I" & lastRow).Font.ColorIndex = xlAutomatic
    End If
End Sub

Private Sub ClearImportInfoOutput(ByVal ws As Worksheet)
    Dim lastRow As Long

    lastRow = FindSheetLastRow(ws)
    If lastRow >= 2 Then
        ws.Range("A2:N" & lastRow).ClearContents
        ws.Range("A2:N" & lastRow).Font.ColorIndex = xlAutomatic
    End If
End Sub

Private Sub ClearImportInfoBaselineSheetContents()
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = GetImportInfoBaselineSheet
    lastRow = FindSheetLastRow(ws)
    If lastRow >= 2 Then
        ws.Range("A2:J" & lastRow).ClearContents
    End If
    ws.Visible = xlSheetVeryHidden
End Sub

Private Sub ApplyNomenclatureFilter(ByVal ws As Worksheet, ByVal lastRow As Long)
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If

    ws.Range(ws.Cells(1, 1), ws.Cells(Application.Max(lastRow, 1), 9)).AutoFilter
End Sub

Private Sub ApplyImportInfoFilter(ByVal ws As Worksheet, ByVal lastRow As Long)
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If

    ws.Range(ws.Cells(1, 1), ws.Cells(Application.Max(lastRow, 1), 14)).AutoFilter
End Sub

Private Sub ApplyNomenclatureWarningPlaceholders(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim autoColumns As Variant
    Dim columnIndex As Variant
    Dim rowIndex As Long
    Dim targetCell As Range

    If lastRow < 2 Then Exit Sub

    ResetWarningFormatting ws, lastRow, 9
    autoColumns = Array( _
        FindColumn(ws, NOMENCLATURE_HEADER_ORDER), _
        FindColumn(ws, NOMENCLATURE_HEADER_SUPPLIER), _
        FindColumn(ws, NOMENCLATURE_HEADER_CODE), _
        FindColumn(ws, NOMENCLATURE_HEADER_NAME), _
        FindColumn(ws, NOMENCLATURE_HEADER_STATEMENT), _
        FindColumn(ws, NOMENCLATURE_HEADER_QTY), _
        FindColumn(ws, NOMENCLATURE_HEADER_STICKER), _
        FindColumn(ws, NOMENCLATURE_HEADER_MIL))

    For rowIndex = 2 To lastRow
        For Each columnIndex In autoColumns
            Set targetCell = ws.Cells(rowIndex, CLng(columnIndex))
            If Len(NormalizeStageText(targetCell.Value)) = 0 Then
                SetWarningCell targetCell, MISSING_DATA_TEXT
            ElseIf IsWarningText(targetCell.Value) Then
                targetCell.Font.Color = RGB(192, 0, 0)
            End If
        Next columnIndex
    Next rowIndex
End Sub

Private Sub ApplyImportInfoWarningPlaceholders(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim autoColumns As Variant
    Dim manualColumns As Variant
    Dim columnIndex As Variant
    Dim rowIndex As Long
    Dim targetCell As Range

    If lastRow < 2 Then Exit Sub

    ResetWarningFormatting ws, lastRow, 14
    autoColumns = Array( _
        FindColumn(ws, IMPORT_INFO_HEADER_ORDER), _
        FindColumn(ws, IMPORT_INFO_HEADER_SUPPLIER), _
        FindColumn(ws, IMPORT_INFO_HEADER_CODE), _
        FindColumn(ws, IMPORT_INFO_HEADER_NAME), _
        FindColumn(ws, IMPORT_INFO_HEADER_STATEMENT), _
        FindColumn(ws, IMPORT_INFO_HEADER_QTY), _
        FindColumn(ws, IMPORT_INFO_HEADER_STICKER), _
        FindColumn(ws, IMPORT_INFO_HEADER_SHIPMENT_INVOICE), _
        FindColumn(ws, IMPORT_INFO_HEADER_SHIPMENT_DATE), _
        FindColumn(ws, IMPORT_INFO_HEADER_MIL))
    manualColumns = Array( _
        FindColumn(ws, IMPORT_INFO_HEADER_ALCOHOL), _
        FindColumn(ws, IMPORT_INFO_HEADER_VOLUME), _
        FindColumn(ws, IMPORT_INFO_HEADER_VINTAGE))

    For rowIndex = 2 To lastRow
        For Each columnIndex In autoColumns
            Set targetCell = ws.Cells(rowIndex, CLng(columnIndex))
            If Len(NormalizeStageText(targetCell.Value)) = 0 Then
                SetWarningCell targetCell, MISSING_DATA_TEXT
            ElseIf IsWarningText(targetCell.Value) Then
                targetCell.Font.Color = RGB(192, 0, 0)
            End If
        Next columnIndex

        For Each columnIndex In manualColumns
            SetWarningCell ws.Cells(rowIndex, CLng(columnIndex)), MANUAL_FILL_TEXT
        Next columnIndex
    Next rowIndex
End Sub

Private Sub ResetWarningFormatting(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastColumn As Long)
    If lastRow < 2 Then Exit Sub
    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastColumn)).Font.ColorIndex = xlAutomatic
End Sub

Private Sub SetWarningCell(ByVal targetCell As Range, ByVal warningText As String)
    targetCell.Value = warningText
    targetCell.Font.Color = RGB(192, 0, 0)
End Sub

Private Function IsWarningText(ByVal value As Variant) As Boolean
    Dim normalizedValue As String

    normalizedValue = NormalizeStageText(value)
    IsWarningText = StrComp(normalizedValue, MISSING_DATA_TEXT, vbTextCompare) = 0 Or _
                    StrComp(normalizedValue, MANUAL_FILL_TEXT, vbTextCompare) = 0 Or _
                    StrComp(normalizedValue, MULTIPLE_MATCHES_TEXT, vbTextCompare) = 0
End Function

Private Sub ShowNomenclatureRequestError(ByVal message As String, Optional ByVal activateFsmSheet As Boolean = False, Optional ByVal style As VbMsgBoxStyle = vbCritical)
    ShowStagingRequestError message, NOMENCLATURE_REQUEST_SHEET_NAME, "PrepareNomenclatureRequest", activateFsmSheet, style
End Sub

Private Sub ShowImportInfoRequestError(ByVal message As String, Optional ByVal activateFsmSheet As Boolean = False, Optional ByVal style As VbMsgBoxStyle = vbCritical)
    ShowStagingRequestError message, IMPORT_INFO_REQUEST_SHEET_NAME, "PrepareImportInfoRequest", activateFsmSheet, style
End Sub

Private Sub ShowStagingRequestError(ByVal message As String, ByVal requestSheetName As String, ByVal sourceProcedure As String, Optional ByVal activateFsmSheet As Boolean = False, Optional ByVal style As VbMsgBoxStyle = vbCritical)
    On Error Resume Next
    EnsureInteractiveSheetProtection
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If activateFsmSheet Then
        GetFsmRequestSheet.Activate
    Else
        ThisWorkbook.Worksheets(requestSheetName).Activate
    End If
    On Error GoTo 0

    If AutomationModeEnabled Then
        SetAutomationResult "error", message
        Err.Raise AUTOMATION_ERROR_NUMBER, sourceProcedure, message
    End If

    MsgBox message, style
    End
End Sub

Private Sub ResetAutomationResult()
    AutomationLastStatus = ""
    AutomationLastMessage = ""
End Sub

Private Sub ResetAutomationContext()
    ResetAutomationResult
    Set AutomationConfirmations = Nothing
End Sub

Private Sub EndAutomationRun()
    AutomationModeEnabled = False
    Set AutomationConfirmations = Nothing
End Sub

Public Sub SetAutomationResult(ByVal statusValue As String, ByVal messageValue As String)
    AutomationLastStatus = statusValue
    AutomationLastMessage = messageValue
End Sub

Private Sub TraceAutomationStep(ByVal stepText As String)
    Dim tracePath As String
    Dim fileNumber As Integer

    If Not AutomationModeEnabled Then Exit Sub
    If Len(ThisWorkbook.Path) = 0 Then Exit Sub

    tracePath = ThisWorkbook.Path & Application.PathSeparator & AUTOMATION_TRACE_FILE_NAME
    fileNumber = FreeFile

    On Error Resume Next
    Open tracePath For Append Access Write Lock Write As #fileNumber
    If Err.Number = 0 Then
        Print #fileNumber, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & SanitizeTraceText(stepText)
    End If
    Close #fileNumber
    On Error GoTo 0
End Sub

Private Function SanitizeTraceText(ByVal value As String) As String
    Dim index As Long
    Dim codePoint As Long
    Dim result As String
    Dim currentChar As String

    For index = 1 To Len(value)
        currentChar = Mid$(value, index, 1)
        codePoint = AscW(currentChar)
        If codePoint >= 32 And codePoint <= 126 Then
            result = result & currentChar
        Else
            result = result & "_"
        End If
    Next index

    SanitizeTraceText = result
End Function
