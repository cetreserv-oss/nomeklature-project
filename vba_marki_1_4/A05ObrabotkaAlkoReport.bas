Option Explicit

Public Sub ObrabotkaAlkoReport()
    On Error GoTo ErrHandler

    Dim wb As Workbook
    Dim wsAlko As Worksheet
    Dim wsWork As Worksheet
    Dim headersAlko As Variant
    Dim colMapAlko As Object
    Dim i As Long
    Dim foundCol As Range
    Dim srcWorkOrderCol As Range
    Dim workOrderCol As Long
    Dim colProvider As Range
    Dim colCode As Range
    Dim colPos As Range
    Dim colQty As Range
    Dim colAction As Range
    Dim lastRowAlko As Long
    Dim lastRowWork As Long
    Dim r As Long
    Dim dictWorkOrders As Object
    Dim dictAlkoOrders As Object
    Dim wRow As Long
    Dim ord As String
    Dim key As Variant
    Dim ordA As String
    Dim colQtyInDelivery As Long
    Dim valQtyInDelivery As Variant

    Set wb = ThisWorkbook
    Set wsAlko = wb.Worksheets(ALCO_REPORT_SHEET_NAME)
    Set wsWork = GetFsmRequestSheet

    wsAlko.AutoFilterMode = False
    wsAlko.Cells.EntireColumn.Hidden = False

    headersAlko = Array("Номер заказа", "Поставщик", "Код номенклатуры", "Наименование товара (рус)", "Заказанный объем", "Статус заказа", "Комментарии КМ")
    Set colMapAlko = CreateObject("Scripting.Dictionary")

    For i = LBound(headersAlko) To UBound(headersAlko)
        Set foundCol = wsAlko.Rows(1).Find(What:=headersAlko(i), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        If foundCol Is Nothing Then
            Err.Raise vbObjectError + 901, , "Столбец '" & headersAlko(i) & "' отсутствует в выгрузке Алкоотчета. Проверьте корректность выгрузки."
        End If

        colMapAlko(headersAlko(i)) = foundCol.Column
    Next i

    Set srcWorkOrderCol = wsWork.Rows(1).Find(What:="Заказ", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If srcWorkOrderCol Is Nothing Then
        Err.Raise vbObjectError + 902, , "Столбец 'Заказ' отсутствует на листе '" & wsWork.Name & "'."
    End If

    workOrderCol = srcWorkOrderCol.Column

    Set colProvider = wsWork.Rows(1).Find(What:="Поставщик (новый)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    Set colCode = wsWork.Rows(1).Find(What:="Код (новый)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    Set colPos = wsWork.Rows(1).Find(What:="Позиция (новая)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    Set colQty = wsWork.Rows(1).Find(What:="Кол-во (новое)", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    Set colAction = wsWork.Rows(1).Find(What:="Действие", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)

    If colProvider Is Nothing Or colCode Is Nothing Or colPos Is Nothing Or colQty Is Nothing Or colAction Is Nothing Then
        Err.Raise vbObjectError + 903, , "На листе '" & wsWork.Name & "' отсутствуют один или несколько целевых столбцов: 'Поставщик (новый)', 'Код (новый)', 'Позиция (новая)', 'Кол-во (новое)', 'Действие'."
    End If

    lastRowAlko = wsAlko.Cells(wsAlko.Rows.Count, 1).End(xlUp).Row
    lastRowWork = wsWork.Cells(wsWork.Rows.Count, workOrderCol).End(xlUp).Row

    For r = 2 To lastRowAlko
        Dim commVal As String
        commVal = Trim(CStr(wsAlko.Cells(r, colMapAlko("Комментарии КМ")).Value))
        If Len(commVal) >= 7 Then
            If LCase$(Left$(commVal, 7)) = LCase$("GKF-000") Then
                wsAlko.Cells(r, colMapAlko("Номер заказа")).Value = commVal
            End If
        End If
    Next r

    Set dictWorkOrders = CreateObject("Scripting.Dictionary")
    dictWorkOrders.CompareMode = vbTextCompare

    For wRow = 2 To lastRowWork
        ord = Trim(CStr(wsWork.Cells(wRow, workOrderCol).Value))
        If Len(ord) > 0 Then
            If Not dictWorkOrders.Exists(ord) Then dictWorkOrders.Add ord, 0
        End If
    Next wRow

    Set dictAlkoOrders = CreateObject("Scripting.Dictionary")
    dictAlkoOrders.CompareMode = vbTextCompare

    For r = 2 To lastRowAlko
        ordA = Trim(CStr(wsAlko.Cells(r, colMapAlko("Номер заказа")).Value))
        If Len(ordA) > 0 Then
            If Not dictAlkoOrders.Exists(ordA) Then dictAlkoOrders.Add ordA, 0
        End If
    Next r

    For Each key In dictWorkOrders.Keys
        If Not dictAlkoOrders.Exists(key) Then
            Err.Raise vbObjectError + 904, , "Заказа " & key & " нет в выгрузке Алкоотчета. Проверьте корректность выгрузки или верно ли указан заказ."
        End If
    Next key

    For r = lastRowAlko To 2 Step -1
        ordA = Trim(CStr(wsAlko.Cells(r, colMapAlko("Номер заказа")).Value))
        If Len(ordA) = 0 Or Not dictWorkOrders.Exists(ordA) Then
            wsAlko.Rows(r).Delete
        End If
    Next r

    lastRowAlko = wsAlko.Cells(wsAlko.Rows.Count, 1).End(xlUp).Row

    Set foundCol = wsAlko.Rows(1).Find(What:="Кол-во шт / кг в поставке", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If foundCol Is Nothing Then
        Err.Raise vbObjectError + 905, , "Столбец 'Кол-во шт / кг в поставке' отсутствует в выгрузке Алкоотчета."
    End If

    colQtyInDelivery = foundCol.Column

    For r = 2 To lastRowAlko
        valQtyInDelivery = wsAlko.Cells(r, colQtyInDelivery).Value

        If Not IsEmpty(valQtyInDelivery) Then
            If IsNumeric(valQtyInDelivery) Then
                If CDbl(valQtyInDelivery) <> 0 Then
                    wsAlko.Cells(r, colMapAlko("Заказанный объем")).Value = valQtyInDelivery
                End If
            End If
        End If
    Next r

    Exit Sub

ErrHandler:
    Dim originalErrorNumber As Long
    Dim originalErrorDescription As String

    originalErrorNumber = Err.Number
    originalErrorDescription = Err.Description

    On Error Resume Next
    GetAlcoReportSheet.Visible = xlSheetHidden
    On Error GoTo 0
    Err.Raise originalErrorNumber, "ObrabotkaAlkoReport", originalErrorDescription
End Sub
