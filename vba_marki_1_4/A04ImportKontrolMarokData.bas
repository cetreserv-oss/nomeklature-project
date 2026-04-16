Option Explicit

Public Sub ImportKontrolMarokData()
    On Error GoTo ErrHandler

    Dim wbKM As Workbook
    Dim wsFSM As Worksheet
    Dim wsWork As Worksheet
    Dim kontrolMarokPath As String
    Dim lastRowWork As Long
    Dim i As Long
    Dim orders As Collection
    Dim orderVal As String
    Dim ordersFound As Object
    Dim headerFSM As Object
    Dim headerWork As Object
    Dim dataFSM As Variant
    Dim r As Long
    Dim lastRowFSM As Long
    Dim lastColFSM As Long
    Dim c As Long
    Dim colName As String
    Dim needCols As Variant
    Dim key As Variant
    Dim rowOut As Long
    Dim missingOrders As Collection
    Dim missingOrderText As String

    Set wsWork = GetFsmRequestSheet
    kontrolMarokPath = Trim(GetSettingsValue("Контроль марок", GetSettingsSheet.Range("B3").Value))

    If Len(kontrolMarokPath) = 0 Then
        Err.Raise vbObjectError + 513, , "Не указан путь к файлу 'Контроль марок.xlsx' на листе '" & SETTINGS_SHEET_NAME & "'."
    End If

    If Dir(kontrolMarokPath) = "" Then
        Err.Raise vbObjectError + 514, , "Файл 'Контроль марок.xlsx' не найден: " & kontrolMarokPath
    End If

    Set wbKM = ReopenKontrolMarokWorkbook(kontrolMarokPath)

    On Error Resume Next
    Set wsFSM = wbKM.Worksheets("ФСМ")
    On Error GoTo ErrHandler

    If wsFSM Is Nothing Then
        Err.Raise vbObjectError + 515, , "Не найдена вкладка 'ФСМ' в файле 'Контроль марок.xlsx'."
    End If

    wsFSM.AutoFilterMode = False
    wsFSM.Cells.EntireColumn.Hidden = False

    lastRowWork = wsWork.Cells(wsWork.Rows.Count, "A").End(xlUp).Row
    Set orders = New Collection

    For i = 2 To lastRowWork
        orderVal = Trim(CStr(wsWork.Cells(i, 1).Value))
        If orderVal <> "" Then orders.Add orderVal
    Next i

    If orders.Count = 0 Then
        MsgBox "Во вкладке '" & wsWork.Name & "' нет заказов.", vbInformation
        wbKM.Close SaveChanges:=False
        Exit Sub
    End If

    Set ordersFound = CreateObject("Scripting.Dictionary")
    Set headerFSM = CreateObject("Scripting.Dictionary")
    Set headerWork = CreateObject("Scripting.Dictionary")
    Set missingOrders = New Collection

    lastRowFSM = wsFSM.Cells(wsFSM.Rows.Count, 1).End(xlUp).Row
    lastColFSM = wsFSM.Cells(1, wsFSM.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastColFSM
        colName = Trim(CStr(wsFSM.Cells(1, c).Value))
        If colName <> "" Then headerFSM(colName) = c
    Next c

    needCols = Array("Заказ", "Заявление", "Поставщик", "Код", "Позиция", "Кол-во")
    For Each key In needCols
        If Not headerFSM.Exists(CStr(key)) Then
            Err.Raise vbObjectError + 516, , "Не найден столбец '" & CStr(key) & "' на листе 'ФСМ' в файле 'Контроль марок.xlsx'."
        End If
    Next key

    headerWork("Заказ") = FindColumn(wsWork, "Заказ")
    headerWork("Заявление") = FindColumn(wsWork, "Заявление (КМ)")
    headerWork("Поставщик") = FindColumn(wsWork, "Поставщик (КМ)")
    headerWork("Код") = FindColumn(wsWork, "Код (КМ)")
    headerWork("Позиция") = FindColumn(wsWork, "Позиция (КМ)")
    headerWork("Кол-во") = FindColumn(wsWork, "Кол-во (КМ)")

    dataFSM = wsFSM.Range(wsFSM.Cells(2, 1), wsFSM.Cells(lastRowFSM, lastColFSM)).Value
    rowOut = wsWork.Cells(wsWork.Rows.Count, "A").End(xlUp).Row + 1

    For i = 1 To orders.Count
        orderVal = orders(i)
        Dim found As Boolean
        found = False

        For r = 1 To UBound(dataFSM, 1)
            If NormalizeCyrLat(Trim(CStr(dataFSM(r, headerFSM("Заказ"))))) = NormalizeCyrLat(orderVal) Then
                found = True

                wsWork.Cells(rowOut, headerWork("Заказ")).Value = NormalizeCyrLat(CStr(dataFSM(r, headerFSM("Заказ"))))
                wsWork.Cells(rowOut, headerWork("Заявление")).Value = dataFSM(r, headerFSM("Заявление"))
                wsWork.Cells(rowOut, headerWork("Поставщик")).Value = dataFSM(r, headerFSM("Поставщик"))
                wsWork.Cells(rowOut, headerWork("Код")).Value = dataFSM(r, headerFSM("Код"))
                wsWork.Cells(rowOut, headerWork("Позиция")).Value = dataFSM(r, headerFSM("Позиция"))
                wsWork.Cells(rowOut, headerWork("Кол-во")).Value = dataFSM(r, headerFSM("Кол-во"))

                rowOut = rowOut + 1
            End If
        Next r

        If found Then ordersFound(orderVal) = True
    Next i

    For i = 1 To orders.Count
        orderVal = CStr(orders(i))
        If Not ordersFound.Exists(orderVal) Then
            missingOrders.Add orderVal
        End If
    Next i

    If missingOrders.Count > 0 Then
        If missingOrders.Count = 1 Then
            missingOrderText = CStr(missingOrders(1))
            Err.Raise vbObjectError + 518, , "Заказ '" & missingOrderText & "' не найден среди данных файла 'Контроль марок.xlsx'. Проверьте номер заказа. Буквы 'ТК' можно вводить кириллицей: макрос сам приводит их к 'TK'."
        Else
            missingOrderText = JoinOrdersForMessage(missingOrders)
            Err.Raise vbObjectError + 518, , "Заказы " & missingOrderText & " не найдены среди данных файла 'Контроль марок.xlsx'. Проверьте номера заказов. Буквы 'ТК' можно вводить кириллицей: макрос сам приводит их к 'TK'."
        End If
    End If

    For i = lastRowWork To 2 Step -1
        orderVal = Trim(CStr(wsWork.Cells(i, 1).Value))
        If ordersFound.Exists(orderVal) Then
            wsWork.Rows(i).Delete
        End If
    Next i

    wbKM.Close SaveChanges:=False
    Exit Sub

ErrHandler:
    Dim originalErrorNumber As Long
    Dim originalErrorDescription As String

    originalErrorNumber = Err.Number
    originalErrorDescription = Err.Description

    On Error Resume Next
    If Not wbKM Is Nothing Then wbKM.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise originalErrorNumber, "ImportKontrolMarokData", originalErrorDescription
End Sub

Private Function ReopenKontrolMarokWorkbook(ByVal kontrolMarokPath As String) As Workbook
    Dim wb As Workbook
    Dim fileName As String

    fileName = Dir(kontrolMarokPath)

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, kontrolMarokPath, vbTextCompare) = 0 Or StrComp(wb.Name, fileName, vbTextCompare) = 0 Then
            If Not wb.ReadOnly Then
                On Error GoTo SaveError
                wb.Save
                On Error GoTo 0
            End If

            wb.Close SaveChanges:=False
            Exit For
        End If
    Next wb

    Set ReopenKontrolMarokWorkbook = Workbooks.Open(FileName:=kontrolMarokPath, ReadOnly:=True, UpdateLinks:=False)
    Exit Function

SaveError:
    Err.Raise vbObjectError + 517, , "Не удалось сохранить открытый файл 'Контроль марок.xlsx'. Сохраните его вручную или закройте, затем повторите попытку."
End Function

Private Function JoinOrdersForMessage(ByVal orders As Collection) As String
    Dim i As Long

    For i = 1 To orders.Count
        If Len(JoinOrdersForMessage) > 0 Then
            JoinOrdersForMessage = JoinOrdersForMessage & ", "
        End If
        JoinOrdersForMessage = JoinOrdersForMessage & "'" & CStr(orders(i)) & "'"
    Next i
End Function
