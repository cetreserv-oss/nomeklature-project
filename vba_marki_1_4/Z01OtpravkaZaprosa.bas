Option Explicit

Public Sub OtpravkaZaprosa()
    On Error GoTo ErrHandler

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim newWb As Workbook
    Dim newWs As Worksheet
    Dim filePath As String
    Dim fileName As String
    Dim userName As String
    Dim outApp As Object
    Dim outMail As Object
    Dim saveFolder As String
    Dim dateStr As String
    Dim timeStr As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim c As Long
    Dim r As Long

    Set wb = ThisWorkbook
    Set ws = GetFsmRequestSheet

    userName = Trim(GetSettingsValue("МИЛ", GetSettingsSheet.Range("B2").Value))
    dateStr = Format(Date, "dd.mm.yyyy")
    timeStr = Format(Time, "hh-mm-ss")

    fileName = "Запрос по маркам (корректировка) " & dateStr & " " & timeStr & " " & userName & ".xlsx"

    saveFolder = EnsureTrailingSlash(GetSettingsValue("Папка корректировка КМ", GetSettingsSheet.Range("B4").Value))
    If Len(saveFolder) = 0 Then
        Err.Raise vbObjectError + 1001, , "Не указан путь для сохранения запроса на листе '" & SETTINGS_SHEET_NAME & "'."
    End If

    If Dir(saveFolder, vbDirectory) = "" Then
        Err.Raise vbObjectError + 1002, , "Папка для сохранения запроса не найдена: " & saveFolder
    End If

    filePath = saveFolder & fileName

    Set newWb = Workbooks.Add(xlWBATWorksheet)
    Set newWs = newWb.Worksheets(1)
    newWs.Name = Left$(ws.Name, 31)

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Copy
    newWs.Range("A1").PasteSpecial xlPasteAll

    For c = 1 To lastCol
        newWs.Columns(c).ColumnWidth = ws.Columns(c).ColumnWidth
    Next c

    For r = 1 To lastRow
        newWs.Rows(r).RowHeight = ws.Rows(r).RowHeight
    Next r

    Application.CutCopyMode = False

    Application.DisplayAlerts = False
    newWb.SaveAs filePath, 51
    Application.DisplayAlerts = True

    On Error Resume Next
    Set outApp = GetObject(Class:="Outlook.Application")
    If outApp Is Nothing Then Set outApp = CreateObject("Outlook.Application")
    On Error GoTo ErrHandler

    Set outMail = outApp.CreateItem(0)

    With outMail
        .To = "sviii@grandtrade.world; slobodchikovas@grandtrade.world; fakira_oe@grandtrade.world; bakhtiarov_rs@grandtrade.world"
        .Subject = fileName
        .Body = "Добрый день, во вложении запрос по ФСМ."
        .Attachments.Add filePath
        .Display
    End With

    newWb.Close SaveChanges:=False
    Exit Sub

ErrHandler:
    On Error Resume Next
    Application.DisplayAlerts = True
    If Not newWb Is Nothing Then newWb.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "Ошибка при формировании запроса: " & Err.Description, vbCritical
End Sub
