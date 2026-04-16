Option Explicit

Public Sub UpdateDataSheets()
    Dim downloadsPath As String
    Dim logosOtchetFile As String

    downloadsPath = EnsureTrailingSlash(GetSettingsValue("Папка загрузки", GetSettingsSheet.Range("B1").Value))
    If Len(downloadsPath) = 0 Then
        Err.Raise vbObjectError + 801, , "Не удалось получить путь к папке загрузки из листа '" & SETTINGS_SHEET_NAME & "'."
    End If

    If Dir(downloadsPath, vbDirectory) = "" Then
        Err.Raise vbObjectError + 802, , "Папка загрузки не найдена: " & downloadsPath
    End If

    GetAlcoReportSheet.Visible = xlSheetVisible

    logosOtchetFile = FindLatestFile(downloadsPath, "_ALCOHOL_REPORT")
    If logosOtchetFile = "" Then
        Err.Raise vbObjectError + 803, , "Файл Алкоотчета не найден в папке загрузки: " & downloadsPath
    End If

    ReplaceSheetWithFile ALCO_REPORT_SHEET_NAME, downloadsPath & logosOtchetFile
End Sub

Public Function FindLatestFile(folderPath As String, pattern As String, Optional exclude As String = "") As String
    Dim fileName As String
    Dim latestFile As String
    Dim latestDate As Date
    Dim fileDate As Date
    Dim shouldExclude As Boolean
    Dim excludeParts() As String
    Dim ex As Variant

    fileName = Dir(folderPath & "*.xlsx")
    Do While fileName <> ""
        If InStr(1, fileName, pattern, vbTextCompare) > 0 Then
            shouldExclude = False

            If exclude <> "" Then
                excludeParts = Split(exclude, "|")
                For Each ex In excludeParts
                    If InStr(1, fileName, CStr(ex), vbTextCompare) > 0 Then
                        shouldExclude = True
                        Exit For
                    End If
                Next ex
            End If

            If Not shouldExclude Then
                fileDate = GetFileDateFromName(fileName)
                If fileDate > latestDate Then
                    latestDate = fileDate
                    latestFile = fileName
                End If
            End If
        End If

        fileName = Dir
    Loop

    FindLatestFile = latestFile
End Function

Public Function GetFileDateFromName(fileName As String) As Date
    On Error Resume Next

    Dim datePart As String

    datePart = Left$(fileName, 15)
    GetFileDateFromName = CDate(Left$(datePart, 4) & "-" & Mid$(datePart, 6, 2) & "-" & Mid$(datePart, 9, 2) & " " & _
                                Mid$(datePart, 12, 2) & ":" & Mid$(datePart, 14, 2) & ":00")

    On Error GoTo 0
End Function

Public Sub ReplaceSheetWithFile(sheetName As String, sourceFilePath As String)
    Dim wbSource As Workbook
    Dim wsTarget As Worksheet
    Dim wsSource As Worksheet

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo FileOpenError

    Set wbSource = Workbooks.Open(fileName:=sourceFilePath, UpdateLinks:=False, ReadOnly:=True, CorruptLoad:=xlRepairFile)
    Set wsSource = wbSource.Sheets(1)
    Set wsTarget = ThisWorkbook.Sheets(sheetName)

    wsTarget.Cells.Clear
    wsSource.UsedRange.Copy wsTarget.Range("A1")

    wbSource.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

FileOpenError:
    On Error Resume Next
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    On Error GoTo 0
    Err.Raise vbObjectError + 804, , "Не удалось открыть файл: " & sourceFilePath
End Sub
