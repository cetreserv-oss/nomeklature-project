$ErrorActionPreference = "Stop"
Add-Type -AssemblyName UIAutomationClient

$RootDir = "E:\Projects\nomenklature project"
$ReferenceDir = Join-Path $RootDir "reference"
$VbaSourceDir = Join-Path $RootDir "vba_marki_1_4"
$SourceWorkbook = Join-Path $ReferenceDir "Запрос по маркам 1.3.xlsm"
$OutputWorkbook = Join-Path $ReferenceDir "Запрос по маркам 1.4.xlsm"
$DefaultNomenclatureWorkbook = Join-Path $ReferenceDir "Номенклатура New 2.xlsx"
$DefaultArchiveWorkbook = Join-Path $ReferenceDir "Архив исходных запросов.xlsx"
$DefaultDownloadsDir = Join-Path $env:USERPROFILE "Downloads"

$FsmSheetName = "Отправить запрос по ФСМ"
$NomenclatureSheetName = "Отправка марок (номенклатура)"
$ImportInfoSheetName = "Сведения о ввозе (номенклатура)"
$ImportInfoBaselineSheetName = "__Ввоз_База"

$NomenclatureHeaders = @(
    "Номер заказа",
    "Поставщик",
    "Код УТ",
    "Номенклатура",
    "Заявление на выдачу ФСМ",
    "Кол-во",
    "Оклейщик",
    "МИЛ",
    "Комментарий МИЛ"
)
$ImportInfoHeaders = @(
    "Номер заказа",
    "Поставщик",
    "Код УТ",
    "Номенклатура",
    "Заявление на выдачу ФСМ",
    "Количество",
    "Оклейщик",
    "Накладная отправки",
    "Дата отправки",
    "Градус алкоголя",
    "Объем бутылки",
    "Год урожая",
    "МИЛ",
    "Комментарий МИЛ"
)
$ImportInfoBaselineHeaders = @(
    "Номер заказа",
    "Поставщик",
    "Код УТ",
    "Номенклатура",
    "Заявление на выдачу ФСМ",
    "Количество",
    "Оклейщик",
    "Накладная отправки",
    "Дата отправки",
    "МИЛ"
)

$NomenclatureWidths = @(18, 24, 16, 36, 24, 12, 24, 18, 24)
$ImportInfoWidths = @(18, 24, 16, 36, 24, 12, 24, 22, 16, 18, 18, 16, 18, 24)
$ImportInfoBaselineWidths = @(18, 24, 16, 36, 24, 12, 24, 22, 16, 18)
$SendArchiveSheetName = "Исх. запросы (отпр ФСМ)"
$SendCorrectionArchiveSheetName = "Корр. (отпр ФСМ)"
$ImportInfoArchiveSheetName = "Исх. запросы (ввоз)"
$ImportInfoCorrectionArchiveSheetName = "Корр. (ввоз)"
$ArchiveWorkbookPassword = "7777"
$SendArchiveHeaders = @(
    "Номер заказа",
    "Код УТ",
    "Номенклатура",
    "Заявление на выдачу ФСМ",
    "Кол-во",
    "Оклейщик",
    "Поставщик",
    "Комментарий МИЛ",
    "МИЛ",
    "Дата внесения строки",
    "Подтверждение повторной отправки"
)
$SendCorrectionArchiveHeaders = @(
    "Номер заказа",
    "Код УТ",
    "Номенклатура",
    "Заявление на выдачу ФСМ",
    "Кол-во",
    "Оклейщик",
    "Поставщик",
    "Комментарий МИЛ",
    "МИЛ",
    "Дата внесения строки"
)
$ImportInfoArchiveHeaders = $ImportInfoHeaders + @(
    "Дата внесения строки",
    "Подтверждение повторной отправки",
    "Подтверждение несоответствия количества"
)
$ImportInfoCorrectionArchiveHeaders = $ImportInfoHeaders + @(
    "Дата внесения строки",
    "Подтверждение несоответствия количества"
)
$SendArchiveWidths = @(18, 16, 36, 24, 12, 24, 24, 24, 18, 18, 24)
$SendCorrectionArchiveWidths = @(18, 16, 36, 24, 12, 24, 24, 24, 18, 18)
$ImportInfoArchiveWidths = @(18, 24, 16, 36, 24, 12, 24, 22, 16, 18, 18, 16, 18, 24, 18, 24, 32)
$ImportInfoCorrectionArchiveWidths = @(18, 24, 16, 36, 24, 12, 24, 22, 16, 18, 18, 16, 18, 24, 18, 32)

$ButtonWidth = 250
$ButtonHeight = 54
$ButtonVerticalGap = 6
$ButtonLeftOffset = 12
$PrepareCaption = "Подготовить строки к внесению в номенклатуру"
$SendCaption = "Внести строчки в номенклатуру"
$CorrectionCaption = "Внести корректировку в номенклатуру"

$SettingDownloads = "Папка загрузки"
$SettingNomenclature = "Номенклатура"
$SettingArchive = "Архив исходных запросов"
$SettingNomenclaturePassword = "Пароль номенклатуры"
$SettingProtectedSheetsPassword = "Пароль защищенных вкладок номенклатуры"
$SettingsNote = "Перед нажатием кнопок нужно сделать выгрузку из Алкоотчета и актуализировать Логос"

$ModuleSourceFiles = [ordered]@{
    "A00MainModule" = Join-Path $VbaSourceDir "A00MainModule.bas"
    "A02FormatZakaz" = Join-Path $VbaSourceDir "A02FormatZakaz.bas"
    "A03UpdateDataSheets" = Join-Path $VbaSourceDir "A03UpdateDataSheets.bas"
    "A04ImportKontrolMarokData" = Join-Path $VbaSourceDir "A04ImportKontrolMarokData.bas"
    "A05ObrabotkaAlkoReport" = Join-Path $VbaSourceDir "A05ObrabotkaAlkoReport.bas"
    "Z01OtpravkaZaprosa" = Join-Path $VbaSourceDir "Z01OtpravkaZaprosa.bas"
    "Z02NomenclatureRequest" = Join-Path $VbaSourceDir "Z02NomenclatureRequest.bas"
}

function Release-ComObject($ComObject) {
    if ($null -ne $ComObject) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

function Invoke-ComRetry([scriptblock]$Script, [int]$Retries = 30, [int]$DelayMilliseconds = 500) {
    for ($attempt = 1; $attempt -le $Retries; $attempt++) {
        try {
            return & $Script
        }
        catch {
            if ($attempt -eq $Retries) {
                throw
            }
            Start-Sleep -Milliseconds $DelayMilliseconds
        }
    }
}

function Close-OfficeActivationWizard {
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        $closed = $false
        $processes = Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }
        foreach ($process in $processes) {
            try {
                $root = [System.Windows.Automation.AutomationElement]::FromHandle($process.MainWindowHandle)
                if ($null -eq $root) { continue }

                $dialogs = $root.FindAll(
                    [System.Windows.Automation.TreeScope]::Descendants,
                    [System.Windows.Automation.PropertyCondition]::new(
                        [System.Windows.Automation.AutomationElement]::NameProperty,
                        "Microsoft Office Activation Wizard"
                    )
                )

                for ($dialogIndex = 0; $dialogIndex -lt $dialogs.Count; $dialogIndex++) {
                    $dialog = $dialogs.Item($dialogIndex)
                    $buttons = $dialog.FindAll(
                        [System.Windows.Automation.TreeScope]::Descendants,
                        [System.Windows.Automation.PropertyCondition]::new(
                            [System.Windows.Automation.AutomationElement]::NameProperty,
                            "Close"
                        )
                    )

                    for ($i = 0; $i -lt $buttons.Count; $i++) {
                        $button = $buttons.Item($i)
                        $pattern = $null
                        if ($button.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$pattern)) {
                            $pattern.Invoke()
                            $closed = $true
                            Start-Sleep -Seconds 1
                            break
                        }
                    }
                    if ($closed) { break }
                }
            }
            catch {
            }
        }

        if ($closed) { return }
        Start-Sleep -Milliseconds 500
    }
}

function Find-SettingRow($Worksheet, [string]$Label) {
    for ($row = 1; $row -le 50; $row++) {
        $value = [string](Invoke-ComRetry { $Worksheet.Cells.Item($row, 1).Value2 })
        if ($value.Trim().ToLowerInvariant() -eq $Label.ToLowerInvariant()) {
            return $row
        }
    }
    return $null
}

function Ensure-SettingRow($Worksheet, [string]$Label, [int]$PreferredRow) {
    $existing = Find-SettingRow $Worksheet $Label
    if ($null -ne $existing) {
        return $existing
    }

    $Worksheet.Rows.Item($PreferredRow).Insert() | Out-Null
    $Worksheet.Cells.Item($PreferredRow, 1).Value2 = $Label
    return $PreferredRow
}

function Get-OrCreateSheet($Workbook, [string]$SheetName, $AfterSheet = $null) {
    try {
        return $Workbook.Worksheets($SheetName)
    }
    catch {
        if ($null -eq $AfterSheet) {
            $AfterSheet = $Workbook.Worksheets($Workbook.Worksheets.Count)
        }
        $sheet = Invoke-ComRetry { $Workbook.Worksheets.Add() }
        $sheet.Name = $SheetName
        return $sheet
    }
}

function Format-StagingSheet($Excel, $Worksheet, [string[]]$Headers, [int[]]$Widths) {
    try { $Worksheet.Unprotect("") | Out-Null } catch {}
    $Worksheet.Cells.Clear() | Out-Null

    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $Worksheet.Cells.Item(1, $i + 1).Value2 = $Headers[$i]
    }

    $headerRange = $Worksheet.Range($Worksheet.Cells.Item(1, 1), $Worksheet.Cells.Item(1, $Headers.Count))
    $headerRange.Interior.Color = 65535
    $headerRange.Font.Bold = $true
    $headerRange.HorizontalAlignment = -4108
    $headerRange.VerticalAlignment = -4108
    $headerRange.WrapText = $true
    $Worksheet.Rows.Item(1).RowHeight = 30

    for ($i = 0; $i -lt $Widths.Count; $i++) {
        $Worksheet.Columns.Item($i + 1).ColumnWidth = $Widths[$i]
    }

    if ($Worksheet.AutoFilterMode) {
        $Worksheet.AutoFilterMode = $false
    }
    $Worksheet.Range($Worksheet.Cells.Item(1, 1), $Worksheet.Cells.Item(1, $Headers.Count)).AutoFilter() | Out-Null

    $Worksheet.Activate() | Out-Null
    $Excel.ActiveWindow.SplitColumn = 0
    $Excel.ActiveWindow.SplitRow = 1
    $Excel.ActiveWindow.FreezePanes = $false
    $Excel.ActiveWindow.FreezePanes = $true
}

function Configure-StagingButtons($Worksheet, [int]$HeaderCount, [object[]]$ButtonSpecs) {
    while ($Worksheet.Buttons().Count -gt 0) {
        $Worksheet.Buttons().Item(1).Delete()
    }

    $left = $Worksheet.Cells.Item(2, $HeaderCount + 1).Left + $ButtonLeftOffset
    $top = $Worksheet.Cells.Item(2, $HeaderCount + 1).Top

    for ($i = 0; $i -lt $ButtonSpecs.Count; $i++) {
        $spec = $ButtonSpecs[$i]
        $button = $Worksheet.Buttons().Add(
            $left,
            $top + $i * ($ButtonHeight + $ButtonVerticalGap),
            $ButtonWidth,
            $ButtonHeight
        )
        $button.Caption = $spec.Caption
        $button.OnAction = $spec.Action
        $button.Locked = $false
    }

    foreach ($shape in $Worksheet.Shapes) {
        $shape.Placement = 3
    }
}

function Find-HeaderColumn($Worksheet, [string]$Header) {
    for ($column = 1; $column -le 80; $column++) {
        $value = [string]$Worksheet.Cells.Item(1, $column).Value2
        if ($value.Trim().ToLowerInvariant() -eq $Header.ToLowerInvariant()) {
            return $column
        }
    }
    return $null
}

function Format-ArchiveSheet($Worksheet, [object[]]$Headers, [int[]]$Widths) {
    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $Worksheet.Cells.Item(1, $i + 1).Value2 = $Headers[$i]
    }

    $headerRange = $Worksheet.Range($Worksheet.Cells.Item(1, 1), $Worksheet.Cells.Item(1, $Headers.Count))
    $headerRange.Interior.Color = 65535
    $headerRange.Font.Bold = $true
    $headerRange.HorizontalAlignment = -4108
    $headerRange.VerticalAlignment = -4108
    $headerRange.WrapText = $true
    $Worksheet.Rows.Item(1).RowHeight = 30

    for ($i = 0; $i -lt $Widths.Count; $i++) {
        $Worksheet.Columns.Item($i + 1).ColumnWidth = $Widths[$i]
    }

    $Worksheet.Activate() | Out-Null
    $Worksheet.Application.ActiveWindow.SplitColumn = 0
    $Worksheet.Application.ActiveWindow.SplitRow = 1
    $Worksheet.Application.ActiveWindow.FreezePanes = $false
    $Worksheet.Application.ActiveWindow.FreezePanes = $true
}

function Get-OrCreateArchiveSheet($Workbook, [string]$SheetName, [string]$LegacyName = "") {
    try {
        return $Workbook.Worksheets($SheetName)
    }
    catch {
        if (-not [string]::IsNullOrWhiteSpace($LegacyName)) {
            try {
                $sheet = $Workbook.Worksheets($LegacyName)
                $sheet.Name = $SheetName
                return $sheet
            }
            catch {
            }
        }
        $sheet = $Workbook.Worksheets.Add()
        $sheet.Name = $SheetName
        return $sheet
    }
}

function Create-ArchiveWorkbook($Excel, [string]$ArchivePath) {
    $archiveDir = Split-Path -Parent $ArchivePath
    if (-not (Test-Path -LiteralPath $archiveDir)) {
        New-Item -ItemType Directory -Path $archiveDir -Force | Out-Null
    }

    $archiveWorkbook = $Excel.Workbooks.Add()
    while ($archiveWorkbook.Worksheets.Count -gt 1) {
        $archiveWorkbook.Worksheets.Item($archiveWorkbook.Worksheets.Count).Delete()
    }

    $mainSheet = $archiveWorkbook.Worksheets.Item(1)
    $mainSheet.Name = $SendArchiveSheetName
    Format-ArchiveSheet $mainSheet $SendArchiveHeaders $SendArchiveWidths

    $sendCorrectionSheet = $archiveWorkbook.Worksheets.Add()
    $sendCorrectionSheet.Name = $SendCorrectionArchiveSheetName
    Format-ArchiveSheet $sendCorrectionSheet $SendCorrectionArchiveHeaders $SendCorrectionArchiveWidths

    $importInfoSheet = $archiveWorkbook.Worksheets.Add()
    $importInfoSheet.Name = $ImportInfoArchiveSheetName
    Format-ArchiveSheet $importInfoSheet $ImportInfoArchiveHeaders $ImportInfoArchiveWidths

    $importInfoCorrectionSheet = $archiveWorkbook.Worksheets.Add()
    $importInfoCorrectionSheet.Name = $ImportInfoCorrectionArchiveSheetName
    Format-ArchiveSheet $importInfoCorrectionSheet $ImportInfoCorrectionArchiveHeaders $ImportInfoCorrectionArchiveWidths

    $archiveWorkbook.SaveAs($ArchivePath, 51, $ArchiveWorkbookPassword) | Out-Null
    $archiveWorkbook.Close($true) | Out-Null
}

function Sync-ArchiveWorkbook($Excel, [string]$ArchivePath) {
    $archiveWorkbook = $Excel.Workbooks.Open($ArchivePath, 0, $false, 5, $ArchiveWorkbookPassword)
    try {
        $mainSheet = Get-OrCreateArchiveSheet $archiveWorkbook $SendArchiveSheetName "Архив исходных запросов"
        $commentColumn = Find-HeaderColumn $mainSheet "Комментарий МИЛ"
        $milColumn = Find-HeaderColumn $mainSheet "МИЛ"
        if ($null -eq $commentColumn) {
            if ($null -eq $milColumn) { $milColumn = 8 }
            $mainSheet.Columns.Item($milColumn).Insert() | Out-Null
        }
        Format-ArchiveSheet $mainSheet $SendArchiveHeaders $SendArchiveWidths

        $sendCorrectionSheet = Get-OrCreateArchiveSheet $archiveWorkbook $SendCorrectionArchiveSheetName "Архив корректировок"
        Format-ArchiveSheet $sendCorrectionSheet $SendCorrectionArchiveHeaders $SendCorrectionArchiveWidths

        $importInfoSheet = Get-OrCreateArchiveSheet $archiveWorkbook $ImportInfoArchiveSheetName
        Format-ArchiveSheet $importInfoSheet $ImportInfoArchiveHeaders $ImportInfoArchiveWidths

        $importInfoCorrectionSheet = Get-OrCreateArchiveSheet $archiveWorkbook $ImportInfoCorrectionArchiveSheetName
        Format-ArchiveSheet $importInfoCorrectionSheet $ImportInfoCorrectionArchiveHeaders $ImportInfoCorrectionArchiveWidths

        for ($index = $archiveWorkbook.Worksheets.Count; $index -ge 1; $index--) {
            $sheet = $archiveWorkbook.Worksheets.Item($index)
            if (@($SendArchiveSheetName, $SendCorrectionArchiveSheetName, $ImportInfoArchiveSheetName, $ImportInfoCorrectionArchiveSheetName) -contains $sheet.Name) {
                continue
            }
            if ($archiveWorkbook.Worksheets.Count -le 4) { break }
            if (-not [string]::IsNullOrWhiteSpace([string]$sheet.Cells.Item(1, 1).Value2)) {
                continue
            }
            $sheet.Delete()
        }

        $archiveWorkbook.Save()
    }
    finally {
        $archiveWorkbook.Close($true) | Out-Null
    }
}

function Ensure-ArchiveWorkbook($Excel, [string]$ArchivePath) {
    if ([string]::IsNullOrWhiteSpace($ArchivePath)) { return }
    if ($ArchivePath -match "://") { return }
    if (Test-Path -LiteralPath $ArchivePath) {
        Sync-ArchiveWorkbook $Excel $ArchivePath
    }
    else {
        Create-ArchiveWorkbook $Excel $ArchivePath
    }
}

function Set-ModuleSource($Component, [string]$Source) {
    $module = $Component.CodeModule
    if ($module.CountOfLines -gt 0) {
        $module.DeleteLines(1, $module.CountOfLines)
    }
    $normalized = $Source -replace "`r?`n", "`r`n"
    $module.InsertLines(1, $normalized)
}

function Get-OrCreateStandardModule($Project, [string]$ModuleName) {
    try {
        return $Project.VBComponents.Item($ModuleName)
    }
    catch {
        $component = $Project.VBComponents.Add(1)
        $component.Name = $ModuleName
        return $component
    }
}

function Apply-ModuleUpdates($Workbook) {
    $project = $Workbook.VBProject

    foreach ($moduleName in $ModuleSourceFiles.Keys) {
        $component = Get-OrCreateStandardModule $project $moduleName
        $source = [System.IO.File]::ReadAllText($ModuleSourceFiles[$moduleName], [System.Text.Encoding]::UTF8)
        Set-ModuleSource $component $source
    }

    $transforms = @(
        @{ Module = "A01CollectIznachalnieZakazi"; Old = 'Set ws = ThisWorkbook.Sheets("Рабочий")'; New = "Set ws = GetFsmRequestSheet()" },
        @{ Module = "A06TransferAlcoReportToRabochiy"; Old = 'Set wsRab = ThisWorkbook.Worksheets("Рабочий")'; New = "Set wsRab = GetFsmRequestSheet()" },
        @{ Module = "A07DeleteNeispolzovannieStroki"; Old = 'Set wsWork = wb.Worksheets("Рабочий")'; New = "Set wsWork = GetFsmRequestSheet()" },
        @{ Module = "A08CallHighlightIzhlishekFSM"; Old = 'Set ws = ThisWorkbook.Worksheets("Рабочий")'; New = "Set ws = GetFsmRequestSheet()" },
        @{ Module = "A09NaitiIVidelitIzmenenia"; Old = 'Set ws = ThisWorkbook.Worksheets("Рабочий")'; New = "Set ws = GetFsmRequestSheet()" },
        @{ Module = "A09NaitiIVidelitIzmenenia"; Old = 'Array("Заявление (КМ)", "Заявление (Новый)")'; New = 'Array("Заявление (КМ)", "Заявление (новое)")' },
        @{ Module = "A10AddDeystvieToZapros"; Old = 'Set wsRab = ThisWorkbook.Worksheets("Рабочий")'; New = "Set wsRab = GetFsmRequestSheet()" },
        @{ Module = "A11CenterAndWrapColumnsAtoM"; Old = 'Set ws = ThisWorkbook.Worksheets("Рабочий")'; New = "Set ws = GetFsmRequestSheet()" }
    )

    foreach ($transform in $transforms) {
        $component = $project.VBComponents.Item($transform.Module)
        $module = $component.CodeModule
        $source = $module.Lines(1, $module.CountOfLines)
        $source = $source.Replace($transform.Old, $transform.New)
        Set-ModuleSource $component $source
    }
}

function Ensure-SettingsSheet($Workbook) {
    $worksheet = $Workbook.Worksheets("Настройка")

    $downloadsRow = Find-SettingRow $worksheet $SettingDownloads
    if ($null -eq $downloadsRow) {
        $downloadsRow = 1
        $worksheet.Cells.Item($downloadsRow, 1).Value2 = $SettingDownloads
    }
    if ((Test-Path $DefaultDownloadsDir) -and -not (Test-Path ([string]$worksheet.Cells.Item($downloadsRow, 2).Value2))) {
        $worksheet.Cells.Item($downloadsRow, 2).Value2 = $DefaultDownloadsDir
    }

    $nomenclatureRow = Ensure-SettingRow $worksheet $SettingNomenclature 5
    $archiveRow = Ensure-SettingRow $worksheet $SettingArchive 6
    $passwordRow = Ensure-SettingRow $worksheet $SettingNomenclaturePassword 7
    $protectedPasswordRow = Ensure-SettingRow $worksheet $SettingProtectedSheetsPassword 8

    $worksheet.Cells.Item($nomenclatureRow, 2).Value2 = $DefaultNomenclatureWorkbook
    $currentArchivePath = [string]$worksheet.Cells.Item($archiveRow, 2).Value2
    if ([string]::IsNullOrWhiteSpace($currentArchivePath) -or $currentArchivePath.Contains("?") -or -not (Test-Path -LiteralPath $currentArchivePath)) {
        $worksheet.Cells.Item($archiveRow, 2).Value2 = $DefaultArchiveWorkbook
    }
    $worksheet.Cells.Item($passwordRow, 1).Value2 = $SettingNomenclaturePassword
    $worksheet.Cells.Item($passwordRow, 2).NumberFormat = "@"
    if ([string]::IsNullOrWhiteSpace([string]$worksheet.Cells.Item($passwordRow, 2).Value2)) {
        $worksheet.Cells.Item($passwordRow, 2).Value2 = "5623"
    }
    $worksheet.Cells.Item($protectedPasswordRow, 1).Value2 = $SettingProtectedSheetsPassword
    $worksheet.Cells.Item($protectedPasswordRow, 2).NumberFormat = "@"
    if ([string]::IsNullOrWhiteSpace([string]$worksheet.Cells.Item($protectedPasswordRow, 2).Value2)) {
        $worksheet.Cells.Item($protectedPasswordRow, 2).Value2 = "2356"
    }

    if ($null -eq (Find-SettingRow $worksheet $SettingsNote)) {
        $worksheet.Cells.Item(9, 1).Value2 = $SettingsNote
    }
}

if (-not (Test-Path -LiteralPath $OutputWorkbook)) {
    Copy-Item -LiteralPath $SourceWorkbook -Destination $OutputWorkbook -Force
    Start-Sleep -Seconds 2
}

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1

    Write-Output "Opening $OutputWorkbook"
    Write-Output "Exists: $(Test-Path -LiteralPath $OutputWorkbook)"
    $workbook = $excel.Workbooks.Open($OutputWorkbook, 0, $false)
    Close-OfficeActivationWizard
    $excel.Visible = $false
    for ($attempt = 1; $attempt -le 20; $attempt++) {
        $worksheetCount = Invoke-ComRetry { $workbook.Worksheets.Count } 5 500
        if ($worksheetCount -gt 0) { break }
        Start-Sleep -Milliseconds 500
    }
    Write-Output "Worksheets: $worksheetCount"
    if ($worksheetCount -le 0) {
        throw "Excel opened '$OutputWorkbook' without visible worksheets."
    }
    Start-Sleep -Seconds 2

    try {
        $fsmSheet = Invoke-ComRetry { $workbook.Worksheets($FsmSheetName) }
    }
    catch {
        try {
            $fsmSheet = Invoke-ComRetry { $workbook.Worksheets("Рабочий") }
            $fsmSheet.Name = $FsmSheetName
        }
        catch {
            $fsmSheet = Invoke-ComRetry { $workbook.Worksheets(1) }
            if ($fsmSheet.Name -ne $FsmSheetName) {
                $fsmSheet.Name = $FsmSheetName
            }
        }
    }

    Ensure-SettingsSheet $workbook
    $archiveRow = Find-SettingRow $workbook.Worksheets("Настройка") $SettingArchive
    if ($null -ne $archiveRow) {
        Ensure-ArchiveWorkbook $excel ([string]$workbook.Worksheets("Настройка").Cells.Item($archiveRow, 2).Value2)
    }

    $nomenclatureSheet = Get-OrCreateSheet $workbook $NomenclatureSheetName $fsmSheet
    Format-StagingSheet $excel $nomenclatureSheet $NomenclatureHeaders $NomenclatureWidths
    Configure-StagingButtons $nomenclatureSheet $NomenclatureHeaders.Count @(
        [pscustomobject]@{ Caption = $PrepareCaption; Action = "PrepareNomenclatureRequest" },
        [pscustomobject]@{ Caption = $SendCaption; Action = "SendNomenclatureRequest" },
        [pscustomobject]@{ Caption = $CorrectionCaption; Action = "SendNomenclatureCorrectionRequest" }
    )

    $importInfoSheet = Get-OrCreateSheet $workbook $ImportInfoSheetName $nomenclatureSheet
    Format-StagingSheet $excel $importInfoSheet $ImportInfoHeaders $ImportInfoWidths
    Configure-StagingButtons $importInfoSheet $ImportInfoHeaders.Count @(
        [pscustomobject]@{ Caption = $PrepareCaption; Action = "PrepareImportInfoRequest" },
        [pscustomobject]@{ Caption = $SendCaption; Action = "SendImportInfoRequest" },
        [pscustomobject]@{ Caption = $CorrectionCaption; Action = "SendImportInfoCorrectionRequest" }
    )

    $baselineSheet = Get-OrCreateSheet $workbook $ImportInfoBaselineSheetName $importInfoSheet
    $baselineSheet.Cells.Clear() | Out-Null
    for ($i = 0; $i -lt $ImportInfoBaselineHeaders.Count; $i++) {
        $baselineSheet.Cells.Item(1, $i + 1).Value2 = $ImportInfoBaselineHeaders[$i]
        $baselineSheet.Columns.Item($i + 1).ColumnWidth = $ImportInfoBaselineWidths[$i]
    }
    $baselineHeaderRange = $baselineSheet.Range($baselineSheet.Cells.Item(1, 1), $baselineSheet.Cells.Item(1, $ImportInfoBaselineHeaders.Count))
    $baselineHeaderRange.Interior.Color = 65535
    $baselineHeaderRange.Font.Bold = $true
    $baselineHeaderRange.HorizontalAlignment = -4108
    $baselineHeaderRange.VerticalAlignment = -4108
    $baselineHeaderRange.WrapText = $true
    $baselineSheet.Rows.Item(1).RowHeight = 30
    $baselineSheet.Visible = 2

    Apply-ModuleUpdates $workbook
    $excel.Run("'$($workbook.Name)'!EnsureInteractiveSheetProtection")

    $workbook.Save()
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($true) | Out-Null } catch {}
        Release-ComObject $workbook
    }
    if ($null -ne $excel) {
        try { $excel.Quit() | Out-Null } catch {}
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Output "Built $OutputWorkbook"



