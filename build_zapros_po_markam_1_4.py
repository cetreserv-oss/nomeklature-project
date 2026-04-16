from __future__ import annotations

import shutil
from pathlib import Path

import pythoncom
import win32com.client as win32


ROOT_DIR = Path(r"E:\Projects\nomenklature project")
REFERENCE_DIR = ROOT_DIR / "reference"
VBA_SOURCE_DIR = ROOT_DIR / "vba_marki_1_4"
DEFAULT_DOWNLOADS_DIR = Path.home() / "Downloads"
DEFAULT_NOMENCLATURE_PASSWORD = "5623"
DEFAULT_NOMENCLATURE_PROTECTED_SHEETS_PASSWORD = "2356"

SOURCE_WORKBOOK = REFERENCE_DIR / "Запрос по маркам 1.3.xlsm"
OUTPUT_WORKBOOK = REFERENCE_DIR / "Запрос по маркам 1.4.xlsm"
DEFAULT_NOMENCLATURE_WORKBOOK = REFERENCE_DIR / "Номенклатура New 2.xlsx"
DEFAULT_ARCHIVE_WORKBOOK = REFERENCE_DIR / "Архив исходных запросов.xlsx"

FSM_SHEET_NAME = "Отправить запрос по ФСМ"
NOMENCLATURE_SHEET_NAME = "Отправка марок (номенклатура)"
IMPORT_INFO_SHEET_NAME = "Сведения о ввозе (номенклатура)"
IMPORT_INFO_BASELINE_SHEET_NAME = "__Ввоз_База"
NOMENCLATURE_HEADERS = [
    "Номер заказа",
    "Поставщик",
    "Код УТ",
    "Номенклатура",
    "Заявление на выдачу ФСМ",
    "Кол-во",
    "Оклейщик",
    "МИЛ",
    "Комментарий МИЛ",
]
IMPORT_INFO_HEADERS = [
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
    "Комментарий МИЛ",
]
IMPORT_INFO_BASELINE_HEADERS = [
    "Номер заказа",
    "Поставщик",
    "Код УТ",
    "Номенклатура",
    "Заявление на выдачу ФСМ",
    "Количество",
    "Оклейщик",
    "Накладная отправки",
    "Дата отправки",
    "МИЛ",
]
SEND_ARCHIVE_HEADERS = [
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
    "Подтверждение повторной отправки",
]
SEND_ARCHIVE_SHEET_NAME = "Исх. запросы (отпр ФСМ)"
SEND_CORRECTION_ARCHIVE_HEADERS = [
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
]
SEND_CORRECTION_ARCHIVE_SHEET_NAME = "Корр. (отпр ФСМ)"
IMPORT_INFO_ARCHIVE_HEADERS = IMPORT_INFO_HEADERS + [
    "Дата внесения строки",
    "Подтверждение повторной отправки",
    "Подтверждение несоответствия количества",
]
IMPORT_INFO_ARCHIVE_SHEET_NAME = "Исх. запросы (ввоз)"
IMPORT_INFO_CORRECTION_ARCHIVE_HEADERS = IMPORT_INFO_HEADERS + [
    "Дата внесения строки",
    "Подтверждение несоответствия количества",
]
IMPORT_INFO_CORRECTION_ARCHIVE_SHEET_NAME = "Корр. (ввоз)"
ARCHIVE_WORKBOOK_PASSWORD = "7777"

SETTING_DOWNLOADS = "Папка загрузки"
SETTING_NOMENCLATURE = "Номенклатура"
SETTING_ARCHIVE = "Архив исходных запросов"
SETTING_NOMENCLATURE_PASSWORD = "Пароль номенклатуры"
SETTING_NOMENCLATURE_PROTECTED_SHEETS_PASSWORD = "Пароль защищенных вкладок номенклатуры"
SETTINGS_NOTE = "Перед нажатием кнопок нужно сделать выгрузку из Алкоотчета и актуализировать Логос"

NOMENCLATURE_PREPARE_BUTTON_CAPTION = "Подготовить строки к внесению в номенклатуру"
NOMENCLATURE_SEND_BUTTON_CAPTION = "Внести строчки в номенклатуру"
NOMENCLATURE_CORRECTION_BUTTON_CAPTION = "Внести корректировку в номенклатуру"

BUTTON_WIDTH = 250
BUTTON_HEIGHT = 54
BUTTON_VERTICAL_GAP = 6
BUTTON_LEFT_OFFSET = 12

NOMENCLATURE_COLUMN_WIDTHS = (18, 24, 16, 36, 24, 12, 24, 18, 24)
IMPORT_INFO_COLUMN_WIDTHS = (18, 24, 16, 36, 24, 12, 24, 22, 16, 18, 18, 16, 18, 24)
IMPORT_INFO_BASELINE_COLUMN_WIDTHS = (18, 24, 16, 36, 24, 12, 24, 22, 16, 18)
SEND_ARCHIVE_COLUMN_WIDTHS = (18, 16, 36, 24, 12, 24, 24, 24, 18, 18, 24)
SEND_CORRECTION_ARCHIVE_COLUMN_WIDTHS = (18, 16, 36, 24, 12, 24, 24, 24, 18, 18)
IMPORT_INFO_ARCHIVE_COLUMN_WIDTHS = (18, 24, 16, 36, 24, 12, 24, 22, 16, 18, 18, 16, 18, 24, 18, 24, 32)
IMPORT_INFO_CORRECTION_ARCHIVE_COLUMN_WIDTHS = (18, 24, 16, 36, 24, 12, 24, 22, 16, 18, 18, 16, 18, 24, 18, 32)

MODULE_SOURCE_FILES = {
    "A00MainModule": VBA_SOURCE_DIR / "A00MainModule.bas",
    "A02FormatZakaz": VBA_SOURCE_DIR / "A02FormatZakaz.bas",
    "A03UpdateDataSheets": VBA_SOURCE_DIR / "A03UpdateDataSheets.bas",
    "A04ImportKontrolMarokData": VBA_SOURCE_DIR / "A04ImportKontrolMarokData.bas",
    "A05ObrabotkaAlkoReport": VBA_SOURCE_DIR / "A05ObrabotkaAlkoReport.bas",
    "Z01OtpravkaZaprosa": VBA_SOURCE_DIR / "Z01OtpravkaZaprosa.bas",
    "Z02NomenclatureRequest": VBA_SOURCE_DIR / "Z02NomenclatureRequest.bas",
}

MODULE_TRANSFORMS = {
    "A01CollectIznachalnieZakazi": [
        ('Set ws = ThisWorkbook.Sheets("Рабочий")', "Set ws = GetFsmRequestSheet()"),
    ],
    "A06TransferAlcoReportToRabochiy": [
        ('Set wsRab = ThisWorkbook.Worksheets("Рабочий")', "Set wsRab = GetFsmRequestSheet()"),
    ],
    "A07DeleteNeispolzovannieStroki": [
        ('Set wsWork = wb.Worksheets("Рабочий")', "Set wsWork = GetFsmRequestSheet()"),
    ],
    "A08CallHighlightIzhlishekFSM": [
        ('Set ws = ThisWorkbook.Worksheets("Рабочий")', "Set ws = GetFsmRequestSheet()"),
    ],
    "A09NaitiIVidelitIzmenenia": [
        ('Set ws = ThisWorkbook.Worksheets("Рабочий")', "Set ws = GetFsmRequestSheet()"),
        ('Array("Заявление (КМ)", "Заявление (Новый)")', 'Array("Заявление (КМ)", "Заявление (новое)")'),
    ],
    "A10AddDeystvieToZapros": [
        ('Set wsRab = ThisWorkbook.Worksheets("Рабочий")', "Set wsRab = GetFsmRequestSheet()"),
    ],
    "A11CenterAndWrapColumnsAtoM": [
        ('Set ws = ThisWorkbook.Worksheets("Рабочий")', "Set ws = GetFsmRequestSheet()"),
    ],
}


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8").replace("\n", "\r\n")


def get_module_source(component) -> str:
    code_module = component.CodeModule
    if code_module.CountOfLines == 0:
        return ""
    return code_module.Lines(1, code_module.CountOfLines)


def set_module_source(component, source: str) -> None:
    code_module = component.CodeModule
    if code_module.CountOfLines:
        code_module.DeleteLines(1, code_module.CountOfLines)
    code_module.InsertLines(1, source)


def get_or_create_standard_module(project, module_name: str):
    try:
        return project.VBComponents(module_name)
    except Exception:
        component = project.VBComponents.Add(1)
        component.Name = module_name
        return component


def apply_module_updates(workbook) -> None:
    project = workbook.VBProject

    for module_name, path in MODULE_SOURCE_FILES.items():
        component = get_or_create_standard_module(project, module_name)
        set_module_source(component, read_text(path))

    for module_name, replacements in MODULE_TRANSFORMS.items():
        component = project.VBComponents(module_name)
        source = get_module_source(component)
        for old, new in replacements:
            source = source.replace(old, new)
        set_module_source(component, source)


def ensure_output_copy() -> None:
    if OUTPUT_WORKBOOK.exists():
        OUTPUT_WORKBOOK.unlink()
    shutil.copy2(SOURCE_WORKBOOK, OUTPUT_WORKBOOK)


def ensure_fsm_sheet(workbook):
    worksheet = workbook.Worksheets("Рабочий")
    worksheet.Name = FSM_SHEET_NAME
    return worksheet


def find_setting_row(worksheet, label: str):
    for row in range(1, 50):
        value = str(worksheet.Cells(row, 1).Value or "").strip()
        if value.lower() == label.lower():
            return row
    return None


def ensure_setting_row(worksheet, label: str, preferred_row: int) -> int:
    existing_row = find_setting_row(worksheet, label)
    if existing_row is not None:
        return existing_row

    worksheet.Rows(preferred_row).Insert()
    worksheet.Cells(preferred_row, 1).Value = label
    return preferred_row


def is_probably_local_path(path_value: str) -> bool:
    path_value = path_value.strip()
    return bool(path_value) and "://" not in path_value


def build_default_archive_path(nomenclature_path: str) -> Path:
    if is_probably_local_path(nomenclature_path):
        try:
            return Path(nomenclature_path).expanduser().resolve().parent / DEFAULT_ARCHIVE_WORKBOOK.name
        except OSError:
            pass
    return DEFAULT_ARCHIVE_WORKBOOK


def find_header_column(worksheet, header: str, max_columns: int = 50):
    for column_index in range(1, max_columns + 1):
        value = str(worksheet.Cells(1, column_index).Value or "").strip()
        if value.lower() == header.lower():
            return column_index
    return None


def get_or_create_sheet(workbook, sheet_name: str, *, after_sheet=None):
    try:
        return workbook.Worksheets(sheet_name)
    except Exception:
        if after_sheet is None:
            return workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        return workbook.Worksheets.Add(After=after_sheet)


def format_staging_sheet(worksheet, headers: list[str], widths: tuple[int, ...]) -> None:
    worksheet.Cells.Clear()

    for column_index, header in enumerate(headers, start=1):
        worksheet.Cells(1, column_index).Value = header

    header_range = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(1, len(headers)))
    header_range.Interior.Color = 65535
    header_range.Font.Bold = True
    header_range.HorizontalAlignment = -4108
    header_range.VerticalAlignment = -4108
    header_range.WrapText = True
    worksheet.Rows(1).RowHeight = 30

    for column_index, width in enumerate(widths, start=1):
        worksheet.Columns(column_index).ColumnWidth = width

    worksheet.Activate()
    active_window = worksheet.Application.ActiveWindow
    active_window.SplitColumn = 0
    active_window.SplitRow = 1
    active_window.FreezePanes = False
    active_window.FreezePanes = True


def configure_staging_buttons(worksheet, header_count: int, button_specs: list[tuple[str, str, str]]) -> None:
    while worksheet.Buttons().Count > 0:
        worksheet.Buttons(1).Delete()

    button_left = worksheet.Cells(2, header_count + 1).Left + BUTTON_LEFT_OFFSET
    button_top = worksheet.Cells(2, header_count + 1).Top

    for index, (_, caption, action) in enumerate(button_specs):
        button = worksheet.Buttons().Add(
            button_left,
            button_top + index * (BUTTON_HEIGHT + BUTTON_VERTICAL_GAP),
            BUTTON_WIDTH,
            BUTTON_HEIGHT,
        )
        button.Caption = caption
        button.OnAction = action
        button.Locked = False

    for shape in worksheet.Shapes:
        shape.Placement = 3


def format_archive_sheet(worksheet, headers: list[str], widths: tuple[int, ...]) -> None:
    for column_index, header in enumerate(headers, start=1):
        worksheet.Cells(1, column_index).Value = header

    header_range = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(1, len(headers)))
    header_range.Interior.Color = 65535
    header_range.Font.Bold = True
    header_range.HorizontalAlignment = -4108
    header_range.VerticalAlignment = -4108
    header_range.WrapText = True
    worksheet.Rows(1).RowHeight = 30

    for column_index, width in enumerate(widths, start=1):
        worksheet.Columns(column_index).ColumnWidth = width

    worksheet.Activate()
    active_window = worksheet.Application.ActiveWindow
    active_window.SplitColumn = 0
    active_window.SplitRow = 1
    active_window.FreezePanes = False
    active_window.FreezePanes = True


def create_archive_workbook(excel, archive_path: Path) -> None:
    archive_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = excel.Workbooks.Add()
    while workbook.Worksheets.Count > 1:
        workbook.Worksheets(workbook.Worksheets.Count).Delete()

    main_sheet = workbook.Worksheets(1)
    main_sheet.Name = SEND_ARCHIVE_SHEET_NAME
    format_archive_sheet(main_sheet, SEND_ARCHIVE_HEADERS, SEND_ARCHIVE_COLUMN_WIDTHS)

    send_correction_sheet = workbook.Worksheets.Add(After=main_sheet)
    send_correction_sheet.Name = SEND_CORRECTION_ARCHIVE_SHEET_NAME
    format_archive_sheet(
        send_correction_sheet,
        SEND_CORRECTION_ARCHIVE_HEADERS,
        SEND_CORRECTION_ARCHIVE_COLUMN_WIDTHS,
    )

    import_info_sheet = workbook.Worksheets.Add(After=send_correction_sheet)
    import_info_sheet.Name = IMPORT_INFO_ARCHIVE_SHEET_NAME
    format_archive_sheet(
        import_info_sheet,
        IMPORT_INFO_ARCHIVE_HEADERS,
        IMPORT_INFO_ARCHIVE_COLUMN_WIDTHS,
    )

    import_info_correction_sheet = workbook.Worksheets.Add(After=import_info_sheet)
    import_info_correction_sheet.Name = IMPORT_INFO_CORRECTION_ARCHIVE_SHEET_NAME
    format_archive_sheet(
        import_info_correction_sheet,
        IMPORT_INFO_CORRECTION_ARCHIVE_HEADERS,
        IMPORT_INFO_CORRECTION_ARCHIVE_COLUMN_WIDTHS,
    )

    workbook.SaveAs(str(archive_path), FileFormat=51, Password=ARCHIVE_WORKBOOK_PASSWORD)
    workbook.Close(SaveChanges=True)


def sync_archive_workbook(excel, archive_path: Path) -> None:
    workbook = excel.Workbooks.Open(str(archive_path), Password=ARCHIVE_WORKBOOK_PASSWORD, ReadOnly=False)
    try:
        try:
            main_sheet = workbook.Worksheets(SEND_ARCHIVE_SHEET_NAME)
        except Exception:
            try:
                main_sheet = workbook.Worksheets("Архив исходных запросов")
                main_sheet.Name = SEND_ARCHIVE_SHEET_NAME
            except Exception:
                main_sheet = get_or_create_sheet(workbook, SEND_ARCHIVE_SHEET_NAME)

        comment_column = find_header_column(main_sheet, "Комментарий МИЛ")
        mil_column = find_header_column(main_sheet, "МИЛ")
        if comment_column is None:
            main_sheet.Columns(mil_column or 8).Insert()
        format_archive_sheet(main_sheet, SEND_ARCHIVE_HEADERS, SEND_ARCHIVE_COLUMN_WIDTHS)

        try:
            send_correction_sheet = workbook.Worksheets(SEND_CORRECTION_ARCHIVE_SHEET_NAME)
        except Exception:
            try:
                send_correction_sheet = workbook.Worksheets("Архив корректировок")
                send_correction_sheet.Name = SEND_CORRECTION_ARCHIVE_SHEET_NAME
            except Exception:
                send_correction_sheet = get_or_create_sheet(
                    workbook,
                    SEND_CORRECTION_ARCHIVE_SHEET_NAME,
                    after_sheet=main_sheet,
                )
        format_archive_sheet(
            send_correction_sheet,
            SEND_CORRECTION_ARCHIVE_HEADERS,
            SEND_CORRECTION_ARCHIVE_COLUMN_WIDTHS,
        )

        import_info_sheet = get_or_create_sheet(
            workbook,
            IMPORT_INFO_ARCHIVE_SHEET_NAME,
            after_sheet=send_correction_sheet,
        )
        format_archive_sheet(
            import_info_sheet,
            IMPORT_INFO_ARCHIVE_HEADERS,
            IMPORT_INFO_ARCHIVE_COLUMN_WIDTHS,
        )

        import_info_correction_sheet = get_or_create_sheet(
            workbook,
            IMPORT_INFO_CORRECTION_ARCHIVE_SHEET_NAME,
            after_sheet=import_info_sheet,
        )
        format_archive_sheet(
            import_info_correction_sheet,
            IMPORT_INFO_CORRECTION_ARCHIVE_HEADERS,
            IMPORT_INFO_CORRECTION_ARCHIVE_COLUMN_WIDTHS,
        )

        for index in range(workbook.Worksheets.Count, 0, -1):
            worksheet = workbook.Worksheets(index)
            if worksheet.Name in {
                SEND_ARCHIVE_SHEET_NAME,
                SEND_CORRECTION_ARCHIVE_SHEET_NAME,
                IMPORT_INFO_ARCHIVE_SHEET_NAME,
                IMPORT_INFO_CORRECTION_ARCHIVE_SHEET_NAME,
            }:
                continue
            if workbook.Worksheets.Count <= 4:
                break
            if str(worksheet.Cells(1, 1).Value or "").strip():
                continue
            worksheet.Delete()

        workbook.Save()
    finally:
        workbook.Close(SaveChanges=True)


def ensure_archive_workbook(excel, archive_path: str) -> None:
    if not is_probably_local_path(archive_path):
        return

    archive_file = Path(archive_path)
    if archive_file.exists():
        sync_archive_workbook(excel, archive_file)
        return

    create_archive_workbook(excel, archive_file)


def ensure_settings_sheet(workbook, excel) -> None:
    worksheet = workbook.Worksheets("Настройка")
    default_nomenclature_path = str(DEFAULT_NOMENCLATURE_WORKBOOK)

    downloads_row = find_setting_row(worksheet, SETTING_DOWNLOADS)
    if downloads_row is None:
        downloads_row = 1
        worksheet.Cells(downloads_row, 1).Value = SETTING_DOWNLOADS

    current_downloads_path = str(worksheet.Cells(downloads_row, 2).Value or "").strip()
    current_downloads_path_exists = False
    if current_downloads_path:
        try:
            current_downloads_path_exists = Path(current_downloads_path).exists()
        except OSError:
            current_downloads_path_exists = False

    if DEFAULT_DOWNLOADS_DIR.exists() and not current_downloads_path_exists:
        worksheet.Cells(downloads_row, 2).Value = str(DEFAULT_DOWNLOADS_DIR)

    nomenclature_row = ensure_setting_row(worksheet, SETTING_NOMENCLATURE, 5)
    archive_row = ensure_setting_row(worksheet, SETTING_ARCHIVE, 6)
    password_row = ensure_setting_row(worksheet, SETTING_NOMENCLATURE_PASSWORD, 7)
    protected_sheets_password_row = ensure_setting_row(
        worksheet,
        SETTING_NOMENCLATURE_PROTECTED_SHEETS_PASSWORD,
        8,
    )

    current_nomenclature_path = str(worksheet.Cells(nomenclature_row, 2).Value or "").strip()
    if not current_nomenclature_path or Path(current_nomenclature_path).name.lower() == "номенклатура new.xlsx":
        worksheet.Cells(nomenclature_row, 2).Value = default_nomenclature_path
        current_nomenclature_path = default_nomenclature_path

    current_archive_path = str(worksheet.Cells(archive_row, 2).Value or "").strip()
    current_archive_path_exists = False
    if current_archive_path and "?" not in current_archive_path:
        try:
            current_archive_path_exists = Path(current_archive_path).exists()
        except OSError:
            current_archive_path_exists = False
    if not current_archive_path or "?" in current_archive_path or not current_archive_path_exists:
        current_archive_path = str(build_default_archive_path(current_nomenclature_path))
        worksheet.Cells(archive_row, 2).Value = current_archive_path

    worksheet.Cells(password_row, 1).Value = SETTING_NOMENCLATURE_PASSWORD
    worksheet.Cells(password_row, 2).NumberFormat = "@"
    current_password = str(worksheet.Cells(password_row, 2).Value or "").strip()
    if not current_password:
        worksheet.Cells(password_row, 2).Value = DEFAULT_NOMENCLATURE_PASSWORD

    worksheet.Cells(protected_sheets_password_row, 1).Value = SETTING_NOMENCLATURE_PROTECTED_SHEETS_PASSWORD
    worksheet.Cells(protected_sheets_password_row, 2).NumberFormat = "@"
    current_protected_sheets_password = str(
        worksheet.Cells(protected_sheets_password_row, 2).Value or ""
    ).strip()
    if not current_protected_sheets_password:
        worksheet.Cells(protected_sheets_password_row, 2).Value = DEFAULT_NOMENCLATURE_PROTECTED_SHEETS_PASSWORD

    note_row = find_setting_row(worksheet, SETTINGS_NOTE)
    if note_row is None:
        note_row = 9
        worksheet.Cells(note_row, 1).Value = SETTINGS_NOTE

    ensure_archive_workbook(excel, current_archive_path)


def ensure_nomenclature_sheet(workbook, fsm_sheet) -> None:
    try:
        worksheet = workbook.Worksheets(NOMENCLATURE_SHEET_NAME)
    except Exception:
        worksheet = workbook.Worksheets.Add(After=fsm_sheet)
        worksheet.Name = NOMENCLATURE_SHEET_NAME

    format_staging_sheet(worksheet, NOMENCLATURE_HEADERS, NOMENCLATURE_COLUMN_WIDTHS)
    button_specs = [
        ("btnPrepareNomenclatureRequest", NOMENCLATURE_PREPARE_BUTTON_CAPTION, "PrepareNomenclatureRequest"),
        ("btnSendNomenclatureRequest", NOMENCLATURE_SEND_BUTTON_CAPTION, "SendNomenclatureRequest"),
        (
            "btnSendNomenclatureCorrectionRequest",
            NOMENCLATURE_CORRECTION_BUTTON_CAPTION,
            "SendNomenclatureCorrectionRequest",
        ),
    ]
    configure_staging_buttons(worksheet, len(NOMENCLATURE_HEADERS), button_specs)
    return worksheet


def ensure_import_info_sheet(workbook, after_sheet) -> None:
    try:
        worksheet = workbook.Worksheets(IMPORT_INFO_SHEET_NAME)
    except Exception:
        worksheet = workbook.Worksheets.Add(After=after_sheet)
        worksheet.Name = IMPORT_INFO_SHEET_NAME

    format_staging_sheet(worksheet, IMPORT_INFO_HEADERS, IMPORT_INFO_COLUMN_WIDTHS)
    button_specs = [
        ("btnPrepareImportInfoRequest", NOMENCLATURE_PREPARE_BUTTON_CAPTION, "PrepareImportInfoRequest"),
        ("btnSendImportInfoRequest", NOMENCLATURE_SEND_BUTTON_CAPTION, "SendImportInfoRequest"),
        (
            "btnSendImportInfoCorrectionRequest",
            NOMENCLATURE_CORRECTION_BUTTON_CAPTION,
            "SendImportInfoCorrectionRequest",
        ),
    ]
    configure_staging_buttons(worksheet, len(IMPORT_INFO_HEADERS), button_specs)
    return worksheet


def ensure_import_info_baseline_sheet(workbook, after_sheet) -> None:
    try:
        worksheet = workbook.Worksheets(IMPORT_INFO_BASELINE_SHEET_NAME)
    except Exception:
        worksheet = workbook.Worksheets.Add(After=after_sheet)
        worksheet.Name = IMPORT_INFO_BASELINE_SHEET_NAME

    worksheet.Cells.Clear()
    for column_index, header in enumerate(IMPORT_INFO_BASELINE_HEADERS, start=1):
        worksheet.Cells(1, column_index).Value = header

    header_range = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(1, len(IMPORT_INFO_BASELINE_HEADERS)))
    header_range.Interior.Color = 65535
    header_range.Font.Bold = True
    header_range.HorizontalAlignment = -4108
    header_range.VerticalAlignment = -4108
    header_range.WrapText = True
    worksheet.Rows(1).RowHeight = 30

    for column_index, width in enumerate(IMPORT_INFO_BASELINE_COLUMN_WIDTHS, start=1):
        worksheet.Columns(column_index).ColumnWidth = width

    worksheet.Visible = 2


def ensure_button_actions(fsm_sheet) -> None:
    for shape in fsm_sheet.Shapes:
        action = ""
        text = ""

        try:
            action = shape.OnAction or ""
        except Exception:
            action = ""

        try:
            text = shape.TextFrame.Characters().Text or ""
        except Exception:
            text = ""

        if "Main" in action or "Подтянуть данные" in text:
            shape.OnAction = "Main"
        elif "OtpravkaZaprosa" in action or "Отправить запрос" in text:
            shape.OnAction = "OtpravkaZaprosa"


def run_macro(excel, workbook, macro_name: str) -> None:
    excel.Run(f"'{workbook.Name}'!{macro_name}")


def build_workbook() -> Path:
    ensure_output_copy()

    pythoncom.CoInitialize()
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None

    try:
        workbook = excel.Workbooks.Open(str(OUTPUT_WORKBOOK))

        fsm_sheet = ensure_fsm_sheet(workbook)
        ensure_settings_sheet(workbook, excel)
        nomenclature_sheet = ensure_nomenclature_sheet(workbook, fsm_sheet)
        import_info_sheet = ensure_import_info_sheet(workbook, nomenclature_sheet)
        ensure_import_info_baseline_sheet(workbook, import_info_sheet)
        ensure_button_actions(fsm_sheet)
        apply_module_updates(workbook)
        run_macro(excel, workbook, "EnsureInteractiveSheetProtection")

        workbook.Save()
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        excel.Quit()
        pythoncom.CoUninitialize()

    return OUTPUT_WORKBOOK


if __name__ == "__main__":
    created_path = build_workbook()
    print(created_path)
