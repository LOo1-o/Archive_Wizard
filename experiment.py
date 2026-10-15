# -*- coding: utf-8 -*-
"""
Улучшенная версия скрипта обработки Excel (openpyxl + COM).
Основные изменения:
- Единая функция calculate_table_dimensions (возвращает width и height).
- Явное выключение FitToPages перед установкой Zoom (openpyxl и COM).
- Последовательность: openpyxl -> save -> COM (гарантированное применение Excel).
- Расширенное логирование PageSetup до/после (openpyxl и COM).
- Убраны дубликаты функций, приведены к одному стилю.
- Добавлен безопасный clamp для Zoom и диагностические логи.
"""
from pathlib import Path
import zipfile
import logging
import os
import time
import tkinter as tk
import traceback
from tkinter import filedialog, messagebox, ttk
import re

from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties

# Excel / COM
import pythoncom
import win32com.client as win32

# openpyxl
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.page import PageMargins
from openpyxl.styles import PatternFill

try:
    import rarfile
except ImportError:
    rarfile = None

# Настройка логгера
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

# Константы
#PAGE_WIDTH_A4_PORTRAIT = 7.5  # Ширина A4 в дюймах (портрет)
PAGE_HEIGHT_A4_PORTRAIT = 11.69  # Высота A4 в дюймах (портрет)

PAGE_WIDTH_A4_PORTRAIT = 8.27  # дюймы (A4)
#PAGE_HEIGHT_A4_PORTRAIT = 11.69  # дюймы (A4)
EXCEL_POINT_TO_INCH = 1 / 72.0
CALIBRATION_FACTOR = 1.05
#CALIBRATION_FACTOR = 1.02
MAX_FILES_TO_PROCESS = 1000


# -----------------------
# Утилиты для архивов
# -----------------------
def extract_archive_with_correct_encoding(archive_path, output_folder=None):
    archive_path = Path(archive_path)
    if output_folder is None:
        output_folder = archive_path.parent
    folder_name = archive_path.stem
    encodings_to_try = ['cp866', 'cp1251', 'utf-8', 'iso-8859-1', 'koi8-r']
    extract_to = output_folder / folder_name
    try:
        extract_to.mkdir(parents=True, exist_ok=True)
        logger.info(f"Распаковка архива: {archive_path.name} → {folder_name}/")
        extracted_files = []
        if archive_path.suffix.lower() == ".zip":
            for encoding in encodings_to_try:
                try:
                    with zipfile.ZipFile(archive_path, 'r') as zip_ref:
                        file_list = zip_ref.namelist()
                        for file_name in file_list:
                            try:
                                decoded_name = file_name.encode('cp437').decode(encoding)
                            except Exception:
                                decoded_name = file_name
                            target_path = extract_to / decoded_name
                            target_path.parent.mkdir(parents=True, exist_ok=True)
                            with zip_ref.open(file_name) as source, open(target_path, 'wb') as target:
                                target.write(source.read())
                            extracted_files.append(target_path)
                        logger.info(f"ZIP архив распакован с кодировкой {encoding} ({len(extracted_files)} файлов)")
                        break
                except Exception as e:
                    logger.debug(f"Кодировка {encoding} не подошла: {e}")
                    continue
        elif archive_path.suffix.lower() == ".rar" and rarfile:
            with rarfile.RarFile(archive_path, 'r') as rar_ref:
                file_list = rar_ref.namelist()
                rar_ref.extractall(extract_to)
                for file_name in file_list:
                    file_path = extract_to / file_name
                    if file_path.exists():
                        extracted_files.append(file_path)
                logger.info(f"RAR архив распакован ({len(extracted_files)} файлов)")
        else:
            logger.warning(f"Формат {archive_path.suffix} не поддерживается")
            return None
        # рекурсивно распаковываем вложенные архивы
        nested_archives = []
        for item in extracted_files:
            if item.is_file() and item.suffix.lower() in (".zip", ".rar"):
                nested_archives.append(item)
        for nested_archive in nested_archives:
            logger.info(f"Найден вложенный архив: {nested_archive.name}")
            extract_archive_with_correct_encoding(nested_archive, extract_to)
        return extract_to
    except Exception as e:
        logger.error(f"Ошибка распаковки архива {archive_path.name}: {e}")
        return None


def process_archives_in_folder(folder_path, delete_originals=False):
    folder_path = Path(folder_path)
    processed_folders = []
    root_archives = []
    for item in folder_path.iterdir():
        if item.is_file() and item.suffix.lower() in (".zip", ".rar"):
            root_archives.append(item)
    logger.info(f"Найдено архивов в '{folder_path.name}': {len(root_archives)}")
    for archive in root_archives:
        try:
            extracted_folder = extract_archive_with_correct_encoding(archive)
            if extracted_folder:
                processed_folders.append(extracted_folder)
                if delete_originals:
                    try:
                        archive.unlink()
                        logger.info(f"Удален исходный архив: {archive.name}")
                    except Exception as e:
                        logger.warning(f"Не удалось удалить архив {archive.name}: {e}")
                else:
                    logger.info(f"Оригинальный архив сохранен: {archive.name}")
        except Exception as e:
            logger.error(f"Ошибка обработки архива {archive.name}: {e}")
            continue
    for subfolder in folder_path.iterdir():
        if subfolder.is_dir():
            if any(subfolder.name == f.stem for f in root_archives):
                continue
            logger.debug(f"Проверка подпапки: {subfolder.name}")
            sub_folders = process_archives_in_folder(subfolder, delete_originals)
            processed_folders.extend(sub_folders)
    return processed_folders


def clean_empty_folders(folder_path):
    try:
        deleted_count = 0
        for item in folder_path.iterdir():
            if item.is_dir():
                deleted_count += clean_empty_folders(item)
                try:
                    if not any(item.iterdir()):
                        item.rmdir()
                        deleted_count += 1
                        logger.debug(f"Удалена пустая папка: {item.name}")
                except:
                    pass
        return deleted_count
    except Exception as e:
        logger.warning(f"Ошибка при очистке пустых папок: {e}")
        return 0


# -----------------------
# Вспомогательные для Excel
# -----------------------
def cleanup_old_print_settings(wb):
    try:
        # Удаляет определения печати _xlnm.Print_Area и _xlnm.Print_Titles
        try:
            dn = wb.defined_names
            if hasattr(dn, 'definedName'):
                # старый стиль
                items = list(dn.definedName)
            else:
                items = list(dn)
            names_to_remove = []
            for name in items:
                n = getattr(name, 'name', None)
                if n in ['_xlnm.Print_Area', '_xlnm.Print_Titles']:
                    names_to_remove.append(name)
            # попытка удалить
            for name in names_to_remove:
                try:
                    if hasattr(wb.defined_names, 'definedName'):
                        wb.defined_names.definedName.remove(name)
                    else:
                        del wb.defined_names[name.name]
                except Exception:
                    pass
            logger.info(f"Удалено старых настроек печати: {len(names_to_remove)}")
        except Exception:
            pass
        return True
    except Exception as e:
        logger.warning(f"Не удалось очистить старые настройки: {e}")
        return False


def reset_page_setup_openpyxl(ws):
    """Сбрасывает настройки печати для листа openpyxl (без удаления содержимого)"""
    try:
        ws.print_area = None
        ws.print_title_rows = None
        ws.print_title_cols = None
        # Используем строковые orientation ('portrait' / 'landscape')
        try:
            ws.page_setup.orientation = 'portrait'
        except Exception:
            pass
        # Выключаем FitTo (чтобы Zoom применялся)
        #try:
        #    ws.page_setup.fitToWidth = False
        #    ws.page_setup.fitToHeight = False
        #except Exception:
            # может не поддерживаться, игнорируем
        #    pass

        # Выключаем "вписать в страницу" — ГЛАВНЫЙ переключатель
        try:
            if ws.sheet_properties.pageSetUpPr is None:
                ws.sheet_properties.pageSetUpPr = PageSetupProperties()
            ws.sheet_properties.pageSetUpPr.fitToPage = False
            ws.page_setup.fitToWidth = None
            ws.page_setup.fitToHeight = None
        except Exception:
            pass

        try:
            ws.page_setup.scale = 100
        except Exception:
            pass
        ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.5, bottom=0.75, header=0.25, footer=0.25)
        try:
            if hasattr(ws, '_horizontal_page_breaks'):
                ws._horizontal_page_breaks = []
            if hasattr(ws, '_vertical_page_breaks'):
                ws._vertical_page_breaks = []
        except:
            pass
        logger.info(f"[{ws.title}] Настройки печати сброшены (openpyxl)")
        return True
    except Exception as e:
        logger.warning(f"Ошибка сброса openpyxl: {e}")
        return False


def reset_page_setup_com(ws, excel_app=None):
    """Сбрасывает настройки печати через COM (Win32)"""
    try:
        # удаляем page breaks
        try:
            while ws.HPageBreaks.Count > 0:
                try:
                    ws.HPageBreaks(1).Delete()
                except:
                    break
        except Exception:
            pass
        try:
            while ws.VPageBreaks.Count > 0:
                try:
                    ws.VPageBreaks(1).Delete()
                except:
                    break
        except Exception:
            pass
        try:
            ws.PageSetup.PrintArea = ""
            ws.PageSetup.PrintTitleRows = ""
            ws.PageSetup.PrintTitleColumns = ""
        except Exception:
            pass
        try:
            # Выключаем FitToPages
            ws.PageSetup.FitToPagesWide = False
            ws.PageSetup.FitToPagesTall = False
        except Exception:
            pass
        logger.info(f"[{ws.Name}] Настройки печати сброшены (COM)")
        return True
    except Exception as e:
        logger.warning(f"Ошибка сброса COM: {e}")
        return False


def is_sheet_empty(ws):
    """Проверяет лист openpyxl на пустоту"""
    try:
        if ws.max_row <= 1 and ws.max_column <= 1:
            for row in ws.iter_rows(max_row=1, max_column=1):
                for cell in row:
                    if cell.value is not None and str(cell.value).strip():
                        return False
            return True
        return False
    except Exception:
        return True


def is_sheet_empty_com(ws):
    """Проверяет лист в COM на пустоту"""
    try:
        used = ws.UsedRange
        if used is None:
            return True
        val = getattr(used, "Value", None)
        if val is None:
            return True
        # Если это одиночное значение
        if not isinstance(val, tuple):
            if val in (None, "", 0):
                return True
            return False
        # в случае массива
        for row in val:
            if isinstance(row, tuple):
                for cell in row:
                    if cell not in (None, "", 0):
                        return False
            else:
                if row not in (None, "", 0):
                    return False
        return True
    except Exception:
        return True


def hard_clear_fills_only(wb):
    empty_fill = PatternFill(fill_type=None)
    for ws in wb.worksheets:
        try:
            if is_sheet_empty(ws):
                continue
            max_row = ws.max_row or 1
            max_col = ws.max_column or 1
            for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
                for cell in row:
                    try:
                        cell.fill = empty_fill
                    except:
                        continue
            try:
                for rng in list(ws.merged_cells.ranges):
                    try:
                        master = ws.cell(rng.min_row, rng.min_col)
                        master.fill = empty_fill
                    except:
                        continue
            except:
                pass
        except Exception as e:
            logger.warning(f"Ошибка очистки заливки на листе {ws.title}: {e}")
    logger.info("Удалена только заливка во всех листах книги")


# -----------------------
# Определение заголовков и области печати
# -----------------------
def find_header_rows(ws, first_data_row=1):
    try:
        marker_row = None
        for col in [1, 2, 3]:
            for row in range(first_data_row, min(first_data_row + 25, ws.max_row + 1)):
                cell_value = ws.cell(row=row, column=col).value
                cell_str = str(cell_value).strip().lower() if cell_value else ""
                if cell_str in ['a', 'b', 'c', 'а', 'б', 'в', 'п/п', '№']:
                    marker_row = row
                    break
                elif re.fullmatch(r'^[123]$', cell_str):
                    marker_row = row
                    break
            if marker_row is not None:
                break
        if marker_row is None:
            logger.info(f"[{ws.title}] Маркер не найден → переходим к поиску по границам.")
            return find_header_rows_any_border_advanced(ws)
        header_block = [marker_row]
        current_row = marker_row - 1
        while current_row >= 1:
            has_right_border = False
            for col in range(1, min(12, ws.max_column + 1)):
                border = ws.cell(row=current_row, column=col).border
                if border.right and border.right.style:
                    has_right_border = True
                    break
            if has_right_border:
                header_block.append(current_row)
                current_row -= 1
            else:
                break
        start_row = min(header_block)
        end_row = max(header_block)
        logger.info(f"[{ws.title}] ✅ Заголовок найден: строки {start_row}-{end_row}")
        return start_row, end_row
    except Exception as e:
        logger.error(f"[{ws.title}] Ошибка при поиске заголовков: {e}")
        return 1, 1


def find_header_rows_any_border_advanced(ws, max_search_rows=30):
    try:
        all_border_rows = []
        for row in range(1, min(max_search_rows, ws.max_row + 1)):
            for col in range(1, min(20, ws.max_column + 1)):
                border = ws.cell(row=row, column=col).border
                if any([
                    border.top and border.top.style,
                    border.right and border.right.style
                ]):
                    all_border_rows.append(row)
                    break
        if not all_border_rows:
            logger.info(f"[{ws.title}] ⚠️ Границы не найдены — используем первую строку.")
            return 1, 1
        blocks = []
        current_block = [all_border_rows[0]]
        for i in range(1, len(all_border_rows)):
            if all_border_rows[i] == all_border_rows[i - 1] + 1:
                current_block.append(all_border_rows[i])
            else:
                blocks.append(current_block)
                current_block = [all_border_rows[i]]
        blocks.append(current_block)
        first_block = blocks[0]
        start_row, end_row = min(first_block), max(first_block)
        logger.info(f"[{ws.title}] 🧩 Границы найдены: {start_row}-{end_row}")
        return start_row, end_row
    except Exception as e:
        logger.error(f"[{ws.title}] Ошибка поиска по границам: {e}")
        return 1, 1


def find_header_rows_smart(ws):
    try:
        start_row, end_row = find_header_rows(ws)
        if end_row - start_row > 8:
            logger.warning(f"[{ws.title}] Заголовок слишком большой ({start_row}:{end_row}) → ограничиваем 2–4")
            start_row, end_row = 2, min(4, ws.max_row)
        return start_row, end_row
    except Exception as e:
        logger.error(f"[{ws.title}] Ошибка при комбинированном поиске: {e}")
        return 1, 1


def safe_add_page_number(ws):
    try:
        # openpyxl Footer API может быть разной версии — используем простую запись в footer string
        try:
            ws.oddFooter.center.text = "Страница &P из &N"
        except Exception:
            try:
                ws.oddFooter.center = "Страница &P из &N"
            except Exception:
                pass
        try:
            ws.evenFooter.center.text = "Страница &P из &N"
        except Exception:
            try:
                ws.evenFooter.center = "Страница &P из &N"
            except Exception:
                pass
        return True
    except Exception as e:
        logger.warning(f"Ошибка добавления номера страницы: {e}")
        return False


# -----------------------
# Размеры таблицы / печатной области
# -----------------------
def calculate_table_dimensions(ws):
    """
    Возвращает (total_width_inches, total_height_inches) оценочные размеры таблицы.
    Использует column.width и row.height, с запасом по конвертации.
    """
    try:
        default_col_width = getattr(ws.sheet_format, 'defaultColWidth', None) or 8.43
        default_row_height = getattr(ws.sheet_format, 'defaultRowHeight', None) or 15.0

        total_width_inches = 0.0
        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            col_dim = ws.column_dimensions.get(col_letter)
            if col_dim and getattr(col_dim, "hidden", False):
                continue
            width = col_dim.width if (col_dim and col_dim.width is not None) else default_col_width
            # приближённое преобразование ширины столбца Excel -> пиксели -> дюймы
            width_in_pixels = (width * 7 * CALIBRATION_FACTOR) + 5
            width_in_inches = width_in_pixels / 96.0
            total_width_inches += width_in_inches

        total_height_points = 0.0
        for row_idx in range(1, ws.max_row + 1):
            row_dim = ws.row_dimensions.get(row_idx)
            if row_dim and getattr(row_dim, "hidden", False):
                continue
            height_points = row_dim.height if (row_dim and row_dim.height is not None) else default_row_height
            total_height_points += height_points
        total_height_inches = total_height_points * EXCEL_POINT_TO_INCH

        logger.info(f"[{ws.title}] Размеры таблицы: {total_width_inches:.2f}\" x {total_height_inches:.2f}\"")
        return total_width_inches, total_height_inches
    except Exception as e:
        logger.error(f"Ошибка расчёта размеров таблицы: {e}")
        return 0.0, 0.0


def is_real_data_cell(ws, row, col):
    try:
        cell = ws.cell(row=row, column=col)
        if cell.value is not None:
            cell_str = str(cell.value).strip()
            if cell_str and cell_str not in ['', ' ', '\t', '\n', '\r']:
                return True
        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_row <= row <= merged_range.max_row and merged_range.min_col <= col <= merged_range.max_col:
                master_cell = ws.cell(merged_range.min_row, merged_range.min_col)
                if master_cell.value is not None:
                    return True
        return False
    except Exception:
        return False


def auto_fit_print_area_improved(ws):
    try:
        logger.info(f"[{ws.title}] УЛУЧШЕННЫЙ поиск реальных данных для PrintArea")
        max_real_row = 1
        max_real_col = 1
        empty_block_count = 0
        for row in range(1, min(ws.max_row, 1000) + 1):
            row_has_real_data = False
            for col in range(1, min(ws.max_column, 100) + 1):
                if is_real_data_cell(ws, row, col):
                    row_has_real_data = True
                    max_real_row = max(max_real_row, row)
                    max_real_col = max(max_real_col, col)
            if not row_has_real_data and row > 10:
                empty_block_count += 1
            else:
                empty_block_count = 0
            if empty_block_count >= 5:
                logger.info(f"[{ws.title}] Найдены 5+ пустых строк подрят - КОНЕЦ ДАННЫХ")
                break
        print_max_row = min(max_real_row + 2, ws.max_row)
        print_max_col = min(max_real_col + 2, ws.max_column)
        if print_max_row > 200:
            logger.warning(f"[{ws.title}] Слишком большая область, ограничиваем")
            print_max_row = min(max_real_row, 100)
        if print_max_row > 1 or print_max_col > 1:
            print_area = f"A1:{get_column_letter(print_max_col)}{print_max_row}"
            ws.print_area = print_area
            logger.info(f"[{ws.title}] ✅ УСТАНОВЛЕНА СТРОГАЯ область печати: {print_area}")
            return True
        else:
            logger.info(f"[{ws.title}] ❌ Не найдены реальные данные")
            return False
    except Exception as e:
        logger.warning(f"Ошибка улучшенной авто-области печати: {e}")
        return False


def aggressive_clean_empty_regions(ws):
    try:
        last_real_row = 1
        last_real_col = 1
        for merged_range in ws.merged_cells.ranges:
            master_cell = ws.cell(merged_range.min_row, merged_range.min_col)
            if master_cell.value is None:
                try:
                    ws.unmerge_cells(str(merged_range))
                    logger.info(f"[{ws.title}] Удалено объединение: {merged_range}")
                except Exception:
                    pass
            if master_cell.value is not None:
                last_real_row = max(last_real_row, merged_range.max_row)
                last_real_col = max(last_real_col, merged_range.max_col)
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                try:
                    if is_real_data_cell(ws, cell.row, cell.column):
                        last_real_row = max(last_real_row, cell.row)
                        last_real_col = max(last_real_col, cell.column)
                except:
                    continue
        logger.info(f"[{ws.title}] Реальные данные до строки {last_real_row}, столбца {last_real_col}")
        for row in range(last_real_row + 5, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                try:
                    cell = ws.cell(row=row, column=col)
                    border = cell.border
                    has_border = (border.top and border.top.style) or (border.bottom and border.bottom.style) or (
                                border.left and border.left.style) or (border.right and border.right.style)
                    if has_border and cell.value is None:
                        cell.border = None
                        logger.debug(f"[{ws.title}] Очищена граница в строке {row}, столбце {col}")
                except:
                    continue
        logger.info(f"[{ws.title}] Агрессивно очищены пустые области")
        return True
    except Exception as e:
        logger.warning(f"Ошибка агрессивной очистки: {e}")
        return False


# -----------------------
# Ориентация и масштаб (openpyxl -> COM)
# -----------------------
def clamp_zoom(n):
    try:
        n = int(n)
    except Exception:
        return None
    if n < 30:
        return 30
    if n > 100:
        return 100
    return n


def set_orientation_and_scale_openpyxl(ws):
    """
    Рассчитывает и задаёт orientation и scale через openpyxl.
    Возвращает recommended_scale (int) или None.
    """
    try:
        logger.info(f"[{ws.title}] === Расчёт ориентации и масштаба (openpyxl) ===")
        reset_page_setup_openpyxl(ws)
        total_width, total_height = calculate_table_dimensions(ws)

        if total_width <= 0:
            logger.warning(f"[{ws.title}] Не удалось рассчитать размеры таблицы (width=0)")
            try:
                ws.page_setup.orientation = 'portrait'
                ws.page_setup.scale = 100
            except:
                pass
            return None
        # установить бумагу A4
        try:
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
        except Exception:
            # не критично
            pass
        margins = PageMargins(left=0.25, right=0.25, top=0.5, bottom=0.75, header=0.25, footer=0.25)
        ws.page_margins = margins
        available_width_portrait = PAGE_WIDTH_A4_PORTRAIT - (margins.left + margins.right)
        available_width_landscape = PAGE_HEIGHT_A4_PORTRAIT - (margins.left + margins.right)

        logger.info(
            f"[{ws.title}] total_width={total_width:.2f}\" available_portrait={available_width_portrait:.2f}\" available_landscape={available_width_landscape:.2f}\"")
        # попытка портрет
        portrait_scale = None
        for scale in range(100, 59, -1):
            if total_width * (scale / 100.0) <= available_width_portrait:
                portrait_scale = scale
                break
        if portrait_scale:
            try:
                ws.page_setup.orientation = 'portrait'
                ws.page_setup.fitToWidth = False
                ws.page_setup.fitToHeight = False
                ws.page_setup.scale = max(60, portrait_scale)
                logger.info(f"[{ws.title}] Выбрана портретная ориентация, scale={ws.page_setup.scale}%")
                return clamp_zoom(ws.page_setup.scale)
            except Exception:
                pass
        # попытка альбом
        landscape_scale = None
        for scale in range(100, 59, -1):
            if total_width * (scale / 100.0) <= available_width_landscape:
                landscape_scale = scale
                break
        if landscape_scale:
            try:
                ws.page_setup.orientation = 'landscape'
                ws.page_setup.fitToWidth = False
                ws.page_setup.fitToHeight = False
                ws.page_setup.scale = max(60, landscape_scale)
                logger.info(f"[{ws.title}] Выбрана альбомная ориентация, scale={ws.page_setup.scale}%")
                return clamp_zoom(ws.page_setup.scale)
            except Exception:
                pass
        # форсирование
        try:
            ws.page_setup.orientation = 'landscape'
            ws.page_setup.fitToWidth = False
            ws.page_setup.fitToHeight = False
            ws.page_setup.scale = 60
            logger.warning(f"[{ws.title}] Форсированная альбомная ориентация (scale=60%)")
            return 60
        except Exception:
            pass
        return None
    except Exception as e:
        logger.error(f"[{ws.title}] Ошибка установки ориентации/масштаба (openpyxl): {e}", exc_info=True)
        return None


# -----------------------
# Диагностика (печать текущих PageSetup)
# -----------------------
def log_openpyxl_pagesetup(ws):
    try:
        ps = ws.page_setup
        logger.info(
            f"[{ws.title}] openpyxl PageSetup: orientation={getattr(ps, 'orientation', None)}, scale={getattr(ps, 'scale', None)}, fitToWidth={getattr(ps, 'fitToWidth', None)}, fitToHeight={getattr(ps, 'fitToHeight', None)}, paperSize={getattr(ps, 'paperSize', None)}")
    except Exception:
        pass


def log_com_pagesetup(ws):
    try:
        ps = ws.PageSetup
        logger.info(
            f"[{ws.Name}] COM PageSetup: Zoom={getattr(ps, 'Zoom', None)}, FitToPagesWide={getattr(ps, 'FitToPagesWide', None)}, FitToPagesTall={getattr(ps, 'FitToPagesTall', None)}, Orientation={getattr(ps, 'Orientation', None)}, PaperSize={getattr(ps, 'PaperSize', None)}")
    except Exception:
        pass


# -----------------------
# Диагностика (скриптовая)
# -----------------------
def diagnose_scale(xlsx_path):
    wb = load_workbook(xlsx_path)
    for ws in wb.worksheets:
        total_width, total_height = calculate_table_dimensions(ws)
        margins = ws.page_margins
        available_portrait = PAGE_WIDTH_A4_PORTRAIT - (margins.left + margins.right)
        available_landscape = PAGE_HEIGHT_A4_PORTRAIT - (margins.left + margins.right)
        print(f"\nФайл: {xlsx_path}, Лист: {ws.title}")
        print(f"Ширина таблицы (дюймы): {total_width:.2f}")
        print(f"Доступно (портрет): {available_portrait:.2f}, (альбом): {available_landscape:.2f}")
        if total_width <= available_portrait:
            print("→ Уместится в портрет при 100%")
        elif total_width <= available_landscape:
            print("→ Уместится в альбом при 100%")
        else:
            for scale in range(100, 59, -1):
                if total_width * (scale / 100.0) <= available_landscape:
                    print(f"→ Нужно уменьшить до {scale}% (альбом)")
                    break
        ps = ws.page_setup
        print("Текущие настройки PageSetup (openpyxl):")
        print(f"  orientation: {getattr(ps, 'orientation', None)}")
        print(f"  scale: {getattr(ps, 'scale', None)}")
        print(f"  fitToWidth: {getattr(ps, 'fitToWidth', None)}")
        print(f"  fitToHeight: {getattr(ps, 'fitToHeight', None)}")


# -----------------------
# COM / Excel instance
# -----------------------
def create_excel_instance(visible=False):
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = visible
    excel.DisplayAlerts = False
    return excel


# -----------------------
# Основная обработка файла
# -----------------------
def get_modified_path(original_path, root_folder):
    original_path = Path(original_path)
    root_folder = Path(root_folder)
    try:
        if root_folder not in original_path.parents:
            rel_path = Path(original_path.name)
        else:
            rel_path = original_path.relative_to(root_folder)
    except ValueError:
        rel_path = Path(original_path.name)
    output_root = original_path.parent
    return output_root / "modified" / rel_path


def process_excel_file(file_path, log_func=None, excel_app=None):
    """
    Обрабатывает один файл: openpyxl -> save -> COM финализируем PageSetup
    Возвращает True/False
    """
    temp_xlsx = None
    wb = None
    local_excel = False
    header_rows_map = {}
    try:
        file_path = str(file_path)
        original_path_for_modified = Path(file_path)
        if not os.path.exists(file_path):
            logger.info(f"Файл не существует: {file_path}")
            return False
        if excel_app is None:
            excel = create_excel_instance(visible=False)
            local_excel = True
        else:
            excel = excel_app
        # Конвертируем .xls -> .xlsx если нужно
        if file_path.lower().endswith('.xls') and not file_path.lower().endswith('.xlsx'):
            temp_xlsx = file_path + 'x'
            try:
                logger.info(f"Конвертируем {file_path} -> {temp_xlsx}")
                wb_com = None
                try:
                    wb_com = excel.Workbooks.Open(file_path)
                    wb_com.SaveAs(temp_xlsx, FileFormat=51)
                finally:
                    if wb_com is not None:
                        try:
                            wb_com.Close()
                        except Exception:
                            pass
                file_path_for_openpyxl = temp_xlsx
            except Exception as e:
                logger.info(f"Ошибка конвертации: {e}")
                if local_excel:
                    try:
                        excel.Quit()
                    except:
                        pass
                if temp_xlsx and os.path.exists(temp_xlsx):
                    try:
                        os.remove(temp_xlsx)
                    except:
                        pass
                return False
        else:
            file_path_for_openpyxl = file_path
        # === OPENPYXL ОБРАБОТКА ===
        try:
            wb = load_workbook(filename=file_path_for_openpyxl, read_only=False, keep_vba=False)
            cleanup_old_print_settings(wb)
            hard_clear_fills_only(wb)
            # сохранение в modified пути
            modified_path = get_modified_path(original_path_for_modified,
                                              os.path.dirname(original_path_for_modified)).with_suffix('.xlsx')
            modified_path.parent.mkdir(parents=True, exist_ok=True)
            # iterate worksheets
            for ws in wb.worksheets:
                try:
                    if is_sheet_empty(ws):
                        logger.info(f"[{ws.title}] Пустой лист - пропускаем")
                        continue
                    reset_page_setup_openpyxl(ws)
                    start_row, end_row = find_header_rows_smart(ws)
                    if end_row - start_row > 5:
                        logger.warning(f"[{ws.title}] Слишком большой заголовок {start_row}:{end_row}, используем 1:2")
                        start_row, end_row = 1, 2
                    header_rows_map[ws.title] = (start_row, end_row)
                    ws.print_title_rows = f"{start_row}:{end_row}"
                    ws.print_title_cols = "A:A"
                    if safe_add_page_number(ws):
                        logger.info(f"[{ws.title}] Добавлен номер страницы")
                    aggressive_clean_empty_regions(ws)
                    auto_fit_print_area_improved(ws)


                    # Рассчитать и установить ориентацию/масштаб через openpyxl (рекомендуемое значение)
                    rec_scale = set_orientation_and_scale_openpyxl(ws)
                    log_openpyxl_pagesetup(ws)
                    # сохраняем recommended scale в свойства листа (для дальнейшего COM применения)
                    if rec_scale:
                        try:
                            ws.sheet_properties.customProperties = ws.sheet_properties.customProperties if hasattr(
                                ws.sheet_properties, 'customProperties') else None
                        except:
                            pass
                except Exception as e:
                    logger.info(f"Ошибка обработки листа {ws.title}: {e}")
                    continue
            wb.save(str(modified_path))
            logger.info(f"Сохранено после openpyxl обработки: {modified_path}")
        except Exception as e:
            logger.info(f"Ошибка загрузки/очистки openpyxl: {e}")
            if local_excel:
                try:
                    excel.Quit()
                except:
                    pass
            return False
        # === COM финализация (гарантированное применение Zoom/PrintTitles) ===
        wb_com = None
        try:
            wb_com = excel.Workbooks.Open(str(modified_path))
            # попытка установить конкретный ActivePrinter (опционально)
            try:
                # оставляем как есть — изменение принтера может быть нежелательно
                current_printer = excel.ActivePrinter
                logger.debug(f"ActivePrinter (текущий): {current_printer}")
            except Exception:
                pass
            # Пройдём по листам COM и применим PageSetup согласно вычислениям openpyxl
            for ws in wb_com.Sheets:
                try:
                    if is_sheet_empty_com(ws):
                        logger.info(f"[{ws.Name}] Пустой лист (COM) - пропускаем")
                        continue
                    # Сбрасываем FitTo (чтобы Zoom применялся)
                    try:
                        ws.PageSetup.FitToPagesWide = False
                        ws.PageSetup.FitToPagesTall = False
                    except Exception:
                        pass
                    # Применяем PrintTitleRows / PrintTitleColumns (если были рассчитаны)
                    if ws.Name in header_rows_map:
                        start_row, end_row = header_rows_map[ws.Name]
                        try:
                            ws.PageSetup.PrintTitleRows = f"${start_row}:${end_row}"
                            ws.PageSetup.PrintTitleColumns = "$A:$A"
                            logger.info(f"[{ws.Name}] Сквозные строки установлены в COM {start_row}:{end_row}")
                        except Exception:
                            pass
                    # Попробуем получить значение scale из openpyxl (если записали туда)
                    rec_scale = None
                    try:
                        # читаем из openpyxl workbook (wb) той же книги
                        if wb and ws.Name in wb.sheetnames:
                            ws_ox = wb[ws.Name]
                            rec_scale = getattr(ws_ox.page_setup, "scale", None)
                    except Exception:
                        rec_scale = None
                    # Если openpyxl предложил масштаб < 100, устанавливаем его через COM
                    if rec_scale and isinstance(rec_scale, (int, float)) and rec_scale < 100:
                        # Дополнительная диагностика размеров в COM
                        try:
                            used = ws.UsedRange
                            used_width_pts = float(getattr(used, "Width", 0.0) or 0.0)
                            used_height_pts = float(getattr(used, "Height", 0.0) or 0.0)
                            logger.info(
                                f"[{ws.Name}] UsedRange.Width pts = {used_width_pts:.2f}, Height pts = {used_height_pts:.2f}")
                        except Exception:
                            used_width_pts = used_height_pts = 0.0
                        # Printable area pts
                        try:
                            paper_w = PAGE_WIDTH_A4_PORTRAIT
                            paper_h = PAGE_HEIGHT_A4_PORTRAIT
                            if getattr(ws.PageSetup, "Orientation", None) == 2:  # xlLandscape = 2
                                paper_w, paper_h = PAGE_HEIGHT_A4_PORTRAIT, PAGE_WIDTH_A4_PORTRAIT
                            left_margin = float(getattr(ws.PageSetup, "LeftMargin", 0.0) or 0.0)
                            right_margin = float(getattr(ws.PageSetup, "RightMargin", 0.0) or 0.0)
                            top_margin = float(getattr(ws.PageSetup, "TopMargin", 0.0) or 0.0)
                            bottom_margin = float(getattr(ws.PageSetup, "BottomMargin", 0.0) or 0.0)
                        except Exception:
                            left_margin = right_margin = top_margin = bottom_margin = excel.InchesToPoints(0.25)
                        printable_width_pts = excel.InchesToPoints(paper_w) - (left_margin + right_margin)
                        printable_height_pts = excel.InchesToPoints(paper_h) - (top_margin + bottom_margin)
                        logger.info(
                            f"[{ws.Name}] printable_width_pts = {printable_width_pts:.2f}, printable_height_pts = {printable_height_pts:.2f}")
                        # Попытка суммировать ширины колонок через COM Range.Width
                        total_cols_width_pts = 0.0
                        try:
                            # определяем диапазон колонок: используем PrintArea если есть
                            col_min = 1
                            col_max = None
                            pa = ""
                            try:
                                pa = ws.PrintArea
                            except Exception:
                                try:
                                    pa = ws.Range("A1").Worksheet.PrintArea
                                except Exception:
                                    pa = ""
                            if pa:
                                rng = pa.split('!')[-1].replace('$', '')
                                if ':' in rng:
                                    a, b = rng.split(':')
                                    col_min = column_index_from_string(re.sub(r'\d', '', a))
                                    col_max = column_index_from_string(re.sub(r'\d', '', b))
                            if col_max is None:
                                # fallback: UsedRange.Columns.Count
                                try:
                                    col_max = int(ws.UsedRange.Columns.Count)
                                except Exception:
                                    # ещё fallback: максимально возможное из Excel (редко нужно)
                                    col_max = int(ws.Cells.Columns.Count)
                            from openpyxl.utils import get_column_letter as gcl
                            for ci in range(col_min, col_max + 1):
                                cl = get_column_letter(ci)
                                try:
                                    w = float(ws.Range(f"{cl}:{cl}").Width or 0.0)
                                except Exception:
                                    try:
                                        # ColumnWidth -> approximate
                                        w = float(ws.Columns(ci).ColumnWidth or 0.0)
                                    except Exception:
                                        w = 0.0
                                total_cols_width_pts += w
                        except Exception as e:
                            logger.warning(f"[{ws.Name}] Не удалось посчитать ширину колонок в COM: {e}")
                        logger.info(f"[{ws.Name}] total_cols_width_pts = {total_cols_width_pts:.2f}")
                        width_for_calc = total_cols_width_pts if total_cols_width_pts > 0 else used_width_pts
                        try:
                            zoom_w = int((
                                                     printable_width_pts / width_for_calc) * 100) if width_for_calc > 0 and printable_width_pts > 0 else None
                        except Exception:
                            zoom_w = None
                        try:
                            zoom_h = int((
                                                     printable_height_pts / used_height_pts) * 100) if used_height_pts > 0 and printable_height_pts > 0 else None
                        except Exception:
                            zoom_h = None
                        logger.info(f"[{ws.Name}] zoom_w={zoom_w}, zoom_h={zoom_h}")
                        candidates = [z for z in (zoom_w, zoom_h) if isinstance(z, int)]
                        computed_zoom = min(candidates) if candidates else None
                        if computed_zoom:
                            computed_zoom = max(30, min(100, int(computed_zoom)))
                        # окончательное значение
                        final_zoom = clamp_zoom(rec_scale) or computed_zoom or 100
                        try:
                            ws.PageSetup.FitToPagesWide = False
                            ws.PageSetup.FitToPagesTall = False
                        except Exception:
                            pass
                        try:
                            ws.PageSetup.Zoom = int(final_zoom)
                            logger.info(
                                f"[{ws.Name}] COM: установлен Zoom={final_zoom}% (rec_scale={rec_scale}, computed={computed_zoom})")
                        except Exception as e:
                            logger.warning(f"[{ws.Name}] Не удалось установить Zoom через COM: {e}")
                    # лог PageSetup после попытки
                    try:
                        log_com_pagesetup(ws)
                    except:
                        pass
                    # Order
                    try:
                        ws.PageSetup.Order = 2
                    except:
                        pass
                except Exception as e:
                    logger.info(f"[{ws.Name}] Ошибка Win32 настроек: {e}")
            # сохранить COM книгу
            wb_com.Save()
        except Exception as e:
            logger.info(f"Win32COM ошибка: {e}")
            return False
        finally:
            if wb_com is not None:
                try:
                    wb_com.Close()
                except Exception:
                    pass
        return True
    finally:
        if wb:
            try:
                wb.close()
            except:
                pass
        if temp_xlsx and os.path.exists(temp_xlsx):
            try:
                os.remove(temp_xlsx)
                logger.info(f"Удалён временный файл: {temp_xlsx}")
            except Exception as e:
                logger.info(f"Не удалось удалить временный файл: {e}")
        if local_excel and 'excel' in locals():
            try:
                excel.Quit()
            except:
                pass


# -----------------------
# GUI и обработка папки
# -----------------------
class ExcelProcessor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.folder_path = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar(value="Готово")
        self.log_messages = []
        self.delete_archives_var = tk.BooleanVar(value=False)
        self.progress_bar = None
        self.setup_ui()

    def setup_ui(self):
        self.root.deiconify()
        self.root.title("Wizard Archive")
        self.root.geometry("700x500")
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(main_frame, text="Wizard Archive", font=('Helvetica', 14, 'bold')).pack(pady=10)
        ttk.Label(main_frame, text="Папка с файлами:").pack(anchor="w")
        ttk.Entry(main_frame, textvariable=self.folder_path, width=60).pack(fill="x", pady=5)
        delete_frame = ttk.Frame(main_frame)
        delete_frame.pack(pady=5)
        ttk.Checkbutton(delete_frame, text="Удалить архивы после распаковки", variable=self.delete_archives_var).pack(
            side=tk.LEFT)
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Обзор", command=self.select_folder).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Только распаковать архивы", command=self.unpack_only).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Полная обработка", command=self.start_processing).pack(side=tk.LEFT, padx=5)
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", pady=10)
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill="x", pady=5)
        ttk.Label(status_frame, text="Статус:").pack(side=tk.LEFT)
        ttk.Label(status_frame, textvariable=self.status_var).pack(side=tk.LEFT, padx=10)
        log_frame = ttk.LabelFrame(main_frame, text="Лог выполнения", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        self.log_text = tk.Text(log_frame, height=12, width=80, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(self.log_text)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.log_text.yview)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def log_to_gui(self, message):
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        print(message)

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path.set(folder)
            self.log_to_gui(f"Выбрана папка: {folder}")

    def unpack_only(self):
        folder = self.folder_path.get()
        if not folder:
            messagebox.showerror("Ошибка", "Сначала выберите папку!")
            return
        folder_path = Path(folder)
        self.status_var.set("Распаковка архивов...")
        self.progress_var.set(10)
        self.log_to_gui("Начинаю распаковку архивов...")
        self.root.update()
        try:
            processed_folders = process_archives_in_folder(folder_path, self.delete_archives_var.get())
            deleted_count = clean_empty_folders(folder_path)
            if deleted_count > 0:
                self.log_to_gui(f"Удалено {deleted_count} пустых папок")
            self.status_var.set(f"Распаковано {len(processed_folders)} архивов")
            self.progress_var.set(100)
            if processed_folders:
                self.log_to_gui(f"Распаковка завершена! Создано папок: {len(processed_folders)}")
                delete_status = "УДАЛЕНЫ" if self.delete_archives_var.get() else "СОХРАНЕНЫ"
                messagebox.showinfo("Готово",
                                    f"Распаковано {len(processed_folders)} архивов\nОригинальные архивы: {delete_status}")
            else:
                self.log_to_gui("Архивы не найдены")
                messagebox.showinfo("Готово", "Архивы не найдены")
        except Exception as e:
            self.log_to_gui(f"Ошибка распаковки: {str(e)}")
            messagebox.showerror("Ошибка", f"Ошибка распаковки:\n{str(e)}")

    def start_processing(self):
        folder = self.folder_path.get()
        if not folder:
            messagebox.showerror("Ошибка", "Сначала выберите папку!")
            return
        self.log_to_gui("=" * 60)
        self.log_to_gui("НАЧАЛО ПОЛНОЙ ОБРАБОТКИ")
        self.log_to_gui("=" * 60)
        import threading
        threading.Thread(target=self._process_all, args=(folder,), daemon=True).start()

    def _process_all(self, folder):
        excel = None
        try:
            pythoncom.CoInitialize()
            folder_path = Path(folder)
            self.status_var.set("Распаковка архивов...")
            self.progress_var.set(20)
            self.log_to_gui("Этап 1: Распаковка архивов")
            self.root.update()
            processed_folders = process_archives_in_folder(folder_path, self.delete_archives_var.get())
            if processed_folders:
                self.log_to_gui(f"Распаковано архивов: {len(processed_folders)}")
                deleted_count = clean_empty_folders(folder_path)
                if deleted_count > 0:
                    self.log_to_gui(f"Удалено {deleted_count} пустых папок")
            else:
                self.log_to_gui("Архивы не найдены")
            self.status_var.set("Поиск Excel файлов...")
            self.progress_var.set(40)
            self.log_to_gui("\nЭтап 2: Обработка Excel файлов")
            self.root.update()
            files = list(folder_path.rglob("*.xls")) + list(folder_path.rglob("*.xlsx"))
            files = [f for f in files if "modified" not in str(f).lower() and not f.name.startswith('~$')]
            total_files = min(len(files), MAX_FILES_TO_PROCESS)
            if total_files == 0:
                self.status_var.set("Excel файлы не найдены")
                self.log_to_gui("Excel файлы не найдены")
                messagebox.showinfo("Завершено", "В папке не найдено файлов .xls или .xlsx")
                return
            excel = create_excel_instance(visible=False)
            self.log_to_gui(f"Найдено Excel файлов: {total_files}")
            success_count = 0
            error_count = 0
            error_log = []
            for i, file in enumerate(files):
                if not file.is_file():
                    continue
                self.status_var.set(f"Обработка {file.name}...")
                progress = 40 + (i * 60 / total_files)
                self.progress_var.set(progress)
                self.root.update()
                self.log_to_gui(f"   Обработка: {file.name}")
                start_time = time.time()
                try:
                    result = process_excel_file(str(file), excel_app=excel)
                    elapsed = time.time() - start_time
                    if result:
                        success_count += 1
                        self.log_to_gui(f"   Успешно ({elapsed:.1f} сек)")
                    else:
                        error_count += 1
                        error_log.append(str(file))
                        self.log_to_gui(f"   Ошибка ({elapsed:.1f} сек)")
                except Exception as e:
                    error_count += 1
                    error_log.append(f"{file} - {str(e)}")
                    self.log_to_gui(f"   Критическая ошибка: {str(e)}")
            if error_log:
                log_dir = folder_path / "modified" / "logs"
                log_dir.mkdir(parents=True, exist_ok=True)
                log_file = log_dir / "error_log.txt"
                with open(log_file, 'w', encoding='utf-8') as f:
                    f.write("Файлы, при обработке которых произошли ошибки:\n")
                    f.write("\n".join(error_log))
                self.log_to_gui(f"Лог ошибок сохранен: {log_file}")
            self.status_var.set("Готово")
            self.progress_var.set(100)
            self.log_to_gui("\n" + "=" * 60)
            self.log_to_gui("ОБРАБОТКА ЗАВЕРШЕНА")
            self.log_to_gui(f"Успешно: {success_count}")
            self.log_to_gui(f"Ошибки: {error_count}")
            self.log_to_gui("=" * 60)
            messagebox.showinfo("Завершено",
                                f"Обработка завершена!\n\nУспешно: {success_count}\nОшибки: {error_count}\n\nОбработанные файлы сохранены в папке 'modified'")
        except Exception as ex:
            self.log_to_gui(f"Критическая ошибка: {str(ex)}")
            messagebox.showerror("Ошибка", f"Процесс прерван:\n{str(ex)}")
            traceback.print_exc()
        finally:
            if excel:
                try:
                    excel.Quit()
                except Exception:
                    pass
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    def on_close(self):
        if messagebox.askokcancel("Выход", "Вы действительно хотите выйти?"):
            self.root.destroy()

    def run(self):
        self.root.mainloop()


# -----------------------
# Запуск
# -----------------------
if __name__ == "__main__":
    try:
        logger.info("=== Запуск скрипта обработки Excel (fixed) ===")
        app = ExcelProcessor()
        app.run()
    except KeyboardInterrupt:
        logger.info("Остановка по инициативе пользователя.")
    except Exception as e:
        logger.error(f"Произошла непредвиденная ошибка: {e}")
        traceback.print_exc()
