from openpyxl.cell import cell
from openpyxl.worksheet.page import PageMargins
# !/usr/bin/env python3
# -*- coding: utf-8 -*-
import logging
import os
import time
import tkinter as tk
import traceback
from tkinter import filedialog, messagebox, ttk
import re
from pathlib import Path
import zipfile
import math

try:
    import rarfile
except ImportError:
    rarfile = None

import pythoncom
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

import win32com.client as win32
from openpyxl.worksheet.page import PageMargins
from openpyxl.styles import PatternFill

# Настройка логгера
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

# Константы для формата бумаги A4 в дюймах
PAGE_WIDTH_A4_PORTRAIT = 6.5  # Ширина A4 в дюймах (портрет)
PAGE_HEIGHT_A4_PORTRAIT = 11.69  # Высота A4 в дюймах (портрет)

# Коэффициенты
EXCEL_POINT_TO_INCH = 1 / 72
CALIBRATION_FACTOR = 1.02
MAX_FILES_TO_PROCESS = 1000


def smart_autofit_rows_exact(ws_com):
    """
    Точная подгонка высоты строк с учётом длины wrap-текста и ширины столбца.
    Работает без потери WrapText.
    """
    try:
        used = ws_com.UsedRange
        last_row = used.Rows.Count
        last_col = used.Columns.Count

        for r in range(1, last_row + 1):
            row = ws_com.Rows(r)
            max_height = 0
            for c in range(1, last_col + 1):
                try:
                    cell = ws_com.Cells(r, c)
                    val = str(cell.Value).strip() if cell.Value is not None else ""
                    if not val:
                        continue

                    if cell.WrapText:
                        font = cell.Font
                        font_size = font.Size or 11
                        col_width = ws_com.Columns(c).ColumnWidth
                        avg_char_width = 1.05  # приблизительно для Calibri 11pt
                        max_chars_per_line = col_width / avg_char_width

                        # сколько строк займёт текст
                        num_lines = max(1, math.ceil(len(val) / max_chars_per_line))
                        height = num_lines * (font_size * 1.3)

                        if height > max_height:
                            max_height = height
                except Exception:
                    continue

            if max_height > 0:
                row.RowHeight = max_height

        logger.info(f"[{ws_com.Name}] ✅ Smart AutoFit (по содержимому) выполнен")

    except Exception as e:
        logger.warning(f"[{ws_com.Name}] Ошибка Smart AutoFit (точный): {e}")



# ============================================
# 1. ФУНКЦИИ ДЛЯ АРХИВОВ (ИСПРАВЛЕННЫЕ - НЕ УДАЛЯЕМ ОРИГИНАЛЫ!)
# ============================================

def extract_archive_with_correct_encoding(archive_path, output_folder=None):
    """
    Распаковывает архив с правильной кодировкой и НЕ удаляет оригинал
    """
    archive_path = Path(archive_path)

    # Определяем папку для распаковки
    if output_folder is None:
        output_folder = archive_path.parent

    # Создаем папку с именем архива (без расширения)
    folder_name = archive_path.stem

    # Попробуем разные кодировки для имен файлов
    encodings_to_try = ['cp866', 'cp1251', 'utf-8', 'iso-8859-1', 'koi8-r']

    extract_to = output_folder / folder_name

    try:
        # Создаем папку для распаковки
        extract_to.mkdir(parents=True, exist_ok=True)
        logger.info(f"Распаковка архива: {archive_path.name} → {folder_name}/")

        extracted_files = []

        # Распаковываем ZIP с обработкой кодировки
        if archive_path.suffix.lower() == ".zip":
            # Пробуем разные кодировки
            for encoding in encodings_to_try:
                try:
                    with zipfile.ZipFile(archive_path, 'r') as zip_ref:
                        # Получаем список всех файлов в архиве
                        file_list = zip_ref.namelist()

                        # Извлекаем все файлы
                        for file_name in file_list:
                            try:
                                # Пробуем декодировать имя файла
                                decoded_name = file_name.encode('cp437').decode(encoding)
                            except:
                                # Если не получается, используем оригинальное имя
                                decoded_name = file_name

                            # Полный путь для извлечения
                            target_path = extract_to / decoded_name
                            target_path.parent.mkdir(parents=True, exist_ok=True)

                            # Извлекаем файл
                            with zip_ref.open(file_name) as source, open(target_path, 'wb') as target:
                                target.write(source.read())

                            extracted_files.append(target_path)

                        logger.info(f"ZIP архив распакован с кодировкой {encoding} ({len(extracted_files)} файлов)")
                        break  # Успешно распаковали с этой кодировкой

                except Exception as e:
                    logger.debug(f"Кодировка {encoding} не подошла: {e}")
                    continue

        # Распаковываем RAR
        elif archive_path.suffix.lower() == ".rar" and rarfile:
            with rarfile.RarFile(archive_path, 'r') as rar_ref:
                # Получаем список всех файлов в архиве
                file_list = rar_ref.namelist()

                # Извлекаем все файлы
                rar_ref.extractall(extract_to)

                # Формируем список извлеченных файлов
                for file_name in file_list:
                    file_path = extract_to / file_name
                    if file_path.exists():
                        extracted_files.append(file_path)

                logger.info(f"RAR архив распакован ({len(extracted_files)} файлов)")

        else:
            logger.warning(f"Формат {archive_path.suffix} не поддерживается")
            return None

        # Рекурсивно обрабатываем вложенные архивы
        nested_archives = []
        for item in extracted_files:
            if item.is_file() and item.suffix.lower() in (".zip", ".rar"):
                nested_archives.append(item)

        # Обрабатываем все вложенные архивы
        for nested_archive in nested_archives:
            logger.info(f"Найден вложенный архив: {nested_archive.name}")
            # Рекурсивно распаковываем вложенный архив ВНУТРИ текущей папки
            extract_archive_with_correct_encoding(
                nested_archive,
                extract_to  # распаковываем внутрь текущей папки
            )
            # НЕ удаляем вложенный архив после распаковки!

        return extract_to

    except Exception as e:
        logger.error(f"Ошибка распаковки архива {archive_path.name}: {e}")
        return None


def process_archives_in_folder(folder_path, delete_originals=False):
    """
    Обрабатывает все архивы в папке (рекурсивно)
    delete_originals: если True - удаляет оригинальные архивы (по умолчанию False)
    """
    folder_path = Path(folder_path)
    processed_folders = []

    # Сначала обрабатываем архивы в корне папки
    root_archives = []
    for item in folder_path.iterdir():
        if item.is_file() and item.suffix.lower() in (".zip", ".rar"):
            root_archives.append(item)

    logger.info(f"Найдено архивов в '{folder_path.name}': {len(root_archives)}")

    # Обрабатываем каждый архив
    for archive in root_archives:
        try:
            extracted_folder = extract_archive_with_correct_encoding(archive)

            if extracted_folder:
                processed_folders.append(extracted_folder)

                # Удаляем исходный архив ТОЛЬКО если явно указано
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

    # Затем рекурсивно обрабатываем подпапки
    for subfolder in folder_path.iterdir():
        if subfolder.is_dir():
            # Пропускаем папки, которые мы только что создали при распаковке
            if any(subfolder.name == f.stem for f in root_archives):
                continue

            logger.debug(f"Проверка подпапки: {subfolder.name}")
            sub_folders = process_archives_in_folder(subfolder, delete_originals)
            processed_folders.extend(sub_folders)

    return processed_folders


def clean_empty_folders(folder_path):
    """
    Рекурсивно удаляет пустые папки
    """
    try:
        deleted_count = 0
        for item in folder_path.iterdir():
            if item.is_dir():
                # Рекурсивно очищаем подпапки
                deleted_count += clean_empty_folders(item)

                # Проверяем, пуста ли папка после очистки
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


# ============================================
# 2. ФУНКЦИИ ДЛЯ РАБОТЫ С EXCEL
# ============================================
def cleanup_old_print_settings(wb):
    """Удаляет старые настройки печати из книги"""
    try:
        if hasattr(wb, 'defined_names') and wb.defined_names.definedName:
            names_to_remove = []
            for name in wb.defined_names.definedName:
                if name.name in ['_xlnm.Print_Area', '_xlnm.Print_Titles']:
                    names_to_remove.append(name)

            for name in names_to_remove:
                wb.defined_names.definedName.remove(name)
            logger.info(f"Удалено старых настроек печати: {len(names_to_remove)}")
        return True
    except Exception as e:
        logger.warning(f"Не удалось очистить старые настройки: {e}")
        return False


def reset_page_setup_openpyxl(ws):
    """Сбрасывает настройки печати для листа openpyxl"""
    try:
        ws.print_area = None
        ws.print_title_rows = None
        ws.print_title_cols = None
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_setup.scale = 100
        ws.page_margins = PageMargins(
            left=0.25, right=0.25,
            top=0.5, bottom=0.75,
            header=0.25, footer=0.25
        )

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


def reset_page_setup_com(ws):
    """Сбрасывает настройки печати для листа win32com"""
    try:
        while ws.HPageBreaks.Count > 0:
            try:
                ws.HPageBreaks(1).Delete()
            except:
                break

        while ws.VPageBreaks.Count > 0:
            try:
                ws.VPageBreaks(1).Delete()
            except:
                break

        ws.PageSetup.PrintArea = ""
        ws.PageSetup.PrintTitleRows = ""
        ws.PageSetup.PrintTitleColumns = ""
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.FitToPagesTall = False

        logger.info(f"[{ws.Name}] Настройки печати сброшены (COM)")
        return True
    except Exception as e:
        logger.warning(f"Ошибка сброса COM: {e}")
        return False


def is_sheet_empty(ws):
    """Проверяет, действительно ли лист пустой"""
    try:
        if ws.max_row <= 1 and ws.max_column <= 1:
            for row in ws.iter_rows(max_row=1, max_column=1):
                for cell in row:
                    if cell.value is not None and str(cell.value).strip():
                        return False
            return True
        return False
    except:
        return True


def is_sheet_empty_com(ws):
    """Проверяет пустой ли лист в COM"""
    try:
        used = ws.UsedRange
        if used.Count == 1 and (used.Value is None or used.Value == ""):
            return True
        if hasattr(used, 'Value'):
            if isinstance(used.Value, tuple):
                for row in used.Value:
                    if isinstance(row, tuple):
                        for cell in row:
                            if cell not in (None, "", 0):
                                return False
                    elif row not in (None, "", 0):
                        return False
            else:
                if used.Value not in (None, "", 0):
                    return False
        return True
    except:
        return True


def hard_clear_fills_only(wb):
    """Убирает только заливку (фон) во всех листах книги"""
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


def find_header_rows(ws, first_data_row=1):
    """Поиск заголовков таблицы по маркерам"""
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
    """Резервный поиск заголовков по границам"""
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
    """Комбинированный умный поиск заголовков"""
    try:
        start_row, end_row = find_header_rows(ws)
        if end_row - start_row > 8:
            logger.warning(f"[{ws.title}] Заголовок слишком большой ({start_row}:{end_row}) → ограничиваем 2–4")
            start_row, end_row = 2, min(4, ws.max_row)
        return start_row, end_row
    except Exception as e:
        logger.error(f"[{ws.title}] Ошибка при комбинированном поиске: {e}")
        return 1, 1


def create_excel_instance():
    """Создаёт новый независимый экземпляр Excel"""
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    return excel


def safe_add_page_number(ws):
    """Безопасно добавляет номер страницы в нижний колонтитул"""
    try:
        if ws.oddFooter.center and "&P" in getattr(ws.oddFooter.center, 'text', ''):
            return False

        ws.oddFooter.center.text = "Страница &P из &N"
        ws.evenFooter.center.text = "Страница &P из &N"
        return True
    except Exception as e:
        logger.warning(f"Ошибка добавления номера страницы: {e}")
        return False


def calculate_printable_area(ws):
    """Рассчитывает доступную область печати в дюймах"""
    page_setup = ws.page_setup
    margins = ws.page_margins

    paper_sizes = {
        1: (8.5, 11),
        9: (8.27, 11.69),
        5: (8.5, 14),
    }

    paper_code = page_setup.paperSize or 9
    paper_width, paper_height = paper_sizes.get(paper_code, (8.27, 11.69))

    orientation = "portrait"
    if page_setup.orientation == "landscape":
        paper_width, paper_height = paper_height, paper_width
        orientation = "landscape"

    printable_width = paper_width - (margins.left + margins.right)
    printable_height = paper_height - (margins.top + margins.bottom)

    return printable_width, printable_height, orientation


def is_real_data_cell(ws, row, col):
    """Проверка - является ли ячейка реальными данными"""
    try:
        cell = ws.cell(row=row, column=col)

        if cell.value is not None:
            cell_str = str(cell.value).strip()
            if cell_str and cell_str not in ['', ' ', '\t', '\n', '\r']:
                return True

        for merged_range in ws.merged_cells.ranges:
            if merged_range.min_row <= row <= merged_range.max_row and \
                    merged_range.min_col <= col <= merged_range.max_col:
                master_cell = ws.cell(merged_range.min_row, merged_range.min_col)
                if master_cell.value is not None:
                    return True

        return False
    except Exception as e:
        logger.warning(f"Ошибка проверки ячейки: {e}")
        return False


def auto_fit_print_area_improved(ws):
    """УЛУЧШЕННАЯ функция установки области печати"""
    try:
        logger.info(f"[{ws.title}] УЛУЧШЕННЫЙ поиск реальных данных")

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
    """Агрессивно очищает пустые области с форматированием"""
    try:
        last_real_row = 1
        last_real_col = 1

        for merged_range in ws.merged_cells.ranges:
            master_cell = ws.cell(merged_range.min_row, merged_range.min_col)

            if master_cell.value is None:
                try:
                    ws.unmerge_cells(str(merged_range))
                    logger.info(f"[{ws.title}] Удалено объединение: {merged_range}")
                except Exception as e:
                    logger.warning(f"[{ws.title}] Не удалось удалить объединение: {e}")

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
                    has_border = (border.top and border.top.style) or \
                                 (border.bottom and border.bottom.style) or \
                                 (border.left and border.left.style) or \
                                 (border.right and border.right.style)

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


def set_orientation_and_scale(ws):
    """Устанавливает ориентацию страницы и масштаб печати"""
    try:
        reset_page_setup_openpyxl(ws)
        total_width, total_height = calculate_table_dimensions(ws)

        if total_width == 0 or total_height == 0:
            logger.warning(f"[{ws.title}] Не удалось рассчитать размеры")
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
            ws.page_setup.scale = 100
            return True

        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        margins = PageMargins(
            left=0.25, right=0.25,
            top=0.5, bottom=0.75,
            header=0.25, footer=0.25
        )
        ws.page_margins = margins

        available_width_portrait = PAGE_WIDTH_A4_PORTRAIT - margins.left - margins.right
        available_width_landscape = PAGE_HEIGHT_A4_PORTRAIT - margins.left - margins.right

        portrait_scale = None
        for scale in range(100, 59, -1):
            scaled_width = total_width * (scale / 100.0)
            if scaled_width <= available_width_portrait:
                portrait_scale = scale
                break

        if portrait_scale is not None:
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
            ws.page_setup.scale = portrait_scale
            logger.info(f"[{ws.title}] Портрет: масштаб {portrait_scale}%")
            return True

        landscape_scale = None
        for scale in range(100, 59, -1):
            scaled_width = total_width * (scale / 100.0)
            if scaled_width <= available_width_landscape:
                landscape_scale = scale
                break

        if landscape_scale is not None:
            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
            ws.page_setup.scale = landscape_scale
            logger.info(f"[{ws.title}] Альбомная: масштаб {landscape_scale}%")
            return True

        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.scale = 60
        logger.info(f"[{ws.title}] Форсированная альбомная (масштаб 60%)")
        return True

    except Exception as e:
        logger.error(f"Ошибка установки ориентации/масштаба: {e}")
        return False


def calculate_table_dimensions(ws):
    """Рассчитывает размеры таблицы в дюймах"""
    try:
        max_col = 1
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    max_col = max(max_col, cell.column)

        default_col_width = ws.sheet_format.defaultColWidth or 8.43
        default_row_height = ws.sheet_format.defaultRowHeight or 15.0

        total_width_inches = 0.0
        for col_idx in range(1, max_col + 1):
            col_letter = get_column_letter(col_idx)
            col_dim = ws.column_dimensions.get(col_letter)

            if col_dim and col_dim.hidden:
                continue

            if col_dim and col_dim.width is not None:
                actual_width = col_dim.width
            else:
                actual_width = default_col_width

            width_in_pixels = actual_width * 7 * CALIBRATION_FACTOR + 5
            width_in_inches = width_in_pixels / 96
            total_width_inches += width_in_inches

        total_height_points = 0.0
        for row_idx in range(1, ws.max_row + 1):
            row_dim = ws.row_dimensions.get(row_idx)

            if row_dim and row_dim.hidden:
                continue

            if row_dim and row_dim.height is not None:
                height_points = row_dim.height
            else:
                height_points = default_row_height

            total_height_points += height_points

        total_height_inches = total_height_points * EXCEL_POINT_TO_INCH

        logger.info(f"[{ws.title}] Размеры: {total_width_inches:.2f}x{total_height_inches:.2f} дюймов")
        return total_width_inches, total_height_inches

    except Exception as e:
        logger.error(f"Ошибка расчёта размеров таблицы: {e}")
        return 0, 0

# ============================================
# 3. ОСНОВНАЯ ФУНКЦИЯ ОБРАБОТКИ ФАЙЛА
# ============================================

def process_excel_file(file_path, log_func=None):
    """Основная функция обработки Excel файла"""
    temp_xlsx = None
    excel = None
    wb = None
    header_rows_map = {}

    try:
        file_path = str(file_path)
        original_path_for_modified = Path(file_path)

        if not os.path.exists(file_path):
            logger.info(f"Файл не существует: {file_path}")
            return False

        # Конвертация .xls в .xlsx если нужно
        if file_path.lower().endswith('.xls'):
            temp_xlsx = file_path + 'x'
            try:
                logger.info(f"Конвертируем {file_path} -> {temp_xlsx}")
                excel = create_excel_instance()
                excel.DisplayAlerts = False
                wb_com = excel.Workbooks.Open(file_path)
                wb_com.SaveAs(temp_xlsx, FileFormat=51)
                wb_com.Close()
                excel.Quit()
                excel = None
                file_path_for_openpyxl = temp_xlsx
            except Exception as e:
                logger.info(f"Ошибка конвертации: {e}")
                if excel:
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
            # Исправляем вызов функции (у openpyxl другая структура)
            try:
                if wb.defined_names:
                    # Удаляем старые Print_Area и Print_Titles
                    names_to_remove = []
                    for name, def_name in wb.defined_names.items():
                        if name in ['Print_Area', 'Print_Titles', '_xlnm.Print_Area', '_xlnm.Print_Titles']:
                            names_to_remove.append(name)

                    for name in names_to_remove:
                        del wb.defined_names[name]
                    logger.info(f"Удалено старых настроек печати: {len(names_to_remove)}")
            except Exception as e:
                logger.warning(f"Не удалось очистить старые настройки: {e}")

            hard_clear_fills_only(wb)

            modified_path = get_modified_path(
                original_path_for_modified,
                os.path.dirname(original_path_for_modified)
            ).with_suffix('.xlsx')
            modified_path.parent.mkdir(parents=True, exist_ok=True)
            wb.save(str(modified_path))
            logger.info(f"Сохранено после openpyxl обработки: {modified_path}")

        except Exception as e:
            logger.info(f"Ошибка загрузки/очистки: {e}")
            return False

        # Определяем заголовки и настраиваем печать
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
                logger.info(f"[{ws.title}] Сквозные строки: {start_row}-{end_row}, столбец A")

                if safe_add_page_number(ws):
                    logger.info(f"[{ws.title}] Добавлен номер страницы")

                aggressive_clean_empty_regions(ws)
                auto_fit_print_area_improved(ws)

                if set_orientation_and_scale(ws):
                    logger.info(f"[{ws.title}] Установлена ориентация и масштаб")

                calculate_printable_area(ws)
                logger.info(f"[{ws.title}] Автоматически установлена область печати по содержимому")

            except Exception as e:
                logger.info(f"Ошибка обработки листа {ws.title}: {e}")
                continue

        wb.save(str(modified_path))
        logger.info(f"Сохранено как .xlsx: {modified_path}")

        # === COM ОБРАБОТКА ===
        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            wb_com = excel.Workbooks.Open(str(modified_path))

            for ws in wb_com.Sheets:
                try:
                    if is_sheet_empty_com(ws):
                        logger.info(f"[{ws.Name}] Пустой лист (COM) - пропускаем")
                        continue

                    ws.Cells.FormatConditions.Delete()
                    ws.Cells.Interior.Pattern = -4142
                    ws.Cells.Interior.TintAndShade = 0

                    if hasattr(ws, "ListObjects") and ws.ListObjects.Count > 0:
                        for tbl in ws.ListObjects:
                            try:
                                tbl.TableStyle = ""
                            except Exception:
                                pass

                    # ★★★★АВТО-ПОДГОНКА С WRAP ★★★★
                    try:
                        smart_autofit_rows_exact(ws)
                    except Exception as e_fit:
                        logger.warning(f"[{ws.Name}] Ошибка smart_autofit_rows_exact: {e_fit}")
                    # ★★★★★★★★★★★★★★★★★★★★★★★★

                    logger.info(f"[{ws.Name}] Заливка очищена через COM")

                except Exception as e:
                    logger.warning(f"[{ws.Name}] Ошибка при очистке через COM: {e}")

            for ws in wb_com.Sheets:
                try:
                    if is_sheet_empty_com(ws):
                        continue

                    reset_page_setup_com(ws)

                    if ws.Name in header_rows_map:
                        start_row, end_row = header_rows_map[ws.Name]
                        ws.PageSetup.PrintTitleRows = f"${start_row}:${end_row}"
                        ws.PageSetup.PrintTitleColumns = "$A:$A"
                        logger.info(f"[{ws.Name}] Сквозные строки установлены в Win32")

                    ws.PageSetup.Order = 2

                except Exception as e:
                    logger.info(f"[{ws.Name}] Ошибка Win32 настроек: {e}")

            wb_com.Save()
            wb_com.Close()

        except Exception as e:
            logger.info(f"Win32COM ошибка: {e}")
        finally:
            if excel:
                try:
                    excel.Quit()
                except:
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
            except (PermissionError, FileNotFoundError, OSError) as e:
                logger.info(f"Не удалось удалить временный файл: {e}")


# ============================================
# 4. ГРАФИЧЕСКИЙ ИНТЕРФЕЙС
# ============================================

def get_modified_path(original_path, root_folder):
    """Сохраняет структуру папок в modified/"""
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


class ExcelProcessor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.folder_path = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar(value="Готово")
        self.log_messages = []
        self.delete_archives_var = tk.BooleanVar(value=False)  # НЕ удалять архивы по умолчанию

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

        # Checkbox для удаления архивов (по умолчанию выключен)
        delete_frame = ttk.Frame(main_frame)
        delete_frame.pack(pady=5)
        ttk.Checkbutton(delete_frame, text="Удалить архивы после распаковки",
                        variable=self.delete_archives_var).pack(side=tk.LEFT)

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

        # Лог-окно
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
        """Выводит сообщение в GUI лог"""
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
        """Только распаковка архивов"""
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
                self.log_to_gui(f"Распаковка завершена!")
                self.log_to_gui(f"Создано папок: {len(processed_folders)}")
                for i, folder in enumerate(processed_folders[:10], 1):
                    self.log_to_gui(f"   {i}. {folder.relative_to(folder_path)}")
                if len(processed_folders) > 10:
                    self.log_to_gui(f"   ... и еще {len(processed_folders) - 10} папок")

                delete_status = "УДАЛЕНЫ" if self.delete_archives_var.get() else "СОХРАНЕНЫ"
                messagebox.showinfo("Готово",
                                    f"Распаковано {len(processed_folders)} архивов\n"
                                    f"Оригинальные архивы: {delete_status}\n"
                                    f"Структура папок сохранена!")
            else:
                self.log_to_gui("Архивы не найдены")
                messagebox.showinfo("Готово", "Архивы не найдены")

        except Exception as e:
            self.log_to_gui(f"Ошибка распаковки: {str(e)}")
            messagebox.showerror("Ошибка", f"Ошибка распаковки:\n{str(e)}")

    def start_processing(self):
        """Полная обработка: архивы + Excel"""
        folder = self.folder_path.get()
        if not folder:
            messagebox.showerror("Ошибка", "Сначала выберите папку!")
            return

        self.log_to_gui("=" * 60)
        self.log_to_gui("НАЧАЛО ПОЛНОЙ ОБРАБОТКИ")
        self.log_to_gui("=" * 60)

        # Запуск в отдельном потоке
        import threading
        threading.Thread(target=self._process_all, args=(folder,), daemon=True).start()

    def _process_all(self, folder):
        """Внутренняя функция для обработки в отдельном потоке"""
        try:
            pythoncom.CoInitialize()

            folder_path = Path(folder)

            # Шаг 1: Распаковка архивов (НЕ удаляем оригиналы)
            self.status_var.set("Распаковка архивов...")
            self.progress_var.set(20)
            self.log_to_gui("Этап 1: Распаковка архивов")
            self.root.update()

            processed_folders = process_archives_in_folder(folder_path, self.delete_archives_var.get())

            if processed_folders:
                self.log_to_gui(f"Распаковано архивов: {len(processed_folders)}")
                if self.delete_archives_var.get():
                    self.log_to_gui("Оригинальные архивы УДАЛЕНЫ")
                else:
                    self.log_to_gui("Оригинальные архивы СОХРАНЕНЫ")

                deleted_count = clean_empty_folders(folder_path)
                if deleted_count > 0:
                    self.log_to_gui(f"Удалено {deleted_count} пустых папок")
            else:
                self.log_to_gui("Архивы не найдены")

            # Шаг 2: Обработка Excel файлов
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

            self.log_to_gui(f"Найдено Excel файлов: {total_files}")

            success_count = 0
            error_count = 0
            error_log = []

            # Обработка файлов
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
                    result = process_excel_file(str(file))
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

            # Сохранение лога ошибок
            if error_log:
                log_dir = folder_path / "modified" / "logs"
                log_dir.mkdir(parents=True, exist_ok=True)
                log_file = log_dir / "error_log.txt"
                with open(log_file, 'w', encoding='utf-8') as f:
                    f.write("Файлы, при обработке которых произошли ошибки:\n")
                    f.write("\n".join(error_log))
                self.log_to_gui(f"Лог ошибок сохранен: {log_file}")

            # Итог
            self.status_var.set("Готово")
            self.progress_var.set(100)
            self.log_to_gui("\n" + "=" * 60)
            self.log_to_gui("ОБРАБОТКА ЗАВЕРШЕНА")
            self.log_to_gui(f"Успешно: {success_count}")
            self.log_to_gui(f"Ошибки: {error_count}")
            self.log_to_gui("=" * 60)

            messagebox.showinfo("Завершено",
                                f"Обработка завершена!\n\n"
                                f"Успешно: {success_count}\n"
                                f"Ошибки: {error_count}\n\n"
                                f"Обработанные файлы сохранены в папке 'modified'")

        except Exception as ex:
            self.log_to_gui(f"Критическая ошибка: {str(ex)}")
            messagebox.showerror("Ошибка", f"Процесс прерван:\n{str(ex)}")
            traceback.print_exc()
        finally:
            self.status_var.set("Готово")
            self.progress_var.set(0)
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    def on_close(self):
        if messagebox.askokcancel("Выход", "Вы действительно хотите выйти?"):
            self.root.destroy()

    def run(self):
        self.root.mainloop()


# ============================================
# 5. ЗАПУСК ПРОГРАММЫ
# ============================================

if __name__ == "__main__":
    try:
        app = ExcelProcessor()
        app.run()
    except KeyboardInterrupt:
        print("Остановка по инициативе пользователя.")
    except Exception as e:
        print(f"Произошла непредвиденная ошибка: {e}")
        traceback.print_exc()
