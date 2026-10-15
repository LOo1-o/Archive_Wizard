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

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

try:
    import rarfile
except ImportError:
    rarfile = None

import pythoncom

import win32com.client as win32
from openpyxl.worksheet.page import PageMargins
from openpyxl.styles import PatternFill

# Настройка логгера
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

# Константы для формата бумаги A4 в дюймах
# PAGE_WIDTH_A4_PORTRAIT = 6.5  # Ширина A4 в дюймах (портрет)
# PAGE_HEIGHT_A4_PORTRAIT = 11.69  # Высота A4 в дюймах (портрет)

PAGE_WIDTH_A4_PORTRAIT = 8.27  # Ширина A4 в дюймах (портрет)
PAGE_HEIGHT_A4_PORTRAIT = 11.69  # Высота A4 в дюймах (портрет)

# Коэффициенты
EXCEL_POINT_TO_INCH = 1 / 72
CALIBRATION_FACTOR = 1.02
MAX_FILES_TO_PROCESS = 1000


def debug_border_rows(ws, max_rows=50, max_cols=40):
    rows_with_borders = []
    for r in range(1, min(max_rows, ws.max_row) + 1):
        cols = []
        for c in range(1, min(max_cols, ws.max_column) + 1):
            b = ws.cell(row=r, column=c).border
            if any([
                getattr(b, 'top', None) and getattr(b.top, 'style', None),
                getattr(b, 'bottom', None) and getattr(b.bottom, 'style', None),
                getattr(b, 'left', None) and getattr(b.left, 'style', None),
                getattr(b, 'right', None) and getattr(b.right, 'style', None)
            ]):
                cols.append(c)
        if cols:
            rows_with_borders.append((r, cols))
    print(f"Лист {ws.title} — строки с границами (первые {min(max_rows, ws.max_row)}):")
    for r, cols in rows_with_borders:
        print(f"  row {r}: cols with border = {cols}")


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

def log_com_pagesetup(ws):
    try:
        ps = ws.PageSetup
        logger.info(
            f"[{ws.Name}] COM PageSetup: Zoom={getattr(ps, 'Zoom', None)}, FitToPagesWide={getattr(ps, 'FitToPagesWide', None)}, FitToPagesTall={getattr(ps, 'FitToPagesTall', None)}, Orientation={getattr(ps, 'Orientation', None)}, PaperSize={getattr(ps, 'PaperSize', None)}")
    except Exception:
        pass


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
        ws.PageSetup.FitToPagesWide = False
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


# Поиск сквозных строк
def find_header_rows(ws, first_data_row=1):
    """Улучшенный поиск заголовков: приоритет границам, игнорирование маркеров в данных"""
    try:
        # Шаг 1: Ищем заголовок по границам (основной метод)
        start_by_border, end_by_border = find_header_rows_by_borders(ws, first_data_row)

        # Если нашли разумный заголовок по границам (не слишком большой)
        if end_by_border - start_by_border <= 8 and end_by_border >= start_by_border:
            logger.info(f"[{ws.title}] ✅ Заголовок найден по границам: строки {start_by_border}-{end_by_border}")
            return start_by_border, end_by_border

        # Шаг 2: Если не нашли по границам, ищем по маркерам
        marker_row = None

        # Поиск маркеров, НО ИГНОРИРУЕМ строки, которые уже входят в найденный по границам блок
        ignore_above = end_by_border if end_by_border > 1 else 0

        for col in [1, 2, 3]:
            for row in range(max(first_data_row, ignore_above + 1),
                             min(first_data_row + 25, ws.max_row + 1)):
                cell_value = ws.cell(row=row, column=col).value
                cell_str = str(cell_value).strip().lower() if cell_value else ""

                # Маркеры заголовка
                if cell_str in ['a', 'b', 'c', 'а', 'б', 'в', 'п/п', '№']:
                    marker_row = row
                    break
                # Осторожно с цифрами 1,2,3 - они могут быть и в данных!
                elif re.fullmatch(r'^[123]$', cell_str):
                    # Проверяем, есть ли границы в этой строке
                    has_border = False
                    for check_col in range(1, min(5, ws.max_column + 1)):
                        border = ws.cell(row=row, column=check_col).border
                        if (border.top and border.top.style) or \
                                (border.bottom and border.bottom.style):
                            has_border = True
                            break

                    # Если есть границы - считаем маркером заголовка
                    if has_border:
                        marker_row = row
                        break
                    # Иначе - это, скорее всего, данные, пропускаем
            if marker_row is not None:
                break

        if marker_row is None:
            # Если ничего не нашли, используем разумный дефолт
            if end_by_border > 1:
                logger.info(f"[{ws.title}] Маркер не найден, используем заголовок по границам")
                return start_by_border, end_by_border
            else:
                logger.info(f"[{ws.title}] Ничего не найдено → используем первую строку")
                return 1, 1

        # Шаг 3: Ищем блок заголовков выше маркера
        header_block = [marker_row]
        current_row = marker_row - 1

        while current_row >= 1:
            has_any_border = False
            for col in range(1, min(12, ws.max_column + 1)):
                border = ws.cell(row=current_row, column=col).border
                # Ищем ЛЮБЫЕ границы (не только правые!)
                if (border.top and border.top.style) or \
                        (border.bottom and border.bottom.style) or \
                        (border.left and border.left.style) or \
                        (border.right and border.right.style):
                    has_any_border = True
                    break

            if has_any_border:
                header_block.append(current_row)
                current_row -= 1
            else:
                break

        start_row = min(header_block)
        end_row = max(header_block)

        # Ограничиваем размер заголовка
        if end_row - start_row > 6:
            logger.warning(f"[{ws.title}] Заголовок слишком большой ({start_row}:{end_row}) → обрезаем")
            end_row = min(start_row + 3, ws.max_row)

        logger.info(f"[{ws.title}] ✅ Заголовок найден: строки {start_row}-{end_row}")
        return start_row, end_row

    except Exception as e:
        logger.error(f"[{ws.title}] Ошибка при поиске заголовков: {e}")
        return 1, 1


def find_header_rows_by_borders(ws, first_data_row=1, max_search_rows=30):
    """Поиск заголовков ТОЛЬКО по границам"""
    try:
        border_rows = []

        # Сканируем строки на наличие границ
        for row in range(first_data_row, min(first_data_row + max_search_rows, ws.max_row + 1)):
            for col in range(1, min(15, ws.max_column + 1)):
                border = ws.cell(row=row, column=col).border

                # Заголовки обычно имеют ВЕРХНИЕ или НИЖНИЕ границы
                if (border.top and border.top.style) or \
                        (border.bottom and border.bottom.style):
                    border_rows.append(row)
                    break

        if not border_rows:
            return 1, 1

        # Находим самый первый блок подряд идущих строк с границами
        blocks = []
        current_block = [border_rows[0]]

        for i in range(1, len(border_rows)):
            if border_rows[i] == border_rows[i - 1] + 1:
                current_block.append(border_rows[i])
            else:
                blocks.append(current_block)
                current_block = [border_rows[i]]

        blocks.append(current_block)

        if blocks:
            # Берём первый (самый верхний) блок
            first_block = blocks[0]
            start_row = min(first_block)
            end_row = max(first_block)

            # Ограничение размера
            if end_row - start_row > 6:
                end_row = start_row + 2

            return start_row, end_row

        return 1, 1

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


def find_header_rows_any_border_advanced(ws, max_search_rows=30):
    """Резервный поиск заголовков по границам (рамкам)"""
    try:
        all_border_rows = []
        for row in range(1, min(max_search_rows, ws.max_row + 1)):
            for col in range(1, min(20, ws.max_column + 1)):
                border = ws.cell(row=row, column=col).border
                if any([
                    border.top and border.top.style,
                    border.bottom and border.bottom.style,
                    border.right and border.right.style,
                    border.left and border.left.style
                ]):
                    all_border_rows.append(row)
                    break

        if not all_border_rows:
            logger.info(f"[{ws.title}] ⚠️ Границы не найдены — используем первую строку.")
            return 1, 1

        # --- Группируем смежные строки с границами ---
        blocks = []
        current_block = [all_border_rows[0]]
        for i in range(1, len(all_border_rows)):
            if all_border_rows[i] == all_border_rows[i - 1] + 1:
                current_block.append(all_border_rows[i])
            else:
                blocks.append(current_block)
                current_block = [all_border_rows[i]]
        blocks.append(current_block)

        # --- Берем первую "плотную" группу как заголовок ---
        first_block = blocks[0]
        start_row, end_row = min(first_block), max(first_block)

        logger.info(f"[{ws.title}] 🧩 Заголовок найден по границам: строки {start_row}-{end_row}")
        return start_row, end_row

    except Exception as e:
        logger.error(f"[{ws.title}] Ошибка поиска по границам: {e}")
        return 1, 1


# ____окончание блока --> поиск сквозных строк

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
    """Рассчитывает доступную область печати в дюймах с учетом ориентации"""
    # 1. Получаем настройки страницы
    page_setup = ws.page_setup
    margins = ws.page_margins

    # 2. Определяем размер бумаги (в дюймах)
    paper_sizes = {
        1: (8.5, 11),  # Letter
        9: (8.27, 11.69),  # A4
        5: (8.5, 14),  # Legal
        # Добавьте другие форматы при необходимости
    }

    paper_code = page_setup.paperSize or 9  # По умолчанию A4
    paper_width, paper_height = paper_sizes.get(paper_code, (8.27, 11.69))

    # 3. Учитываем ориентацию
    orientation = "portrait"
    if page_setup.orientation == "landscape":
        paper_width, paper_height = paper_height, paper_width
        orientation = "landscape"

    # 4. Рассчитываем доступную область
    printable_width = paper_width - (margins.left + margins.right)
    printable_height = paper_height - (margins.top + margins.bottom)

    return printable_width, printable_height, orientation


def calculate_table_dimensions(ws):
    """Рассчитывает размеры таблицы в дюймах"""
    try:
        # Определяем последний непустой столбец
        max_col = 1
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    max_col = max(max_col, cell.column)

        # Стандартные размеры
        default_col_width = ws.sheet_format.defaultColWidth or 8.43
        default_row_height = ws.sheet_format.defaultRowHeight or 15.0

        # Рассчитываем ширину таблицы
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

        # Рассчитываем высоту таблицы
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


#________________
def set_orientation_and_scale_openpyxl(ws):
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

        # Устанавливаем бумагу A4
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        margins = PageMargins(left=0.25, right=0.25, top=0.5, bottom=0.75, header=0.25, footer=0.25)
        ws.page_margins = margins

        available_width_portrait = PAGE_WIDTH_A4_PORTRAIT - (margins.left + margins.right)
        available_width_landscape = PAGE_HEIGHT_A4_PORTRAIT - (margins.left + margins.right)

        logger.info(
            f"[{ws.title}] total_width={total_width:.2f}\" available_portrait={available_width_portrait:.2f}\" available_landscape={available_width_landscape:.2f}\"")

        # ===== ВСЕГДА сначала пробуем портретную! =====

        # 1. Проверяем, помещается ли в портретную при 100%
        if total_width <= available_width_portrait:
            try:
                ws.page_setup.orientation = 'portrait'
                ws.page_setup.fitToWidth = False
                ws.page_setup.fitToHeight = False
                ws.page_setup.scale = 100
                logger.info(f"[{ws.title}] ✅ Помещается в портретную при 100% -> ВСЕГДА выбираем портретную")
                return 100
            except Exception:
                pass

        # 2. Если не помещается при 100%, пробуем уменьшить масштаб для портретной
        portrait_scale = None
        for scale in range(100, 59, -1):  # от 100% до 60%
            if total_width * (scale / 100.0) <= available_width_portrait:
                portrait_scale = scale
                break

        if portrait_scale is not None:
            try:
                ws.page_setup.orientation = 'portrait'
                ws.page_setup.fitToWidth = False
                ws.page_setup.fitToHeight = False
                ws.page_setup.scale = max(60, portrait_scale)
                logger.info(f"[{ws.title}] ✅ Выбрана портретная ориентация (масштаб {ws.page_setup.scale}%)")
                return clamp_zoom(ws.page_setup.scale)
            except Exception:
                pass

        # ===== ТОЛЬКО если не помещается в портретную даже при 60% =====

        # 3. Проверяем, помещается ли в альбомную при 100%
        if total_width <= available_width_landscape:
            try:
                ws.page_setup.orientation = 'landscape'
                ws.page_setup.fitToWidth = False
                ws.page_setup.fitToHeight = False
                ws.page_setup.scale = 100
                logger.info(f"[{ws.title}] ❌ В портретную не помещается -> выбрана альбомная при 100%")
                return 100
            except Exception:
                pass

        # 4. Если не помещается в альбомную при 100%, уменьшаем масштаб
        landscape_scale = None
        for scale in range(100, 59, -1):
            if total_width * (scale / 100.0) <= available_width_landscape:
                landscape_scale = scale
                break

        if landscape_scale is not None:
            try:
                ws.page_setup.orientation = 'landscape'
                ws.page_setup.fitToWidth = False
                ws.page_setup.fitToHeight = False
                ws.page_setup.scale = max(60, landscape_scale)
                logger.info(
                    f"[{ws.title}] ❌ В портретную не помещается -> выбрана альбомная (масштаб {ws.page_setup.scale}%)")
                return clamp_zoom(ws.page_setup.scale)
            except Exception:
                pass

        # 5. Если ничего не подошло - форсируем альбомную с минимальным масштабом
        try:
            ws.page_setup.orientation = 'landscape'
            ws.page_setup.fitToWidth = False
            ws.page_setup.fitToHeight = False
            ws.page_setup.scale = 60
            logger.warning(f"[{ws.title}] ⚠️ Форсированная альбомная ориентация (scale=60%)")
            return 60
        except Exception:
            pass

        return None
    except Exception as e:
        logger.error(f"[{ws.title}] Ошибка установки ориентации/масштаба (openpyxl): {e}", exc_info=True)
        return None

def set_orientation_and_scale(ws):
    """
    Устанавливает ориентацию страницы и масштаб печати
    """
    try:
        reset_page_setup_openpyxl(ws)

        # Рассчитываем размеры таблицы
        total_width, total_height = calculate_table_dimensions(ws)

        if total_width == 0 or total_height == 0:
            logger.warning(f"[{ws.title}] Не удалось рассчитать размеры")
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
            ws.page_setup.scale = 100
            return True

        # Устанавливаем формат A4
        ws.page_setup.paperSize = ws.PAPERSIZE_A4

        # Настройка полей страницы
        margins = PageMargins(
            left=0.25, right=0.25,
            top=0.5, bottom=0.75,
            header=0.25, footer=0.25
        )
        ws.page_margins = margins

        # Рассчитываем доступную область печати
        available_width_portrait = PAGE_WIDTH_A4_PORTRAIT - margins.left - margins.right
        available_width_landscape = PAGE_HEIGHT_A4_PORTRAIT - margins.left - margins.right

        # Пробуем портретную ориентацию
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

        # Пробуем альбомную ориентацию
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

        # Форсируем альбомную ориентацию
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.scale = 60
        logger.info(f"[{ws.title}] Форсированная альбомная (масштаб 60%)")
        return True

    except Exception as e:
        logger.error(f"Ошибка установки ориентации/масштаба: {e}")
        return False
#_________________

def diagnose_scale(xlsx_path):
    wb = load_workbook(xlsx_path)
    for ws in wb.worksheets:
        total_width = calculate_table_dimensions(ws)
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
        print("Текущие настройки PageSetup:")
        print(f"  orientation: {ps.orientation}")
        print(f"  scale: {ps.scale}")
        print(f"  fitToWidth: {ps.fitToWidth}")
        print(f"  fitToHeight: {ps.fitToHeight}")


# ============================================
# 3. ОСНОВНАЯ ФУНКЦИЯ ОБРАБОТКИ ФАЙЛА
# ============================================
def create_excel_instance(visible=False):
    """
    Создаёт отдельный экземпляр Excel (DispatchEx) и настраивает его.
    visible: показывать ли окно Excel (обычно False).
    """
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = visible
    excel.DisplayAlerts = False
    # Можно дополнительно настроить: excel.ScreenUpdating = False
    return excel


def process_excel_file(file_path, log_func=None, excel_app=None):
    """
    Основная функция обработки Excel файла.
    Если передан excel_app, он будет использован для всех COM-операций (рекомендовано).
    Если excel_app is None — создаётся временный Excel и корректно закрывается.
    Возвращает True/False.
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

        # Создаём локальный Excel только если не передан внешний
        if excel_app is None:
            excel = create_excel_instance(visible=False)
            local_excel = True
        else:
            excel = excel_app

        # Конвертация .xls -> .xlsx если нужно (через переданный/локальный excel)
        if file_path.lower().endswith('.xls'):
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
                # не закрываем excel здесь, т.к. мы используем общий/локальный excel
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

            # Корректная очистка имен печати, если есть
            try:
                # wb.defined_names — mapping-like в новых версиях
                names_to_remove = []
                for name in list(
                        wb.defined_names.definedName if hasattr(wb.defined_names, 'definedName') else wb.defined_names):
                    # Здесь safe-check для нескольких API-версий
                    # В простейшем виде просто игнорируем/логируем
                    pass
            except Exception:
                pass

            # Удаляем только заливки (твоя существующая функция)
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
            # если локальный excel был создан, корректно закроем
            if local_excel:
                try:
                    excel.Quit()
                except:
                    pass
            return False

        # Обработка openpyxl-микроизменений (заголовки, печать и т.д.)
        try:
            for ws in wb.worksheets:
                try:
                    if is_sheet_empty(ws):
                        logger.info(f"[{ws.title}] Пустой лист - пропускаем")
                        continue

                    # _____проверка где есть границы без маркера
                    debug_border_rows(ws)
                    # _____

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

                    if set_orientation_and_scale_openpyxl(ws):
                        logger.info(f"[{ws.title}] Установлена ориентация и масштаб")

                except Exception as e:
                    logger.info(f"Ошибка обработки листа {ws.title}: {e}")
                    continue

            wb.save(str(modified_path))
            logger.info(f"Сохранено как .xlsx: {modified_path}")

        except Exception as e:
            logger.info(f"Ошибка при post-openpyxl обработке: {e}")
            if local_excel:
                try:
                    excel.Quit()
                except:
                    pass
            return False

        # === COM ОБРАБОТКА (через тот же excel экземпляр) ===
        wb_com = None
        try:
            wb_com = excel.Workbooks.Open(str(modified_path))

            MIN_ZOOM = 60  # минимально допустимый масштаб (ваше требование)

            # Этап 1: очистка заливки/условного форматирования, сброс TableStyle
            for ws in wb_com.Sheets:
                try:
                    if is_sheet_empty_com(ws):
                        logger.info(f"[{ws.Name}] Пустой лист (COM) - пропускаем")
                        continue

                    try:
                        ws.Cells.FormatConditions.Delete()
                    except Exception:
                        pass
                    try:
                        ws.Cells.Interior.Pattern = -4142  # xlNone
                        ws.Cells.Interior.TintAndShade = 0
                    except Exception:
                        pass

                    # Сбрасываем стиль таблиц, если есть
                    try:
                        if hasattr(ws, "ListObjects") and ws.ListObjects.Count > 0:
                            for idx in range(ws.ListObjects.Count, 0, -1):
                                try:
                                    ws.ListObjects(idx).TableStyle = ""
                                except Exception:
                                    pass
                    except Exception:
                        pass

                    logger.info(f"[{ws.Name}] Заливка очищена через COM")

                except Exception as e:
                    logger.warning(f"[{ws.Name}] Ошибка при очистке через COM: {e}")

            # Этап 2: установка PageSetup/PrintTitles/Order и принудительная COM-диагностика/Zoom
            for ws in wb_com.Sheets:
                try:
                    if is_sheet_empty_com(ws):
                        continue

                    # Получаем rec_scale из openpyxl (если есть)
                    rec_scale = None
                    try:
                        if wb and ws.Name in wb.sheetnames:
                            rec_scale = getattr(wb[ws.Name].page_setup, "scale", None)
                    except Exception:
                        rec_scale = None

                    # Всегда приводим COM-настройки в порядок перед диагностикой
                    try:
                        ws.PageSetup.FitToPagesWide = 0
                        ws.PageSetup.FitToPagesTall = 0
                    except Exception:
                        pass

                    # Устанавливаем сквозные строки, если есть
                    if ws.Name in header_rows_map:
                        start_row, end_row = header_rows_map[ws.Name]
                        try:
                            ws.PageSetup.PrintTitleRows = f"${start_row}:${end_row}"
                            ws.PageSetup.PrintTitleColumns = "$A:$A"
                            logger.info(f"[{ws.Name}] Сквозные строки установлены в COM {start_row}:{end_row}")
                        except Exception:
                            pass

                    # ---- DIAGNOSTICS & расчет final_zoom на основе COM ----
                    used_width_pts = used_height_pts = 0.0
                    try:
                        used = ws.UsedRange
                        used_width_pts = float(getattr(used, "Width", 0.0) or 0.0)
                        used_height_pts = float(getattr(used, "Height", 0.0) or 0.0)
                        logger.info(
                            f"[{ws.Name}] UsedRange.Width pts = {used_width_pts:.2f}, Height pts = {used_height_pts:.2f}")
                    except Exception:
                        used_width_pts = used_height_pts = 0.0

                    # printable area в pts (учитываем ориентацию)
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

                    # Сумма ширин колонок через COM (fallback на UsedRange.Width)
                    total_cols_width_pts = 0.0
                    try:
                        col_min = 1
                        col_max = None
                        pa = ""
                        try:
                            pa = ws.PrintArea
                        except Exception:
                            pa = ""
                        if pa:
                            rng = pa.split('!')[-1].replace('$', '')
                            if ':' in rng:
                                a, b = rng.split(':')
                                try:
                                    from openpyxl.utils import column_index_from_string
                                    col_min = column_index_from_string(re.sub(r'\d', '', a))
                                    col_max = column_index_from_string(re.sub(r'\d', '', b))
                                except Exception:
                                    col_min = 1
                                    col_max = None
                        if col_max is None:
                            try:
                                col_max = int(ws.UsedRange.Columns.Count)
                            except Exception:
                                try:
                                    col_max = int(ws.Cells.Columns.Count)
                                except Exception:
                                    col_max = None
                        if col_max is not None:
                            for ci in range(col_min, col_max + 1):
                                cl = get_column_letter(ci)
                                try:
                                    w = float(ws.Range(f"{cl}:{cl}").Width or 0.0)
                                except Exception:
                                    try:
                                        w = float(ws.Columns(ci).ColumnWidth or 0.0)
                                    except Exception:
                                        w = 0.0
                                total_cols_width_pts += w
                    except Exception as e:
                        logger.warning(f"[{ws.Name}] Не удалось посчитать ширину колонок в COM: {e}")

                    logger.info(f"[{ws.Name}] total_cols_width_pts = {total_cols_width_pts:.2f}")

                    width_for_calc = total_cols_width_pts if total_cols_width_pts > 0 else used_width_pts

                    # Вычисляем кандидатов zoom (ширина/высота)
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
                        # не опускаем ниже MIN_ZOOM
                        computed_zoom = max(MIN_ZOOM, min(100, int(computed_zoom)))

                    # Логика выбора final_zoom:
                    # 1) Если openpyxl дал rec_scale и он >= MIN_ZOOM — использовать его (в т.ч. если rec_scale > computed)
                    # 2) Иначе, использовать computed_zoom (уже не ниже MIN_ZOOM)
                    # 3) В крайнем случае — MIN_ZOOM
                    if rec_scale and isinstance(rec_scale, (int, float)) and int(rec_scale) >= MIN_ZOOM:
                        final_zoom = max(MIN_ZOOM, min(100, int(rec_scale)))
                    elif computed_zoom:
                        final_zoom = computed_zoom
                    elif rec_scale and isinstance(rec_scale, (int, float)):
                        final_zoom = max(MIN_ZOOM, int(rec_scale))
                    else:
                        final_zoom = MIN_ZOOM

                    # Применяем final_zoom через COM (гарантированно)
                    try:
                        ws.PageSetup.FitToPagesWide = 0
                        ws.PageSetup.FitToPagesTall = 0
                    except Exception:
                        pass
                    try:
                        ws.PageSetup.Zoom = int(final_zoom)
                        logger.info(
                            f"[{ws.Name}] COM: установлен Zoom={final_zoom}% (computed={computed_zoom}, rec_scale={rec_scale})")
                    except Exception as e:
                        logger.warning(f"[{ws.Name}] Не удалось установить Zoom через COM: {e}")

                    # Лог текущего состояния PageSetup в COM
                    try:
                        log_com_pagesetup(ws)
                    except:
                        pass

                    # Порядок печати
                    try:
                        ws.PageSetup.Order = 2
                    except:
                        pass

                except Exception as e:
                    logger.info(f"[{ws.Name}] Ошибка Win32 настроек: {e}")

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
        # Закрываем openpyxl workbook
        if wb:
            try:
                wb.close()
            except:
                pass
        # Удаляем временный файл конвертации
        if temp_xlsx and os.path.exists(temp_xlsx):
            try:
                os.remove(temp_xlsx)
                logger.info(f"Удалён временный файл: {temp_xlsx}")
            except Exception as e:
                logger.info(f"Не удалось удалить временный файл: {e}")
        # Если excel был создан локально — корректно закрываем
        if local_excel and 'excel' in locals():
            try:
                excel.Quit()
            except:
                pass


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
        """Внутренняя функция для обработки в отдельном потоке (с единым COM Excel)"""
        excel = None
        try:
            pythoncom.CoInitialize()

            folder_path = Path(folder)

            # Шаг 1: Распаковка архивов
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

            # Шаг 2: Поиск Excel файлов
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

            # Создаём один Excel COM для всех файлов (в этом потоке)
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
                    # Передаём excel (один экземпляр) в функцию обработки
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

            # Сохраняем лог ошибок
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
            # Закрываем экземпляр Excel, созданный для всего прогона
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


# ============================================
# 5. ЗАПУСК ПРОГРАММЫ
# ============================================

if __name__ == "__main__":
    try:
        logger.info("=== Запуск скрипта обработки Excel ===")
        app = ExcelProcessor()
        app.run()
    except KeyboardInterrupt:
        logger.info("Остановка по инициативе пользователя.")
    except Exception as e:
        logger.error(f"Произошла непредвиденная ошибка: {e}")
        traceback.print_exc()
