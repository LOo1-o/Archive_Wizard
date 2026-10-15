import logging
import time
import tkinter as tk
import traceback
import zipfile
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pythoncom
import win32com.client as win32
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

try:
    import rarfile
except ImportError:
    rarfile = None
# Настройка логгера
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

PAGE_WIDTH_A4_PORTRAIT = 6.5
PAGE_HEIGHT_A4 = 11.69

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


def find_header_rows_any_border_advanced(ws, max_rows=30):
    header_rows = []

    for row in range(1, min(ws.max_row, max_rows) + 1):
        bordered_cells = 0
        for col in range(1, min(ws.max_column, 20) + 1):
            cell = ws.cell(row=row, column=col)
            b = cell.border
            if (
                    b.top.style or b.bottom.style or
                    b.left.style or b.right.style
            ):
                bordered_cells += 1

        if bordered_cells >= 2:
            header_rows.append(row)
        elif header_rows:
            break

    if header_rows:
        return min(header_rows), max(header_rows)

    return 1, 1


def calculate_table_dimensions(ws):
    default_col_width = ws.sheet_format.defaultColWidth or 8.43
    default_row_height = ws.sheet_format.defaultRowHeight or 15

    total_width_inches = 0

    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        dim = ws.column_dimensions.get(col_letter)

        if dim and dim.hidden:
            continue

        width = dim.width if dim and dim.width else default_col_width
        pixels = width * 7 * CALIBRATION_FACTOR + 5
        total_width_inches += pixels / 96

    total_height_points = 0
    for row_idx in range(1, ws.max_row + 1):
        dim = ws.row_dimensions.get(row_idx)
        if dim and dim.hidden:
            continue
        total_height_points += dim.height if dim and dim.height else default_row_height

    return total_width_inches, total_height_points * EXCEL_POINT_TO_INCH


def set_orientation_and_scale(ws):
    ws.page_setup.paperSize = ws.PAPERSIZE_A4

    margins = PageMargins(
        left=0.25, right=0.25,
        top=0.5, bottom=0.75,
        header=0.25, footer=0.25
    )
    ws.page_margins = margins

    total_width, _ = calculate_table_dimensions(ws)

    avail_portrait = PAGE_WIDTH_A4_PORTRAIT - margins.left - margins.right
    avail_landscape = PAGE_HEIGHT_A4 - margins.left - margins.right

    for scale in range(100, 59, -1):
        if total_width * scale / 100 <= avail_portrait:
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
            ws.page_setup.scale = scale
            return

    for scale in range(100, 59, -1):
        if total_width * scale / 100 <= avail_landscape:
            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
            ws.page_setup.scale = scale
            return

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.scale = 60


def clear_fills_openpyxl(wb):
    empty = PatternFill(fill_type=None)
    align = Alignment()

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                cell.fill = empty
                cell.alignment = align


def clear_fills_openpyxl(wb):
    empty = PatternFill(fill_type=None)
    align = Alignment()

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                cell.fill = empty
                cell.alignment = align


def process_excel_file(path, excel_app):
    """
    Обрабатывает Excel файл с использованием переданного экземпляра Excel.
    Возвращает True при успехе, False при ошибке.
    """
    try:
        # Проверяем существование файла
        path = Path(path)
        if not path.exists():
            logger.error(f"Файл не существует: {path}")
            return False

        logger.debug(f"Начинаю обработку файла: {path.name}")

        # Шаг 1: Обработка с openpyxl
        try:
            wb = load_workbook(path)
            clear_fills_openpyxl(wb)
            header_map = {}

            for ws in wb.worksheets:
                start, end = find_header_rows_any_border_advanced(ws)
                header_map[ws.title] = (start, end)
                ws.print_title_rows = f"{start}:{end}"
                ws.print_title_cols = "A:A"
                set_orientation_and_scale(ws)

            wb.save(path)
            wb.close()
            logger.debug(f"openpyxl обработка завершена для: {path.name}")

        except Exception as e:
            logger.error(f"Ошибка openpyxl при обработке {path.name}: {e}")
            return False

        # Шаг 2: Обработка с Excel COM
        try:
            wb_com = excel_app.Workbooks.Open(str(path))

            for ws in wb_com.Worksheets:
                if ws.Name in header_map:
                    s, e = header_map[ws.Name]
                    ws.PageSetup.PrintTitleRows = f"${s}:${e}"
                    ws.PageSetup.PrintTitleColumns = "$A:$A"
                ws.PageSetup.Order = 2

            wb_com.Save()
            wb_com.Close()
            logger.debug(f"Excel COM обработка завершена для: {path.name}")

        except Exception as e:
            logger.error(f"Ошибка Excel COM при обработке {path.name}: {e}")
            # Пытаемся закрыть книгу, если что-то пошло не так
            try:
                wb_com.Close()
            except:
                pass
            return False

        return True

    except Exception as e:
        logger.error(f"Неожиданная ошибка при обработке {path}: {e}")
        return False


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
