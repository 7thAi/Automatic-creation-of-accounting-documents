"""
Модуль для заполнения расчётных таблиц (РТ) данными из фотографий и АП.
"""
import logging
from pathlib import Path
from typing import Optional, Dict
import xlwings as xw

logger = logging.getLogger(__name__)


class RTFiller:
    """Заполняет Excel-файл РТ фото и берет суммарные значения из АП напрямую из Excel."""

    # Карта листов -> подпапки -> ячейки РТ
    SHEET_CELL_MAP = {
        "ДТ": {
            "1. Проезд АБП": "J7",
            "2. ДТС АБП": "K7",
            "3. Борт АБП": "L7",
            "4. ИДН АБП": "M7",
            "5. Подпорка АБП": "N7",
            "6. Санитарка АБП": "O7",
            "1. Покрытие ДП": "Q7",
            "2. МАФ ДП": "R7",
            "3. Табличка ДП": "S7",
            "4. Санитарка ДП": "T7",
            "1. Покрытие СП": "V7",
            "2. МАФ СП": "W7",
            "3. Табличка СП": "X7",
            "4. Санитарка СП": "Y7",
            "4. КП": "Z7",
            "5. МАФ ДТ": "AA7",
            "6. Газон": "AB7"
        },
        "МКД": {
            "1. Тех": "J8",
            "2. Сан": "K8",
            "2. Отмостка": "L8",
            "3. Цоколь": "M8",
            "4. Ливневка": "N8",
            "5. Надписи": "O8",
            "6. Сосульки": "P8"
        },
        "ОДХ": {
            "1. Проезд ОДХ": "J7",
            "2. Тротуар ОДХ": "K7",
            "3. Борт ОДХ": "L7",
            "4. ИДН ОДХ": "M7",
            "5. Санитарка ОДХ": "N7",
            "2. Тех ограждения": "P7",
            "3. Сан ограждения": "Q7",
            "4. Сан дорожные знаки": "R7",
            "5. МАФ ОДХ": "S7"
        },
        "ОО": {
            "1. Проезд ОО": "J7",
            "2. ДТС ОО": "K7",
            "3. Борт ОО": "L7",
            "4. ИДН ОО": "M7",
            "5. Подпорка ОО": "N7",
            "6. Санитарка ОО": "O7",
            "1. Покрытие ОО ДП": "Q7",
            "2. МАФ ОО ДП": "R7",
            "3. Табличка ОО ДП": "S7",
            "4. Санитарка ОО ДП": "T7",
            "1. Покрытие ОО СП": "V7",
            "2. МАФ ОО СП": "W7",
            "3. Табличка ОО СП": "X7",
            "4. Санитарка ОО СП": "Y7",
            "4. МАФ ОО": "Z7",
            "5. Газон ОО": "AA7"
        }
    }

    # Суммарные значения из АП -> РТ
    SUMMARY_MAP = {
        "ДТ": ("АП ДТ", "F2", "ДТ", "F7"),
        "МКД": ("АП МКД", "G2", "МКД", "F8"),
        "ОДХ": ("АП ОДХ", "F2", "ОДХ", "F7"),
        "ОО": ("АП ОО", "F2", "ОО", "F7")
    }

    # Карта для записи количеств объектов в РТ
    # Формат: ключ из counts -> (лист, ячейка)
    COUNTS_MAP = {
        "ДТ": ("ДТ", "C7"),
        "ДТ_пройденные": ("ДТ", "D7"),
        "МКД": ("МКД", "C8"),
        "МКД_пройденные": ("МКД", "D8"),
        "ОДХ": ("ОДХ", "C7"),
        "ОДХ_пройденные": ("ОДХ", "D7"),
        "ОО": ("ОО", "C7"),
        "ОО_пройденные": ("ОО", "D7")
    }

    def __init__(self, analyzer):
        """
        Инициализация RTFiller.
        
        Args:
            analyzer: Экземпляр PhotoFolderAnalyzer для подсчёта фотографий.
        """
        self.analyzer = analyzer

    def find_folder_recursive(self, root: Path, target_name: str) -> Optional[Path]:
        """Ищет папку с точным именем рекурсивно.
        
        Args:
            root: Корневая директория для поиска.
            target_name: Имя искомой папки.
            
        Returns:
            Path к найденной папке или None.
        """
        for item in root.rglob("*"):
            if item.is_dir() and item.name == target_name:
                return item
        return None

    def _fill_cell_with_photo_count(self, sheet, folder_name: str, cell: str, folder_path: Path) -> None:
        """Подсчитывает фото в папке и записывает количество в ячейку.
        
        Args:
            sheet: Лист Excel.
            folder_name: Имя папки (для логирования).
            cell: Адрес ячейки.
            folder_path: Путь к папке с фотографиями.
        """
        count = self.analyzer.count_photos_in_folder(folder_path)
        sheet.range(cell).value = count
        logger.debug(f"Записано {count} фото из {folder_name} в ячейку {cell}")

    def fill_rt(self, rt_path: Path, photo_root: Path, ap_path: Path, counts: Optional[Dict[str, int]] = None) -> None:
        """Заполняет РТ фото, суммарные значения из АП и количества объектов.
        
        Args:
            rt_path: Путь к файлу РТ.
            photo_root: Корневая папка с фотографиями.
            ap_path: Путь к файлу АП.
            counts: Словарь с количеством объектов (ДТ, МКД, ОДХ, ОО и пройденные).
        """
        app = None
        wb_rt = None
        wb_ap = None
        
        try:
            app = xw.App(visible=False)
            wb_rt = xw.Book(str(rt_path))
            wb_ap = xw.Book(str(ap_path))

            # Кэшируем имена листов для эффективности
            rt_sheet_names = {s.name for s in wb_rt.sheets}
            
            # Заполняем количества объектов
            if counts:
                for key, value in counts.items():
                    if key in self.COUNTS_MAP:
                        sheet_name, cell = self.COUNTS_MAP[key]
                        if sheet_name in rt_sheet_names:
                            wb_rt.sheets[sheet_name].range(cell).value = value
                            logger.debug(f"Записано количество {value} для {key} в {sheet_name}!{cell}")

            # Суммарные значения из АП
            for key, (ap_sheet_name, ap_cell, rt_sheet_name, rt_cell) in self.SUMMARY_MAP.items():
                try:
                    val = wb_ap.sheets[ap_sheet_name].range(ap_cell).value
                    wb_rt.sheets[rt_sheet_name].range(rt_cell).value = val
                    logger.debug(f"Скопировано значение {val} для {key}")
                except Exception as e:
                    logger.warning(f"Ошибка при копировании {key}: {e}")

            # Подсчет по подпапкам
            for sheet_name, folders in self.SHEET_CELL_MAP.items():
                if sheet_name not in rt_sheet_names:
                    logger.warning(f"Лист {sheet_name} не найден в РТ")
                    continue
                    
                sheet = wb_rt.sheets[sheet_name]
                for folder_name, cell in folders.items():
                    folder_path = self.find_folder_recursive(photo_root, folder_name)
                    if folder_path:
                        self._fill_cell_with_photo_count(sheet, folder_name, cell, folder_path)
                    else:
                        logger.debug(f"Папка {folder_name} не найдена")

            wb_rt.save(str(rt_path))
            logger.info("РТ успешно сохранены")
            
        except Exception as e:
            logger.error(f"Ошибка при заполнении РТ: {e}")
            raise
        finally:
            # Гарантированное закрытие ресурсов
            if wb_rt:
                wb_rt.close()
            if wb_ap:
                wb_ap.close()
            if app:
                app.quit()
