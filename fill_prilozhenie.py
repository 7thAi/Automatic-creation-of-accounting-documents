"""
Модуль для заполнения приложений Word документами с фотографиями.
"""
import re
from pathlib import Path
from typing import List, Dict, Optional
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image
from io import BytesIO


class PrilozhenieFiller:
    """Заполняет Word-документы приложений фотографиями и подписями."""

    EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff'}

    PHOTO_WIDTH = Cm(12.98)
    PHOTO_HEIGHT = Cm(6.49)
    FONT_NAME = "Times New Roman"
    FONT_SIZE = Pt(14)

    VIOLATION_MAP = {
        "1. Проезд АБП": "локальные разрушения на проездах",
        "2. ДТС АБП": "локальные разрушения на ДТС",
        "3. Борт АБП": "локальные разрушения бортового камня",
        "4. ИДН АБП": "неудовлетворительное состояние ИДН",
        "5. Подпорка АБП": "неудовлетворительное состояние подпорной стены",
        "6. Санитарка АБП": "не убрана, не очищена территория",
        "1. Покрытие ДП": "неудовлетворительное состояние покрытия ДП",
        "2. МАФ ДП": "неудовлетворительное состояние МАФ на ДП",
        "3. Табличка ДП": "отсутствует информационная табличка на ДП",
        "4. Санитарка ДП": "не убрана, не очищена территория на ДП",
        "1. Покрытие СП": "неудовлетворительное состояние покрытия СП",
        "2. МАФ СП": "неудовлетворительное состояние МАФ на СП",
        "3. Табличка СП": "отсутствует информационная табличка на СП",
        "4. Санитарка СП": "не убрана, не очищена территория на СП",
        "4. КП": "неудовлетворительное состояние КП",
        "5. МАФ ДТ": "неудовлетворительное состояние МАФ на ДТ",
        "6. Газон": "неудовлетворительное состояние газона и зеленых насаждений",
        "1. Тех": "неудовлетворительное техническое состояние входной группы",
        "2. Сан": "неудовлетворительное санитарное состояние входной группы",
        "2. Отмостка": "неудовлетворительное состояние отмостки",
        "3. Цоколь": "неудовлетворительное состояние цоколя",
        "4. Ливневка": "неудовлетворительное состояние ливневых стоков",
        "5. Надписи": "наклейки, надписи на фасаде здания",
        "6. Сосульки": "наличие снежных масс и сосулек на кровлях и выступающих элементах фасадов зданий",
        "1. Проезд ОДХ": "локальные разрушения на проездах",
        "2. Тротуар ОДХ": "локальные разрушения на тротуарах",
        "3. Борт ОДХ": "локальные разрушения бортового камня",
        "4. ИДН ОДХ": "неудовлетворительное состояние ИДН",
        "5. Санитарка ОДХ": "не убрана, не очищена территория",
        "2. Тех ограждения": "неудовлетворительное техническое состояние ограждений",
        "3. Сан ограждения": "неудовлетворительное санитарное состояние ограждений",
        "4. Сан дорожные знаки": "неудовлетворительное санитарное состояние дорожных знаков",
        "5. МАФ ОДХ": "неудовлетворительное состояние МАФ на ОДХ",
        "1. Проезд ОО": "локальные разрушения на проездах",
        "2. ДТС ОО": "локальные разрушения на ДТС",
        "3. Борт ОО": "локальные разрушения бортового камня",
        "4. ИДН ОО": "неудовлетворительное состояние ИДН",
        "5. Подпорка ОО": "неудовлетворительное состояние подпорной стены",
        "6. Санитарка ОО": "не убрана, не очищена территория",
        "1. Покрытие ОО ДП": "неудовлетворительное состояние покрытия ДП",
        "2. МАФ ОО ДП": "неудовлетворительное состояние МАФ на ДП",
        "3. Табличка ОО ДП": "отсутствует информационная табличка на ДП",
        "4. Санитарка ОО ДП": "не убрана, не очищена территория на ДП",
        "1. Покрытие ОО СП": "неудовлетворительное состояние покрытия СП",
        "2. МАФ ОО СП": "неудовлетворительное состояние МАФ на СП",
        "3. Табличка ОО СП": "отсутствует информационная табличка на СП",
        "4. Санитарка ОО СП": "не убрана, не очищена территория на СП",
        "4. МАФ ОО": "неудовлетворительное состояние МАФ на ОО",
        "5. Газон ОО": "неудовлетворительное состояние газона и зеленых насаждений"
    }

    def _extract_gbu_short_name(self, gbu_name: str) -> str:
        """
        Извлекает короткое название ГБУ из полного названия.
        Например: "ГБУ «Автомобильные дороги ЦАО»" -> "«Автомобильные дороги ЦАО»"
        
        Args:
            gbu_name: Полное название ГБУ.
            
        Returns:
            Короткое название в кавычках.
        """
        # Ищем текст в кавычках «...»
        match = re.search(r'«(.+?)»', gbu_name)
        if match:
            return f"«{match.group(1)}»"
        return gbu_name

    def _update_document_headers(self, doc, gbu_name: str, app_number: int) -> None:
        """
        Обновляет заголовки документа: вставляет название ГБУ и номер приложения.
        
        Args:
            doc: Документ Word.
            gbu_name: Название ГБУ.
            app_number: Номер приложения.
        """
        gbu_short = self._extract_gbu_short_name(gbu_name)
        
        # Проходим по всем параграфам документа
        for para in doc.paragraphs:
            text = para.text
            
            # Заменяем "Приложение № ????" на "Приложение № {номер}"
            if "Приложение №" in text and "????" in text:
                for run in para.runs:
                    if "????" in run.text:
                        run.text = run.text.replace("????", str(app_number))
            
            # Ищем строки с "ГБУ" и заменяем название ГБУ
            if "ГБУ" in text:
                # Проходим по всем runs в параграфе и заменяем текст
                full_text = "".join(run.text for run in para.runs)
                
                # Если в параграфе есть шаблонное название ГБУ (с ??? или старым названием)
                if "???" in full_text or ("ГБУ" in full_text and "»" in full_text):
                    # Очищаем все runs в параграфе
                    for run in para.runs:
                        run.text = ""
                    
                    # Вставляем новый текст с правильным названием ГБУ
                    # Восстанавливаем структуру строки
                    if "Фотофиксация" in full_text or "нарушений" in full_text:
                        # Это заголовок вроде "Фотофиксация нарушений... ГБУ..."
                        new_text = full_text
                        # Заменяем все варианты названия ГБУ на правильное
                        new_text = re.sub(
                            r'ГБУ «[^»]*»',
                            f'ГБУ {gbu_short}',
                            new_text
                        )
                        new_text = re.sub(
                            r'ГБУ \?+',
                            f'ГБУ {gbu_short}',
                            new_text
                        )
                        
                        # Вставляем новый текст в первый run
                        para.runs[0].text = new_text
                    
                    # Применяем форматирование
                    for run in para.runs:
                        if run.text:  # Только если есть текст
                            run.font.name = self.FONT_NAME
                            run.font.size = self.FONT_SIZE
                            run.bold = True

    def fill_prilozhenie(self, template_path: Path, photo_root: Path, save_path: Path,
                         gbu_name: str = None, app_number: int = None,
                         show_progress: bool = False) -> None:
        """
        Заполняет приложение фотографиями (обе колонки).
        
        Args:
            template_path: Путь к шаблону документа.
            photo_root: Корневая папка с фотографиями.
            save_path: Путь для сохранения результата.
            gbu_name: Название ГБУ для вставки в заголовки.
            app_number: Номер приложения.
            show_progress: Показывать прогресс выполнения.
        """
        doc = Document(str(template_path))
        
        # Обновляем заголовки если переданы данные
        if gbu_name and app_number:
            self._update_document_headers(doc, gbu_name, app_number)
        
        self._fill_all_tables(doc, photo_root, left_only=False, show_progress=show_progress)
        doc.save(str(save_path))
        if show_progress:
            print(f"  {save_path.name}: 100% заполнено")

    def fill_prilozhenie_ustraneniya(self, template_path: Path, photo_root: Path, 
                                      save_path: Path, gbu_name: str = None, 
                                      app_number: int = None, show_progress: bool = False) -> None:
        """
        Заполняет приложение устранения (только левая колонка).
        
        Args:
            template_path: Путь к шаблону документа.
            photo_root: Корневая папка с фотографиями.
            save_path: Путь для сохранения результата.
            gbu_name: Название ГБУ для вставки в заголовки.
            app_number: Номер приложения.
            show_progress: Показывать прогресс выполнения.
        """
        doc = Document(str(template_path))
        
        # Обновляем заголовки если переданы данные
        if gbu_name and app_number:
            self._update_document_headers(doc, gbu_name, app_number)
        
        self._fill_all_tables(doc, photo_root, left_only=True, show_progress=show_progress)
        doc.save(str(save_path))
        if show_progress:
            print(f"  {save_path.name}: 100% заполнено")

    def _fill_all_tables(self, doc, photo_root: Path, left_only: bool = False, 
                         show_progress: bool = False) -> None:
        """
        Заполняет все таблицы в документе.
        
        Args:
            doc: Документ Word.
            photo_root: Корневая папка с фотографиями.
            left_only: Заполнять только левую колонку.
            show_progress: Показывать прогресс выполнения.
        """
        folders = ["1. ДТ", "2. МКД", "3. ОДХ", "4. ОО"]
        total_folders = len(folders)
        
        for table_index, folder in enumerate(folders):
            if table_index >= len(doc.tables):
                if show_progress:
                    print(f"  Таблица {table_index} не найдена в документе")
                continue
                
            table = doc.tables[table_index]
            folder_path = photo_root / folder
            
            if not folder_path.exists():
                if show_progress:
                    print(f"  Папка не существует: {folder_path}")
                continue
                
            photos = self._collect_photos(folder_path)
            self._fill_table(table, photos, left_only)
            
            if show_progress:
                progress = int((table_index + 1) / total_folders * 100)
                print(f"  {folder}: {len(photos)} фото ({progress}%)")

    def _collect_photos(self, root: Path) -> List[Dict]:
        """
        Собирает информацию о всех фотографиях в папке.
        
        Args:
            root: Корневая папка для сканирования.
            
        Returns:
            Список словарей с информацией о фотографиях.
        """
        result = []
        for item in sorted(root.rglob('*')):
            if item.is_file() and item.suffix.lower() in self.EXTENSIONS:
                subfolder = item.parent.name
                result.append({"path": item, "subfolder": subfolder.strip()})
        return result

    def _fill_table(self, table, photos: List[Dict], left_only: bool = False) -> None:
        """
        Заполняет таблицу фотографиями.
        
        Args:
            table: Таблица Word.
            photos: Список информации о фотографиях.
            left_only: Заполнять только левую колонку.
        """
        # Для left_only вставляем только в левую колонку (col_idx=0)
        if left_only:
            for photo_idx, info in enumerate(photos):
                row_idx = photo_idx * 2  # Каждое фото занимает 2 строки (фото + подпись)

                while row_idx + 1 >= len(table.rows):
                    table.add_row()
                    table.add_row()

                # Вставляем фото и подпись только в левую колонку (col_idx=0)
                self._insert_photo(table, row_idx, 0, info["path"])
                self._insert_caption(table, row_idx + 1, 0, info)
                
                # Очищаем правую колонку (col_idx=1)
                self._clear_cell(table, row_idx, 1)
                self._clear_cell(table, row_idx + 1, 1)
        else:
            # Обычный режим - заполняем обе колонки
            photo_idx = 0
            for info in photos:
                col_idx = photo_idx % 2
                row_idx = (photo_idx // 2) * 2

                while row_idx + 1 >= len(table.rows):
                    table.add_row()
                    table.add_row()

                self._insert_photo(table, row_idx, col_idx, info["path"])
                self._insert_caption(table, row_idx + 1, col_idx, info)

                photo_idx += 1

    def _clear_cell(self, table, row: int, col: int) -> None:
        """
        Очищает содержимое ячейки таблицы.
        
        Args:
            table: Таблица Word.
            row: Номер строки.
            col: Номер колонки.
        """
        try:
            cell = table.cell(row, col)
            for p in cell.paragraphs:
                p.clear()
        except Exception:
            pass  # Игнорируем ошибки очистки

    def _insert_photo(self, table, row: int, col: int, path: Path) -> None:
        """
        Вставляет фотографию в ячейку таблицы.
        
        Args:
            table: Таблица Word.
            row: Номер строки.
            col: Номер колонки.
            path: Путь к файлу изображения.
        """
        try:
            with Image.open(path) as img:
                resized_img = self._resize_image(img, 250)
                
                img_bytes = BytesIO()
                resized_img.save(img_bytes, format='PNG')
                img_bytes.seek(0)

                cell = table.cell(row, col)
                cell.paragraphs[0].clear()
                run = cell.paragraphs[0].add_run()
                run.add_picture(img_bytes, width=self.PHOTO_WIDTH, height=self.PHOTO_HEIGHT)
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        except Exception as e:
            print(f"  Ошибка при вставке фото {path}: {e}")

    def _insert_caption(self, table, row: int, col: int, info: Dict) -> None:
        """
        Вставляет подпись под фотографией.
        
        Args:
            table: Таблица Word.
            row: Номер строки.
            col: Номер колонки.
            info: Информация о фотографии.
        """
        cell = table.cell(row, col)
        for p in cell.paragraphs:
            p.clear()

        address = self._clean_address(info["path"])
        violation = self.VIOLATION_MAP.get(info["subfolder"], "неизвестный тип нарушения").strip()

        p = cell.paragraphs[0]
        p.paragraph_format.line_spacing = 1.0

        run1 = p.add_run("Адрес: ")
        run1.font.name = self.FONT_NAME
        run1.font.size = self.FONT_SIZE
        run1.bold = True

        run2 = p.add_run(address)
        run2.font.name = self.FONT_NAME
        run2.font.size = self.FONT_SIZE
        run2.bold = False

        p.add_run().add_break()

        run3 = p.add_run("Нарушение: ")
        run3.font.name = self.FONT_NAME
        run3.font.size = self.FONT_SIZE
        run3.bold = True

        run4 = p.add_run(violation)
        run4.font.name = self.FONT_NAME
        run4.font.size = self.FONT_SIZE
        run4.bold = False

    def _clean_address(self, path: Path) -> str:
        """
        Очищает адрес из имени файла.
        
        Args:
            path: Путь к файлу.
            
        Returns:
            Очищенный адрес.
        """
        name = re.sub(r"\s*\(\d+\)$", "", path.stem)
        return name.replace("_", "/")

    def _resize_image(self, img: Image.Image, target_dpi: int) -> Image.Image:
        """
        Изменяет размер изображения для целевого DPI.
        
        Args:
            img: Исходное изображение.
            target_dpi: Целевое разрешение DPI.
            
        Returns:
            Изменённое изображение.
        """
        original_dpi = img.info.get("dpi", (72, 72))[0]
        x_inch = img.width / original_dpi
        y_inch = img.height / original_dpi
        new_width = int(x_inch * target_dpi)
        new_height = int(y_inch * target_dpi)
        
        # Используем Image.Resampling.LANCZOS для совместимости с Pillow 10+
        try:
            resampling = Image.Resampling.LANCZOS
        except AttributeError:
            # Для старых версий Pillow
            resampling = Image.LANCZOS
            
        resized = img.resize((new_width, new_height), resampling)
        resized.info['dpi'] = (target_dpi, target_dpi)
        return resized
