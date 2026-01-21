import shutil
import sys
import logging
from pathlib import Path
from typing import Dict, Optional, Tuple
from photo_analyzer import PhotoFolderAnalyzer
from fill_rt import RTFiller
from fill_ap import APFiller
from fill_prilozhenie import PrilozhenieFiller
from report_data import ReportData

logger = logging.getLogger(__name__)


class ProjectConfig:
    """Конфигурация проекта с путями и константами."""
    TEMPLATES_DIR_NAME = "Шаблоны НЕ ТРОГАТЬ!!!"
    PHOTOS_DIR_NAME = "Фото"
    TEMPLATE_FILES = {"rt": "РТ.xlsx", "ap": "АП.xlsm", "prilozhenie": "Шаблон.docx"}
    OUTPUT_FILES = {
        "rt_copy": "Расчетные таблицы.xlsx",
        "ap_copy": "Адресный перечень.xlsm",
        "prilozhenie_template": "Шаблон.docx",
        "prilozhenie": "Приложение.docx",
        "prilozhenie_ustraneniya": "Приложение устранения.docx"
    }


def get_int_input(prompt: str) -> int:
    """Получает целое число от пользователя с валидацией."""
    while True:
        try:
            return int(input(prompt))
        except ValueError:
            print("Введите число!")


def check_paths_exist(*paths: Path) -> bool:
    """Проверяет существование путей и выводит список отсутствующих."""
    missing = [p for p in paths if not p.exists()]
    if missing:
        print("❌ Ошибка: Не найдены следующие файлы:")
        for p in missing:
            print(f"   - {p}")
        return False
    return True


def get_base_path() -> Path:
    """Получает базовую папку приложения (работает как для .py, так и для .exe)."""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent


def init_project_paths(base_path: Path) -> Optional[Tuple[Path, Path, Dict[str, Path]]]:
    """Инициализирует и проверяет пути проекта.
    
    Returns:
        Кортеж (photo_root, templates_dir, template_paths) или None при ошибке.
    """
    photo_root = base_path / ProjectConfig.PHOTOS_DIR_NAME
    templates_dir = base_path / ProjectConfig.TEMPLATES_DIR_NAME
    
    if not photo_root.exists():
        print(f"❌ Ошибка: Папка '{ProjectConfig.PHOTOS_DIR_NAME}' не найдена!")
        print(f"   Ищу в: {photo_root}")
        return None
    
    if not templates_dir.exists():
        print(f"❌ Ошибка: Папка '{ProjectConfig.TEMPLATES_DIR_NAME}' не найдена!")
        print(f"   Ищу в: {templates_dir}")
        return None
    
    template_paths = {
        key: templates_dir / ProjectConfig.TEMPLATE_FILES[key]
        for key in ProjectConfig.TEMPLATE_FILES
    }
    
    if not check_paths_exist(*template_paths.values()):
        return None
    
    return photo_root, templates_dir, template_paths


def collect_user_counts() -> Dict[str, int]:
    """Собирает от пользователя количество объектов по категориям."""
    categories = [
        ("ДТ", "Количество ДТ: "),
        ("ДТ_пройденные", "Количество пройденных ДТ: "),
        ("МКД", "Количество МКД: "),
        ("МКД_пройденные", "Количество пройденных МКД: "),
        ("ОДХ", "Количество ОДХ: "),
        ("ОДХ_пройденные", "Количество пройденных ОДХ: "),
        ("ОО", "Количество ОО: "),
        ("ОО_пройденные", "Количество пройденных ОО: ")
    ]
    return {key: get_int_input(prompt) for key, prompt in categories}


def copy_templates(base_path: Path, template_paths: Dict[str, Path]) -> Dict[str, Path]:
    """Копирует файлы шаблонов в рабочую директорию.
    
    Returns:
        Словарь с путями скопированных файлов.
    """
    output_paths = {}
    template_mapping = [
        ("rt", "rt_copy"),
        ("ap", "ap_copy"),
        ("prilozhenie", "prilozhenie_template")
    ]
    
    for src_key, dst_key in template_mapping:
        src_path = template_paths[src_key]
        dst_path = base_path / ProjectConfig.OUTPUT_FILES[dst_key]
        shutil.copy2(src_path, dst_path)
        output_paths[dst_key] = dst_path
        logger.debug(f"Скопирован: {src_path.name} -> {dst_path.name}")
    
    return output_paths


def remove_empty_folders(root_path: Path) -> int:
    """Удаляет все пустые папки рекурсивно.
    
    Args:
        root_path: Корневая папка для поиска пустых папок.
        
    Returns:
        Количество удаленных папок.
    """
    removed_count = 0
    
    # Проходим по всем подпапкам в обратном порядке (от листьев к корню)
    for item in sorted(root_path.rglob('*'), key=lambda p: len(p.parts), reverse=True):
        if item.is_dir():
            try:
                # Пытаемся удалить папку - если она не пустая, исключение
                item.rmdir()
                removed_count += 1
                logger.debug(f"Удалена пустая папка: {item}")
            except OSError:
                # Папка не пустая, игнорируем
                pass
    
    return removed_count


def main():
    base_path = get_base_path()
    
    # Инициализируем пути
    result = init_project_paths(base_path)
    if result is None:
        input("\nНажмите любую клавишу для выхода...")
        return
    
    photo_root, templates_dir, template_paths = result
    
    # Получаем данные ГБУ
    report_data = ReportData()
    gbu_name, app_number = report_data.run()
    
    # Копируем шаблоны
    output_paths = copy_templates(base_path, template_paths)
    
    # Получаем количества от пользователя
    print()
    counts = collect_user_counts()
    
    # Заполняем адресный перечень
    print()
    ap_filler = APFiller()
    ap_filler.fill_counts(output_paths["ap_copy"], counts)
    ap_filler.fill_ap(output_paths["ap_copy"], photo_root)
    print("✅ Данные для адресного перечня обработаны")
    
    # Заполняем расчетные таблицы
    analyzer = PhotoFolderAnalyzer()
    rt_filler = RTFiller(analyzer)
    rt_filler.fill_rt(output_paths["rt_copy"], photo_root, output_paths["ap_copy"], counts)
    print("✅ Данные для расчетных таблиц обработаны")
    
    # Создаем приложения
    filler = PrilozhenieFiller()
    
    print("📄 Создание приложения...")
    filler.fill_prilozhenie(
        output_paths["prilozhenie_template"],
        photo_root,
        base_path / ProjectConfig.OUTPUT_FILES["prilozhenie"],
        gbu_name=gbu_name,
        app_number=app_number
    )
    print("✅ Приложение успешно создано")
    
    print("📄 Создание приложения устранения...")
    filler.fill_prilozhenie_ustraneniya(
        output_paths["prilozhenie_template"],
        photo_root,
        base_path / ProjectConfig.OUTPUT_FILES["prilozhenie_ustraneniya"],
        gbu_name=gbu_name,
        app_number=app_number
    )
    print("✅ Приложение устранения успешно создано")
    
    # Удаляем временный файл шаблона
    output_paths["prilozhenie_template"].unlink(missing_ok=True)
    
    # Удаляем пустые папки из папки Фото
    remove_empty_folders(photo_root)
    
    # Вывод результатов
    print("\n🎉 Все документы успешно созданы!")
    print(f"📁 Файлы находятся в папке: {base_path}")
    output_files = [ProjectConfig.OUTPUT_FILES[key] for key in 
                    ["ap_copy", "rt_copy", "prilozhenie", "prilozhenie_ustraneniya"]]
    for file in output_files:
        print(f"   - {file}")
    
    print("\n" + "="*60)
    input("✓ Работа завершена. Нажмите любую клавишу для выхода...")
    print("="*60)


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.WARNING,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    try:
        main()
    except Exception as e:
        logger.error(f"Критическая ошибка: {e}", exc_info=True)
        print(f"\n❌ Критическая ошибка: {e}")
        input("\nНажмите любую клавишу для выхода...")