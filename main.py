import shutil
import sys
from pathlib import Path
from photo_analyzer import PhotoFolderAnalyzer
from fill_rt import RTFiller
from fill_ap import APFiller
from fill_prilozhenie import PrilozhenieFiller
from report_data import ReportData


def get_int_input(prompt: str) -> int:
    while True:
        try:
            return int(input(prompt))
        except ValueError:
            print("Введите число!")


def check_templates_exist(*paths: Path) -> bool:
    missing = [p for p in paths if not p.exists()]
    if missing:
        print("❌ Ошибка: Не найдены следующие файлы шаблонов:")
        for p in missing:
            print(f"   - {p}")
        return False
    return True


def main():
    # Получаем базовую папку (работает как для .py, так и для .exe)
    if getattr(sys, 'frozen', False):
        # Если запускается exe
        base_path = Path(sys.executable).parent
    else:
        # Если запускается .py
        base_path = Path(__file__).parent
    
    photo_root = base_path / "Фото"
    templates_dir = base_path / "Шаблоны НЕ ТРОГАТЬ!!!"

    # Проверяем наличие необходимых папок
    if not photo_root.exists():
        print(f"❌ Ошибка: Папка 'Фото' не найдена!")
        print(f"   Ищу в: {photo_root}")
        input("\nНажмите любую клавишу для выхода...")
        return
    
    if not templates_dir.exists():
        print(f"❌ Ошибка: Папка 'Шаблоны НЕ ТРОГАТЬ!!!' не найдена!")
        print(f"   Ищу в: {templates_dir}")
        input("\nНажмите любую клавишу для выхода...")
        return

    rt_file = templates_dir / "РТ.xlsx"
    ap_file = templates_dir / "АП.xlsm"
    pril_template = templates_dir / "Шаблон.docx"

    if not check_templates_exist(rt_file, ap_file, pril_template):
        input("\nНажмите любую клавишу для выхода...")
        return

    report_data = ReportData()
    gbu_name, app_number = report_data.run()

    rt_copy = base_path / "Расчетные таблицы.xlsx"
    ap_copy = base_path / "Адресный перечень.xlsm"
    pril_template_copy = base_path / "Шаблон.docx"

    shutil.copy2(rt_file, rt_copy)
    shutil.copy2(ap_file, ap_copy)
    shutil.copy2(pril_template, pril_template_copy)

    analyzer = PhotoFolderAnalyzer()
    ap_filler = APFiller()

    counts = {
        "ДТ": get_int_input("Количество ДТ: "),
        "ДТ_пройденные": get_int_input("Количество пройденных ДТ: "),
        "МКД": get_int_input("Количество МКД: "),
        "МКД_пройденные": get_int_input("Количество пройденных МКД: "),
        "ОДХ": get_int_input("Количество ОДХ: "),
        "ОДХ_пройденные": get_int_input("Количество пройденных ОДХ: "),
        "ОО": get_int_input("Количество ОО: "),
        "ОО_пройденные": get_int_input("Количество пройденных ОО: ")
    }

    print()
    ap_filler.fill_counts(ap_copy, counts)
    ap_filler.fill_ap(ap_copy, photo_root)
    print("✅ Данные для адресного перечня обработаны")

    rt_filler = RTFiller(analyzer)
    rt_filler.fill_rt(rt_copy, photo_root, ap_copy, counts)
    print("✅ Данные для расчетных таблиц обработаны")

    filler = PrilozhenieFiller()

    print("📄 Создание приложения...")
    filler.fill_prilozhenie(
        pril_template_copy,
        photo_root,
        base_path / "Приложение.docx",
        gbu_name=gbu_name,
        app_number=app_number
    )
    print("✅ Приложение успешно создано")

    print("📄 Создание приложения устранения...")
    filler.fill_prilozhenie_ustraneniya(
        pril_template_copy,
        photo_root,
        base_path / "Приложение устранения.docx",
        gbu_name=gbu_name,
        app_number=app_number
    )
    print("✅ Приложение устранения успешно создано")

    pril_template_copy.unlink(missing_ok=True)

    print("\n🎉 Все документы успешно созданы!")
    print(f"📁 Файлы находятся в папке: {base_path}")
    print("   - Адресный перечень.xlsm")
    print("   - Расчетные таблицы.xlsx")
    print("   - Приложение.docx")
    print("   - Приложение устранения.docx")
    
    print("\n" + "="*50)
    input("Нажмите любую клавишу для выхода...")
    print("="*50)


if __name__ == "__main__":
    main()