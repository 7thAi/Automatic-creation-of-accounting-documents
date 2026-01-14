import shutil
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
    for path in paths:
        if not path.exists():
            print(f"Ошибка: Шаблон не найден: {path}")
            return False
    return True


def main():
    base_path = Path(__file__).parent
    photo_root = base_path / "Фото"

    templates_dir = base_path / "Шаблоны НЕ ТРОГАТЬ!!!"

    rt_file = templates_dir / "РТ.xlsx"
    ap_file = templates_dir / "АП.xlsm"
    pril_template = templates_dir / "Шаблон.docx"

    if not check_templates_exist(rt_file, ap_file, pril_template):
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
    shutil.rmtree(templates_dir, ignore_errors=True)

    print("\n🎉 Все документы успешно созданы!")


if __name__ == "__main__":
    main()