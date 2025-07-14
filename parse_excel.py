import os
import json
import pandas as pd
from pathlib import Path


def parse_directory_excel(directory_path):
   
    dir_path = Path(directory_path)

    if not dir_path.is_dir():
        print(f"Ошибка: Директория {directory_path} не существует!")
        return

    output_dir = dir_path / "parsed_results/Обработанные excel"
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Начало обработки Excel-файлов в директории: {dir_path}")
    excel_files = list(dir_path.glob("**/*.xlsx")) + list(dir_path.glob("**/*.xls"))

    if not excel_files:
        print("Не найдено Excel-файлов для обработки!")
        return

    print(f"Найдено файлов: {len(excel_files)}")

    for excel_path in excel_files:
        process_excel_file(excel_path, output_dir)

    print("\nОбработка всех файлов завершена!")


def process_excel_file(excel_path, output_dir):

    base_name = excel_path.stem
    json_output = output_dir / f"{base_name}.json"

    results = {
        "source_file": str(excel_path),
        "sheets": []
    }

    try:
        
        xls = pd.ExcelFile(excel_path)

        for sheet_name in xls.sheet_names:
            
            df = pd.read_excel(
                xls,
                sheet_name=sheet_name,
                header=None,
                dtype=str,
                na_filter=False
            )

            # Заменяем NaN на пустые строки и преобразуем в список
            sheet_data = df.fillna("").values.tolist()

            results["sheets"].append({
                "sheet_name": sheet_name,
                "data": sheet_data
            })

        
        with open(json_output, "w", encoding="utf-8") as json_file:
            json.dump(results, json_file, ensure_ascii=False, indent=4)

        print(f"Файл обработан: {excel_path.name}")
        print(f"  Листов: {len(results['sheets'])}" )
        print(f"  Результаты сохранены в: {json_output}")

    except Exception as e:
        print(f"Ошибка при обработке файла {excel_path}: {str(e)}")


#Пример использования
target_directory = "D:\\Тесты"
parse_directory_excel(target_directory)
