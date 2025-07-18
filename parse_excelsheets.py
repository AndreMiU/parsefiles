import os
import json
import pandas as pd
from pathlib import Path


def parse_directory_excel(directory_path):
    """
    Парсит все Excel-файлы в директории
    Args:
        directory_path (str): Путь к директории с Excel-файлами
    Returns:
        list: Список путей к созданным JSON-файлам
    """
    dir_path = Path(directory_path)

    if not dir_path.is_dir():
        print(f"Ошибка: Директория {directory_path} не существует!")
        return []

    output_dir = dir_path / "parsed_results/Обработанные excel"
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Начало обработки Excel-файлов в директории: {dir_path}")
    excel_files = list(dir_path.glob("**/*.xlsx")) + list(dir_path.glob("**/*.xls"))

    if not excel_files:
        print("Не найдено Excel-файлов для обработки!")
        return []

    print(f"Найдено файлов: {len(excel_files)}")

    all_json_files = []  # Список всех созданных JSON-файлов

    for excel_path in excel_files:
        # Обрабатываем файл и получаем список JSON-файлов
        json_files = process_excel_file(excel_path, output_dir)
        if json_files:
            print(f"  Обработано листов: {len(json_files)}")
            all_json_files.extend(json_files)

    print("\nОбработка всех файлов завершена!")
    return all_json_files


def process_excel_file(excel_path, output_dir):
    """
    Обрабатывает Excel-файл: каждый лист сохраняет в отдельный JSON
    и возвращает список путей к созданным JSON-файлам.
    Args:
        excel_path (Path): Путь к Excel-файлу
        output_dir (Path): Директория для сохранения результатов
    Returns:
        list: Список путей к созданным JSON-файлам
    """
    base_name = excel_path.stem
    file_output_dir = output_dir / base_name
    file_output_dir.mkdir(parents=True, exist_ok=True)

    json_files = []  # Список путей к JSON-файлам
    preview_dir = file_output_dir / "previews"
    preview_dir.mkdir(exist_ok=True)

    try:
        print(f"\nОбработка файла: {excel_path.name}")
        xls = pd.ExcelFile(excel_path)

        for sheet_name in xls.sheet_names:
            # Читаем лист как строки без преобразования типов
            df = pd.read_excel(
                xls,
                sheet_name=sheet_name,
                header=None,
                dtype=str,
                na_filter=False
            )

            # Очищаем имя листа для использования в имени файла
            safe_sheet_name = "".join(
                c if c.isalnum() or c in " _-" else "_"
                for c in sheet_name
            ).strip()

            if safe_sheet_name == "":
                safe_sheet_name = "empty_sheet_name"

            # Формируем путь для JSON
            json_path = file_output_dir / f"{safe_sheet_name}.json"
            json_files.append(json_path)

            # Преобразуем DataFrame в список списков (двумерный массив)
            sheet_data = df.fillna("").values.tolist()

            # Сохраняем данные в JSON
            with open(json_path, "w", encoding="utf-8") as json_file:
                json.dump(sheet_data, json_file, ensure_ascii=False, indent=4)

            # Сохраняем превью листа (первые 5 строк)
            preview_path = preview_dir / f"{safe_sheet_name}_preview.txt"
            with open(preview_path, "w", encoding="utf-8") as preview_file:
                preview_file.write(f"Preview of sheet: {sheet_name}\n")
                preview_file.write(
                    f"Total rows: {len(sheet_data)}, columns: {len(sheet_data[0]) if sheet_data else 0}\n\n")
                preview_file.write("\n".join("\t".join(map(str, row[:10])) for row in sheet_data[:5]))

            print(f"  Лист '{sheet_name}' -> {json_path}")
            print(f"    Превью сохранено в: {preview_path}")

        return json_files

    except Exception as e:
        print(f"Ошибка при обработке файла {excel_path}: {str(e)}")
        return []  # Возвращаем пустой список в случае ошибки


def load_json_to_dataframe(json_path):
    """
    Загружает JSON-файл в pandas DataFrame
    Args:
        json_path (str/Path): Путь к JSON-файлу
    Returns:
        pd.DataFrame: DataFrame с данными из JSON
    """
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return pd.DataFrame(data)


if __name__ == "__main__":
    target_directory = "C:\\Users/ANDREY/PyCharmMiscProject/test_files"

    # Обрабатываем файлы и получаем список JSON
    json_files = parse_directory_excel(target_directory)

    # Демонстрация работы с JSON
    if json_files:
        first_json = json_files[0]
        print(f"\nДемонстрация загрузки JSON в DataFrame: {first_json}")

        # Загружаем данные в DataFrame
        df = load_json_to_dataframe(first_json)

        # Показываем информацию о DataFrame
        print("\nИнформация о DataFrame:")
        print(f"Размер: {df.shape[0]} строк, {df.shape[1]} столбцов")
        print("\nПервые 5 строк:")
        print(df.head())

        # Сохраняем полный DataFrame в CSV для просмотра
        csv_path = Path(first_json).with_suffix('.csv')
        df.to_csv(csv_path, index=False, encoding='utf-8')
        print(f"\nПолный DataFrame сохранен в CSV: {csv_path}")
    else:
        print("\nНет обработанных файлов для демонстрации")
