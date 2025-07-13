import os
import json
import pdfplumber
from pathlib import Path


def parse_directory_pdfs(directory_path):
    """
    Парсит все PDF-файлы в директории с умеренной точностью извлечения таблиц
    Args:
        directory_path (str): Путь к директории с PDF-файлами
    """
    dir_path = Path(directory_path)

    if not dir_path.is_dir():
        print(f"Ошибка: Директория {directory_path} не существует!")
        return

    output_dir = dir_path / "parsed_results/ Обработанные pdf"
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Начало обработки PDF-файлов в директории: {dir_path}")
    pdf_files = list(dir_path.glob("**/*.pdf"))

    if not pdf_files:
        print("Не найдено PDF-файлов для обработки!")
        return

    print(f"Найдено файлов: {len(pdf_files)}")

    for pdf_path in pdf_files:
        process_pdf_file(pdf_path, output_dir)

    print("\nОбработка всех файлов завершена!")


def process_pdf_file(pdf_path, output_dir):
    """
    Обрабатывает PDF-файл с умеренной точностью извлечения таблиц
    Args:
        pdf_path (Path): Путь к PDF-файлу
        output_dir (Path): Директория для сохранения результатов
    """
    base_name = pdf_path.stem
    json_output = output_dir / f"{base_name}.json"

    results = {
        "source_file": str(pdf_path),
        "pages": []
    }

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                page_data = {"page_number": page_num + 1}

                # Извлечение текста
                text = page.extract_text(
                    x_tolerance=1,
                    y_tolerance=1,
                    layout=False,
                    keep_blank_chars=False
                )
                page_data["text"] = text if text else ""

                # Извлечение таблиц
                page_data["tables"] = extract_tables_with_lines(page)

                results["pages"].append(page_data)

    except Exception as e:
        print(f"Ошибка при обработке файла: {str(e)}")
        return

    # Сохранение результатов в JSON
    try:
        with open(json_output, "w", encoding="utf-8") as json_file:
            json.dump(results, json_file, ensure_ascii=False, indent=4)

        total_tables = sum(len(page['tables']) for page in results['pages'])
        print(f"  Страниц: {len(results['pages'])}")
        print(f"  Таблиц: {total_tables}")
        print(f"  Результаты сохранены в: {json_output}")
    except Exception as e:
        print(f"Ошибка при сохранении JSON: {str(e)}")


def extract_tables_with_lines(page):
    """
    Извлекает таблицы с умеренной точностью определения линий
    Args:
        page: Страница PDF из pdfplumber
    Returns:
        list: Список таблиц в формате JSON
    """
    tables = []

    # Упрощенные настройки для извлечения таблиц
    table_settings = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 5,
        "join_tolerance": 5,
        "edge_min_length": 10,
        "text_tolerance": 5,
        "text_x_tolerance": 5,
        "text_y_tolerance": 5,
    }

    try:
        # Извлекаем таблицы
        raw_tables = page.find_tables(table_settings)

        for table_num, table in enumerate(raw_tables):
            # Извлекаем данные таблицы без очистки
            table_data = table.extract()

            # Пропускаем слишком маленькие таблицы
            if len(table_data) < 2 or len(table_data[0]) < 2:
                continue

            # Получаем координаты таблицы
            bbox = table.bbox
            tables.append({
                "table_number": table_num + 1,
                "position": {
                    "x": round(bbox[0], 1),
                    "y": round(bbox[1], 1),
                    "width": round(bbox[2] - bbox[0], 1),
                    "height": round(bbox[3] - bbox[1], 1)
                },
                "data": table_data
            })

    except Exception as e:
        print(f"Ошибка при извлечении таблиц: {str(e)}")

    return tables


if __name__ == "__main__":
    target_directory = "D:\\Тесты"
    parse_directory_pdfs(target_directory)