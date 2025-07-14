import os
import json
import pdfplumber
from pathlib import Path

#Поиск файлов формата pdf в указанной директории
def parse_directory(directory_path):
  
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
        process_pdF(pdf_path, output_dir)

    print("\nОбработка всех файлов завершена!")

#Обработка одного файла
def process_pdf(pdf_path, output_dir):
   
    base_name = pdf_path.stem
    json_output = output_dir / f"{base_name}.json"

    results = {
        "source_file": str(pdf_path),
        "pages": []
    }

    
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
            page_data["tables"] = extract_tables(page)

            results["pages"].append(page_data)

    # Сохранение результатов в JSON без обработки исключений
    with open(json_output, "w", encoding="utf-8") as json_file:
        json.dump(results, json_file, ensure_ascii=False, indent=4)

    total_tables = sum(len(page['tables']) for page in results['pages'])
    print(f"  Страниц: {len(results['pages'])}")
    print(f"  Таблиц: {total_tables}")
    print(f"  Результаты сохранены в: {json_output}")


def extract_tables(page):
    
    tables = []
    table_settings = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 4,
        "join_tolerance": 4,
        "edge_min_length": 10,
        "text_tolerance": 4,
        "text_x_tolerance": 4,
        "text_y_tolerance": 4,
    }

    # Извлечение таблиц
    raw_tables = page.find_tables(table_settings)

    for table_num, table in enumerate(raw_tables):
        table_data = table.extract()

        if len(table_data) < 2 or len(table_data[0]) < 2:
            continue
        #Определние координат таблицы в файле
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

    return tables



target_directory = "D:\\Тесты"
parse_directory(target_directory)
