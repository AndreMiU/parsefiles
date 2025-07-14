import os
import json
from pathlib import Path
import docx
from docx.document import Document as _Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

#Генератор, который последовательно возвращает все блоки (параграфы и таблицы) в документе или ячейке таблицы в порядке их появления.
def iter_block_items(parent):
    
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    else:
        parent_elm = parent

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

#Поиск файло формата docx в директории
def parse_directory_docs(directory_path):
  
    dir_path = Path(directory_path)

    if not dir_path.is_dir():
        print(f"Ошибка: Директория {directory_path} не существует!")
        return

    output_dir = dir_path / "parsed_results/Обработанные docx"
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Начало обработки DOCX-файлов в директории: {dir_path}")

    docx_files = list(dir_path.glob("**/*.docx"))

    if not docx_files:
        print("Не найдено DOCX-файлов для обработки!")
        return

    print(f"Найдено файлов: {len(docx_files)}")

    for docx_path in docx_files:
        try:
            print(f"\nОбработка файла: {docx_path.name}")
            process_docx_file(docx_path, output_dir)
        except Exception as e:
            print(f"Ошибка при обработке {docx_path.name}: {str(e)}")

    print("\nОбработка всех файлов завершена!")

#Обработка одного файла
def process_docx_file(docx_path, output_dir):
   
    base_name = docx_path.stem
    document_structure = extract_document_structure(docx_path)

    json_output = output_dir / f"{base_name}.json"
    with open(json_output, "w", encoding="utf-8") as json_file:
        json.dump(document_structure, json_file, ensure_ascii=False, indent=2)

    print(f"Результаты сохранены в: {json_output}")

Извлечение структуры документа
def extract_document_structure(docx_path):

    doc = docx.Document(docx_path)
    document_data = {
        "file_name": docx_path.name,
        "elements": [],
        "statistics": {
            "paragraphs": 0,
            "tables": 0,
            "table_rows": 0,
            "table_cells": 0
            }
        }

    element_counter = 0
    table_counter = 0

    for block in iter_block_items(doc):
        element_counter += 1
        element_data = {
            "element_id": element_counter,
            "type": None,
            "content": None
        }

        if isinstance(block, Paragraph):
            text = block.text.strip()
            if text:
                element_data["type"] = "paragraph"
                element_data["content"] = text
                document_data["elements"].append(element_data)
                document_data["statistics"]["paragraphs"] += 1

        elif isinstance(block, Table):
            table_counter += 1
            table_data = []
            total_rows = len(block.rows)
            total_cells = 0

            for row_idx, row in enumerate(block.rows):
                row_data = []

                for cell in row.cells:
                    cell_text = cell.text.strip().replace("\n", " ")
                    row_data.append(cell_text)
                    total_cells += 1

                table_data.append(row_data)

            element_data["type"] = "table"
            element_data["content"] = {
                "table_id": table_counter,
                "rows": total_rows,
                "columns": len(block.columns) if total_rows > 0 else 0,
                "cells": total_cells,
                "data": table_data
            }

            document_data["elements"].append(element_data)
            document_data["statistics"]["tables"] += 1
            document_data["statistics"]["table_rows"] += total_rows
            document_data["statistics"]["table_cells"] += total_cells

    document_data["statistics"]["total_elements"] = element_counter

    return document_data

#Пример использования
target_directory = "Входная директория"
parse_directory_docs(target_directory)
