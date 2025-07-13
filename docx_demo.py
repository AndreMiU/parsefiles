import gradio as gr
import io
import json
import os
from pathlib import Path
import docx
from docx.document import Document as _Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph


def iter_block_items(parent):
    """
    Генератор, который последовательно возвращает все элементы документа
    (параграфы и таблицы) в порядке их следования
    """
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    else:
        parent_elm = parent

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def extract_document_structure(docx_path):
    """
    Извлекает структуру документа (текст и таблицы) в формате JSON
    """
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


def process_docx(file_obj):
    """
    Обрабатывает DOCX-файл и возвращает JSON и статистику
    """
    if file_obj is None:
        return None, "Загрузите DOCX-файл"

    try:
        # Получаем временный путь к загруженному файлу
        docx_path = Path(file_obj.name)

        # Извлекаем структуру документа
        data = extract_document_structure(docx_path)

        # Форматируем статистику
        stats = data["statistics"]
        stats_text = (
            f"📊 Статистика документа:\n"
            f"• Имя файла: {data['file_name']}\n"
            f"• Всего элементов: {stats['total_elements']}\n"
            f"• Параграфов: {stats['paragraphs']}\n"
            f"• Таблиц: {stats['tables']}\n"
            f"• Строк в таблицах: {stats['table_rows']}\n"
            f"• Ячеек в таблицах: {stats['table_cells']}"
        )

        # Создаем JSON для скачивания
        json_data = json.dumps(data, ensure_ascii=False, indent=2)
        json_bytes = json_data.encode("utf-8")
        json_file = io.BytesIO(json_bytes)
        json_file.name = f"{docx_path.stem}.json"

        return json_file, stats_text

    except Exception as e:
        return None, f"⛔ Ошибка обработки: {str(e)}"


# Создаем интерфейс Gradio
with gr.Blocks(title="DOCX Parser", theme="soft") as demo:
    gr.Markdown("# 🗂️ Парсер DOCX-документов")
    gr.Markdown("Загрузите DOCX-файл для извлечения структуры (текст и таблицы) в формате JSON")

    with gr.Row():
        file_input = gr.File(
            label="Выберите DOCX-файл",
            file_types=[".docx"],
            type="file"
        )

    with gr.Row():
        process_btn = gr.Button("🚀 Обработать документ", variant="primary")

    with gr.Row():
        json_output = gr.File(label="Скачать JSON результат")

    with gr.Row():
        stats_output = gr.Textbox(
            label="Статистика документа",
            interactive=False,
            lines=6
        )

    process_btn.click(
        fn=process_docx,
        inputs=file_input,
        outputs=[json_output, stats_output]
    )

# Запускаем приложение
if __name__ == "__main__":
    demo.launch()