import gradio as gr
import pdfplumber
import os
import json
from pathlib import Path


def extract_tables_with_lines(page):
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

    raw_tables = page.find_tables(table_settings)

    for table_num, table in enumerate(raw_tables):
        table_data = table.extract()

        if len(table_data) < 2 or len(table_data[0]) < 2:
            continue

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


def parse_pdf(file_path):
    results = {"pages": []}

    with pdfplumber.open(file_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            page_data = {"page_number": page_num + 1}

            # Extract text
            text = page.extract_text(
                x_tolerance=1,
                y_tolerance=1,
                layout=False,
                keep_blank_chars=False
            )
            page_data["text"] = text if text else ""

            # Extract tables
            page_data["tables"] = extract_tables_with_lines(page)

            results["pages"].append(page_data)

    return results


def process_file(file):
    if not file:
        return None, None, None

    # Save uploaded file temporarily
    file_path = file.name
    with open(file_path, "wb") as f:
        f.write(file.read())

    # Parse PDF
    results = parse_pdf(file_path)
    os.remove(file_path)  # Cleanup temporary file

    # Prepare outputs
    page_texts = [f"Страница {p['page_number']}:\n{p['text']}" for p in results["pages"]]
    text_output = "\n\n" + "-" * 50 + "\n\n".join(page_texts) + "\n\n" + "-" * 50

    tables_list = []
    for page in results["pages"]:
        for table in page["tables"]:
            tables_list.append({
                "Страница": page["page_number"],
                "Таблица": table["table_number"],
                "Позиция": table["position"],
                "Данные": table["data"]
            })

    json_output = json.dumps(results, ensure_ascii=False, indent=4)

    return text_output, tables_list, json_output


with gr.Blocks() as demo:
    gr.Markdown("## 🧠 PDF Parser Demo")
    gr.Markdown("Загрузите PDF-файл для извлечения текста и таблиц")

    with gr.Row():
        file_input = gr.File(label="PDF файл", file_types=[".pdf"])
        btn = gr.Button("Обработать")

    with gr.Tabs():
        with gr.Tab("Текст"):
            text_output = gr.Textbox(label="Извлеченный текст", lines=20)

        with gr.Tab("Таблицы"):
            table_output = gr.JSON(label="Извлеченные таблицы")

        with gr.Tab("Полные результаты (JSON)"):
            json_output = gr.JSON(label="JSON результат")

    btn.click(
        fn=process_file,
        inputs=file_input,
        outputs=[text_output, table_output, json_output]
    )

if __name__ == "__main__":
    demo.launch()
