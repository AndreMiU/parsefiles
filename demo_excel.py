import gradio as gr
import pandas as pd
import json
from pathlib import Path
from io import BytesIO
import tempfile


def parse_excel_file(uploaded_file):
    """
    Обрабатывает загруженный Excel-файл и возвращает данные в формате JSON
    """
    try:
        # Создаем временный файл
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file)
            tmp_path = tmp.name

        # Читаем файл с помощью pandas
        xls = pd.ExcelFile(tmp_path)
        results = {
            "source_file": uploaded_file.name,
            "sheets": []
        }

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(
                xls,
                sheet_name=sheet_name,
                header=None,
                dtype=str,
                na_filter=False
            )
            sheet_data = df.fillna("").values.tolist()

            results["sheets"].append({
                "sheet_name": sheet_name,
                "data": sheet_data
            })

        # Конвертируем в красивый JSON
        json_output = json.dumps(results, ensure_ascii=False, indent=4)

        # Очищаем временные файлы
        Path(tmp_path).unlink()

        return json_output, results

    except Exception as e:
        return f"Ошибка обработки файла: {str(e)}", None


def display_sheet_data(data, sheet_index=0):
    """
    Преобразует данные листа в DataFrame для отображения
    """
    if not data or not data["sheets"]:
        return pd.DataFrame()

    sheet = data["sheets"][sheet_index]
    df = pd.DataFrame(sheet["data"])
    return df


# Создаем интерфейс Gradio
with gr.Blocks(title="Excel Parser Demo") as demo:
    gr.Markdown("## 📊 Парсер Excel-файлов")
    gr.Markdown("Загрузите Excel-файл (.xlsx или .xls) для просмотра его содержимого в формате JSON")

    with gr.Row():
        file_input = gr.File(label="Выберите Excel-файл", type="binary")

    with gr.Row():
        parse_btn = gr.Button("Парсить файл", variant="primary")

    with gr.Row():
        with gr.Column():
            json_output = gr.JSON(label="Результат в JSON")
        with gr.Column():
            sheet_selector = gr.Dropdown(label="Выберите лист")
            data_table = gr.DataFrame(label="Данные листа", wrap=True)


    # Обработчики событий
    def handle_file_upload(file):
        if file is None:
            return None, None, None, pd.DataFrame()

        json_data, full_data = parse_excel_file(file)
        sheet_names = [s["sheet_name"] for s in full_data["sheets"]] if full_data else []

        return (
            json_data,
            gr.Dropdown(choices=sheet_names, value=sheet_names[0] if sheet_names else None),
            full_data,
            display_sheet_data(full_data, 0) if full_data else pd.DataFrame()
        )


    def handle_sheet_select(full_data, sheet_name):
        if not full_data or not sheet_name:
            return pd.DataFrame()

        sheet_index = next((i for i, s in enumerate(full_data["sheets"]) if s["sheet_name"] == sheet_name), 0)
        return display_sheet_data(full_data, sheet_index)


    file_input.upload(
        fn=handle_file_upload,
        inputs=file_input,
        outputs=[json_output, sheet_selector, sheet_selector, data_table]
    )

    parse_btn.click(
        fn=handle_file_upload,
        inputs=file_input,
        outputs=[json_output, sheet_selector, sheet_selector, data_table]
    )

    sheet_selector.change(
        fn=handle_sheet_select,
        inputs=[sheet_selector, sheet_selector],
        outputs=data_table
    )

# Запуск приложения
if __name__ == "__main__":
    demo.launch()