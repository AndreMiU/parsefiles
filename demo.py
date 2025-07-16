import streamlit as st
import os
import sys
import json
import tempfile
from pathlib import Path
import importlib.util


# Функция для импорта модуля из файла
def import_module_from_path(module_name, file_path):
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


# Импорт функций из файлов
current_dir = Path(__file__).parent

# Импорт DOCX парсера
docx_parser = import_module_from_path(
    "docx_parser",
    current_dir / "parse_docx.py"
)

# Импорт PDF парсера
pdf_parser = import_module_from_path(
    "pdf_parser",
    current_dir / "parse_pdf.py"
)

# Импорт Excel парсера
excel_parser = import_module_from_path(
    "excel_parser",
    current_dir / "parse_excel.py"
)


# Основное приложение Streamlit
def main():
    st.set_page_config(page_title="Парсер документов", layout="wide")
    st.title("📄 Парсер документов (DOCX, Excel, PDF)")
    st.write("Обработка документов и извлечение структурированных данных")

    # Вкладки для выбора режима обработки
    tab_dir, tab_file = st.tabs(["Обработка директории", "Обработка одного файла"])

    with tab_dir:
        st.subheader("Пакетная обработка документов в директории")
        with st.form("directory_processing"):
            directory_path = st.text_input("Путь к директории:", "documents")
            process_docx = st.checkbox("Обрабатывать DOCX", value=True)
            process_excel = st.checkbox("Обрабатывать Excel", value=True)
            process_pdf = st.checkbox("Обрабатывать PDF", value=True)

            if st.form_submit_button("Запустить обработку", type="primary"):
                run_directory_processing(directory_path, process_docx, process_excel, process_pdf)

        st.info("""
        **Инструкция:**
        1. Укажите путь к директории с документами
        2. Выберите типы файлов для обработки
        3. Нажмите кнопку "Запустить обработку"
        4. Результаты будут сохранены в подпапках внутри указанной директории
        """)

    with tab_file:
        st.subheader("Обработка отдельного документа")
        with st.form("file_processing"):
            file_type = st.radio("Тип файла:", ["DOCX", "PDF", "Excel"], index=0)
            uploaded_file = st.file_uploader(
                f"Загрузите {file_type} файл",
                type=["docx"] if file_type == "DOCX" else ["pdf"] if file_type == "PDF" else ["xlsx", "xls"]
            )

            if st.form_submit_button("Обработать файл", type="primary") and uploaded_file:
                run_single_file_processing(uploaded_file, file_type)

        st.info("""
        **Инструкция:**
        1. Выберите тип файла
        2. Загрузите документ
        3. Нажмите кнопку "Обработать файл"
        4. Результат будет отображён ниже и доступен для скачивания
        """)


def run_directory_processing(directory_path, process_docx, process_excel, process_pdf):
    dir_path = Path(directory_path)

    if not dir_path.is_dir():
        st.error(f"Директория {directory_path} не существует!")
        return

    output_root = dir_path / "parsed_results"
    output_root.mkdir(parents=True, exist_ok=True)

    progress_bar = st.progress(0)
    status_text = st.empty()

    # Обработка DOCX
    if process_docx:
        output_dir = output_root / "Обработанные docx"
        output_dir.mkdir(exist_ok=True)

        with st.spinner("Поиск DOCX файлов..."):
            docx_files = list(dir_path.glob("**/*.docx"))

        if docx_files:
            status_text.text(f"Найдено DOCX файлов: {len(docx_files)}")
            for i, docx_path in enumerate(docx_files):
                progress = (i + 1) / len(docx_files)
                progress_bar.progress(progress)

                with st.expander(f"Обработка: {docx_path.name}"):
                    try:
                        docx_parser.process_docx_file(docx_path, output_dir)
                        st.success(f"Успешно обработан!")
                        st.info(f"Результат: {output_dir / docx_path.stem}.json")
                    except Exception as e:
                        st.error(f"Ошибка обработки: {str(e)}")
        else:
            st.info("DOCX файлы не найдены")

    # Обработка Excel
    if process_excel:
        output_dir = output_root / "Обработанные excel"
        output_dir.mkdir(exist_ok=True)

        with st.spinner("Поиск Excel файлов..."):
            excel_files = list(dir_path.glob("**/*.xlsx")) + list(dir_path.glob("**/*.xls"))

        if excel_files:
            status_text.text(f"Найдено Excel файлов: {len(excel_files)}")
            for i, excel_path in enumerate(excel_files):
                progress = (i + 1) / len(excel_files)
                progress_bar.progress(progress)

                with st.expander(f"Обработка: {excel_path.name}"):
                    try:
                        excel_parser.process_excel_file(excel_path, output_dir)
                        st.success("Успешно обработан!")
                        st.info(f"Результат: {output_dir / excel_path.stem}.json")
                    except Exception as e:
                        st.error(f"Ошибка обработки: {str(e)}")
        else:
            st.info("Excel файлы не найдены")

    # Обработка PDF
    if process_pdf:
        output_dir = output_root / "Обработанные pdf"
        output_dir.mkdir(exist_ok=True)

        with st.spinner("Поиск PDF файлов..."):
            pdf_files = list(dir_path.glob("**/*.pdf"))

        if pdf_files:
            status_text.text(f"Найдено PDF файлов: {len(pdf_files)}")
            for i, pdf_path in enumerate(pdf_files):
                progress = (i + 1) / len(pdf_files)
                progress_bar.progress(progress)

                with st.expander(f"Обработка: {pdf_path.name}"):
                    try:
                        pdf_parser.process_pdf_file(pdf_path, output_dir)
                        st.success("Успешно обработан!")
                        st.info(f"Результат: {output_dir / pdf_path.stem}.json")
                    except Exception as e:
                        st.error(f"Ошибка обработки: {str(e)}")
        else:
            st.info("PDF файлы не найдены")

    progress_bar.empty()
    status_text.success("Обработка завершена!")
    st.balloons()


def run_single_file_processing(uploaded_file, file_type):
    with st.spinner("Обработка файла..."):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_dir = Path(temp_dir)
            file_path = temp_dir / uploaded_file.name

            # Сохраняем загруженный файл во временную директорию
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getvalue())

            try:
                # Обработка в зависимости от типа файла
                if file_type == "DOCX":
                    result = docx_parser.extract_document_structure(file_path)
                    result_type = "DOCX"

                elif file_type == "PDF":
                    # Создаем временную папку для результата
                    output_dir = temp_dir / "output"
                    output_dir.mkdir()

                    # Обрабатываем PDF
                    pdf_parser.process_pdf_file(file_path, output_dir)

                    # Читаем результат
                    result_path = output_dir / (file_path.stem + ".json")
                    with open(result_path, "r", encoding="utf-8") as f:
                        result = json.load(f)
                    result_type = "PDF"

                else:  # Excel
                    # Создаем временную папку для результата
                    output_dir = temp_dir / "output"
                    output_dir.mkdir()

                    # Обрабатываем Excel
                    excel_parser.process_excel_file(file_path, output_dir)

                    # Читаем результат
                    result_path = output_dir / (file_path.stem + ".json")
                    with open(result_path, "r", encoding="utf-8") as f:
                        result = json.load(f)
                    result_type = "Excel"

                # Отображаем результат
                st.success(f"Файл успешно обработан! Тип: {file_type}")
                st.subheader("Результат обработки:")
                st.json(result)

                # Кнопка скачивания
                st.download_button(
                    label="Скачать результат в JSON",
                    data=json.dumps(result, ensure_ascii=False, indent=2),
                    file_name=f"{uploaded_file.name}_result.json",
                    mime="application/json"
                )

            except Exception as e:
                st.error(f"Ошибка обработки файла: {str(e)}")


if __name__ == "__main__":
    main()