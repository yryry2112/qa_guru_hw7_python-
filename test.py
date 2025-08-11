import os
import zipfile
from io import BytesIO
import pypdf
from openpyxl import load_workbook


# Общая функция для чтения файла из архива
def read_file_from_zip(archive_path, filename):
    with zipfile.ZipFile(archive_path, 'r') as zf:
        with zf.open(filename) as f:
            return f.read()


# Проверка наличия файла в архиве
def assert_file_in_archive(archive_path, filename):
    with zipfile.ZipFile(archive_path, 'r') as zf:
        assert filename in zf.namelist(), f"'{filename}' не найден"


# Тест создания архива
def test_create_archive():
    os.makedirs('my_archives', exist_ok=True)
    archive_path = os.path.join('my_archives', 'archive_file.zip')
    files = ['Файл_CSV.csv', 'Файл_xlsx.xlsx', 'Файл_PDF.pdf']

    with zipfile.ZipFile(archive_path, 'w') as zf:
        for file in files:
            file_path = os.path.join(os.getcwd(), file)
            if os.path.exists(file_path):
                zf.write(file_path, os.path.basename(file_path))
    assert os.path.exists(archive_path)
    for file in files:
        assert_file_in_archive(archive_path, file)


# Проверка CSV
def test_csv():
    archive_path = 'my_archives/archive_file.zip'
    target_file = 'Файл_CSV.csv'
    content = read_file_from_zip(archive_path, target_file).decode('utf-8')
    assert content == 'файл csv'


# Проверка PDF
def test_pdf():
    archive_path = 'my_archives/archive_file.zip'
    pdf_name = 'Файл_PDF.pdf'
    expected_text = 'Файл  пдф  '
    pdf_bytes = read_file_from_zip(archive_path, pdf_name)
    reader = pypdf.PdfReader(BytesIO(pdf_bytes))
    full_text = ''.join(page.extract_text() for page in reader.pages)
    assert full_text == expected_text


# Проверка XLSX
def test_xlsx():
    archive_path = 'my_archives/archive_file.zip'
    xlsx_name = 'Файл_xlsx.xlsx'
    expected_cell_value = 'Файл xlsx'

    xlsx_bytes = read_file_from_zip(archive_path, xlsx_name)
    workbook = load_workbook(filename=BytesIO(xlsx_bytes), read_only=True)
    sheet = workbook.active
    actual_value = sheet['A1'].value
    assert actual_value == expected_cell_value