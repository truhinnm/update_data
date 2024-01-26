import io
import os
import msoffcrypto
from users_passwords import *
from win32com.client.gencache import EnsureDispatch
from openpyxl import load_workbook


def decrypt_file(file_path, file_password):
    decrypted_workbook = io.BytesIO()
    with open(file_path, 'rb') as cur_file:
        office_file = msoffcrypto.OfficeFile(cur_file)
        office_file.load_key(password=file_password)
        office_file.decrypt(decrypted_workbook)
    return decrypted_workbook


def find_cell(sheet, string_column, target_column, search_string):
    for cell in sheet[string_column]:
        if cell.value is not None:
            if cell.value == search_string:
                return target_column + str(cell.row)


def edit_workbook(workbook, file_path):
    wb = load_workbook(workbook)
    sheet = wb.active
    sheet["H96"] = 0.05
    wb.save(file_path)


def set_wb_pass(file_dir_path, read_password, write_password):
    xl_file = EnsureDispatch("Excel.Application")
    wb = xl_file.Workbooks.Open(file_dir_path)
    xl_file.DisplayAlerts = False
    wb.SaveAs(file_dir_path, Password=read_password, WriteResPassword=write_password)
    wb.Close()
    xl_file.Quit()


def open_main_tables(mt_path):
    tables = {}
    for m_root, m_dirs, m_files in os.walk(mt_path):
        for m_file in m_files:
            if m_file.endswith('.xlsx'):
                path = os.path.join(m_root, m_file)
                tables[m_file] = str(decrypt_file(path, main_tables[m_file][0]))
    return tables
