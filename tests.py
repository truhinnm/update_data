# for root, dirs, files in os.walk(r"C:\Users\nick_\Desktop\kpi formulas"):
#     for file in files:
#         if file.endswith('.xlsx') and not file.startswith('~$'):
#             if ("KPI_Архив" not in root) & ("ДЕМО" not in root) & ("Общая таблица" not in root):
#                 filepath = os.path.join(root, file)
#                 print(file[:-5])
#                 print(filepath)
#                 wb = load_workbook(filepath)
#                 ws = wb.active
#                 target = find_cell(ws, "A", "A", "Исполнительская дисциплина")
#                 if target is not None:
#                     print(target)
#                     print(ws[target])
#                 else:
#                     print("None")
#                 password = "Amanat" + str(random.randint(1000, 9999))
#                 print('"' + file + '": ["' + password + '" ,"' + password + 'Ed"], ')

# start = time.time()
# main_path = r"C:\Users\nick_\Desktop\kpi formulas\Общая таблица"
# m_tables = open_main_tables(main_path)
# id_wb = load_workbook(m_tables["Исполнительская дисциплина.xlsx"], data_only=True)
# td_wb = load_workbook(m_tables["Трудовая дисциплина.xlsx"], data_only=True)
# ot_wb = load_workbook(m_tables["KPI общая таблица.xlsx"])
#
# ind_path = r"C:\Users\nick_\Desktop\kpi formulas\Директор по развитию\Трушин Н.М..xlsx"
# ind_wb = load_workbook(ind_path)
# ind_val_wb = load_workbook(ind_path, data_only=True)
#
# transfer_data(id_wb, "Оценка", ind_wb, "Трушин Н.М..xlsx", "Исполнительская дисциплина")
# transfer_data(td_wb, "Статистика", ind_wb, "Трушин Н.М..xlsx", "Трудовая дисциплина")
# transfer_to_main("Трушин Н.М..xlsx", ot_wb, ind_wb, ind_val_wb)
# ind_wb.save(ind_path)
# ot_wb.save(os.path.join(main_path, "KPI общая таблица.xlsx"))
# set_wb_pass(os.path.join(main_path, "KPI общая таблица.xlsx"),
#             main_tables["KPI общая таблица.xlsx"][0],
#             main_tables["KPI общая таблица.xlsx"][1])
# end = time.time()
# print(end - start)
