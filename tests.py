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