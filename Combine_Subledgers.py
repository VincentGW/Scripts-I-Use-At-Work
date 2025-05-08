import xlwings as xw
import os
import inspect

directory = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
output_file = '#22 2022A Project Statements & Subledger Analysis.xlsx'
file_paths = []
template = ""

for file in os.listdir(directory):
    file_path = os.path.join(directory, file)
    if file_path.endswith(".xlsx") or file_path.endswith(".xls"):
        if file == "Subledger Analysis Template.xlsx":
            template = file_path
            pass
        else:
            file_paths.append(file_path)

def combine_excel_files(file_paths, output_file):
    new_wb = xw.Book()
    template_wb = xw.Book(template)
    template_wb.sheets[0].api.Copy(Before = new_wb.sheets[0].api)
    template_wb.close()
    new_wb.sheets['Sheet1'].delete()
    for file in file_paths:
        temp_wb = xw.Book(file)
        cell_value = temp_wb.sheets[0].range('A7').value
        cell_value_trimmed = cell_value.strip()
        project_num = cell_value_trimmed[8:13]
        temp_wb.sheets[0].name = project_num
        temp_wb.sheets[0].api.Copy(After = new_wb.sheets[-1].api)
        selected_range = new_wb.sheets[-1].range('A1:E200')
        selected_range.api.WrapText = False
        exp = 0
        rev_63175 = 0
        rev_63183 = 0
        rev = 0
        exp_list = []
        rev_list = []
        for cell in new_wb.sheets[-1].range('A1:A100'):
            y = cell.row
            highlight_area = new_wb.sheets[-1].range('A' + str(y) +':E' + str(y))
            if '80450' in str(cell.value):
                highlight_area.color = (251, 226, 213)
                exp = "'" + str(project_num) + "'!E" + str(y)
                exp_list.append(exp)
            elif '63175' in str(cell.value):
                highlight_area.color = (251, 226, 213)
                rev_63175 = "'" + str(project_num) + "'!E" + str(y)
                rev_list.append(rev_63175)
            elif '63183' in str(cell.value):
                highlight_area.color = (251, 226, 213)
                rev_63183 = "'" + str(project_num) + "'!E" + str(y)
                rev_list.append(rev_63183)
        temp_wb.close()
        if rev_63175 == 0 and rev_63183 == 0:
            rev = 0
        else:
            rev = "=" + '+'.join(rev_list)
        for cell in new_wb.sheets[0].range('A1:A30'):
            y = cell.row
            revenue = new_wb.sheets[0].range('E' + str(y))
            expense = new_wb.sheets[0].range('F' + str(y))
            if str(project_num) in str(cell.value):
                revenue.value = rev
                expense.value = exp
    new_wb.save(output_file)            
    new_wb.close()

combine_excel_files(file_paths,output_file)
