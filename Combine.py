import xlwings as xw
import os
import inspect

directory = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
output_file = 'Combined_Files.xlsx'
file_paths = []

for file in os.listdir(directory):
    file_path = os.path.join(directory, file)
    if file_path.endswith(".xlsx") or file_path.endswith(".xls"):
        if file == "Subledger Analysis Template.xlsx":
            pass
        else:
            file_paths.append(file_path)

print(file_paths)

def combine_excel_files(file_paths, output_file):
    new_wb = xw.Book()
    for file in file_paths:
        temp = xw.Book(file)
        temp.sheets[0].api.Copy(Before = new_wb.sheets[0].api)
        cell_value = temp.sheets[0].range('A7').value
        cell_value_trimmed = cell_value.strip()
        project_num = cell_value_trimmed[8:13]
        new_wb.sheets[0].name = project_num        
        selected_range = new_wb.sheets[0].range('A1:E200')
        selected_range.api.WrapText = False
        exp_80450 = 0
        rev_63175 = 0
        rev_63183 = 0
        for cell in new_wb.sheets[0].range('A1:A100'):
            y = cell.row
            area = new_wb.sheets[0].range('A' + str(y) +':E' + str(y))
            if '80450' in str(cell.value):
                area.color = (251, 226, 213)
                exp_80450 = new_wb.sheets[0].range('E' + str(y)).value
                print('Project: ',project_num, 'Rev: ',exp_80450,'Exp: ',rev_63175 + rev_63183)
            elif '63175' in str(cell.value):
                area.color = (251, 226, 213)
                rev_63175 = new_wb.sheets[0].range('E' + str(y)).value
                print('Project: ',project_num, 'Rev: ',exp_80450,'Exp: ',rev_63175 + rev_63183)
            elif '63183' in str(cell.value):
                area.color = (251, 226, 213)
                rev_63183 = new_wb.sheets[0].range('E' + str(y)).value
                print('Project: ',project_num, 'Rev: ',exp_80450,'Exp: ',rev_63175 + rev_63183)
            
        temp.close()
    new_wb.sheets['Sheet1'].delete()
    new_wb.save(output_file)
    new_wb.close()

combine_excel_files(file_paths,output_file)