import string
import pandas as pd
import numpy as np
import openpyxl
teachers_list=[]
classroom_list=[]


# Create a new workbook
workbook = openpyxl.Workbook()

# Select the worksheet you want to edit (by default, there is one called 'Sheet')
worksheet = workbook.active
worksheet.append(["DAY", "DIVISION", "START", "END", "SUBJECT", "BATCH", "CLASSROOM", "TEACHER", "TYPE"])

# Load Excel file using pandas
df = pd.read_excel('/home/darpan/vscode/XLS_MYSQL-converter/Sem_2.xlsx', sheet_name='Final Copy')

# Convert pandas DataFrame to numpy array
data = np.array(df)

for count, day in enumerate(["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY"], start=1):
    i = (count - 1) * 3 + 4
    ascii = 65
    row = (count - 1) * 60 + 2
    while i < 53:
        div = chr(ascii)
        j = 1 if count == 1 else 12
        while j < (9 if count == 1 else 19):
            column = 1
            cell_value = str(data[i][j])
            teacher = str(data[i+1][j])
            class_room = str(data[i+2][j])
            timing = str(data[3][j]).replace(".", ":")
            batch = "0"
            type = "T"
            a, b = 1, 1
            c = len(timing)
            while b < len(timing):
                if timing[b] == "-":
                    start = timing[0:b]
                    end = timing[b+1:c]
                b = b+1
            if "/" in cell_value:
                cell_values = cell_value.split("/")
            else:
                cell_values = [cell_value]
                
            if "/" in teacher:
                teachers = teacher.split("/")
            else:
                teachers = [teacher]
                
            if "/" in class_room:
                class_rooms = class_room.split("/")
            else:
                class_rooms = [class_room]
                
            for cv, t, cr in zip(cell_values, teachers, class_rooms):
                if cv == "nan" or t == "nan" or cr == "nan":
                    cv = "Nan"
                elif cv == "*Incase of theory lecture it will end at 1:10 pm":
                    cv = "Lunch Break"
                else:
                    print(day + "     " + div + "     "+start+"     "+end + "     " + cv +
                          "    " + batch+"     " + cr + "    " + t+"     "+type)
                    worksheet.append([day, div, start, end, cv, batch, cr, t, type])
            j += 1
        ascii += 1
        row += 1
        i += 3

workbook.save('Time_Table_output.xlsx')
