import string
import pandas as pd
import numpy as np
import openpyxl

# Create a new workbook
workbook = openpyxl.Workbook()

# Select the worksheet you want to edit (by default, there is one called 'Sheet')
worksheet = workbook.active
worksheet.append(["DAY", "DIVISION", "START", "END", "SUBJECT",
                 "BATCH", "CLASSROOM", "TEACHER", "TYPE"])

# Load Excel file using pandas
df = pd.read_excel('C:/Users/mywor/OneDrive\Desktop/programming/opensource/XLS_MYSQL-converter/Sem_2.xlsx',sheet_name='Final Copy')

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
            while a < len(cell_value):
                if cell_value[a] == "/" or cell_value[a] == "(":
                    cell_value = "nan"
                    break
                a = a+1
            if cell_value == "nan" or teacher == "nan" or class_room == "nan":
                cell_value = "Nan"
            elif cell_value == "*Incase of theory lecture it will end at 1:10 pm":
                cell_value = "Lunch Break"
            else:
                print(day + "     " + div + "     "+start+"     "+end + "     " + cell_value +
                      "    " + batch+"     " + class_room + "    " + teacher+"     "+type)
                worksheet.append(
                    [day, div, start, end, cell_value, batch, class_room, teacher, type])
            j += 1
        ascii += 1
        row += 1
        i += 3

workbook.save('C:/Users/mywor/OneDrive/Desktop/programming/opensource/XLS_MYSQL-converter/Time_Table_output.xlsx')
