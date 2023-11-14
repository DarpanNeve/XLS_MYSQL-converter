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
df = pd.read_excel('C:/Users/mywor/OneDrive/Desktop/programming/opensource/XLS_MYSQL-converter/Sem_2_old.xlsx', sheet_name='Final Copy')

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
            cell_value = str(data[i, j])
            teacher = str(data[i + 1, j])
            class_room = str(data[i + 2, j])
            timing = str(data[3, j]).replace(":", ".")
            batch = "0"
            type = "T"
            
            # Handle the case where timing has multiple parts separated by "-"
            if "-" in timing:
                start, end = timing.split("-")
            else:
                start = timing
                end = timing

            # Get the value of the next cell
            next_cell_value = str(data[i, j + 1])

            # Check if the next cell is empty and the current cell is a practical
            if next_cell_value == "nan" and "/" in cell_value:
                # If it is, update the end time to the end time of the next cell
                next_timing = str(data[3, j + 1]).replace(":", ".")
                if "-" in next_timing:
                    end = next_timing.split("-")[1]

            if "/" in cell_value:
                cell_values = cell_value.split("/")
                batch = str(len(cell_values))
                type = "P"
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
                if cv == "*Incase of theory lecture it will end at 1:10 pm":
                    cv = "Lunch Break"
                elif cv == "Life Skills":
                    j += 1
                elif cv == "Theory lecture will end at 12:10pm":
                    cv = "Lunch Break"
                elif "1" in cv:
                    batch = "1"
                elif "2" in cv:
                    batch = "2"
                elif "3" in cv:
                    batch = "3"
                    j += 1
                elif cv == "nan" or t == "nan" or cr == "nan":
                    cv = " "
                    t = " "
                    cr = " "
                print(day + "     " + div + "     " + start + "     " + end + "     " + cv + "    " + batch + "     " + cr + "    " + t + "     " + type)
                worksheet.append([day, div, start, end, cv, batch, cr, t, type])
            j += 1
        ascii += 1
        row += 1
        i += 3

workbook.save('C:/Users/mywor/OneDrive/Desktop/programming/opensource/XLS_MYSQL-converter/Time_Table_output.xlsx')
