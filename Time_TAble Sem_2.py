
import pandas as pd
import numpy as np

# Load Excel file using pandas
df = pd.read_excel('Test.xlsx', sheet_name='Table 3')

# Convert pandas DataFrame to numpy array
data = np.array(df)

i=1
j=1
while i<10:
    while j<10:
     
     cell_value=str(df.iloc[i,j])

     a=1
     if(cell_value[a]=='/'):
        break
     else:
        print(cell_value)
        j=j+1
    i=i+1
    

