import pandas as pd

path = r"Steuer.xlsx"
excel = pd.read_excel(path, engine = 'openpyxl')
print(excel)
last_nr = excel['Nr'].values[-1]
new_nr = last_nr+1
new_row =pd.DataFrame([[new_nr,"Schlatt2, Anja",13.00]],columns=["Nr",'Name','Rechnung'])


new_excel = pd.concat([excel,new_row],ignore_index=True)
print(new_excel)