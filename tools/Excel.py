import pandas as pd
pd.options.display.max_rows = 999
pd.options.display.max_columns = 999

#excel = pd.read_excel(path, engine = 'openpyxl')
#path = r"2021.xlsx"
#last_nr = excel['Nr'].values[-1]
#new_nr = last_nr+1
#new_row =pd.DataFrame([[new_nr,"Schlatt2, Anja",13.00]],columns=["Nr",'Name','Rechnung'])
#new_excel = pd.concat([excel,new_row],ignore_index=True)


class Tax:
    def __init__(self,month_int):
        dic = {"01":"Januar","02":"Februar","03":"März","04":"April","05":"Mai","06":"Juni",
               "07":"Juli","08":"August","09":"September","10":"Oktober",
               "11":"November","12":"Dezember"}
        month = dic.get(month_int)
        xls = pd.ExcelFile(r'2021.xlsx')
        self.df = pd.read_excel(xls, month)


    def add_billRow(self,index):
        new_line = pd.DataFrame({"ReNr": 30.0, "Datum": 1.3,"Name":None,"Tier": None,"Rechnung":None,"Anteil Medikamente":None,"19%MWSt":None,"7% MWSt":None,"Dkm":None,"Besuche":None,"Anteil.Laborkosten brutto, vom Gesamtbetrag abziehen!": None,"bar":None,"Fahrtkosten":None,"Praxisanteil":None}, index=[3])
        df2 = pd.concat([self.df.iloc[:2], new_line, self.df.iloc[2:]]).reset_index(drop=True)
