from io import StringIO

import numpy as np
import pdfminer
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter#process_pdf
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import TextConverter
from pdfminer.layout import LTTextContainer
from pdfminer.layout import LAParams
from pdfminer.high_level import extract_pages
from pdfminer.high_level import extract_text
import camelot
import tabula
import pandas as pd
pd.options.display.max_rows =  999
pd.options.display.max_columns= 999


class GOT:

    def __init__(self, path):
        self.paragraph =[]
        self.leistungen =[]
        self.tier =[]
        self.ein_fach =[]
        self.zwei_fach =[]
        self.drei_fach =[]
        self.last_paragraph = 0
        self.last_leistung = ""
        self.pdf = path
        self.dfs =tabula.read_pdf(self.pdf, pages='all')

    def check_tables(self):
        #last 6 pages inhaltsverzeichnis mit ziffern
        for i in range(1,len(self.dfs)-6):
            #if self.dfs[i].shape[1] == 4:
            print(self.dfs[i].shape)
            print(self.dfs[i])#.iloc[[0]])
            print("-----------")


    def get_values_per_dataframe(self):
        print(len(self.dfs))
        #print(self.dfs[2])
        self.paragraph.append(10)
        self.paragraph.append(11)
        for i in range(len(self.dfs[2])):
            self.leistungen.append(self.dfs[2].iat[i,0])
            self.tier.append(np.nan)
            self.ein_fach.append(self.dfs[2].iat[i,1])
            self.zwei_fach.append(self.dfs[2].iat[i,2])
            self.drei_fach.append(self.dfs[2].iat[i,3])

        #print(self.dfs[3])
        for i in range(2,len(self.dfs[3])):
            if self.dfs[3].iat[i,2] is np.nan:
                print(self.dfs[3].iloc[[i]])
                continue
            else:
                self.paragraph.append(int(self.dfs[3].iat[1,0]))
                self.leistungen.append(self.dfs[3].iat[1,1])
                self.tier.append(self.dfs[3].iat[i,1])
                self.ein_fach.append(self.dfs[3].iat[i,2])
                self.zwei_fach.append(self.dfs[3].iat[i,3])
                self.drei_fach.append(self.dfs[3].iat[i,4])

        #print(self.dfs[4])
        for i in range (len(self.dfs[4])):
            self.paragraph.append(21)
            self.leistungen.append("Folgeuntersuchung im gleichen Behandlungsfall mit Beratung")
            self.tier.append(self.dfs[4].iat[i,0])
            self.ein_fach.append(self.dfs[4].iat[i,1])
            self.zwei_fach.append(self.dfs[4].iat[i,2])
            self.drei_fach.append(self.dfs[4].iat[i,3])
        #print(self.dfs[5])
        for i in range (len(self.dfs[5])):
            self.paragraph.append(21)
            self.leistungen.append("Folgeuntersuchung im gleichen Behandlungsfall mit Beratung")
            self.tier.append(self.dfs[5].iat[i,0])
            self.ein_fach.append(self.dfs[5].iat[i,1])
            self.zwei_fach.append(self.dfs[5].iat[i,2])
            self.drei_fach.append(self.dfs[5].iat[i,3])

        self.last_paragraph = 21
        self.last_leistung = "Folgeuntersuchung im gleichen Behandlungsfall mit Beratung"
        for dataframe in range(6,len(self.dfs)-6):
            print(dataframe)
            print(self.dfs[dataframe].shape)

            #dataframe with 4 cols
            if self.dfs[dataframe].shape[1] == 4:
                for row in range (len(self.dfs[dataframe])):
                    if self.dfs[dataframe].iat[row,2] is np.nan:
                        continue
                    self.leistungen.append(self.last_leistung)
                    self.paragraph.append(int(self.last_paragraph))
                    self.tier.append(self.dfs[dataframe].iat[row,0])
                    self.ein_fach.append(self.dfs[dataframe].iat[row,1])
                    self.zwei_fach.append(self.dfs[dataframe].iat[row,2])
                    self.drei_fach.append(self.dfs[dataframe].iat[row,3])
            #dataframe with 5 cols
            if self.dfs[dataframe].shape[1] == 5:
                if self.dfs[dataframe].iat[0,0] is np.nan and self.dfs[dataframe].iat[0,1] is np.nan:
                    i = 1
                else:
                    i = 0
                self.last_paragraph = self.dfs[dataframe].iat[i,0]
                for row in range (i,len(self.dfs[dataframe])):
                    if self.dfs[dataframe].iat[row,2] is np.nan:
                        continue
                    if self.dfs[dataframe].iat[row,0] is not np.nan:
                        self.last_paragraph = self.dfs[dataframe].iat[row,0]
                        self.last_leistung = self.dfs[dataframe].iat[row,1]

                    self.leistungen.append(self.last_leistung)
                    self.paragraph.append(int(self.last_paragraph))
                    self.tier.append(self.dfs[dataframe].iat[row,0])
                    self.ein_fach.append(self.dfs[dataframe].iat[row,1])
                    self.zwei_fach.append(self.dfs[dataframe].iat[row,2])
                    self.drei_fach.append(self.dfs[dataframe].iat[row,3])

            #dataframe with 6 cols
            if self.dfs[dataframe].shape[1] == 6:
                if self.dfs[dataframe].iat[0,0] is np.nan and self.dfs[dataframe].iat[0,1] is np.nan:
                    i = 1
                else:
                    i = 0
                self.last_paragraph = self.dfs[dataframe].iat[i,0]
                for row in range (1,len(self.dfs[dataframe])):
                    if self.dfs[dataframe].iat[row,1] is np.nan:
                        continue
                    if self.dfs[dataframe].iat[row,0] is not np.nan:
                        self.last_paragraph = self.dfs[dataframe].iat[row,0]
                        self.last_leistung = self.dfs[dataframe].iat[row,1]

                    self.leistungen.append(self.last_leistung)
                    self.paragraph.append(int(self.last_paragraph))
                    self.tier.append(np.nan)
                    self.ein_fach.append(self.dfs[dataframe].iat[row,3])
                    self.zwei_fach.append(self.dfs[dataframe].iat[row,4])
                    self.drei_fach.append(self.dfs[dataframe].iat[row,5])


    def create_leistungen(self):
        #create new dataframe with all Leistungen
        d = {'Paragraph':self.paragraph,'Leistung': self.leistungen,'Tier': self.tier, '1-fach': self.ein_fach, '2-fach': self.zwei_fach, '3-fach': self.drei_fach}
        df = pd.DataFrame(data=d)
        print(df)
    # Tabelle [1]: Zeitgebühr
    # Tabelle [3]: Teil A Grundleistung §10,§11
    # Tabelle [4], [5], [7], [8], [9], [10]: Teil A Grundleistung 20
    # Tabelle [14]:Teil A Grundleistung §60
    # Tabelle [15]: Teil B Besondere Leistugen: Bescheinigungen/Gutachten
    # Tabelle [16]: Teil B Besondere Leistugen: Sonstige Untersuchungen


    # for page_layout in extract_pages(pdf, page_numbers=[16,17,18]):
    #     print(page_layout)
    #     for element in page_layout:
    #         if isinstance(element, LTTextContainer):
    #             print(element.get_text())

    def exampleDF(self):
        paragraph=[10,11,20,20,20,20,20,20,20,20,20,20,20,20,20,20,20]
        leistungen= ["Beratung im einzelnen Fall ohne Untersuchung","Eingehende Anamneseerhebung oder Beratung","Allgemeine Untersuchung mit Beratung","Pferd","Rind","Zuchtschwein","Mastschwein","Kalb","Ferkel,Schaf,Ziege","Hund","Katze","Nutzgeflügel","Fische","Pelztiere, sonstige Farmtiere","Wildtiere, Zootiere","Heimtiere","Ziergeflügel"]
        #tier= ["","","","Pferd","Rind","Zuchtschwein","Mastschwein","Kalb","Ferkel,Schaf,Ziege","Hund","Katze","Nutzgeflügel","Fische","Pelztiere, sonstige Farmtiere","Wildtiere, Zootiere","Heimtiere","Ziergeflügel"]
        ein_fach = [7.04,19.24,np.NAN,19.24,12.84,12.84,9.62,12.84,7.71,13.47,8.98,3.21,15.39,15.39,23.09,9.62,7.04]
        zwei_fach = [14.08,38.48,np.NAN,38.48,25.68,25.68,19.24,25.68,15.42,26.94,17.96,6.42,30.78,30.78,46.18,19.24,14.08]
        drei_fach = [21.12,57.72,np.NAN,57.72,38.52,38.52,28.86,38.52,23.13,40.41,26.94,9.63,46.17,46.17,69.27,28.86,21.12]
        #print(len(paragraph))
        #print(len(leistungen))
        #print(len(tier))
        #print(len(ein_fach))
        #print(len(zwei_fach))
        #print(len(drei_fach))
        d = {'Paragraph':paragraph,'Leistung': leistungen, '1-fach': ein_fach, '2-fach': zwei_fach, '3-fach': drei_fach}
        df = pd.DataFrame(data=d)
        return df

    #red pdf file GOT
    #PDF: r'D:\Rechnung_template_projekt\GOT_2020.pdf'
#pdf = r'D:\Rechnung_template_projekt\GOT_2020.pdf'
#got = GOT(pdf)
#got.check_tables()
#got.get_values_per_dataframe()
#got.create_leistungen()
#result =got.exampleDF()
#print(result)

def exampleDF():
    paragraph=["10","11",
               "20a","20b","20ca","20cb","20d","20e","20f","20g","20h","20i","20j","20k","20l","20m",
               "21a","21b","21ca","21cb","21d","21e","21f","21g","21h","21i","21j","21k","21l","21m",
               "22",
               "60","60",
               "101","102","103","104","105","106",
               "302","303a","303b","305a","305b","305c",
               "501",
               "502a","502b","502c","502d","502e","502f",
               "503",
               "504aa","504ab","504ac","504ac","504b","504ba","504c","504c","504d","504e","504e","504f","504g",
               "505a","505b","505c","505c","505d",
               "510a","510b","510c","510d",
               "602a", "602e", "602g", "602h",
               "509","509",
               "A1","A2",
               "B 5.1","B 5.1",
               "Bl1","Bl2","Bl2","Bl3","Bl3","Bl4","Bl4",
               "Bl5a","Bl5a","Bl5b","Bl5b",
               "Bl7a","Bl7b","Bl7c","Bl7d","Bl7e",
               "G 2.4 aa","G 2.4 ab","G 2.4e",
               "G 2.16a","G 2.16a","G 2.16d","G 2.16d","G 2.16e","G 2.16e",
               "G 5.3a","G 5.3b","G 5.4a","G 5.4b","G 5.7a","G 5.7b","G 5.7",
               "Z4.1a","Z4.1b","Z4.1ca","Z4.1cb","Z4.1d","Z4.1d","Z4.1d",
               "Z4.3a","Z4.3e","Z4.3f"]
    leistungen= ["Beratung im einzelnen Fall ohne Untersuchung (auch schriftlich oder fernmündlich)","Eingehende Anamneseerhebung oder Beratung",
                 "Allgemeine Untersuchung mit Beratung","Allgemeine Untersuchung mit Beratung","Allgemeine Untersuchung mit Beratung","Allgemeine Untersuchung mit Beratung","Allgemeine Untersuchung mit Beratung","Allgemeine Untersuchung mit Beratung","Allgemeine Untersuchung mit Beratung","Allgemeine Untersuchung mit Beratung","Allgemeine Untersuchung mit Beratung","Allgemeine Untersuchung mit Beratung","Allgemeine Untersuchung mit Beratung","Allgemeine Untersuchung mit Beratung","Allgemeine Untersuchung mit Beratung","Allgemeine Untersuchung mit Beratung",
                 "Folgeuntersuchung im gleichen Behandlungsfall mit Beratung","Folgeuntersuchung im gleichen Behandlungsfall mit Beratung","Folgeuntersuchung im gleichen Behandlungsfall mit Beratung","Folgeuntersuchung im gleichen Behandlungsfall mit Beratung","Folgeuntersuchung im gleichen Behandlungsfall mit Beratung","Folgeuntersuchung im gleichen Behandlungsfall mit Beratung","Folgeuntersuchung im gleichen Behandlungsfall mit Beratung","Folgeuntersuchung im gleichen Behandlungsfall mit Beratung","Folgeuntersuchung im gleichen Behandlungsfall mit Beratung","Folgeuntersuchung im gleichen Behandlungsfall mit Beratung","Folgeuntersuchung im gleichen Behandlungsfall mit Beratung","Folgeuntersuchung im gleichen Behandlungsfall mit Beratung","Folgeuntersuchung im gleichen Behandlungsfall mit Beratung","Folgeuntersuchung im gleichen Behandlungsfall mit Beratung",
                 "Eilbesuche",
                 "Überwachung von Intesivpatienten bei Tag","Überwachung von Intesivpatienten bei Nacht",
                 "Impfbescheinigung","Sonstige Bescheinigung","Einfache Guthaben","Ausführliche Gutachten","Rezeptgebühr für Wiederholungsrezept ohne Beratung bei einer Inanspruchnahme des Tierarztes soweit keine weiteren Leistungen berechnet werden","Verschreibung eines Fütterungsarzneimittels",
                 "Bearbeitung von Proben zum Versand","Bakteriologische Untersuchung einfacher Art ohne Resistenzbestimmung","Bakteriologische Untersuchung einfacher Art mit Resistenzbestimmung","Mikroskopische Untersuchung, Nativpräparat, auch Harnsediment","Mikroskopische Untersuchung mit Anwendung einfacher Färbeverfahren","Mikroskopische Untersuchung mit Anwendung besonderer (differenzierender) Färbeverfahren",
                 "Eingeben von Medikamenten",
                 "Tötung (Euthanasie) durch Injektion","Tötung (Euthanasie) durch Injektion","Tötung (Euthanasie) durch Injektion","Tötung (Euthanasie) durch Injektion","Tötung (Euthanasie) durch Injektion","Tötung (Euthanasie) durch Injektion",
                 "Implantation eines Arzneimittels",
                 "Injektion, Instillation, Infusion: subkutan, intrakutan, intramuskulär, intraingluvial","Injektion, Instillation, Infusion: subkutan, intrakutan, intramuskulär, intraingluvial","Injektion, Instillation, Infusion: subkutan, intrakutan, intramuskulär, intraingluvial","Injektion, Instillation, Infusion: subkutan, intrakutan, intramuskulär, intraingluvial","Injektion, Instillation, Infusion: intravenös, intratracheal, subkonjunktival","Injektion, Instillation, Infusion: Venenkatheter einlegen","Injektion, Instillation, Infusion: extradural, intraartikulär, intrabulbär","Injektion, Instillation, Infusion: extradural, intraartikulär, intrabulbär","Injektion, Instillation, Infusion: intrarektal, intrapräputial, intravaginal","Injektion, Instillation, Infusion: intrauterin, intraabdominal","Injektion, Instillation, Infusion: intrauterin, intraabdominal","Injektion, Instillation, Infusion: intranasal","Injektion, Instillation, Infusion: Infusion",
                 "Kennzeichnen: Einziehen von Ohrmarken","Kennzeichnen: Tätowieren","Kennzeichnen: Implantation eines Mikrochips","Kennzeichnen: Implantation eines Mikrochips","Kennzeichnen: Ablesen eines Mikrochips",
                 "Verband anlegen/abnehmen einfach","Verband anlegen/abnehmen schwierig","Verband anlegen/abnehmen Robert-Jones-Verband", "Verband anlegen/abnehmen Gipsverband oder ähnl. Schienung",
                 "Schutzimpfungen","Schutzimpfungen","Schutzimpfungen","Schutzimpfungen",
                 "Tupferprobenentnahme","Tupferprobenentnahme gynäkologisch",
                 "Eingehende Untersuchung, einzelner Organe","Inhalation",
                 "Kürzen der Krallen","Kürzen der Krallen",
                 "Aderlass","Blut-chemische Untersuchung: photometrische Einzelparameter","Blut-chemische Untersuchung:pro Parameter bei mehr als 3 Parameter","Blutdruckmessung unblutig", "Blutdruckmessung operativ","Blutgasanalyse erste Messung","Blutgasanalyse jede weitere Messung",
                 "Blutprobenentnahme Einzeltier: venös", "Blutprobenentnahme Einzeltier: arteriell", "Blutprobenentnahme Reihenentnahme pro Tier","Rind Laufstall bzw. Ammenkuhhaltung",
                 "Blutuntersuchung: Blutausstrich mit Färbung und Differenzierung", "Blutuntersuchung: Blusenkungsreaktion","Blutuntersuchung: Hämatokriwert","Blutuntersuchung: Leukozytenzählung,Erythrozytenzählung,Thrombozytenzählung","Blutuntersuchung: Blutungs- und/oder Gerinnungszeit",
                 "Geburtshilfe einfach", "Geburtshilfe schwierig","Geburtshilfe",
                 "Trächtigkeitsuntersuchung","Trächtigkeitsuntersuchung einschließlich Ultraschall","Trächtigkeitsuntersuchung","Trächtigkeitsuntersuchung einschließlich Ultraschall","Trächtigkeitsuntersuchung","Trächtigkeitsuntersuchung einschließlich Ultraschall",
                 "Kastration und Sterilisation","Kastration und Sterilisation","Kastration und Sterilisation","Kastration und Sterilisation","Kastration und Sterilisation","Kastration und Sterilisation","Kastration und Sterilisation",
                 "Lokalanästhesie", "Leitungsanästhesie", "Epidurale/intrakartikuläre Anästhesie","Epidurale/intrakartikuläre Anästhesie","Neuraltherapie, systemisch, intravenös", "Neuraltherapie, lokal (Gelosen, Narben)", "Neuraltherapie, segmental",
                 "Injektionsnarkose","Injektionsnarkose","Injektionsnarkose"]
    tier= [np.NAN,np.NAN,
           "Pferd","Rind","Zuchtschwein","Mastschwein","Kalb","Ferkel,Schaf,Ziege","Hund","Katze","Nutzgeflügel","Fische","Pelztiere, sonstige Farmtiere","Wildtiere, Zootiere","Heimtiere","Ziergeflügel",
           "Pferd","Rind","Zuchtschwein","Mastschwein","Kalb","Ferkel,Schaf,Ziege","Hund","Katze","Nutzgeflügel","Fische","Pelztiere, sonstige Farmtiere","Wildtiere, Zootiere","Heimtiere","Ziergeflügel",
           np.NAN,
           np.NAN,np.NAN,
           np.NAN,np.NAN,np.NAN,np.NAN,np.NAN,np.NAN,
           np.NAN,np.NAN,np.NAN,np.NAN,np.NAN,np.NAN,
           np.NAN,
           "Pferd","Hund","Katze","Rind","Schwein, Kalb, Schaf, Ziege", "Tiere im Säuglingsalter, Heimtiere, Ziergeflügel, Pelztiere",
           np.NAN,
           "Hund, Katze, Pferd", "Heimtiere, Wildtiere, Zootiere, Rind, Schwein, Schaf, Geflügel, Ziege","Lamm, Ferkel bis 5 Tiere, je Tier", "Lamm, Ferkel jedes weitere Tier",np.NAN,np.NAN,"Hund, Katze, Pferd, Wildtiere, Zootiere", "Sonstige",np.NAN,"Alles außer Pferd","Pferd",np.NAN,np.NAN,
           np.NAN,np.NAN,np.NAN,"ab dem 5. Tier",np.NAN,
           np.NAN,np.NAN,np.NAN,np.NAN,
           "Pferd", "Pelztier", "Hund, Katze", "Bestandsgebühr",
           np.NAN,np.NAN,
           np.NAN,np.NAN,
           "Hund, Katze, alle Extremitäten", "Heimtiere, Geflügel",
           np.NAN,np.NAN,np.NAN,np.NAN,np.NAN,np.NAN,np.NAN,
           np.NAN,np.NAN,"Pferd,Rind,Schwein,Schaf,Fische",np.NAN,
           np.NAN,np.NAN,np.NAN,np.NAN,np.NAN,
           "Pferd","Pferd","Hund, Katze",
           "Pferd","Pferd","Hund,Katze","Hund,Katze","Heimtiere","Heimtiere",
           "Hund männlich", "Hund weiblich", "Katze männlich", "Katze weiblich","Männlich Einzel: Kanninchen/Heimtiere", "Männlich,jedes weitere: Kanninchen/Heimtiere","Weiblich: Kanninchen/Heimtiere",
           np.NAN,np.NAN,"Hund,Katze,Pferd,Wildtiere,Zootiere","Heimtiere,Rind,Schwein,Schaf,Ziege",np.NAN,np.NAN,np.NAN,
           "Pferd,Wildtiere,Zootiere","Hund,Katze","Kleine Heimtiere,Geflügel"]
    ein_fach = [7.04,19.24,
                19.24,12.84,12.84,9.62,12.84,7.71,13.47,8.98,3.21,15.39,15.39,23.09,9.62,7.04,
                15.39,6.41,10.26,7.71,10.26,5.13,10.90,7.71,2.58,6.41,12.19,17.96,7.71,5.77,
                25.65,
                19.24,38.48,
                3.85,6.41,25.65,83.38,1.92,6.41,
                6.41,6.41,9.62,6.41,7.71,12.84,
                2.58,
                92.37, 19.24,19.24,22.46,22.46,6.41,
                5.77,
                5.77, 3.85, 1.28, 0.64, 7.71,15.39,19.24,9.62,5.13,6.41,16.03,3.21,12.84,
                1.6,6.41,6.41,5.13,3.21,
                5.13,7.71,19.24,51.31,
                4.49,1.28,4.49,16.03,
                5.13,12.84,
                9.62,9.62,
                6.41,5.13,
                23.09,5.77,3.85,9.62,32.07,9.62,6.41,
                6.41,9.62,3.85,7.71,
                9.62,4.49,5.13,4.49,9.62,
                96.20,160.34,38.48,
                22.46,38.48,12.84,38.48,9.62,38.48,
                51.31,160.34,19.24,57.72,19.24,12.84,57.72,
                7.71,9.62,19.24,9.62,16.03,16.03,25.65,
                38.48,19.24,6.41]
    zwei_fach = [14.08,38.48,
                 38.48,25.68,25.68,19.24,25.68,15.42,26.94,17.96,6.42,30.78,30.78,46.18,19.24,14.08,
                 30.78,12.82,20.52,15.42,20.52,10.26,21.80,15.42,5.16,12.82,24.38,35.92,15.42,11.54,
                 51.30,
                 38.48,76.96,
                 7.70,12.82,51.30,166.76,3.84,12.82,
                 12.82,12.82,19.24,12.82,15.42,25.68,
                 5.16,
                 184.74,38.48,38.48,44.92,44.92,12.82,
                 11.54,
                 11.54, 7.70,2.56,1.28,15.42,30.78,38.48,19.24,10.26,12.82,32.06,6.42,25.68,
                 3.2,12.82,12.82,10.26,6.42,
                 10.26,15.42,38.48,102.62,
                 8.98,2.56,8.98,32.06,
                 19.26,25.68,
                 19.24,19.24,
                 12.82,10.26,
                 46.18,11.54,7.7,19.24,64.14,19.24,12.82,
                 12.82,19.24,7.7,15.42,
                 19.24,8.98,10.26,0.98,19.24,
                 192.4,320.68,76.96,
                 44.92,76.96,25.68,76.96,19.24,76.96,
                 102.62,320.68,38.48,115.44,38.48,25.68,115.44,
                 15.42,19.24,38.48,19.23,32.06,32.06,51.30,
                 76.96,38.48,12.82]
    drei_fach = [21.12,57.72,
                 57.72,38.52,38.52,28.86,38.52,23.13,40.41,26.94,9.63,46.17,46.17,69.27,28.86,21.12,
                 46.17,19.23,30.78,23.13,30.78,15.39,32.70,23.13,7.74,19.23,36.57,53.88,23.13,17.31,
                 76.95,
                 57.72,115.44,
                 11.55,19.23,76.95,250.14,5.76,19.23,
                 19.23,19.23,28.86,19.23,23.13,38.52,
                 7.74,
                 277.11,57.72,57.72,67.38,67.38,19.23,
                 17.31,
                 17.31,11.55,3.84,1.92,23.13,46.17,57.72,28.86,15.39,19.23,48.09,9.63,38.52,
                 4.8,19.23,19.23,15.39,9.63,
                 15.39,23.13,57.72,153.93,
                 13.47,3.84,13.47,48.09,
                 15.39,38.52,
                 28.86,28.86,
                 19.23,15.39,
                 69.27,17.31,11.55,28.86,96.21,28.86,19.23,
                 19.23,28.86,11.55,23.13,
                 28.86,13.47,15.39,13.47,28.86,
                 288.6,481.02,115.44,
                 67.38,115.44,38.52,115.44,28.86,115.44,
                 153.93,481.02,57.72,173.16,57.72,38.52,173.16,
                 23.13,28.86,57.72,28.86,48.09,48.09,76.95,
                 115.44,57.72,19.23]
    # print(len(paragraph))
    # print(len(leistungen))
    # print(len(tier))
    # print(len(ein_fach))
    # print(len(zwei_fach))
    # print(len(drei_fach))


    d = {'Paragraph':paragraph,'Leistung': leistungen,'Tier': tier, '1-fach': ein_fach, '2-fach': zwei_fach, '3-fach': drei_fach}
    df = pd.DataFrame(data=d)
    return df


# df = exampleDF()
# print(df)

