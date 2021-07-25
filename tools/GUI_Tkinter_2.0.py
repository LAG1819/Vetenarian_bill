import tkinter.filedialog
from tkinter import *
from tkinter import ttk
import tkinter.simpledialog as sd

import Excel as Excel
import numpy as np
from ttkthemes import ThemedTk
import ttkbootstrap

import GOT_Parser
import Template
from docx import Document
from tkcalendar import DateEntry
import ttkSimpleDialog
import Patient
from datetime import datetime
import json
import os


class GUI:

    def __init__(self, root, got):

        self.r = root
        self.got = got
        #self.r.state('zoomed') #vollbild
        #self.pathc = r'C:\Users\Pepi\AppData\Local\Rechnung_template\counter.json'
        #self.pathk = r'C:\Users\Pepi\AppData\Local\Rechnung_template\kontaktdaten.json'
        self.pathc= r"counter.json"
        self.pathk= r'kontaktdaten.json'

        self.r.title("Tierarzt Rechnung")

        self.bill_counter = 1
        self.patientenliste = Patient.Patientenliste()

        self.vorname = None
        self.frame = None
        self.befund = None
        self.befundsliste = []
        self.tiername = None
        self.plz = None
        self.tierart = None
        self.mail = None
        self.hausnummer = None
        self.nachname = None
        self.current = None
        self.date_behandlung=None
        self.date_month = "00"
        self.several_counter = 1
        self.werktag = None
        self.woende = None
        self.feiertag = None
        self.notdienst = None
        self.tag = IntVar()

        self.row_leistugen = 5
        self.leistungen =[]
        self.inputWerte_leistungen=[]
        self.all_entry = []
        self.boxes =[]
        self.labels =[]
        self.labels2 = []
        self.anrede_spin_box = None
        self.autocompleteName =[]
        self.auto_listbox = []

        #kontaktdaten for autocomplete
        self.autoCompleteList =[]

    def start(self):

        self.load_allFiles()
        tax = Excel.Tax(self.date_month)
        style = ttkbootstrap.Style()

        container1 = ttk.Frame(self.r, style = 'TFrame')
        container2 = ttk.Frame(self.r,style = 'TFrame')
        container3 = ttk.Frame(self.r,style = 'TFrame')

        canvas2 = Canvas(container2)
        scrollbar2 = ttk.Scrollbar(container2, orient=VERTICAL, command=canvas2.yview)
        scrollable_frame2 = ttk.Frame(canvas2)
        scrollable_frame2.bind("<Configure>",lambda e: canvas2.configure(scrollregion=canvas2.bbox("all")))
        canvas2.create_window((0, 0), window=scrollable_frame2, anchor="nw")
        canvas2.configure(yscrollcommand=scrollbar2.set)

        self.create_customer_input(container1)
        container1.pack(expand=False, side ="top",fill="both",pady=2,padx=5)#,fill="both", expand=True)
        self.create_leistungen_input(container1,container2,container3)

        container2.pack(expand=False, side ="top",fill="both",pady=2,padx=5)
        container3.pack(expand=False, side ="top",fill="both",pady=2,padx=5)

        self.r.mainloop()

    def create_customer_input(self,frame_customer):
        ttk.Label(frame_customer, text='Tierbesitzer:',font=("", 11, 'bold', 'underline'),style='primary.TLabel').grid(sticky = W, row = 1,pady=2)

        self.current = StringVar()
        self.anrede_spin_box= ttk.Spinbox(frame_customer, from_ =0, to =1, values = ("Herr", "Frau"), textvariable = self.current,wrap = True, width = 8)
        self.anrede_spin_box.grid(row=2, column = 1, sticky = W,pady=2)

        ttk.Label(frame_customer, text='Anrede:', anchor = "w",style='primary.TLabel' ).grid(row=2, column = 0,sticky = "w",pady=2)
        ttk.Label(frame_customer, text='Nachname:',anchor="w",style='primary.TLabel').grid(row=3, column = 0,sticky = "w",pady=2)
        ttk.Label(frame_customer, text='Vorname:', anchor = "w" ,style='primary.TLabel').grid(row=4, column = 0,sticky = "w",pady=2)
        ttk.Label(frame_customer, text='Straße, Hausnr.: ',anchor="w",style='primary.TLabel').grid(row=5, column = 0,sticky = "w",pady=2)
        ttk.Label(frame_customer, text='PLZ, Stadt:',style='primary.TLabel').grid(row=6, column = 0,sticky = "w",pady=2)
        ttk.Label(frame_customer, text='Email:',style='primary.TLabel').grid(row=7, column = 0,sticky = "w",pady=2)
        ttk.Label(frame_customer, text='Haustier:',font=("", 11, 'bold', 'underline'),style='primary.TLabel').grid(sticky = W, row = 8,pady=2)
        ttk.Label(frame_customer, text='Tierart:',style='primary.TLabel').grid(row=9, column = 0,sticky = "w",pady=2)
        ttk.Label(frame_customer, text='Tiername:',style='primary.TLabel').grid(row=10, column = 0,sticky = "w",pady=2)
        ttk.Label(frame_customer, text='Befund:',style='primary.TLabel').grid(row=11, column = 0,sticky = "w",pady=2)

        #creating text box for each input
        vornamen = [p.getVorname() for p in self.patientenliste.patientenliste]
        nachnamen = [p.getNachname() for p in self.patientenliste.patientenliste]

        self.nachname = ttk.Combobox(frame_customer,values=nachnamen)
        #self.nachname.bind('<<ComboboxSelected>>', lambda event: self.fillout_customer(event,self.nachname,nachnamen))
        self.nachname.grid(row=3, column=1,pady=2)
        self.autofill(self.nachname,nachnamen,frame_customer)

        self.vorname = ttk.Combobox(frame_customer,values=vornamen)
        self.vorname.grid(row=4, column=1, sticky = W,pady=2)
        self.autofill(self.vorname,vornamen,frame_customer)


        self.hausnummer = ttk.Entry(frame_customer)
        self.hausnummer.grid(row=5, column=1,pady=2)

        self.plz = ttk.Entry(frame_customer)
        self.plz.grid(row=6, column=1,pady=2)

        self.mail = ttk.Entry(frame_customer)
        self.mail.grid(row=7, column=1,pady=2)

        tiere = ["Hund", "Katze","Kaninchen","Meerschwein","Hamster","Chinchilla","Degu","Maus","Ratte","Pferd","Rind","Zuchtschwein","Mastschwein","Kalb","Ferkel","Schaf","Ziege","Fisch"]
        self.tierart = ttk.Combobox(frame_customer,values=tiere,height = 3,width = 18)
        self.tierart.grid(row=9, column=1,sticky = W,pady=2)

        self.tiername = ttk.Combobox(frame_customer,values=[],height = 3,width = 18)
        self.tiername.grid(row=10, column=1,pady=2)

        self.befund = ttk.Entry(frame_customer)
        self.befund.grid(row=11, column=1,pady=2,sticky =W)

        self.all_entry = [self.vorname, self.nachname, self.hausnummer, self.plz, self.mail, self.tierart,
                          self.tiername, self.befund]

    def create_leistungen_input(self,frame1,frame,frame2):
        #Leistungen
        ttk.Label(frame, text ='        ').grid()
        ttk.Label(frame, text='Leistungen:',font=("", 11, 'bold', 'underline'),style='primary.TLabel').grid(sticky = W, row = 0)

        #Leistungstag
        ttk.Label(frame, text ='        ').grid()
        self.werktag= ttk.Checkbutton(frame, text='Werktag (Mo-Fr)',width =16, var=self.tag, onvalue= 1, offvalue = 0).grid(row = 1, column = 0)
        self.woende = ttk.Checkbutton(frame, text='Wochenende (Sa,So)', width =20,var=self.tag, onvalue=2, offvalue = 0).grid(row = 1, column = 1)
        self.feiertag = ttk.Checkbutton(frame, text='Feiertag', var=self.tag,width =10, onvalue=3, offvalue = 0).grid(row = 1, column = 2)
        self.notdienst = ttk.Checkbutton(frame, text='Notdienst', var=self.tag,width =10, onvalue=4, offvalue = 0).grid(row = 1, column = 3,sticky = W)
        ttk.Label(frame, text ='        ').grid(row = 2)

        #Leistungstable
        ttk.Label(frame, text='Paragraph:', width = 13).grid(sticky = W,column = 1, row = 3)
        ttk.Label(frame, text='Tierart:', width = 10).grid(sticky = W,column = 4, row = 3, padx = 15)
        ttk.Label(frame, text='Stufe:', width = 5).grid(sticky = W,column = 5, row = 3, padx = 15)
        ttk.Label(frame, text='Nahrungs- \nergänzung:').grid(sticky = W,column = 2, row = 3,)
        ttk.Label(frame, text='Leistungen/\nMedikamente:',width = 35).grid(sticky = W,column = 3, row = 3)
        ttk.Label(frame,text ='Medikamente \n Netto in €:').grid(sticky = W, column = 6, row = 3, padx = 2)
        ttk.Label(frame,text ='Betrag\nNetto in €:').grid(sticky = W,column = 7, row = 3)
        ttk.Label(frame,text ='Datum der\nBehandlung').grid(sticky = W, column = 0, row = 3)

        ttk.Button(frame2,text="Leistung hinzufügen",style = 'secondary.TButton', command = lambda :self.add_row(frame), width = 22).grid(column = 3, row = self.row_leistugen+101,pady=2)
        ttk.Button(frame2,text="Erzeuge Rechnung",command = self.make_bill, style = 'success.TButton', width = 17).grid(padx = 2,column = 1, row = self.row_leistugen+104,pady=2)
        ttk.Button(frame2,text="Neu", style = 'secondary.TButton', command = lambda :self.restart(frame1,frame), width = 5).grid(padx = 2,column = 2, row =self.row_leistugen+104,pady=2)
        ttk.Button(frame2,text="Weiteres Tier hinzufügen", command = self.restartTier, style = 'secondary.TButton', width = 22).grid(padx = 2,column = 0, row =self.row_leistugen+104,pady=2)

    #add Leistungsrow
    def add_row(self,frame_leistungen):
        self.row_leistugen += 1

        tagN = BooleanVar()
        start_leistung_values = sorted(list(set(self.got['Leistung'].values)))

        nahrungsergänzungsmittel = ttk.Checkbutton(frame_leistungen,width =2, var=tagN, onvalue= True, offvalue = False,style = 'primary.TCheckbutton')
        tierart = ttk.Combobox(frame_leistungen,values=[], width = 18, height = 4)
        leistung = ttk.Combobox(frame_leistungen,values=start_leistung_values, width = 50, height = 4)
        leistung.bind('<KeyRelease>', lambda eventK: self.checkkey(eventK,leistung,start_leistung_values,combobox = True))
        leistung.bind('<<ComboboxSelected>>', lambda eventS: self.reset_leistung(event=eventS,leistung=leistung,stufe=stufe,betrag=betrag,paragraph=paragraph,tierart=tierart))
        tierart.bind('<<ComboboxSelected>>', lambda event:self.reset_tierart(tierart,stufe,betrag,leistung,paragraph,event,medi))

        try:
            tier = self.tierart.get()
        except:
            pass
        if tier in ["Kaninchen","Meerschwein","Hamster","Chinchilla","Degu","Maus","Ratte"]:
            tierart.insert(0,"Heimtiere")
        else:
            tierart.insert(0,tier)

        medi =ttk.Entry(frame_leistungen, width = 15)
        betrag = ttk.Entry(frame_leistungen, width = 15)
        date_label = DateEntry(frame_leistungen,date_pattern='dd/MM/yyyy')
        date_label.configure(validate='none')

        paragraph = self.parapgraph_scrollbar(frame_leistungen,leistung,betrag)
        stufe = self.stufe_spinbox(frame_leistungen,medi,betrag,leistung,paragraph,tierart)

        #b= ttk.Button(self.frame_leistungen, text="Datum", command= lambda: self.show_calender(), width = 10).grid(sticky = W, row = self.row_leistugen, column = 6)
        leistungsreihe = nahrungsergänzungsmittel.grid(row = self.row_leistugen, column = 2,pady=2),leistung.grid(column = 3, row = self.row_leistugen, sticky = W,padx = 2,pady=2),\
                         tierart.grid(column = 4, row = self.row_leistugen, sticky = W,padx=2,pady=2),medi.grid(column = 6, row = self.row_leistugen,sticky = W,padx = 2,pady=2), betrag.grid(padx=2,column = 7, row = self.row_leistugen,sticky = W,pady=2), \
                         date_label.grid(sticky = W,column = 0, row = self.row_leistugen,pady=2),

        date_label.delete(0, "end")
        self.leistungen.append([leistung,medi,betrag,tagN, date_label,self.tierart.get(),self.tiername.get(),paragraph])
        self.labels.append(leistung)
        self.labels.append(medi)
        self.labels.append(betrag)
        self.labels.append(nahrungsergänzungsmittel)
        self.labels.append(date_label)
        self.boxes.append(tierart)

    def parapgraph_scrollbar(self,canvas,leistung,betrag):
        values = self.got['Paragraph'].tolist()
        values = sorted(list(set(values)))
        pg = ttk.Combobox(canvas,values=values, width = 10,height=3)
        pg.bind('<<ComboboxSelected>>', lambda event: self.paragraph_update_leistung(event,leistung,betrag))
        pg.grid(column = 1, row =self.row_leistugen, sticky = 'W')#column = 1, row =self.row_leistugen,sticky='ns'
        self.boxes.append(pg)
        return pg

    def paragraph_update_leistung(self,event,leistung,betrag):
        leistung.delete(0,END)
        betrag.delete(0,END)
        selected = event.widget.get()
        found = self.got.loc[self.got['Paragraph']==selected]['Leistung']#.values
        values = found.values#.iloc[0])
        values = [v for v in found]
        values = list(set(values))
        leistung['values']= values#[values]

    def reset_leistung(self,event,stufe,betrag,leistung,paragraph,tierart):
        stufe.delete(0,END)
        betrag.delete(0,END)
        leist = leistung.get()
        tier = tierart.get()

        found = self.got.loc[self.got['Leistung']==leist]
        para = found['Paragraph']



        #isfound = found.loc[found['Tier']==tier]
        isfound = found[found['Tier'].str.contains(tier,na=False)]
        print(isfound)

        if isfound.empty == True:
            tierart.delete(0,END)
            found = found['Tier'].values
            values = ["Alle" if v is np.NAN else v for v in found]

            if len(values) > 1 :
                tierart['values']=values
            else:
                tierart.insert(0,values[0])

        else:
            tierart['values'] = isfound['Tier'].values()
            para = isfound['Paragraph'].iloc[0]
            paragraph.delete(0,END)
            paragraph.insert(0,para)


    def reset_tierart(self,tierart,stufe,betrag,leistung,paragraph,event,mediEntry):
        stufe.delete(0,END)
        paragraph.delete(0,END)
        betrag.delete(0,END)
        mediEntry.delete(0,END)

        leist = leistung.get()
        found = self.got.loc[self.got['Leistung']==leist]
        foundTier = found.loc[found['Tier']==event.widget.get()]

        if foundTier.empty == True:
            para = found['Paragraph'].values[0]
        else:
           para = foundTier['Paragraph'].values[0]

        paragraph.insert(0,para)

    def stufe_spinbox(self,canvas,mediEntry,betragEntry,leistung,paragraph,tierart):
        pg = ttk.Combobox(canvas,values=['1-fach','2-fach', '3-fach'], width = 8,height=4)
        pg.bind('<<ComboboxSelected>>', lambda event: self.stufe_update_preise(event,mediEntry,betragEntry,leistung,paragraph,tierart))
        pg.grid(column = 5, row =self.row_leistugen, padx = 2)
        self.boxes.append(pg)
        return pg

    def stufe_update_preise(self,event,mediEntry,betragEntry,leistung,paragraph,tierart):
        mediEntry.delete(0,END)
        betragEntry.delete(0,END)

        leist = leistung.get()
        selected = event.widget.get()
        para = paragraph.get()
        found = self.got.loc[self.got['Leistung']==leist]

        # check with paragraph:
        try:
            found = found.loc[found['Paragraph']==para]

            #if found.empty == True:
            #found = self.got.loc[self.got['Leistung']==leist]
            tier = tierart.get()
            found = found.loc[found['Tier']==tier]

            if len(found[selected].values) >1:
                tier = tierart.get()
                found = found.loc[found['Tier']==tier]
                found = found[selected].values[0]
            else:
                found = found[selected].values[0]
        except:
            found = self.got.loc[self.got['Leistung']==leist]
            para = found['Paragraph'].values[0]
            if len(found[selected].values) >1:
                tier = tierart.get()
                found = found.loc[found['Tier']==tier]
                found = found[selected].values[0]
            else:
                found = found[selected].values[0]
            paragraph.delete(0,END)
            paragraph.insert(0,para)

        betragEntry.insert(0,str(found).replace(".",","))

    def autofill(self,entry,data,frame):
        entry.bind('<KeyRelease>', lambda event: self.checkkey(event,entry,data,combobox = True))
        self.update_combobox(data,entry)
        entry.bind('<<ComboboxSelected>>', lambda event: self.fillout_customer(event,entry))


    def checkkey(self,event, listbox,elem, row=None, col=None,combobox = False):
        #input char
        value = event.widget.get()

        # get data from l
        if value == '':
            data = elem
        else:
            data = []
            for item in elem:
                if value.lower() in item.lower():
                    data.append(item)
                    # update data in listbox
        if combobox == False:
            self.update(data,listbox)
        else:
            self.update_combobox(data,listbox)
        #show
        try:
            listbox.grid(row = row,column = col)
        except:
            pass

    def update_combobox(self,data,lb):
        # put new data
        lb['values'] = data

    def fillout_customer(self,event,box):
        print("Filling out:")
        #box.delete(0,END)
        #box.insert(0, event.widget.get(ACTIVE))
        nachnameBox = False
        if box == self.nachname:
            nachnameBox = True
        foundPatienten = self.patientenliste.fillout_patient(self.patientenliste.patientenliste, event.widget.get(),box,nachnameBox)
        self.fill_in_patient(foundPatienten)

    def fill_on_patienten(self,patienten):
        print(patienten)

    def reset(self, listbox):
        listbox.grid_forget()
#--------------------------------------------------------------------------------------------------------------------#
    #create word document
    def make_bill(self):
        tierart =""
        tiername =""

        try:
            #get values of leistungen, l=leistung, m=medikament,b=behandlung,n=nahrungsergänzung, date= Behandlungsdatum
            for leistung in self.leistungen:
                l = leistung[0].get()
                if leistung[1].get():
                    m = float(leistung[1].get().replace(",", "."))
                else:
                    m = 0
                if leistung[2].get():
                    b = float(leistung[2].get().replace(",", "."))
                else:
                    b = 0
                n = leistung[3].get()

                date= leistung[4].get()
                tierart =leistung[5]
                tiername =leistung[6]
                paragraph = leistung[7]
                if [l,m,b,n,date,leistung[5],leistung[6],self.tag.get(),paragraph.get()] not in self.inputWerte_leistungen:
                    self.inputWerte_leistungen.append([l,m,b,n,date,leistung[5],leistung[6],self.tag.get(),paragraph.get()])
            if self.befund.get() not in self.befundsliste:
                self.befundsliste.append(self.befund.get())
        #add daily fee
        except:
            tkinter.messagebox.showwarning(title="Error",message="Etwas ist bei der Erstellung Ihrer Rechnung schief gegangen.\nBitte überprüfen Sie Ihre Eingaben!")
        #print(self.inputWerte_leistungen)
        #create word
        document, name, timestamp,rnr = Template.make_word(self.anrede_spin_box.get(), self.nachname.get(),self.vorname.get(),
                                                           self.hausnummer.get(),self.plz.get(),self.mail.get(),
                                                           self.befundsliste, self.inputWerte_leistungen,self.bill_counter, self.several_counter,self.tag.get())
        #save kontaktdaten
        self.save_contacts()

        #save bill_counter
        self.save_jsonCounter()

        #save word document
        self.save_doc(document,name, timestamp,rnr)

    def save_doc(self,saveDocument,name, timestamp,rnr):
        is_present = False
        files = [('Word', '*.docx')]

        #let user choose dir
        dir = tkinter.filedialog.asksaveasfile(filetypes = files, defaultextension = files)

        document = Document()
        try:
            document = Document(dir.name)
            is_present =True
        except:
            pass

        if is_present:
            os.remove(dir.name) # delete old

        saveDocument.save(dir.name)
        tkinter.messagebox.showinfo(title="Rechnung gespeichert", message="Ihre Rechnung mit der RechnungsNr: "+str(rnr)+"\n für "+self.anrede_spin_box.get()+" "+self.vorname.get()+" "+ self.nachname.get()+" wurde erzeugt!")

    def save_contacts(self):
        person = Patient.Person(self.anrede_spin_box.get(),self.vorname.get(), self.nachname.get(),
                                self.hausnummer.get(),self.plz.get(),self.mail.get(),
                                self.tierart.get(),self.tiername.get())

        if not self.patientenliste.existPatient(person):
            self.patientenliste.addPatient(person)
        self.patientenliste.savePatients()


    def add_daily_fee(self,tierart,tiername):
        m = 0
        n = False
        date = None
        l=""
        b=""
        if self.tag.get() == 4:
            l = "Notdienst"
            b = 50.0
            tagParagraph = "3a"
        elif self.tag.get() == 3:
            l = "Feiertag"
            b = 10.0
            tagParagraph = "3a"
        elif self.tag.get() == 2:
            l = "Wochenende"
            b = 15.0
            tagParagraph = "3a"
        self.inputWerte_leistungen.append([l,m,b,n,date,tierart,tiername])

    def restart(self,frame1,frame2):
        self.r.destroy()
        root = ThemedTk(theme="arc")
        gui = GUI(root,df_got)
        gui.start()
        #subprocess.Popen(["GUI_Tkinter_2.0.py"],shell = True)
        #python = sys.executable
        #os.execl(python, python, * sys.argv)



    def restartTier(self):
        if not self.tiername.get() or not self.tierart.get():
            tkinter.messagebox.showinfo(title="Tierinformationen fehlen", message="Bitte geben Sie zunächst alle Patienten- und Tierinformationen ein!")
            return
        #save inputs of pet for new pet inputs
        for leistung in self.leistungen:
            l = leistung[0].get()
            if leistung[1].get():
                m = float(leistung[1].get().replace(",", "."))
            else:
                m = 0
            if leistung[2].get():
                b = float(leistung[2].get().replace(",", "."))
            else:
                b = 0
            n = leistung[3].get()
            date= leistung[4].get()
            tierart = leistung[5]
            tiername = leistung[6]
            paragraph = leistung[7]
            if [l,m,b,n,date,tierart,tiername,self.tag.get(),paragraph.get()] not in self.inputWerte_leistungen:
                self.inputWerte_leistungen.append([l,m,b,n,date,tierart,tiername,self.tag.get(),paragraph.get()])

        if self.befund.get() not in self.befundsliste:
            self.befundsliste.append(self.befund.get())

        #delete Entrys in GUI
        self.tiername.delete(0,END)
        self.tierart.delete(0,END)
        self.befund.delete(0,END)
        for l in self.labels:
            l.destroy()
        for box in self.boxes:
            box.grid_forget()
        self.leistungen =[]
        self.several_counter +=1

        tkinter.messagebox.showinfo(title="Eingaben gespeichert", message="Ihre Eingaben zum Patiententier wurden gespeichert.\nBitte geben Sie alle Daten zum nächsten Tier ein.")

    def save_jsonCounter(self):
        self.bill_counter += 1

        out_file = open(self.pathc, "w")
        data =[self.bill_counter, self.date_month]
        json.dump(data, out_file, indent = 3)
        out_file.close()

    def load_jsonContacts(self):
        #C:\Users\Pepi\AppData\Local\Rechnung_template\
        with open(self.pathk, 'r') as fp:
            data = json.load(fp)
        for dict in data:
            p = Patient.Person(dict.get("Anrede"),dict.get("Vorname"),dict.get("Nachname"),dict.get("Straße"),dict.get("Addresse"),
                               dict.get("Mail"),dict.get("Tierart"),dict.get("Tiername"))
            self.patientenliste.patientenliste.append(p)
        self.autocompleteName= [[p.getVorname(),p.getNachname()] for p in self.patientenliste.patientenliste]

    def load_jsonCounter(self):
        date = datetime.now()
        actual_month = date.strftime("%m")
        #C:\Users\Pepi\AppData\Local\Rechnung_template\
        with open(self.pathc, 'r') as fp:
            data = json.load(fp)
        #if new month reset bill_counter
        if actual_month == data[1]:
            self.bill_counter = data[0]
            self.date_month = data[1]
        else:
            self.bill_counter = 1
            self.date_month = actual_month

    ###upload all existing contacts in rechnungs_template
    def load_allFiles(self):
        try:
            self.load_jsonContacts()
            self.load_jsonCounter()
        except:
            #C:\Users\Pepi\AppData\Local\Rechnung_template\
            out_file = open(self.pathk, "w")
            json.dump(self.patientenliste.patientenliste, out_file, indent = 9)
            out_file.close()
            #C:\Users\Pepi\AppData\Local\Rechnung_template\
            out_file2 = open(self.pathc, "w")
            json.dump([self.bill_counter,self.date_month], out_file2, indent = 9)
            out_file2.close()
            self.load_jsonContacts()
            self.load_jsonCounter()






if __name__ == "__main__":
    pdf = r'D:\Rechnung_template_projekt\GOT_2020.pdf'
    df_got = GOT_Parser.exampleDF()
    root = Tk() #ThemedTk(theme="arc")#
    gui = GUI(root,df_got)
    gui.start()