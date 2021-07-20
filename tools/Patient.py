import json
from tkinter import END


class Patientenliste:
    def __init__(self):
        self.patientenliste =[]
        #self.pathk = r'C:\Users\Pepi\AppData\Local\Rechnung_template\kontaktdaten.json'
        self.pathk= r'kontaktdaten.json'

    def addPatient(self, patient):
        self.patientenliste.append(patient)

    def searchPatientByNachname(self,inputPatientenListe,nachname):
        matches =[]
        for p in inputPatientenListe:
            nachn= p.getNachname()
            if nachname == nachn:
                matches.append(p)
        return matches

    def searchPatientByVorname(self,inputPatientenListe,vorname):
        matches =[]
        for p in inputPatientenListe:
            vorn= p.getVorname()
            if vorname == vorn:
                matches.append(p)
        return matches

    def savePatients(self):
        contacts=[]
        tiernamen = []
        for p in self.patientenliste:
            tiernamen =[]
            tiernamen.append(p.getTiername())
            data = {"Anrede":p.getAnrede(),
                         "Vorname":p.getVorname(),
                         "Nachname":p.getNachname(),
                         "Straße":p.getStraße(),
                         "Addresse": p.getPLZ(),
                         "Mail": p.getMail(),
                         "Tierart": p.getTierart(),
                         "Tiername":tiernamen}
            contacts.append(data)
        #C:\Users\Pepi\AppData\Local\Rechnung_template\
        out_file = open(self.pathk, "w")
        json.dump(contacts, out_file, indent = 9)
        out_file.close()

    def existPatient(self,person):
        exist = False
        try:
            personAnrede = person.getAnrede()
            personVorname = person.getVorname()
            personNachname = person.getNachname()
            #personTier= person.tiername()
        except:
            exist = True
            return exist
        for p in self.patientenliste:
            anrede = p.getAnrede()
            vorname = p.getVorname()
            nachname = p.getNachname()
            tiername = p.getTiername()

            if (personAnrede==anrede)and (personVorname==vorname) and (personNachname==nachname):
                exist = True

        return exist

    # def check_for_patient(self, entry,vornameEntry,nachnameEntry, entryInput, anredeEntry, mailEntry,plzEntry,hausnrEntry):
    #     if entry == vornameEntry:
    #         patienten, found = self.searchPatient(vorname = entryInput, nachname = self.nachname.get())
    #         if found:
    #             anredeEntry.delete(0, END)
    #             anredeEntry.insert('end',patienten[0].getAnrede())
    #
    #             nachnameEntry.delete(0, END)
    #             nachnameEntry.insert('end',patienten[0].getVorname())
    #
    #             mailEntry.delete(0, END)
    #             mailEntry.insert('end',patienten[0].getMail())
    #
    #             plzEntry.delete(0, END)
    #             plzEntry.insert('end',patienten[0].getPLZ())
    #
    #             hausnrEntry.delete(0, END)
    #             hausnrEntry.insert('end',patienten[0].getStraße())
    #         patienten=[]
    #         found = False
    #
    #     if entry == nachnameEntry:
    #         patienten, found=self.searchPatient(vorname = vornameEntry.get(),nachname = entryInput)
    #         if found and anredeEntry.get() == patienten[0].getAnrede():
    #             anredeEntry.delete(0, END)
    #             anredeEntry.insert('end',patienten[0].getAnrede())
    #             vornameEntry.delete(0, END)
    #             vornameEntry.insert('end',patienten[0].getVorname())
    #             mailEntry.delete(0, END)
    #             mailEntry.insert('end',patienten[0].getMail())
    #             plzEntry.delete(0, END)
    #             plzEntry.insert('end',patienten[0].getPLZ())
    #             hausnrEntry.delete(0, END)
    #             hausnrEntry.insert('end',patienten[0].getStraße())
    #         patienten=[]
    #         found = False

    def fillout_patient(self,patientenliste,selectedName,selectedBox,nachnameBox):
        foundPatienten = None
        if nachnameBox:
            nachname = selectedName
            foundPatienten = self.searchPatientByNachname(patientenliste,nachname)

        else:
            vorname = selectedName
            foundPatienten = self.searchPatientByVorname(patientenliste,vorname)
        return foundPatienten

class Person:
    def __init__(self,anrede,vorn,nachn,straße_haus, plz_ort, mail,tierart,tiername):
        self.id = None
        self.anrede = anrede
        self.vorname = vorn
        self.nachname = nachn
        self.straße_hausnr = straße_haus
        self.addresse = plz_ort
        self.email = mail
        self.tierart= tierart
        self.tiername = tiername

    def getId(self):
        return self.id

    def getAnrede(self):
        return self.anrede

    def getVorname(self):
        return self.vorname

    def getNachname(self):
        return self.nachname

    def getStraße(self):
        return self.straße_hausnr

    def getPLZ(self):
        return self.addresse

    def getMail(self):
        return self.email

    def getTierart(self):
        return self.tierart

    def getTiername(self):
        return self.tiername

