import locale
import os
import os.path
import logging
import getpass
from Cryptodome.Cipher import AES
from Cryptodome.Hash import SHA256
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import *

import ebics
from gsheets import pyinst
# from paypal.paypalhttp.http_client import HttpClient

# not allowed to allow serviceaccount access to spreadsheet!?
# import gspread
# from oauth2client.service_account import ServiceAccountCredentials


class RadioBoxes(Frame):
    def __init__(self, master, gui, labeltext, klasses):
        super().__init__(master)
        self.master = master
        self.gui = gui
        self.names = list(klasses.keys())
        self.boxes = []
        self.value = IntVar()
        self.klasses = klasses
        self.defaults = []
        self.label = Label(self, text=labeltext, bd=4, width=15, height=0, relief=RIDGE)
        self.label.grid(row=0, column=0, sticky="w")
        for i, s in enumerate(self.names):
            box = Radiobutton(self, text=s, value=i+1, variable=self.value, command=self.setVal)
            box.grid(row=0, column=i + 1, sticky="w")
            self.boxes.append(box)
            klass = klasses.get(s)
            self.defaults.append(klass.getDefaults())

    def setVal(self):
        i = self.value.get()-1
        b, z, m = self.defaults[i]
        self.gui.mandat = m
        self.gui.kursArgLE.set("")
        self.gui.betragLE.set(b)
        self.gui.zweckLE.set(z)
        self.gui.mandatLE.set(m)
        self.gui.kursArgLE.config(state=NORMAL)
        self.gui.betragLE.config(state=NORMAL)
        self.gui.zweckLE.config(state=NORMAL)
        self.gui.mandatLE.config(state=NORMAL)

    def get(self):
        i = self.value.get() -1
        return {self.names[i]: i}


class ButtonEntry(Frame):
    def __init__(self, master, buttontext, stringtext, w, cmd):
        super().__init__(master)
        self.btn = Button(self, text=buttontext, bg="red", bd=4, width=15, height=0, relief=RAISED, command=cmd)
        self.svar = StringVar()
        self.svar.set(stringtext)
        self.entry = Entry(self, textvariable=self.svar, bd=4,
                           width=w, borderwidth=2)
        self.grid_rowconfigure(0, weight=0)
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.btn.grid(row=0, column=0, sticky="w")
        self.entry.grid(row=0, column=1, sticky="we")

    def get(self):
        return self.svar.get()

    def set(self, s):
        return self.svar.set(s)


class LabelEntry(Frame):
    def __init__(self, master, labeltext, stringtext, w, command=None):
        super().__init__(master)
        self.label = Label(self, text=labeltext, bd=4, width=15, height=0, relief=RIDGE)
        self.svar = StringVar()
        self.svar.set(stringtext)
        if command is None:
            self.entry = Entry(self, textvariable=self.svar, bd=4,
                               width=w, borderwidth=2)
        else:
            cmd = self.register(command)
            self.entry = Entry(self, textvariable=self.svar, bd=4,
                               width=w, borderwidth=2,
                               validatecommand=(cmd,"%P"), validate="key")
        self.grid_rowconfigure(0, weight=0)
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.label.grid(row=0, column=0, sticky="w")
        self.entry.grid(row=0, column=1, sticky="we")

    def get(self):
        return self.svar.get()

    def config(self, **kwargs):
        self.entry.config(**kwargs)

    def set(self, s):
        return self.svar.set(s)


class LabelOM(Frame):
    def __init__(self, master, labeltext, options, initVal, **kwargs):
        super().__init__(master)
        self.options = options
        self.label = Label(self, text=labeltext, bd=4, width=15, relief=RIDGE)
        self.svar = StringVar()
        self.svar.set(initVal)
        self.optionMenu = OptionMenu(self, self.svar, *options, **kwargs)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.label.grid(row=0, column=0, sticky="w")
        self.optionMenu.grid(row=0, column=1, sticky="w")

    def get(self):
        return self.svar.get()

    def set(self, s):
        self.svar.set(s)


class MyApp(Frame):
    def setMandat(self, txt):
        self.mandatLE.set(self.mandat + "-" + txt.replace("_", "-"))
        return True

    def __init__(self, master):
        super().__init__(master)
        w = 50
        self.sheetsBE = RadioBoxes(master, self, "Sheets", ebics.getKlasses())
        # self.templateBE = ButtonEntry(master, "Template-Datei", ebics.templateFileDefault, w, self.templFileSetter)
        self.outputLE = LabelEntry(master, "Ausgabedatei", "ebics.xml", w)
        self.kursArgLE = LabelEntry(master, "Kursname", "", w, command=self.setMandat)
        self.kursArgLE.config(state=DISABLED)
        self.betragLE = LabelEntry(master, "Betrag", "", w)
        self.betragLE.config(state=DISABLED)
        self.zweckLE = LabelEntry(master, "Zweck", "", w)
        self.zweckLE.config(state=DISABLED)
        self.mandatLE = LabelEntry(master, "Mandat", "", w)
        self.mandatLE.config(state=DISABLED)
        self.checkBtn = Button(master, text="Prüfen", bd=4, bg="red", width=15, command=self.check)
        self.testBtn = Button(master, text="Test", bd=4, bg="red", width=15, command=self.testen)
        self.startBtn = Button(master, text="EBICS", bd=4, bg="red", width=15, command=self.starten)

        for x in range(1):
            Grid.columnconfigure(master, x, weight=1)
        for y in range(9):
            Grid.rowconfigure(master, y, weight=1)

        self.sheetsBE.grid(row=0, column=0, sticky="we")
        # self.templateBE.grid(row=1, column=0, sticky="we")
        self.outputLE.grid(row=2, column=0, sticky="we")
        self.kursArgLE.grid(row=3, column=0, sticky="we")
        self.betragLE.grid(row=4, column=0, sticky="we")
        self.zweckLE.grid(row=5, column=0, sticky="we")
        self.mandatLE.grid(row=6, column=0, sticky="we")
        self.checkBtn.grid(row=7, column=0, sticky="w")
        self.testBtn.grid(row=8, column=0, sticky="w")
        self.startBtn.grid(row=9, column=0, sticky="w")

    # def templFileSetter(self):
    #     x = askopenfilename(title="Template Datei auswählen", defaultextension=".xml", filetypes=[("XML", ".xml")])
    #     self.templateBE.set(x)

    def check(self):
        sels = self.sheetsBE.get()
        sammel = sels.get("SAMMEL") is not None

        eb = ebics.Ebics(
                   sels,
                   self.outputLE.get(),
                   self.kursArgLE.get(),
                   self.betragLE.get(),
                   self.zweckLE.get(),
                   self.mandatLE.get(),
                   # self.templateBE.get()
            )
        try:
            eb.check()
            stats = eb.getStatistics()
            if sammel:
                msg = f"Anzahl Überweisungsaufträge: {stats[0]}\nÜberwiesen: {stats[3]}\nNoch zu überweisen: {stats[4]}"
            else:
                msg = f"Anzahl Abbuchungsaufträge: {stats[0]}\nUnverifizierte Email-Adresse: {stats[1]}\nSchon bezahlt: {stats[2]}\nAbgebucht: {stats[3]}\nNoch abzubuchen: {stats[4]}"
            showinfo("Ergebnis", msg)
        except Exception as e:
            logging.exception("Fehler")
            showerror("Fehler", str(e))

    def testen(self):
        self.writeEbics(False)

    def starten(self):
        self.writeEbics(True)

    def writeEbics(self, setEingezogen):
        if self.outputLE.get() == "":
            showerror("Fehler", "keine Ausgabedatei")
            return
        sels = self.sheetsBE.get()
        sammel = sels.get("SAMMEL") is not None
        eb = ebics.Ebics(
                   self.sheetsBE.get(),
                   self.outputLE.get(),
                   self.kursArgLE.get(),
                   self.betragLE.get(),
                   self.zweckLE.get(),
                   self.mandatLE.get(),
                   # self.templateBE.get()
        )
        try:
            res = eb.createEbicsXml(setEingezogen)
            stats = eb.getStatistics()
            if sammel:
                msg = f"Anzahl Überweisungsaufträge: {stats[0]}\nÜberwiesen: {stats[3]}\nNoch zu überweisen: {stats[4]}"
            else:
                msg = f"Anzahl Abbuchungsaufträge: {stats[0]}\nUnverifizierte Email-Adresse: {stats[1]}\nSchon bezahlt: {stats[2]}\nAbgebucht: {stats[3]}\nNoch abzubuchen: {stats[4]}"
            if res is None:
                showinfo("Nichts gefunden", msg)
            else:
                showinfo("Erfolg", msg + f"\nAusgabe in Datei {self.outputLE.get()} erzeugt")
        except Exception as e:
            logging.exception("Fehler")
            showerror("Fehler", str(e))


def cred_decrypt():
    if os.path.exists("credentials.json"):
        return True
    if os.path.exists("password.txt"):
        with open("password.txt", "r") as p:
            pwd = p.read().strip()
    else:
        pwd = getpass.getpass("Passwort: ")
    sha = SHA256.new()
    sha.update(pwd.encode('utf-8'))
    key = sha.digest()
    chiffre = AES.new(key, AES.MODE_CFB, "1111111122222222".encode('utf-8'))
    with open(pyinst("credentials.json.encr"), "rb") as inp:
        b = inp.read()
        credb = chiffre.decrypt(b)
        if credb.find(b"client_id") == -1:
            print("Konnte credentials.json.encr nicht dekodieren. Falsches Passwort?")
            return False
        with open("credentials.json", "wb") as out:
            out.write(credb)
        with open("password.txt", "w") as p:
            p.write(pwd)
    return True


def cred_encrypt():
    if not os.path.exists("credentials.json.encr"):
        pwd = getpass.getpass("Passwort: ")
        with open("password.txt", "w") as p:
            p.write(pwd)
        sha = SHA256.new()
        sha.update(pwd.encode('utf-8'))
        key = sha.digest()
        chiffre = AES.new(key, AES.MODE_CFB, "1111111122222222".encode('utf-8'))
        with open("credentials.json.orig", "rb") as inp, open("credentials.json.encr", "wb") as out:
            b = inp.read()
            out.write(chiffre.encrypt(b))


if __name__ == '__main__':
    import sys

    if len(sys.argv) == 2 and sys.argv[1] == "e":
        cred_encrypt()
        sys.exit(0)
    locale.setlocale(locale.LC_TIME, "German")
    if not cred_decrypt():
        sys.exit(1)
    root = Tk()
    app = MyApp(root)
    app.master.title("ADFC Lastschrift")
    app.mainloop()
