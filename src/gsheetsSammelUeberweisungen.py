import datetime

import gsheets
from decimal import Decimal


class GSheetSammelueberweisungen(gsheets.GSheet):
    def __init__(self, kursArg, _stdBetrag, stdZweck):
        super().__init__(kursArg, "", stdZweck)
        self.spreadSheetId = "1j7dmwEqro7beNGztODHuiEKyKQoW7I1jtLPdVqoCPs4"  # Sammelueberweisungen

        # diese Felder brauchen wir für den Einzug
        self.ebicsnames = ebicsnames = ["Name des Kontoinhabers", "IBAN-Kontonummer", "Betrag", "Zweck", "Betrag"]
        self.ktoinh = ebicsnames[0]
        self.iban = ebicsnames[1]
        self.betrag = ebicsnames[2]
        self.zweck = ebicsnames[3]
        self.zahlungsbetrag = ebicsnames[4]
        self.mandat = "Mandat" # wird für Überweisungen nicht gebraucht (anscheinend)
        self.datum = "useToday"

        # Felder die wir überprüfen
        # self.formnames = formnames = ["Vorname", "Name"]
        # self.vorname = formnames[0]
        # self.name = formnames[1]

        # diese Felder fügen wir hinzu
        self.zusatzFelder = zusatzFelder = ["Überwiesen", "Kommentar"]
        self.eingezogen = zusatzFelder[0]
        self.kommentar = zusatzFelder[1]

    @classmethod
    def getDefaults(cls):
        return "", "Sammelüberweisung", ""

    def validSheetName(self, sname):
        return sname.startswith("Buchungen")

    def checkKtoinh(self, row):
        inh = row.get(self.ktoinh)
        if not inh or len(inh) < 5:
            print("Kein gültiger Kontoinhaber in ", row["Sheet"])
            return False
        return True

    def checkBetrag(self, row):
        if not row.get(self.betrag):
            print("Kein Betrag in ", row["Sheet"], "KontoInhaber:", row[self.ktoinh])
            return False
        row[self.betrag] = Decimal(row[self.betrag].replace(',', '.'))  # 3,14 -> 3.14
        return True
