from decimal import Decimal
from tkinter.messagebox import *

import gsheets


class GSheetMTT(gsheets.GSheet):
    def __init__(self, _stdBetrag, stdZweck):
        super().__init__("", stdZweck)
        self.spreadSheetId = "1LNopZRbkggBm4OtRRp-YtRBFKKv1z8no0xLEhk-K8mo"  # Test-Backend-MTT

        # Buchungen
        # diese Felder brauchen wir für den Einzug
        self.ebicsnames = ebicsnames = ["Name des Kontoinhabers", "IBAN-Kontonummer",
                                        "Betrag", "Zweck", "Zeitstempel"]
        self.ktoinh = ebicsnames[0]
        self.iban = ebicsnames[1]
        self.betrag = ebicsnames[2]
        self.zweck = ebicsnames[3]
        self.datum = ebicsnames[4]

        # Felder die wir überprüfen
        self.formnames = formnames = ["Anrede", "Vorname", "Name", "ADFC-Mitgliedsnummer",
                                      "Bei welchen Touren möchten Sie mitfahren?",
                                      "Reisen Sie alleine oder zu zweit?",
                                      "Zustimmung zur SEPA-Lastschrift", "Bestätigung",
                                      "Verifikation", "Anmeldebestätigung"]
        self.anrede = formnames[0]
        self.vorname = formnames[1]
        self.name = formnames[2]
        self.mitglied = formnames[3]
        self.reisenB = formnames[4]
        self.einzeln = formnames[5]
        self.zustimmung = formnames[6]
        self.bestätigung = formnames[7]  # Bestätigung der Teilnahmebedingungen
        self.verifikation = formnames[8]
        self.anmeldebest = formnames[9]  # wird vom Skript Radfahrschule/Anmeldebestätigung senden ausgefüllt

        # diese Felder fügen wir hinzu
        self.zusatzFelder = zusatzFelder = ["Eingezogen", "Zahlungseingang",
                                            "Kommentar", "Zahlungsbetrag"]
        self.eingezogen = zusatzFelder[0]
        self.zahlungseingang = zusatzFelder[1]  # händisch
        self.kommentar = zusatzFelder[2]
        self.zahlungsbetrag = zusatzFelder[3]


        # Reisen
        self.reisenNames = reisennames = ["Reise", "DZ-Preis", "EZ-Preis"]
        self.reiseR = reisennames[0]
        self.dzPreis = reisennames[1]
        self.ezPreis = reisennames[2]

        self.reisenIndex = None  # column no of "Reise" in Reisen sheet
        self.ezPreisIndex = None
        self.dzPreisIndex = None

    @classmethod
    def getDefaults(cls):
        return "je nach Reise", "ADFC Mehrtagestouren", "ADFC-M-MTT-2023"

    def validSheetName(self, sname):
        return sname == "Buchungen" or sname == "Reisen"

    def checkKtoinh(self, row):
        inh = row[self.ktoinh]
        if len(inh) < 5 or inh.startswith("dto") or inh.startswith("ditto"):
            row[self.ktoinh] = row[self.vorname] + " " + row[self.name]
        return True

    def getReisenIndices(self):
        reisenRows = self.data["Reisen"]
        headers = reisenRows[0]
        self.reisenIndices = {}
        for i, h in enumerate(headers):
            self.reisenIndices[h] = i
        self.reisenIndex = self.reisenIndices[self.reiseR]
        self.ezPreisIndex = self.reisenIndices[self.ezPreis]
        self.dzPreisIndex = self.reisenIndices[self.dzPreis]

    def reisenBetrag(self, row):
        if self.reisenIndex is None:
            self.getReisenIndices()

        reisenName = row[self.reisenB]
        reisenRows = self.data["Reisen"]
        # if rrow has a note, then buchungen should also have a note,
        # i.e. here we should see no matching rrow with a note
        reisenRows = list(filter(lambda rrow: rrow[self.reisenIndex] == reisenName, reisenRows))
        if len(reisenRows) != 1:
            msg = "Reise '" + reisenName + "' nicht in Reisen"
            showerror("Fehler", msg)
            raise ValueError(msg)
        reisenRow = reisenRows[0]
        einzeln = row[self.einzeln].startswith("Alleine")
        zimmerpreis = reisenRow[self.ezPreisIndex if einzeln else self.dzPreisIndex]
        return zimmerpreis if einzeln else (zimmerpreis * 2)

    def checkBetrag(self, row):
        mitglied = row[self.mitglied] != ""
        if not mitglied:
            anrede = row[self.anrede] + " " + row[self.name]
            showerror("Fehlbuchung?", anrede + " ist nicht Mitglied?")
            return False
        if self.betrag not in row:
            row[self.betrag] = self.reisenBetrag(row)
        row[self.betrag] = Decimal(row[self.betrag].replace(',', '.'))  # 3,14 -> 3.14
        return True
