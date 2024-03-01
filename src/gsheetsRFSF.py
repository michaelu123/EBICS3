import gsheets
from decimal import Decimal
from tkinter.messagebox import *


class GSheetRFSF(gsheets.GSheet):
    def __init__(self, kursArg, stdBetrag, stdZweck):
        super().__init__(kursArg, stdBetrag, stdZweck)
        self.spreadSheetId = "1K5wHJrEP0vP-tuM7gqSOUdEfq7OzFt5iCjnnjyzWzxA"
        self.kursFrage = "Welchen Kurs möchten Sie belegen?"
        # diese Felder brauchen wir für den Einzug
        self.ebicsnames = ebicsnames = ["Lastschrift: Name des Kontoinhabers", "Lastschrift: IBAN-Kontonummer",
                                        "Betrag", "Zweck", "Zeitstempel"]
        self.ktoinh = ebicsnames[0]
        self.iban = ebicsnames[1]
        self.betrag = ebicsnames[2]
        self.zweck = ebicsnames[3]
        self.datum = ebicsnames[4]

        # Felder die wir überprüfen
        self.formnames = formnames = ["Vorname", "Name", "ADFC-Mitgliedsnummer falls Mitglied",
                                      "Zustimmung zur SEPA-Lastschrift", "Bestätigung",
                                      "Verifikation", "Anmeldebestätigung", "Welchen Kurs möchten Sie belegen?"]
        self.vorname = formnames[0]
        self.name = formnames[1]
        self.mitglied = formnames[2]
        self.zustimmung = formnames[3]
        self.bestätigung = formnames[4]  # Bestätigung der Teilnahmebedingungen
        self.verifikation = formnames[5]
        self.anmeldebest = formnames[6]  # wird vom Skript Radfahrschule/Anmeldebestätigung senden ausgefüllt
        self.kursName = formnames[7]

        # diese Felder fügen wir hinzu
        self.zusatzFelder = zusatzFelder = ["Eingezogen", "Zahlungseingang", "Kommentar", "Zahlungsbetrag"]
        self.eingezogen = zusatzFelder[0]
        self.zahlungseingang = zusatzFelder[1]  # händisch
        self.kommentar = zusatzFelder[2]
        self.zahlungsbetrag = zusatzFelder[3]

    @classmethod
    def getDefaults(cls):
        return "20/35 oder 30/45", "ADFC Radfahrschule", "ADFC-M-RFSF-2024"

    def validSheetName(self, sname):
        return sname == "Buchungen"

    def checkKtoinh(self, row):
        inh = row[self.ktoinh]
        if len(inh) < 5 or inh.startswith("dto") or inh.startswith("ditto"):
            row[self.ktoinh] = row[self.vorname] + " " + row[self.name]
        return True

    def checkBetrag(self, row):
        mitglied = row[self.mitglied] != ""
        kurs = row[self.kursFrage]
        if self.betrag not in row:
            if kurs.endswith('G'):
                row[self.betrag] = "20" if mitglied else "35"
            elif kurs.endswith('A'):
                row[self.betrag] = "30" if mitglied else "45"
            elif kurs.endswith('S'):
                row[self.betrag] = "20" if mitglied else "35"
            else:
                showerror("Kursname", "Unbekannt:" + kurs)
        row[self.betrag] = Decimal(row[self.betrag].replace(',', '.'))  # 3,14 -> 3.14
        return True
