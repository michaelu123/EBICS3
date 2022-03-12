import gsheets
from decimal import Decimal


class GSheetTK(gsheets.GSheet):
    def __init__(self, _stdBetrag, stdZweck):
        super().__init__("", stdZweck)
        self.spreadSheetId = "1r4WEgWskyJrHNRgWLOZ2n7pgho-8S6kSo7cm1po1pt4"  # Backend-Technikkurse
        self.spreadSheetName = "Backend-Technikkurse"

        # diese Felder brauchen wir für den Einzug
        self.ebicsnames = ebicsnames = ["Name des Kontoinhabers", "IBAN-Kontonummer", "Betrag", "Zweck", "Zeitstempel"]
        self.ktoinh = ebicsnames[0]
        self.iban = ebicsnames[1]
        self.betrag = ebicsnames[2]
        self.zweck = ebicsnames[3]
        self.datum = ebicsnames[4]

        # Felder die wir überprüfen
        self.formnames = formnames = ["Vorname", "Name", "ADFC-Mitgliedsnummer", "Zustimmung zur SEPA-Lastschrift",
                                      "Bestätigung", "Verifikation", "Anmeldebestätigung"]
        self.vorname = formnames[0]
        self.name = formnames[1]
        self.mitglied = formnames[2]
        self.zustimmung = formnames[3]
        self.bestätigung = formnames[4]  # Bestätigung der Teilnahmebedingungen
        self.verifikation = formnames[5]
        self.anmeldebest = formnames[6]  # wird vom Skript Radfahrschule/Anmeldebestätigung senden ausgefüllt

        # diese Felder fügen wir hinzu
        self.zusatzFelder = zusatzFelder = ["Eingezogen", "Zahlungseingang",
                                            "Kommentar"]
        self.eingezogen = zusatzFelder[0]
        self.zahlungseingang = zusatzFelder[1]  # händisch
        self.kommentar = zusatzFelder[2]

    @classmethod
    def getDefaults(cls):
        return "10/15", "ADFC Technikkurse", "ADFC-M-TK-2022"

    def validSheetName(self, sname):
        return sname.startswith("Buchungen")

    def checkKtoinh(self, row):
        inh = row[self.ktoinh]
        if len(inh) < 5 or inh.startswith("dto") or inh.startswith("ditto"):
            row[self.ktoinh] = row[self.vorname] + " " + row[self.name]
        return True

    def checkBetrag(self, row):
        mitglied = row[self.mitglied] != ""
        if self.betrag not in row:
            row[self.betrag] = "10" if mitglied else "15"
        row[self.betrag] = Decimal(row[self.betrag].replace(',', '.'))  # 3,14 -> 3.14
        return True
