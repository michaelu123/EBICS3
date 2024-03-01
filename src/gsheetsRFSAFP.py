import gsheets
from decimal import Decimal

class GSheetRFSAFP(gsheets.GSheet):
    def __init__(self, kursArg, stdBetrag, stdZweck):
        super().__init__(kursArg, stdBetrag, stdZweck)
        self.spreadSheetId = "1KCh-3tpb1ciF3KVxJeRty0CDDd40vE0UTLGC2bqE8uo"

        # diese Felder brauchen wir für den Einzug
        self.ebicsnames = ebicsnames = ["Lastschrift: Name des Kontoinhabers",
                                        "Lastschrift: IBAN-Kontonummer", "Betrag", "Zweck", "Zeitstempel"]
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
        return "20/35", "ADFC Radfahrschule", "ADFC-M-RFSAFP-2024"

    def validSheetName(self, sname):
        return sname == "Buchungen"

    def checkKtoinh(self, row):
        inh = row[self.ktoinh]
        if len(inh) < 5 or inh.startswith("dto") or inh.startswith("ditto"):
            row[self.ktoinh] = row[self.vorname] + " " + row[self.name]
        return True

    def checkBetrag(self, row):
        mitglied = row[self.mitglied] != ""
        if self.betrag not in row:
            row[self.betrag] = "20" if mitglied else "35"
        row[self.betrag] = Decimal(row[self.betrag].replace(',', '.'))  # 3,14 -> 3.14
        return True
