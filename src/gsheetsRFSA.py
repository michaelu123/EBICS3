import gsheets


class GSheetRFSA(gsheets.GSheet):
    def __init__(self, stdBetrag, stdZweck):
        super().__init__(stdBetrag, stdZweck)
        self.spreadSheetId = "163EnaAw4A0_BP6zuo51VSyz02MNHTfI3_t8MyQ2JUP4"

        # diese Felder brauchen wir für den Einzug
        self.ebicsnames = ebicsnames = ["Lastschrift: Name des Kontoinhabers",
                                        "Lastschrift: IBAN-Kontonummer", "Betrag", "Zweck", "Zeitstempel"]
        self.ktoinh = ebicsnames[0]
        self.iban = ebicsnames[1]
        self.betrag = ebicsnames[2]
        self.zweck = ebicsnames[3]
        self.datum = ebicsnames[4]

        # Felder die wir überprüfen
        self.formnames = formnames = ["Vorname", "Name", "Zustimmung zur SEPA-Lastschrift", "Bestätigung",
                                      "Verifikation", "Anmeldebestätigung"]
        self.vorname = formnames[0]
        self.name = formnames[1]
        self.zustimmung = formnames[2]
        self.bestätigung = formnames[3]  # Bestätigung der Teilnahmebedingungen
        self.verifikation = formnames[4]
        self.anmeldebest = formnames[5]  # wird vom Skript Radfahrschule/Anmeldebestätigung senden ausgefüllt

        # diese Felder fügen wir hinzu
        self.zusatzFelder = zusatzFelder = ["Eingezogen", "Zahlungseingang", "Kommentar"]
        self.eingezogen = zusatzFelder[0]
        self.zahlungseingang = zusatzFelder[1]  # händisch
        self.kommentar = zusatzFelder[2]

    @classmethod
    def getDefaults(cls):
        return "100", "ADFC Radfahrschule", "ADFC-M-RFS-2023"

    def validSheetName(self, sname):
        return sname == "Buchungen"

    def checkKtoinh(self, row):
        inh = row[self.ktoinh]
        if len(inh) < 5 or inh.startswith("dto") or inh.startswith("ditto"):
            row[self.ktoinh] = row[self.vorname] + " " + row[self.name]
        return True
