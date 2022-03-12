import gsheets


class GSheetSK(gsheets.GSheet):
    def __init__(self, _stdBetrag, _stdZweck):
        super().__init__("22", "Saisonkarte 2022")
        self.spreadSheetId = "1IsG9HpZlDU97Sf82LG-XkTPY1stI_xONH_pDP3x72BU"  # Saisonkarten-Bestellungen
        # Mit dem aktuellen credentials.json brauchen wir dafür eine externe Linkfreigabe für adfc-muc zum Bearbeiten
        self.spreadSheetName = "Saisonkarten-Bestellungen"

        self.ebicsnames = ebicsnames = ["Name des Kontoinhabers (kann leer bleiben falls gleich Mitgliedsname)",
                                        "IBAN-Kontonummer", "Betrag", "Zweck", "Zeitstempel"]
        self.ktoinh = ebicsnames[0]
        self.iban = ebicsnames[1]
        self.betrag = ebicsnames[2]
        self.zweck = ebicsnames[3]
        self.datum = ebicsnames[4]

        # Felder die wir überprüfen
        self.formnames = formnames = ["ADFC-Mitgliedsname", "ADFC-Mitgliedsnummer",
                                      "Zustimmung zur SEPA-Lastschrift", "Verifikation", "Gesendet"]
        self.mitgliedsname = formnames[0]
        self.mitgliedsnummer = formnames[1]
        self.zustimmung = formnames[2]
        self.verifikation = formnames[3]
        self.gesendet = formnames[4]  # wird vom Skript der Tabelle ausgefüllt

        # diese Felder fügen wir hinzu
        self.zusatzFelder = zusatzFelder = ["Abgebucht", "Bezahlt", "Kommentar"]
        self.eingezogen = zusatzFelder[0]
        self.zahlungseingang = zusatzFelder[1]  # händisch
        self.kommentar = zusatzFelder[2]

    @classmethod
    def getDefaults(cls):
        return "22", "ADFC Saisonkarte", "ADFC-M-SK-2022"

    def validSheetName(self, sname):
        return sname == "Bestellungen"

    def checkKtoinh(self, row):
        if row[self.ktoinh] == "":
            row[self.ktoinh] = row[self.mitgliedsname]
        return True
