import datetime
import os
import pickle
import string
import sys
import logging
from decimal import Decimal

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
# from googleapiclient import errors
from googleapiclient.discovery import build

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


def is_iban(unchecked_iban):  # https://gist.github.com/mperlet/f912b1e57d058bd1b07d78ed13de1f23
    LETTERS = {ord(d): str(i) for i, d in enumerate(string.digits + string.ascii_uppercase)}

    def _number_iban(iban):
        return (iban[4:] + iban[:4]).translate(LETTERS)

    def generate_iban_check_digits(iban):
        number_iban = _number_iban(iban[:2] + '00' + iban[4:])
        return '{:0>2}'.format(98 - (int(number_iban) % 97))

    def valid_iban(iban):
        return int(_number_iban(iban)) % 97 == 1

    unchecked_iban = unchecked_iban.replace(' ', "").upper()
    return generate_iban_check_digits(unchecked_iban) == unchecked_iban[2:4] and valid_iban(unchecked_iban)


#  it seems that with "pyinstaller -F" tkinter (resp. TK) does not find data files relative to the MEIPASS dir
def pyinst(path):
    path = path.strip()
    if os.path.exists(path):
        return path
    if hasattr(sys, "_MEIPASS"):  # i.e. if running as exe produced by pyinstaller
        pypath = sys._MEIPASS + "/" + path
        if os.path.exists(pypath):
            return pypath
    return path


class GSheet:
    def __init__(self, kursArg, stdBetrag, stdZweck):
        self.kursArg = kursArg
        self.stdbetrag = stdBetrag
        self.stdzweck = stdZweck
        self.ssheet = None
        self.data = {}
        self.emailAdresses = {}
        self.nr_einzug = 0
        self.nr_bezahlt = 0
        self.nr_einzuziehen = 0
        self.nr_eingezogen = 0
        self.nr_unverifiziert = 0
        self.eingez = []

        self.spreadSheetId = ""

        # diese Felder brauchen wir für den Einzug
        self.ebicsnames = []
        self.ktoinh = ""
        self.iban = ""
        self.betrag = ""
        self.zweck = ""

        # Felder die wir überprüfen
        self.zustimmung = ""
        self.verifikation = ""
        self.kursName = "---"

        # diese Felder fügen wir hinzu
        self.zusatzFelder = []
        self.eingezogen = ""
        self.zahlungseingang = ""
        self.zahlungsbetrag = ""

    def getStatistics(self):
        return self.nr_einzug, self.nr_unverifiziert, self.nr_bezahlt, self.nr_eingezogen, self.nr_einzuziehen

    def getData(self):
        """Calls the Apps Script API.
        """
        creds = None
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    pyinst('credentials.json'), SCOPES)
                creds = flow.run_local_server(port=53876)
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        service = build('sheets', 'v4', credentials=creds)
        self.ssheet = service.spreadsheets()
        sheet_props = self.ssheet.get(spreadsheetId=self.spreadSheetId, fields="sheets.properties").execute()
        sheet_names = [sheet_prop["properties"]["title"] for sheet_prop in sheet_props["sheets"]]
        all_notes = self.ssheet.get(spreadsheetId=self.spreadSheetId,
                                    fields="sheets.properties,sheets/data/rowData/values/note").execute()

        for i, sname in enumerate(sheet_names):
            if not self.validSheetName(sname):
                continue
            try:
                rows = self.ssheet.values().get(spreadsheetId=self.spreadSheetId, range=sname). \
                    execute().get('values', [])
                rowdata = all_notes["sheets"][i]["data"][0]
                if "rowData" in rowdata:
                    rowdata = rowdata["rowData"]
                    for j in range(len(rowdata)):
                        if "values" in rowdata[j]:
                            rows[j][0] = "Notiz"
                self.data[sname] = rows
            except Exception as e:
                logging.exception("Kann Arbeitsblatt " + sname + " nicht laden")
                raise e

    def checkColumns(self):
        # Prüfen ob im sheet die Zusatzfelder angelegt sind
        for sheet in self.data.keys():
            srows = self.data[sheet]
            headers = srows[0]  # ein Pointer nach data, keine Kopie!
            try:
                _ = headers.index("E-Mail-Adresse")
            except:
                continue
            for h in self.zusatzFelder:
                if h not in headers:
                    self.addColumn(sheet, h)
                    headers.append(h)

    def checkBetrag(self, row):
        if self.betrag not in row:
            if self.stdbetrag == "":
                raise ValueError("Standard-Betrag nicht definiert")
            row[self.betrag] = self.stdbetrag
        row[self.betrag] = Decimal(row[self.betrag].replace(',', '.'))  # 3,14 -> 3.14
        return True

    def checkRow(self, row):
        if self.kursArg != "" and self.kursName in row and not row[self.kursName].startswith(self.kursArg):
            return False
        # IBAN angegeben?
        if self.iban not in row.keys() or row[self.iban] == "":
            return False
        row[self.iban] = row[self.iban].upper()
        if not is_iban(row[self.iban]):
            print("falsche iban in", row["Sheet"], "KontoInhaber:", row[self.ktoinh], "IBAN:", row[self.iban])
            return False
        # Zustimmung erteilt?
        if self.zustimmung not in row.keys() or row[self.zustimmung] == "":
            print("keine Zustimmung in", row["Sheet"], "KontoInhaber:", row[self.ktoinh])
            return False
        self.nr_einzug += 1
        # Email verifiziert?
        if self.verifikation not in row.keys() or row[self.verifikation] == "":
            self.nr_unverifiziert += 1
            print("Emailadresse nicht verifiziert in", row["Sheet"], "KontoInhaber:",
                  row[self.ktoinh], "Emailadresse: ", row["E-Mail-Adresse"])
            return False
        # Schon Zahlungseingang
        if self.zahlungseingang in row and row[self.zahlungseingang] != "":
            self.nr_bezahlt += 1
            return False
        # Schon eingezogen?
        if self.eingezogen in row and row[self.eingezogen] != "":
            self.nr_eingezogen += 1
            return False
        self.nr_einzuziehen += 1
        if not self.checkBetrag(row):
            return False
        if not self.checkKtoinh(row):
            return False
        if self.zweck not in row:
            if self.stdzweck == "":
                raise ValueError("Standard-Verwendungszweck nicht definiert")
            row[self.zweck] = self.stdzweck
        return True

    def parseGS(self):
        vals = []
        for sheet in self.data.keys():
            srows = self.data[sheet]
            headers = srows[0]
            try:
                eingezogenX = headers.index(self.eingezogen)
                zahlungsbetragX = headers.index(self.zahlungsbetrag)
            except:
                continue
            for r, srow in enumerate(srows[1:]):
                if len(srow) == 0:
                    continue
                if srow[0] == "Notiz":
                    print("Mit Notiz in Zeile ", r + 2, srow)
                    continue
                row = {}
                for c, v in enumerate(srow):
                    if c < len(headers) and headers[c] != "":
                        key = headers[c]
                    else:
                        while c >= len(headers):
                            headers.append("")
                        key = headers[c] = chr(ord('A') + c)
                    row[key] = v
                row["Sheet"] = sheet
                if self.checkRow(row):
                    # Hier können wir einziehen!
                    vals.append({x: row[x] for x in self.ebicsnames})
                    # Merken wo das Eingezogen-Datum gespeichert wird, nachdem ebics.xml geschrieben wurde
                    self.eingez.append((sheet, r + 1, eingezogenX, zahlungsbetragX, str(row[self.betrag])))
        return vals

    def getEntries(self):
        self.getData()
        self.checkColumns()
        entries = self.parseGS()
        return entries

    def addColumn(self, sheetName, colName):
        col = 0
        for row in self.data[sheetName]:
            col = max(col, len(row))
        self.addValue(sheetName, 0, col, colName)

    def addValue(self, sheetName, row, col, val):
        # row, col are 0 based
        values = [[val]]
        body = {"values": values}
        col0 = "" if col < 26 else chr(ord('A') + int(col / 26) - 1)  # A B ... Z AA AB ... AZ BA BB ... BZ ...
        col1 = chr(ord('A') + int(col % 26))
        srange = sheetName + "!" + col0 + col1 + str(row + 1)  # 0,0-> A1, 1,2->C2 2,1->B3
        if row == 0:
            try:
                result = self.ssheet.values().update(spreadsheetId=self.spreadSheetId,
                                                     range=srange, valueInputOption="RAW", body=body).execute()
            except:
                result = self.ssheet.values().append(spreadsheetId=self.spreadSheetId,
                                                     range=srange, valueInputOption="RAW", body=body).execute()
        else:
            result = self.ssheet.values().update(spreadsheetId=self.spreadSheetId,
                                                 range=srange, valueInputOption="RAW", body=body).execute()

        logging.log(logging.INFO, "result %s", result)

    def fillEingezogen(self):
        # Spalte "Eingezogen" auf heutiges Datum setzen
        now = datetime.datetime.now()
        d = now.strftime("%Y-%m-%d")
        for t in self.eingez:
            self.addValue(t[0], t[1], t[2], d)
            self.addValue(t[0], t[1], t[3], t[4])
