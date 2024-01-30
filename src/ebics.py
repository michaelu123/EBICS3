import copy
import datetime
import os
import random
from decimal import Decimal, getcontext
from string import digits, ascii_uppercase
from tkinter.messagebox import *
from xml.dom.minidom import *

import gsheetsMTT2
import gsheetsRFSA
import gsheetsRFSAFP
import gsheetsRFSF
import gsheetsSaisonKarte
import gsheetsTK

templateFileDefault = "Default"
decCtx = getcontext()
decCtx.prec = 7  # 5.2 digits, max=99999.99
charset = digits + ascii_uppercase

xmls = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Document xmlns="urn:iso:std:iso:20022:tech:xsd:pain.008.001.02" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:schemaLocation="urn:iso:std:iso:20022:tech:xsd:pain.008.001.02 pain.008.001.02.xsd">
    <CstmrDrctDbtInitn>
        <GrpHdr>
            <MsgId>MSG26022a8fb83cf1a515099ade7bdc3afc</MsgId>
            <CreDtTm>2019-03-27T11:25:44.620Z</CreDtTm>
            <NbOfTxs>1</NbOfTxs>
            <CtrlSum>0.01</CtrlSum>
            <InitgPty>
                <Nm>ALLG. DEUTSCHER FAHRRAD-CLUB KREISVERBAND MÜNCH. ADFC</Nm>
            </InitgPty>
        </GrpHdr>
        <PmtInf>
            <PmtInfId>PIIa671997ba9d14b0085f75f1353e9d008</PmtInfId>
            <PmtMtd>DD</PmtMtd>
            <NbOfTxs>1</NbOfTxs>
            <CtrlSum>0.01</CtrlSum>
            <PmtTpInf>
                <SvcLvl>
                    <Cd>SEPA</Cd>
                </SvcLvl>
                <LclInstrm>
                    <Cd>CORE</Cd>
                </LclInstrm>
                <SeqTp>OOFF</SeqTp>
            </PmtTpInf>
            <ReqdColltnDt>2019-03-29</ReqdColltnDt>
            <Cdtr>
                <Nm>ALLG. DEUTSCHER FAHRRAD-CLUB KREISVERBAND MÜNCH. ADFC</Nm>
            </Cdtr>
            <CdtrAcct>
                <Id>
                    <IBAN>DE62701500000904157781</IBAN>
                </Id>
            </CdtrAcct>
            <CdtrAgt>
                <FinInstnId>
                    <BIC>SSKMDEMMXXX</BIC>
                </FinInstnId>
            </CdtrAgt>
            <ChrgBr>SLEV</ChrgBr>
            <CdtrSchmeId>
                <Id>
                    <PrvtId>
                        <Othr>
                            <Id>DE44ZZZ00000793122</Id>
                            <SchmeNm>
                                <Prtry>SEPA</Prtry>
                            </SchmeNm>
                        </Othr>
                    </PrvtId>
                </Id>
            </CdtrSchmeId>
            <DrctDbtTxInf>
                <PmtId>
                    <EndToEndId>NOTPROVIDED</EndToEndId>
                </PmtId>
                <InstdAmt Ccy="EUR">0.01</InstdAmt>
                <DrctDbtTx>
                    <MndtRltdInf>
                        <MndtId>ADFC-M-RFS-2018</MndtId>
                        <DtOfSgntr>2019-03-27</DtOfSgntr>
                    </MndtRltdInf>
                </DrctDbtTx>
                <DbtrAgt>
                    <FinInstnId>
                        <Othr>
                            <Id>NOTPROVIDED</Id>
                        </Othr>
                    </FinInstnId>
                </DbtrAgt>
                <Dbtr>
                    <Nm>Vorname Nachname</Nm>
                </Dbtr>
                <DbtrAcct>
                    <Id>
                        <IBAN>DE12341234123412341234</IBAN>
                    </Id>
                </DbtrAcct>
                <RmtInf>
                    <Ustrd>Zweck</Ustrd>
                </RmtInf>
            </DrctDbtTxInf>
        </PmtInf>
    </CstmrDrctDbtInitn>
</Document>
"""


def randomId(length):
    r1 = random.choice(ascii_uppercase)  # first a letter
    r2 = [random.choice(charset) for _ in range(length - 1)]  # then any mixture of capitalletters and numbers
    return r1 + ''.join(r2)


latin = set()
for c in "0123456789":
    latin.add(c)
for c in "abcdefghijklmnopqrstuvwxyz":
    latin.add(c)
for c in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
    latin.add(c)
for c in "':?,-(+.)/ ":
    latin.add(c)
for c in "ÄÖÜäöüß&*$%":
    latin.add(c)


def isLatin(s):
    for c in s:
        if c not in latin:
            return False
    return True


def convertToIsoDate(ts):  # 06.03.2022 17:28:38 -> 2022-03-06
    if ts[2] == '.' and ts[5] == '.':
        return ts[6:10] + "-" + ts[3:5] + "-" + ts[0:2]
    return "2024-01-01"  # ??


klasses = {
    "RFSA": gsheetsRFSA.GSheetRFSA,
    "RFSAFP": gsheetsRFSAFP.GSheetRFSAFP,
    "RFSF": gsheetsRFSF.GSheetRFSF,
    "SK": gsheetsSaisonKarte.GSheetSK,
    "TK": gsheetsTK.GSheetTK,
    "MTT": gsheetsMTT2.GSheetMTT2,  # edoobox
    # "MTT": gsheetsMTT.GSheetMTT # Google Forms/Sheets
}


def getKlasses():
    return klasses


class Ebics:
    def __init__(self, sels, outputfile, stdbetrag, stdzweck, mandat, ebics):
        self.sels = sels
        self.outputFile = outputfile
        self.stdbetrag = stdbetrag
        self.stdzweck = stdzweck
        self.mandat = mandat
        self.ebics = ebics
        self.gsheets = []
        self.xmlt = None
        self.gsheet = None

    def fillinIDs(self):
        msgid = self.xmlt.getElementsByTagName("MsgId")
        val = "MSG" + randomId(32)
        msgid[0].childNodes[0] = self.xmlt.createTextNode(val)
        piid = self.xmlt.getElementsByTagName("PmtInfId")
        val = "PII" + randomId(32)
        piid[0].childNodes[0] = self.xmlt.createTextNode(val)

    def fillinSumme(self, summe, cnt):
        ctrlSum = self.xmlt.getElementsByTagName("CtrlSum")
        for cs in ctrlSum:
            cs.childNodes[0] = self.xmlt.createTextNode(str(summe))
        nbOfTxs = self.xmlt.getElementsByTagName("NbOfTxs")
        for nr in nbOfTxs:
            nr.childNodes[0] = self.xmlt.createTextNode(str(cnt))

    def fillin1(self):
        pmtInf = self.xmlt.getElementsByTagName("PmtInf")[0]
        drctDbtTxInf = pmtInf.getElementsByTagName("DrctDbtTxInf")[0]
        x = pmtInf.childNodes.index(drctDbtTxInf)
        drctDbtTxInf = copy.deepcopy(drctDbtTxInf)
        nl1 = pmtInf.childNodes[x - 1]
        nl2 = pmtInf.childNodes[x + 1]
        pmtInf.childNodes = pmtInf.childNodes[0:x]
        return pmtInf, drctDbtTxInf, nl1, nl2

    def fillin2(self, ttuple, entries):
        pmtInf, drctDbtTxInf, nl1, nl2 = ttuple
        res = []
        for entry in entries:
            ktoInh = entry[self.gsheet.ktoinh]
            if not isLatin(ktoInh):
                showerror("Zeichensatz", "Kontoinhaber " + ktoInh + " mit ungültigen Zeichen")
                self.gsheet.nr_einzuziehen -= 1
                continue
            zweck = str(entry[self.gsheet.zweck])
            if not isLatin(zweck):
                showerror("Zeichensatz", "Zweck " + zweck + " mit ungültigen Zeichen")
                self.gsheet.nr_einzuziehen -= 1
                continue
            res.append(entry)
            newtx = copy.deepcopy(drctDbtTxInf)
            nm = newtx.getElementsByTagName("Nm")
            nm[0].childNodes[0] = self.xmlt.createTextNode(ktoInh)
            ibn = newtx.getElementsByTagName("IBAN")
            ibn[0].childNodes[0] = self.xmlt.createTextNode(entry[self.gsheet.iban])
            amt = newtx.getElementsByTagName("InstdAmt")
            amt[0].childNodes[0] = self.xmlt.createTextNode(str(entry[self.gsheet.betrag]))
            ustrd = newtx.getElementsByTagName("Ustrd")
            ustrd[0].childNodes[0] = self.xmlt.createTextNode(str(zweck))
            mndtId = newtx.getElementsByTagName("MndtId")
            mandat = entry.get("mandat")
            if mandat is None:
                mandat = self.mandat
            mndtId[0].childNodes[0] = self.xmlt.createTextNode(mandat)

            dtOfSgntr = newtx.getElementsByTagName("DtOfSgntr")
            timestamp = entry.get(self.gsheet.datum)
            timestamp = convertToIsoDate(timestamp)
            dtOfSgntr[0].childNodes[0] = self.xmlt.createTextNode(timestamp)

            pmtInf.childNodes.append(newtx)
            pmtInf.childNodes.append(copy.copy(nl1))
        return res

    def fillin3(self, ttuple):
        pmtInf, drctDbtTxInf, nl1, nl2 = ttuple
        pmtInf.childNodes[len(pmtInf.childNodes) - 1] = copy.copy(nl2)

    def fillinDates(self):
        creDtTm = self.xmlt.getElementsByTagName("CreDtTm")
        now = datetime.datetime.utcnow()
        d = now.isoformat(timespec="milliseconds") + "Z"
        creDtTm[0].childNodes[0] = self.xmlt.createTextNode(d)
        reqdColltnDt = self.xmlt.getElementsByTagName("ReqdColltnDt")
        day2 = datetime.date.today() + datetime.timedelta(days=2)
        d = day2.isoformat()
        reqdColltnDt[0].childNodes[0] = self.xmlt.createTextNode(d)

    def check(self):
        for i, sel in enumerate(self.sels.keys()):
            if self.sels[sel] == 0:
                continue
            klass = klasses[sel]
            defaults = klass.getDefaults()
            stdBetrag = self.stdbetrag if self.stdbetrag != "" else defaults[0]
            stdZweck = self.stdzweck if self.stdzweck != "" else defaults[1]
            self.mandat = defaults[2]
            self.gsheet = klass(stdBetrag, stdZweck)
            self.gsheets.append(self.gsheet)
            self.gsheet.getEntries()

    def addBetraege(self, entries):
        ssum = Decimal("0.00")
        for row in entries:
            ssum = ssum + row[self.gsheet.betrag]
        return ssum

    def createEbicsXml(self, setEingezogen):
        template = xmls
        if self.ebics is not None and self.ebics != "" and self.ebics != templateFileDefault:
            with open(self.ebics, "r", encoding="utf-8") as f:
                template = f.read()
        if os.path.exists(self.outputFile):
            raise Exception("Datei " + self.outputFile + " existiert schon")
        self.xmlt = parseString(template)
        self.fillinIDs()
        self.fillinDates()
        ttuple = self.fillin1()

        summe = 0
        entryCnt = 0
        for i, sel in enumerate(self.sels.keys()):
            if self.sels[sel] == 0:
                continue
            klass = klasses[sel]
            defaults = klass.getDefaults()
            stdBetrag = self.stdbetrag if self.stdbetrag != "" else defaults[0]
            stdZweck = self.stdzweck if self.stdzweck != "" else defaults[1]
            if self.mandat == "":  # GUI may have overridden mandat, but only for one spreadsheet
                self.mandat = defaults[2]
            self.gsheet = klass(stdBetrag, stdZweck)
            self.gsheets.append(self.gsheet)
            entries = self.gsheet.getEntries()
            entries = self.fillin2(ttuple, entries)
            entryCnt += len(entries)
            summe += self.addBetraege(entries)
            self.mandat = ""
        self.fillin3(ttuple)

        if entryCnt == 0:
            return
        self.fillinSumme(summe, entryCnt)
        # self.xmlt.standalone = True or "yes" has no effect
        pr = self.xmlt.toxml(encoding="utf-8")
        pr = pr[0:36] + b' standalone="yes"' + pr[36:]  # minidom has problems with standalone param
        with open(self.outputFile, "wb") as o:
            o.write(pr)
        if not setEingezogen:
            return "TESTED OK"
        # Spalte "Eingezogen" auf heutiges Datum setzen
        for gsheet in self.gsheets:
            gsheet.fillEingezogen()
        return "OK"

    def getStatistics(self):
        nr_einzug = nr_unverifiziert = nr_bezahlt = nr_eingezogen = nr_einzuziehen = 0
        for gsheet in self.gsheets:
            st = gsheet.getStatistics()
            nr_einzug += st[0]
            nr_unverifiziert += st[1]
            nr_bezahlt += st[2]
            nr_eingezogen += st[3]
            nr_einzuziehen += st[4]
        return nr_einzug, nr_unverifiziert, nr_bezahlt, nr_eingezogen, nr_einzuziehen
