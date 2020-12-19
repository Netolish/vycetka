import threading
import time

import uno
import unohelper
from com.sun.star.beans import PropertyValue
from com.sun.star.sheet import XRangeSelectionListener


class Vycetka(object):

    BANKOVKY = [5000, 2000, 1000, 500, 200, 100, 50, 20, 10, 5, 2, 1]

    """Docstring for Vycetka. """

    def __init__(self):
        self.doc = XSCRIPTCONTEXT.getDocument()
        self.sheet = None
        self.first = (0, 0)
        self.last = (0, 0)
        
    def handleRange(self, rangeDescriptor):
        self.sheet = self.doc.getSheets()[Vycetka.getSheet(rangeDescriptor)]
        self.parseRange(rangeDescriptor)
        self.genHeader()
        self.fillVycetka()
        self.fillSum()

    def genHeader(self): 
        (col, row) = self.first
        self.sheet.getRows().insertByIndex(row, 2)
        self.first = (col, row + 2)
        self.last = (self.last[0], self.last[1] + 2)
        cell = self.sheet.getCellByPosition(col, row)
        cell.setString('Bankovka/Mince')
        bankcol = col + 1
        for b in Vycetka.BANKOVKY:
            self.sheet.getCellByPosition(bankcol, row).setValue(b)
            bankcol += 1
        cell = self.sheet.getCellByPosition(col, row + 1)
        cell.setString('Počet')

    def fillVycetka(self):
        for row in range(self.first[1], self.last[1] + 1):
            self.vycetkaRow(self.first[0], row)

    def fillSum(self):
        (col, row) = self.first
        for i in range(len(Vycetka.BANKOVKY)):
            cell = self.sheet.getCellByPosition(col + 1 + i, row - 1)
            formula = '=SUM({}:{})'.format(Vycetka.addr(col + i + 1, self.first[1] + 1),
                    Vycetka.addr(col + 1 + i, self.last[1] + 1))
            cell.setFormula(formula)

    def vycetkaRow(self, col, row):
        nrow = row + 1
        data = self.sheet.getCellByPosition(col, row).getValue()
        if data == 0:
            return
        for i in range(len(Vycetka.BANKOVKY)):
            sumprev = ""
            for j in range(0, i):
                sumprev += " - {}*{}".format(Vycetka.addr(col + 1 + j, nrow),
                                              Vycetka.addr(col + 1 + j,
                                                           self.first[1] - 1, rowabs=True))
            formula = "=FLOOR(({}{})/{})".format(Vycetka.addr(col, nrow), sumprev,
                                                 Vycetka.addr(col + 1 + i,
                                                              self.first[1] - 1, rowabs=True))
            cell = self.sheet.getCellByPosition(col + 1 + i, row)
            cell.setFormula(formula)

    @staticmethod
    def getSheet(rangeDescriptor):
        parts = rangeDescriptor.split('.')
        if len(parts) == 2:
            return parts[0].lstrip('$')
        return ''

    @staticmethod
    def addr(col, row, colabs=False, rowabs=False):
        if colabs:
            cprefix = '$'
        else:
            cprefix = ''
        if rowabs:
            rprefix = '$'
        else:
            rprefix = ''
        return "{}{}{}{}".format(cprefix, Vycetka.colName(col), rprefix, row)

    @staticmethod
    def colName(colnum):
        if colnum < 26:
            return chr(colnum + ord('A'))
        div = colnum // 26
        rem = colnum % 26
        return chr(div - 1 + ord('A')) + chr(rem + ord('A')) 

    @staticmethod
    def colIdx(colname):
        if len(colname) == 1:
            return ord(colname) - ord('A')
        return (ord(colname[0]) - ord('A') + 1) * 26 + ord(colname[1]) - ord('A')

    def parseRange(self, rangeDescriptor):
        parts = rangeDescriptor.split('.')
        if len(parts) == 2:
            begend = parts[1].split(':')
            if len(begend) == 2:
                (dummy, col, row) = begend[0].split('$')
                self.first = (Vycetka.colIdx(col), int(row) - 1)
                (dummy, col, row) = begend[1].split('$')
                self.last = (Vycetka.colIdx(col), int(row) - 1)
    


def vycetka(dummy_context=None):
    controller = XSCRIPTCONTEXT.getDocument().getCurrentController()
    xRngSel = controller
    aListener = ExampleRangeListener()
    xRngSel.addRangeSelectionListener(aListener)
    aArguments = (
        createProp("Title", "Vyberte oblast částek"),
        createProp("CloseOnMouseRelease", False)
        )
    xRngSel.startRangeSelection(aArguments)
    vycetka = Vycetka()
    t1 = WaiterThread(xRngSel, aListener, vycetka)
    t1.start()


class ExampleRangeListener(XRangeSelectionListener, unohelper.Base):
    def __init__(self):
        self.aResult = "not yet"

    def done(self, aEvent):
        self.aResult = aEvent.RangeDescriptor

    def aborted(self, dummy_aEvent):
        self.aResult = "nothing"

    def disposing(self, dummy_aEvent):
        pass

class WaiterThread(threading.Thread):
    def __init__(self, xRngSel, aListener, handler):
        threading.Thread.__init__(self)
        self.xRngSel = xRngSel
        self.aListener = aListener
        self.handler = handler

    def run(self):
        for dummy in range(120):  # don't wait more than 60 seconds
            if self.aListener.aResult != "not yet":
                break
            time.sleep(0.5)
        self.xRngSel.removeRangeSelectionListener(self.aListener)
        if self.aListener.aResult != "not yet":
            self.handler.handleRange(self.aListener.aResult)

def createProp(name, value):
    """Creates an UNO property."""
    prop = PropertyValue()
    prop.Name = name
    prop.Value = value
    return prop

