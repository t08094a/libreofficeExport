from odf.opendocument import load
from odf.table import TableRow
from odf.text import P
from odf.namespaces import *

class OdfReader:
    m_TableKey = (TABLENS, 'name')
    m_ExpectedTableName = "Stammliste_aktuell"

    def __init__(self, docFilename):
        """
        Open an ODF file.
        """
        self.filename = docFilename
        print("operate on: " + docFilename)

    def __validate(self, ods):
        """
        validates the given ods document if it has a table 'Stammliste_aktuell'
        :param ods: the ods document
        :return: nothing
        """
        childNodes = ods.spreadsheet.childNodes

        assert len(childNodes) > 0
        attributes = childNodes[1].attributes

        exists = self.m_TableKey in attributes
        assert exists


        tableName = attributes[self.m_TableKey]
        assert tableName == self.m_ExpectedTableName

        print("Found table: " + self.m_ExpectedTableName)

    def __findStammlisteTable(self, ods):

        childNodes = ods.spreadsheet.childNodes

        for table in childNodes:
            if len(table.attributes) <= 0:
                continue

            if self.m_TableKey not in table.attributes:
                continue

            tableName = table.attributes[self.m_TableKey]

            if tableName != self.m_ExpectedTableName:
                continue

            return table

        return None

    def __getRelevantRows(self, table):
        result = []

        rows = table.getElementsByType(TableRow)
        for row in rows:
            isRelevant = self.__isCurrentRowRelevant(row)

            if isRelevant:
                print("    append row")
                result.append(row)


        return result

    def __isCurrentRowRelevant(self, row):

        firstCell = row.firstChild

        if firstCell is None:
            return False

        textElementList = firstCell.getElementsByType(P)
        if len(textElementList) == 0 or len(textElementList[0].childNodes) == 0:
            return False

        content = textElementList[0].firstChild.data
        print("found row: " + content)

        if not content and content.isdigit():
            return False

        rowNumber = int(content)
        if rowNumber > 0:
            return True

        return False

    def __convertRowsToXml(self, rows):
        xml = "<?xml version='1.0' encoding='UTF-8'?>\n" \
              "<items>"

        for row in rows:
            rowAsXml = self.__convertRowToXml(row)
            xml += "\n" + rowAsXml

        xml += "\n</items>"

        return xml

    def __convertRowToXml(self, row):
        offsetLp = 34+4
        offsetAusbildung = offsetLp + 12 + 7
        offsetVerfügbarkeit = offsetAusbildung + 8 + 1

        ids = {0:"Nummer", 2:"aktivUeber18", 3:"aktivUnter18", 4:"maennlich", 5:"weiblich",
               7:"vereinAktiv", 9:"vereinPassiv", 11:"vereinFoerdernd",
               14:"rang",15:"gruppe",
               17:"nachname", 18:"vorname", 19:"strasse", 20:"hausnummer", 21:"plz", 22:"ort", 23:"geburtsdatum", 25:"telefon", 26:"mobil", 27: "email", 28: "infoPerMail", 29:"sonstigeErreichbarkeit",
               33:"eintrittAktiv", 34:"endeAktiv",
               offsetLp+1:"hl1", offsetLp+2:"hl2", offsetLp+3:"hl3", offsetLp+4: "hl4", offsetLp+5:"hl5", offsetLp+6:"hl6", offsetLp+7:"wa1", offsetLp+8:"wa2", offsetLp+9:"wa3", offsetLp+10:"wa4", offsetLp+11:"wa5", offsetLp+12:"wa6",
               offsetAusbildung+1:"ausbildungGA", offsetAusbildung+2:"ausbildungTM", offsetAusbildung+3:"ausbildungGF", offsetAusbildung+4:"ausbildungZF", offsetAusbildung+5:"ausbildungVF", offsetAusbildung+6:"ausbildungFunk", offsetAusbildung+7:"ausbildungMA", offsetAusbildung+8:"ausbildungAT",
               offsetVerfügbarkeit+1:"verfWocheTag", offsetVerfügbarkeit+3:"verfWocheNacht", offsetVerfügbarkeit+5:"verfWochenendeTag", offsetVerfügbarkeit+7:"verWochenendeNacht"}

        xml = "    <item>\n"

        idsAsList = ids.items()

        cells = []
        for cell in row.childNodes:
            cells.append(cell)

            # if cell has a repeated attribute -> add it n-times
            repeatAttr = cell.getAttribute('numbercolumnsrepeated') # auch: number-columns-spanned
            if repeatAttr is not None:
                strValue = str(repeatAttr)

                if strValue.isdigit():
                    value = int(strValue)

                    for x in range(0, value - 1):
                        cells.append(cell)

        for pair in idsAsList:
            index = pair[0]
            name = pair[1]

            if index >= len(cells):
                continue

            cell = cells[index]
            value = str(cell)
            xml += "        <{0}>{1}</{0}>\n".format(name, value)

        xml += "    </item>"
        return xml

    def parse(self):
        """
        parse the content
        """
        ods = load(self.filename)

        self.__validate(ods)

        table = self.__findStammlisteTable(ods)
        assert(table is not None)

        rows = self.__getRelevantRows(table)

        xml = self.__convertRowsToXml(rows)

        return xml
