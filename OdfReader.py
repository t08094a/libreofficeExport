from odf.opendocument import load
from odf.table import TableRow
from odf.text import P
from odf.namespaces import *


class OdfReader:
    m_TableKey = (TABLENS, 'name')
    m_ExpectedTableName = "Stammliste_aktuell"

    def __init__(self, doc_filename):
        """
        Open an ODF file.
        """
        self.filename = doc_filename
        print("operate on: " + doc_filename)

    def __validate(self, ods):
        """
        validates the given ods document if it has a table 'Stammliste_aktuell'
        :param ods: the ods document
        :return: nothing
        """
        child_nodes = ods.spreadsheet.childNodes

        assert len(child_nodes) > 0
        attributes = child_nodes[1].attributes

        exists = self.m_TableKey in attributes
        assert exists

        table_name = attributes[self.m_TableKey]
        assert table_name == self.m_ExpectedTableName

        print("Found table: " + self.m_ExpectedTableName)

    def __find_stammliste_table(self, ods):

        child_nodes = ods.spreadsheet.childNodes

        for table in child_nodes:
            if len(table.attributes) <= 0:
                continue

            if self.m_TableKey not in table.attributes:
                continue

            table_name = table.attributes[self.m_TableKey]

            if table_name != self.m_ExpectedTableName:
                continue

            return table

        return None

    def __get_relevant_rows(self, table):
        result = []

        rows = table.getElementsByType(TableRow)
        for row in rows:
            is_relevant = self.__is_current_row_relevant(row)

            if is_relevant:
                print("    append row")
                result.append(row)

        return result

    @staticmethod
    def __is_current_row_relevant(row):

        first_cell = row.firstChild

        if first_cell is None:
            return False

        text_element_list = first_cell.getElementsByType(P)
        if len(text_element_list) == 0 or len(text_element_list[0].childNodes) == 0:
            return False

        content = text_element_list[0].firstChild.data
        print("found row: " + content)

        if not content and content.isdigit():
            return False

        row_number = int(content)
        if row_number > 0:
            return True

        return False

    def __convert_rows_to_xml(self, rows):
        xml = "<?xml version='1.0' encoding='UTF-8'?>\n" \
              "<items>"

        for row in rows:
            row_as_xml = self.__convert_row_to_xml(row)
            xml += "\n" + row_as_xml

        xml += "\n</items>"

        return xml

    @staticmethod
    def __convert_row_to_xml(row):
        offset_lp = 34 + 4
        offset_ausbildung = offset_lp + 12 + 7
        offset_verfuegbarkeit = offset_ausbildung + 8 + 1

        ids = {2: "aktivUeber18", 3: "aktivUnter18", 4: "maennlich", 5: "weiblich",
               7: "vereinAktiv", 9: "vereinPassiv", 11: "vereinFoerdernd",
               14: "rang", 15: "gruppe",
               17: "nachname", 18: "vorname", 19: "strasse", 20: "hausnummer", 21: "plz", 22: "ort", 23: "geburtsdatum",
               25: "telefon", 26: "mobil", 27: "email", 28: "infoPerMail", 29: "sonstigeErreichbarkeit",
               33: "eintrittAktiv", 34: "endeAktiv",
               offset_lp + 1: "hl1", offset_lp + 2: "hl2", offset_lp + 3: "hl3", offset_lp + 4: "hl4", offset_lp + 5: "hl5",
               offset_lp + 6: "hl6", offset_lp + 7: "wa1", offset_lp + 8: "wa2", offset_lp + 9: "wa3", offset_lp + 10: "wa4",
               offset_lp + 11: "wa5", offset_lp + 12: "wa6",
               offset_ausbildung + 1: "ausbildungGA", offset_ausbildung + 2: "ausbildungTM",
               offset_ausbildung + 3: "ausbildungGF", offset_ausbildung + 4: "ausbildungZF",
               offset_ausbildung + 5: "ausbildungVF", offset_ausbildung + 6: "ausbildungFunk",
               offset_ausbildung + 7: "ausbildungMA", offset_ausbildung + 8: "ausbildungAT",
               offset_verfuegbarkeit + 1: "verfWocheTag", offset_verfuegbarkeit + 3: "verfWocheNacht",
               offset_verfuegbarkeit + 5: "verfWochenendeTag", offset_verfuegbarkeit + 7: "verWochenendeNacht"}

        xml = "    <item>\n"

        ids_as_list = ids.items()

        cells = []
        for cell in row.childNodes:
            cells.append(cell)

            # if cell has a repeated attribute -> add it n-times
            repeat_attr = cell.getAttribute('numbercolumnsrepeated')  # auch: number-columns-spanned
            if repeat_attr is not None:
                str_value = str(repeat_attr)

                if str_value.isdigit():
                    value = int(str_value)

                    for x in range(0, value - 1):
                        cells.append(cell)

        for pair in ids_as_list:
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

        table = self.__find_stammliste_table(ods)
        assert (table is not None)

        rows = self.__get_relevant_rows(table)

        xml = self.__convert_rows_to_xml(rows)

        return xml
