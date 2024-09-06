from __future__ import unicode_literals
from zipfile import ZipFile, ZIP_STORED
import decimal
import datetime
from xml.dom.minidom import parseString

from . import ods_components
from .formula import Formula

# Basic compatibility setup for Python 2 and Python 3.

try:
    long
except NameError:
    long = int

try:
    unicode
except NameError:
    unicode = str

# End compatibility setup.


class ODSWriter(object):
    """
    Utility for writing OpenDocument Spreadsheets. Can be used in simple 1 sheet mode (use writerow/writerows) or with
    multiple sheets (use new_sheet). It is suggested that you use with object like a context manager.
    """
    def __init__(self, odsfile, compression=ZIP_STORED, first_row_bold=False):
        self.zipf = ZipFile(odsfile, "w", compression)
        # Make the skeleton of an ODS.
        self.dom = parseString(ods_components.content_xml)
        self.zipf.writestr("mimetype",
                           ods_components.mimetype.encode("utf-8"))
        self.zipf.writestr("META-INF/manifest.xml",
                           ods_components.manifest_xml.encode("utf-8"))
        self.zipf.writestr("styles.xml",
                           ods_components.styles_xml.encode("utf-8"))
        self.default_sheet = None
        self.number_of_sheets = 1
        self.sheets = []
        self.first_row_bold = first_row_bold

    def __enter__(self):
        return self

    def __exit__(self, *args, **kwargs):
        self.close()

    def close(self):
        """
        Finalises the compressed version of the spreadsheet. If you aren't using the context manager ('with' statement,
        you must call this manually, it is not triggered automatically like on a file object.
        :return: Nothing.
        """

        if self.default_sheet is not None:
            self.default_sheet.finalizeFormat(1)
        sheet_index = 2
        for finalSheet in self.sheets:
            finalSheet.finalizeFormat(sheet_index)
            sheet_index += 1

        xmlContent = self.dom.toprettyxml(encoding="utf-8")
        open("TestXML.xml", "wb").write(xmlContent)
        self.zipf.writestr("content.xml", self.dom.toxml().encode("utf-8"))
        self.zipf.close()

    def writerow(self, cells):
        """
        Write a row of cells into the default sheet of the spreadsheet.
        :param cells: A list of cells (most basic Python types supported).
        :return: Nothing.
        """
        if self.default_sheet is None:
            self.default_sheet = self.new_sheet(first_row_bold = self.first_row_bold)
        self.default_sheet.writerow(cells)

    def writerows(self, rows):
        """
        Write rows into the default sheet of the spreadsheet.
        :param rows: A list of rows, rows are lists of cells - see writerow.
        :return: Nothing.
        """
        for row in rows:
            self.writerow(row)

    def new_sheet(self, name=None, cols=None, first_row_bold=False):
        """
        Create a new sheet in the spreadsheet and return it so content can be added.
        :param name: Optional name for the sheet.
        :param cols: Specify the number of columns, needed for compatibility in some cases
        :return: Sheet object
        """
        sheet = Sheet(self.dom, name, cols, sheet_first_row_bold=first_row_bold)
        self.sheets.append(sheet)
        self.number_of_sheets += 1
        return sheet

# Note the way the columns are described in the XML file, specifically, how they're repeated.
# Say that ONLY the 3rd operties fo:break-before="auto" style:column-width="81.04pt"/>
#    </style:style>

 #<table:table-column table:style-name="co2" table:default-cell-style-name="Default"/>

class Sheet(object):

    def __init__(self, dom, name="Sheet 1", cols=None, sheet_first_row_bold=False):
        self.dom = dom
        self.cols = cols
        self.max_cols_content = 0
        self.col_widths = []
        self.sheet_name = name
        self.first_row_bold = sheet_first_row_bold
        self.width_factor = 10  # Number of points each charachter is wide
        spreadsheet = self.dom.getElementsByTagName("office:spreadsheet")[0]
        self.table = self.dom.createElement("table:table")
        if name:
            self.table.setAttribute("table:name", name)
        self.table.setAttribute("table:style-name", "ta1")

        if self.cols is not None:
            col = self.dom.createElement("table:table-column")
            col.setAttribute("table:number-columns-repeated", unicode(self.cols))
            self.table.appendChild(col)

        spreadsheet.appendChild(self.table)

    def finalizeFormat(self, sheet_index):

        styles = self.dom.getElementsByTagName("office:automatic-styles")[0]
        first_row = self.table.getElementsByTagName("table:table-row")[0]

        for i in range(0, self.max_cols_content):
            style = self.dom.createElement("style:style")
            style.setAttribute("style:name", "sh" + str(sheet_index) + "col" + str(i))
            style.setAttribute("style:family", "table-column")
            styleprops = self.dom.createElement("style:table-column-properties")
            styleprops.setAttribute("fo:break-before", "auto")
            styleprops.setAttribute("style:column-width", str(self.col_widths[i]) + "pt")
            style.appendChild(styleprops)
            styles.appendChild(style)

        for i in range(0, self.max_cols_content):
            col = self.dom.createElement("table:table-column")
            col.setAttribute("table:style-name", "sh" + str(sheet_index) + "col" + str(i))
            col.setAttribute("table:default-cell-style-name", "Default")
            self.table.insertBefore(col, first_row)



    def writerow(self, cells):
        row = self.dom.createElement("table:table-row")
        content_cells = 0

        for cell_data in cells:
            cell = self.dom.createElement("table:table-cell")
            text = None
            curr_cell_width = 0

            if isinstance(cell_data, (datetime.date, datetime.datetime)):
                cell.setAttribute("office:value-type", "date")
                date_str = cell_data.isoformat()
                cell.setAttribute("office:date-value", date_str)
                cell.setAttribute("table:style-name", "cDateISO")
                text = date_str

            elif isinstance(cell_data, datetime.time):
                cell.setAttribute("office:value-type", "time")
                cell.setAttribute("office:time-value",
                                  cell_data.strftime("PT%HH%MM%SS"))
                cell.setAttribute("table:style-name", "cTime")
                text = cell_data.strftime("%H:%M:%S")

            elif isinstance(cell_data, bool):
                # Bool condition must be checked before numeric because:
                # isinstance(True, int): True
                # isinstance(True, bool): True
                cell.setAttribute("office:value-type", "boolean")
                cell.setAttribute("office:boolean-value",
                                  "true" if cell_data else "false")
                cell.setAttribute("table:style-name", "cBool")
                text = "TRUE" if cell_data else "FALSE"

            elif isinstance(cell_data, (float, int, decimal.Decimal, long)):
                cell.setAttribute("office:value-type", "float")
                float_str = unicode(cell_data)
                cell.setAttribute("office:value", float_str)
                text = float_str

            elif isinstance(cell_data, Formula):
                cell.setAttribute("table:formula", str(cell_data))

            elif cell_data is None:
                pass  # Empty element

            else:
                # String and unknown types become string cells
                cell.setAttribute("office:value-type", "string")
                text = unicode(cell_data)

            if text:
                # Multiline strings mean multiple text:p elements
                # TODO: Handle different sorts of carriage returns properly
                # https://stackoverflow.com/questions/1059559/split-strings-with-multiple-delimiters

                multiline = False
                chop_line = False

                for datum in text.split('\n'):
                    if multiline:
                        cell.setAttribute("table:style-name", "wrapStyle")
                        chop_line = True
                    p = self.dom.createElement("text:p")
                    datum = datum.strip()
                    p.appendChild(self.dom.createTextNode(datum))
                    cell.appendChild(p)
                    if curr_cell_width < len(datum):
                        curr_cell_width = len(datum)
                    multiline = True

                if chop_line and len(datum.strip()) >= 20:
                    curr_cell_width = 20

            if self.first_row_bold:
                cell.setAttribute("table:style-name", "boldStyle")


            row.appendChild(cell)

            curr_cell_width *= self.width_factor

            if len(self.col_widths) > content_cells and len(self.col_widths) > 0:
                if self.col_widths[content_cells] < curr_cell_width:
                    self.col_widths[content_cells] = curr_cell_width
            else:
                self.col_widths.append(curr_cell_width)

            content_cells += 1
            curr_cell_width = 0

        if self.cols is not None:
            if content_cells > self.cols:
                raise Exception("More cells than cols.")

            for _ in range(content_cells, self.cols):
                cell = self.dom.createElement("table:table-cell")
                row.appendChild(cell)

        if content_cells > self.max_cols_content:
            self.max_cols_content = content_cells

        self.table.appendChild(row)
        self.first_row_bold = False

    def writerows(self, rows):
        for row in rows:
            self.writerow(row)


def writer(odsfile, *args, **kwargs):
    """
        Returns an ODSWriter object.

        Python 3: Make sure that the file you pass is mode b:
        f = open("spreadsheet.ods", "wb")
        odswriter.writer(f)
        ...
        Otherwise you will get "TypeError: must be str, not bytes"
    """
    return ODSWriter(odsfile, *args, **kwargs)
