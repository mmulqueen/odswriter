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


class BaseWriter(object):
    """
    Base class for ODS/FODS writing, contains methods for setting up and managing sheets. Should 
    not be instantiated directly.
    """
    def __init__(self):
        # Make the skeleton of an ODS.
        self.dom = parseString(ods_components.content_xml)
        # Setup sheets
        self.default_sheet = None
        self.sheets = []

    def __enter__(self):
        return self

    def __exit__(self, *args, **kwargs):
        self.close()

    def close(self):
        raise NotImplementedError(
            "BaseWriter should not be instantiated directly. Please use either ODSWriter for .ods "
            "(zipped) spreadsheets or FODSWriter for .fods (human-readable) spreadsheets."
        )

    def writerow(self, cells):
        """
        Write a row of cells into the default sheet of the spreadsheet.
        :param cells: A list of cells (most basic Python types supported).
        :return: Nothing.
        """
        if self.default_sheet is None:
            self.default_sheet = self.new_sheet()
        self.default_sheet.writerow(cells)

    def writerows(self, rows):
        """
        Write rows into the default sheet of the spreadsheet.
        :param rows: A list of rows, rows are lists of cells - see writerow.
        :return: Nothing.
        """
        for row in rows:
            self.writerow(row)

    def new_sheet(self, name=None, cols=None):
        """
        Create a new sheet in the spreadsheet and return it so content can be added.
        :param name: Optional name for the sheet.
        :param cols: Specify the number of columns, needed for compatibility in some cases
        :return: Sheet object
        """
        sheet = Sheet(self.dom, name, cols)
        self.sheets.append(sheet)
        return sheet


class ODSWriter(BaseWriter):
    """
    Utility for writing OpenDocument Spreadsheets (.ods). Can be used in simple 1 sheet mode (use 
    writerow/writerows) or with multiple sheets (use new_sheet). It is suggested that you use with 
    object like a context manager.
    """
    def __init__(self, odsfile, compression=ZIP_STORED):
        # Initialise parent class
        BaseWriter.__init__(self)
        # Setup zip file
        self.zipf = ZipFile(odsfile, "w", compression)
        self.zipf.writestr("mimetype",
                           ods_components.mimetype.encode("utf-8"))
        self.zipf.writestr("META-INF/manifest.xml",
                           ods_components.manifest_xml.encode("utf-8"))
        self.zipf.writestr("styles.xml",
                           ods_components.styles_xml.encode("utf-8"))
    
    def close(self):
        """
        Finalises the compressed version of the spreadsheet. If you aren't using the context 
        manager ('with' statement, you must call this manually, it is not triggered automatically 
        like on a file object.
        :return: Nothing.
        """
        self.zipf.writestr("content.xml", self.dom.toxml().encode("utf-8"))
        self.zipf.close()


class FODSWriter(BaseWriter):
    """
    Utility for writing Flat OpenDocument Spreadsheets (.fods). Can be used in simple 1 sheet mode 
    (use writerow/writerows) or with multiple sheets (use new_sheet). It is suggested that you use 
    with object like a context manager.
    """
    def __init__(self, fodsfile):
        # Initialise parent class
        BaseWriter.__init__(self)
        # Setup xml file
        self.xmlf = fodsfile
    
    def close(self):
        """
        Finalises the flat version of the spreadsheet. If you aren't using the context 
        manager ('with' statement, you must call this manually, it is not triggered automatically 
        like on a file object.
        :return: Nothing.
        """
        self.xmlf.write(self.dom.toxml())
        self.xmlf.close()


class Sheet(object):
    def __init__(self, dom, name="Sheet 1", cols=None):
        self.dom = dom
        self.cols = cols
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

    def writerow(self, cells):
        row = self.dom.createElement("table:table-row")
        content_cells = 0

        for cell_data in cells:
            cell = self.dom.createElement("table:table-cell")
            text = None

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
                p = self.dom.createElement("text:p")
                p.appendChild(self.dom.createTextNode(text))
                cell.appendChild(p)

            row.appendChild(cell)

            content_cells += 1

        if self.cols is not None:
            if content_cells > self.cols:
                raise Exception("More cells than cols.")

            for _ in range(content_cells, self.cols):
                cell = self.dom.createElement("table:table-cell")
                row.appendChild(cell)

        self.table.appendChild(row)

    def writerows(self, rows):
        for row in rows:
            self.writerow(row)


def writer(odsfile, *args, **kwargs):
    """
        Returns an ODSWriter (.osd) or FODSWriter (.fosd) object, depending on the file extension.

        Python 3: Make sure that the file you pass is mode "wb" for .ods files and mode "w" for 
        .fosd files:
        f = open("spreadsheet.ods", "wb")
        odswriter.writer(f)
        f = open("spreadsheet.fods", "w")
        odswriter.writer(f)
        ...
        Otherwise you will get "TypeError"
    """
    if odsfile.name.endswith(".fods"):
        # FODS mode
        return FODSWriter(odsfile, *args, **kwargs)
    else:
        # ODS mode
        return ODSWriter(odsfile, *args, **kwargs)
