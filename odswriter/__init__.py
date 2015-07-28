# The MIT License (MIT)
#
# Copyright (c) 2014 Michael Mulqueen (http://michael.mulqueen.me.uk/)
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

from __future__ import unicode_literals
from zipfile import ZipFile
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
    def __init__(self, odsfile):
        self.zipf = ZipFile(odsfile, "w")
        # Make the skeleton of an ODS.
        self.dom = parseString(ods_components.content_xml)
        self.zipf.writestr("mimetype",
                           ods_components.mimetype.encode("utf-8"))
        self.zipf.writestr("META-INF/manifest.xml",
                           ods_components.manifest_xml.encode("utf-8"))
        self.zipf.writestr("styles.xml",
                           ods_components.styles_xml.encode("utf-8"))
        self.default_sheet = None

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
        self.zipf.writestr("content.xml", self.dom.toxml().encode("utf-8"))
        self.zipf.close()

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

    def new_sheet(self, name=None):
        """
        Create a new sheet in the spreadsheet and return it so content can be added.
        :param name: Optional name for the sheet.
        :return: Sheet object
        """
        return Sheet(self.dom, name)


class Sheet(object):
    def __init__(self, dom, name="Sheet 1"):
        self.dom = dom
        spreadsheet = self.dom.getElementsByTagName("office:spreadsheet")[0]
        self.table = self.dom.createElement("table:table")
        if name:
            self.table.setAttribute("table:name", name)
        self.table.setAttribute("table:style-name", "ta1")
        spreadsheet.appendChild(self.table)

    def writerow(self, cells):
        row = self.dom.createElement("table:table-row")
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

        self.table.appendChild(row)

    def writerows(self, rows):
        for row in rows:
            self.writerow(row)


def writer(odsfile):
    """
        Returns an ODSWriter object.

        Python 3: Make sure that the file you pass is mode b:
        f = open("spreadsheet.ods", "wb")
        odswriter.writer(f)
        ...
        Otherwise you will get "TypeError: must be str, not bytes"
    """
    return ODSWriter(odsfile, *args, **kwargs)
