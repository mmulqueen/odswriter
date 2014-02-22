from __future__ import unicode_literals
from zipfile import ZipFile
import decimal
import datetime
from xml.dom.minidom import parseString

import odswriter.ods_components as ods_components
from odswriter.formula import Formula

try:
    long
except NameError:
    long = int

try:
    unicode
except NameError:
    unicode = str

class ODSWriter(object):
    def __init__(self, odsfile):
        self.zipf = ZipFile(odsfile,"w")
        self.make_skeleton()
    def make_skeleton(self):
        self.dom = parseString(ods_components.content_xml)
        self.table = self.dom.getElementsByTagName("table:table")[0]
        self.zipf.writestr("manifest.rdf",ods_components.manifest_rdf.encode("utf-8"))
        self.zipf.writestr("mimetype",ods_components.mimetype.encode("utf-8"))
        self.zipf.writestr("META-INF/manifest.xml",ods_components.manifest_xml.encode("utf-8"))

    def __enter__(self):
        return self

    def __exit__(self, type, value, tb):
        self.close()

    def close(self):
        self.zipf.writestr("content.xml", self.dom.toxml().encode("utf-8"))
        self.zipf.close()

    def writerow(self, cells):
        row = self.dom.createElement("table:table-row")
        for cell_data in cells:
            cell = self.dom.createElement("table:table-cell")
            text = None
            cell_type = type(cell_data)
            if cell_type in (datetime.date,datetime.datetime):
                cell.setAttribute("office:value-type", "date")
                date_str = text=cell_data.isoformat()
                cell.setAttribute("office:date-value", date_str)
                cell.setAttribute("table:style-name", "cDateISO")
                text=date_str
            elif cell_type == datetime.time:
                cell.setAttribute("office:value-type", "time")
                cell.setAttribute("office:time-value", cell_data.strftime("PT%HH%MM%SS"))
                cell.setAttribute("table:style-name", "cTime")
                text=cell_data.strftime("%H:%M:%S")
            elif cell_type in (float, int, decimal.Decimal, long):
                cell.setAttribute("office:value-type", "float")
                float_str = unicode(cell_data)
                cell.setAttribute("office:value", float_str)
                text=float_str
            elif cell_type == bool:
                cell.setAttribute("office:value-type", "boolean")
                cell.setAttribute("office:boolean-value", "true" if cell_data else "false")
                cell.setAttribute("table:style-name", "cBool")
                text="TRUE" if cell_data else "FALSE"
            elif isinstance(cell_data, Formula):
                 cell.setAttribute("table:formula", str(cell_data))
            elif cell_type == type(None):
                pass # Empty element
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

def writer(odsfile, **kwargs):
    """
        Returns an ODSWriter object.

        Python 3: Make sure that the file you pass is mode b:
        f = open("spreadsheet.ods", "wb")
        odswriter.writer(f)
        ...
        Otherwise you will get "TypeError: must be str, not bytes"
    """
    return ODSWriter(odsfile, **kwargs)