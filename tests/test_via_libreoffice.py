
# The MIT License (MIT)
#
# Copyright (c) 2015 Michael Mulqueen (http://michael.mulqueen.me.uk/)
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

import unittest

import tempfile
import os
import subprocess
import csv
import shutil

import decimal
import datetime

import odswriter as ods

class TempDir(object):
    """
        A simple context manager for temporary directories.
    """

    def __init__(self):
        self.path = tempfile.mkdtemp()

    def __enter__(self):
        return self

    def __exit__(self, *args, **kwargs):
        shutil.rmtree(self.path)

def command_is_executable(args):
    try:
        subprocess.call(args)
        return True
    except OSError:
        return False

def launder_through_lo(rows):
    """
        Saves rows into an ods, uses LibreOffice to convert to a CSV and loads
        the rows from that CSV.
    """
    with TempDir() as d:
        # Make an ODS
        temp_ods = os.path.join(d.path, "test.ods")
        with ods.writer(open(temp_ods, "wb")) as odsfile:
            odsfile.writerows(rows)

        # Convert it to a CSV
        p = subprocess.Popen(["libreoffice", "--headless", "--convert-to",
                              "csv", "test.ods"],
                             cwd=d.path)

        p.wait()

        # Read the CSV
        temp_csv = os.path.join(d.path,"test.csv")
        csvfile =  csv.reader(open(temp_csv))
        return list(csvfile)


@unittest.skipUnless(command_is_executable(["libreoffice", "--version"]), "LibreOffice not found")
class TestViaLibreOffice(unittest.TestCase):
    def test_string(self):
        lrows = launder_through_lo([["String", "ABCDEF123456", "123456"]])
        self.assertEqual(lrows,[["String", "ABCDEF123456", "123456"]])

    def test_numeric(self):
        lrows = launder_through_lo([["Float",
                                     1,
                                     123,
                                     123.123,
                                     decimal.Decimal("10.321")]])
        self.assertEqual(lrows,
                         [["Float",
                           "1",
                           "123",
                           "123.123",
                           "10.321"]])
    def test_datetime(self):
        lrows = launder_through_lo([["Date/DateTime",
                                     datetime.date(1989,11,9)]])

        # Locales may effect how LibreOffice outputs the dates, so I'll
        # settle for checking for the presence of substrings.

        self.assertEqual(lrows[0][0],"Date/DateTime")
        self.assertIn("1989", lrows[0][1])
        self.assertIn("11", lrows[0][1])
        self.assertIn("09", lrows[0][1])

    def test_time(self):
        lrows = launder_through_lo([["Time",
                                     datetime.time(13,37),
                                     datetime.time(16,17,18)]])

        # Again locales may be important.

        self.assertEqual(lrows[0][0],"Time")
        self.assertTrue("13" in lrows[0][1] or "01" in lrows[0][1])
        self.assertIn("37",lrows[0][1])
        self.assertNotIn("AM", lrows[0][1])
        self.assertTrue("16" in lrows[0][2] or "04" in lrows[0][2])
        self.assertNotIn("AM", lrows[0][2])
        self.assertIn("17",lrows[0][2])
        self.assertIn("18",lrows[0][2])

    def test_bool(self):
        lrows = launder_through_lo([["Bool",True,False,True]])

        self.assertEqual(lrows,[["Bool", "TRUE", "FALSE", "TRUE"]])

    def test_formula(self):
        lrows = launder_through_lo([["Formula",  # A1
                                     1,  # B1
                                     2,  # C1
                                     3,  # D1
                                     ods.Formula("IF(C1=2;B1;C1)"),  # E1
                                         ods.Formula("SUM(B1:D1)")]])  # F1
        self.assertEqual(lrows, [["Formula","1","2","3","1","6"]])

    def test_nested_formula(self):
        lrows = launder_through_lo([["Formula",  # A1
                                     1,  # B1
                                     2,  # C1
                                     3,  # D1
                                     ods.Formula("IF(C1=2;MIN(B1:D1);MAX(B1:D1))"),  # E1
                                     ods.Formula("IF(C1=3;MIN(B1:D1);MAX(B1:D1))")]])  # F1
        self.assertEqual(lrows, [["Formula","1","2","3","1","3"]])

    def test_escape(self):
        """
            Make sure that special characters are actually being escaped.
        """
        lrows = launder_through_lo([["<table:table-cell>",
                                     "</table:table-cell>",
                                     "<br />",
                                     "&",
                                     "&amp;"]])
        self.assertEqual(lrows, [["<table:table-cell>",
                                  "</table:table-cell>",
                                  "<br />",
                                  "&",
                                  "&amp;"]])