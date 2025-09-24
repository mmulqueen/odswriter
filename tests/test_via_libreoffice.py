import unittest

import tempfile
import os
import subprocess
import csv

import decimal
import datetime

import odswriter as ods


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
    with tempfile.TemporaryDirectory() as temp_dir:
        # Make an ODS
        temp_ods = os.path.join(temp_dir, "test.ods")
        with open(temp_ods, "wb") as temp_ods_file:
            with ods.writer(temp_ods_file) as odsfile:
                odsfile.writerows(rows)

        # Convert it to a CSV
        p = subprocess.Popen(["libreoffice", "--headless", "--convert-to",
                              "csv", "test.ods"],
                             cwd=temp_dir)

        p.wait()

        # Read the CSV
        temp_csv = os.path.join(temp_dir, "test.csv")
        with open(temp_csv) as temp_csv_file:
            csvfile = csv.reader(temp_csv_file)
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