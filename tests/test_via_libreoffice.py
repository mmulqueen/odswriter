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

@unittest.skipUnless(command_is_executable(["libreoffice", "--version"]), "LibreOffice not found")
class TestViaLibreOffice(unittest.TestCase):
    def launder_through_lo(self, rows):
        """
            Saves rows into both ods and fods formats, uses LibreOffice to convert to CSV and loads
            the rows from that CSV. Uses subtests to test both formats.
        """
        formats = [
            ("ods", lambda f: ods.writer(f), "wb"),
            ("fods", lambda f: ods.fods_writer(f), "w")
        ]

        results = []

        for format_name, writer_func, file_mode in formats:
            with self.subTest(format=format_name):
                with tempfile.TemporaryDirectory() as temp_dir:
                    # Make the file
                    temp_file = os.path.join(temp_dir, f"test.{format_name}")
                    with open(temp_file, file_mode) as temp_file_handle:
                        with writer_func(temp_file_handle) as odsfile:
                            odsfile.writerows(rows)

                    # Convert it to a CSV
                    subprocess.check_output(["libreoffice", "--headless", "--convert-to",
                                                 "csv", f"test.{format_name}"],
                                                cwd=temp_dir)

                    # Read the CSV
                    temp_csv = os.path.join(temp_dir, "test.csv")
                    with open(temp_csv) as temp_csv_file:
                        csvfile = csv.reader(temp_csv_file)
                        result = list(csvfile)
                        results.append(result)

        # Return the first result for backward compatibility, but both formats are tested
        return results[0] if results else []
    def test_string(self):
        lrows = self.launder_through_lo([["String", "ABCDEF123456", "123456"]])
        self.assertEqual(lrows,[["String", "ABCDEF123456", "123456"]])

    def test_numeric(self):
        lrows = self.launder_through_lo([["Float",
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
        lrows = self.launder_through_lo([["Date/DateTime",
                                     datetime.date(1989,11,9)]])

        # Locales may effect how LibreOffice outputs the dates, so I'll
        # settle for checking for the presence of substrings.

        self.assertEqual(lrows[0][0],"Date/DateTime")
        self.assertIn("1989", lrows[0][1])
        self.assertIn("11", lrows[0][1])
        self.assertIn("09", lrows[0][1])

    def test_time(self):
        lrows = self.launder_through_lo([["Time",
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
        lrows = self.launder_through_lo([["Bool",True,False,True]])

        self.assertEqual(lrows,[["Bool", "TRUE", "FALSE", "TRUE"]])

    def test_formula(self):
        lrows = self.launder_through_lo([["Formula",  # A1
                                     1,  # B1
                                     2,  # C1
                                     3,  # D1
                                     ods.Formula("IF(C1=2;B1;C1)"),  # E1
                                         ods.Formula("SUM(B1:D1)")]])  # F1
        self.assertEqual(lrows, [["Formula","1","2","3","1","6"]])

    def test_nested_formula(self):
        lrows = self.launder_through_lo([["Formula",  # A1
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
        lrows = self.launder_through_lo([["<table:table-cell>",
                                     "</table:table-cell>",
                                     "<br />",
                                     "&",
                                     "&amp;"]])
        self.assertEqual(lrows, [["<table:table-cell>",
                                  "</table:table-cell>",
                                  "<br />",
                                  "&",
                                  "&amp;"]])