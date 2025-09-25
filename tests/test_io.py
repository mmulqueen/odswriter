from unittest import TestCase

import io
import tempfile
import decimal
import datetime

import odswriter as ods

class TestIO(TestCase):
    def setUp(self):
        self.rows = [
            ["String", "ABCDEF123456", "123456"],
            ["Float", 1, 123, 123.123, decimal.Decimal("10.321")],
            ["Date/DateTime", datetime.datetime.now(), datetime.date(1989,11,9)],
            ["Time",datetime.time(13,37),datetime.time(16,17,18)],
            ["Bool",True,False,True],
            ["Formula",1,2,3,ods.Formula("IF(A1=2,B1,C1)")]
        ]

    def test_BytesIO(self):
        f = io.BytesIO()
        with ods.writer(f) as odsfile:
            odsfile.writerows(self.rows)
        val = f.getvalue()
        self.assertGreater(len(val), 0) # Produces non-empty output

    def test_FileIO(self):
        f = tempfile.TemporaryFile(mode="wb")
        with ods.writer(f) as odsfile:
            odsfile.writerows(self.rows)

    def test_fods_StringIO(self):
        f = io.StringIO()
        with ods.fods_writer(f) as odsfile:
            odsfile.writerows(self.rows)
        val = f.getvalue()
        self.assertGreater(len(val), 0) # Produces non-empty output

    def test_fods_FileIO(self):
        f = tempfile.TemporaryFile(mode="w+")
        with ods.fods_writer(f) as odsfile:
            odsfile.writerows(self.rows)