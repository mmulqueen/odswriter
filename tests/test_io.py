from unittest import TestCase

import io
import tempfile
import decimal
import datetime
import zipfile

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

    def test_structured_and_compressed_correctly(self):
        # ODF spec: mimetype must be the first entry and stored uncompressed.
        # Other entries must use the compression the caller requested.
        for compression in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED):
            f = io.BytesIO()
            with ods.writer(f, compression=compression) as odsfile:
                odsfile.writerows(self.rows)
            with zipfile.ZipFile(io.BytesIO(f.getvalue())) as z:
                self.assertEqual(z.infolist()[0].filename, "mimetype",
                                 msg="mimetype not first (compression={})".format(compression))
                self.assertEqual(z.getinfo("mimetype").compress_type,
                                 zipfile.ZIP_STORED,
                                 msg="mimetype not stored (compression={})".format(compression))
                self.assertEqual(z.getinfo("content.xml").compress_type,
                                 compression,
                                 msg="content.xml not using requested compression={}".format(compression))