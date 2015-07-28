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