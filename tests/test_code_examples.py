# The MIT License (MIT)
#
# Copyright (c) 2014 - 2015 Michael Mulqueen (http://michael.mulqueen.me.uk/)
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


class TestCodeExamples(TestCase):
    def test_single_sheet(self):
        f = io.BytesIO()
        with ods.writer(f) as odsfile:
            odsfile.writerow(["String", "ABCDEF123456", "123456"])
            # Lose the 2L below if you want to run this example code on Python 3, Python 3 has no long type.
            odsfile.writerow(["Float", 1, 123, 123.123, decimal.Decimal("10.321")])
            odsfile.writerow(["Date/DateTime", datetime.datetime.now(), datetime.date(1989, 11, 9)])
            odsfile.writerow(["Time", datetime.time(13, 37), datetime.time(16, 17, 18)])
            odsfile.writerow(["Bool", True, False, True])
            odsfile.writerow(["Formula", 1, 2, 3, ods.Formula("IF(A1=2,B1,C1)")])
        val = f.getvalue()
        self.assertGreater(len(val), 0)

    def test_multi_sheet(self):
        f = io.BytesIO()
        with ods.writer(f) as odsfile:
            bears = odsfile.new_sheet("Bears")
            bears.writerow(["American Black Bear", "Asiatic Black Bear", "Brown Bear", "Giant Panda", "Qinling Panda",
                             "Sloth Bear", "Sun Bear", "Polar Bear", "Spectacled Bear"])
            sloths = odsfile.new_sheet("Sloths")
            sloths.writerow(["Pygmy Three-Toed Sloth", "Maned Sloth", "Pale-Throated Sloth", "Brown-Throated Sloth",
                             "Linneaeus's Two-Twoed Sloth", "Hoffman's Two-Toed Sloth"])

    def test_col_padding(self):
        f = io.BytesIO()
        with ods.writer(f) as odsfile:
            my_sheet = odsfile.new_sheet("My Sheet", cols=3)
            my_sheet.writerows([["One"],
                                ["Two", "Four", "Sixteen"],
                                ["Three", "Nine", "Twenty seven"]])