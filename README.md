odswriter
=========

This library receives minimal maintenance, but it is luckily very simple.

A pure-Python module for writing OpenDocument spreadsheets (similar to csv.writer).

Features
-------------
 - Pure python
 - Automatically converts Python types into OpenDocument equivalents
 - Includes support for datetime, date and time types
 - Includes support for Decimal type
 - Tested on Python 3.7, 3.8, 3.9, 3.10, 3.11, 3.12, 3.13
 - Probably still mostly compatible with Python 2.7 and 3.3 to 3.6, but no longer tested.
 - Support for writing formulae (but not evaluating their results)

License
-----------
The MIT License (MIT), refer to LICENSE.txt

Example
---------
```python
import datetime
import decimal
import odswriter as ods


# Single sheet mode
with open("test.ods", "wb") as f:
    with ods.writer(f) as odsfile:
        odsfile.writerow(["String", "ABCDEF123456", "123456"])
        # Lose the 2L below if you want to run this example code on Python 3, Python 3 has no long type.
        odsfile.writerow(["Float", 1, 123, 123.123, decimal.Decimal("10.321")])
        odsfile.writerow(["Date/DateTime", datetime.datetime.now(), datetime.date(1989, 11, 9)])
        odsfile.writerow(["Time",datetime.time(13, 37),datetime.time(16, 17, 18)])
        odsfile.writerow(["Bool", True, False, True])
        odsfile.writerow(["Formula", 1, 2, 3, ods.Formula("IF(A1=2,B1,C1)")])

# Multiple sheet mode
with open("test-multi.ods", "wb") as f:
    with ods.writer(f) as odsfile:
        bears = odsfile.new_sheet("Bears")
        bears.writerow(["American Black Bear", "Asiatic Black Bear", "Brown Bear", "Giant Panda", "Qinling Panda",
                         "Sloth Bear", "Sun Bear", "Polar Bear", "Spectacled Bear"])
        sloths = odsfile.new_sheet("Sloths")
        sloths.writerow(["Pygmy Three-Toed Sloth", "Maned Sloth", "Pale-Throated Sloth", "Brown-Throated Sloth",
                         "Linneaeus's Two-Twoed Sloth", "Hoffman's Two-Toed Sloth"])
```

Compatibility
-------------
Odswriter is tested for compatibility with LibreOffice and Gnumeric. 

jOpenDocument is not compatible out-of-the-box, but by specifying the number of columns (odswriter will pad with empty
cells up to that number) it can be made compatible. Code example:

```python
import odswriter as ods

with open("test-multi.ods", "wb") as f:
    with ods.writer(f) as odsfile:
        my_sheet = odsfile.new_sheet("My Sheet", cols=3)
        my_sheet.writerows([["One"],
                            ["Two", "Four", "Sixteen"],
                            ["Three", "Nine", "Twenty seven"]])

```

Testing
-------
LibreOffice and Gnumeric are used to verify the output of odswriter if installed.

```bash
poetry run python3 -m unittest discover -s tests
```