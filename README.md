odswriter
=========

A pure-Python module for writing OpenDocument spreadsheets (similar to csv.writer).

Features
-------------
 - Pure python
 - Automatically converts Python types into OpenDocument equivalents
 - Includes support for datetime, date and time types
 - Includes support for Decimal type
 - Compatible with Python 2.7 and 3.3 (and probably others)

License
-----------
The MIT License (MIT), refer to LICENSE.txt

Example
---------
```python
import datetime
import decimal
import odswriter as ods
    
with ods.writer(open("test.ods","wb")) as odsfile:
    odsfile.writerow(["String", "ABCDEF123456", "123456"])
    # Lose the 2L below if you want to run this example code on Python 3, Python 3 has no long type.
    odsfile.writerow(["Float", 1, 123, 123.123, 2L, decimal.Decimal("10.321")])
    odsfile.writerow(["Date/DateTime", datetime.datetime.now(), datetime.date(1989,11,9)])
    odsfile.writerow(["Time",datetime.time(13,37),datetime.time(16,17,18)])
    odsfile.writerow(["Bool",True,False,True])
```
