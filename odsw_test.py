import datetime
import decimal
import odswriter as ods


# Single sheet mode
with ods.writer(open("test.ods","wb"), first_row_bold=True) as odsfile:
    odsfile.writerow(["String", "ABCDEF123456", "123456"])
    odsfile.writerow(["Multiline", "This is one line\nThis is two line\nThis is three line\nThis is a really stupendously long line that is just way too long, isn't it?  Maybe it is we'll never know!"])
    # Lose the 2L below if you want to run this example code on Python 3, Python 3 has no long type.
    odsfile.writerow(["Float", 1, 123, 123.123, decimal.Decimal("10.321")])
    odsfile.writerow(["Date/DateTime", datetime.datetime.now(), datetime.date(1989, 11, 9)])
    odsfile.writerow(["Time",datetime.time(13, 37),datetime.time(16, 17, 18)])
    odsfile.writerow(["Bool", True, False, True])
    odsfile.writerow(["Formula", 1, 2, 3, ods.Formula("IF(A1=2,B1,C1)")])


# Multiple sheet mode
with ods.writer(open("test-multi.ods","wb")) as odsfile:
    bears = odsfile.new_sheet("Bears", first_row_bold=True)
    bears.writerow(["American Black Bear", "Asiatic Black Bear", "Brown Bear", "Giant Panda", "Qinling Panda", 
                     "Sloth Bear", "Sun Bear", "Polar Bear", "Spectacled Bear"])
    sloths = odsfile.new_sheet("Sloths")
    sloths.writerow(["Pygmy Three-Toed Sloth", "Maned Sloth", "Pale-Throated Sloth", "Brown-Throated Sloth",
                     "Linneaeus's Two-Twoed Sloth", "Hoffman's Two-Toed Sloth"])

print ("All done!")