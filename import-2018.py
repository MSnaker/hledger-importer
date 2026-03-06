from openpyxl import load_workbook
from datetime import datetime

wb = load_workbook(filename="INPUT/Balance_2018.xlsx")

def len_str(*args) -> int:
    totlen = 0
    for arg in args:
        if type(arg) != str:
            raise TypeError
        foo = len(arg)
        totlen += foo

    return totlen     

sheetimport = wb['IMPORTBATCH']
reason = "2018 EXCEL IMPORT"
source = "Marco:assets:BancoBPM"


with open("./accounting/.hledger.journal", "a", encoding="utf-8") as fw:
    for row in sheetimport.rows:
        date = datetime.date(row[0].value)
        destination = row[1].value
        amount = row[2].value
        line1 = f"\n{date} {reason}\n"

        len2 = len_str(destination, str(amount)) + 8
        len3 = len_str(source, str(amount)) + 9
        padding = max(len2,len3)
        
        line2 = f"\t{destination:<{padding}}{amount}\n"
        line3 = f"\t{source:<{padding-1}}-{amount}\n"
        string_to_append = line1 + line2 + line3
        # print(string_to_append)
        fw.write(string_to_append)