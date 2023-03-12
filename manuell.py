from datetime import datetime
import os
from datetime import timedelta
import openpyxl
import methoden
from time import  sleep
print("Das einzugebende format ist mit 0 ")
dan = input("Tag anfang")
man= input ("Monat anfang")
jan= input("Jahr anfang")

den = input("Tag Ende ")
men= input ("Monat Ende")
jen= input("Jahr Ende")
anfang = datetime.strptime(f"{dan}-{man}-{jan}",'%d-%m-%Y')
ende= datetime.strptime(f"{den}-{men}-{jen}",'%d-%m-%Y')
akt = anfang
while akt<= ende:
    if os.path.exists(f"Z:/Kassenblatt/{akt.year}/{akt.month}/{akt.day}.xlsx"):
        wb=openpyxl.load_workbook(f"Z:/Kassenblatt/{akt.year}/{akt.month}/{akt.day}.xlsx")
        sheet= wb.active
        sheet=methoden.bar(sheet,akt)
        wb.save(f"Z:/Kassenblatt/{akt.year}/{akt.month}/{akt.day}.xlsx")
        sleep(1)
    else:
        methoden.haupt(akt)
    os.startfile(f"Z:/Kassenblatt/{akt.year}/{akt.month}/{akt.day}.xlsx","print")

    akt=akt+timedelta(1)



