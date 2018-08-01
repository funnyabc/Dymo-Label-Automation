import csv
import sys
from win32com.client import Dispatch

dict = {}
labelCom = Dispatch('Dymo.DymoAddIn')
labelText = Dispatch('Dymo.DymoLabels')
isOpen = labelCom.Open('Assy, PN, Desc Generic.label')
selectPrinter = 'DYMO LabelWriter 450'
labelCom.SelectPrinter(selectPrinter)

with open('2018 Commercial Instrument Refurb Tracking Spreadsheet.csv') as refurb_list:
    inst_list = csv.reader(refurb_list, delimiter=',', quotechar='"')
    for instrument in inst_list:
        if len(list(instrument)) != 5 or list(instrument)[0] == "Instrument Type":
            continue
        items = list(instrument)
        text = ['TEXT__1', 'TEXT__2', 'TEXT___1', 'TEXT____1', 'TEXT_1']
        labelText.SetField('TEXT__1', items[0]) # Inst Type
        labelText.SetField('TEXT__2', items[1]) # Manufacturer
        labelText.SetField('TEXT___1', items[2]) # MFG Model
        labelText.SetField('TEXT____1', 'SN: ' + items[3]) # SN
        labelText.SetField('TEXT_1', items[4]) # AT #
        print(items)
        next = input("'n' to stop, 's' to skip")
        if next == 'n':
            sys.exit(0)
        elif next == 's':
            continue
        labelCom.StartPrintJob()
        labelCom.Print(1,False)
        labelCom.EndPrintJob()
        
