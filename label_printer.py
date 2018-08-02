import csv
import sys
from pprint import pprint
#from win32com.client import Dispatch

dict = {}
labelCom = Dispatch('Dymo.DymoAddIn')
labelText = Dispatch('Dymo.DymoLabels')
isOpen = labelCom.Open('Assy, PN, Desc Generic.label')
selectPrinter = 'DYMO LabelWriter 450'
labelCom.SelectPrinter(selectPrinter)

with open('2018 Commercial Instrument Refurb Tracking Spreadsheet.csv') as refurb_list:
    inst_list = csv.reader(refurb_list, delimiter=',', quotechar='"')
    for instrument in inst_list:
        if list(instrument)[0] == "Instrument Type":
            continue
        items = list(instrument)
        values = {'Instrument Type': items[0], 'Manufacturer': items[1], 'Mfg Model': items[2], \
               'Serial Number': items[3], 'Asset Tracking Number': items[4]}
        text = ['TEXT__1', 'TEXT__2', 'TEXT___1', 'TEXT____1', 'TEXT_1']
        labelText.SetField('TEXT__1', values['Instrument Type']) # Inst Type
        labelText.SetField('TEXT__2', values['Manufacturer']) # Manufacturer
        labelText.SetField('TEXT___1', values['Mfg Model']) # MFG Model
        labelText.SetField('TEXT____1', 'SN: ' + values['Serial Number']) # SN
        labelText.SetField('TEXT_1', values['Asset Tracking Number']) # AT #
        pprint(values)
        next = input("'n' to stop, 's' to skip")
        if next == 'n':
            sys.exit(0)
        elif next == 's':
            continue
        labelCom.StartPrintJob()
        labelCom.Print(1,False)
        labelCom.EndPrintJob()
