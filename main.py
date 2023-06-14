
import os
import re
import time
from prettytable import PrettyTable
import win32com.client
import serial
from datautils.excellreader import ExcelToGenerator
from datautils.db import MachineMSSQLServer
from commutils.machine import NG_PLC, OK_PLC, PRINT_SIG, RFID_MSG, SCAN_BCODE, STATUS_MSG, MachineCommLink
from rfidutils.rfid_helper import RFIDHelper

LABEL_PATH = "C:\\Users\\PT-PC\\Desktop\\machinelabel.btw"
PRINTER_NAME = "TSC PEX-1231"
RFID_DELAY = 0.2

def print_btw_label(filename, printer_name, data_dict, quantity=1):
    # Create an instance of BarTender Application
    bt_app = win32com.client.Dispatch('BarTender.Application')

    # Open the label template
    bt_format = bt_app.Formats.Open(filename, False, '')

    # Set printer
    bt_format.Printer = printer_name

    # Specify the number of copies
    bt_format.IdenticalCopiesOfLabel = quantity

    # Set field values
    for field, value in data_dict.items():
        bt_format.SetNamedSubStringValue(field, value)

    # Print the label
    bt_format.PrintOut(False, False)

# Call the function




if __name__ == '__main__':

    _machine_link = MachineCommLink(parity="O")
    response = _machine_link.send_and_receive(RFID_MSG)
    time.sleep(.5)
    response = _machine_link.send_and_receive(SCAN_BCODE)
    time.sleep(.5)
    response = _machine_link.send_and_receive(NG_PLC)
    os.system('cls' if os.name == 'nt' else 'clear')
    _barcode_link = serial.Serial("COM1", 9600, timeout=1,parity="N")
    _rfid_link = RFIDHelper(device='COM3')
    #_rfid_link.inventory()

    # file_name = 'a.xlsx'
    # sheet_name = 'Sheet1'
    # excel_gen = ExcelToGenerator(file_name, sheet_name)

    # # Create a generator
    # gen = excel_gen.data_generator()

    _machinedb = MachineMSSQLServer('172.16.20.1', 'ActiveSooperWizerNCL', 'sa', 'wimetrix')
    _lbldata = _machinedb.load_data()


    _cltp = True
    _card_mappings = {}
    pattern = re.compile(r'^(01|02)\d{6}$')
    table = PrettyTable(['RFID', 'QR', 'Status','Cut','Bundle','Part'])
    # ANSI color codes
    RED = "\033[1;31m"
    GREEN = "\033[0;32m"
    BLUE = "\033[1;34m"
    RESET = "\033[0;0m"

    _stats = {"Total":len(_lbldata),"printed":0,"Failed":0}


    
    # Use the generator
    for row in _lbldata:
        _lbl = row._asdict()
        print(f"new label printing {_lbl['GroupID']}")
        _qr = None
        while(True):
            response = _machine_link.send_and_receive(STATUS_MSG)
            e_val = _machine_link.extract_status_value(response)
            if e_val is None:
                print(f"Invalid response = {response} Extracted value = {e_val}")
                time.sleep(1)
                continue
            if int(e_val[1])==1 and _cltp == True:
                print(f"Printing {_lbl['GroupID']}")
                response = _machine_link.send_and_receive(PRINT_SIG)
                params = {"WorkOrder":_lbl['OrId'],"PO":_lbl['ProductionOrderCode'],"Bundle":_lbl['BundleCode'],"CutNo":_lbl['CutNo'],"Color":_lbl['Color'],"Part":_lbl['GarPanelDesc'],"QTY":_lbl['BundleQuantity'],"Size":_lbl['Size'],"LOT":_lbl['Lotno'],"QRCode":_lbl['GroupID']}
                print_btw_label(LABEL_PATH, PRINTER_NAME,params, 1)
                _cltp = False
                #time.sleep(1)
                #continue


            if int(e_val[2]):
                print("RFID Signal is High !")
                response = _machine_link.send_and_receive(RFID_MSG)
                time.sleep(RFID_DELAY)
                if _qr is None or _qr=="no ready":
                    print("QR Error Throwing Out")
                    response = _machine_link.send_and_receive(NG_PLC)
                    _stats["Failed"] += 1
                    _cltp = True
                    continue
                else:
                    _st = _rfid_link.inventory()
                    if _st[0] == 1:
                        print("No Card Trying again")
                        time.sleep(RFID_DELAY)
                        _st = _rfid_link.inventory()
                        if _st[0] == 1:
                            print("No Card Throwing out")
                            response = _machine_link.send_and_receive(NG_PLC)
                            _stats["Failed"] += 1
                            _cltp = True
                            continue
                    if _st[0] == 2:
                        print("Multi Read Throwing Out")
                        response = _machine_link.send_and_receive(NG_PLC)
                        _stats["Failed"] += 1
                        _cltp = True
                        continue

                    if _st[1] in _card_mappings:
                        print(f"Already Initialized Card {_st[1]} Read Again Throwing Out")
                        response = _machine_link.send_and_receive(NG_PLC)
                        _stats["Failed"] += 1
                        _cltp = True
                        continue

                        
                    if not pattern.match(_st[1]):
                        print("Bad EPC !")
                        response = _machine_link.send_and_receive(NG_PLC)
                        _stats["Failed"] += 1
                        _cltp = True
                        continue
                    if not _qr == str(_lbl['GroupID']):
                        print(f"Bad QR ! Read = {_qr} Expected {_lbl['GroupID']}")
                        response = _machine_link.send_and_receive(NG_PLC)
                        _stats["Failed"] += 1
                        _cltp = True
                        continue
                    else:
                        print (f"QR {_qr} RFID = {_st[1]}")
                        _card_mappings[_st[1]] = _qr
                        #params = {"WorkOrder":_lbl['OrId'],"PO":_lbl['ProductionOrderCode'],"Bundle":_lbl['BundleCode'],"CutNo":_lbl['CutNo'],"Color":_lbl['Color'],"Part":_lbl['GarPanelDesc'],"QTY":_lbl['BundleQuantity'],"Size":_lbl['Size'],"LOT":_lbl['Lotno'],"QRCode":_lbl['GroupID']}
                        table.add_row([RED + _st[1] + RESET, _qr, 'Success',_lbl['CutNo'],_lbl['BundleCode'],_lbl['GarPanelDesc']])
                        _stats["printed"] += 1
                        os.system('cls' if os.name == 'nt' else 'clear')     
                        print(table)
                        print(f"Total = {_stats['Total']} Printed = {_stats['printed']} Failed = {_stats['Failed']}")
                        if _machinedb.upload_data(_st[1],_qr):
                            print("Error Uploading to Database! ")
                            response = _machine_link.send_and_receive(NG_PLC)
                        response = _machine_link.send_and_receive(OK_PLC)
                        _machinedb.save_id_to_file('id.txt',_lbl['ID'])
                        _qr = None
                        _st = None
                        _cltp = True
                        break
            if int(e_val[3]):
                response = _machine_link.send_and_receive(SCAN_BCODE)
                _qr = _barcode_link.readline().decode().strip() 
                #print(f"Barcode Read = {_qr}")  
                #
            
            time.sleep(.1)
    # _rfid_transport.close()