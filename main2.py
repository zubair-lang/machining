
import datetime
import os
import re
import sys
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
RFID_DELAY = 0.8
RFID_SCAN_DELAY = 0.05
REJECT_DELAY = 0.6

bt_app_g = win32com.client.Dispatch('BarTender.Application')
bt_format_g = bt_app_g.Formats.Open(LABEL_PATH, False, '')
bt_format_g.Printer = PRINTER_NAME
bt_format_g.IdenticalCopiesOfLabel = 1



# def print_btw_label(filename, printer_name, data_dict, quantity=1):
#     # Create an instance of BarTender Application
#     bt_app = win32com.client.Dispatch('BarTender.Application')

#     # Open the label template
#     bt_format = bt_app.Formats.Open(filename, False, '')

#     # Set printer
#     bt_format.Printer = printer_name

#     # Specify the number of copies
#     bt_format.IdenticalCopiesOfLabel = quantity

#     # Set field values
#     for field, value in data_dict.items():
#         bt_format.SetNamedSubStringValue(field, value)

#     # Print the label
#     bt_format.PrintOut(False, False)

# Call the function

def reset_all(mlink):
    response = _machine_link.send_and_receive(STATUS_MSG)
    e_val = _machine_link.extract_status_value(response)
    # if int(e_val[1])==1:
    #     print("Clearing Print !")
    #     response = _machine_link.send_and_receive(PRINT_SIG)
    if int(e_val[2])==1:
        print("Clearing RFID !")
        response = _machine_link.send_and_receive(RFID_MSG)
        time.sleep(RFID_DELAY)
        response = _machine_link.send_and_receive(OK_PLC)
    if int(e_val[3])==1:
        print("Clearing Barcode !")
        response = _machine_link.send_and_receive(SCAN_BCODE)



if __name__ == '__main__':

    _machine_link = MachineCommLink(parity="O")
    reset_all(_machine_link)

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
    _clt_rfid = False
    _card_mappings = {}
    pattern = re.compile(r'^(01|02)\d{6}$')
    table = PrettyTable(['RFID', 'QR', 'Status','Cut','Bundle','Part'])
    # ANSI color codes
    RED = "\033[1;31m"
    GREEN = "\033[0;32m"
    BLUE = "\033[1;34m"
    RESET = "\033[0;0m"

    _stats = {"Total":len(_lbldata),"printed":0,"Failed":0}

    x= 10
    while(x > 0):
        x = x - 1
        _st = _rfid_link.inventory()
        
        if _st[0] == 0 or _st[0] == 2:
            print(RED + f"Tags in Range ! {_st} Please clear surrounding"+ RESET)
            sys.exit()
        time.sleep(RFID_SCAN_DELAY*1.5)



    _labels = []
    _tbl_dt = []
    # Use the generator
    for row in _lbldata:
        _lblt = row._asdict()
        _labels.append(_lblt)
        _tbl_dt.append([RED + _lblt['RFID'] + RESET, _lblt['GroupID'], 'loaded',_lblt['CutNo'],_lblt['BundleCode'],_lblt['GarPanelDesc']])
    
    table.clear_rows()
    table.add_rows(_tbl_dt)
    print(table)

    for idx,_lbl in enumerate(_labels):
        
        _tbl_dt[idx] = [BLUE + '>>>>>>>>>'  + RESET, '__', 'Printing',_lbl['CutNo'],_lbl['BundleCode'],_lbl['GarPanelDesc']]
        table.clear_rows()
        table.add_rows(_tbl_dt)
        os.system('cls' if os.name == 'nt' else 'clear')
        print(table)
        
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
                params = {"TS":datetime.datetime.now().strftime("%d/%m %H:%M"),
                          "WorkOrder":_lbl['OrId'],
                          "PO":_lbl['ProductionOrderCode'],
                          "Bundle":_lbl['BundleCode'],
                          "CutNo":_lbl['CutNo'],
                          "Color":_lbl['Color'],
                          "Part":_lbl['GarPanelDesc'],
                          "QTY":_lbl['BundleQuantity'],
                          "Size":_lbl['Size'],
                          "LOT":_lbl['Lotno'],
                          "QRCode":_lbl['GroupID'],
                          "FLR":_lbl['FLRSrNo'],
                          "KIT": "__",
                          "SR": f"{_lbl['StrtPcs']} - {_lbl['EndPcs']}"
                          }
                #print_btw_label(LABEL_PATH, PRINTER_NAME,params, 1)
                for field, value in params.items():
                    bt_format_g.SetNamedSubStringValue(field, value)
                bt_format_g.PrintOut(False, False)
                _cltp = False
                #time.sleep(1)
                continue

            if int(e_val[3]):
                response = _machine_link.send_and_receive(SCAN_BCODE)
                _qr = _barcode_link.readline().decode().strip() 
                

            if int(e_val[2]):
                
                print("RFID Signal is High !")
                response = _machine_link.send_and_receive(RFID_MSG)
                x = 10
                _st = None
                if _qr is None or _qr=="no ready":
                    _tbl_dt[idx] = [BLUE + '>>>>>>>>>'  + RESET, '__', 'QR Read Error',_lbl['CutNo'],_lbl['BundleCode'],_lbl['GarPanelDesc']]
                    table.clear_rows()
                    table.add_rows(_tbl_dt)
                    os.system('cls' if os.name == 'nt' else 'clear')
                    print(table)
                    
                    print(f"QR Error Throwing Out eval {e_val}")

                    time.sleep(REJECT_DELAY)
                    response = _machine_link.send_and_receive(NG_PLC)
                    _stats["Failed"] += 1
                    _cltp = True
                    print(f"Total = {_stats['Total']} Printed = {_stats['printed']} Failed = {_stats['Failed']}")
                    _decision = int(input("Press 1 to continue and 0 to exit ")) or 0
                    if _decision == 1:
                        reset_all(_machine_link)
                        continue
                    else:
                        sys.exit()

                else:
                    #time.sleep(RFID_DELAY)
                    while(x > 0):
                        x = x - 1
                        _st = _rfid_link.inventory()
                        
                        if _st[0] == 0:
                            break;
                        time.sleep(RFID_SCAN_DELAY*1)
                    # time.sleep(RFID_DELAY)
                
                
                    # _st = _rfid_link.inventory()
                    
                    if _st[0] == 1:
                        _tbl_dt[idx] = [BLUE + '>>>>>>>>>'  + RESET, '__', 'No Card Detected',_lbl['CutNo'],_lbl['BundleCode'],_lbl['GarPanelDesc']]
                        table.clear_rows()
                        table.add_rows(_tbl_dt)
                        os.system('cls' if os.name == 'nt' else 'clear')
                        print(table)
                        
                        print("No Card Throwing out")
                        time.sleep(REJECT_DELAY)
                        response = _machine_link.send_and_receive(NG_PLC)
                        _stats["Failed"] += 1
                        _cltp = True
                        print(f"Total = {_stats['Total']} Printed = {_stats['printed']} Failed = {_stats['Failed']}")
                        continue
                    if _st[0] == 2:
                        _tbl_dt[idx] = [BLUE + '>>>>>>>>>'  + RESET, '__', 'Multi Card',_lbl['CutNo'],_lbl['BundleCode'],_lbl['GarPanelDesc']]
                        table.clear_rows()
                        table.add_rows(_tbl_dt)
                        os.system('cls' if os.name == 'nt' else 'clear')
                        print(table)

                        print("Multi Read Throwing Out")
                        time.sleep(REJECT_DELAY)
                        response = _machine_link.send_and_receive(NG_PLC)
                        _stats["Failed"] += 1
                        _cltp = True
                        print(f"Total = {_stats['Total']} Printed = {_stats['printed']} Failed = {_stats['Failed']}")
                        _decision = int(input("Press 1 to continue and 0 to exit ")) or 0
                        if _decision == 1:
                            reset_all(_machine_link)
                            continue
                        else:
                            sys.exit()


                    if _st[1] in _card_mappings:
                        _tbl_dt[idx] = [BLUE + '>>>>>>>>>'  + RESET, '__', 'Same Card Read Check Nearby',_lbl['CutNo'],_lbl['BundleCode'],_lbl['GarPanelDesc']]
                        table.clear_rows()
                        table.add_rows(_tbl_dt)
                        os.system('cls' if os.name == 'nt' else 'clear')
                        print(table)
                        
                        print(f"Already Initialized Card {_st[1]} Read Again Throwing Out")
                        time.sleep(REJECT_DELAY)
                        response = _machine_link.send_and_receive(NG_PLC)
                        _stats["Failed"] += 1
                        _cltp = True
                        _decision = int(input("Press 1 to continue and 0 to exit ")) or 0
                        print(f"Total = {_stats['Total']} Printed = {_stats['printed']} Failed = {_stats['Failed']}")
                        if _decision == 1:
                            reset_all(_machine_link)
                            continue
                        else:
                            sys.exit()

                        
                    if not pattern.match(_st[1]):
                        _tbl_dt[idx] = [BLUE + '>>>>>>>>>'  + RESET, '__', 'Wrong/Bad EPC',_lbl['CutNo'],_lbl['BundleCode'],_lbl['GarPanelDesc']]
                        table.clear_rows()
                        table.add_rows(_tbl_dt)
                        os.system('cls' if os.name == 'nt' else 'clear')
                        print(table)
                        
                        print("Bad EPC !")

                        time.sleep(REJECT_DELAY)
                        response = _machine_link.send_and_receive(NG_PLC)
                        _stats["Failed"] += 1
                        _cltp = True
                        print(f"Total = {_stats['Total']} Printed = {_stats['printed']} Failed = {_stats['Failed']}")
                        continue
                    if not _qr == str(_lbl['GroupID']):
                        _tbl_dt[idx] = [BLUE + '>>>>>>>>>'  + RESET, '__', 'Wrong QR Read',_lbl['CutNo'],_lbl['BundleCode'],_lbl['GarPanelDesc']]
                        table.clear_rows()
                        table.add_rows(_tbl_dt)
                        os.system('cls' if os.name == 'nt' else 'clear')
                        print(table)
                        print(f"Bad QR ! Read = {_qr} Expected {_lbl['GroupID']}")
                        time.sleep(REJECT_DELAY)
                        response = _machine_link.send_and_receive(NG_PLC)
                        _stats["Failed"] += 1
                        _cltp = True
                        _qr = None
                        print(f"Total = {_stats['Total']} Printed = {_stats['printed']} Failed = {_stats['Failed']}")
                        print("Restart Print Job !")
                        sys.exit()
                    else:
                        print (f"QR {_qr} RFID = {_st[1]}")
                        _card_mappings[_st[1]] = _qr
                        #params = {"WorkOrder":_lbl['OrId'],"PO":_lbl['ProductionOrderCode'],"Bundle":_lbl['BundleCode'],"CutNo":_lbl['CutNo'],"Color":_lbl['Color'],"Part":_lbl['GarPanelDesc'],"QTY":_lbl['BundleQuantity'],"Size":_lbl['Size'],"LOT":_lbl['Lotno'],"QRCode":_lbl['GroupID']}
                        #table.add_row([RED + _st[1] + RESET, _qr, 'Success',_lbl['CutNo'],_lbl['BundleCode'],_lbl['GarPanelDesc']])
                        _tbl_dt[idx] = [GREEN + _st[1] + RESET, _qr,GREEN+ 'Success' +RESET,_lbl['CutNo'],_lbl['BundleCode'],_lbl['GarPanelDesc']]
                        table.clear_rows()
                        table.add_rows(_tbl_dt)

                        _stats["printed"] += 1
                        os.system('cls' if os.name == 'nt' else 'clear')     
                        print(table)
                        print(f"Total = {_stats['Total']} Printed = {_stats['printed']} Failed = {_stats['Failed']}")
                        if _machinedb.upload_data(_st[1],_qr):
                            print("Error Uploading to Database! ")
                            response = _machine_link.send_and_receive(NG_PLC)
                            _decision = int(input("Press 1 to continue and 0 to exit ")) or 0
                        
                            if _decision == 1:
                                reset_all(_machine_link)
                                continue
                            else:
                                sys.exit()
                        response = _machine_link.send_and_receive(OK_PLC)
                        _machinedb.save_id_to_file('id.txt',_lbl['ID'])
                        _qr = None
                        _st = None
                        _cltp = True
                        break
            
            
            time.sleep(.1)
    # _rfid_transport.close()