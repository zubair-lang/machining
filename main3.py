
import datetime
import logging
import os
import re
import sys
import time
from prettytable import PrettyTable
import win32com.client
import serial
from datautils.excellreader import ExcelToGenerator
from datautils.db import MachineMSSQLServer
from commutils.machine import ACK_SIG, NG_PLC, OK_PLC, PRINT_SIG, RFID_MSG, SCAN_BCODE, STATUS_MSG, MachineCommLink
from rfidutils.rfid_helper import RFIDHelper



# Get logger
logger = logging.getLogger('machine')

# Set up logger
logger.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s')

# Create and configure console handler
console_handler = logging.StreamHandler()
console_handler.setFormatter(formatter)

# Add console handler to the logger
logger.addHandler(console_handler)
_stats = {"Total":"-","printed":0,"Failed":0}

def update_print_table(idx,r,tbl_dt,_tbl,cl=True):
    _tbl_dt[idx] = r
    _tbl.clear_rows()
    _tbl.add_rows(_tbl_dt)
    if cl:
        os.system('cls' if os.name == 'nt' else 'clear')
    logger.info(_tbl)
    logger.info(_stats)
def print_table(tbl,cl=True):
    if cl:
        os.system('cls' if os.name == 'nt' else 'clear')
    logger.info(tbl)

### Labeling Vars
LABEL_PATH = "C:\\Users\\PT-PC\\Desktop\\machinelabel.btw"
PRINTER_NAME = "TSC PEX-1231"
bt_app_g = win32com.client.Dispatch('BarTender.Application')
bt_format_g = bt_app_g.Formats.Open(LABEL_PATH, False, '')
bt_format_g.Printer = PRINTER_NAME
bt_format_g.IdenticalCopiesOfLabel = 1
#####

##### RFID Delay Rejection Vars
RFID_DELAY = 0.8
RFID_SCAN_DELAY = 0.02
REJECT_DELAY = 120
STATUS_SCAN_DELAY = 0.02
NUM_SCAN = 20
WAIT_TIMER = 5.0
##########
####### TIMERS ######
_last_com_t = 0
_last_print_t = 0
_last_bcode_t = 0
_last_paste_t = 0
_last_rfid_t = 0
_qr = None
_rejected = True
_rejected_wait = False

###### HELPER VARS ###
_card_mappings = {}
pattern = re.compile(r'^(01|02)\d{6}$')
table = PrettyTable(['RFID', 'QR', 'Status','Cut','Bundle','Part'])


# ANSI color codes
RED = "\033[1;31m"
GREEN = "\033[0;32m"
BLUE = "\033[1;34m"
RESET = "\033[0;0m"
#################

###### STATE VARS ######
machine_states = {
    "MACHINE_INIT": 1,
    "MACHINE_LABEL_SENT": 2,
    "MACHINE_LABEL_PASTED": 3,
    "MACHINE_BARCODE_TRIGGERED": 4,
    "MACHINE_BARCODE_READ": 5,
    "MACHINE_RFID_READING": 6,
    "MACHINE_REJECTED_CARD": 7,
    "MACHINE_BCODE_TIMEOUT": 8,
    "MACHINE_RFID_TIMEOUT": 9,
    "MACHINE_KEEP_HOLD": 10,
    "MACHINE_READY_TO_PRINT": 11,
    "MACHINE_SHOULD_PRINT": 12,
    "MACHINE_PRINTED_WAITING_PASTING": 13
}

##########

_machine_state = 0

def get_set_bits(machine_state):

    set_bits = []

    for bit_name, bit_number in machine_states.items():
        if machine_state & (1 << (bit_number-1)) != 0:
            set_bits.append(bit_name)

    return set_bits


def set_bit(number, position):
    mask = 1 << (position -1) 
    return number | mask

def clear_bit(number, position):
    mask = ~(1 << (position -1))
    return number & mask

def check_bit(number, position):
    mask = 1 << (position-1)
    return (number & mask) != 0

######## Reset signals
def reset_all(mlink,p=False):
    response = _machine_link.send_and_receive(STATUS_MSG)
    e_val = _machine_link.extract_status_value(response)
    if p == True:
        if int(e_val[1])==1:
            print("Clearing Print !")
            response = _machine_link.send_and_receive(PRINT_SIG) # Reset Print Signal
    if int(e_val[2])==1:
        print("Clearing RFID !")
        response = _machine_link.send_and_receive(RFID_MSG)
        time.sleep(RFID_DELAY)
        response = _machine_link.send_and_receive(OK_PLC)
    if int(e_val[3])==1:
        print("Clearing Barcode !")
        response = _machine_link.send_and_receive(SCAN_BCODE)
    time.sleep(0.1)

    response = _machine_link.send_and_receive(STATUS_MSG)
    return _machine_link.extract_status_value(response)

##################### RFID READ LOOP ################################
def scan_rfid(x,link):
    while(x > 0):
        x = x - 1
        _st = link.inventory()
        
        if _st[0] == 0 or _st[0] == 2:
            return _st
        time.sleep(RFID_SCAN_DELAY)
    return _st


############################# MAAIN ##################################


if __name__ == '__main__':

    ###### COM LINKS TO HARDWARE
    _machine_link = MachineCommLink(parity="O")
    _barcode_link = serial.Serial("COM1", 9600, timeout=1,parity="N")
    _barcode_link.reset_input_buffer()
    _rfid_link = RFIDHelper(device='COM3')
    ##############  

    ######## TAGS CHECK############
    _sres = scan_rfid(10,_rfid_link)
    if _sres[0] == 0 or _sres[0] == 2:
        logger.error(f"Tags in range {_sres}")
        sys.exit()

    ####### MACHINE STATE CHECK #############
    _machine_state = set_bit(_machine_state, machine_states['MACHINE_KEEP_HOLD'])
    logger.debug(BLUE + f"Machine Status {get_set_bits(_machine_state)}" + RESET)
    _status = reset_all(_machine_link)
    while int(_status[2]) == 1 or int(_status[3]) == 1:
        logger.warning(f"{_status} Not reset !")
        time.sleep(0.5)
        _status = reset_all(_machine_link)

    if int(_status[1]) == 1:
        _machine_state = clear_bit(_machine_state,machine_states['MACHINE_KEEP_HOLD'])
        _machine_state = set_bit(_machine_state, machine_states['MACHINE_READY_TO_PRINT'])
        _machine_state = set_bit(_machine_state, machine_states['MACHINE_INIT'])
        _machine_state = set_bit(_machine_state,machine_states['MACHINE_SHOULD_PRINT'])
        logger.debug(GREEN + f"Machine Okay {get_set_bits(_machine_state)} Starting Print Job !\n\n" + RESET)
    else:
        logger.error(RED + f"Machine Not Ready {get_set_bits(_machine_state)}  Status = {_status} Start from HMI and Print Again !\n\n" + RESET)
        sys.exit()
    
    ########################


    ##### DB LINK AND LOAD######
    _machinedb = MachineMSSQLServer('172.16.20.1', 'ActiveSooperWizerNCL', 'sa', 'wimetrix')
    _lbldata = _machinedb.load_data()
    _labels = []
    _tbl_dt = []
    if len(_lbldata) < 1:
        logger.error("Nothing to Print !")
        sys.exit()
    for row in _lbldata:
        _lblt = row._asdict()
        _labels.append(_lblt)
        _tbl_dt.append([RED + _lblt['RFID'] + RESET, _lblt['GroupID'], 'loaded',_lblt['CutNo'],_lblt['BundleCode'],_lblt['GarPanelDesc']])
    
    table.clear_rows()
    table.add_rows(_tbl_dt)
    #print_table(table,True)
    _stats = {"Total":len(_lbldata),"printed":0,"Failed":0}
    logger.info(_stats)
    ##############
    ####### iterate over labels #############
    for idx,_lbl in enumerate(_labels):
        while(True):
            response = _machine_link.send_and_receive(STATUS_MSG)
            e_val = _machine_link.extract_status_value(response)
            
            _last_com_t = time.perf_counter()
            if e_val is None:
                logger.warning(f"Invalid response = {response} Extracted value = {e_val}")
                time.sleep(1)
                continue
            _p,_r,_b = int(e_val[1]),int(e_val[2]),int(e_val[3])

            if _p == 1 and check_bit(_machine_state,machine_states['MACHINE_SHOULD_PRINT']) == True:
                logger.debug(f"Print Condition True! ")
                _machine_state = set_bit(_machine_state,machine_states['MACHINE_READY_TO_PRINT'])
                if check_bit(_machine_state,machine_states['MACHINE_SHOULD_PRINT']):
                    response = _machine_link.send_and_receive(PRINT_SIG) # Reset Print Signal
                    if  ACK_SIG != response:
                        logger.debug(f"Print Reset Signal Failed {response}")
                        _machine_state = clear_bit(_machine_state,machine_states['MACHINE_READY_TO_PRINT']) # Reset Print Sig
                    
                    if check_bit(_machine_state,machine_states['MACHINE_INIT']):
                        logger.debug("First print !")
                        _machine_state = clear_bit(_machine_state,machine_states['MACHINE_INIT']) # Reset Init State
                    _separator = ""
                    if _lbl['GarPanelDesc'] == "FR" or _lbl['GarPanelDesc'] == "FR-WB":
                        _separator = "/\\/\\/\\/\\/\\/\\/\\/\\/\\/\\/\\/\\/\\/\\/\\/\\/\\/\\/\\/\\/\\/\\/\\/\\/\\/"
                        
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
                        "KIT": _lbl['Kit'],
                        "SR": f"{_lbl['StrtPcs']} - {_lbl['EndPcs']}",
                        "Seperator": _separator
                        }
                    #print_btw_label(LABEL_PATH, PRINTER_NAME,params, 1)
                    for field, value in params.items():
                        bt_format_g.SetNamedSubStringValue(field, value)
                    bt_format_g.PrintOut(False, False)
                    _last_print_t = time.perf_counter()
                    update_print_table(idx,[BLUE + '>>>>>>>>>'  + RESET, '__', 'Printing',_lbl['CutNo'],_lbl['BundleCode'],_lbl['GarPanelDesc']],_tbl_dt,table,True)
                    _machine_state = set_bit(_machine_state,machine_states['MACHINE_PRINTED_WAITING_PASTING'])
                    _machine_state = clear_bit(_machine_state,machine_states['MACHINE_SHOULD_PRINT']) # Reset Ready to Print Sig
                    _machine_state = clear_bit(_machine_state,machine_states['MACHINE_READY_TO_PRINT'])

                    ############# Poll Print High Again for Confirmed Pasting #########################
                    while(True):
                        response = _machine_link.send_and_receive(STATUS_MSG)
                        e_val = _machine_link.extract_status_value(response)
                        _p,_r,_b = int(e_val[1]),int(e_val[2]),int(e_val[3])
                        if _p != 1:
                            #logger.debug(get_set_bits(_machine_state))
                            pass
                        else:
                            break
                        time.sleep(STATUS_SCAN_DELAY)
                    logger.debug(f"Pasted ! after {time.perf_counter() - _last_print_t:0.3f}")
                    _machine_state = clear_bit(_machine_state,machine_states['MACHINE_PRINTED_WAITING_PASTING'])
                    _machine_state = set_bit(_machine_state,machine_states['MACHINE_READY_TO_PRINT']) # Reset Ready to Print Sig
                    _machine_state = set_bit(_machine_state,machine_states['MACHINE_LABEL_PASTED']) # Barcode Pasted
                    _last_paste_t = time.perf_counter()
                    _rejected = True
                    logger.debug(get_set_bits(_machine_state))
            
            
            ####################### Barcode Read State ##############################
            
            if check_bit(_machine_state,machine_states['MACHINE_LABEL_PASTED']) == True:
                logger.debug(f"Barcode Condition True!")
                _timeout = False
                logger.debug(f"Waiting for Barcode Trigger !")
                while(True):
                    response = _machine_link.send_and_receive(STATUS_MSG)
                    e_val = _machine_link.extract_status_value(response)
                    _p,_r,_b = int(e_val[1]),int(e_val[2]),int(e_val[3])
                    if _b != 1:
                        logger.debug(get_set_bits(_machine_state))
                    else:
                        break
                    time.sleep(STATUS_SCAN_DELAY)

                    if int((time.perf_counter() - _last_paste_t)*10) > 11:
                        logger.warning(f"Read Barcode Triger Failed! Remove Stuck card and Restart {time.perf_counter() - _last_paste_t:0.3f}")
                        _timeout = True
                        break 
                if _timeout:
                    logger.error(f"Timeout not trigger in {time.perf_counter() - _last_paste_t:0.3f} seconds !")
                    logger.info(RED + f"Stop Machine Remove stuck or missprinted card and press 1 to resume or 0 to start Again !" + RESET)
                    _decision = int(input("Press 1 to continue and 0 to exit ")) or 0
                    if _decision == 1:
                        _rejected = False
                        reset_all(_machine_link,True)
                        _machine_state = set_bit(_machine_state,machine_states['MACHINE_SHOULD_PRINT']) # Reset Ready to Print Sig
                        _machine_state = clear_bit(_machine_state,machine_states['MACHINE_LABEL_PASTED']) # Barcode Pasted
                        continue
                    else:
                        sys.exit()
                logger.debug(f"Read Barcode Triggered ! after {time.perf_counter() - _last_paste_t:0.3f}")
                _last_bcode_t = time.perf_counter()
                response = _machine_link.send_and_receive(SCAN_BCODE)
                if  ACK_SIG != response:
                    logger.warn("Reset Barcode Sig Failed !")
                
                _qr = _barcode_link.readline().decode().strip()
                
                ######## Barcode Read Error Occured
                if _qr is None or _qr=="no ready" or _qr != str(_lbl['GroupID']):
                    _barcode_link.reset_input_buffer()
                    _stats['Failed'] = _stats['Failed'] + 1
                    logger.error(f"Barcode value read Error {_qr}")
                    logger.info(RED + f"Stop Machine Remove stuck or missprinted card and press 1 to resume or 0 to start Again !" + RESET)
                    _decision = int(input("Press 1 to continue and 0 to exit ")) or 0
                    if _decision == 1:
                        _rejected = False
                        reset_all(_machine_link,True)
                        _machine_state = set_bit(_machine_state,machine_states['MACHINE_SHOULD_PRINT']) # Reset Ready to Print Sig
                        _machine_state = clear_bit(_machine_state,machine_states['MACHINE_LABEL_PASTED']) # Barcode Pasted
                        continue
                    else:
                        sys.exit()
                ################################
                else:
                    update_print_table(idx,[BLUE + '>>>>>>>>>'  + RESET,GREEN +  _qr + RESET, 'QR Read',_lbl['CutNo'],_lbl['BundleCode'],_lbl['GarPanelDesc']],_tbl_dt,table,True)
                    _machine_state = set_bit(_machine_state,machine_states['MACHINE_BARCODE_READ']) # Barcode Pasted
                    _machine_state = clear_bit(_machine_state,machine_states['MACHINE_LABEL_PASTED']) # Barcode Pasted
                    _barcode_link.reset_input_buffer()
                    logger.debug(get_set_bits(_machine_state))

            if check_bit(_machine_state,machine_states['MACHINE_BARCODE_READ']) == True:
                logger.debug(f"RFID Condition True!")
                _timeout = False
                logger.debug(f"Waiting for RFID Trigger !")
                _poll_count = 0
                # while(True):
                #     response = _machine_link.send_and_receive(STATUS_MSG)
                #     e_val = _machine_link.extract_status_value(response)
                #     _p,_r,_b = int(e_val[1]),int(e_val[2]),int(e_val[3])
                #     if _r != 1:
                #         logger.debug(get_set_bits(_machine_state))
                #     else:
                #         break
                #     _poll_count = _poll_count + 1
                #     if _poll_count > 10:
                #         break
                #     time.sleep(STATUS_SCAN_DELAY)
                #     response = _machine_link.send_and_receive(RFID_MSG)
                #     if  ACK_SIG != response:
                #         logger.warn("Reset RFID Sig Failed !")
                    
                #     if int((time.perf_counter() - _last_paste_t)*100) > REJECT_DELAY:
                #         logger.error(f"Card Rejected Time Passed Rejecting {time.perf_counter() - _last_paste_t:0.3f}")
                #         response = _machine_link.send_and_receive(NG_PLC)
                #         _rejected = False
                #         _machine_state = clear_bit(_machine_state,machine_states['MACHINE_BARCODE_READ'])
                #         _machine_state = set_bit(_machine_state,machine_states['MACHINE_SHOULD_PRINT']) # Reset Ready to Print Sig
                #         break
                    
                # logger.debug(f"Read RFID Triggered ! after {time.perf_counter() - _last_bcode_t:0.3f}")
                # response = _machine_link.send_and_receive(RFID_MSG)
                # if  ACK_SIG != response:
                #     logger.warn("Reset RFID Sig Failed !")
                
                while(True):
                    _st = _rfid_link.inventory()
                    if _st[0] == 0 or _st[0] == 2:
                        break
                    if int((time.perf_counter() - _last_paste_t)*100) > REJECT_DELAY:
                        logger.error(f"Card Rejected Time Passed Rejecting {time.perf_counter() - _last_paste_t:0.3f}")
                        response = _machine_link.send_and_receive(NG_PLC)
                        _rejected = False
                        _stats['Failed'] = _stats['Failed'] + 1
                        _machine_state = clear_bit(_machine_state,machine_states['MACHINE_BARCODE_READ'])
                        _machine_state = set_bit(_machine_state,machine_states['MACHINE_SHOULD_PRINT']) # Reset Ready to Print Sig
                        break
                    time.sleep(RFID_SCAN_DELAY)
                _last_rfid_t = time.perf_counter()
                if _st[0] == 2:
                    logger.error(f"Card Rejected Multi Read {time.perf_counter() - _last_paste_t:0.3f}")
                    _machine_state = clear_bit(_machine_state,machine_states['MACHINE_BARCODE_READ'])     

                if _st is None or _st[0] == 1:
                    update_print_table(idx,[RED + 'XXXXXXXXX'  + RESET,GREEN +  _qr + RESET, 'No Detection',_lbl['CutNo'],_lbl['BundleCode'],_lbl['GarPanelDesc']],_tbl_dt,table,True)
                    logger.error(f"No RFID Detected !")
                    _machine_state = clear_bit(_machine_state,machine_states['MACHINE_BARCODE_READ']) 
                elif not pattern.match(_st[1]):
                    update_print_table(idx,[RED + _st[1]  + RESET,GREEN +  _qr + RESET, 'Bad EPC',_lbl['CutNo'],_lbl['BundleCode'],_lbl['GarPanelDesc']],_tbl_dt,table,True)
                    logger.error(f"BAD EPC Detected !")
                    _machine_state = clear_bit(_machine_state,machine_states['MACHINE_BARCODE_READ']) 
                    logger.debug(get_set_bits(_machine_state))

                elif _st[1] in _card_mappings:
                    _rejected_wait = True
                    update_print_table(idx,[RED + _st[1]  + RESET,GREEN +  _qr + RESET, 'Same Card Again',_lbl['CutNo'],_lbl['BundleCode'],_lbl['GarPanelDesc']],_tbl_dt,table,True)
                    logger.error(f"Previous card read Check Nearby !")
                    _machine_state = clear_bit(_machine_state,machine_states['MACHINE_BARCODE_READ']) 
                    logger.debug(get_set_bits(_machine_state))
                else:
                    response = _machine_link.send_and_receive(OK_PLC) 
                    _rejected = False
                    _machine_state = set_bit(_machine_state,machine_states['MACHINE_SHOULD_PRINT']) # Reset Ready to Print Sig
                    _machine_state = clear_bit(_machine_state,machine_states['MACHINE_BARCODE_READ'])
                    _ex = f"'{_lbl['GarPanelDesc']}__{_lbl['BundleCode']}'"
                    if _machinedb.upload_data(_st[1],_qr,_lbl['BundleID'],_ex):
                            logger.error("Error Uploading to Database! ")
                            _decision = int(input("Press 1 to continue and 0 to exit ")) or 0
                            if _decision == 1:
                                reset_all(_machine_link)
                                continue
                            else:
                                sys.exit()
                    _card_mappings[_st[1]] = _qr
                    _machinedb.save_id_to_file('id.txt',_lbl['ID'])
                    logger.debug(f"Passed {_st[1]}")
                    _stats['printed'] = _stats['printed'] + 1
                    update_print_table(idx,[GREEN + _st[1]  + RESET,GREEN +  _qr + RESET, 'SUCCESS',_lbl['CutNo'],_lbl['BundleCode'],_lbl['GarPanelDesc']],_tbl_dt,table,True)
                    break

            if _rejected and int((time.perf_counter() - _last_paste_t)* 100 > REJECT_DELAY):
                    _stats['Failed'] = _stats['Failed'] + 1
                    _machine_state = set_bit(_machine_state,machine_states['MACHINE_SHOULD_PRINT']) # Reset Ready to Print Sig
                    logger.error(f"Reject Delay")
                    _rejected = False
                    response = _machine_link.send_and_receive(NG_PLC)
                    logger.debug(f"Rjecting Card !")     
                    if _rejected_wait:
                        logger.info(f"Waiting {WAIT_TIMER}")
                        time.sleep(WAIT_TIMER)
                        _rejected_wait = False 




    