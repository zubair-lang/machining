import asyncio
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
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s')

# Create and configure console handler
console_handler = logging.StreamHandler()
console_handler.setFormatter(formatter)

# Add console handler to the logger
logger.addHandler(console_handler)


def update_print_table(idx,r,tbl_dt,_tbl,cl=True):
    _tbl_dt[idx] = r
    _tbl.clear_rows()
    _tbl.add_rows(_tbl_dt)
    if cl:
        os.system('cls' if os.name == 'nt' else 'clear')
    print(_tbl)

def print_table(tbl,cl=True):
    if cl:
        os.system('cls' if os.name == 'nt' else 'clear')
    print(tbl)

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
RFID_SCAN_DELAY = 0.05
REJECT_DELAY = 0.450
##########
####### TIMERS ######
_last_com_t = 0
_last_print_t = 0
_last_bcode_t = 0
_last_paste_t = 0
_last_rfid_t = 0

###### HELPER VARS ###
_card_mappings = {}
pattern = re.compile(r'^(01|02)\d{6}$')
table = PrettyTable(['RFID', 'QR', 'Status','Cut','Bundle','Part'])
_stats = {"Total":"-","printed":0,"Failed":0}

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
def reset_all(mlink):
    response = _machine_link.send_and_receive(STATUS_MSG)
    e_val = _machine_link.extract_status_value(response)
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




async def print_label(ready_print,should_print,printed):
    print("Print Label Task started Waiting for signal !")
    while(True):
        should_print.Wait() # Wait For Go Ahead
        ready_print.Wait()  # Wait for Ready Event
        item = await queue.get()
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
async def barcode_read(e_bcode):
    print("Barcode Read Task started Waiting for signal !")
    while(True):
        e_bcode.Wait()  # Wait for Event

async def rfid_read(e_bcode):
    print("RFID Read Task started Waiting for signal !")
    while(True):
        e_bcode.Wait()  # Wait for Event

async def com_mgr(_machine_link):
    print("Com Manager Started !")
    while(True):
        response = _machine_link.send_and_receive(STATUS_MSG)
        e_val = _machine_link.extract_status_value(response)
        
        _last_com_t = time.perf_counter()
        if e_val is None:
            logger.warning(f"Invalid response = {response} Extracted value = {e_val}")
            time.sleep(1)
            continue
        _p,_r,_b = int(e_val[1]),int(e_val[2]),int(e_val[3])
        if _p == 1 and check_bit(_machine_state,machine_states['MACHINE_SHOULD_PRINT']):
            set_bit(_machine_state,machine_states['MACHINE_READY_TO_PRINT'])
            if check_bit(_machine_state,machine_states['MACHINE_SHOULD_PRINT']):
                response = _machine_link.send_and_receive(PRINT_SIG) # Reset Print Signal
                if  ACK_SIG != response:
                    logger.debug(f"Print Reset Signal Failed {response}")
                    _machine_state = clear_bit(_machine_state,machine_states['MACHINE_READY_TO_PRINT']) # Reset Print Sig
                
                if check_bit(_machine_state,machine_states['MACHINE_INIT']):
                    logger.debug("First print !")
                    _machine_state = clear_bit(_machine_state,machine_states['MACHINE_INIT']) # Reset Init State



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
#print_table(table,False)
_stats = {"Total":len(_lbldata),"printed":0,"Failed":0}
logger.info(_stats)
##############



async def main():

    # Create event and tasks
    ready_print = asyncio.Event()
    should_print = asyncio.Event()
    printed =  asyncio.Event()
    barcode_event = asyncio.Event()
    rfid_event = asyncio.Event()
    queue = asyncio.Queue()


    tasks = [print_label(ready_print,should_print,printed,queue), barcode_read(barcode_event), rfid_read(rfid_event),com_mgr(_machine_link,queue)]
    await asyncio.gather(*tasks)

# Run the main coroutine
asyncio.run(main())
