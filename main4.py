import datetime
import logging
import os
import re
import sys
import time
from prettytable import PrettyTable
from datautils.excellreader import ExcelToGenerator
from datautils.db import MachineMSSQLServer
import requests

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
_stats = {"Total": "-", "printed": 0, "Failed": 0}

def update_print_table(idx, r, tbl_dt, _tbl, cl=True):
    _tbl_dt[idx] = r
    _tbl.clear_rows()
    _tbl.add_rows(_tbl_dt)
    if cl:
        os.system('cls' if os.name == 'nt' else 'clear')
    logger.info(_tbl)
    logger.info(_stats)

def print_table(tbl, cl=True):
    if cl:
        os.system('cls' if os.name == 'nt' else 'clear')
    logger.info(tbl)


###### HELPER VARS ###
_card_mappings = {}
pattern = re.compile(r'^(01|02)\d{6}$')
table = PrettyTable(['RFID', 'QR', 'Status', 'Cut', 'Bundle', 'Part'])


# ANSI color codes
RED = "\033[1;31m"
GREEN = "\033[0;32m"
BLUE = "\033[1;34m"
RESET = "\033[0;0m"
#################


############################# MAAIN ##################################

if __name__ == '__main__':
    ##### DB LINK AND LOAD ######
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
        _tbl_dt.append([RED + _lblt['RFID'] + RESET, _lblt['GroupID'], 'loaded', _lblt['CutNo'], _lblt['BundleCode'], _lblt['GarPanelDesc'] + f" - {_lblt['StrtPcs']} - {_lblt['EndPcs']} - " + _lblt['Size']])

    table.clear_rows()
    table.add_rows(_tbl_dt)
    print_table(table, True)
    _stats = {"Total": len(_lbldata), "printed": 0, "Failed": 0}
    logger.info(_stats)
    ##############
    ####### iterate over labels #############
    for idx, _lbl in enumerate(_labels):
        while(True):
            if _lbl['GarPanelDesc'] == "FR" or _lbl['GarPanelDesc'] == "FR-WB":
                pass
            update_print_table(idx, [BLUE + '>>>>>>>>>' + RESET, GREEN + str(_lbl['GroupID']) + RESET, 'QR Read', _lbl['CutNo'], _lbl['BundleCode'], _lbl['GarPanelDesc'] + f" - {_lbl['StrtPcs']} - {_lbl['EndPcs']} - " + _lbl['Size']], _tbl_dt, table, True)
            
            params = {
                "orid": _lbl['OrId'],
                "po": _lbl['ProductionOrderCode'],
                "bundle": _lbl['BundleCode'],
                "cut": _lbl['CutNo'],
                "color": _lbl['Color'],
                "part": _lbl['GarPanelDesc'],
                "qty": _lbl['BundleQuantity'],
                "size": _lbl['Size'],
                "lot": _lbl['Lotno'],
                "range": f"{_lbl['StrtPcs']} - {_lbl['EndPcs']}",
                "flr": _lbl['FLRSrNo'],
                "kit": _lbl['Kit'],
                "speed": 2,
                "usr": f"Cutting Depart {_lbl['BundleID']}"
            }
            url = 'http://localhost:5000/api/CardPrinter/PrintLabel'
            response = requests.get(url, params=params, headers={'accept': 'text/plain'})
            logger.info(response.text)
            
            response_json = response.json()
            if response_json['printerStatus'] == 0:
                logger.info("Printed Successfully")
                _stats['printed'] += 1
                update_print_table(idx, [GREEN + response_json['rfidInfo']['blockDataStr'][:8] + RESET, GREEN + str(_lbl['GroupID']) + RESET, 'SUCCESS', _lbl['CutNo'], _lbl['BundleCode'], _lbl['GarPanelDesc'] + f" - {_lbl['StrtPcs']} - {_lbl['EndPcs']} - " + _lbl['Size']], _tbl_dt, table, True)
                _ex = f"'{_lbl['GarPanelDesc']}__{_lbl['BundleCode']}'"
                if _machinedb.upload_data(response_json['rfidInfo']['blockDataStr'][:8],_lbl['GroupID'],_lbl['BundleID'],_ex):
                    logger.error("Error Uploading to Database! ")
                    sys.exit()
                _machinedb.save_id_to_file('id.txt',_lbl['ID'])
                break
            else:
                logger.error("Printing Failed")
                _stats['Failed'] += 1
                update_print_table(idx, [RED + "Failed" + RESET, RED + str(_lbl['GroupID']) + RESET, 'FAILED', _lbl['CutNo'], _lbl['BundleCode'], _lbl['GarPanelDesc'] + f" - {_lbl['StrtPcs']} - {_lbl['EndPcs']} - " + _lbl['Size']], _tbl_dt, table, True)
                sys.exit()
                
