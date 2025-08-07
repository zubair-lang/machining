import datetime
import logging
import os
import re
import sys
import time
import threading

# Required for web server functionality
from flask import Flask, request, jsonify

# Required for hardware and label printing
import win32com.client
import serial
import pythoncom  # For CoInitialize
from flask_cors import CORS
# Your custom utility modules
from commutils.machine import ACK_SIG, NG_PLC, OK_PLC, PRINT_SIG, RFID_MSG, SCAN_BCODE, STATUS_MSG, MachineCommLink
from rfidutils.rfid_helper import RFIDHelper

# --- Flask App and State Management ---
app = Flask(__name__)
CORS(app)  # Enables CORS for all routes and origins
_processing_lock = threading.Lock()

# --- Logger Configuration ---
logger = logging.getLogger('machine')
logger.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s')
console_handler = logging.StreamHandler()
console_handler.setFormatter(formatter)
logger.addHandler(console_handler)

# --- Global Handles & State (Hardware Only) ---
_stats = {"Total": 0, "printed": 0, "Failed": 0}
_machine_link = None
_barcode_link = None
_rfid_link = None

# --- Configuration & Constants ---
LABEL_PATH = "C:\\Users\\PT-PC\\Desktop\\rolemachinelabel.btw"
PRINTER_NAME = "TSC PEX-1231"
RFID_SCAN_DELAY = 0.02
REJECT_DELAY_SECONDS = 11
STATUS_SCAN_DELAY = 0.02
pattern = re.compile(r'^(01|02)\d{6}$')
RED = "\033[1;31m"
GREEN = "\033[0;32m"
RESET = "\033[0;0m"


def reset_all(mlink, p=False):
    # This function remains the same
    response = mlink.send_and_receive(STATUS_MSG)
    e_val = mlink.extract_status_value(response)
    if e_val is None:
        return [0, 0, 0, 0]
    if p and int(e_val[1]) == 1: mlink.send_and_receive(PRINT_SIG)
    if int(e_val[2]) == 1:
        mlink.send_and_receive(RFID_MSG)
        time.sleep(0.8)
        mlink.send_and_receive(OK_PLC)
    if int(e_val[3]) == 1: mlink.send_and_receive(SCAN_BCODE)
    time.sleep(0.1)
    response = mlink.send_and_receive(STATUS_MSG)
    return mlink.extract_status_value(response)


@app.route('/print_label', methods=['POST'])
def print_label_sync():
    if not _processing_lock.acquire(blocking=False):
        return jsonify({"status": "ERROR_MACHINE_BUSY", "message": "The machine is currently processing another label."}), 409
    try:
        label_data = request.get_json()
        if not label_data or not isinstance(label_data, dict):
            return jsonify({"status": "ERROR_BAD_REQUEST", "message": "Invalid data format. Expected a single JSON object."}), 400

        required_keys = ['WorkOrder', 'ItemCode']
        if not all(k in label_data for k in required_keys):
            return jsonify({"status": "ERROR_BAD_REQUEST", "message": f"Label data is missing one of required keys: {required_keys}"}), 400

        # Use RollNo as fallback for QRCode if missing
        if 'QRCode' not in label_data or not label_data['QRCode']:
            if 'RollNo' in label_data and label_data['RollNo']:
                label_data['QRCode'] = label_data['RollNo']
                logger.info(f"QRCode not provided. Using RollNo '{label_data['RollNo']}' as fallback.")
            else:
                return jsonify({"status": "ERROR_BAD_REQUEST", "message": "QRCode is missing and RollNo not available for fallback."}), 400

        _stats["Total"] += 1
        logger.info(f"Received job for QRCode: {label_data['QRCode']}. Total jobs: {_stats['Total']}")

        status, message, result_data = process_single_label(label_data)

        if status == 'SUCCESS':
            _stats["printed"] += 1
            http_status_code = 200
        else:
            _stats["Failed"] += 1
            http_status_code = 500

        response = {"status": status, "message": message, **result_data}
        return jsonify(response), http_status_code

    except Exception as e:
        logger.error(f"An unexpected error occurred in the API handler: {e}", exc_info=True)
        return jsonify({"status": "ERROR_UNEXPECTED", "message": str(e)}), 500
    finally:
        logger.info(f"Machine is now idle. Stats: {_stats}")
        _processing_lock.release()



def process_single_label(lbl):
    bt_app = None
    bt_format = None
    try:
        # 1. Initialize this thread for COM
        pythoncom.CoInitialize()
        
        # 2. Start the BarTender application.
        logger.debug("Dispatching new BarTender.Application instance...")
        bt_app = win32com.client.Dispatch('BarTender.Application')
        
        # 3. Check Machine Readiness and open label
        status_val = reset_all(_machine_link)
        if status_val is None or int(status_val[1]) != 1:
            return "FAILED_MACHINE_NOT_READY", f"Machine not ready. PLC Status: {status_val}.", {}

        bt_format = bt_app.Formats.Open(LABEL_PATH, False, '')
        bt_format.Printer = PRINTER_NAME
        bt_format.IdenticalCopiesOfLabel = 1
        
        logger.info(f"Printing label for QRCode: {lbl['QRCode']}...")
        params = {
            "DocNo": lbl.get('DocNo', ''),
            "WorkOrder": lbl.get('WorkOrder', ''), "COLOR": lbl.get('COLOR', ''),
            "LotNo": lbl.get('LotNo', ''), "QRCode": lbl.get('QRCode', ''),
            "Roll": lbl.get('Roll', ''), "RollNo": lbl.get('RollNo', ''),
            "RollLength": lbl.get('RollLength', ''),
            "InvoiceNo": lbl.get('InvoiceNo', ''),
            "Supplier": lbl.get('Supplier', ''),
            "ItemCode": lbl.get('ItemCode', ''),
            "ItemDescription": lbl.get('ItemDescription', ''),
            "TS": lbl.get('TS', datetime.datetime.now().strftime("%d/%m/%Y %H:%M")),
        }
        for field, value in params.items():
            bt_format.SetNamedSubStringValue(field, str(value))

        # 4. Trigger Print
        _machine_link.send_and_receive(PRINT_SIG)
        bt_format.PrintOut(False, False)
        _last_paste_t = time.perf_counter()

        # 5. Barcode Reading
        logger.debug("Waiting for Barcode Trigger...")
        while (time.perf_counter() - _last_paste_t) < REJECT_DELAY_SECONDS:
            status_val = reset_all(_machine_link)
            if status_val is not None and int(status_val[3]) == 1: break
        else:
            return "FAILED_BARCODE_TIMEOUT", "Timeout waiting for barcode trigger.", {}
        
        _machine_link.send_and_receive(SCAN_BCODE)
        _qr = _barcode_link.readline().decode().strip()
        _barcode_link.reset_input_buffer()
        if not _qr or _qr != str(lbl['QRCode']):
            msg = f"Barcode read error. Read: '{_qr}', Expected: '{lbl['QRCode']}'"
            _machine_link.send_and_receive(NG_PLC)
            return "FAILED_BARCODE_MISMATCH", msg, {"expected_qr": str(lbl['QRCode']), "read_qr": _qr}
        
        logger.info(f"Barcode read OK: {_qr}")

        # 6. RFID Reading
        rfid_tag = None
        while (time.perf_counter() - _last_paste_t) < REJECT_DELAY_SECONDS:
            _st = _rfid_link.inventory()
            if _st[0] == 0:
                if pattern.match(_st[1]):
                    rfid_tag = _st[1]
                    break
                else:
                    _machine_link.send_and_receive(NG_PLC)
                    return "FAILED_RFID_BAD_FORMAT", f"Invalid RFID EPC format detected: {_st[1]}", {"read_rfid": _st[1]}
            elif _st[0] == 2:
                _machine_link.send_and_receive(NG_PLC)
                return "FAILED_RFID_MULTI_TAG", f"Multiple RFID tags detected: {_st}", {"tags_detected": _st}
            time.sleep(RFID_SCAN_DELAY)

        if not rfid_tag:
            _machine_link.send_and_receive(NG_PLC)
            return "FAILED_RFID_TIMEOUT", "No RFID tag detected within the rejection window.", {}

        # 7. Success
        _machine_link.send_and_receive(OK_PLC)
        msg = f"Label processed successfully. RFID: {rfid_tag}, QR: {_qr}"
        logger.info(GREEN + msg + RESET)
        return "SUCCESS", msg, {"rfid": rfid_tag, "qr_code": _qr}

    finally:
        # This block ensures everything is cleaned up in the correct order.
        if bt_format is not None:
            bt_format.Close(2)  # 2 = btDoNotSaveChanges
            logger.debug("BarTender format closed.")
        if bt_app is not None:
            # --- vvv THIS IS THE FIX vvv ---
            bt_app.Quit(2)  # 2 = btDoNotSaveChanges
            # --- ^^^ END OF FIX ^^^ ---
            logger.debug("BarTender application quit.")
        
        # Uninitialize this thread from COM
        pythoncom.CoUninitialize()
        logger.debug("Thread uninitialized from COM.")


if __name__ == '__main__':
    try:
        # Initialize hardware connections at startup.
        _machine_link = MachineCommLink(parity="O")
        _barcode_link = serial.Serial("COM1", 9600, timeout=1, parity="N")
        _rfid_link = RFIDHelper(device='COM3')
        
        logger.info("Hardware initialized successfully.")
        logger.info("Starting Flask web server...")
        logger.info(f"Machine is now idle. Awaiting requests at http://0.0.0.0:5000/print_label")
        
    except Exception as e:
        logger.error(f"FATAL: Failed to initialize hardware on startup: {e}", exc_info=True)
        sys.exit(1)

    # Start the Flask web server
    app.run(host='0.0.0.0', port=5000)