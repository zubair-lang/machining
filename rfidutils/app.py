import serial
import time
from chafon_rfid.base import CommandRunner,ReaderCommand
from chafon_rfid.transport_serial import SerialTransport
from chafon_rfid.command import CF_GET_READER_INFO, G2_TAG_INVENTORY
from chafon_rfid.uhfreader18 import G2InventoryResponseFrame as G2InventoryResponseFrame18
from chafon_rfid.response import G2_TAG_INVENTORY_STATUS_MORE_FRAMES

class SerialCommunication:
    def __init__(self, port='COM2', baudrate=115200, timeout=10,parity="N"):
        self.serial = serial.Serial(port, baudrate, timeout=timeout,parity=parity)
 
    def send_and_receive(self, message):
        self.serial.write(message)
        response = self.serial.read_until(expected=b'\r').decode().strip()
        return response

    def extract_value(self, response):
        start_index = response.find("%01$RC") + 6
        if start_index != -1:
            extracted_value = str(response[start_index:start_index + 4])
            #print (f"start ind = {start_index} Extracted value = {extracted_value}")
            return extracted_value
        else:
            return None



if __name__ == '__main__':

    get_inventory_uhfreader18 = ReaderCommand(G2_TAG_INVENTORY)
    _rfid_transport = SerialTransport(device='COM3')
    runner = CommandRunner(_rfid_transport)


    _machine_link = SerialCommunication(parity="O")
    _barcode_link = SerialCommunication(port='COM1',baudrate=9600,timeout=1)
    
    # Send a message and receive the response
    print_sig = bytes.fromhex('25 30 31 23 57 43 50 34 52 30 31 36 31 30 52 30 31 36 32 30 52 30 31 36 33 30 52 30 31 36 34 30 2a 2a 0d 0a')
    status_msg = bytes.fromhex('25 30 31 23 52 43 50 34 52 30 31 36 30 52 30 31 36 31 52 30 31 36 35 52 30 31 36 38 2a 2a 0d 0a')
    rfid_msg = bytes.fromhex('25 30 31 23 57 43 50 33 52 30 31 36 35 30 52 30 31 36 36 30 52 30 31 36 37 30 2a 2a 0d 0a')
    ok_plc = bytes.fromhex('25 30 31 23 57 43 50 31 52 30 31 36 36 31 2a 2a 0d 0a')
    ng_plc = bytes.fromhex('25 30 31 23 57 43 50 31 52 30 31 36 37 31 2a 2a 0d 0a')
    scan_bcode = bytes.fromhex('25 30 31 23 57 43 50 31 52 30 31 36 38 30 2a 2a 0d 0a')
    _qr = None
    while(True):
        response = _machine_link.send_and_receive(status_msg)
        e_val = _machine_link.extract_value(response)
        #print(f"response = {response} Extracted value = {e_val}")
        if e_val is None:
            print("Not valid status Response Found !")
        if int(e_val[1]):
            #print("Print Signal is High !")
            response = _machine_link.send_and_receive(print_sig)
        if int(e_val[2]):
            #print("RFID Signal is High !")
            response = _machine_link.send_and_receive(rfid_msg)
            time.sleep(.2)
            if _qr is None or _qr=="no ready":
                print("QR Error Throwing Out")
                response = _machine_link.send_and_receive(ng_plc)
            else:
                _rfid_transport.write(get_inventory_uhfreader18.serialize())
                inventory_status = None
                while inventory_status is None or inventory_status == G2_TAG_INVENTORY_STATUS_MORE_FRAMES:
                    #g2_response = G2InventoryResponseFrame288(transport.read_frame())
                    g2_response = G2InventoryResponseFrame18(_rfid_transport.read_frame())
                    inventory_status = g2_response.result_status
                    if inventory_status == 251:
                        print("No Card Throwing Out")
                        response = _machine_link.send_and_receive(ng_plc)
                        break
                    if inventory_status == G2_TAG_INVENTORY_STATUS_MORE_FRAMES:
                        print("Multi Read Throwing Out")
                        response = _machine_link.send_and_receive(ng_plc)
                        break
                    for tag in g2_response.get_tag():
                        #print('Antenna %d: EPC %s, RSSI %s' % (tag.antenna_num, tag.epc.hex(), tag.rssi))
                        print (f"QR {_qr} RFID = {tag.epc.hex()}")
                        response = _machine_link.send_and_receive(ok_plc)
                        _qr = None
                        break
                #print("Letting Pass")
                
            #print(f"resp {response}")
        if int(e_val[3]):
            response = _machine_link.send_and_receive(scan_bcode)
            _qr = _barcode_link.serial.readline().decode().strip() 
            #print(f"Barcode Read = {_qr}")      

        time.sleep(.2)
    _rfid_transport.close()