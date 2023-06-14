import serial

# Send a message and receive the response
PRINT_SIG   = bytes.fromhex('25 30 31 23 57 43 50 34 52 30 31 36 31 30 52 30 31 36 32 30 52 30 31 36 33 30 52 30 31 36 34 30 2a 2a 0d 0a')
STATUS_MSG  = bytes.fromhex('25 30 31 23 52 43 50 34 52 30 31 36 30 52 30 31 36 31 52 30 31 36 35 52 30 31 36 38 2a 2a 0d 0a')
RFID_MSG    = bytes.fromhex('25 30 31 23 57 43 50 33 52 30 31 36 35 30 52 30 31 36 36 30 52 30 31 36 37 30 2a 2a 0d 0a')
OK_PLC      = bytes.fromhex('25 30 31 23 57 43 50 31 52 30 31 36 36 31 2a 2a 0d 0a')
NG_PLC      = bytes.fromhex('25 30 31 23 57 43 50 31 52 30 31 36 37 31 2a 2a 0d 0a')
SCAN_BCODE  = bytes.fromhex('25 30 31 23 57 43 50 31 52 30 31 36 38 30 2a 2a 0d 0a')


ACK_SIG = '%01$WC14'


class MachineCommLink:
    def __init__(self, port='COM2', baudrate=115200, timeout=10,parity="N"):
        self.serial = serial.Serial(port, baudrate, timeout=timeout,parity=parity)
 
    def send_and_receive(self, message):
        self.serial.write(message)
        response = self.serial.read_until(expected=b'\r').decode().strip()
        return response

    def extract_status_value(self, response):
        start_index = response.find("%01$RC") + 6
        if start_index != -1:
            extracted_value = str(response[start_index:start_index + 4])
            #print (f"start ind = {start_index} Extracted value = {extracted_value}")
            return extracted_value
        else:
            return None

