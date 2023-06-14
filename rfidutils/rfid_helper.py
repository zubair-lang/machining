from chafon_rfid.base import CommandRunner,ReaderCommand
from chafon_rfid.transport_serial import SerialTransport
from chafon_rfid.command import G2_TAG_INVENTORY
from chafon_rfid.uhfreader18 import G2InventoryResponseFrame as G2InventoryResponseFrame18
from chafon_rfid.response import G2_TAG_INVENTORY_STATUS_MORE_FRAMES

class RFIDHelper:
    def __init__(self, device='COM3'):
        self._rfid_transport = SerialTransport(device=device)
        self._runner = CommandRunner(self._rfid_transport)
        self._get_inventory_uhfreader18 = ReaderCommand(G2_TAG_INVENTORY)

    def inventory(self):
        self._rfid_transport.write(self._get_inventory_uhfreader18.serialize())
        inventory_status = None
        while inventory_status is None or inventory_status == G2_TAG_INVENTORY_STATUS_MORE_FRAMES:
            g2_response = G2InventoryResponseFrame18(self._rfid_transport.read_frame())
            inventory_status = g2_response.result_status
            if inventory_status == 251:
                return [1,{}]
            if inventory_status == G2_TAG_INVENTORY_STATUS_MORE_FRAMES:
                return [2,{}]
            for tag in g2_response.get_tag():
                #print(f'Antenna {tag.antenna_num}: EPC {tag.epc.hex()}, RSSI {tag.rssi}')
                return [0,tag.epc.hex()]
