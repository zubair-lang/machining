import win32com.client

def print_btw_label(filename, printer_name, quantity=1):
    # Create an instance of BarTender Application
    bt_app = win32com.client.Dispatch('BarTender.Application')

    # Open the label template
    bt_format = bt_app.Formats.Open(filename, False, '')

    # Set printer
    bt_format.Printer = printer_name

    # Specify the number of copies
    bt_format.IdenticalCopiesOfLabel = quantity

    # Print the label
    bt_format.PrintOut(False, False)

    # Close BarTender
    bt_format.Close(1)  # btSaveOptions.btDoNotSaveChanges: Discards any changes that have been made
    bt_app.Quit(1)  # btSaveOptions.btDoNotSaveChanges: Discards any changes that have been made

# Call the function
print_btw_label("C:\\Users\\PT-PC\\Desktop\\machinelabel.btw", "TSC PEX-1231", 1)
