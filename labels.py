import tkinter as tk
import win32print
import win32api
from datetime import datetime

def print_label():
    part_nbr = part_nbr_entry.get()
    quantity = quantity_entry.get()
    date_time = date_entry.get()
    
    # ZPL
    zpl = f"""
    ^XA
    ^PW550
    ^LL380
    ^CF0,50
    ^FO50,20^FD{part_nbr}^FS
    ^BY2,2,100
    ^FO50,70^BCN,100,N,N,N^FD{part_nbr}^FS
    ^FO50,190^FD{date_time}^FS
    ^FO50,240^FDQuantity: {quantity} pcs^FS
    ^BY2,2,80
    ^FO50,290^BCN,80,N,N,N^FD{quantity}^FS
    ^XZ
    """
    
    # Send to printer
    printer_name = win32print.GetDefaultPrinter()
    hprinter = win32print.OpenPrinter(printer_name)
    pdc = win32print.StartDocPrinter(hprinter, 1, ("Label", None, "RAW"))
    win32print.StartPagePrinter(hprinter)
    win32print.WritePrinter(hprinter, zpl.encode('utf-8'))
    win32print.EndPagePrinter(hprinter)
    win32print.EndDocPrinter(hprinter)
    win32print.ClosePrinter(hprinter)

# Main window
root = tk.Tk()
root.title("Друк етикеток")
root.geometry("500x350")
root.resizable(False, False)

# Styles
label_style = ('Arial', 14, 'bold')
entry_style = ('Arial', 18, 'bold')

# Enter Part Number
tk.Label(root, text="Part Number:", font=label_style).pack(pady=5)
part_nbr_entry = tk.Entry(root, width=30, font=entry_style)
part_nbr_entry.pack()

# Enter Date
tk.Label(root, text="Дата:", font=label_style).pack(pady=5)
date_entry = tk.Entry(root, width=30, font=entry_style)
current_datetime = datetime.now().strftime("%d/%m/%Y %H:%M")
date_entry.insert(0, current_datetime)
date_entry.pack()

# Enter Quantity
tk.Label(root, text="Кількість (pcs):", font=label_style).pack(pady=5)
quantity_entry = tk.Entry(root, width=30, font=entry_style)
quantity_entry.pack()

# Print Button
print_button = tk.Button(
    root, 
    text="ДРУК", 
    command=print_label, 
    height=2, 
    width=20,
    font=('Arial', 14, 'bold'),
    bg='#4CAF50',
    fg='white',
    relief=tk.RAISED,
    borderwidth=3
)
print_button.pack(pady=30)

root.mainloop()