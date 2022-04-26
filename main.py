from openpyxl import Workbook
from data import data
from gui import *

PATH = "C:/Users/ACER/Desktop/"
workbook = Workbook()
sheet = workbook.active

layout = [layout,
          [sg.TabGroup([[sg.Tab('TEES', layout_tee), sg.Tab('BS', layout_bs, key="BS"), sg.Tab('CAP', layout_cap),
                         sg.Tab('DEBS', layout_deb), sg.Tab('SWS', layout_sws), sg.Tab('POLOS', layout_polo)]])],
          [sg.Button('Ok'), sg.Button('Cancel')]]

# Create the headers
sheet["A1"] = "partner_id"
sheet["B1"] = "order_line/product_id/id"
sheet["C1"] = "order_line/name"
sheet["D1"] = "order_line/product_uom_qty"

# Create all the references in excel
for i in data:
    sheet[i] = data[i]

# Create the Window
window = sg.Window('Import Odoo', layout, grab_anywhere=True, no_titlebar=False)

# Event Loop to process "events"
while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Cancel'):
        break
    if event == 'Ok':
        FILENAME = f"{values['A2']}.xlsx"
        for i in values:
            try:
                sheet[i] = values[i]
            except:
                pass
        workbook.save(filename=f"{PATH}{FILENAME}")
        sg.popup('Saved!', auto_close=True, auto_close_duration=2)
        window.close()
window.close()
