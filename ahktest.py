import openpyxl as xl
import ahk as ahk
import sys
import os

DATA_WB = xl.load_workbook(filename="data.xlsx")
DATA_WS = DATA_WB.active
data = []
for row in DATA_WS.values:
    data.append(row)

auto = ahk.AHK()
print("starting")

def hotkey():
    for s in data:
        print(s)
        auto.set_clipboard(str(s[0]).upper())
        auto.send_input("^v")
        auto.set_clipboard(str(s[0]))
        auto.send_input("^v")
        auto.key_press('Tab')
    print("stopping")
    auto.stop_hotkeys()
    os._exit(status=0)

auto.add_hotkey("^n", callback=hotkey)
auto.start_hotkeys()
auto.block_forever()
auto.ex
