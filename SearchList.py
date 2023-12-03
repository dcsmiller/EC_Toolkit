import tkinter as tk
import openpyxl as xl


# Masterlist excel file location
MASTER_LOC: str = "Data/masterlist.xlsx"

#Lookout list location
LOOKOUTLIST_LOC = "lookoutlist.txt"

# Total number of columns in masterlist
TOTAL_COLS: int = 6

# Zero indexed column number of values in report
SERIAL_COL: int  = 0  # Serial Number
ITEM_NUM_COL: int  = 3  # Item number
ITEM_TYPE1_COL: int  = 6  # Item type
ITEM_TYPE2_COL: int  = 5  # Item type 2
LOCATOR_COL: int  = 1  # Item's locator
SUBINV_COL: int  = 2  # Subinventory
DESC_COL: int  = 4  # Item's description

EXCLUDE_INUSE_COUNT: list[str] = ["In-Use.IT_Hub.0.0.0.0.0.0"]
EXCLUDE_SPARE_COUNT: list[str] = []
INCLUDE_INUSE_COUNT: list[str] = []
INCLUDE_SPARE_COUNT: list[str] = ["In-Use.IT_Hub.0.0.0.0.0.0"]

SPARE_SUBINV: list[str] = ["Spare", "Reserved", "Storage"]


class SearchMasterlist:
    def __init__(self) -> None:
        #Initialize the window and frames
        #The lower frame/dashboard still needs to be completed
        self.window: tk.Tk = tk.Tk()
        self.main_frame: tk.Frame = tk.Frame(self.window)
        self.upper_frame: tk.Frame = tk.Frame(self.main_frame)
        self.lower_frame: tk.Frame = tk.Frame(self.main_frame)

        #Set up the geometry managers
        self.main_frame.grid()
        self.upper_frame.pack()
        self.lower_frame.pack()
        self.window.resizable(0, 0)

        self.lb: tk.Listbox = tk.Listbox(self.lower_frame)
        self.lb.pack(expand = 1, fill = 'both')

        #Load excel file, make dicts/lists to hold the data, initialize
        # varibles with a value so there is something to display initially
        self.WORKBOOK: list[list] = xl.load_workbook(MASTER_LOC)
        self.WORKSHEET: list[str] = self.WORKBOOK.active
        self.masterlist: dict[list] = {}
        self.lookoutlist: list[str] = open(LOOKOUTLIST_LOC).read().splitlines()
        print(self.lookoutlist)
        self.spare_inventory_count: dict[str, int] = {"N/A": 0}
        self.inuse_inventory_count: dict[str, int] = {"N/A": 0}
        self.current_inuse_qty: int = 0
        self.current_spare_qty: int = 0
        self.process_excel_data()

        #Set up the bottom indicator - think of a better way
        self.current_indicator: str = ""
        self.indicator_dict: dict[str, str] = {
            "success": ["", "green2"],
            "duplicate": [
                "This serial number is in inventory twice, in both upper and lower case",
                "yellow",
            ],
            "not_found": [
                "Item with this serial number is either not in inventory, liquidated, or not virtually at SAV4",
                "red",
            ],
            "lookout": ["Item is on lookout list", "gold"],
            "missing": ["Item is missing", "blue"]
        }

        # Make a blank row
        self.blank_row: list[str] = []
        for i in range(TOTAL_COLS + 1):
            self.blank_row.append("N/A")

        # Set current row to the blank row, so there is something to be displayed
        self.current_row: list[str] = self.blank_row.copy()

        #Make the widgets
        self.create_widgets()

    def create_widgets(self) -> None:
        padding = 3
        self.searchbox: tk.Entry = tk.Entry(self.upper_frame)
        self.searchbox.bind("<Return>", self.submit)
        self.searchbox.focus_set()

        self.serial_display: SelectableText = SelectableText(
            self.upper_frame, text=self.current_row[SERIAL_COL]
        )
        self.desc_display: SelectableText = SelectableText(
            self.upper_frame, text=self.current_row[DESC_COL]
        )
        self.itemnum_display: SelectableText = SelectableText(
            self.upper_frame, text=self.current_row[ITEM_NUM_COL]
        )
        self.itemtype_display: SelectableText = SelectableText(
            self.upper_frame, text=self.current_row[ITEM_TYPE1_COL]
        )
        self.itemtype2_display = SelectableText(
            self.upper_frame, text=self.current_row[ITEM_TYPE2_COL]
        )
        self.subinventory_display: SelectableText = SelectableText(
            self.upper_frame, text=self.current_row[SUBINV_COL]
        )
        self.location_display: SelectableText = SelectableText(
            self.upper_frame, text=self.current_row[LOCATOR_COL]
        )

        self.search_displays: dict[SelectableText, int] = {
            self.serial_display: SERIAL_COL,
            self.desc_display: DESC_COL,
            self.itemnum_display: ITEM_NUM_COL,
            self.itemtype_display: ITEM_TYPE1_COL,
            self.itemtype2_display: ITEM_TYPE2_COL,
            self.subinventory_display: SUBINV_COL,
            self.location_display: LOCATOR_COL,
        }

        self.serial_lbl: tk.Label = tk.Label(self.upper_frame, text="Serial: ")
        self.desc_lbl: tk.Label = tk.Label(self.upper_frame, text="Description: ")
        self.itemnum_lbl: tk.Label = tk.Label(self.upper_frame, text="ItemNum: ")
        self.itemtype_lbl: tk.Label = tk.Label(self.upper_frame, text="ItemType: ")
        self.itemtype2_lbl: tk.Label = tk.Label(self.upper_frame, text="ItemType2: ")
        self.subinventory_lbl: tk.Label = tk.Label(
            self.upper_frame, text="Subinventory: "
        )
        self.location_lbl: tk.Label = tk.Label(self.upper_frame, text="Location: ")

        self.indicator_lbl: tk.Label = tk.Label(
            self.upper_frame,
            text="Scan an item.",
            bg="white",
            fg="black",
            font="Terminal 10",
        )

        self.searchbox.grid(
            row=0, column=0, columnspan=12, sticky="ew", padx=padding, pady=padding
        )

        self.serial_lbl.grid(row=1, column=0, sticky="w", pady=padding)
        self.serial_display.grid(
            row=1, column=1, columnspan=2, sticky="w", pady=padding
        )
        self.itemnum_lbl.grid(row=1, column=3, sticky="w", pady=padding)
        self.itemnum_display.grid(
            row=1, column=4, columnspan=2, sticky="w", pady=padding
        )
        self.itemtype_lbl.grid(row=1, column=6, sticky="w", pady=padding)
        self.itemtype_display.grid(
            row=1, column=7, columnspan=2, sticky="w", pady=padding
        )
        self.itemtype2_lbl.grid(row=1, column=9, sticky="w", pady=padding)
        self.itemtype2_display.grid(
            row=1, column=10, columnspan=2, sticky="w", pady=padding
        )

        self.desc_lbl.grid(row=2, column=0, sticky="w", pady=padding)
        self.desc_display.grid(
            row=2, column=1, columnspan=11, sticky="ew", pady=padding
        )

        self.subinventory_lbl.grid(row=3, column=0, sticky="w", pady=padding)
        self.subinventory_display.grid(
            row=3, column=1, columnspan=2, sticky="w", pady=padding
        )
        self.location_lbl.grid(row=3, column=3, sticky="w", pady=padding)
        self.location_display.grid(
            row=3, column=4, columnspan=8, sticky="ew", pady=padding
        )

        self.indicator_lbl.grid(
            row=4, column=0, sticky="ew", columnspan=12, pady=padding
        )

    def process_excel_data(self) -> None:
        for row in self.WORKSHEET.values:
            self.masterlist.update({row[SERIAL_COL]: row})
            if row[SUBINV_COL] == "Missing":
                self.lookoutlist.append(row[SERIAL_COL])
            self.count_item(row)
        print(self.spare_inventory_count)

    def count_item(self, row: list[str]) -> None:
        current_subinv = row[SUBINV_COL]
        current_loc = row[LOCATOR_COL]
        current_item_num = row[ITEM_NUM_COL]

        #Perform surplus count
        if (
            current_subinv in SPARE_SUBINV or current_loc in INCLUDE_SPARE_COUNT
        ) and current_loc not in EXCLUDE_SPARE_COUNT:
            if current_item_num in self.spare_inventory_count:
                self.spare_inventory_count[current_item_num] += 1
            else:
                self.spare_inventory_count[current_item_num] = 1

        #Perform In-Use count
        if (
            current_subinv == "In-Use" or current_loc in INCLUDE_INUSE_COUNT
        ) and current_loc not in EXCLUDE_INUSE_COUNT:
            if current_item_num in self.inuse_inventory_count:
                self.inuse_inventory_count[current_item_num] += 1
            else:
                self.inuse_inventory_count[current_item_num] = 1

    def submit(self, event: tk.Event) -> None:
        searchvalue: str = self.searchbox.get()
        self.current_row = self.search(searchvalue)
        self.check_lookout(searchvalue)
        self.check_missing()
        self.update_search_displays()
        self.update_inventory_qty()
        self.update_indicator()
        self.lb.insert(tk.END, self.current_row)
        print(type(self.lb.get(0)))
        self.searchbox.delete(0, tk.END)

    def search(self, serial: str) -> list:
        if serial.upper() in self.masterlist and serial.upper() in self.masterlist:
            self.current_indicator = "duplicate"
            return self.masterlist[serial]

        if serial in self.masterlist:
            self.current_indicator = "success"
            return self.masterlist[serial]
        else:
            if serial.lower() in self.masterlist:
                self.current_indicator = "success"
                return self.masterlist[serial.lower()]
            elif serial.upper() in self.masterlist:
                self.current_indicator = "success"
                return self.masterlist[serial.upper()]
            else:
                self.current_indicator = "not_found"
                return self.blank_row

    def check_lookout(self, serial: str) -> None:
        if serial in self.lookoutlist:
            self.current_indicator = "lookout"
    
    def check_missing(self) -> None:
        if self.current_row[SUBINV_COL] == "Missing":
            self.current_indicator = "missing"

    def update_search_displays(self) -> None:
        for display in self.search_displays:
            col: int = self.search_displays[display]
            display.update(text=self.current_row[col])

    def update_indicator(self) -> None:
        indicator: list[str] = self.indicator_dict[self.current_indicator].copy()
        if self.current_indicator == "success":
            indicator[0] += f"In-Use QTY: {self.current_inuse_qty}    |    Spare QTY: {self.current_spare_qty}"
        self.indicator_lbl.configure(text=indicator[0], bg=indicator[1])

    def update_inventory_qty(self) -> None:
        current_item_number = self.current_row[ITEM_NUM_COL]
        if current_item_number in self.spare_inventory_count:
            self.current_spare_qty = self.spare_inventory_count[current_item_number]
        else:
            self.current_spare_qty = 0

        if current_item_number in self.inuse_inventory_count:
            self.current_inuse_qty = self.inuse_inventory_count[current_item_number]
        else:
            self.current_inuse_qty = 0


class SelectableText(tk.Entry):
    def __init__(self, parent: tk.Frame, text: str, **kwargs: dict[str, str]) -> None:
        kwargs["font"] = "Terminal 16"
        kwargs["bd"] = 2
        kwargs["readonlybackground"] = "white"
        kwargs["relief"] = "groove"
        kwargs["takefocus"] = "0"
        super().__init__(parent, **kwargs)
        self.update(text)

    def update(self, text: str) -> None:
        self["state"] = "normal"
        self.delete(0, tk.END)
        self.insert(0, text)
        self["state"] = "readonly"


if __name__ == "__main__":
    app = SearchMasterlist()
    app.window.mainloop()
    exit()
