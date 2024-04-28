import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk
import json
from classes import *


class SpreadsheetApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.notebook = None
        self.file_menu = None
        self.menu_bar = None
        self.title("Spreadsheet App")
        self.workbook = Workbook()  # משתמש במחלקה הקיימת
        self.current_sheet_name = None
        self.entries = {}  # Dictionary to hold Entry widgets
        self.setup_menus()
        self.sheet_frame = None
        self.setup_notebook()  # קריאה להגדרת ה-notebook

    def setup_notebook(self):
        self.notebook = ttk.Notebook(self)  # יצירת אובייקט Notebook
        self.notebook.pack(fill='both', expand=True)  # הצגת ה-notebook בחלון

    def add_sheet_tab(self, sheet_name):
        frame = ttk.Frame(self.notebook)  # יצירת פריים לכל טאב
        self.notebook.add(frame, text=sheet_name)  # הוספת הטאב ל-notebook
        self.notebook.select(frame)  # בחירת הטאב החדש
        return frame

    def setup_menus(self):
        self.menu_bar = tk.Menu(self)
        self.config(menu=self.menu_bar)

        self.file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.file_menu.add_command(label="New Sheet", command=self.create_new_sheet)
        self.file_menu.add_command(label="Open Workbook", command=self.open_workbook)
        self.file_menu.add_command(label="Save Workbook", command=self.save_workbook)
        self.file_menu.add_command(label="Save Workbook As...", command=self.save_workbook_as)
        self.menu_bar.add_cascade(label="File", menu=self.file_menu)

    def create_grid(self, rows=10, columns=10):
        if self.sheet_frame:
            self.sheet_frame.destroy()
        self.sheet_frame = tk.Frame(self)
        self.sheet_frame.pack(fill=tk.BOTH, expand=True)
        title_label = tk.Label(self.sheet_frame, text=f"Current Sheet: {self.current_sheet_name}")
        title_label.grid(row=0, column=0, columnspan=columns, sticky="ew")

        for j in range(columns):
            col_header = tk.Label(self.sheet_frame, text=chr(65 + j))  # Headers A, B, C, etc.
            col_header.grid(row=1, column=j + 1)

        for i in range(rows):
            row_header = tk.Label(self.sheet_frame, text=str(i + 1))  # Row numbers 1, 2, 3, etc.
            row_header.grid(row=i + 2, column=0)

            for j in range(columns):
                entry = tk.Entry(self.sheet_frame, width=10)
                entry.grid(row=i + 2, column=j + 1)
                entry.bind('<FocusOut>',
                           lambda e, r=i, c=j: self.cell_updated(r, c, e.widget))  # Bind for updating cell values

                self.entries[(i, j)] = entry

    def cell_updated(self, row, col, widget):
        text = widget.get()
        sheet = self.workbook.get_sheet(self.current_sheet_name)
        cell = sheet.get_cell(row, col)
        cell.insert_text(text)
        self.refresh_ui()


    def show_formula(self, row, col, widget):
        cell = self.workbook.get_sheet(self.current_sheet_name).get_cell(row, col)
        # Debug output to verify the cell's status
        print(f"Checking cell at row {row}, col {col}: formula={cell.formula}, value={cell.value}")
        if cell.formula:
            messagebox.showinfo("Formula", f"Formula in cell {chr(65 + col)}{row + 1}: {cell.formula}")
        else:
            messagebox.showinfo("Formula", "No formula defined for this cell.")

    def refresh_ui(self):
        sheet = self.workbook.get_sheet(self.current_sheet_name)
        for (row, col), entry in self.entries.items():
            cell = sheet.get_cell(row, col)
            entry.delete(0, tk.END)
            entry.insert(0, cell.get_display_value())

    def create_new_sheet(self):
        new_sheet_name = simpledialog.askstring("New Sheet", "Enter the name of the new sheet:")
        if new_sheet_name:
            self.workbook.add_sheet(new_sheet_name)
            self.current_sheet_name = new_sheet_name
            self.create_grid()

    def open_workbook(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if file_path:
            with open(file_path, 'r') as file:
                data = json.load(file)
                self.workbook.load_from_json(data)  # טעינת הנתונים למחלקת Workbook
                self.current_sheet_name = next(iter(self.workbook.sheets))  # בחירת הדף הראשון להצגה
                print(self.workbook.sheets[self.current_sheet_name].table[0][0].text)

                self.create_grid()  # יצירת הגריד עם הנתונים החדשים
                self.refresh_ui()


    def save_workbook(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json")
        if file_path:
            self.save_workbook_as(file_path)

    def save_workbook_as(self, file_path):
        with open(file_path, 'w') as file:
            json.dump(self.workbook.to_json(), file)
            print("Workbook saved.")


if __name__ == "__main__":
    app = SpreadsheetApp()
    app.mainloop()
