"""
Original code for excel compare app.
Compares two datasets from excel sheets and analyzes for changes.

Future:
    - Modify 'xlcompare' or create new modules to handle other file types like csv, tab-delim (txt), json
    - organize the various widgets to be more modular
    - like placeholder..
"""

import tkinter as tk
from tkinter import filedialog, messagebox

from xlcompare import ExcelComparator    

class ExcelCompareApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Excel Compare App")

        # File selection boxes
        ## File 1 
        self.input1_label = tk.Label(master, text="File 1:")
        self.input1_label.grid(row=0, column=0, padx=(1,1), pady=10, sticky="e")
        self.input1_entry = tk.Entry(master)
        self.input1_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        self.input1_button = tk.Button(master, text="Browse", command=lambda: self.browse_file(self.input1_entry))
        self.input1_button.grid(row=0, column=2, padx=(10,1), pady=10)
        # placeholder
        self.set_placeholder(self.input1_entry, "Enter file name...", "gray")
        # Bind an event to clear placeholder text when entry is clicked
        self.input1_entry.bind("<FocusIn>", lambda event: self.clear_placeholder(self.input1_entry, "black"))
        # Bind an event to restore placeholder text and color when focus is lost and entry is empty
        self.input1_entry.bind("<FocusOut>", lambda event: self.restore_placeholder(self.input1_entry, "Enter file name...", "gray"))

        ## File 2
        self.input2_label = tk.Label(master, text="File 2:")
        self.input2_label.grid(row=2, column=0, padx=(1,1), pady=10, sticky="e")
        self.input2_entry = tk.Entry(master)
        self.input2_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        self.input2_button = tk.Button(master, text="Browse", command=lambda: self.browse_file(self.input2_entry))
        self.input2_button.grid(row=2, column=2, padx=(10,1), pady=10)
        # placeholder
        self.set_placeholder(self.input2_entry, "Enter file name...", "gray")
        # Bind an event to clear placeholder text when entry is clicked
        self.input2_entry.bind("<FocusIn>", lambda event: self.clear_placeholder(self.input2_entry, "black"))
        # Bind an event to restore placeholder text and color when focus is lost and entry is empty
        self.input2_entry.bind("<FocusOut>", lambda event: self.restore_placeholder(self.input2_entry, "Enter file name...", "gray"))

        # Text input boxes
        # Input box for File1 sheet name
        # self.input1_sheet_label = tk.Label(master, text="Sheet Name:")
        # self.input1_sheet_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.input1_sheet_entry = tk.Entry(master)
        self.input1_sheet_entry.grid(row=1, column=1, padx=(10, 10), pady=10, sticky="ew")
        # Set placeholder text and color
        self.set_placeholder(self.input1_sheet_entry, "Enter file sheet name...", "gray")
        # Bind an event to clear the placeholder text when entry is clicked
        self.input1_sheet_entry.bind("<FocusIn>", lambda event: self.clear_placeholder(self.input1_sheet_entry, "black"))
        # Bind an event to restore the placeholder text and color when focus is lost and entry is empty
        self.input1_sheet_entry.bind("<FocusOut>", lambda event: self.restore_placeholder(self.input1_sheet_entry, "Enter file sheet name...", "gray"))


        # Input box for File2 sheet name
        # self.input2_sheet_label = tk.Label(master, text="Sheet Name:")
        # self.input2_sheet_label.grid(row=3, column=0, padx=10, pady=10, sticky="e")
        self.input2_sheet_entry = tk.Entry(master)
        self.input2_sheet_entry.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        # Set placeholder text and color
        self.set_placeholder(self.input2_sheet_entry, "Enter file sheet name...", "gray")
        # Bind an event to clear the placeholder text when entry is clicked
        self.input2_sheet_entry.bind("<FocusIn>", lambda event: self.clear_placeholder(self.input2_sheet_entry, "black"))
        # Bind an event to restore placeholder text and color when focus is lost and entry is empty
        self.input2_sheet_entry.bind("<FocusOut>", lambda event: self.restore_placeholder(self.input2_sheet_entry, "Enter file sheet name...", "gray"))


        # Input box for common column name
        self.input3_unique_col_label = tk.Label(master, text="Column(s):")
        self.input3_unique_col_label.grid(row=4, column=0, columnspan=None, padx=(1,1), pady=10, sticky="e")
        self.input3_unique_col_entry = tk.Entry(master)
        self.input3_unique_col_entry.grid(row=4, column=1, padx=10, pady=10, sticky="ew")
        # Set placeholder text and color
        self.set_placeholder(self.input3_unique_col_entry, "ie. empl_id, order_id", "gray")
        # Bind an event to clear placeholder text when entry is clicked
        self.input3_unique_col_entry.bind("<FocusIn>", lambda event: self.clear_placeholder(self.input3_unique_col_entry, "black"))
        # Bind an event to restore placeholder text and color when focus is lost and entry is empty
        self.input3_unique_col_entry.bind("<FocusOut>", lambda event: self.restore_placeholder(self.input3_unique_col_entry, "ie. empl_id, order_id", "gray"))

        # Run button
        self.run_button = tk.Button(master, text="Run", command=self.run_comparison)
        self.run_button.grid(row=5, column=0, columnspan=3, pady=20)

        # Column configuration for expansion
        self.master.columnconfigure(1, weight=1)
        # self.master.columnconfigure(2, weight=1)

    def set_placeholder(self, entry, placeholder_text, placeholder_color):
        # Set the placeholder text and color
        entry.insert(0, placeholder_text)
        entry.config(fg=placeholder_color)

    def clear_placeholder(self, entry, text_color):
        # Delete placeholder text when entry is clicked and set text color
        # if entry.get() == "Enter sheet name...": add this so that the placeholder doesnt clear the text if the user already typed something...
        entry.delete(0, tk.END)
        entry.config(fg=text_color)

    def restore_placeholder(self, entry, placeholder_text, placeholder_color):
        # Restore placeholder text and color when focus is lost and entry is empty
        if not entry.get():
            self.set_placeholder(entry, placeholder_text, placeholder_color)


    def browse_file(self, entry):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            entry.delete(0, tk.END)
            entry.insert(0, file_path)

    def run_comparison(self):
        try:
            file_path_1 = self.input1_entry.get() or self.input1_button.cget("text")
            file_path_2 = self.input2_entry.get() or self.input2_button.cget("text")
            sheet_1 = self.input1_sheet_entry.get()
            sheet_2 = self.input2_sheet_entry.get()
            unique_col = self.input3_unique_col_entry.get()

            # Run comparison
            excelCompare = ExcelComparator(file_path_1, file_path_2, sheet_1, sheet_2, unique_col)
            excelCompare.main()

            # Show dialog box after comparison is complete
            messagebox.showinfo("Comparison Complete", "Comparison Complete! Check the output.xlsx file.")
            print("Comparison complete. Check the output.xlsx file.")
        except Exception as e:
            print(e)
            messagebox.showinfo("An Error Has Occurred!", f"{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("350x300")
    app = ExcelCompareApp(root)
    root.mainloop()


    