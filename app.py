import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import shutil
from openpyxl import load_workbook
from config import config
from main import get_control_dict, process_subject_file, save_to_control_file

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Exam Grade Consolidator")
        self.geometry("600x650")
        
        self.control_file_path = tk.StringVar()
        self.subject_files = []
        
        self.var_row_start = tk.StringVar(value=str(config.ROW_STARTING))
        self.var_header_row = tk.StringVar(value=str(config.COURSE_PAGE_HEADER_WITH_COURSE_TITLE_ROW_NO))

        self._build_ui()

    def _build_ui(self):
        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        settings_frame = ttk.LabelFrame(main_frame, text="Configuration Settings", padding="10")
        settings_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(settings_frame, text="Control Sheet Data Start Row:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(settings_frame, textvariable=self.var_row_start, width=10).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(settings_frame, text="Subject Sheet Header Row:").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(settings_frame, textvariable=self.var_header_row, width=10).grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)

        lbl_control = ttk.Label(main_frame, text="1. Select Control Sheet (Master):", font=("Segoe UI", 10, "bold"))
        lbl_control.pack(anchor=tk.W, pady=(0, 5))

        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.entry_control = ttk.Entry(control_frame, textvariable=self.control_file_path)
        self.entry_control.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        btn_browse_control = ttk.Button(control_frame, text="Browse...", command=self.select_control)
        btn_browse_control.pack(side=tk.RIGHT)

        lbl_subjects = ttk.Label(main_frame, text="2. Add Subject Files (Result Sheets):", font=("Segoe UI", 10, "bold"))
        lbl_subjects.pack(anchor=tk.W, pady=(0, 5))

        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        self.listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, height=10)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 20))
        
        btn_add = ttk.Button(btn_frame, text="Add Files", command=self.add_subject_files)
        btn_add.pack(side=tk.LEFT, padx=(0, 5))
        
        btn_clear = ttk.Button(btn_frame, text="Clear List", command=self.clear_list)
        btn_clear.pack(side=tk.LEFT)

        sep = ttk.Separator(main_frame, orient='horizontal')
        sep.pack(fill=tk.X, pady=(0, 15))

        self.btn_process = ttk.Button(main_frame, text="Process & Save As...", command=self.process_and_save)
        self.btn_process.pack(fill=tk.X, ipady=10)

        self.status_var = tk.StringVar(value="Ready")
        lbl_status = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        lbl_status.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))

    def select_control(self):
        filename = filedialog.askopenfilename(
            title="Select Control Sheet",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if filename:
            self.control_file_path.set(filename)

    def add_subject_files(self):
        filenames = filedialog.askopenfilenames(
            title="Select Subject Files",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if filenames:
            for f in filenames:
                if f not in self.subject_files:
                    self.subject_files.append(f)
                    self.listbox.insert(tk.END, os.path.basename(f))

    def clear_list(self):
        self.subject_files = []
        self.listbox.delete(0, tk.END)

    def process_and_save(self):
        control_path = self.control_file_path.get()
        if not control_path:
            messagebox.showwarning("Missing Input", "Please select a Control Sheet.")
            return
        if not self.subject_files:
            messagebox.showwarning("Missing Input", "Please add at least one Subject File.")
            return

        try:
            config.ROW_STARTING = int(self.var_row_start.get())
            config.COURSE_PAGE_HEADER_WITH_COURSE_TITLE_ROW_NO = int(self.var_header_row.get())
        except ValueError:
            messagebox.showerror("Invalid Settings", "Row numbers must be integers.")
            return

        output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile="Updated_Control_Sheet.xlsx",
            title="Save Updated Sheet As"
        )
        if not output_file:
            return

        self.status_var.set("Processing... Please wait.")
        self.update_idletasks()
        
        try:
            control_wb = load_workbook(control_path)
            control_sheet = control_wb.active
            
            control_dict = get_control_dict(control_sheet)
            
            newdict = {}
            namelookup = {}
            
            for sub_file in self.subject_files:
                process_subject_file(sub_file, control_dict, newdict, namelookup)
            
            save_to_control_file(control_sheet, control_dict, newdict, namelookup)
            
            default_output = "control_updated.xlsx"
            if os.path.exists(default_output):
                if os.path.abspath(default_output) != os.path.abspath(output_file):
                    shutil.move(default_output, output_file)
            
            self.status_var.set("Done!")
            messagebox.showinfo("Success", f"File saved successfully at:\n{output_file}")
            
        except Exception as e:
            self.status_var.set("Error occurred")
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
            print(e)

if __name__ == "__main__":
    app = App()
    app.mainloop()