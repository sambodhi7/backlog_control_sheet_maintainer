import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import shutil
from openpyxl import load_workbook
from a import codeLookup, get_control_dict, make_new_dict_from_control, process_subject_file2, save_to_control_file2


class AppA(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Backlog Control Sheet Maintainer")
        self.geometry("620x620")
        self.control_file_path = tk.StringVar()
        self.subject_files = []
        self._build_ui()

    def _build_ui(self):
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        label = ttk.Label(main_frame, text="A.py Backlog Processor", font=("Segoe UI", 14, "bold"))
        label.pack(anchor=tk.W, pady=(0, 10))

        control_frame = ttk.LabelFrame(main_frame, text="1. Control Sheet", padding=10)
        control_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(control_frame, text="Control Excel File:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        entry = ttk.Entry(control_frame, textvariable=self.control_file_path)
        entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=5)
        btn_browse = ttk.Button(control_frame, text="Browse...", command=self.select_control)
        btn_browse.grid(row=0, column=2, padx=5, pady=5)
        control_frame.columnconfigure(1, weight=1)

        subjects_frame = ttk.LabelFrame(main_frame, text="2. Subject Files", padding=10)
        subjects_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        self.subject_listbox = tk.Listbox(subjects_frame, selectmode=tk.EXTENDED, height=12)
        self.subject_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        scrollbar = ttk.Scrollbar(subjects_frame, orient=tk.VERTICAL, command=self.subject_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.subject_listbox.config(yscrollcommand=scrollbar.set)

        subject_buttons = ttk.Frame(main_frame)
        subject_buttons.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(subject_buttons, text="Add Files", command=self.add_subject_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(subject_buttons, text="Remove Selected", command=self.remove_selected_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(subject_buttons, text="Clear List", command=self.clear_list).pack(side=tk.LEFT)

        process_frame = ttk.LabelFrame(main_frame, text="3. Process and Save", padding=10)
        process_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Button(process_frame, text="Process & Save As...", command=self.process_and_save).pack(fill=tk.X, ipady=10)

        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_label.pack(fill=tk.X, pady=(10, 0))

    def select_control(self):
        path = filedialog.askopenfilename(
            title="Select Control Sheet",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if path:
            self.control_file_path.set(path)

    def add_subject_files(self):
        paths = filedialog.askopenfilenames(
            title="Select Subject Files",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if paths:
            for path in paths:
                if path not in self.subject_files:
                    self.subject_files.append(path)
                    self.subject_listbox.insert(tk.END, os.path.basename(path))

    def remove_selected_files(self):
        selected = list(self.subject_listbox.curselection())
        if not selected:
            return
        for index in reversed(selected):
            self.subject_listbox.delete(index)
            self.subject_files.pop(index)

    def clear_list(self):
        self.subject_files = []
        self.subject_listbox.delete(0, tk.END)

    def process_and_save(self):
        control_path = self.control_file_path.get().strip()
        if not control_path:
            messagebox.showwarning("Missing Control File", "Please select a control sheet first.")
            return
        if not self.subject_files:
            messagebox.showwarning("Missing Subject Files", "Please add at least one subject file.")
            return

        output_path = filedialog.asksaveasfilename(
            title="Save Updated Control Workbook",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile="updated_control2.xlsx"
        )
        if not output_path:
            return

        self.status_var.set("Processing...")
        self.update_idletasks()

        try:
            control_wb = load_workbook(control_path)
            control_sheet = control_wb.active
            control_dict = get_control_dict(control_sheet)
            last_row = control_dict.get("last_row", 1000)
            control_dict.pop("last_row", None)

            newdict = make_new_dict_from_control(control_dict)
            for subject_file in self.subject_files:
                process_subject_file2(subject_file, control_dict, newdict, codeLookup)

            save_to_control_file2(output_path, control_dict, newdict, {}, last_row)

            self.status_var.set("Done")
            messagebox.showinfo("Success", f"Updated file saved successfully:\n{output_path}")
        except Exception as exc:
            self.status_var.set("Error")
            messagebox.showerror("Error", f"Processing failed:\n{exc}")
            print(exc)


if __name__ == "__main__":
    app = AppA()
    app.mainloop()
