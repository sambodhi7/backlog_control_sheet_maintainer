import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from shortFormData import shortFormData

class ShortFormEditor(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Edit Short Form Data")
        self.geometry("800x600")
        self.resizable(True, True)

        self.data = shortFormData.copy()  # Work with a copy

        self._build_ui()
        self._populate_tree()

    def _build_ui(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Treeview for displaying data
        columns = ("Course Code", "Short Form", "Full Names")
        self.tree = ttk.Treeview(main_frame, columns=columns, show="headings", height=15)
        self.tree.heading("Course Code", text="Course Code")
        self.tree.heading("Short Form", text="Short Form")
        self.tree.heading("Full Names", text="Full Names")
        self.tree.column("Course Code", width=150)
        self.tree.column("Short Form", width=150)
        self.tree.column("Full Names", width=300)

        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(btn_frame, text="Add Entry", command=self.add_entry).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Edit Selected", command=self.edit_selected).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Delete Selected", command=self.delete_selected).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Save Changes", command=self.save_changes).pack(side=tk.RIGHT)

    def _populate_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for key, values in self.data.items():
            short_form = values[0] if len(values) > 0 else ""
            full_names = "; ".join(values[1:]) if len(values) > 1 else ""
            self.tree.insert("", tk.END, values=(key, short_form, full_names))

    def add_entry(self):
        dialog = AddEditDialog(self, "Add Entry")
        if dialog.result:
            key, short_form, full_names_str = dialog.result
            if key in self.data:
                messagebox.showerror("Error", "Course Code already exists.")
                return
            full_names = [name.strip() for name in full_names_str.split(",") if name.strip()]
            self.data[key] = [short_form] + full_names
            self._populate_tree()

    def edit_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an entry to edit.")
            return
        item = selected[0]
        values = self.tree.item(item, "values")
        key = values[0]
        short_form = values[1]
        full_names_str = values[2]
        dialog = AddEditDialog(self, "Edit Entry", key, short_form, full_names_str)
        if dialog.result:
            new_key, new_short, new_full_str = dialog.result
            if new_key != key and new_key in self.data:
                messagebox.showerror("Error", "Course Code already exists.")
                return
            full_names = [name.strip() for name in new_full_str.split(",") if name.strip()]
            del self.data[key]
            self.data[new_key] = [new_short] + full_names
            self._populate_tree()

    def delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an entry to delete.")
            return
        if messagebox.askyesno("Confirm", "Are you sure you want to delete the selected entry?"):
            item = selected[0]
            key = self.tree.item(item, "values")[0]
            del self.data[key]
            self._populate_tree()

    def save_changes(self):
        # Write back to shortFormData.py
        with open("shortFormData.py", "w", encoding='utf-8') as f:
            f.write("shortFormData = {\n")
            for key, values in self.data.items():
                f.write(f'    "{key}" : {values},\n')
            f.write("}\n")
        messagebox.showinfo("Success", "Changes saved successfully.")
        self.destroy()

class AddEditDialog(tk.Toplevel):
    def __init__(self, parent, title, key="", short_form="", full_names_str=""):
        super().__init__(parent)
        self.title(title)
        self.geometry("500x400")
        self.resizable(False, False)
        self.result = None

        ttk.Label(self, text="Course Code:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        self.entry_key = ttk.Entry(self)
        self.entry_key.insert(0, key)
        self.entry_key.grid(row=0, column=1, sticky=tk.EW, padx=10, pady=5)

        ttk.Label(self, text="Short Form:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        self.entry_short = ttk.Entry(self)
        self.entry_short.insert(0, short_form)
        self.entry_short.grid(row=1, column=1, sticky=tk.EW, padx=10, pady=5)

        ttk.Label(self, text="Full Names:").grid(row=2, column=0, sticky=tk.NW, padx=10, pady=5)
        
        # Frame for listbox and buttons
        list_frame = ttk.Frame(self)
        list_frame.grid(row=2, column=1, sticky=tk.EW, padx=10, pady=5)
        
        self.listbox = tk.Listbox(list_frame, height=5)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        btn_frame = ttk.Frame(list_frame)
        btn_frame.pack(side=tk.RIGHT, fill=tk.Y)
        
        ttk.Button(btn_frame, text="Add", command=self.add_name).pack(fill=tk.X, pady=2)
        ttk.Button(btn_frame, text="Remove", command=self.remove_name).pack(fill=tk.X, pady=2)
        
        # Populate listbox
        if full_names_str:
            for name in full_names_str.split("; "):
                if name.strip():
                    self.listbox.insert(tk.END, name.strip())

        btn_frame_main = ttk.Frame(self)
        btn_frame_main.grid(row=3, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame_main, text="OK", command=self.ok).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame_main, text="Cancel", command=self.cancel).pack(side=tk.LEFT)

        self.columnconfigure(1, weight=1)

    def add_name(self):
        name = tk.simpledialog.askstring("Add Full Name", "Enter full name:")
        if name and name.strip():
            self.listbox.insert(tk.END, name.strip())

    def remove_name(self):
        selected = self.listbox.curselection()
        if selected:
            self.listbox.delete(selected)

    def ok(self):
        key = self.entry_key.get().strip()
        short = self.entry_short.get().strip()
        full_names = [self.listbox.get(i) for i in range(self.listbox.size())]
        full_str = "; ".join(full_names)
        if not key:
            messagebox.showerror("Error", "Course Code cannot be empty.")
            return
        self.result = (key, short, full_str)
        self.destroy()

    def cancel(self):
        self.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Hide main window
    editor = ShortFormEditor(root)
    root.mainloop()