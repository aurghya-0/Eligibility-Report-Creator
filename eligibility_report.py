import os
from tkinter import *
from tkinter import ttk, filedialog, messagebox
import threading
from eligibility_processor import extract_subject_codes, process_file

class EligibilityApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Eligibility Report Generator")
        self.root.geometry("800x600")

        self.input_filepath = ""
        self.output_folder_path = ""
        self.subject_checkboxes = {}
        self.subject_vars = {}
        self.combine_subjects = BooleanVar()

        # Top Buttons
        frame = Frame(root)
        frame.pack(pady=10)

        Button(frame, text="Select Excel File", command=self.select_file).grid(row=0, column=0, padx=5)
        Button(frame, text="Select Output Folder", command=self.select_folder).grid(row=0, column=1, padx=5)

        # Combine Checkbox
        Checkbutton(root, text="Combine Subjects into One PDF", variable=self.combine_subjects).pack(pady=5)

        # Search Entry
        search_frame = Frame(root)
        search_frame.pack(pady=5)
        Label(search_frame, text="Search Subject:").pack(side=LEFT, padx=5)
        self.search_var = StringVar()
        self.search_var.trace("w", self.filter_subjects)
        Entry(search_frame, textvariable=self.search_var, width=50).pack(side=LEFT)

        # Scrollable Subject Selection
        self.subject_frame = Frame(root)
        self.subject_frame.pack(fill=BOTH, expand=True, pady=10)
        canvas = Canvas(self.subject_frame)
        scrollbar = Scrollbar(self.subject_frame, orient=VERTICAL, command=canvas.yview)
        self.scrollable_frame = Frame(canvas)

        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        # Bottom Buttons
        Button(root, text="Export Selected Subjects", command=self.export_selected).pack(pady=10)
        Button(root, text="Export All PDFs", command=self.export_all).pack()

    def select_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            self.input_filepath = filepath
            self.load_subjects()

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.output_folder_path = folder_path

    def load_subjects(self):
        subjects, df = extract_subject_codes(self.input_filepath)
        self.df = df
        self.subject_checkboxes.clear()
        self.subject_vars.clear()

        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        for code, name in subjects:
            var = BooleanVar()
            cb = Checkbutton(self.scrollable_frame, text=f"{code} - {name}", variable=var, anchor="w", width=80, justify=LEFT)
            cb.pack(fill=X, anchor="w")
            self.subject_checkboxes[f"{code} - {name}"] = cb
            self.subject_vars[f"{code} - {name}"] = (var, code)

    def filter_subjects(self, *args):
        search_term = self.search_var.get().lower()
        for key, cb in self.subject_checkboxes.items():
            if search_term in key.lower():
                cb.pack(fill=X, anchor="w")
            else:
                cb.pack_forget()

    def export_selected(self):
        selected_codes = [code for key, (var, code) in self.subject_vars.items() if var.get()]
        if not self.input_filepath or not self.output_folder_path:
            messagebox.showwarning("Missing Info", "Please select both input file and output folder.")
            return
        if not selected_codes:
            messagebox.showinfo("No Subjects", "Please select at least one subject to export.")
            return

        threading.Thread(target=self.run_process, args=(selected_codes,), daemon=True).start()

    def export_all(self):
        all_codes = [code for _, code in self.subject_vars.values()]
        if not self.input_filepath or not self.output_folder_path:
            messagebox.showwarning("Missing Info", "Please select both input file and output folder.")
            return

        threading.Thread(target=self.run_process, args=(all_codes,), daemon=True).start()

    def run_process(self, codes):
        try:
            output_file = process_file(self.df.copy(), codes, self.output_folder_path, self.combine_subjects.get())
            messagebox.showinfo("Success", f"Reports generated successfully.\n\nSaved to:\n{output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

if __name__ == "__main__":
    root = Tk()
    app = EligibilityApp(root)
    root.mainloop()
