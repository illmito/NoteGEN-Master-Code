import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime


class WorkNotepad(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        # Labels and fields
        ttk.Label(self, text="Ref No:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.ref_entry = ttk.Entry(self, width=50)
        self.ref_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(self, text="Name:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.name_entry = ttk.Entry(self, width=50)
        self.name_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(self, text="Phone:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.phone_entry = ttk.Entry(self, width=50)
        self.phone_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(self, text="Email:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.email_entry = ttk.Entry(self, width=50)
        self.email_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(self, text="Notes:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.description_text = tk.Text(self, width=38, height=10)
        self.description_text.grid(row=4, column=1, padx=5, pady=5, sticky="w")

        # Buttons
        self.copy_button = ttk.Button(self, text="Copy", command=self.copy_note)
        self.copy_button.grid(row=6, column=0, padx=5, pady=10, sticky="w")

        self.clear_button = ttk.Button(self, text="Clear", command=self.clear_form)
        self.clear_button.grid(row=6, column=1, sticky="e", padx=5, pady=10)

    def copy_note(self):
        reference_number = self.ref_entry.get()
        name = self.name_entry.get()
        phone = self.phone_entry.get()
        email = self.email_entry.get()
        description = self.description_text.get("1.0", tk.END).strip()

        if not name:
            messagebox.showwarning("Incomplete Data", "POC Name Required")
            self.name_entry.focus()
            return
        elif not phone and not email:
            messagebox.showwarning("Missing Contact", f"No Contact for {name}?")
            self.phone_entry.focus()
            return

        timestamp = datetime.now().strftime("%H:%M:%S - %d/%m/%y ")
        note = (
            f"Date Time: {timestamp}\n\n"
            f"Ref no: {reference_number}\n"
            f"Name:   {name}\n"
            f"Phone : {phone}\n"
            f"Email:  {email}\n\n"
            f"Notes:\n{description}"
        )

        self.clipboard_clear()
        self.clipboard_append(note)
        self.update()

        messagebox.showinfo("Copied", "Note copied to clipboard successfully.")
        self.copy_button.focus()

    def clear_form(self):
        self.ref_entry.delete(0, tk.END)
        self.name_entry.delete(0, tk.END)
        self.phone_entry.delete(0, tk.END)
        self.email_entry.delete(0, tk.END)
        self.description_text.delete("1.0", tk.END)


