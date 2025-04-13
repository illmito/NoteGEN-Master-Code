import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime


class WorkNotepad(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)

        # Row 0: Ref No 🆔
        ttk.Label(self, text="Ref No:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.ref_entry = ttk.Entry(self, width=50)
        self.ref_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(self, text="🆔", font=("Segoe UI Emoji", 12)).grid(row=0, column=2, padx=5, pady=5, sticky="w")

        # Row 1: Name 👤
        ttk.Label(self, text="Name:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.name_entry = ttk.Entry(self, width=50)
        self.name_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(self, text="👤", font=("Segoe UI Emoji", 12)).grid(row=1, column=2, padx=5, pady=5, sticky="w")

        # Row 2: Phone 📞
        ttk.Label(self, text="Phone:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.phone_entry = ttk.Entry(self, width=50)
        self.phone_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(self, text="📞", font=("Segoe UI Emoji", 12)).grid(row=2, column=2, padx=5, pady=5, sticky="w")

        # Row 3: Email ✉️
        ttk.Label(self, text="Email:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.email_entry = ttk.Entry(self, width=50)
        self.email_entry.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(self, text="✉️", font=("Segoe UI Emoji", 12)).grid(row=3, column=2, padx=5, pady=5, sticky="w")

        # Row 4: Notes 📝 (spanning column 1 & 2)
# Row 4: Notes label
        ttk.Label(self, text="Notes:").grid(row=4, column=0, sticky="nw", padx=5, pady=5)

        # Frame to hold the text widget and scrollbar
        notes_frame = ttk.Frame(self)
        notes_frame.grid(row=4, column=1, columnspan=2, padx=5, pady=5, sticky="w")

        # Text widget
        self.description_text = tk.Text(notes_frame, width=50, height=15, wrap="word")
        self.description_text.pack(side="left", fill="both", expand=True)

        # Scrollbar widget
        scrollbar = ttk.Scrollbar(notes_frame, orient="vertical", command=self.description_text.yview)
        scrollbar.pack(side="right", fill="y")

        # Connect scrollbar to text widget
        self.description_text.config(yscrollcommand=scrollbar.set)

        # Buttons
        self.copy_button = ttk.Button(self, text="Copy", command=self.copy_note)
        self.copy_button.grid(row=6, column=1, padx=5,columnspan=2, pady=10, sticky="e")

        self.clear_button = ttk.Button(self, text="Clear", command=self.clear_form, underline=1)
        self.clear_button.grid(row=6, column=0,columnspan=2, sticky="w", padx=5, pady=10)

        self.name_entry.focus()




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
            f"Date Time: {timestamp}\n"
            "- - - - - - - - - - - - - - -\n"
            f"Ref no: {reference_number}\n"
            f"Name:   {name}\n"
            f"Phone:  {phone}\n"
            f"Email:  {email}\n\n"
            f"Notes:\n{description}"
        )

        self.clipboard_clear()
        self.clipboard_append(note)
        self.update()

        messagebox.showinfo("Copied", "Note copied to clipboard successfully.")
        self.copy_button.focus()

    def clear_form(self):
        confirm = messagebox.askyesno("Confirm Clear", "Are you sure you want to clear the form?")
        if confirm:
            self.ref_entry.delete(0, tk.END)
            self.name_entry.delete(0, tk.END)
            self.phone_entry.delete(0, tk.END)
            self.email_entry.delete(0, tk.END)
            self.description_text.delete("1.0", tk.END)
            self.name_entry.focus()
        else:
            self.name_entry.focus()
            # Do nothing
            return


