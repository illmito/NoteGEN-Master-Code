from tkinter import filedialog
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime
import os
import csv

from email_window import EmailTab


from pystray import Icon, Menu, MenuItem
from PIL import Image, ImageDraw, ImageFont
import threading

# Import the history window functionality
from history_window import open_history_window




def create_tray_icon_image():
    """Generate a simple tray icon."""
    # Create a 16x16 white image
    image = Image.new("RGB", (16, 16), "white")
    draw = ImageDraw.Draw(image)

    # Draw a rounded rectangle that fits within the 16x16 canvas
    draw.rounded_rectangle((0, 0, 15, 15), radius=4, fill="white", outline="black")

    # Optional: Load a font for better text rendering
    try:
        # Load a default font; adjust size as needed
        font = ImageFont.truetype("arial.ttf", 12)
    except IOError:
        # Fallback to the default PIL font
        font = ImageFont.load_default()

    # Calculate the text position to center it
    text = "NG"
    bbox = font.getbbox(text)  # Get the bounding box of the text
    text_width = bbox[2] - bbox[0]  # Calculate text width
    text_height = bbox[3] - bbox[1]  # Calculate text height
    text_x = (16 - text_width) // 2
    text_y = (16 - text_height) // 2

    # Draw the text in black
    draw.text((text_x, text_y), text, fill="black", font=font)

    return image


# Start the system tray
def start_tray_icon():
    """Start the system tray icon."""
    def on_show_window(icon, item):
        """Restore the main application window from the tray."""
        root.deiconify()  # Show the Tkinter window
        icon.stop()  # Stop the tray icon

    def on_exit(icon, item):
        """Exit the application."""
        icon.stop()  # Stop the tray icon
        confirm_exit()
        root.destroy()  # Close the Tkinter window

    # Define menu items for the tray icon
    menu = Menu(
        MenuItem("Open", on_show_window),
        MenuItem("History", open_history_window),
        MenuItem("Email", open_email_window),
        MenuItem("Close", on_exit),
    )



    # Create and configure the tray icon
    icon = Icon("NoteGen", create_tray_icon_image(), "NoteGen Application", menu)

    # Run the tray icon in a separate thread
    tray_thread = threading.Thread(target=icon.run, daemon=True)
    tray_thread.start()


def minimize_to_tray():
    """Minimize the application to the system tray."""
    root.withdraw()  # Hide the Tkinter window
    start_tray_icon()  # Start the system tray icon




# Function to generate and display the note in real time
def update_note(*args):
    # Always update note regardless of preview toggle
    note_output.config(state=tk.NORMAL)
    note_output.delete("1.0", tk.END)

    reference_number_entry = reference_number_entry_entry.get().strip()
    site_contact = site_contact_entry.get().strip().title()
    new_priority = priority_var.get().strip()
    contractor = contractor_entry.get().strip().title()
    reason_note_menu = reason_note_var.get()
    note = note_entry.get("1.0", tk.END).strip()

    generated_note = ""
    if reference_number_entry:
        generated_note += f"Reference No:   {reference_number_entry}\n"
    if new_priority:
        generated_note += f"New Priority:   {new_priority}\n"
    if site_contact:
        generated_note += f"Site Contact:   {site_contact}\n"
    if contractor:
        generated_note += f"Contractor:     {contractor}\n"

    actions = {
        "Contacted": contacted_var.get(),
        "Dispatched": dispatched_var.get(),
        "Emailed": emailed_var.get(),
        "Confirmed": confirmed_var.get(),
        "Cancelled": cancelled_var.get(),
        "POC Callback": poc_callback_var.get(),
    }

    actions_added = False
    for action, is_checked in actions.items():
        if is_checked:
            if not actions_added:
                generated_note += "\nAction:\n"
                actions_added = True
            generated_note += f"   - {action}:\n"

    if reason_note_menu == "Select Reason for Cancel and/or Priority Change":
        f""
    elif reason_note_menu:
        generated_note += f"\nNotes:\n  - {reason_note_menu}\n"

    if note:
        if not reason_note_menu:
            generated_note += "\nNotes:\n - "
        generated_note += f"{note}\n"

    # Always insert the generated note (even if hidden)
    note_output.insert("1.0", generated_note)
    note_output.yview(tk.END)
    note_output.config(state=tk.DISABLED)

    # Enable save button if there is data
    save_button.config(state=tk.NORMAL if reference_number_entry else tk.DISABLED)

    return reference_number_entry


# Function to toggle note preview (visibility only)
def toggle_preview():
    if preview_var.get():
        # Show scrolled text
        note_output.place(x=10, y=400, width=500, height=200)
        root.geometry("520x680")
    else:
        # Hide scrolled text
        note_output.place_forget()
        root.geometry("520x480")


# Function to copy the note text to clipboard
def copy_to_clipboard():
    root.clipboard_clear()
    root.clipboard_append(note_output.get("1.0", tk.END).strip())
    root.update()




# Function to clear all fields
def clear_all():
    reference_number_entry_entry.delete(0, tk.END)
    site_contact_entry.delete(0, tk.END)
    contractor_entry.delete(0, tk.END)
    priority_var.set("")
    reason_note_var.set("")
    note_entry.delete("1.0", tk.END)
    
    for var in [contacted_var, dispatched_var, emailed_var, confirmed_var, cancelled_var, poc_callback_var]:
        var.set(False)

    note_output.config(state=tk.NORMAL)
    note_output.delete("1.0", tk.END)
    note_output.insert("1.0", "Notes cleared.")  # Optional feedback
    note_output.config(state=tk.DISABLED)

    # Reset scrolled text height and root window
    # note_output.place_configure(height=250)
    # root.geometry("520x500")

    hide_success_label()
    save_button.config(state=tk.DISABLED)


    # Hide success label
    hide_success_label()
    
    # Disable save button initially
    save_button.config(state=tk.DISABLED)


def validate_inputs():
    reference_number_entry = reference_number_entry_entry.get().strip()
    if not reference_number_entry:  # Check if the Reference Number field is empty
        messagebox.showerror("Input Error", "The 'Reference Number' field is mandatory.")
        return False
    return True



# Function to save note and display confirmation
def save_note(event=None):  # Accept an optional event parameter for hotkey binding
    if not validate_inputs():
        return  # Exit the function if validation fails

    # Extract data for CSV
    reference_number_save = reference_number_entry_entry.get().strip()
    site_contact = site_contact_entry.get().strip().title()
    new_priority = priority_var.get().strip()
    contractor = contractor_entry.get().strip().title()
    reason_note_menu = reason_note_var.get()
    note = note_entry.get("1.0", tk.END).strip()
    actions = []
    if contacted_var.get(): actions.append("Contacted")
    if dispatched_var.get(): actions.append("Dispatched")
    if emailed_var.get(): actions.append("Emailed")
    if confirmed_var.get(): actions.append("Confirmed")
    if cancelled_var.get(): actions.append("Cancelled")
    if poc_callback_var.get(): actions.append("POC Callback")
    actions_str = ", ".join(actions)

    # Save data to CSV
    file_exists = os.path.isfile("notegenhistory.csv")
    with open("notegenhistory.csv", "a", newline='') as csvfile:
        fieldnames = ["Date Time", "Reference Number", "Site Contact", "New Priority", "Contractor", "Actions", "Notes",]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        if not file_exists:
            writer.writeheader()
        writer.writerow({
            "Date Time": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "Reference Number": reference_number_save,
            "Site Contact": site_contact,
            "New Priority": new_priority,
            "Contractor": contractor,
            "Actions": actions_str,
            "Notes": f"{reason_note_menu}\n{note}".strip()
        })

    messagebox.showinfo("Save Confirmation", "Note saved to notegenhistory.csv successfully.")
    show_success_label()
    save_button.config(state=tk.DISABLED)  # Disable save button until changes are made

# Function to show "Successfully Saved" label
def show_success_label():
    reference_number = reference_number_entry_entry.get().strip()  # Get the text from the entry
    timestamp = datetime.now().strftime(" %d-%m-%Y %H:%M")
    success_label.config(text=f"Reference: {reference_number} saved   |    {timestamp}")
    success_label.place(relx=1.0, rely=1.0, anchor="se")
    root.after(60000, hide_success_label)  # Hide after 3 seconds (90000 is 90 seconds)



# Function to hide the "Successfully Saved" label
def hide_success_label():
    success_label.place_forget()


# Function to confirm exit
def confirm_exit():
    if messagebox.askokcancel("Quit",  "Any unsaved work will be lost\n\nDo you want to continue?"):
        root.destroy()  # Quit the application if the user confirms


# info / about window
def open_about_window():
    about_window = tk.Toplevel()
    about_window.title("About - NoteGEN")
    about_window.geometry("300x280")
    
    label1 = ttk.Label(about_window, text="NoteGEN v3.0", font=("Arial", 12, "bold"))
    label1.pack(pady=(5, 10))  # Add padding to the first label

    label5 = ttk.Label(about_window, 
                       text="Ctrl + H : Opens History \nCtrl + E : Opens Email Templates\nCtrl + S : Save Note\n", 
                       font=("Arial", 10))  # Slightly larger font
    label5.pack(pady=(0, 5))  # Add bottom padding

    emaildesc = ttk.Label(about_window, text="Email Templates:", font=("Arial", 10, "bold"))  # Slightly larger font
    emaildesc.pack(pady=(0, 0), anchor="center")  # Add bottom padding
    emaildesc_info = ttk.Label(about_window, text="Set a Default Email Signature\nplace signature on top of list.", font=("Arial", 8,))  # Slightly larger font
    emaildesc_info.pack(pady=(0, 5), anchor="center")  # Add bottom padding
#:\n\nPlease ensure a default Email Signature is set and on the top of the list.


    label3 = ttk.Label(about_window, text="\u00A9 2025 Thomas Brayovic", font=("Arial", 8, "bold"))  # Slightly larger font
    label3.pack(pady=(10, 10))  # Add bottom padding

   
    close_button = ttk.Button(about_window, text="sure", command=about_window.destroy)
    close_button.pack(pady=10)
    close_button.focus_force()
    #function that allows to press escape to close window.
    def close_window(event=None):
        
        about_window.destroy()

    about_window.bind("<Escape>", close_window )


# Root setup
root = tk.Tk()
root.title("NoteGen")
root.geometry("520x480")
root.resizable(True, True)


# notebook = ttk.Notebook(root)
# notebook.pack(expand=True, fill="both")

# # Main Tab
# main_tab = ttk.Frame(root)
# notebook.add(main_tab, text="Main Tab")

# # Email Tab
# email_tab = EmailTab(ttk.Frame)
# notebook.add(email_tab, text="Email Tab")





# Create menu
menu = tk.Menu(root)
root.config(menu=menu)

fmenu = tk.Menu(root)
root.config(menu=menu)

file_menu = tk.Menu(menu, tearoff=0)
file_menu.add_command(label="About", command=open_about_window)
file_menu.add_cascade(menu="disabled")
file_menu.add_separator()
file_menu.add_command(label="Exit", command=confirm_exit) 

options_menu = tk.Menu(menu, tearoff=0)
options_menu.add_command(label="Send An Email", command=lambda: open_email_window())
options_menu.add_separator()
options_menu.add_command(label="History", command=lambda: open_history_window(ttk.Frame))

menu.add_cascade(label="File", menu=file_menu)
menu.add_cascade(label="Options",menu=options_menu)

options_menu = tk.Menu(file_menu, tearoff=0)


 # Use confirm_exit function
# Add a "Minimize to Tray" button


# Bind the close button (Windows X button) to the confirm_exit function
root.protocol("WM_DELETE_WINDOW", confirm_exit)





# short cuts
# Bind Ctrl+S to save the note
root.bind("<Control-s>", save_note)
root.bind("<Control-S>", save_note)
# Bind Ctrl+H to open the history window
root.bind("<Control-h>", open_history_window)
root.bind("<Control-H>", open_history_window)





# input  frame 
input_frame = ttk.LabelFrame(root, relief="groove", border=2)
input_frame.place(x=10, y=15, width=500, height=130)

input_frame.grid_columnconfigure(0, weight=1)
input_frame.grid_columnconfigure(1, weight=1)

ttk.Label(input_frame, relief="flat", text="Reference Number:", anchor="w").grid(row=0, column=0, sticky="w", padx=5, pady=0)
reference_number_entry_entry = ttk.Entry(input_frame, width=35)
reference_number_entry_entry.grid(row=1, column=0, padx=5, pady=5)
reference_number_entry_entry.bind("<KeyRelease>", update_note)

ttk.Label(input_frame,relief="flat", text="Site Contact:", anchor="w").grid(row=2, column=0, sticky="w", padx=5, pady=0)
site_contact_entry = ttk.Entry(input_frame, width=35, )
site_contact_entry.grid(row=3, column=0, padx=5, pady=5, )
site_contact_entry.bind("<KeyRelease>", update_note)


ttk.Label(input_frame,relief="flat", text="Contractor:", anchor="w").grid(row=2, column=1, sticky="w", padx=5, pady=0)
contractor_entry = ttk.Entry(input_frame, width=35)
contractor_entry.grid(row=3, column=1, padx=5, pady=5)
contractor_entry.bind("<KeyRelease>", update_note)


ttk.Label(input_frame,relief="flat", text="New Priority?", anchor="w").grid(row=0, column=1, sticky="w", padx=5, pady=0)
priority_var = tk.StringVar()
priority_options = ["", "P1 Emergency", "P2 Immediate", "P3 Urgent", "P3.5 Escalated Routine", "P4 Routine", "P5 Specific Date"]
priority_menu = ttk.Combobox(input_frame, textvariable=priority_var, values=priority_options, width=33, state="readonly")
priority_menu.grid(row=1, column=1, padx=5, pady=5)
priority_menu.bind("<<ComboboxSelected>>", update_note) 


side_frame = ttk.Frame(root)
side_frame.place(x=10, y=150, width=560, height=240)

side_frame.grid_columnconfigure(0, weight=1)

actions_frame = ttk.LabelFrame(side_frame, text="Action" ,relief="groove", borderwidth=2,)
actions_frame.place(x=0, y=0, width=130, height=240)
# ttk.Label(actions_frame, anchor="center").pack(anchor="n", padx=10, pady=5)

def create_action_row(label_text, var):
    checkbox = ttk.Checkbutton(actions_frame, text=label_text, variable=var, command=update_note)
    checkbox.pack(anchor="sw", padx=10, pady=2)

contacted_var = tk.BooleanVar()
dispatched_var = tk.BooleanVar()
emailed_var = tk.BooleanVar()
confirmed_var = tk.BooleanVar()
cancelled_var = tk.BooleanVar()
poc_callback_var = tk.BooleanVar()

create_action_row("Contacted", contacted_var)
create_action_row("Dispatched", dispatched_var)
create_action_row("Emailed", emailed_var)
create_action_row("Confirmed", confirmed_var)
# create_action_row("Cancelled", cancelled_var)
create_action_row("POC Callback", poc_callback_var)

# Cancel heck box.
cancelled_checkbox = ttk.Checkbutton(actions_frame, text="Cancelled", variable=cancelled_var, command=update_note)
cancelled_checkbox.pack(anchor="sw", padx=10, pady=2)

notes_frame = ttk.LabelFrame(root, text="Notes", relief="groove", border=4)
notes_frame.place(x=150, y=150, width=360, height=240,)

notes_frame.grid_columnconfigure(0, weight=1)

reason_note_var = tk.StringVar()
reason_note_var.trace_add("write", update_note)

reason_note_options = [
    "Cancellation Reason: Duplicate Request: <ref-number>",
    "Cancellation Reason: No Longer Required",
    "Cancellation Reason: Site Contact Advised",
    "- - -",
    "Priority Change Reason: Non-Urgent",
    "Priority Change Reason: Site Contact Request",
    "Priority Change Reason: Unable to Allocate Resources.",
    "",

]


# ttk.Label(notes_frame, anchor="center").grid(row=0, column=0, sticky="w", padx=2, pady=2,)

# ttk.Label(notes_frame, text="Notes:", anchor="w").grid(row=2, column=0, sticky="w", padx=2, pady=2,)


note_entry = tk.Text(notes_frame, width=41, height=9, wrap=tk.WORD, relief="flat")
note_entry.grid(row=2, column=0, padx=5, pady=5)
note_entry.bind("<KeyRelease>", update_note)



# Scrolled Text (Fixed Size, when visible)
note_output = scrolledtext.ScrolledText(root, relief="flat", width=65, height=12, wrap=tk.WORD, state="disabled")
note_output.place(x=10, y=380, width=500, height=200)
note_output.place_forget()


# Place a new frame for the buttons at the bottom of the GUI
button_frame = ttk.Frame(root, relief="groove", borderwidth=2)
button_frame.place(relx=0, rely=1.01, anchor="sw", width=520, height=90)

# # Create Preview Checkbox
# preview_var = tk.BooleanVar(value=False)
# preview_checkbox = ttk.Checkbutton(button_frame, text="Preview Note", variable=preview_var, command=toggle_preview)
# preview_checkbox.place(x=25, y=50,)

preview_var = tk.BooleanVar(value=False)
preview_checkbox = ttk.Checkbutton(notes_frame, text="Preview Note", variable=preview_var, command=toggle_preview,)
preview_checkbox.grid(row=3, column=0, sticky="e" , padx=2, pady=2,)

# separator = ttk.Separator(button_frame, orient="horizontal")
# separator.place(width=480,)

# Reposition buttons to this frame
clear_button = ttk.Button(button_frame, text="Clear All", command=clear_all,)
clear_button.place(x=20, y=5)

copy_button = ttk.Button(button_frame, text="Copy", command=copy_to_clipboard)
copy_button.place(x=310, y=5)

save_button = ttk.Button(button_frame, text="Save", command=save_note)
save_button.place(x=410, y=5)
save_button.config(state=tk.DISABLED)


# Add the priority reason label (initially hidden)
cancel_priority_warning_label = ttk.Label(notes_frame, text="State Reason: Cancel / Priority Change", foreground="dark red")
cancel_priority_warning_label.place_forget()  # Hide the label initially

# Success label setup
ref_saved = update_note
success_label = ttk.Label(root, foreground="Black")

# # Reason Note Combobox (Initially Disabled)
# reason_note_menu = ttk.Combobox(notes_frame, textvariable=reason_note_var, values=reason_note_options, width=52, state="disabled")
# reason_note_menu.grid(row=1, column=0, padx=5, pady=5)

def toggle_reason_combobox(*args):
    cancel_selected = cancelled_var.get()
    priority_selected = priority_var.get() not in [""]

    # Disable Cancel if Priority is selected
    if priority_selected:
        cancelled_checkbox.config(state="disabled")
        cancelled_var.set(False)  # Uncheck if previously selected
    else:
        cancelled_checkbox.config(state="normal")

    # Disable Priority if Cancel is selected
    if cancel_selected:
        priority_menu.config(state="disabled")
        priority_var.set("")  # Clear priority selection if Cancel is selected
    else:
        priority_menu.config(state="readonly")

    # Show or hide the reason label based on conditions
    if cancel_selected and priority_selected:
        cancel_priority_warning_label.config(text="Cancel or Priority Change", foreground="dark red")
        cancel_priority_warning_label.place(x=30, y=200, anchor="w")
        reason_note_var.set("")
        reason_note_menu.config(state="disabled")

    elif cancel_selected:
        cancel_priority_warning_label.config(text="Reason For Cancellation", foreground="dark red")
        cancel_priority_warning_label.place(x=30, y=200, anchor="w")
        reason_note_var.set("Cancellation Reason: ")
        reason_note_menu.config(state="normal")
        
    elif priority_selected:
        cancel_priority_warning_label.config(text="Reason For Priority Change", foreground="dark red")
        cancel_priority_warning_label.place(x=30, y=200, anchor="w")
        reason_note_var.set("Priority Change Reason: ")
        reason_note_menu.config(state="normal")
    else:
        cancel_priority_warning_label.place_forget()
        reason_note_menu.config(state="disabled")
        reason_note_var.set("")






# Allow the user to reset the placeholder on selection
def validate_selection(event):
    if reason_note_var.get() == "Select Reason or Type Your Own":
        reason_note_var.set("")
        # reason_note_menu.config(state="disabled")



def open_email_window(event=None):
    window = tk.Toplevel()
    EmailTab(window).pack(expand=True, fill="both")
    window.title("EmailGEN")
    window.geometry("520x860")

# Bind Ctrl+E to open the email window
root.bind("<Control-e>", open_email_window)
root.bind("<Control-E>", open_email_window)




def enable_undo_for_text(widget):
    """Enable undo functionality for Text widgets."""
    widget.config(undo=True)  # Enable the undo option
    widget.bind("<Control-z>", lambda event: safe_edit_undo(widget))
    widget.bind("<Control-y>", lambda event: safe_edit_redo(widget))

def safe_edit_undo(widget):
    """Safely call edit_undo to avoid TclError."""
    try:
        widget.edit_undo()
    except tk.TclError:
        pass  # Ignore if there's nothing to undo

def safe_edit_redo(widget):
    """Safely call edit_redo to avoid TclError."""
    try:
        widget.edit_redo()
    except tk.TclError:
        pass  # Ignore if there's nothing to redo

def enable_undo_for_entry(widget):
    """Enable undo functionality for Entry widgets."""
    undo_stack = [widget.get()]

    def save_state(*args):
        """Save the current state to the undo stack."""
        if not undo_stack or undo_stack[-1] != widget.get():
            undo_stack.append(widget.get())

    def undo_action(event=None):
        """Perform the undo action."""
        if len(undo_stack) > 1:
            undo_stack.pop()
            previous_value = undo_stack[-1]
            widget.delete(0, tk.END)
            widget.insert(0, previous_value)

    # Bind events to track changes and handle undo
    widget.bind("<KeyRelease>", lambda event: save_state())
    widget.bind("<Control-z>", undo_action)


# For Entry widgets
enable_undo_for_entry(reference_number_entry_entry)
enable_undo_for_entry(site_contact_entry)
enable_undo_for_entry(contractor_entry)

# For Text widgets
enable_undo_for_text(note_entry)
enable_undo_for_text(note_output)

# Enable undo for Entry widget
enable_undo_for_entry(reference_number_entry_entry)




# Create the combobox for Reason Note
reason_note_menu = ttk.Combobox(notes_frame, textvariable=reason_note_var, values=reason_note_options, width=52, state="disabled")
reason_note_menu.grid(row=1, column=0, padx=5, pady=5)

# Set the default value to the placeholder
reason_note_menu.set("")

# Bind the combobox to reset placeholder if reselected
reason_note_menu.bind("<<ComboboxSelected>>", validate_selection)

# Attach tracing to monitor changes in priority and cancelled checkbox
cancelled_var.trace_add("write", toggle_reason_combobox)
priority_var.trace_add("write", toggle_reason_combobox)






start_tray_icon()

root.mainloop()













# pyinstaller -w --onefile main.py