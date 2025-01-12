import os
import csv
from tkinter import ttk, messagebox, filedialog
import tkinter as tk
from datetime import datetime


# Function to open the History window
def open_history_window(event=None):
    global search_var, history_tree, selected_item_details_text, history_count_label

    
    history_window = tk.Toplevel()
    history_window.title("History")
    history_window.geometry("675x750")
    history_window.transient()

    # Create a search entry and button in the history window
    search_frame = ttk.Frame(history_window)
    search_frame.pack(pady=5)

    search_label = ttk.Label(search_frame, text="Search:")
    search_label.pack(side="left", padx=5)

    # Use a StringVar to track the search input in real time
    search_var = tk.StringVar()
    search_var.trace("w", lambda name, index, mode: search_history())

    search_entry = ttk.Entry(search_frame, textvariable=search_var, width=30)
    search_entry.pack(side="left", padx=5)

    next_button = ttk.Button(search_frame, text="Next", command=next_result)
    next_button.pack(side="left", padx=5)

    # Scrollbars for the history tree widget
    history_tree_frame = ttk.Frame(history_window)
    history_tree_frame.pack(expand=True, fill="x")

    history_scrollbar_y = ttk.Scrollbar(history_tree_frame, orient="vertical")
    history_scrollbar_y.pack(side="right", fill="y")

    history_scrollbar_x = ttk.Scrollbar(history_tree_frame, orient="horizontal")
    history_scrollbar_x.pack(side="bottom", fill="x")

    # Create Treeview widget
    history_tree = ttk.Treeview(history_tree_frame, height=15, yscrollcommand=history_scrollbar_y.set, xscrollcommand=history_scrollbar_x.set)

    # Configure the scrollbars
    history_scrollbar_y.config(command=history_tree.yview)
    history_scrollbar_x.config(command=history_tree.xview)

    history_tree["show"] = "headings"
    history_tree.pack(expand=True, fill="both")

    # Bind selection event to show details
    history_tree.bind("<<TreeviewSelect>>", show_selected_item_details)

    # Create a Text widget to display the details of the selected item
    selected_item_details_frame = ttk.Frame(history_window)
    selected_item_details_frame.pack(expand=True, fill="both", padx=25, pady=15)

    selected_item_details_text = tk.Text(selected_item_details_frame, wrap="word", height=15, width=80)
    selected_item_details_text.pack(expand=True, fill="x")

    # Create a frame to hold the Export, Load buttons, Edit Record, and History Count label
    button_frame = ttk.Frame(history_window)
    button_frame.pack(pady=10)

    # Create an "Export to CSV" button
    export_button = ttk.Button(button_frame, text="Export to CSV", command=export_csv)
    export_button.pack(side="left", padx=10, pady=10)

    # Create a "Load History" button
    load_button = ttk.Button(button_frame, text="Reload History", command=load_history)
    load_button.pack(side="left", padx=10, pady=10)

    # Add an "Edit Record" button to the button frame
    edit_button = ttk.Button(button_frame, text="Edit Record", command=edit_selected_record)
    edit_button.pack(side="left", padx=10, pady=10)

    # Create a label to display the count of rows in the history
    history_count_label = ttk.Label(button_frame, text="Total Records: 0")
    history_count_label.pack(side="right", padx=10, pady=10)

    # Load history data as soon as the window opens
    load_history_on_open(history_window)
    search_entry.focus_force()

    def close_window(event=None):
        
        history_window.destroy()



    history_window.bind("<Escape>", close_window )


    # history_window.bind("<Enter>", next_result)

# Define all supporting functions (search_history, next_result, etc.)
# Ensure all these functions are included in the file, with required global variables passed or managed.

def search_history():
    global search_var, history_tree, search_results, current_result_index, history_count_label

    # Get the search term from the search entry box
    search_term = search_var.get().strip().lower()

    # Clear existing items in the Treeview
    for item in history_tree.get_children():
        history_tree.delete(item)

    # Initialize the search results
    search_results = []
    current_result_index = -1

    # Reload the data based on search input
    if os.path.exists("notegenhistory.csv"):
        with open("notegenhistory.csv", "r") as csvfile:
            csvreader = csv.DictReader(csvfile)
            headers = csvreader.fieldnames

            if headers:
                # Exclude Actions and Notes from the columns to display
                display_headers = [header for header in headers if header not in ["Actions", "Notes"]]

                # Set Treeview columns
                history_tree["columns"] = display_headers
                for header in display_headers:
                    history_tree.heading(header, text=header)
                    history_tree.column(header, anchor='center', width=100)

                # Insert rows into the Treeview that match the search term
                for row in csvreader:
                    if any(search_term in row[header].lower() for header in headers if row[header]):
                        values = [row[header] for header in display_headers]
                        item_id = history_tree.insert("", "end", values=values)
                        search_results.append(item_id)  # Track matching rows

    # Update the history count label
    history_count_label.config(text=f"Total Records: {len(history_tree.get_children())}")

def next_result():
    global current_result_index
    if search_results:
        current_result_index = (current_result_index + 1) % len(search_results)
        focus_on_result()


# Variables to track search results and index
search_results = []
current_result_index = -1

# Function to remove previous highlights
def remove_highlights():
    # Iterate through all the rows and remove the highlight tag
    for item in history_tree.get_children():
        history_tree.item(item, tags=())  # Removing all tags

# Modified focus function to highlight and scroll to search results
def focus_on_result():
    # Remove previous highlights
    remove_highlights()

    # Check if there are valid search results to highlight
    if search_results and 0 <= current_result_index < len(search_results):
        # Focus on the current result
        history_tree.selection_remove(history_tree.selection())
        result_item = search_results[current_result_index]
        history_tree.selection_add(result_item)
        history_tree.see(result_item)  # Scroll to the item

        # Highlight the result by adding a tag to it
        history_tree.item(result_item, tags=("highlight",))
        history_tree.tag_configure("highlight", background="yellow")

# Updated search_history() to reset and prepare for highlighting



# Function to move to the next search result
def next_result():
    global current_result_index
    if search_results:
        current_result_index = (current_result_index + 1) % len(search_results)
        focus_on_result()




def load_history_on_open(history_window):
    if os.path.exists("notegenhistory.csv"):
        load_history()  # Load history if CSV exists
    else:
        # If CSV does not exist, notify user and prompt to select a file
        response = messagebox.askyesno("File Not Found", "The history CSV file does not exist. Would you like to choose a location to load a CSV?")
        if response:
            file_path = filedialog.askopenfilename(defaultextension=".csv",
                                                   filetypes=[("CSV files", "*.csv")],
                                                   title="Select CSV File")
            if file_path:
                try:
                    with open(file_path, "r") as csvfile:
                        reader = csv.reader(csvfile)
                        with open("notegenhistory.csv", "w", newline='') as outfile:
                            writer = csv.writer(outfile)
                            for row in reader:
                                writer.writerow(row)
                    load_history()  # Load the CSV after it has been copied

                except Exception as e:
                    messagebox.showerror("Error", f"An error occurred while loading the CSV file: {e}")
                    history_tree.focus_set() 
            else:
                messagebox.showinfo("Cancelled", "No file selected. History will not be loaded.")
                history_tree.focus_set() 
        else:
            messagebox.showinfo("Cancelled", "History will not be loaded.")
            history_tree.focus_set() 

# Function to export and save CSV file to selected location
def export_csv():
    # Open a file save dialog to select save location
    file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                             filetypes=[("CSV files", "*.csv")],
                                             title="Save CSV File")
    if file_path:
        try:
            # Copy the CSV data from the current file to the selected location
            with open("notegenhistory.csv", "r") as infile, open(file_path, "w", newline='') as outfile:
                reader = csv.reader(infile)
                writer = csv.writer(outfile)

                # Write each row from infile to outfile
                for row in reader:
                    writer.writerow(row)

            messagebox.showinfo("Export Successful", f"File has been saved successfully to {file_path}")
            history_tree.focus_set() 

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while exporting the CSV file: {e}")
            history_tree.focus_set() 

# Function to edit selected record
def edit_selected_record():


    selected_item = history_tree.focus()
    if not selected_item:
        messagebox.showerror("No Selection", "Please select a record to edit.")
        history_tree.focus_set() 
        return

    # Get the selected record values and match the corresponding Reference Number in the CSV
    item_values = history_tree.item(selected_item, "values")
    headers = history_tree["columns"]

    # Load the complete record from the CSV, including Actions and Notes
    full_record = {}
    if os.path.exists("notegenhistory.csv"):
        with open("notegenhistory.csv", "r") as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                if row["Reference Number"] == item_values[headers.index("Reference Number")]:
                    full_record = row
                    break

    if not full_record:
        messagebox.showerror("Error", "Unable to find the full record for editing.")
        return


    
    # Open the Edit Record window
   
    edit_window = tk.Toplevel()
    edit_window.title("Edit Record")
    edit_window.geometry("530x560")
    edit_window.transient()

    # Frame for the form
    form_frame = ttk.Frame(edit_window, padding=10)
    form_frame.pack(fill=tk.BOTH, expand=True)

    # Input variables and widgets
    input_widgets = {}

    # Reference Number
    ttk.Label(form_frame, text="Reference Number", font=("Arial", 10,)).grid(row=0, column=0, sticky="w", pady=5)
    work_order_var = tk.StringVar(value=full_record.get("Reference Number", "").strip())
    ttk.Entry(form_frame, textvariable=work_order_var, width=50).grid(row=0, column=1, pady=5, padx=10)
    input_widgets["Reference Number"] = work_order_var

    # Site Contact
    ttk.Label(form_frame, text="Site Contact", font=("Arial", 10,)).grid(row=1, column=0, sticky="w", pady=5)
    site_contact_var = tk.StringVar(value=full_record.get("Site Contact", "").strip())
    ttk.Entry(form_frame, textvariable=site_contact_var, width=50).grid(row=1, column=1, pady=5, padx=10)
    input_widgets["Site Contact"] = site_contact_var

    # Contractor
    ttk.Label(form_frame, text="Contractor", font=("Arial", 10,)).grid(row=2, column=0, sticky="w", pady=5)
    contractor_var = tk.StringVar(value=full_record.get("Contractor", "").strip())
    ttk.Entry(form_frame, textvariable=contractor_var, width=50).grid(row=2, column=1, pady=5, padx=10)
    input_widgets["Contractor"] = contractor_var

    # New Priority
    ttk.Label(form_frame, text="New Priority", font=("Arial", 10,)).grid(row=3, column=0, sticky="w", pady=5)
    priority_var = tk.StringVar(value=full_record.get("New Priority", "").strip())
    priority_options = ["", "P1 Emergency", "P2 Immediate", "P3 Urgent", "P3.5 Escalated Routine", "P4 Routine", "P5 Specific Date"]
    ttk.Combobox(form_frame, textvariable=priority_var, values=priority_options, width=47, state="readonly").grid(row=3, column=1, pady=5, padx=10)
    input_widgets["New Priority"] = priority_var

    # Actions (Checkboxes)
    ttk.Label(form_frame, text="Actions", font=("Arial", 10,)).grid(row=4, column=0, sticky="nw", pady=5)
    action_vars = {
        "Contacted": tk.BooleanVar(value="Contacted" in full_record.get("Actions", "")),
        "Dispatched": tk.BooleanVar(value="Dispatched" in full_record.get("Actions", "")),
        "Emailed": tk.BooleanVar(value="Emailed" in full_record.get("Actions", "")),
        "Confirmed": tk.BooleanVar(value="Confirmed" in full_record.get("Actions", "")),
        "Cancelled": tk.BooleanVar(value="Cancelled" in full_record.get("Actions", "")),
        "POC Callback": tk.BooleanVar(value="POC Callback" in full_record.get("Actions", ""))
    }
    actions_frame = ttk.Frame(form_frame)
    actions_frame.grid(row=4, column=1, sticky="w", pady=5, padx=10)
    for idx, (action, var) in enumerate(action_vars.items()):
        ttk.Checkbutton(actions_frame, text=action, variable=var).grid(row=idx, column=0, sticky="w", pady=2)
    input_widgets["Actions"] = action_vars

    # Notes
    ttk.Label(form_frame, text="Notes").grid(row=5, column=0, sticky="nw", pady=5)
    notes_widget = tk.Text(form_frame, height=10, width=40, wrap=tk.WORD)
    notes_widget.insert("1.0", full_record.get("Notes", "").strip())
    notes_widget.grid(row=5, column=1, pady=5, padx=10)
    input_widgets["Notes"] = notes_widget

    # Save the edited record
    def save_edited_record():
        updated_record = {}

        # Retrieve updated values from input widgets
        updated_record["Reference Number"] = work_order_var.get().strip()
        updated_record["Site Contact"] = site_contact_var.get().strip()
        updated_record["Contractor"] = contractor_var.get().strip()
        updated_record["New Priority"] = priority_var.get().strip()

        # Gather actions from checkboxes
        updated_record["Actions"] = ", ".join(action for action, var in action_vars.items() if var.get())

        # Get notes
        updated_record["Notes"] = notes_widget.get("1.0", tk.END).strip()

        # Validation: Ensure Reference Number is not empty
        if not updated_record["Reference Number"]:
            messagebox.showerror("Input Error", "Reference Number is mandatory.")
            return

        # Update the record in the CSV file
        with open("notegenhistory.csv", "r") as infile:
            rows = list(csv.DictReader(infile))

        for row in rows:
            if row["Reference Number"] == full_record["Reference Number"]:
                row.update(updated_record)
                break

        # Write updated rows back to the CSV
        with open("notegenhistory.csv", "w", newline="") as outfile:
            writer = csv.DictWriter(outfile, fieldnames=rows[0].keys())
            writer.writeheader()
            writer.writerows(rows)

        load_history()
        edit_window.destroy()
        messagebox.showinfo("Success", "Record updated successfully.")

    # Buttons
    button_frame = ttk.Frame(edit_window, padding=10)
    button_frame.pack(fill=tk.X)

    ttk.Button(button_frame, text="Save", command=save_edited_record).pack(side=tk.RIGHT, padx=10, pady=5)
    ttk.Button(button_frame, text="Cancel", command=edit_window.destroy).pack(side=tk.LEFT, padx=10, pady=5)


# Function to load history data into the Treeview
def load_history():
    global history_tree, selected_item_details_text, history_count_label

    # Clear existing items in the Treeview
    for item in history_tree.get_children():
        history_tree.delete(item)

    # Set Treeview column headings without Actions and Notes
    if os.path.exists("notegenhistory.csv"):
        with open("notegenhistory.csv", "r") as csvfile:
            csvreader = csv.DictReader(csvfile)
            headers = csvreader.fieldnames

            if headers:
                # Exclude Actions and Notes from the columns to display
                display_headers = [header for header in headers if header not in ["Actions", "Notes"]]

                # Set Treeview columns
                history_tree["columns"] = display_headers
                for header in display_headers:
                    history_tree.heading(header, text=header)
                    history_tree.column(header, anchor='center', width=100)

                # Read data into a list and sort it by date in descending order
                rows = list(csvreader)
                rows = sorted(rows, key=lambda x: datetime.strptime(x["Date Time"], "%d/%m/%Y %H:%M"), reverse=True)

                # Insert rows into the Treeview without Actions and Notes
                for row in rows:
                    values = [row[header] for header in display_headers]
                    history_tree.insert("", "end", values=values)

    # Update the history count label
    history_count_label.config(text=f"Total Records: {len(history_tree.get_children())}")

# Function to display details of the selected item with exact format
def show_selected_item_details(event):
    selected_item = history_tree.focus()  # Get the selected item in the Treeview
    if selected_item:
        item_values = history_tree.item(selected_item, "values")
        headers = history_tree["columns"]

        # Extract the values that are displayed in the Treeview
        work_order = item_values[headers.index("Reference Number")].strip() if "Reference Number" in headers else ""
        new_priority = item_values[headers.index("New Priority")].strip() if "New Priority" in headers else ""
        site_contact = item_values[headers.index("Site Contact")].strip() if "Site Contact" in headers else ""
        contractor = item_values[headers.index("Contractor")].strip() if "Contractor" in headers else ""

        # Load Actions and Notes separately from the CSV since they are not displayed in Treeview
        actions = ""
        notes = ""
        if work_order:  # Proceed only if there's a valid Reference Number to match
            with open("notegenhistory.csv", "r") as csvfile:
                csvreader = csv.DictReader(csvfile)
                for row in csvreader:
                    if row["Reference Number"].strip() == work_order:
                        actions = row.get("Actions", "").strip()
                        notes = row.get("Notes", "").strip()
                        break

        # Start formatting details, only include non-empty values
        formatted_details = ""

        if work_order:
            formatted_details += f"Reference Number:     {work_order}\n"
        if new_priority:
            formatted_details += f"New Priority:   {new_priority}\n"
        if site_contact:
            formatted_details += f"Site Contact:   {site_contact}\n"
        if contractor:
            formatted_details += f"Contractor:     {contractor}\n"

        if actions:
            action_items = actions.split(", ")
            formatted_details += "\nAction:\n"
            for action in ["Contacted", "Dispatched", "Emailed", "Confirmed", "Cancelled", "POC Callback"]:
                if action in action_items:
                    formatted_details += f"   - {action}:\n"

        if notes:
            formatted_details += f"\nNotes:\n  - {notes}\n"

        # Clear existing content and insert formatted details
        selected_item_details_text.delete("1.0", tk.END)
        selected_item_details_text.insert("1.0", formatted_details)


