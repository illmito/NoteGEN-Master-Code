import win32com.client
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Text
import os

class EmailTab(ttk.Frame):

    def __init__(self, parent):
        super().__init__(parent)
        self.attachments_list = []
        self.task_widgets = {}

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)


        # self.task_widgets['cc_entry'] = ""


        # Task Selection Frame
        self.task_selection_frame = ttk.LabelFrame(self, text="Task Selection", relief="groove", borderwidth=2)
        self.task_selection_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nwe", columnspan=1)

        ttk.Label(self.task_selection_frame, text="Task:").grid(row=0, column=0, sticky="w", padx=10, pady=10)
        self.task_var = tk.StringVar()
        self.task_combobox = ttk.Combobox(
            self.task_selection_frame, textvariable=self.task_var, state="readonly",
            values=["Leased Property", "Waste Management", "Cancellation Request"], width=45
        )
        self.task_combobox.grid(row=0, column=1, padx=10, pady=10)
        self.task_combobox.set("Select Template")
        self.task_combobox.bind("<<ComboboxSelected>>", self.display_task_widgets)

        # Task Widget Frame
        self.task_frame = ttk.LabelFrame(self, text="Email Template", relief="groove", borderwidth=2, )
        self.task_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nwe", columnspan=1)

        # # Attachment Frame
        # self.attachment_frame = ttk.LabelFrame(self, text="Attachments", relief="groove", borderwidth=2)
        # self.attachment_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nwe")

        # self.attachment_label = ttk.Label(self.attachment_frame, text="No attachments added.", anchor="w")
        # self.attachment_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        # self.attach_button = ttk.Button(self.attachment_frame, text="Add Attachment", command=self.add_attachment)
        # self.attach_button.grid(row=1, column=0, padx=5, pady=10, sticky="w")

        # Button Frame
        self.button_frame = ttk.Frame(self)
        self.button_frame.grid(row=3, column=0, padx=10, pady=20, sticky="e",)

        # self.ready_button = ttk.Button(self.button_frame, text="Ready")
 

        # self.grid_columnconfigure(0, weight=1)
        # self.grid_rowconfigure(0, weight=1)

    def add_attachment(self):
        selected_file = filedialog.askopenfilename(title="Select File to Attach")
        if selected_file:
            self.attachments_list.append(selected_file)
            self.attachment_label.config(text="\n".join(self.attachments_list))

    def display_task_widgets(self, event=None):
        task = self.task_var.get()
        for widget in self.task_widgets.values():
            widget.destroy()
        self.task_widgets.clear()

        if task == "Leased Property":
            self.create_leased_property_widgets()
        elif task == "Waste Management":
            self.create_waste_management_widgets()
        elif task == "Blank Email":
            self.create_blank_email_widgets()
        elif task == "Cancellation Request":
            self.create_cancellation_request_widgets()






    def create_cancellation_request_widgets(self):
        # Clear existing widgets
        for widget in self.task_widgets.values():
            widget.destroy()
        self.task_widgets.clear()
        

        # To
        self.task_widgets['to_label'] = ttk.Label(self.task_frame, text="To:",)
        self.task_widgets['to_label'].grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.task_widgets['to_entry'] = ttk.Entry(self.task_frame, width=40,)
        self.task_widgets['to_entry'].insert(0, "BSSC@ventia.com")
        self.task_widgets['to_entry'].config(state="disabled")
        self.task_widgets['to_entry'].grid(row=0, column=1, padx=10, pady=5, sticky="W")




        # CC
        self.task_widgets['cc_label'] = ttk.Label(self.task_frame, text="CC:",)
        self.task_widgets['cc_label'].grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.task_widgets['cc_entry'] = ttk.Entry(self.task_frame, width=40)

        self.task_widgets['cc_entry'].grid(row=1, column=1, padx=10, pady=5, sticky="w")

        # Subject
        # self.task_widgets['subject_label'] = ttk.Label(self.task_frame, text="Subject:")
        # self.task_widgets['subject_label'].grid(row=2, column=0, padx=10, pady=5, sticky="w")
        # self.task_widgets['subject_entry'] = ttk.Entry(self.task_frame, width=50, state="disabled")
        # self.task_widgets['subject_entry'].grid(row=2, column=1, padx=10, pady=5)




        # Service Line
        self.task_widgets['service_label'] = ttk.Label(self.task_frame, text="Service Line:")
        self.task_widgets['service_label'].grid(row=2, column=0, padx=10, pady=5, sticky="w")

        self.task_widgets['service_entry'] = ttk.Combobox(
            self.task_frame,
            values=["EU", "WM", "CL", "PV", "HC", "TS", "HK", "LM", "RS", "AC", "SR"],
            state="readonly",
            width=25,
        )
        self.task_widgets['service_entry'].set("Select")
        self.task_widgets['service_entry'].grid(row=2, column=1, padx=10, pady=5, sticky="w")

        # Priority
        self.task_widgets['priority_label'] = ttk.Label(self.task_frame, text="Priority:")
        self.task_widgets['priority_label'].grid(row=3, column=0, padx=10, pady=5, sticky="w")

        self.task_widgets['priority_entry'] = ttk.Combobox(
            self.task_frame,
            values=["P1", "P2", "P3", "P3.5", "P4", "P5"],
            state="readonly",
            width=25,
        )
        self.task_widgets['priority_entry'].set("Select")
        self.task_widgets['priority_entry'].grid(row=3, column=1, padx=10, pady=5, sticky="w")

        # SAP Work Order / Notification Number
        self.task_widgets['sap_label'] = ttk.Label(self.task_frame, text="WO / Notification:")
        self.task_widgets['sap_label'].grid(row=4, column=0, padx=10, pady=5, sticky="w")

        self.task_widgets['sap_entry'] = ttk.Entry(self.task_frame, width=40)
        self.task_widgets['sap_entry'].grid(row=4, column=1, padx=10, pady=5, sticky="w")

        # Requested By
        self.task_widgets['requested_label'] = ttk.Label(self.task_frame, text="Requested By:")
        self.task_widgets['requested_label'].grid(row=5, column=0, padx=10, pady=5, sticky="w")

        self.task_widgets['requested_entry'] = ttk.Combobox(
            self.task_frame,
            values=["POC: ", "SS: ", "Contractor: "],
            width=25,
        )
        self.task_widgets['requested_entry'].set("Select")
        self.task_widgets['requested_entry'].grid(row=5, column=1, padx=10, pady=5, sticky="w")
        

        # Create a hidden label for the message
        self.task_widgets['requested_warning_label'] = tk.Label(
            self.task_frame, 
            text="*Input Name of Person", 
            fg="dark red",  # Set text color to dark red
        )
        self.task_widgets['requested_warning_label'].grid(row=5, column=1, padx=1, pady=5, sticky="e")
        self.task_widgets['requested_warning_label'].grid_remove()  # Hide initially

        # Function to handle selection changes in the combobox
        def on_supervisor_selection(event):
            selection = self.task_widgets['requested_entry'].get()
            if selection:
                self.task_widgets['requested_warning_label'].grid()  # Show the warning label
            else:
                self.task_widgets['requested_warning_label'].grid_remove()  # Hide the warning label

        # Attach the event to the combobox
        self.task_widgets['requested_entry'].bind("<<ComboboxSelected>>", on_supervisor_selection)




        # Reason for Cancellation
        self.task_widgets['reason_label'] = ttk.Label(self.task_frame, text="Reason for Cancellation:")
        self.task_widgets['reason_label'].grid(row=6, column=0, padx=10, pady=5, sticky="nw")

        self.task_widgets['reason_entry'] = tk.Text(self.task_frame, height=4, width=35)
        self.task_widgets['reason_entry'].grid(row=6, column=1, padx=10, pady=5, sticky="w")

        # Site Supervisor Notified
        self.task_widgets['supervisor_label'] = ttk.Label(self.task_frame, text="Site Supervisor Notified:")
        self.task_widgets['supervisor_label'].grid(row=7, column=0, padx=10, pady=5, sticky="w")

        self.task_widgets['supervisor_combobox'] = ttk.Combobox(
            self.task_frame,
            values=["YES", "N/A"],
            state="normal",  # Allow text entry
            width=25,
        )
        self.task_widgets['supervisor_combobox'].set("Select")
        self.task_widgets['supervisor_combobox'].grid(row=7, column=1, padx=10, pady=5, sticky="w")



        # Create a hidden label for the message
        self.task_widgets['supervisor_warning_label'] = tk.Label(
            self.task_frame, 
            text="*Input Site Supervisor name", 
            fg="dark red",  # Set text color to dark red
        )
        self.task_widgets['supervisor_warning_label'].grid(row=7, column=1, padx=1, pady=5, sticky="e")
        self.task_widgets['supervisor_warning_label'].grid_remove()  # Hide initially

        # Function to handle selection changes in the combobox
        def on_supervisor_selection(event):
            selection = self.task_widgets['supervisor_combobox'].get()
            if selection == "YES":
                self.task_widgets['supervisor_warning_label'].grid()  # Show the warning label
            else:
                self.task_widgets['supervisor_warning_label'].grid_remove()  # Hide the warning label

        # Attach the event to the combobox
        self.task_widgets['supervisor_combobox'].bind("<<ComboboxSelected>>", on_supervisor_selection)


        # POC Notified By
        self.task_widgets['poc_label'] = ttk.Label(self.task_frame, text="POC Notified By:")
        self.task_widgets['poc_label'].grid(row=8, column=0, padx=10, pady=5, sticky="w")

        self.task_widgets['poc_combobox'] = ttk.Combobox(
            self.task_frame,
            values=["Phone", "Email", "N/A"],
            state="readonly",
            width=25,
        )
        self.task_widgets['poc_combobox'].set("Select")
        self.task_widgets['poc_combobox'].grid(row=8, column=1, padx=10, pady=5, sticky="w")

        # Contractor Notified
        self.task_widgets['contractor_label'] = ttk.Label(self.task_frame, text="Contractor Notified:")
        self.task_widgets['contractor_label'].grid(row=9, column=0, padx=10, pady=5, sticky="w")

        self.task_widgets['contractor_combobox'] = ttk.Combobox(
            self.task_frame,
            values=["YES", "N/A"],
            state="readonly",
            width=25,
        )
        self.task_widgets['contractor_combobox'].set("Select")
        self.task_widgets['contractor_combobox'].grid(row=9, column=1, padx=10, pady=5, sticky="w")

        # Request Re-Raised
        self.task_widgets['re_raised_label'] = ttk.Label(self.task_frame, text="Request Re-Raised:")
        self.task_widgets['re_raised_label'].grid(row=10, column=0, padx=10, pady=5, sticky="w")

        self.task_widgets['re_raised_entry'] = ttk.Entry(self.task_frame, width=40)
        self.task_widgets['re_raised_entry'].grid(row=10, column=1, padx=10, pady=5, sticky="w")

        # Notes Added to SAP
        self.task_widgets['sap_notes_label'] = ttk.Label(self.task_frame, text="Notes Added to SAP:")
        self.task_widgets['sap_notes_label'].grid(row=11, column=0, padx=10, pady=5, sticky="w")

        self.task_widgets['sap_notes_combobox'] = ttk.Combobox(
            self.task_frame,
            values=["YES", "NO"],
            state="readonly",
            width=10,
        )
        self.task_widgets['sap_notes_combobox'].set("Select")
        self.task_widgets['sap_notes_combobox'].grid(row=11, column=1, padx=10, pady=5, sticky="w")

        # Cancel Email Button
        self.ready_button = ttk.Button(
            self.button_frame, 
            text="Ready", 
            command=lambda: self.format_email("cancel"),
        )
        self.ready_button.grid(row=0, column=0, padx=10, pady=5, sticky="e")



           


    def create_leased_property_widgets(self):
        # Clear existing widgets
        for widget in self.task_widgets.values():
            widget.destroy()
        self.task_widgets.clear()

        # To
        self.task_widgets['to_label'] = ttk.Label(self.task_frame, text="To:")
        self.task_widgets['to_entry'] = ttk.Entry(self.task_frame, width=50)
        self.task_widgets['to_entry'].insert(0, "bssc@ventia.com")  # Default email
        self.task_widgets['to_entry'].config(state="disabled")
        self.task_widgets['to_label'].grid(row=0, column=0, sticky="w",  padx=10, pady=5, )
        self.task_widgets['to_entry'].grid(row=0, column=1, padx=10, pady=5, sticky="w")

        # CC
        self.task_widgets['cc_label'] = ttk.Label(self.task_frame, text="CC:")
        self.task_widgets['cc_label'].grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.task_widgets['cc_entry'] = ttk.Entry(self.task_frame, width=50)
        self.task_widgets['cc_entry'].grid(row=1, column=1, padx=10, pady=5, sticky="w")

        # # Subject
        # self.task_widgets['subject_label'] = ttk.Label(self.task_frame, text="Subject:")
        # self.task_widgets['subject_label'].grid(row=2, column=0, padx=10, pady=5, sticky="w")
        # self.task_widgets['subject_entry'] = ttk.Entry(self.task_frame, width=50)
        # self.task_widgets['subject_entry'].grid(row=2, column=1, padx=10, pady=5)

        # SAP Notification Number
        self.task_widgets['sap_label'] = ttk.Label(self.task_frame, text="Notification Number:")
        self.task_widgets['sap_label'].grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.task_widgets['sap_entry'] = ttk.Entry(self.task_frame, width=40)
        self.task_widgets['sap_entry'].grid(row=3, column=1, padx=10, pady=5, sticky="w")

        # Priority
        self.task_widgets['priority_label'] = ttk.Label(self.task_frame, text="Priority:")
        self.task_widgets['priority_label'].grid(row=4, column=0, padx=10, pady=5, sticky="w")

        self.task_widgets['priority_entry'] = ttk.Combobox(
            self.task_frame,
            values=["P1", "P2", "P3", "P3.5", "P4", "P5"],
            state="readonly",
            width=20,
        )
        self.task_widgets['priority_entry'].set("Select")
        self.task_widgets['priority_entry'].grid(row=4, column=1, padx=10, pady=5, sticky="w")


        # Long Text Template
        self.task_widgets['long_text_label'] = ttk.Label(self.task_frame, text="Long Text:")
        self.task_widgets['long_text_label'].grid(row=5, column=0, padx=5, pady=5, sticky="ns",columnspan=2)
        self.task_widgets['reason_entry'] = tk.Text(self.task_frame, height=25, width=60, relief="flat")
        self.task_widgets['reason_entry'].grid(row=6, column=0, padx=5, pady=5, sticky="ew",columnspan=2)



        # self.task_widgets['long_text_label'] = ttk.Label(self.task_frame, text="Long Text:")
        # self.task_widgets['long_text_label'].grid(row=5, column=0, padx=10, pady=5, sticky="w")
        # self.task_widgets['reason_entry'] = tk.Text(self.task_frame, height=30, width=50,)
        # self.task_widgets['reason_entry'].grid(row=5, column=1, padx=10, pady=5, sticky="w",)

        # Button to format and generate the email
        self.ready_button = ttk.Button(
            self.button_frame,
            text="Ready",
            command=lambda: self.format_email("leased")
        )
        self.ready_button.grid(row=0, column=0, padx=10, pady=10, sticky="e")





    def create_waste_management_widgets(self):
        # Clear existing widgets
        for widget in self.task_widgets.values():
            widget.destroy()
        self.task_widgets.clear()
        

        # To
        self.task_widgets['to_label'] = ttk.Label(self.task_frame, text="To:")
        self.task_widgets['to_label'].grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.task_widgets['to_entry'] = ttk.Entry(self.task_frame, width=50)
        self.task_widgets['to_entry'].insert(0, "BSSC@ventia.com")
        self.task_widgets['to_entry'].config(state="disabled")  # Default email for waste management
        self.task_widgets['to_entry'].grid(row=0, column=1, padx=10, pady=5, sticky="w")

        # CC
        self.task_widgets['cc_label'] = ttk.Label(self.task_frame, text="CC:")
        self.task_widgets['cc_label'].grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.task_widgets['cc_entry'] = ttk.Entry(self.task_frame, width=50)
        self.task_widgets['cc_entry'].insert(0, "Chloe.Barrett@ventia.com") 
        self.task_widgets['cc_entry'].config(state="disabled") # Default CC
        self.task_widgets['cc_entry'].grid(row=1, column=1, padx=10, pady=5, sticky="w")

        # # Subject
        # self.task_widgets['subject_label'] = ttk.Label(self.task_frame, text="Subject:")
        # self.task_widgets['subject_label'].grid(row=2, column=0, padx=10, pady=5, sticky="w")
        # self.task_widgets['subject_entry'] = ttk.Entry(self.task_frame, width=50)
        # self.task_widgets['subject_entry'].grid(row=2, column=1, padx=10, pady=5)

        # SAP Notification Number
        self.task_widgets['sap_label'] = ttk.Label(self.task_frame, text="Notification Number:")
        self.task_widgets['sap_label'].grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.task_widgets['sap_entry'] = ttk.Entry(self.task_frame, width=40)
        self.task_widgets['sap_entry'].grid(row=3, column=1, padx=10, pady=5, sticky="w")

        # Short Text
        self.task_widgets['short_text_label'] = ttk.Label(self.task_frame, text="Short Text")
        self.task_widgets['short_text_label'].grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.task_widgets['short_text_entry'] = ttk.Entry(self.task_frame, width=40)
        self.task_widgets['short_text_entry'].grid(row=4, column=1, padx=10, pady=5, sticky="w")

        # Priority

    
        self.task_widgets['priority_label'] = ttk.Label(self.task_frame, text="Priority:")
        self.task_widgets['priority_label'].grid(row=5, column=0, padx=10, pady=5, sticky="w")

        self.task_widgets['priority_entry'] = ttk.Combobox(
            self.task_frame,
            values=["P1", "P2", "P3", "P3.5", "P4", "P5"],
            state="readonly",
            width=20,
        )
        self.task_widgets['priority_entry'].set("Select")
        self.task_widgets['priority_entry'].grid(row=5, column=1, padx=10, pady=5, sticky="w")



        # self.task_widgets['priority_label'] = ttk.Label(self.task_frame, text="Priority:")
        # self.task_widgets['priority_label'].grid(row=5, column=0, padx=10, pady=5, sticky="w")
        # self.task_widgets['priority_entry'] = ttk.Entry(self.task_frame, width=40)
        # self.task_widgets['priority_entry'].grid(row=5, column=1, padx=10, pady=5, sticky="w")

        # Called Through
        self.task_widgets['called_through_label'] = ttk.Label(self.task_frame, text="Called Through?")
        self.task_widgets['called_through_label'].grid(row=6, column=0, padx=10, pady=5, sticky="w")
        self.task_widgets['called_through_combobox'] = ttk.Combobox(
            self.task_frame, state="readonly", values=["Yes", "No"], width=20)
        self.task_widgets['called_through_combobox'].set("Select")
        self.task_widgets['called_through_combobox'].grid(row=6, column=1, padx=10, pady=5, sticky="w")

        # Long Text (Body Details)
        # Long Text Template
        self.task_widgets['long_text_label'] = ttk.Label(self.task_frame, text="Long Text:")
        self.task_widgets['long_text_label'].grid(row=7, column=0, padx=5, pady=5, sticky="ns",columnspan=2)
        self.task_widgets['reason_entry'] = tk.Text(self.task_frame, height=25, width=60, relief="flat")
        self.task_widgets['reason_entry'].grid(row=8, column=0, padx=5, pady=5, sticky="ew",columnspan=2)





        # Button to format the email
        self.ready_button = ttk.Button(
            self.button_frame,
            text="Ready",
            command=lambda: self.format_email("waste")
        )
        self.ready_button.grid(row=0, column=0, padx=10, pady=10, sticky="e")

        


    def create_blank_email_widgets(self):
        for widget in self.task_widgets.values():
            widget.destroy()
        self.task_widgets.clear()

        self.task_widgets['to_label'] = ttk.Label(self.task_frame, text="To:")
        self.task_widgets['to_label'].grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.task_widgets['to_entry'] = ttk.Entry(self.task_frame, width=50)
        self.task_widgets['to_entry'].grid(row=0, column=1, padx=10, pady=5, sticky="w")

        self.task_widgets['cc_label'] = ttk.Label(self.task_frame, text="CC:")
        self.task_widgets['cc_label'].grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.task_widgets['cc_entry'] = ttk.Entry(self.task_frame, width=50)
        self.task_widgets['cc_entry'].grid(row=1, column=1, padx=10, pady=5)

        self.task_widgets['subject_label'] = ttk.Label(self.task_frame, text="Subject:")
        self.task_widgets['subject_label'].grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.task_widgets['subject_entry'] = ttk.Entry(self.task_frame, width=50)
        self.task_widgets['subject_entry'].grid(row=2, column=1, padx=10, pady=5)

        self.ready_button = ttk.Button(self.button_frame, text="Ready", command=lambda: self.format_email("blank"))
        self.ready_button.grid(row=0, column=0, padx=10, pady=10, sticky="e")


    def format_email(self, task_type=None):
        subject = ""
        body = ""

        if task_type == "cancel":
            service = self.task_widgets['service_entry'].get()
            sap = self.task_widgets['sap_entry'].get()
            priority = self.task_widgets['priority_entry'].get()
            requested_by = self.task_widgets['requested_entry'].get()
            reason = self.task_widgets['reason_entry'].get("1.0", tk.END).strip()
            supervisor = self.task_widgets['supervisor_combobox'].get()
            poc = self.task_widgets['poc_combobox'].get()
            contractor = self.task_widgets['contractor_combobox'].get()
            re_raised = self.task_widgets['re_raised_entry'].get()
            notes = self.task_widgets['sap_notes_combobox'].get()

            subject = f"Cancel {service} - {sap} - {priority}"
            body = (f"Hi Team,\n\nPlease cancel {service}, {sap}, {priority}.\n\n"
                    f"Requested by: {requested_by}\n"
                    f"Reason given for cancellation request: {reason}\n"
                    f"Site Supervisor that has been notified: {supervisor}\n"
                    f"POC has been notified by: {poc}\n"
                    f"Contractor has been notified: {contractor}\n"
                    f"Request re-raised: {re_raised}\n"
                    f"Notes about cancellation have been added to SAP: {notes}")

        elif task_type == "leased":
            sap = self.task_widgets['sap_entry'].get()
            priority = self.task_widgets['priority_entry'].get()
            long_text = self.task_widgets['reason_entry'].get("1.0", tk.END).strip()

            subject = f"EU Leased Property - {sap} - {priority}"
            body = (f"Hi Team,\n\n"
                    f"Please see {priority} Estate Upkeep Leased Property request.\n"
                    f"Notification Number: {sap}\n\n"
                    f"*****************\n{long_text}\n********************")

        elif task_type == "waste":
            short_text = self.task_widgets['short_text_entry'].get()
            sap = self.task_widgets['sap_entry'].get()
            priority = self.task_widgets['priority_entry'].get()
            called_through = self.task_widgets['called_through_combobox'].get()
            long_text = self.task_widgets['reason_entry'].get("1.0", tk.END).strip()

            subject = f"{short_text} - {sap} - {priority}"
            body = (f"Hi Team,\n\n"
                    f"Please see {priority} Waste Management request.\n"
                    f"Notification Number: {sap}\n"
                    f"Has this request been called through? {called_through}\n\n"
                    f"*****************\n{long_text}\n********************")

        elif task_type == "blank":
            subject = self.task_widgets['subject_entry'].get()
            body = "This is a blank email."

        # Pass the formatted email to the create_email function
        self.create_email(subject, body)

    
    
    # clearing high to type into box
    def clear_highlight(widget):
        """Clears text selection and moves the cursor to the end for the given widget."""
        # Get the widget's internal entry (focus_get() ensures we target the right element)
        entry_widget = widget.master.focus_get()
        if entry_widget and hasattr(entry_widget, 'selection_clear'):
            entry_widget.selection_clear()  # Clear any text selection
            entry_widget.icursor("end")  # Move the cursor to the end of the text







    def get_default_signature(self):
        """Retrieve the default Outlook email signature."""
        appdata = os.getenv('APPDATA')
        signature_dir = os.path.join(appdata, "Microsoft", "Signatures")

        if not os.path.exists(signature_dir):
            return ""

        for file in os.listdir(signature_dir):
            if file.endswith(".htm"):
                with open(os.path.join(signature_dir, file), 'r', encoding='utf-8') as f:
                    return f.read()
        return ""    


    def append_signature_to_body(self, body):
        """Append the default Outlook signature to the email body."""
        signature = self.get_default_signature()
        # Convert plain text body to HTML
        body_html = body.replace("\n", "<br>")  # Replace newlines with HTML line breaks
        return f"<html><body>{body_html}<br>{signature}</body></html>"    

    def create_email(self, subject, body):
        try:
            formatted_body = self.append_signature_to_body(body)

            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)

            # Retrieve recipient details
            to_address = self.task_widgets['to_entry'].get().strip()
            cc_email = self.task_widgets['cc_entry'].get().strip()

            if not to_address:
                messagebox.showerror("Error", "Recipient email address is required.")
                return

            mail.To = to_address
            mail.CC = cc_email
            mail.Subject = subject

            # Add formatted body with signature
            mail.HTMLBody = formatted_body

            # Display the email for review
            mail.Display()

            messagebox.showinfo("Success", "Email ready in Outlook.")
            self.task_selection_frame.focus_force()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create email: {str(e)}")
