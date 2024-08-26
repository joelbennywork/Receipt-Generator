import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from docx import Document
from reportlab.pdfgen import canvas
import pywhatkit as kit
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Load data from Excel
def load_data():
    df = pd.read_excel('invoice.xlsx')
    return df

# Save data to Excel
def save_data(df):
    df.to_excel('invoice.xlsx', index=False)

# Add row
def add_row():
    def submit():
        # Get data from entry fields
        first_name = entry_first_name.get()
        last_name = entry_last_name.get()
        phone_number = entry_phone_number.get()
        email = entry_email.get()
        item1 = entry_item1.get()
        item2 = entry_item2.get()
        total_price = entry_total_price.get()

        # Insert data into the table
        tree.insert("", "end", values=(first_name, last_name, phone_number, email, item1, item2, total_price))

        # Update the Excel file
        df = load_data()
        new_row = {
            "First Name": first_name,
            "Last Name": last_name,
            "Phone Number": phone_number,
            "Email": email,
            "Item 1": item1,
            "Item 2": item2,
            "Total Price": total_price
        }
        df = df.append(new_row, ignore_index=True)
        save_data(df)

        # Close the add row window
        add_window.destroy()

    # Create a new window for adding a row
    add_window = tk.Toplevel(root)
    add_window.title("Add Row")

    # Create entry fields for each column
    tk.Label(add_window, text="First Name").grid(row=0, column=0)
    entry_first_name = tk.Entry(add_window)
    entry_first_name.grid(row=0, column=1)

    tk.Label(add_window, text="Last Name").grid(row=1, column=0)
    entry_last_name = tk.Entry(add_window)
    entry_last_name.grid(row=1, column=1)

    tk.Label(add_window, text="Phone Number").grid(row=2, column=0)
    entry_phone_number = tk.Entry(add_window)
    entry_phone_number.grid(row=2, column=1)

    tk.Label(add_window, text="Email").grid(row=3, column=0)
    entry_email = tk.Entry(add_window)
    entry_email.grid(row=3, column=1)

    tk.Label(add_window, text="Item 1").grid(row=4, column=0)
    entry_item1 = tk.Entry(add_window)
    entry_item1.grid(row=4, column=1)

    tk.Label(add_window, text="Item 2").grid(row=5, column=0)
    entry_item2 = tk.Entry(add_window)
    entry_item2.grid(row=5, column=1)

    tk.Label(add_window, text="Total Price").grid(row=6, column=0)
    entry_total_price = tk.Entry(add_window)
    entry_total_price.grid(row=6, column=1)

    # Create a submit button
    submit_button = tk.Button(add_window, text="Submit", command=submit)
    submit_button.grid(row=7, columnspan=2)

# Add the add_row function to the Add Row button in the toolbar
add_btn.config(command=add_row)


# Delete row
def delete_row():
    # Implement delete row functionality
    pass

# Edit row
def edit_row():
    # Implement edit row functionality
    pass

# Generate receipt
def generate_receipt():
    # Implement receipt generation functionality
    pass

# Send to phone
def send_to_phone():
    # Implement send to phone functionality
    pass

# Send to email
def send_to_email():
    # Implement send to email functionality
    pass

# Preview receipt
def preview_receipt():
    # Implement preview receipt functionality
    pass

# Main application
class ReceiptGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Receipt Generator")
        self.create_widgets()

    def create_widgets(self):
        # Create toolbar
        toolbar = tk.Frame(self.root)
        toolbar.pack(side=tk.TOP, fill=tk.X)

        add_btn = tk.Button(toolbar, text="Add Row", command=add_row)
        add_btn.pack(side=tk.LEFT, padx=2, pady=2)

        delete_btn = tk.Button(toolbar, text="Delete Row", command=delete_row)
        delete_btn.pack(side=tk.LEFT, padx=2, pady=2)

        edit_btn = tk.Button(toolbar, text="Edit Row", command=edit_row)
        edit_btn.pack(side=tk.LEFT, padx=2, pady=2)

        generate_btn = tk.Button(toolbar, text="Generate Receipts", command=generate_receipt)
        generate_btn.pack(side=tk.LEFT, padx=2, pady=2)

        send_phone_btn = tk.Button(toolbar, text="Send to Phone", command=send_to_phone)
        send_phone_btn.pack(side=tk.LEFT, padx=2, pady=2)

        send_email_btn = tk.Button(toolbar, text="Send to Email", command=send_to_email)
        send_email_btn.pack(side=tk.LEFT, padx=2, pady=2)

        preview_btn = tk.Button(toolbar, text="Preview Receipt", command=preview_receipt)
        preview_btn.pack(side=tk.LEFT, padx=2, pady=2)

        # Create table
        self.tree = ttk.Treeview(self.root, columns=("First Name", "Last Name", "Phone Number", "Email", "Item 1", "Item 2", "Total Price"), show='headings')
        self.tree.pack(fill=tk.BOTH, expand=True)

        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)

        # Load data into table
        self.load_table_data()

    def load_table_data(self):
        df = load_data()
        for index, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

if __name__ == "__main__":
    root = tk.Tk()
    app = ReceiptGeneratorApp(root)
    root.mainloop()
