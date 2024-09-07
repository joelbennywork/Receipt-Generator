import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from docx import Document
from reportlab.pdfgen import canvas
import pywhatkit as kit
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os

# Load data from Excel
def load_data():
    df = pd.read_excel('invoice.xlsx')
    return df

# Save data to Excel
def save_data(df):
    df.to_excel('invoice.xlsx', index=False)

# Generate receipt
def generate_receipt(dataframe):
    #Implement generate invoice functionality
    pass

# Send to phone
def send_to_phone():
    # Implement send to phone functionality
    pass

# Send to email
def send_to_email():
    # Implement send to email functionality
    pass

df = load_data()

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

        generate_btn = tk.Button(toolbar, text="Generate Receipts", command=lambda: generate_receipt(df))
        generate_btn.pack(side=tk.LEFT, padx=2, pady=2)

        send_phone_btn = tk.Button(toolbar, text="Send to Phone")
        send_phone_btn.pack(side=tk.LEFT, padx=2, pady=2)

        send_email_btn = tk.Button(toolbar, text="Send to Email")
        send_email_btn.pack(side=tk.LEFT, padx=2, pady=2)

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
