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
def add_row(self):
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
        self.tree.insert("", "end", values=(first_name, last_name, phone_number, email, item1, item2, total_price))

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

        # Append the new row to the DataFrame
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        save_data(df)

        # Close the add row window
        add_window.destroy()

    # Create a new window for adding a row
    add_window = tk.Toplevel(self.root)
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

# Delete row
def delete_row():
    selected_item = app.tree.selection()  # Get selected item
    if not selected_item:
        messagebox.showwarning("Warning", "Please select a row to delete")
        return

    # Confirm deletion
    confirm = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete the selected row?")
    if not confirm:
        return

    # Get the values of the selected item
    item_values = app.tree.item(selected_item, "values")

    # Load the current data from the Excel file
    df = load_data()

    # Find the row in the DataFrame that matches the selected item values
    row_to_delete = df[(df['First Name'] == item_values[0]) & 
                       (df['Last Name'] == item_values[1]) & 
                       (df['Phone Number'] == item_values[2]) & 
                       (df['Email'] == item_values[3]) & 
                       (df['Item 1'] == item_values[4]) & 
                       (df['Item 2'] == item_values[5]) & 
                       (df['Total Price'] == item_values[6])]

    # Drop the row from the DataFrame
    df = df.drop(row_to_delete.index)

    # Save the updated DataFrame back to the Excel file
    try:
        save_data(df)
        messagebox.showinfo("Success", "Row deleted successfully from Excel file")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save changes to the Excel file: {e}")
        return

    # Remove the item from the Treeview
    app.tree.delete(selected_item)



# Edit row
def edit_row(self):
    selected_item = self.tree.selection()[0]
    values = self.tree.item(selected_item, "values")

    def submit():
        # Get data from entry fields
        first_name = entry_first_name.get()
        last_name = entry_last_name.get()
        phone_number = entry_phone_number.get()
        email = entry_email.get()
        item1 = entry_item1.get()
        item2 = entry_item2.get()
        total_price = entry_total_price.get()

        # Update the selected row in the table
        self.tree.item(selected_item, values=(first_name, last_name, phone_number, email, item1, item2, total_price))

        # Update the Excel file
        df = load_data()
        df.loc[int(self.tree.index(selected_item))] = [first_name, last_name, phone_number, email, item1, item2, total_price]
        save_data(df)

        # Close the edit row window
        edit_window.destroy()

    # Create a new window for editing a row
    edit_window = tk.Toplevel(self.root)
    edit_window.title("Edit Row")

    # Create entry fields for each column pre-filled with the selected row's data
    tk.Label(edit_window, text="First Name").grid(row=0, column=0)
    entry_first_name = tk.Entry(edit_window)
    entry_first_name.grid(row=0, column=1)
    entry_first_name.insert(0, values[0])

    tk.Label(edit_window, text="Last Name").grid(row=1, column=0)
    entry_last_name = tk.Entry(edit_window)
    entry_last_name.grid(row=1, column=1)
    entry_last_name.insert(0, values[1])

    tk.Label(edit_window, text="Phone Number").grid(row=2, column=0)
    entry_phone_number = tk.Entry(edit_window)
    entry_phone_number.grid(row=2, column=1)
    entry_phone_number.insert(0, values[2])

    tk.Label(edit_window, text="Email").grid(row=3, column=0)
    entry_email = tk.Entry(edit_window)
    entry_email.grid(row=3, column=1)
    entry_email.insert(0, values[3])

    tk.Label(edit_window, text="Item 1").grid(row=4, column=0)
    entry_item1 = tk.Entry(edit_window)
    entry_item1.grid(row=4, column=1)
    entry_item1.insert(0, values[4])

    tk.Label(edit_window, text="Item 2").grid(row=5, column=0)
    entry_item2 = tk.Entry(edit_window)
    entry_item2.grid(row=5, column=1)
    entry_item2.insert(0, values[5])

    tk.Label(edit_window, text="Total Price").grid(row=6, column=0)
    entry_total_price = tk.Entry(edit_window)
    entry_total_price.grid(row=6, column=1)
    entry_total_price.insert(0, values[6])

    # Create a submit button
    submit_button = tk.Button(edit_window, text="Submit", command=submit)
    submit_button.grid(row=7, columnspan=2)

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

        add_btn = tk.Button(toolbar, text="Add Row", command=lambda: add_row(self))
        add_btn.pack(side=tk.LEFT, padx=2, pady=2)

        delete_btn = tk.Button(toolbar, text="Delete Row", command=lambda: delete_row())
        delete_btn.pack(side=tk.LEFT, padx=2, pady=2)

        edit_btn = tk.Button(toolbar, text="Edit Row", command=lambda: edit_row(self))
        edit_btn.pack(side=tk.LEFT, padx=2, pady=2)

        generate_btn = tk.Button(toolbar, text="Generate Receipts")
        generate_btn.pack(side=tk.LEFT, padx=2, pady=2)

        send_phone_btn = tk.Button(toolbar, text="Send to Phone")
        send_phone_btn.pack(side=tk.LEFT, padx=2, pady=2)

        send_email_btn = tk.Button(toolbar, text="Send to Email")
        send_email_btn.pack(side=tk.LEFT, padx=2, pady=2)

        preview_btn = tk.Button(toolbar, text="Preview Receipt")
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
