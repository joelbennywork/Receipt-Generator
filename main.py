import tkinter as tk
from tkinter import ttk
from openpyxl import load_workbook

# Load the Excel file
file_path = 'invoice.xlsx'
wb = load_workbook(file_path)
ws = wb.active

# Create the main window
root = tk.Tk()
root.title("Excel Editor")
root.geometry("1000x600")

# Create a toolbar
toolbar = tk.Frame(root)
toolbar.pack(side=tk.TOP, fill=tk.X)

# Create a frame for entry widgets
entry_frame = tk.Frame(root)
entry_frame.pack(side=tk.TOP, fill=tk.X)

# Undo and redo stacks
undo_stack = []
redo_stack = []

def add_row():
    new_row = [len(tree.get_children()) + 1] + ['']*7
    tree.insert('', 'end', values=new_row)
    undo_stack.append(('add', new_row))

def delete_row():
    selected_item = tree.selection()[0]
    values = tree.item(selected_item)['values']
    tree.delete(selected_item)
    undo_stack.append(('delete', values))

def edit_row():
    selected_item = tree.selection()[0]
    old_values = tree.item(selected_item)['values']
    new_values = [entry.get() if entry.get() else old_values[i] for i, entry in enumerate(entries)]
    tree.item(selected_item, values=new_values)
    undo_stack.append(('edit', old_values, new_values))

def undo():
    if undo_stack:
        action = undo_stack.pop()
        if action[0] == 'add':
            for item in tree.get_children():
                if tree.item(item)['values'] == action[1]:
                    tree.delete(item)
                    break
        elif action[0] == 'delete':
            tree.insert('', 'end', values=action[1])
        elif action[0] == 'edit':
            for item in tree.get_children():
                if tree.item(item)['values'] == action[2]:
                    tree.item(item, values=action[1])
                    break
        redo_stack.append(action)

def redo():
    if redo_stack:
        action = redo_stack.pop()
        if action[0] == 'add':
            tree.insert('', 'end', values=action[1])
        elif action[0] == 'delete':
            for item in tree.get_children():
                if tree.item(item)['values'] == action[1]:
                    tree.delete(item)
                    break
        elif action[0] == 'edit':
            for item in tree.get_children():
                if tree.item(item)['values'] == action[1]:
                    tree.item(item, values=action[2])
                    break
        undo_stack.append(action)

add_button = tk.Button(toolbar, text="Add Row", command=add_row)
add_button.pack(side=tk.LEFT, padx=2, pady=2)

delete_button = tk.Button(toolbar, text="Delete Row", command=delete_row)
delete_button.pack(side=tk.LEFT, padx=2, pady=2)

edit_button = tk.Button(toolbar, text="Edit Row", command=edit_row)
edit_button.pack(side=tk.LEFT, padx=2, pady=2)

undo_button = tk.Button(toolbar, text="Undo", command=undo)
undo_button.pack(side=tk.LEFT, padx=2, pady=2)

redo_button = tk.Button(toolbar, text="Redo", command=redo)
redo_button.pack(side=tk.LEFT, padx=2, pady=2)

# Create the treeview
columns = ['Row Number'] + [ws.cell(row=1, column=i).value for i in range(1, 9)]
tree = ttk.Treeview(root, columns=columns, show='headings')
tree.pack(fill=tk.BOTH, expand=True)

for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=100)

# Load data from Excel into the treeview
for i, row in enumerate(ws.iter_rows(min_row=2, max_col=8, values_only=True), start=1):
    tree.insert('', 'end', values=[i] + list(row))

# Function to update column H with the sum of columns E to G
def update_sums():
    for row in tree.get_children():
        values = tree.item(row)['values']
        if all(isinstance(values[i], (int, float)) for i in range(5, 8)):
            values[8] = sum(values[5:8])
            tree.item(row, values=values)

tree.bind('<FocusOut>', lambda e: update_sums())

# Create entry widgets for editing
entries = []
labels = ['Row Number'] + [ws.cell(row=1, column=i).value for i in range(1, 9)]
for label in labels:
    lbl = tk.Label(entry_frame, text=label)
    lbl.pack(side=tk.LEFT, padx=2, pady=2)
    entry = tk.Entry(entry_frame)
    entry.pack(side=tk.LEFT, padx=2, pady=2)
    entries.append(entry)

def on_row_select(event):
    selected_item = tree.selection()[0]
    values = tree.item(selected_item)['values']
    for i, entry in enumerate(entries):
        entry.delete(0, tk.END)
        entry.insert(0, values[i])

tree.bind('<<TreeviewSelect>>', on_row_select)

root.mainloop()
