import tkinter as tk
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
import os
from tkinter import messagebox
import sqlite3

from numpy import loadtxt
lines = loadtxt("item_stock.txt", dtype=str, usecols=0, unpack=True)

def clear_item():
    qty_spinbox.delete(0, tk.END)
    qty_spinbox.insert(0, "1")
    desc_entry.delete(0, tk.END)
    price_spinbox.delete(0, tk.END)
    price_spinbox.insert(0, "0.0")

invoice_list = []
def add_item():
    qty = int(qty_spinbox.get())
    desc = desc_entry.get()
    price = float(price_spinbox.get())
    line_total = qty * price
    invoice_item = [qty, desc, price, line_total]

    tree.insert('', 0, values=invoice_item)
    clear_item()
    invoice_list.append(invoice_item)

def new_invoice():
    first_name_entry.delete(0, tk.END)
    last_name_entry.delete(0, tk.END)
    phone_entry.delete(0, tk.END)
    clear_item()
    tree.delete(*tree.get_children())
    invoice_list.clear()
    refresh_invoice_listbox()  # Refresh the Listbox when a new invoice is created

def generate_invoice():
    doc = DocxTemplate("invoice_template.docx")
    name = first_name_entry.get() + " " + last_name_entry.get()  # Add space between names
    phone = phone_entry.get()
    subtotal = sum(item[3] for item in invoice_list)
    salestax = 0.1
    total = subtotal * (1 + salestax)  # Add sales tax to the subtotal

    doc.render({"name": name,
                "phone": phone,
                "invoice_list": invoice_list,
                "subtotal": subtotal,
                "salestax": str(salestax * 100) + "%",
                "total": total})
    
    doc_name = os.path.join("./invoices/", f"new_invoice-{name}-{datetime.datetime.now().strftime('%Y-%m-%d-%H%M%S')}.docx")
    doc.save(doc_name)

    messagebox.showinfo("Invoice Generation", "Invoice Complete")
    new_invoice()

    # Create Table in Database
    conn = sqlite3.connect('./invDatabase/data.db')
    table_create_query = '''
        CREATE TABLE IF NOT EXISTS Invoice_Data(
            name TEXT,
            phone TEXT,
            subtotal REAL,
            salestax REAL,
            total REAL
        )
    '''
    conn.execute(table_create_query)

    # Insert Data into Database
    data_insert_query = '''
        INSERT INTO Invoice_Data (name, phone, subtotal, salestax, total) 
        VALUES (?, ?, ?, ?, ?)
    '''
    data_insert_tuple = (name, phone, subtotal, salestax, total)
    cursor = conn.cursor()
    cursor.execute(data_insert_query, data_insert_tuple)
    conn.commit()
    conn.close()

    # Refresh the Listbox with data from the database
    refresh_invoice_listbox()

def refresh_invoice_listbox():
    # Clear the existing content in the Listbox
    invoice_listbox.delete(0, tk.END)
    
    # Fetch data from the database
    conn = sqlite3.connect('./invDatabase/data.db')
    cursor = conn.cursor()
    cursor.execute("SELECT name, phone, subtotal, salestax, total FROM Invoice_Data")
    rows = cursor.fetchall()
    conn.close()

    # Populate the Listbox with the fetched data
    for row in rows:
        invoice_listbox.insert(tk.END, f"Name: {row[0]}, Phone: {row[1]}, Subtotal: {row[2]}, Tax: {row[3]}, Total: {row[4]}")

# Initialize main window
window = tk.Tk()
window.title("Invoice Generator Form")


frame = tk.Frame(window)
frame.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)

# Invoice Data Entry
first_name_label = tk.Label(frame, text="First Name")
last_name_label = tk.Label(frame, text="Last Name")
first_name_label.grid(row=0, column=0)
last_name_label.grid(row=0, column=1)

first_name_entry = tk.Entry(frame)
last_name_entry = tk.Entry(frame)
first_name_entry.grid(row=1, column=0)
last_name_entry.grid(row=1, column=1)

phone_label = tk.Label(frame, text="Phone")
phone_label.grid(row=0, column=2)
phone_entry = tk.Entry(frame)
phone_entry.grid(row=1, column=2)

qty_label = tk.Label(frame, text="Qty")
qty_label.grid(row=2, column=0)
qty_spinbox = tk.Spinbox(frame, from_=1, to=100)
qty_spinbox.grid(row=3, column=0)

desc_label = tk.Label(frame, text="Description")
desc_label.grid(row=2, column=1)
desc_entry = ttk.Combobox(frame, values=list(map(str, lines)))
desc_entry.grid(row=3, column=1)

price_label = tk.Label(frame, text="Unit Price")
price_label.grid(row=2, column=2)
price_spinbox = tk.Spinbox(frame, from_=0.0, to=500, increment=0.5)
price_spinbox.grid(row=3, column=2)

add_item_button = tk.Button(frame, text="Add Item", command=add_item, cursor="hand2")
add_item_button.grid(row=4, column=2, pady=5)

# Treeview for items
columns = ('qty', 'desc', 'price', 'total')
tree = ttk.Treeview(frame, columns=columns, show="headings")
tree.heading('qty', text="Qty")
tree.heading('desc', text="Description")
tree.heading('price', text="Unit Price")
tree.heading('total', text="Total")
tree.grid(row=5, column=0, columnspan=3, padx=20, pady=10)

save_invoice_button = tk.Button(frame, text="Generate Invoice", command=generate_invoice, cursor="hand2")
save_invoice_button.grid(row=6, column=0, columnspan=3, sticky="news", padx=20, pady=5)

new_invoice_button = tk.Button(frame, text="New Invoice", command=new_invoice, cursor="hand2")
new_invoice_button.grid(row=7, column=0, columnspan=3, sticky="news", padx=20, pady=5)

# Right side - LabelFrame for displaying previous invoices
invoice_labelframe = tk.LabelFrame(window, text="Previous Invoices")
invoice_labelframe.pack(side=tk.RIGHT, padx=20, pady=10, fill=tk.BOTH, expand=True)

invoice_listbox = tk.Listbox(invoice_labelframe, height=12)
invoice_listbox.pack(fill=tk.BOTH, expand=True)

# Load existing invoices into Listbox when application starts
refresh_invoice_listbox()

# Start the application
window.mainloop()
