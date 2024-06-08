import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox
import uuid
import locale
import os
import sys

# Set locale for currency formatting
locale.setlocale(locale.LC_ALL, 'en_US.UTF-8')

# List that holds information of the invoice being filled out.
invoiceList = []

# Function to get the absolute path to resource
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Functions
def clearItem():
    job_entry.delete(0, tkinter.END)
    desc_entry.delete(0, tkinter.END)
    unitPrice_entry.delete(0, tkinter.END)


def addItem():
    job = job_entry.get()
    desc = desc_entry.get()
    price = float(unitPrice_entry.get())
    invoiceItem = [job, desc, price]
    # Inserting the values from input onto a treeview in the GUI
    tree.insert('', 0, values= invoiceItem)
    clearItem()
    # Appending values from input onto the list of items.
    invoiceList.append(invoiceItem)


def newInvoice():
    firstName_entry.delete(0, tkinter.END)
    lastName_entry.delete(0, tkinter.END)
    phoneNum_entry.delete(0, tkinter.END)
    address_entry.delete(0, tkinter.END)
    address2_entry.delete(0, tkinter.END)
    tree.delete(*tree.get_children())
    # When Creating a new invoice, delete everything inside the list + all items showcased on GUI
    clearItem()
    invoiceList.clear()

def format_currency(value):
    return locale.currency(value, grouping=True)


def generateInvoice():
    doc_path = resource_path("invoice_template.docx")
    doc = DocxTemplate(doc_path)
    name = firstName_entry.get() + " " + lastName_entry.get()
    phone = phoneNum_entry.get()
    date = datetime.datetime.now().strftime("%m/%d/%Y")
    address = address_entry.get()
    address2 = address2_entry.get()
    estimate_id = str(uuid.uuid4())[:8]

    formatted_invoice_list = [{"job": item[0], "desc": item[1], "price": format_currency(item[2])} for item in invoiceList]

    subtotal = sum(item[2] for item in invoiceList)
    formatted_subtotal = format_currency(subtotal)
    salestax_rate = 0.07                                 # Change this to the percentage of sales tax. Default = 0.1 for 10%
    salestax = subtotal * salestax_rate
    formatted_salestax = format_currency(salestax)
    total = subtotal * (1 + salestax_rate)
    formatted_total = format_currency(total)

    doc.render({"name":name,
                "address":address,
                "date":date,
                "phone":phone,
                "address2":address2,
                "estimate":estimate_id,
                "invoice_list": formatted_invoice_list,
                "subtotal":formatted_subtotal,
                "salestax": formatted_salestax,
                "total":formatted_total})
    
    doc_name = "newInvoice_" + name + "_" + datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S") + ".docx"
    doc.save(doc_name)

    messagebox.showinfo("Invoice Complete", "Invoice Complete")

    newInvoice()


#Creating GUI Window
window = tkinter.Tk()
window.title("Invoice Generator Form")

#Creating a Frame inside Parent Window to store widgets.
frame = tkinter.Frame(window)
frame.pack(padx = 20, pady = 10)

#Creating Grid of Inputs with Labels
firstName_label = tkinter.Label(frame, text = "First Name")
firstName_label.grid(row = 0, column = 0)
firstName_entry = tkinter.Entry(frame)
firstName_entry.grid(row = 1, column = 0)

lastName_label = tkinter.Label(frame, text = "Last Name")
lastName_label.grid(row = 0, column = 1)
lastName_entry = tkinter.Entry(frame)
lastName_entry.grid(row = 1, column = 1)

phoneNum_label = tkinter.Label(frame, text = "Phone Number")
phoneNum_label.grid(row = 0, column = 2)
phoneNum_entry = tkinter.Entry(frame)
phoneNum_entry.grid(row = 1, column = 2)

address_label = tkinter.Label(frame, text = "Address")
address_label.grid(row = 0, column = 3)
address_entry = tkinter.Entry(frame)
address_entry.grid(row = 1, column = 3)

job_label = tkinter.Label(frame, text = "Job")
job_label.grid(row = 2, column = 0)
job_entry = tkinter.Entry(frame)
job_entry.grid(row = 3, column = 0)

desc_label = tkinter.Label(frame, text = "Description")
desc_label.grid(row = 2, column = 1)
desc_entry = tkinter.Entry(frame)
desc_entry.grid(row = 3, column = 1)

unitPrice_label = tkinter.Label(frame, text = "Cost of Job")
unitPrice_label.grid(row = 2, column = 2)
unitPrice_entry = tkinter.Entry(frame)
unitPrice_entry.grid(row = 3, column = 2)

address2_label = tkinter.Label(frame, text = "City/State/Zip")
address2_label.grid(row = 2, column = 3)
address2_entry = tkinter.Entry(frame)
address2_entry.grid(row = 3, column = 3)

addItem_button = tkinter.Button(frame, text = "Add Item", command = addItem)
addItem_button.grid(row = 4, column = 3, pady = 10)

# Using TTK to Create a Tree View of the Items Added onto the List
columns = ('job', 'desc', 'price')
tree = ttk.Treeview(frame, columns = columns, show = "headings")

tree.heading('job', text = 'Job')
tree.heading('desc', text = 'Description')
tree.heading('price', text = 'Job Price')

# Columnspan = 3 says to take up the space that 3 columns do, To keep presentability 
tree.grid(row = 5, column = 0, columnspan = 4, padx = 20, pady = 10)

# Generate New Invoice  / Save Invoice
# sticky = 'news' = North, East, West, South. I'm telling the button to stay expanded in these directions
saveInvoiceButton = tkinter.Button(frame , text = "Generate Invoice", command = generateInvoice)
saveInvoiceButton.grid(row = 6, column = 0, columnspan = 4, sticky = "news", padx= 20, pady = 5)
newInvoiceButton = tkinter.Button(frame , text = "New Invoice", command = newInvoice)
newInvoiceButton.grid(row = 7, column = 0, columnspan = 4, sticky = "news", padx = 20, pady = 5)

window.mainloop()