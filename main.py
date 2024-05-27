import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox


# List that holds information of the invoice being filled out.
invoiceList = []


# Functions
def clearItem():
    qty_spinbox.delete(0, tkinter.END)
    qty_spinbox.insert(0, "1")
    desc_entry.delete(0, tkinter.END)
    unitPrice_entry.delete(0, tkinter.END)


def addItem():
    qty = int(qty_spinbox.get())
    desc = desc_entry.get()
    price = float(unitPrice_entry.get())
    lineTotal = qty*price
    invoiceItem = [qty, desc, price, lineTotal]
    # Inserting the values from input onto a treeview in the GUI
    tree.insert('', 0, values= invoiceItem)
    clearItem()
    # Appending values from input onto the list of items.
    invoiceList.append(invoiceItem)


def newInvoice():
    firstName_entry.delete(0, tkinter.END)
    lastName_entry.delete(0, tkinter.END)
    phoneNum_entry.delete(0, tkinter.END)
    tree.delete(*tree.get_children())
    # When Creating a new invoice, delete everything inside the list + all items showcased on GUI
    clearItem()
    invoiceList.clear()


def generateInvoice():
    doc = DocxTemplate("invoice_template.docx")
    name = firstName_entry.get() + " " + lastName_entry.get()
    phone = phoneNum_entry.get()
    
    
    subtotal = sum(item[3] for item in invoiceList)
    salestax = 0.1          # Change this to the percetange of sales tax. Default = 0.1 for 10%
    total = subtotal*(1-salestax)

    doc.render({"name":name,
                "phone":phone,
                "invoice_list": invoiceList,
                "subtotal":subtotal,
                "salestax":str(salestax*100)+"%",
                "total":total})
    
    doc_name = "newInvoice" + name + datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S") + ".doxc"
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

qty_label = tkinter.Label(frame, text = "Qty")
qty_label.grid(row = 2, column = 0)
qty_spinbox = tkinter.Spinbox(frame, from_= 1, to = 100)
qty_spinbox.grid(row = 3, column = 0)

desc_label = tkinter.Label(frame, text = "Description")
desc_label.grid(row = 2, column = 1)
desc_entry = tkinter.Entry(frame)
desc_entry.grid(row = 3, column = 1)

unitPrice_label = tkinter.Label(frame, text = "Unit Price")
unitPrice_label.grid(row = 2, column = 2)
unitPrice_entry = tkinter.Entry(frame)
unitPrice_entry.grid(row = 3, column = 2)

addItem_button = tkinter.Button(frame, text = "Add Item", command = addItem)
addItem_button.grid(row = 4, column = 2, pady = 10)

# Using TTK to Create a Tree View of the Items Added onto the List
columns = ('qty', 'desc', 'price', 'total')
tree = ttk.Treeview(frame, columns = columns, show = "headings")

tree.heading('qty', text = 'Qty')
tree.heading('desc', text = 'Description')
tree.heading('price', text = 'Unit Price')
tree.heading('total', text = 'Total')

# Columnspan = 3 says to take up the space that 3 columns do, To keep presentability 
tree.grid(row = 5, column = 0, columnspan = 3, padx = 20, pady = 10)

# Generate New Invoice  / Save Invoice
# sticky = 'news' = North, East, West, South. I'm telling the button to stay expanded in these directions

saveInvoiceButton = tkinter.Button(frame , text = "Generate Invoice", command = generateInvoice)
saveInvoiceButton.grid(row = 6, column = 0, columnspan = 3, sticky = "news", padx= 20, pady = 5)
newInvoiceButton = tkinter.Button(frame , text = "New Invoice", command = newInvoice)
newInvoiceButton.grid(row = 7, column = 0, columnspan = 3, sticky = "news", padx = 20, pady = 5)






window.mainloop()