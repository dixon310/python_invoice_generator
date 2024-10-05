import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox


def clear_item():
    qty_spinbox.delete(0, tkinter.END)
    qty_spinbox.insert(0, "1")
    desc_entry.delete(0, tkinter.END)
    price_spinbox.delete(0, tkinter.END)
    price_spinbox.insert(0, "0.0")


invoice_list = []
def add_item():
    qty = int(qty_spinbox.get())
    desc = desc_entry.get()
    price = float(price_spinbox.get())
    line_total = qty*price
    invoice_item = [qty, desc, price, line_total]
    tree.insert('',0, values=invoice_item)
    clear_item()

    invoice_list.append(invoice_item)


def invoiceOrQuote():
    doc_type_selection = str(doc_type_option.get())

def new_invoice():
    first_name_entry.delete(0, tkinter.END)
    last_name_entry.delete(0, tkinter.END)
    phone_entry.delete(0, tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())
    
    invoice_list.clear()

def generate_invoice():
    doc = DocxTemplate("invoice_template.docx")
    name = first_name_entry.get()+ " " + last_name_entry.get()
    phone = phone_entry.get()
    subtotal = sum(item[3] for item in invoice_list) 
    salestax = 0.1
    total = subtotal*(1-salestax)
    doc_type_selection = "Invoice" if doc_type_option.get() == 0 else "Quote"
    
    doc.render({
                "name":name, 
                "phone":phone,
                "invoice_list": invoice_list,
                "subtotal":subtotal,
                "salestax":str(salestax*100)+"%",
                "total":total,
                "doc_type":doc_type_selection
                })
    
    doc_name = doc_type_selection + name + "_" + datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S") + ".docx"
    doc.save(doc_name)
    
    messagebox.showinfo(doc_type_selection+" Complete", doc_type_selection+" Complete")
    
    new_invoice()

window = tkinter.Tk()
window.title("Invoice Generator Form")

frame = tkinter.Frame(window)
frame.pack()

doc_type_option = tkinter.IntVar()

customer_info_frame = tkinter.LabelFrame(frame, text="Customer Information")
customer_info_frame.grid(row=0, column=0)

invoiceOrQuote_frame = tkinter.LabelFrame(frame, text="Invoice / Quote")
invoiceOrQuote_frame.grid(row=0, column=1)

invoice_radiobutton = tkinter.Radiobutton(invoiceOrQuote_frame, text='INVOICE',variable=doc_type_option ,value=0, command=invoiceOrQuote)
invoice_radiobutton.grid(row=0,column=0)
quote_radiobutton = tkinter.Radiobutton(invoiceOrQuote_frame, text='QUOTE',variable=doc_type_option , value=1, command=invoiceOrQuote)
quote_radiobutton.grid(row=1,column=0)

company_name_label = tkinter.Label(customer_info_frame, text="Company Name: ").grid(row=0, column=0)
first_name_label = tkinter.Label(customer_info_frame, text="First Name: ").grid(row=1, column=0)
last_name_label = tkinter.Label(customer_info_frame, text="Last Name: ").grid(row=2,column=0)
company_name_entry = tkinter.Entry(customer_info_frame).grid(row=0, column=1)

first_name_entry = tkinter.Entry(customer_info_frame)
first_name_entry.grid(row=1, column=1)
last_name_entry = tkinter.Entry(customer_info_frame)
last_name_entry.grid(row=2, column=1)

phone_label = tkinter.Label(customer_info_frame, text="Phone")
phone_entry = tkinter.Entry(customer_info_frame)
phone_label.grid(row=3, column=0)
phone_entry.grid(row=3, column=1)

qty_label = tkinter.Label(frame, text="Qty")
qty_label.grid(row=2, column=0)
qty_spinbox = tkinter.Spinbox(frame, from_=1, to=100)
qty_spinbox.grid(row=3, column=0)

desc_label = tkinter.Label(frame, text="Description")
desc_label.grid(row=2, column=1)
desc_entry = tkinter.Entry(frame)
desc_entry.grid(row=3, column=1)

price_label = tkinter.Label(frame, text="Unit Price")
price_label.grid(row=2, column=2)
price_spinbox = tkinter.Spinbox(frame, from_=0.0, to=999999, increment=0.5)
price_spinbox.grid(row=3, column=2)

add_item_button = tkinter.Button(frame, text="Add item", command=add_item)
add_item_button.grid(row=4, column=2, pady=5)


columns = ('qty', 'desc', 'price', 'total')
tree = ttk.Treeview(frame, columns=columns, show="headings")
tree.heading('qty', text='Qty')
tree.heading('desc', text='Description')
tree.heading('price', text='Unit Price')
tree.heading('total', text="Total")
tree.grid(row=5, column=0, columnspan=3, padx=20, pady=10)

save_invoice_button = tkinter.Button(frame, text="Generate document",command=generate_invoice)
save_invoice_button.grid(row=6, column=0, columnspan=3, sticky="news", padx=20, pady=5)
new_invoice_button = tkinter.Button(frame, text="New Invoice", command=new_invoice)
new_invoice_button.grid(row=7, column=0, columnspan=3, sticky="news", padx=20, pady=5)


window.mainloop()