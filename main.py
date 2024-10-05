import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox


# Clear the current item inputs
def clear_item():
    quantity_spinbox.delete(0, tkinter.END)
    quantity_spinbox.insert(0, "1")
    desc_entry.delete(0, tkinter.END)
    price_spinbox.delete(0, tkinter.END)
    price_spinbox.insert(0, "0.0")


# List to store invoice items
invoice_list = []

# Add item to the invoice
def add_item():
    try:
        qty = int(quantity_spinbox.get())
        desc = desc_entry.get().capitalize()
        price = float(price_spinbox.get())
        line_total = round(qty * price, 2)
        invoice_item = [qty, desc, price, line_total]
        tree.insert('', 0, values=invoice_item)
        clear_item()
        invoice_list.append(invoice_item)
    except ValueError:
        messagebox.showerror("Input Error", "Please enter valid quantity and price values.")


# Clear the form for a new invoice
def new_invoice():
    first_name_entry.delete(0, tkinter.END)
    last_name_entry.delete(0, tkinter.END)
    phone_entry.delete(0, tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())
    invoice_list.clear()


# Generate and save the invoice/quote document
def generate_invoice():
    if not invoice_list:
        messagebox.showerror("No Items", "You need to add at least one item to the invoice.")
        return

    try:
        doc = DocxTemplate("invoice_template.docx")
        name = first_name_entry.get() + " " + last_name_entry.get()
        phone = phone_entry.get()
        subtotal = sum(item[3] for item in invoice_list)
        salestax = 0.1
        total = subtotal * (1 + salestax)
        doc_type_selection = "Invoice" if doc_type_option.get() == 0 else "Quote"

        doc.render({
            "name": name,
            "phone": phone,
            "invoice_list": invoice_list,
            "subtotal": subtotal,
            "salestax": str(salestax * 100) + "%",
            "total": total,
            "doc_type": doc_type_selection
        })

        doc_name = f"{doc_type_selection}_{name}_{datetime.datetime.now().strftime('%Y-%m-%d-%H%M%S')}.docx"
        doc.save(doc_name)

        messagebox.showinfo(doc_type_selection + " Complete", doc_type_selection + " document generated successfully.")
        new_invoice()

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while generating the document: {e}")


# Main window
window = tkinter.Tk()
window.title("Invoice Generator Form")

frame = tkinter.Frame(window)
frame.pack()

doc_type_option = tkinter.IntVar()

# Customer Information Frame
customer_info_frame = tkinter.LabelFrame(frame, text="Customer Information")
customer_info_frame.grid(row=0, column=0)

# Invoice/Quote selection frame
invoiceOrQuote_frame = tkinter.LabelFrame(frame, text="Invoice / Quote")
invoiceOrQuote_frame.grid(row=0, column=1)

# Invoice/Quote radio buttons
invoice_radiobutton = tkinter.Radiobutton(invoiceOrQuote_frame, text='INVOICE', variable=doc_type_option, value=0)
invoice_radiobutton.grid(row=0, column=0)
quote_radiobutton = tkinter.Radiobutton(invoiceOrQuote_frame, text='QUOTE', variable=doc_type_option, value=1)
quote_radiobutton.grid(row=1, column=0)

# Customer Information fields
tkinter.Label(customer_info_frame, text="Company Name:").grid(row=0, column=0)
tkinter.Label(customer_info_frame, text="First Name:").grid(row=1, column=0)
tkinter.Label(customer_info_frame, text="Last Name:").grid(row=2, column=0)
tkinter.Label(customer_info_frame, text="Phone:").grid(row=3, column=0)

company_name_entry = tkinter.Entry(customer_info_frame)
company_name_entry.grid(row=0, column=1)
first_name_entry = tkinter.Entry(customer_info_frame)
first_name_entry.grid(row=1, column=1)
last_name_entry = tkinter.Entry(customer_info_frame)
last_name_entry.grid(row=2, column=1)
phone_entry = tkinter.Entry(customer_info_frame)
phone_entry.grid(row=3, column=1)

# Item Details (Qty, Description, Price)
tkinter.Label(frame, text="Qty").grid(row=2, column=0)
quantity_spinbox = tkinter.Spinbox(frame, from_=1, to=100)
quantity_spinbox.grid(row=3, column=0)

tkinter.Label(frame, text="Description").grid(row=2, column=1)
desc_entry = tkinter.Entry(frame)
desc_entry.grid(row=3, column=1)

tkinter.Label(frame, text="Unit Price").grid(row=2, column=2)
price_spinbox = tkinter.Spinbox(frame, from_=0.0, to=999999, increment=0.5)
price_spinbox.grid(row=3, column=2)

# Add item button
add_item_button = tkinter.Button(frame, text="Add item", command=add_item)
add_item_button.grid(row=4, column=2, pady=5)

# Treeview for invoice items
columns = ('qty', 'desc', 'price', 'total')
tree = ttk.Treeview(frame, columns=columns, show="headings")
tree.heading('qty', text='Qty')
tree.heading('desc', text='Description')
tree.heading('price', text='Unit Price')
tree.heading('total', text="Total")
tree.grid(row=5, column=0, columnspan=3, padx=20, pady=10)

# Generate document button
save_invoice_button = tkinter.Button(frame, text="Generate document", command=generate_invoice)
save_invoice_button.grid(row=6, column=0, columnspan=3, sticky="news", padx=20, pady=5)

# New Invoice button
new_invoice_button = tkinter.Button(frame, text="New Invoice", command=new_invoice)
new_invoice_button.grid(row=7, column=0, columnspan=3, sticky="news", padx=20, pady=5)

window.mainloop()