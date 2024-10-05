import customtkinter as ctk
from tkinter import ttk  # Import ttk for Treeview
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox


# Clear the current item inputs
def clear_item():
    quantity_spinbox.delete(0, ctk.END)
    quantity_spinbox.insert(0, "1")
    desc_entry.delete(0, ctk.END)
    price_spinbox.delete(0, ctk.END)
    price_spinbox.insert(0, "0.0")


# List to store invoice items
invoice_list = []


# Add item to the invoice
def add_item():
    try:
        qty = int(quantity_spinbox.get())
        desc = desc_entry.get().capitalize()  # Capitalize the first letter of the description
        price = float(price_spinbox.get())
        line_total = round(qty * price, 2)  # Round the line total to 2 decimal places
        invoice_item = [qty, desc, price, line_total]
        tree.insert('', 0, values=invoice_item)
        clear_item()
        invoice_list.append(invoice_item)
    except ValueError:
        messagebox.showerror("Input Error", "Please enter valid quantity and price values.")


# Clear the form for a new invoice
def new_invoice():
    first_name_entry.delete(0, ctk.END)
    last_name_entry.delete(0, ctk.END)
    phone_entry.delete(0, ctk.END)
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
        subtotal = round(sum(item[3] for item in invoice_list), 2)  # Round subtotal to 2 decimal places
        salestax = 0.1
        total = round(subtotal * (1 + salestax), 2)  # Round the total to 2 decimal places
        doc_type_selection = "Invoice" if doc_type_option.get() == 0 else "Quote"

        doc.render({
            "name": name,
            "phone": phone,
            "invoice_list": invoice_list,
            "subtotal": f"{subtotal:.2f}",  # Format subtotal to 2 decimal places
            "salestax": str(salestax * 100) + "%",
            "total": f"{total:.2f}",  # Format total to 2 decimal places
            "doc_type": doc_type_selection
        })

        doc_name = f"{doc_type_selection}_{name}_{datetime.datetime.now().strftime('%Y-%m-%d-%H%M%S')}.docx"
        doc.save(doc_name)

        messagebox.showinfo(doc_type_selection + " Complete", doc_type_selection + " document generated successfully.")
        new_invoice()

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while generating the document: {e}")


# Main window
ctk.set_appearance_mode("dark")  # Options: "System" (default), "Light", "Dark"
ctk.set_default_color_theme("blue")  # Options: "blue" (default), "green", "dark-blue"

window = ctk.CTk()
window.title("Invoice Generator Form")
window.geometry("600x600")

frame = ctk.CTkFrame(window)
frame.pack(padx=20, pady=20)

doc_type_option = ctk.IntVar()

# Customer Information Frame
customer_info_frame = ctk.CTkFrame(frame, corner_radius=10)
customer_info_frame.grid(row=0, column=0, padx=10, pady=10)

ctk.CTkLabel(customer_info_frame, text="Customer Information").grid(row=0, column=0, columnspan=2, pady=5)
ctk.CTkLabel(customer_info_frame, text="First Name:").grid(row=1, column=0)
ctk.CTkLabel(customer_info_frame, text="Last Name:").grid(row=2, column=0)
ctk.CTkLabel(customer_info_frame, text="Phone:").grid(row=3, column=0)

first_name_entry = ctk.CTkEntry(customer_info_frame)
first_name_entry.grid(row=1, column=1, pady=5)
last_name_entry = ctk.CTkEntry(customer_info_frame)
last_name_entry.grid(row=2, column=1, pady=5)
phone_entry = ctk.CTkEntry(customer_info_frame)
phone_entry.grid(row=3, column=1, pady=5)

# Invoice/Quote selection
invoiceOrQuote_frame = ctk.CTkFrame(frame, corner_radius=10)
invoiceOrQuote_frame.grid(row=0, column=1, padx=10, pady=10)

ctk.CTkLabel(invoiceOrQuote_frame, text="Invoice / Quote").grid(row=0, column=0, columnspan=2, pady=5)
invoice_radiobutton = ctk.CTkRadioButton(invoiceOrQuote_frame, text='Invoice', variable=doc_type_option, value=0)
invoice_radiobutton.grid(row=1, column=0)
quote_radiobutton = ctk.CTkRadioButton(invoiceOrQuote_frame, text='Quote', variable=doc_type_option, value=1)
quote_radiobutton.grid(row=2, column=0)

# Item Details
ctk.CTkLabel(frame, text="Qty").grid(row=1, column=0)
quantity_spinbox = ctk.CTkEntry(frame)
quantity_spinbox.insert(0, "1")
quantity_spinbox.grid(row=2, column=0, pady=5)

ctk.CTkLabel(frame, text="Description").grid(row=1, column=1)
desc_entry = ctk.CTkEntry(frame)
desc_entry.grid(row=2, column=1, pady=5)

ctk.CTkLabel(frame, text="Unit Price").grid(row=1, column=2)
price_spinbox = ctk.CTkEntry(frame)
price_spinbox.insert(0, "0.0")
price_spinbox.grid(row=2, column=2, pady=5)

# Add item button
add_item_button = ctk.CTkButton(frame, text="Add Item", command=add_item)
add_item_button.grid(row=3, column=2, pady=5)

# Treeview for invoice items using ttk
columns = ('qty', 'desc', 'price', 'total')
tree = ttk.Treeview(frame, columns=columns, show="headings")
tree.heading('qty', text='Qty')
tree.heading('desc', text='Description')
tree.heading('price', text='Unit Price')
tree.heading('total', text='Total')
tree.grid(row=4, column=0, columnspan=3, padx=20, pady=10)

# Generate document button
save_invoice_button = ctk.CTkButton(frame, text="Generate Document", command=generate_invoice)
save_invoice_button.grid(row=5, column=0, columnspan=3, sticky="news", padx=20, pady=5)

# New Invoice button
new_invoice_button = ctk.CTkButton(frame, text="New Invoice", command=new_invoice)
new_invoice_button.grid(row=6, column=0, columnspan=3, sticky="news", padx=20, pady=5)

window.mainloop()
