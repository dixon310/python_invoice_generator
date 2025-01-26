import customtkinter as ctk
from tkinter import ttk  # Import ttk for Treeview
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox
import json
import os


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
        invoice_folder = "invoices"
        if not os.path.exists(invoice_folder):
            os.makedirs(invoice_folder)

        doc_path = os.path.join(invoice_folder, doc_name)
        doc.save(doc_path)

        messagebox.showinfo(doc_type_selection + " Complete", doc_type_selection + " document generated successfully.")
        # new_invoice() # Clear the form after generating the document

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while generating the document: {e}")


# Main window UI
ctk.set_appearance_mode("dark")  # Options: "System" (default), "Light", "Dark"
ctk.set_default_color_theme("blue")  # Options: "blue" (default), "green", "dark-blue"



window = ctk.CTk()
window.title("Invoice Generator Form")
window.minsize(800, 600)

# Create a tab view with two tabs
tabview = ctk.CTkTabview(window, width=950, height=550)
tabview.pack(padx=20, pady=20)

# Add two tabs to the tabview
tab_invoice_Gen = tabview.add("Invoice Generator")
tab_clients = tabview.add(" Clients")
tab_history = tabview.add(" History")

frame = ctk.CTkFrame(tab_invoice_Gen)
frame.pack(padx=20, pady=20)

doc_type_option = ctk.IntVar()

# Customer Information Frame
customer_info_frame = ctk.CTkFrame(frame, corner_radius=10 )
customer_info_frame.grid(row=0, column=0, padx=12, pady=12)

ctk.CTkLabel(customer_info_frame, text="Customer Information", 
             font=("Arial", 16, "bold")).grid(row=0, column=0, columnspan=2, pady=10)
ctk.CTkLabel(customer_info_frame, text="Company:").grid(row=1, column=0)
ctk.CTkLabel(customer_info_frame, text="First Name:").grid(row=2, column=0)
ctk.CTkLabel(customer_info_frame, text="Last Name:").grid(row=3, column=0)
ctk.CTkLabel(customer_info_frame, text="Phone:").grid(row=4, column=0)

# Load the JSON data from the file and create a list for the dropdown
with open('clients.json', 'r') as file:
    data = json.load(file)

# Access the clients and extract company names
clients = [client['company_name'] for client in data['clients']]

# Create a dropdown (ComboBox) for company names with a specified width
company_name_entry = ctk.CTkComboBox(customer_info_frame, values=clients, width=200)
company_name_entry.grid(row=1, column=1, pady=5)

# Extract the first client's information (or any specific client's information)
first_client = data['clients'][0]

# Create entry boxes and set their initial text with the first client's information
first_name_entry = ctk.CTkEntry(customer_info_frame)
first_name_entry.grid(row=2, column=1, pady=5)
# first_name_entry.insert(0, first_client['first_name'])

last_name_entry = ctk.CTkEntry(customer_info_frame)
last_name_entry.grid(row=3, column=1, pady=5)
# last_name_entry.insert(0, first_client['last_name'])

phone_entry = ctk.CTkEntry(customer_info_frame)
phone_entry.grid(row=4, column=1, pady=5)
# phone_entry.insert(0, first_client['phone_number'])

# Function to update customer information based on selected company
def update_customer_info():
    selected_company = company_name_entry.get()
    print(f"Selected company: {selected_company}")  # Debugging line
    for client in data['clients']:
        if client['company_name'] == selected_company:
            print(f"Updating entries for: {client}")  # Debugging line
            first_name_entry.delete(0, ctk.END)
            first_name_entry.insert(0, client['first_name'])
            last_name_entry.delete(0, ctk.END)
            last_name_entry.insert(0, client['last_name'])
            phone_entry.delete(0, ctk.END)
            phone_entry.insert(0, client['phone_number'])
            break

# Button to trigger the update_customer_info function
select_customer_button = ctk.CTkButton(customer_info_frame, text="Select Customer", command=update_customer_info)
select_customer_button.grid(row=5, column=1, pady=5)

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

# Client Tab
client_frame = ctk.CTkFrame(tab_clients)
client_frame.pack(padx=20, pady=20)
ctk.CTkLabel(client_frame, text="Client Information", font=("Arial", 16, "bold")).grid(row=0, column=0, columnspan=2, pady=10)

ctk.CTkLabel(client_frame, text="Company:").grid(row=1, column=0)
ctk.CTkLabel(client_frame, text="First Name:").grid(row=2, column=0)
ctk.CTkLabel(client_frame, text="Last Name:").grid(row=3, column=0)
ctk.CTkLabel(client_frame, text="Phone:").grid(row=4, column=0)

cl_company_entry = ctk.CTkEntry(client_frame)
cl_company_entry.grid(row=1, column=1, pady=5)
cl_first_name_entry = ctk.CTkEntry(client_frame)
cl_first_name_entry.grid(row=2, column=1, pady=5)
cl_last_name_entry = ctk.CTkEntry(client_frame)
cl_last_name_entry.grid(row=3, column=1, pady=5)
cl_phone_entry = ctk.CTkEntry(client_frame)
cl_phone_entry.grid(row=4, column=1, pady=5)

# Save client button
def save_client():
    company = cl_company_entry.get()
    first_name = cl_first_name_entry.get()
    last_name = cl_last_name_entry.get()
    phone = cl_phone_entry.get()

    if not company or not first_name or not last_name or not phone:
        messagebox.showerror("Input Error", "All fields must be filled out.")
        return

    new_client = {
        "company_name": company,
        "first_name": first_name,
        "last_name": last_name,
        "phone_number": phone
    }

    try:
        with open('clients.json', 'r+') as file:
            data = json.load(file)
            data['clients'].append(new_client)
            file.seek(0)
            json.dump(data, file, indent=4)

        client_tree.insert('', 'end', values=(company, first_name, last_name, phone))
        cl_company_entry.delete(0, ctk.END)
        cl_first_name_entry.delete(0, ctk.END)
        cl_last_name_entry.delete(0, ctk.END)
        cl_phone_entry.delete(0, ctk.END)
        
        # Update the company_name_entry values
        company_name_entry.configure(values=[client['company_name'] for client in data['clients']])
        
        messagebox.showinfo("Success", "Client added successfully.")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while saving the client: {e}")

save_client_button = ctk.CTkButton(client_frame, text="Save Client", command=save_client)
save_client_button.grid(row=5, column=0, columnspan=2, sticky="news", padx=20, pady=5)
# Function to delete selected client
def delete_client():
    selected_item = client_tree.selection()
    if not selected_item:
        messagebox.showerror("Selection Error", "Please select a client to delete.")
        return

    try:
        with open('clients.json', 'r+') as file:
            data = json.load(file)
            selected_client = client_tree.item(selected_item, 'values')
            data['clients'] = [client for client in data['clients'] if not (
                client['company_name'] == selected_client[0] and
                client['first_name'] == selected_client[1] and
                client['last_name'] == selected_client[2] and
                client['phone_number'] == selected_client[3]
            )]
            file.seek(0)
            file.truncate()
            json.dump(data, file, indent=4)

        client_tree.delete(selected_item)
        messagebox.showinfo("Success", "Client deleted successfully.")
        
        # Update the company_name_entry values
        company_name_entry.configure(values=[client['company_name'] for client in data['clients']])

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while deleting the client: {e}")

delete_client_button = ctk.CTkButton(client_frame, text="Delete Client", command=delete_client)
delete_client_button.grid(row=7, column=0, columnspan=2, sticky="news", padx=20, pady=5)

# Function to edit selected client
def edit_client():
    selected_item = client_tree.selection()
    if not selected_item:
        messagebox.showerror("Selection Error", "Please select a client to edit.")
        return

    selected_client = client_tree.item(selected_item, 'values')
    cl_company_entry.delete(0, ctk.END)
    cl_company_entry.insert(0, selected_client[0])
    cl_first_name_entry.delete(0, ctk.END)
    cl_first_name_entry.insert(0, selected_client[1])
    cl_last_name_entry.delete(0, ctk.END)
    cl_last_name_entry.insert(0, selected_client[2])
    cl_phone_entry.delete(0, ctk.END)
    cl_phone_entry.insert(0, selected_client[3])

    def save_edited_client():
        new_company = cl_company_entry.get()
        new_first_name = cl_first_name_entry.get()
        new_last_name = cl_last_name_entry.get()
        new_phone = cl_phone_entry.get()

        if not new_company or not new_first_name or not new_last_name or not new_phone:
            messagebox.showerror("Input Error", "All fields must be filled out.")
            return

        try:
            with open('clients.json', 'r+') as file:
                data = json.load(file)
                for client in data['clients']:
                    if (client['company_name'] == selected_client[0] and
                        client['first_name'] == selected_client[1] and
                        client['last_name'] == selected_client[2] and
                        client['phone_number'] == selected_client[3]):
                        client['company_name'] = new_company
                        client['first_name'] = new_first_name
                        client['last_name'] = new_last_name
                        client['phone_number'] = new_phone
                        break
                file.seek(0)
                file.truncate()
                json.dump(data, file, indent=4)

            client_tree.item(selected_item, values=(new_company, new_first_name, new_last_name, new_phone))
            cl_company_entry.delete(0, ctk.END)
            cl_first_name_entry.delete(0, ctk.END)
            cl_last_name_entry.delete(0, ctk.END)
            cl_phone_entry.delete(0, ctk.END)
            
            # Update the company_name_entry values
            company_name_entry.configure(values=[client['company_name'] for client in data['clients']])
            
            messagebox.showinfo("Success", "Client edited successfully.")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while editing the client: {e}")

    save_client_button.configure(command=save_edited_client)
    save_client_button.configure(text="Save Changes")

edit_client_button = ctk.CTkButton(client_frame, text="Edit Client", command=edit_client)
edit_client_button.grid(row=8, column=0, columnspan=2, sticky="news", padx=20, pady=5)
# Treeview for clients using ttk
columns = ('company', 'first_name', 'last_name', 'phone')
client_tree = ttk.Treeview(client_frame, columns=columns, show="headings")
client_tree.heading('company', text='Company')
client_tree.heading('first_name', text='First Name')
client_tree.heading('last_name', text='Last Name')
client_tree.heading('phone', text='Phone')
client_tree.grid(row=6, column=0, columnspan=2, padx=20, pady=10)

# Load the JSON data from the file
with open('clients.json', 'r') as file:
    data = json.load(file)

# Access the clients
clients = data['clients']

# Example of inserting the first client into the tree
for client in clients:
    client_tree.insert('', 'end', 
                       values=(client['company_name'],
                               client['first_name'],
                               client['last_name'],
                               client['phone_number']))


window.mainloop()

