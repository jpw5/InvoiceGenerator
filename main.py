import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox


def clear_item():
    qty_spinbox.delete(0, tk.END)
    qty_spinbox.insert(0, "1")
    desc_entry.delete(0, tk.END)
    price_spinbox.delete(0, tk.END)
    price_spinbox.insert(0, "0.0")


invoice_list = []


def add_item():
    try:
        qty = int(qty_spinbox.get())
        desc = desc_entry.get()
        price = float(price_spinbox.get())
        new_price = round(price, 2)
        price = new_price
        line_total = qty * price
        print(line_total)
        invoice_item = [qty, desc, price, line_total]
        tree.insert('', 0, values=invoice_item)
        clear_item()
        invoice_list.append(invoice_item)
    except:
        messagebox.showwarning(title='Error', message='Missing quantity or price. ')


def new_invoice():
    first_name_entry.delete(0, tk.END)
    last_name_entry.delete(0, tk.END)
    phone_entry.delete(0, tk.END)
    clear_item()
    tree.delete(*tree.get_children())

    invoice_list.clear()


def generate_invoice():
    doc = DocxTemplate("invoice_template.docx")
    name = f'{first_name_entry.get()} {last_name_entry.get()}'
    phone = phone_entry.get()
    subtotal = sum(item[3] for item in invoice_list)
    salestax = 0.1
    total = subtotal * (1 - salestax)

    doc.render({"name": name,
                "phone": phone,
                "invoice_list": invoice_list,
                "subtotal": f'${round(subtotal, 2)}',
                "salestax": str(salestax * 100) + "%",
                "total": f'${round(total, 2)}'})

    doc_name = f'{name}\'s New Invoice_{datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S")}.docx'
    doc.save(doc_name)

    messagebox.showinfo("Invoice Complete", "Invoice Complete")

    new_invoice()


root = ctk.CTk()
root.resizable(0, 0)
root.title("Invoice Generator Form")

frame = tk.Frame(root)
frame.pack(padx=20, pady=10)

first_name_label = ctk.CTkLabel(frame, text="First Name")
first_name_label.grid(row=0, column=0)
last_name_label = ctk.CTkLabel(frame, text="Last Name")
last_name_label.grid(row=0, column=1)

first_name_entry = ctk.CTkEntry(frame)
last_name_entry = ctk.CTkEntry(frame)
first_name_entry.grid(row=1, column=0)
last_name_entry.grid(row=1, column=1)

phone_label = ctk.CTkLabel(frame, text="Phone")
phone_label.grid(row=0, column=2)
phone_entry = ctk.CTkEntry(frame)
phone_entry.grid(row=1, column=2)

qty_label = ctk.CTkLabel(frame, text="Qty")
qty_label.grid(row=2, column=0)
qty_spinbox = ctk.CTkEntry(frame)
qty_spinbox.grid(row=3, column=0)

desc_label = ctk.CTkLabel(frame, text="Description")
desc_label.grid(row=2, column=1)
desc_entry = ctk.CTkEntry(frame)
desc_entry.grid(row=3, column=1)

price_label = ctk.CTkLabel(frame, text="Unit Price")
price_label.grid(row=2, column=2)
price_spinbox = ctk.CTkEntry(frame)
price_spinbox.grid(row=3, column=2)

add_item_button = ctk.CTkButton(frame, text="Add item", command=add_item)
add_item_button.grid(row=4, column=0, padx=20, pady=5, sticky='news', columnspan=3)
columns = ('qty', 'desc', 'price', 'total')
tree = ttk.Treeview(frame, columns=columns, show="headings")
tree.heading('qty', text='Qty')
tree.heading('desc', text='Description')
tree.heading('price', text='Unit Price')
tree.heading('total', text="Total")

tree.grid(row=5, column=0, columnspan=3, padx=20, pady=10)

save_invoice_button = ctk.CTkButton(frame, text="Generate Invoice", command=generate_invoice)
save_invoice_button.grid(row=6, column=0, columnspan=3, sticky="news", padx=20, pady=5)
new_invoice_button = ctk.CTkButton(frame, text="New Invoice", command=new_invoice)
new_invoice_button.grid(row=7, column=0, columnspan=3, sticky="news", padx=20, pady=5)

root.mainloop()
