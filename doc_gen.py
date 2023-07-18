import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox

doc = DocxTemplate("invoice_template.docx")

invoice_list = [[2, "pen", 0.5, 1],
                [1, "paper", 5, 5],
                [2, "notebook", 2, 4]]

doc.render({"name": "John Rodgers",
            "phone": "123-456-7890",
            "invoice_list": invoice_list,
            "subtotal": "10%",
            "total": 9
            })
doc.save("new_invoice.docx")

# root = ctk.CTk()
# root.title("Invoice Generator")
#
#
# root.mainloop()
