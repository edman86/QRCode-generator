# Import modules

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm, Pt
import qrcode
import tkinter as tk
from tkinter import ttk
import pyperclip
import requests
from bs4 import BeautifulSoup


# App Business logic --------------------------------------------------------------------

# Functions -----------------------------------------------------------

def getDataFromWebPage(link):
    """Parse web page and get data of item name and item price,
       then return them in array."""

    url = link
    itemDataArr = []

    r = requests.get(url)

    try:
        r.raise_for_status()
    except Exception as exc:
        print("No connect!")

    soup = BeautifulSoup(r.text, 'html.parser')

    itemName = soup.select(".product-name")[0]
    itemPrice = soup.select(".product-inner-price ins")[0]
    
    itemDataArr.append(itemName.getText())
    itemDataArr.append(itemPrice.getText())
    
    return itemDataArr


# Event Listeners ------------------------------------------------------

def generate_qr():
    """ All QR Code generating logic""" 
    
    # Generating of the QR Code Segment -----------------------
    
    # Getting the data from entries
    data_link = entry_link.get()
    data_name = entry_name.get()
    data_price = entry_price.get()
    qr_color = entry_qr_color.get()

    # If entry for color data is empty
    if qr_color == "": 
        qr_color = "#26bae0"

    data = data_link + "?QR" + "\n\n" + data_name + "\n\n" + data_price

    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )

    qr.add_data(data)
    qr.make(fit=True)

    img = qr.make_image(fill_color=qr_color, back_color="white")

    img.save("qr_image.png")

    # Generating of the Word Document Segment ---------------- 

    tpl = DocxTemplate("temp.docx")
    myimage = InlineImage(tpl, "qr_image.png", width=Mm(50))

    context = { 
        "heading": data_name, 
        "img": myimage
    }

    tpl.render(context)
    tpl.save("generated_doc.docx")

    # Showing text message to user
    lbl_message["text"] = "Фаил сгенерирован в директорию программы"



def clear():
    """Clear all entries and text message"""

    entry_link.delete(0, tk.END)
    entry_name.delete(0, tk.END)
    entry_price.delete(0, tk.END)
    entry_qr_color.delete(0, tk.END)
    
    lbl_message["text"] = ""



def autoFill():
    """Fills data automatically"""
 
    # get link from the clipboard
    link = pyperclip.paste()

    # get data from web page
    itemDataArr = getDataFromWebPage(link)

    # set standart color for QR Code
    qrColor = "#26bae0"

    data_link = entry_link.insert(0, link)
    data_name = entry_name.insert(0, itemDataArr[0])
    data_price = entry_price.insert(0, itemDataArr[1])
    data_color = entry_qr_color.insert(0, qrColor)



# Creating GUI ---------------------------------------------------------------------------

root = tk.Tk()
root.title("QR Code Generator")
root.geometry("600x500")

# Style for buttons
style = ttk.Style()
style.configure("TButton", padding=6, relief="flat", font="Arial 20")

# Разделение окна на верхнюю и нижнюю области.
# Верхняя область
top_frame = ttk.Frame(root)
top_frame.pack()
# Нижняя область
bottom_frame = ttk.Frame(root)
bottom_frame.pack(side=tk.BOTTOM)

# label "enter the link"
lbl_link = ttk.Label(top_frame, font=("Arial", 20), text="Введите ссылку")
lbl_link.grid(row=0, columnspan=2)

entry_link = ttk.Entry(top_frame, font=("Arial", 20), width=30)
entry_link.grid(row=1)

# label "enter name"
lbl_name = ttk.Label(top_frame, font=("Arial", 20), text="Введите название")
lbl_name.grid(row=2, columnspan=2)

entry_name = ttk.Entry(top_frame, font=("Arial", 20), width=30)
entry_name.grid(row=3)

# label "enter the price"
lbl_price = ttk.Label(top_frame, font=("Arial", 20), text="Введите цену")
lbl_price.grid(row=4, columnspan=2)

entry_price = ttk.Entry(top_frame, font=("Arial", 20), width=30)
entry_price.grid(row=5)

# label "enter QR Code color"
lbl_qr_color = ttk.Label(top_frame, font=("Arial", 20), text="Введите цвет QR Кода")
lbl_qr_color.grid(row=6, columnspan=2)

entry_qr_color = ttk.Entry(top_frame, font=("Arial", 20), width=30)
entry_qr_color.grid(row=7)

# This label shows message when the QR Code is create
lbl_message = tk.Label(top_frame, font=("Arial", 20), fg="red", text="")
lbl_message.grid(row=8)

# Auto data filling button 
btn_autoFill = ttk.Button(bottom_frame, text="Заполнить автоматически", style="TButton", command=autoFill) #command
btn_autoFill.grid(row=0, columnspan=2, sticky="ew") 

# Main buttons
btn_generate = ttk.Button(bottom_frame, text="Generate", style="TButton", command=generate_qr) #command
btn_generate.grid(row=1, column=0, sticky="ew")

btn_reset = ttk.Button(bottom_frame, text="Reset", style="TButton", command=clear ) #command
btn_reset.grid(row=1, column=1, sticky="ew")

# This label creates space between buttons and bottom border of window
lbl_space = tk.Label(bottom_frame, font=("Ubuntu", 40))
lbl_space.grid(row=2, columnspan=2)

# This label shows tip
lbl_tip = ttk.Label(bottom_frame, font=("Arial", 10), text="Для копирования используйте Ctrl + C\nДля вставки используйте Ctrl + V")
lbl_tip.grid(row=3, columnspan=2)


root.mainloop()