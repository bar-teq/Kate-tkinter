from ast import Global
import imaplib
import os
import email
import time
import io
from email import policy
import mailparser
from time import strftime
from dotenv import load_dotenv
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from tkinter import *

attachment_dir = os.getenv('ATTACHMENT_DIR')
labels_dir = os.getenv('LABEL_DIR')

id_list = []
mail = 0
open("items.txt", "w").close()


def how_many_new_emails():
    global id_list, mail
    load_dotenv()
    LOGIN = os.getenv('LOGIN')
    PASSWORD = os.getenv('PASSWORD')
    IMAP_URL = os.getenv('IMAP_URL')
    mail = imaplib.IMAP4_SSL(IMAP_URL)
    mail.login(LOGIN, PASSWORD)
    mail.select('ETYKIETY')
    typ, data = mail.search(None, 'UNSEEN')
    mail_ids = data[0]
    id_list = mail_ids.split()
    if len(id_list) == 0:
        return "brak"
    else:
        for i in range(int(id_list[-1]), int(id_list[0])-1, -1):
            mail.store(str(i), '-FLAGS', '(\Seen)')
        print(id_list)

        return id_list


def show_items():
    for i in range(int(id_list[-1]), int(id_list[0])-1, -1):
        _, data = mail.fetch(str(i), '(RFC822)')
        raw = email.message_from_bytes(data[0][1])
        mailp = mailparser.parse_from_bytes(data[0][1])
        for response_part in data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                temat = mailp.subject
                print(temat)
                besbefore = temat[-24:]
                crop = temat[0:-46]
                lista = crop.split(',')
                print(crop)
                with open('items.txt', 'a', encoding="UTF-8") as file:
                    file.write(f"\nPACZKA {i} \n")
                for item in lista:
                    with open('items.txt', 'a', encoding="UTF-8") as file:
                        file.write(f"{item.lstrip()} \n")
                with open('items.txt', 'a', encoding="UTF-8") as file:
                    file.write(f"{besbefore.lstrip()} \n***************** \n")
    with open('items.txt', 'r', encoding="UTF-8") as file:
        stuff = file.read()
        my_text.insert(END, stuff)


def change_pdf(title, fileName):
    filePath_lab = os.path.join(labels_dir, fileName)
    filePath_att = os.path.join(attachment_dir, fileName)
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter, encrypt=None)
    can.setFont(psfontname="DejaVuSans", size=14, leading=None)
    can.drawString(0, 100, f"{title}")
    can.save()

    packet.seek(0)
    new_pdf = PdfFileReader(packet)
    with open(filePath_att, "rb") as f:
        existing_pdf = PdfFileReader(f)
        output = PdfFileWriter()
        page = existing_pdf.getPage(0)
        page.mergePage(new_pdf.getPage(0))
        output.addPage(page)
        outputStream = open(filePath_lab, "wb")
        output.write(outputStream)
        outputStream.close()


root = Tk()
root.title('KATE')
root.geometry("500x500")
label1 = Label(root, text=len(how_many_new_emails()))
label1.grid(row=0, column=0)
my_text = Text(root)
my_text.grid(row=1, column=0)
show_items()

root.mainloop()
