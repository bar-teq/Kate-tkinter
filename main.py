import imaplib
import os
import email
import time
import subprocess
import io
from email import policy
from tkinter import *
import mailparser
from time import strftime
from dotenv import load_dotenv
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

pdfmetrics.registerFont(TTFont('DejaVuSans', 'DejaVuSans.ttf'))
load_dotenv()
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
    show_items()

    if len(id_list) == 0:
        return "0"
    else:
        for i in range(int(id_list[-1]), int(id_list[0])-1, -1):
            mail.store(str(i), '-FLAGS', '(\Seen)')

        return len(id_list)


def show_items():
    if len(id_list) == 0:
        return "brak"
    for i in range(int(id_list[-1]), int(id_list[0])-1, -1):
        _, data = mail.fetch(str(i), '(RFC822)')
        my_text.delete(0, 99)
        raw = email.message_from_bytes(data[0][1])
        mailp = mailparser.parse_from_bytes(data[0][1])
        for response_part in data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                temat = mailp.subject
                besbefore = temat[-24:]
                crop = temat[0:-46]
                lista = crop.split(',')
                with open('items.txt', 'a', encoding="UTF-8") as file:
                    file.write(f"{besbefore.lstrip()} \n***************** \n")
                    my_text.insert(0, " ")
                    my_text.insert(0, f"{besbefore.lstrip()}")

                for item in lista:
                    with open('items.txt', 'a', encoding="UTF-8") as file:
                        file.write(f"{item.lstrip()} \n")
                        my_text.insert(0, f" * {item.lstrip()}")

                with open('items.txt', 'a', encoding="UTF-8") as file:
                    file.write(f"\nPACZKA {i} \n")
                    my_text.insert(0, f"PACZKA {i}")
        mail.store(str(i), '-FLAGS', '(\Seen)')


def print_all():
    if len(id_list) == 0:
        return "brak"
    for i in range(int(id_list[-1]), int(id_list[0])-1, -1):
        _, data = mail.fetch(str(i), '(RFC822)')
        raw = email.message_from_bytes(data[0][1])
        mailp = mailparser.parse_from_bytes(data[0][1])
        for part in raw.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            fileName = part.get_filename()
            filePath_lab = os.path.join(labels_dir, fileName)
            filePath_att = os.path.join(attachment_dir, fileName)

            if bool(fileName):
                with open(filePath_att, 'wb') as f:
                    f.write(part.get_payload(decode=True))
        mail.store(str(i), '+FLAGS', '(\Seen)')
        mail.store(str(i), '+X-GM-LABELS', '\\Trash')
        for response_part in data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                temat = mailp.subject
                packet = io.BytesIO()
                can = canvas.Canvas(packet, pagesize=A4, encrypt=None)
                can.setFont(psfontname="DejaVuSans", size=12, leading=None)
                besbefore = temat[-24:]
                crop = temat[0:-46]
                lista = crop.split(',')
                lista.reverse()
                x = 120

                for i in lista:
                    can.drawString(10, x, i.lstrip())
                    x += 20
                can.setFont(psfontname="DejaVuSans", size=10, leading=None)
                can.drawString(10, 90, f"{besbefore}")

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
                time.sleep(5)


def files_deleting():
    for f in os.listdir(attachment_dir):
        try:
            os.remove(os.path.join(attachment_dir, f))
        except PermissionError:
            pass

    for f in os.listdir(labels_dir):
        try:
            os.remove(os.path.join(labels_dir, f))
        except PermissionError:
            pass


def refresh():
    label1 = Label(root, text=f" Masz {how_many_new_emails()} nowych etykiet")
    label1.grid(row=0, column=0)


files_deleting()
root = Tk()
root.title('KATE')
my_text = Listbox(root, width=40, height=20)
my_text.grid(row=2, column=0)
label1 = Label(root, text=f" Masz {how_many_new_emails()} nowych etykiet")
label1.grid(row=0, column=0)
show_items()
button = Button(root, text="Drukuj wszystkie!", command=lambda: print_all())
button.grid(row=1, column=0)
button1 = Button(root, text="odswiez", command=lambda: refresh())
button1.grid(row=3, column=0)
root.mainloop()
