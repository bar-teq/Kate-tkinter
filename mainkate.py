import imaplib
import os
import email
import time
import io
import subprocess
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
root = Tk()
# root.geometry('500x500')
root.columnconfigure(0, minsize=10)
root.columnconfigure(1, minsize=40)
root.columnconfigure(2, minsize=10)
root.columnconfigure(3, minsize=40)
root.columnconfigure(4, minsize=10)

root.rowconfigure(0, minsize=10)
root.rowconfigure(1, minsize=40)
root.rowconfigure(2, minsize=40)
root.rowconfigure(3, minsize=40)
root.rowconfigure(4, minsize=40)
root.rowconfigure(5, minsize=40)
root.rowconfigure(6, minsize=10)


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
    listbox_list_of_items.delete(0, 99)
    for i in range(int(id_list[-1]), int(id_list[0])-1, -1):
        _, data = mail.fetch(str(i), '(RFC822)')
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
                    # file.write(f"{besbefore.lstrip()} \n***************** \n")
                    listbox_list_of_items.insert(0, " ")
                    listbox_list_of_items.insert(0, f"{besbefore.lstrip()}")

                for item in lista:
                    with open('items.txt', 'a', encoding="UTF-8") as file:
                        # file.write(f"{item.lstrip()} \n")
                        listbox_list_of_items.insert(0, f" * {item.lstrip()}")

                with open('items.txt', 'a', encoding="UTF-8") as file:
                    # file.write(f"\nPACZKA {i} \n")
                    listbox_list_of_items.insert(0, f"PACZKA {i}")
        mail.store(str(i), '-FLAGS', '(\Seen)')


def print_all():
    files_deleting()
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
    for f in os.listdir(labels_dir):
        filePath = os.path.join(labels_dir, f)
        subprocess.call(f'PDFtoPrinter.exe /s {filePath}', shell=True)
        # os.system(f'{filePath} print')
        # os.startfile(filePath, "print")
        time.sleep(1)
        # command = "{} {}".format('PDFtoPrinter.exe','report.pdf')


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
    # label1 = Label(root, text=f" Masz {how_many_new_emails()} nowych etykiet")
    # label1.grid(row=0, column=0)
    label_how_many_new_emails = Label(
        root, text=f" Nowych etykiet: {how_many_new_emails()}")
    label_how_many_new_emails.grid(row=1, column=3)


button_printall = Button(root, width=40, padx=0,
                         pady=40, text="DRUKUJ WSZYSTKO", command=lambda: print_all())
button_refresh = Button(root, width=40, padx=0, pady=0,
                        text="Odswież listę", command=lambda: refresh())
button_schedule = Button(root, width=10, padx=0, pady=0, text="Grafik?")
listbox_list_of_items = Listbox(root, width=40, height=30)
label_how_many_new_emails = Label(
    root, text=f" Nowych etykiet: {how_many_new_emails()}")
listbox_list_of_items.grid(row=1, column=1, rowspan=5)
button_printall.grid(row=3, column=3, sticky=NS)
label_how_many_new_emails.grid(row=1, column=3)
button_refresh.grid(row=4, column=3, sticky=N)
button_schedule.grid(row=5, column=3)
label_bottom = Label(root, text="by carmel")
label_bottom.grid(row=6, column=2, columnspan=3, sticky=E)

# E.grid(row=0, column=0)
# F.grid(row=1, column=0)


# def print_days():
#     # lista = e.get()
#     # days = lista.split(',')
#     mylabel = Label(root, text=E.get())
#     mylabel.pack()
#     mylabel2 = Label(root, text=F.get())
#     mylabel2.pack()


# files_deleting()

# my_text = Listbox(root, width=40, height=20)
# my_text.grid(row=2, column=0)
# label1 = Label(root, text=f" Masz {how_many_new_emails()} nowych etykiet")
# label1.grid(row=0, column=0)

# button = Button(root, text="Drukuj wszystkie!", command=lambda: print_all())
# button.grid(row=1, column=0)
# button1 = Button(root, text="odswiez", command=lambda: refresh())
# button1.grid(row=3, column=0)
root.mainloop()
