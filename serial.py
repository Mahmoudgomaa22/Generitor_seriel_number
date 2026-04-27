import random
import string
import os
from barcode import Code128
from barcode.writer import ImageWriter
from openpyxl import Workbook
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4



PREFIX = "WGZ-"
RANDOM_LENGTH = 8
COUNT = 100   # عدد السيريالات
OUTPUT_FOLDER = "output"

# A4 Labels (40 label)
COLS = 4
ROWS = 10

LABEL_WIDTH = 140
LABEL_HEIGHT = 75

START_X = 15
START_Y = 770



if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)


generated = set()

def generate_serial():
    while True:
        code = ''.join(random.choices(string.ascii_uppercase + string.digits, k=RANDOM_LENGTH))
        serial = PREFIX + code
        if serial not in generated:
            generated.add(serial)
            return serial


wb = Workbook()
ws = wb.active
ws.title = "Serials"
ws.append(["Serial Number"])

serials = []

print("Generating serials...")

for i in range(COUNT):
    serial = generate_serial()
    serials.append(serial)
    ws.append([serial])

    # barcode image
    barcode = Code128(serial, writer=ImageWriter())
    barcode.save(f"{OUTPUT_FOLDER}/{serial}")

# save excel
wb.save(f"{OUTPUT_FOLDER}/WGZ_Serials.xlsx")


# Create PDF Labels


pdf = canvas.Canvas(f"{OUTPUT_FOLDER}/WGZ_Labels.pdf", pagesize=A4)

x = START_X
y = START_Y

index = 0

for serial in serials:
    img_path = f"{OUTPUT_FOLDER}/{serial}.png"

    pdf.drawImage(img_path, x, y-35, width=120, height=35)
    pdf.setFont("Helvetica-Bold", 8)
    pdf.drawString(x + 15, y - 45, serial)

    index += 1

    if index % COLS == 0:
        x = START_X
        y -= LABEL_HEIGHT
    else:
        x += LABEL_WIDTH

    if index % (COLS * ROWS) == 0:
        pdf.showPage()
        x = START_X
        y = START_Y

pdf.save()

print("Done Successfully!")
print("Files saved in output folder.")