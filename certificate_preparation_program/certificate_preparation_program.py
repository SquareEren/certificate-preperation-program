from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
import openpyxl
import subprocess

# Upload your certificate design.
foto = Image.open("certificate.png")

# Adjust the font settings for the certificate.
font_path = "font.ttf" # Write your font folder here.
font_size = 124

font = ImageFont.truetype("font.ttf", 124, encoding="utf-8")

#Upload your excel file.
wb = openpyxl.load_workbook("data.xlsx")

# Select The Page
sheet = wb["list"]

# Read the names and save them on the list.
names = []
for row in sheet.iter_rows(values_only=True):
    names.append(row[0])
# create the certificate for each name.
for name in names:
    foto = Image.open("certificate.png")
    draw = ImageDraw.Draw(foto)
    if name is not None:
        text_width, text_height = draw.textbbox((0, 0), name, font=font)[2:]
        x = (foto.width - text_width) / 2
        y = 575
        color = (178, 133, 68)
        draw.text((x, y), name, font=font, fill=color)
        foto.save(f"{name}.png")
    else:
        break