import xlsxwriter
from PIL import Image
from webcolors import *

path = "" #Enter your path

fileName = input()
scale = int(input())

img_path1   =   path + fileName + ".jpg"
xlsx_path   =   path + fileName + ".xlsx"

workbook = xlsxwriter.Workbook(xlsx_path)
worksheet1 = workbook.add_worksheet()
worksheet2 = workbook.add_worksheet()

img = Image.open(img_path1)
width, height = img.size

for i in range(height//scale):
    worksheet1.set_row(i, 1)
worksheet1.set_column(0, width//scale - 1, 0.1)

for i in range(0, height // scale):
    for j in range(0, width // scale):
        color = str(rgb_to_hex(img.getpixel((j*scale, i*scale))).upper())
        cell_format = workbook.add_format({
            "bg_color" : color
        })
        worksheet1.write(i, j, "", cell_format)


workbook.close()