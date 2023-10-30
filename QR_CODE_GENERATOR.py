import base64
import pyqrcode

import pandas as pd
from PIL import Image
from PIL import ImageDraw, ImageFont, ImageFilter

font = ImageFont.truetype("fonts/Lobster.ttf", 48)

df = pd.read_excel("C:/Users/FPTSHOP/Desktop/DS_DangKiThamGiaChaoTanK65.xlsx")

excelFrame = pd.DataFrame()
ids = []
names = []
classes = []
urls = []

count = 0
turn = 1

urlTemplate = "https://raw.githubusercontent.com/VienCNTTKTS/NEUSmartCheckin/main/events/Ch%C3%A0o%20t%C3%A2n%20K65%20Vi%E1%BB%87n%20CNTT%26KTS/"

target = "parents"

size = len(df.index)
count = 0

for i in range(len(df.index)):
    # if (size > 30 and count == 30) or (30 > size and size - 1 == count):
    #     excelFrame["Mã sinh viên"] = ids
    #     excelFrame["Họ tên"] = names
    #     excelFrame["Lớp"] = classes
    #     excelFrame["Mã QR"] = urls
    #
    #     excelFrame.to_excel(str(turn) + ".xlsx")
    #
    #     excelFrame = pd.DataFrame()
    #
    #     ids.clear()
    #     names.clear()
    #     classes.clear()
    #     urls.clear()
    #
    #     turn += 1
    #     count = 0
    #     size -= 30
    #
    #     print(size)

    record = df.iloc[i]
    #
    # image = Image.open("thumoiph.png")
    # canvas = ImageDraw.Draw(image)
    #
    # offsetTen = (905 / 2) - (font.getlength(record.Ten)/2)
    # offsetLop = (750 / 2) - (font.getlength(record.Lop) / 2)
    #
    # canvas.text((106 + offsetTen, 460), record.Ten, font=font, fill=(52, 58, 108), align="center")
    # canvas.text((193 + offsetLop, 600), record.Lop, font=font, fill=(52, 58, 108), align="center")



    # data = "SV_" + str(record.MSV)
    data = "SV_11236202"
    encodedData = base64.b64encode(data.encode('ascii'))

    # ids.append(str(record.MSV))
    # names.append(record.Ten)
    # classes.append(record.Lop)
    # urls.append(urlTemplate + "parents/" + encodedData.decode('utf-8').replace('=', '%3D') + ".png")
    #
    # count += 1
    qrCode = pyqrcode.create(encodedData)
    qrCode.png("data/" + data + ".png", scale=6, )
    count += 1
    print(str(count) + "/" + str(size) + " " + data)
    #
    # image.paste(Image.open("data/" + data + ".png").resize(size=(128*2, 128*2)), (1371, 925))
    #
    # image.save("thumoi/" + data + ".png")