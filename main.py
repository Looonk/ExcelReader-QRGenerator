import os
import pandas as pd
import qrcode
import shutil
from PIL import Image, ImageDraw, ImageFont

l = os.listdir()
for f in l:
    if f.endswith(".xls") or f.endswith(".XLS"):
        df = pd.read_excel(f, usecols="B,D", keep_default_na=False, engine='xlrd')
        # df = pd.read_excel(f, usecols="A,G", keep_default_na=False, engine='xlrd')
        for i in df.itertuples():
            if not i.__contains__(""):
                if i.__getattribute__("_2").__contains__("\""):
                    a = i.__getattribute__("_2").replace("\"", " PULGADAS ")
                    name = str(i.__getattribute__("_1")) + " - " + a + ".jpg"
                elif i.__getattribute__("_2").__contains__("/"):
                    a = i.__getattribute__("_2").replace("/", "ON")
                    name = str(i.__getattribute__("_1")) + " - " + a + ".jpg"
                else:
                    name = str(i.__getattribute__("_1")) + " - " + i.__getattribute__("_2") + ".jpg"
                img = qrcode.make(i.__getattribute__("_1"))
                os.makedirs("output", exist_ok=True)
                os.makedirs("./output/" + str(f)[:-4], exist_ok=True)
                img.save("./output/" + str(f)[:-4] + "/" + name)

                img = Image.open("./output/" + str(f)[:-4] + "/" + name)
                rgb_img = img.convert('RGB')
                rgb_img.save("./output/" + str(f)[:-4] + "/" + name[:-4] + " not.png")

                canvas = Image.new('RGB', (290, 320), 'white')
                img_draw = ImageDraw.Draw(canvas)
                font = ImageFont.truetype("times.ttf", size=19)
                img_draw.text((110, 300), str(i.__getattribute__("_1")), fill='black', font=font)
                canvas.save("./output/" + str(f)[:-4] + "/" + name[:-4] + " bg.png")

                image = Image.open("./output/" + str(f)[:-4] + "/" + name[:-4] + " not.png")
                marco = Image.open("./output/" + str(f)[:-4] + "/" + name[:-4] + " bg.png")
                image_copy = marco.copy()

                image_copy.paste(image, None)
                image_copy.save("./output/" + str(f)[:-4] + "/" + name[:-4] + ".png")

                os.remove("./output/" + str(f)[:-4] + "/" + name)
                os.remove("./output/" + str(f)[:-4] + "/" + name[:-4] + " not.png")
                os.remove("./output/" + str(f)[:-4] + "/" + name[:-4] + " bg.png")

        os.remove("./output/" + str(f)[:-4] + "/Código - Descripción.png")

        shutil.make_archive("./output/" + str(f)[:-4], 'zip', "./output/" + str(f)[:-4])