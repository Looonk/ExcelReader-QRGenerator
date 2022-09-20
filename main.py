import os
import pandas as pd
import qrcode
import shutil
from PIL import Image, ImageDraw, ImageFont

folder = os.listdir("./in")
files = os.listdir("./in/"+folder[0])
os.chdir("./in/"+folder[0]+"/")
for f in files:
    if f.endswith(".xls") or f.endswith(".XLS"):
        df = pd.read_excel(f, usecols="B,D", keep_default_na=False, engine="xlrd")
        #df = pd.read_excel(f, usecols="A,G", keep_default_na=False, engine="xlrd")
        for i in df.itertuples():
            if not i.__contains__(""):
                if str(i.__getattribute__("_1")).__contains__("*"):
                    aux = str(i.__getattribute__("_1")).replace("*", "")
                else:
                    aux = str(i.__getattribute__("_1"))
                if i.__getattribute__("_2").__contains__("\""):
                    a = i.__getattribute__("_2").replace("\"", " PULGADAS ")
                    name = aux + " - " + a + ".jpg"
                elif i.__getattribute__("_2").__contains__("´´"):
                    a = i.__getattribute__("_2").replace("´´", " PULGADAS")
                    name = aux + " - " + a + ".jpg"
                elif i.__getattribute__("_2").__contains__("/"):
                    a = i.__getattribute__("_2").replace("/", "ON")
                    name = aux + " - " + a + ".jpg"
                else:
                    name = aux + " - " + i.__getattribute__("_2") + ".jpg"
                img = qrcode.make(i.__getattribute__("_1"))
                os.makedirs("../../output", exist_ok=True)
                os.makedirs("../../output/" + folder[0] + "/" + str(f)[:-4], exist_ok=True)
                img.save("../../output/" + folder[0] + "/" + str(f)[:-4] + "/" + name)

                img = Image.open("../../output/" + folder[0] + "/" + str(f)[:-4] + "/" + name)
                rgb_img = img.convert("RGB")
                rgb_img.save("../../output/" + folder[0] + "/" + str(f)[:-4] + "/" + name[:-4] + " not.png")

                canvas = Image.new("RGB", (290, 320), "white")
                img_draw = ImageDraw.Draw(canvas)
                font = ImageFont.truetype("../../times.ttf", size=19)
                img_draw.text((110, 300), str(i.__getattribute__("_1")), fill="black", font=font)
                canvas.save("../../output/" + folder[0] + "/" + str(f)[:-4] + "/" + name[:-4] + " bg.png")

                image = Image.open("../../output/" + folder[0] + "/" + str(f)[:-4] + "/" + name[:-4] + " not.png")
                marco = Image.open("../../output/" + folder[0] + "/" + str(f)[:-4] + "/" + name[:-4] + " bg.png")
                image_copy = marco.copy()

                image_copy.paste(image, None)
                image_copy.save("../../output/" + folder[0] + "/" + str(f)[:-4] + "/" + name[:-4] + ".png")

                os.remove("../../output/" + folder[0] + "/" + str(f)[:-4] + "/" + name)
                os.remove("../../output/" + folder[0] + "/" + str(f)[:-4] + "/" + name[:-4] + " not.png")
                os.remove("../../output/" + folder[0] + "/" + str(f)[:-4] + "/" + name[:-4] + " bg.png")

        os.remove("../../output/" + folder[0] + "/" + str(f)[:-4] + "/Código - Descripción.png")
os.chdir("../../")
shutil.make_archive("./output/" + folder[0], "zip", "./output/" + folder[0])
