from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import pytesseract
from pytesseract import Output
import cv2


def extractText(ocr_dir, template_image):
    pytesseract.pytesseract.tesseract_cmd = ocr_dir

    img = cv2.imread(template_image)

    d = pytesseract.image_to_data(img, output_type=Output.DICT)
    x, y = "", ""
    for i in d['text']:
        if i == "$":
            ind = d['text'].index(i)
            x = d['left'][ind]
            y = d['top'][ind]

    if x == "" or y == "":
        print("Sorry. Template Invalid !!")
        exit()

    return (x, y)


def writeName(font_style_local, size_local, conf_local):
    # reading excel
    try:
        df = pd.read_excel(conf_local["data_file"][0], sheet_name=conf_local["data_file"][1], engine='openpyxl')
    except Exception as e:
        print("Error : ", e)
        exit()

    # setting font styles and size
    font = ImageFont.truetype(font_style_local, size_local)
    newImg = conf_local["design_image"]  # Design

    # Automating names into original certificate
    if newImg.endswith(conf_local["ext"]):
        try:
            for index, j in df.iterrows():
                img = Image.open(newImg)
                draw = ImageDraw.Draw(img)
                print("Writing name : {}".format(j['NAME']))
                draw.text(xy=extractText(conf_local['ocr_dir'], conf_local['template_image']), text='{}'.format(j['NAME']), fill=(0, 0, 0), font=font)
                img.save(conf_local['project_output_dir'] + "{}.png".format(j['NAME']))
        except Exception as e:
            print("Error : ", e)
            exit()

    else:
        print("Please input (", conf_local["ext"], ") image")

# ---------------------------------------------------------------------------------------------------------------------------------------------


size = 20
font_style = 'arial.ttf'
# configuration data
conf = {
    "ext": ".png",
    "design_image": "design.png",
    "template_image": "template.png",
    "data_file": ("Data.xlsx", "Sheet1"),
    "project_output_dir": "output/",
    "ocr_dir": r"C:\Program Files\Tesseract-OCR\tesseract.exe"
}

writeName(font_style, size, conf)
