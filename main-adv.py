from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import pytesseract
from pytesseract import Output
import cv2


def name_ext(ocr_dir, text):
    pytesseract.pytesseract.tesseract_cmd = ocr_dir

    img = cv2.imread("output/" + text + ".png")

    d = pytesseract.image_to_data(img, output_type=Output.DICT)
    for i in d['text']:
        # print(i)
        if i == str(text).replace(" ", "-"):
            ind1 = d['text'].index(i)
            x1 = d['left'][ind1]
            y1 = d['top'][ind1]
            w1 = d['width'][ind1]
            h1 = d['height'][ind1]
            x2 = x1 + w1
            y2 = y1 + h1

    dist = ((((x2 - x1)**2) + ((y2-y1)**2))**0.5)

    return dist


def extractText(ocr_dir, template_image):
    x1, y1, x2, y2 = "", "", "", ""
    pytesseract.pytesseract.tesseract_cmd = ocr_dir

    img = cv2.imread(template_image)

    d = pytesseract.image_to_data(img, output_type=Output.DICT)

    for i in d['text']:
        # print(i)
        if i == "$":
            ind1 = d['text'].index(i)
            x1 = d['left'][ind1]
            y1 = d['top'][ind1]
            w1 = d['width'][ind1]
            h1 = d['height'][ind1]

        elif i == "#":
            ind2 = d["text"].index(i)
            x2 = d['left'][ind2]
            y2 = d['top'][ind2]
            w2 = d['width'][ind2]
            h2 = d['height'][ind2]

    # cv2.rectangle(img, (x1, y1), (x1 + w1, y1 + h1), (0, 255, 0), 2)
    # cv2.rectangle(img, (x2, y2), (x2 + w2, y2 + h2), (0, 255, 0), 2)
    #
    # cv2.imshow('img', img)
    # cv2.waitKey(0)

    if x1 == "" or x2 == "" or y1 == "" or y2 == "":
        print("Error : Cannot identify position !")
        exit()

    return ((x1, y1), (x2, y2))


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
                bimg = Image.open("program_assets/blank.png")
                draw = ImageDraw.Draw(img)
                draw2 = ImageDraw.Draw(bimg)
                print("Writing name : {}".format(j['NAME']))
                start_ind = extractText(conf_local['ocr_dir'], conf_local['template_image'])[0]
                end_ind = extractText(conf_local['ocr_dir'], conf_local['template_image'])[1]
                # print(start_ind)
                # print(end_ind)
                temp_name = str(j['NAME']).replace(" ", "-")
                draw2.text(xy=(0, 0), text='{}'.format(temp_name), fill=(0, 0, 0), font=font)
                bimg.save(conf_local['project_output_dir'] + "{}.png".format(j['NAME']))
                str_coord = name_ext(conf_local['ocr_dir'], j['NAME'])
                dist = ((((end_ind[0] - start_ind[0])**2) + ((end_ind[1]-start_ind[1])**2))**0.5)

                answer = dist - str_coord
                answer = answer/2
                x = start_ind[0] + answer
                y = start_ind[1]

                draw.text(xy=(x, y), text='{}'.format(j['NAME']), fill=(0, 0, 0), font=font)
                img.save(conf_local['project_output_dir'] + "{}.png".format(j['NAME']))

                # break
        except Exception as e:
            print("Error : ", e)
            exit()

    else:
        print("Please input (", conf_local["ext"], ") image")

# ---------------------------------------------------------------------------------------------------------------------------------------------


size = 30
font_style = 'arial.ttf'
# configuration data
conf = {
    "ext": ".png",
    "design_image": "design1.png",
    "template_image": "template1.png",
    "data_file": ("Data.xlsx", "Sheet1"),
    "project_output_dir": "output/",
    "ocr_dir": r"C:\Program Files\Tesseract-OCR\tesseract.exe"
}


# extractText(conf['ocr_dir'], conf['template_image'])
writeName(font_style, size, conf)
