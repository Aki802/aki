import flet as ft
import pyocr
from PIL import Image, ImageEnhance
import os
from pathlib import Path

cd = Path.cwd()

Tes = r"Tesseract-OCR"
tess = r"Tesseract-OCR\tessdata"

TESSERACT_PATH = cd / Tes
TESSDATA_PATH = cd / tess

TESSERACT_PATH = str(TESSERACT_PATH)
TESSDATA_PATH = str(TESSDATA_PATH)

print(TESSERACT_PATH)
print(TESSDATA_PATH)


os.environ["PATH"] += os.pathsep + TESSERACT_PATH
os.environ["TESSDATA_PREFIX"] = TESSDATA_PATH

#OCRエンジン取得
tools = pyocr.get_available_tools()
tool = tools[0]

#OCRの設定 ※tesseract_layout=6が精度には重要。デフォルトは3
builder = pyocr.builders.TextBuilder(tesseract_layout=6)



#UI
def main(page: ft.Page):
    page.title = "OCRAI"
    page.bgcolor = ft.colors.WHITE


    def ocr(e: ft.FilePickerResultEvent):
        file_name = e.files[0].path
        #解析画像読み込み
        img = Image.open(file_name) 

        #適当に画像処理
        img_g = img.convert('L') #Gray変換
        enhancer= ImageEnhance.Contrast(img_g) #コントラストを上げる
        img_con = enhancer.enhance(2.0) #コントラストを上げる

        #画像からOCRで日本語を読んで、文字列として取り出す
        la = langa.value

        txt_pyocr = tool.image_to_string(img_con , lang=la, builder=builder)

        #半角スペースを消す ※読みやすくするため
        txt_pyocr = txt_pyocr.replace(' ', '')

        Text_ex.value = txt_pyocr

        page.update()


    pick_files_TXT = ft.FilePicker(on_result=ocr)
    page.overlay.append(pick_files_TXT)


    langa = ft.Dropdown(label="言語を選択", width=120, options=[ft.dropdown.Option('jpn'), ft.dropdown.Option('eng')])

    FileOpen = ft.ElevatedButton(text="Fileを選択", icon = ft.icons.AD_UNITS_SHARP, on_click=lambda _: pick_files_TXT.pick_files(allow_multiple=False))

    Text_ex = ft.TextField(multiline=True,)



    maincon = ft.Column([
        ft.Row([FileOpen, ft.Text("言語を選択"), langa]),
        Text_ex
        ],expand=True)
    
    page.add(maincon)


ft.app(target=main)