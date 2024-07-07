import pyautogui as ag
import time
import pytesseract
import pyperclip as pc
import numpy
import cv2
import flet as ft
import openpyxl as xl
from pathlib import Path
import win32com.client
import os
import xlwings as xw
import pythoncom

#指定した起動中のアプリ選択
def move(key):
    KEY=str(key)
    ag.hotkey("win", KEY)

#アプリのカーソル位置取得
def operation():
    P1 = ag.position()
    time.sleep(2)
    while True:
        P2 = ag.position()
        if P1==P2:
            break
        time.sleep(1)
        P1 = P2
    return P1

#文字入力
def moji(mozi):
    pc.copy(mozi)
    ag.hotkey('ctrl', 'v')

#マウス右クリック
def Rclick():
    ag.rightClick()

#マウス左クリック
def Lclick():
    ag.leftClick()

#ドラッグドロップ
def drag_drop(Pos1, Pos2):
    ag.move(Pos1)
    ag.mouseDown()
    ag.move(Pos2)
    ag.mouseUp()

#スクロール
def scroll(ryou):
    ag.hscroll(ryou)

#コピー
def copy():
    ag.hotkey("ctrl", "C")

#貼り付け
def paste():
    ag.hotkey("ctrl", "V")

#Excelを開く
def xls_open(xlpath):
    wb = xl.load_workbook(xlpath)
    return wb

#新しいExcelを作成
def new_xls(Xname):
    wb = xl.Workbook()
    wb.save(Xname)
    return wb

#シート選択
def xls_sheet(sheet, wb):
    ws = wb[sheet]
    return ws

#Excelに書き込む
def xls_write(X, y, mo, wb, ws, xlpath):
    Y = ord(y)
    Y = Y-64
    ws.cell(row=X, column=Y, value=mo)
    wb.save(xlpath)

#マクロ読み込み
def read_vba_macro(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        vba_code = file.read()
    return vba_code

#VBAマクロ追加
def add_vba_macro_excel(excel_file_path, vba_macro, module_name='Module1'):
    # Excelを起動
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    
    # 指定したExcelファイルを開く
    workbook = excel.Workbooks.Open(excel_file_path)
    
    # モジュールが既に存在するかチェック
    vb_component = None
    for vb_comp in workbook.VBProject.VBComponents:
        if vb_comp.Name == module_name:
            vb_component = vb_comp
            break
    
    if vb_component:
        # 既存のモジュールがある場合、そのコードを更新
        vb_component.CodeModule.DeleteLines(1, vb_component.CodeModule.CountOfLines)
        vb_component.CodeModule.AddFromString(vba_macro)
    else:
        # 新しいモジュールを追加
        vb_component = workbook.VBProject.VBComponents.Add(1)  # 1:標準モジュール
        vb_component.Name = module_name
        vb_component.CodeModule.AddFromString(vba_macro)
    
    workbook.Save()
    workbook.Close()

#VBA起動
def ex_vba_macro(excel_file_path):

    app = xw.App()

    wb = app.books.open(excel_file_path)

    pg = wb.macro("A1_in_test")

    pg()

    wb.save(excel_file_path)

    wb.close()
    app.quit()



#UI
def main(page: ft.Page):
    page.title = "AutoGUI"  # タイトル

    locate = []
    count = []

    #位置情報の追加
    def carsol_posi_clicked(e):
        posi = operation()
        locate.append(posi)
        kazu = len(locate) - 1
        locate_row.controls.append(ft.Text(kazu))
        locate_row.controls.append(ft.Text("="))
        locate_row.controls.append(ft.Text(posi))
        locate_row.controls.append(ft.Text("   "))
        page.update()


    #動作系↓

    #Excel関係↓
    #新しいエクセル作成
    def newxl_ST(e):
        name = newxl_name.value

        #チャートに追加
        count.append('NEWXL')
        count.append(name)
        count_column.controls.append(ft.Row([ft.Text(name), ft.Text(".xlsmを作成")]))
        page.update()

    #VBA取得
    def VBA_select_click(e: ft.FilePickerResultEvent):
        if e.files:
            count.append('VBA')
            count.append(e.files[0].path)
            count_column.controls.append(ft.Row([ft.Text(e.files[0].path), ft.Text("を")]))
            VBA_select.visible = False
            excel_select.visible = True
            page.update()




    def excel_select_click(e: ft.FilePickerResultEvent):
        if e.files:
            count.append(e.files[0].path)
            count_column.controls.append(ft.Row([ft.Text(e.files[0].path), ft.Text("に適応")]))
            VBA_select.visible = True
            excel_select.visible = False
            page.update()
            

    #起動フラグ　新しいエクセル作成(NEWXL,エクセルの名前)　VBA取得(VBA,マクロテキストのパス,エクセルのパス)
    def start_button_click(e):
        count.append("END")
        i=0
        while True:
            if count[i]=="END":
                break
            elif count[i]=="NEWXL":
                i=i+1
                Xname = count[i]+".xlsm"
                cd = Path.cwd()
                path = cd / Xname
                new_xls(path)
                print(count[i])
                i=i+1
            elif count[i]=="VBA":
                #print("DEBUG--------------------------------------")
                i=i+1
                macro_path = count[i]
                i=i+1
                excel_path = count[i]

                print(macro_path)
                print(excel_path)
                
                # VBAマクロの読み込み
                vba_macro = read_vba_macro(macro_path)

                pythoncom.CoInitialize()  # Excelを起動する前にこれを呼び出す
                # VBAマクロをExcelに追加
                add_vba_macro_excel(excel_path, vba_macro)

                #  VBAマクロを起動
                ex_vba_macro(excel_path)
                pythoncom.CoUninitialize()  # Excelを終了した後はこれを呼び出す


                i=i+1
            else:
                break
        
                


    carsol_posi = ft.FilledTonalButton(text="カーソルの位置を取得", on_click=carsol_posi_clicked)

    pick_files_dialog = ft.FilePicker(on_result=VBA_select_click)
    page.overlay.append(pick_files_dialog)
    pick_files_dialog2 = ft.FilePicker(on_result=excel_select_click)
    page.overlay.append(pick_files_dialog2)

    VBA_select = ft.ElevatedButton(text="VBAを取得(txt)", icon = ft.icons.FILE_OPEN, on_click=lambda _: pick_files_dialog.pick_files(allow_multiple=False))
    excel_select = ft.ElevatedButton(text="excelのpathを取得(xlsm)", icon = ft.icons.FILE_OPEN, on_click=lambda _: pick_files_dialog2.pick_files(allow_multiple=False), visible=False, color=ft.colors.GREEN)

    start_button = ft.FilledTonalButton(text="start", on_click=start_button_click)



    #Excel関係
    newxl_name = ft.TextField(hint_text="新しく作るExcelの名前")
    newxl_name_button = ft.FilledTonalButton(text="Add", on_click=newxl_ST)

    #位置情報一覧
    locate_row = ft.Row()
    #チャート
    count_column = ft.Column()


    page.add(
        ft.Text("位置情報記録", color=ft.colors.BLUE),
        locate_row,
        ft.Column(
            [
                ft.Text("追加したい操作を選ぶ"),
                ft.Row([carsol_posi]),

                ft.Text("Excel関連", color=ft.colors.GREEN),
                ft.Row([newxl_name, newxl_name_button]),
                ft.Row([VBA_select, excel_select])
            ]
        ),
        ft.Text("チャート", color=ft.colors.RED),
        count_column,
        #起動ボタン
        ft.Column(
            [start_button]
        )
    )



ft.app(target=main)