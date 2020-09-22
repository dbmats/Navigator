from tkinter import *
import pandas as pd
import xlwings as xw
import win32gui
from pyxlsb import open_workbook

def select(event):
    wb = xw.apps.active
    tt = wb.books
    tt1 = pd.DataFrame(tt)
    tt1[1] = ((tt1[0].astype('str')).str.split('[').str.get(1)).str[0:-2]
    tt1.loc[tt1[1].astype('str').str[0:8] == 'PERSONAL', 1] = 0
    tt1.loc[tt1[1].astype('str').str[0:8] == 'Personal', 1] = 0
    tt1 = tt1.loc[tt1[1] != 0]
    f = list(tt1[1])
    list1 = f
    listbox1.delete('0', 'end')
    for i in list1:
        listbox1.insert(END, i)


def select_item(event):
    wb1 = (listbox1.get(listbox1.curselection()))
    ww = wb1.split('.')[0]
    def window_enum_handler(hwnd, resultList):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
            resultList.append((hwnd, win32gui.GetWindowText(hwnd)))

    def get_app_list(handles=[]):
        mlst = []
        win32gui.EnumWindows(window_enum_handler, handles)
        for handle in handles:
            mlst.append(handle)
        return mlst

    appwindows = dict(get_app_list())
    for value in appwindows.values():
        if ww in value:
            ww2 = value
    ww1 = win32gui.FindWindow(None, ww2)
    win32gui.SetForegroundWindow(ww1)
    if (wb1[-1] != 'b' and wb1[-1] != 'B'):
        t = xw.books(wb1)
        ttt = t.__str__().replace('<Book ', '').replace('>', '')
        kk = '<Sheet ' + ttt
        ss = xw.sheets
        ss1 = pd.DataFrame(ss)
        ss1[0] = ss1[0].astype('str').str[0:-1]
        ss1 = ss1[0].tolist()
        ss1 = ss1.__str__().replace('Sheets([', '').replace('])', '').replace(kk, '') \
            .replace("['", '').replace("']", '').replace("', '", ',')
        ss1 = list(ss1.split(','))
        listbox2.delete('0', 'end')
        for i in ss1:
            listbox2.insert(END, i)
    else:
        t = xw.books(wb1)
        listbox2.delete('0', 'end')
        with open_workbook(t.fullname) as wb:
            for sheetname in wb.sheets:
                listbox2.insert(END, sheetname)


def select_item1(event):
    wb3 = (listbox2.get(listbox2.curselection()))
    t12 = xw.books.active
    t13 = t12.__str__().replace('<Book [', '').replace(']>', '')
    ww = t13.split('.')[0]
    def window_enum_handler(hwnd, resultList):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
            resultList.append((hwnd, win32gui.GetWindowText(hwnd)))

    def get_app_list(handles=[]):
        mlst = []
        win32gui.EnumWindows(window_enum_handler, handles)
        for handle in handles:
            mlst.append(handle)
        return mlst

    appwindows = dict(get_app_list())
    for value in appwindows.values():
        if ww in value:
            ww2 = value
    ww1 = win32gui.FindWindow(None, ww2)
    win32gui.SetForegroundWindow(ww1)
    xw.Book(t13).sheets[wb3].activate()


def sel():
    s = var.get() / 100
    root.wm_attributes("-alpha", s)


root = Tk()
root.iconbitmap(default="rainbow sphere.ico")
root.title("#")
root.geometry('180x1105+5+5')
root.wm_attributes("-alpha", 0.45)
root.attributes("-topmost", True)
listbox1 = Listbox(width=45, height=10)
listbox2 = Listbox(width=45, height=51)
button1 = Button(root, text='Обновить список файлов', width=20, height=1)

var = DoubleVar()
scale = Scale(root, variable=var, orient=HORIZONTAL, from_=45)
scale.pack(anchor=CENTER)
button = Button(root, text="Прозрачно. сдвинь и нажми", command=sel)
button.pack(anchor=CENTER)
label = Label(root)
label.pack()

listbox1.bind('<<ListboxSelect>>', select_item)
listbox2.bind('<<ListboxSelect>>', select_item1)
button1.bind("<Button-1>", select)

button1.pack()
listbox1.pack()
listbox2.pack()
root.mainloop()
