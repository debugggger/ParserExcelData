import ast
import os
import pickle
import tkinter
from tkinter import *

from dictionary import ExcDictionary
from parser import Parser

def browseFiles():
    parser.browseFiles()
def startEvent():
    parser.startEvent()
def stopEvent():
    parser.stopEvent()

class UserApp(tkinter.Tk):
    def __init__(self):
        super().__init__()

window = Tk()
window.title("")
width = window.winfo_screenwidth()
height = window.winfo_screenheight()
window.geometry('%dx%d' % (width/3, height/4))

dictionary = {}
maxData = 'z210'
manListName = 'Рез_Муж'
womanListName = 'Рез_Жен'
try:
    with open('dataFile.txt') as file:
        lines = file.read().splitlines()

    for line in lines:
        key, value = line.split(': ')
        dictionary.update({key: value})

    maxData = dictionary.get('maxData')
    manList = dictionary.get('manList')
    womanList = dictionary.get('womanList')
except:
    print("Используются стандартные настройки листов")


lbl = Label(window, text="Путь к файлу")
lbl.grid(column=0, row=0)
filePath = Entry(window, width=40)
filePath.grid(column=1, row=0)
btn = Button(window, text="поиск", command=browseFiles)
btn.grid(column=2, row=0)

dataPlaceLbl = Label(window, text="Последняя ячейка данных")
dataPlaceLbl.grid(column=0, row=1)
dataPlace = Entry(window, width=10)
dataPlace.grid(column=1, row=1)

btnStart = Button(window, text="начать", command=startEvent)
btnStart.grid(column=0, row=3)

btnStop = Button(window, text="остановить", command=stopEvent)
btnStop.grid_remove()

lblReserve = Label(window, text="Путь к папке")
lblReserve.grid(column=0, row=4)
filePathReserve = Entry(window, width=40)
filePathReserve.grid(column=1, row=4)
filePathReserve.insert(0, r'C:\Users' '\\' + os.getenv('USERNAME') + '\AppData\Roaming\Microsoft\Excel' '\\')

lblManList = Label(window, text="Имя листа мужских соревнований")
lblManList.grid(column=0, row=5)
manList = Entry(window, width=10)
manList.grid(column=1, row=5)

lblWomanList = Label(window, text="Имя листа женских соревнований")
lblWomanList.grid(column=0, row=6)
womanList = Entry(window, width=10)
womanList.grid(column=1, row=6)

if dataPlace.get() != '':
    dictionary.update({'maxData': dataPlace.get()})
else:
    dictionary.update({'maxData': maxData})
if manList.get() != '':
    dictionary.update({'manList': manList.get()})
else:
    dictionary.update({'manList': manListName})
if womanList.get() != '':
    dictionary.update({'womanList': womanList.get()})
else:
    dictionary.update({'womanList': womanListName})

file = open('dataFile.txt', 'w')
for key, value in dictionary.items():
    file.write(f'{key}: {value}\n')
file.close()

manList.insert(0, dictionary.get('manList'))
dataPlace.insert(0, dictionary.get('maxData'))
womanList.insert(0, dictionary.get('womanList'))

parser = Parser(btnStop, filePath, filePathReserve, dataPlace, manList, womanList)

window.mainloop()

if __name__ == "__main__":
    window.mainloop()


