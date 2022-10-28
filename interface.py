import os
from tkinter import *

from parser import Parser

def browseFiles():
    parser.browseFiles()
def startEvent():
    parser.startEvent()
def stopEvent():
    parser.stopEvent()

window = Tk()
window.title("")
width = window.winfo_screenwidth()
height = window.winfo_screenheight()
window.geometry('%dx%d' % (width/3, height/4))

lbl = Label(window, text="Путь к файлу")
lbl.grid(column=0, row=0)
filePath = Entry(window, width=10)
filePath.grid(column=1, row=0)
btn = Button(window, text="поиск", command=browseFiles)
btn.grid(column=2, row=0)

dataPlaceLbl = Label(window, text="Диапазон данных")
dataPlaceLbl.grid(column=0, row=1)
dataPlace = Entry(window, width=10)
dataPlace.grid(column=1, row=1)


varGender = IntVar()
varGender.set(0)
rbtnMan = Radiobutton(text='Мужские соревнования', variable=varGender, value=0)
rbtnWoman = Radiobutton(text='Женские соревнования', variable=varGender, value=1)
rbtnMan.grid(column=0, row=2)
rbtnWoman.grid(column=1, row=2)

btnStart = Button(window, text="начать", command=startEvent)
btnStart.grid(column=0, row=3)

btnStop = Button(window, text="остановить", command=stopEvent)
btnStop.grid_remove()

lblReserve = Label(window, text="Путь к папке")
lblReserve.grid(column=0, row=4)
filePathReserve = Entry(window, width=10)
filePathReserve.grid(column=1, row=4)
filePathReserve.insert(0, r'C:\Users' '\\' + os.getenv('USERNAME') + '\AppData\Roaming\Microsoft\Excel' '\\')

parser = Parser(btnStop, filePath, varGender, filePathReserve, dataPlace)

window.mainloop()

