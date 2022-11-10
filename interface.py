import os
from tkinter import *

from parser import Parser

def browseFiles():
    parser.browseFiles()
def startEvent():
    parser.startEvent()
def stopEvent():
    parser.stopEvent()

class UserApp(Tk):
    def __init__(self):
        super().__init__()
        self.lblhead = Label(text="", background="#323232", pady=5)
        self.lblhead.grid(column=0, row=0, )

        self.lbl = Label(self, background="#323232", foreground="white", text="Путь к рабочему файлу", pady=5)
        self.lbl.grid(column=0, row=1, )
        self.filePath = Entry(self, width=38)
        self.filePath.grid(column=1, row=1)
        self.btn = Button(self, cursor="hand2", text="поиск", command=browseFiles)
        self.btn.grid(column=2, row=1)

        self.lblReserve = Label(self, background="#323232",  foreground="white", text="Путь к резервной папке", pady=5)
        self.lblReserve.grid(column=0, row=2)
        self.filePathReserve = Entry(self, width=38)
        self.filePathReserve.grid(column=1, row=2)
        self.filePathReserve.insert(0, r'C:\Users' '\\' + os.getenv('USERNAME') + '\AppData\Roaming\Microsoft\Excel' '\\')

        self.dataPlaceLbl = Label(self, background="#323232",  foreground="white", text="Последняя ячейка данных", pady=5)
        self.dataPlaceLbl.grid(column=0, row=3)
        self.dataPlace = Entry(self, width=10)
        self.dataPlace.grid(column=1, row=3)

        self.lblManList = Label(self, background="#323232", foreground="white", text="Имя листа мужских соревнований", pady=5)
        self.lblManList.grid(column=0, row=4)
        self.manList = Entry(self, width=10)
        self.manList.grid(column=1, row=4)
        self.btnUpdMLN = Button(self, cursor="hand2", text="обновить", command=self.updDict)
        self.btnUpdMLN.grid(column=2, row=4)

        self.lblWomanList = Label(self, background="#323232",  foreground="white", text="Имя листа женских соревнований", pady=5)
        self.lblWomanList.grid(column=0, row=5)
        self.womanList = Entry(self, width=10)
        self.womanList.grid(column=1, row=5)

        self.btnStart = Button(self, cursor="hand2", text="начать", command=startEvent)
        self.btnStart.grid(column=0, row=6)

        self.btnStop = Button(self, cursor="hand2", text="остановить", command=stopEvent)
        self.btnStop.grid_remove()

        self.errorMsg = Label(self, foreground="red", text="")
        self.errorMsg.grid_remove()

        self.dictionary = {}
        self.maxData = ''
        self.manListName = ''
        self.womanListName = ''
        self.getDict()
        global parser
        parser = Parser(self, self.btnStop, self.filePath, self.filePathReserve, self.maxData, self.manListName, self.womanListName)

    def getDict(self):
        try:
            with open('dataFile.txt') as file:
                lines = file.read().splitlines()

            for line in lines:
                key, value = line.split(': ')
                self.dictionary.update({key: value})

            self.maxData = self.dictionary.get('maxData')
            self.manListName = self.dictionary.get('manList')
            self.womanListName = self.dictionary.get('womanList')
            self.dataPlace.insert(0, self.maxData)
            self.manList.insert(0, self.manListName)
            self.womanList.insert(0, self.womanListName)
        except:
            print("Используются стандартные настройки листов")

    def updDict(self):
        if self.dataPlace.get() != '':
            self.dictionary.update({'maxData': self.dataPlace.get()})
        if self.manList.get() != '':
            self.dictionary.update({'manList': self.manList.get()})
        if self.womanList.get() != '':
            self.dictionary.update({'womanList': self.womanList.get()})

        file = open('dataFile.txt', 'w')
        for key, value in self.dictionary.items():
            file.write(f'{key}: {value}\n')
        file.close()

    def errorMessageNoFile(self):
        stopEvent()
        self.errorMsg = Label(self, background="#323232", foreground="red", text="Ошибка чтения файла")
        self.errorMsg.grid(column=0, row=6)
        browseFiles()

    def errorMessageNoSheet(self):
        stopEvent()
        self.errorMsg = Label(self, background="#323232", foreground="red", text="Проверьте имена листов")
        self.errorMsg.grid(column=0, row=6)

    def errorMessageAdr(self):
        self.errorMsg = Label(self, foreground="red", text="Проверьте указанный адрес ячейки")
        self.errorMsg.grid(column=0, row=6)

    def updAll(self):
        self.dataPlace.delete(0, END)
        self.manList.delete(0, END)
        self.womanList.delete(0, END)
        self.maxData = self.dictionary.get('maxData')
        self.manListName = self.dictionary.get('manList')
        self.womanListName = self.dictionary.get('womanList')
        self.dataPlace.insert(0, self.maxData)
        self.manList.insert(0, self.manListName)
        self.womanList.insert(0, self.womanListName)

    def clearFileName(self):
        self.filePath.delete(0, END)
        self.updAll()



if __name__ == "__main__":
    app = UserApp()
    app.title("")
    app["bg"] = "#323232"
    app.resizable(0,0)
    width = app.winfo_screenwidth()
    height = app.winfo_screenheight()
    x = (width / 2) - (width / 6)
    y = (height / 2) - (height / 8)
    app.geometry('%dx%d+%d+%d' % (width / 3, height / 4, x, y))


    app.mainloop()


