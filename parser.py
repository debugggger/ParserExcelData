import os
import threading
import time
import pandas as pd
import openpyxl

from threading import Thread
from tkinter import filedialog

from peoplesData import People

stopParsingThread = threading.Event()

class Parser(object):
    def __init__(self, btnStop, filePath, varGender, filePathReserve, dataPlace):
        self.fileName = ''
        self.peoples = []
        self.t1 = 0.0
        self.btnStop = btnStop
        self.filePath = filePath
        self.varGender = varGender
        self.filePathReserve = filePathReserve
        self.dataPlace = dataPlace

    def startEvent(self):
        self.t1 = self.getChangeTime()
        threadParsing = Thread(target=self.parsingData)
        threadParsing.start()
        self.btnStop.grid(column=1, row=3)


    def stopEvent(self):
        stopParsingThread.set()
        self.btnStop.grid_remove()

    def browseFiles(self):
        self.fileName = filedialog.askopenfilename(initialdir="/",
                                              title="Select a File",
                                              filetypes=(("Excel files",
                                                          "*.xls*"),
                                                         ("all files",
                                                          "*.*")))
        self.getSheetName()
        self.filePath.insert(0, self.fileName)

    def getSheetName(self):
        sheetName = ""
        if self.varGender.get() == 0:
            sheetName = 'Рез_Муж'
        if self.varGender.get() == 1:
            sheetName = 'Рез_Жен'
        return sheetName

    def getReserveFile(self):
        try:
            reservePath = self.filePathReserve.get()
            name = (self.fileName.split('/')[-1]).split('.')[0]

            checkT = 0
            maxTime = 0.01
            truePath = ''
            for dir in os.listdir(reservePath):
                if len(name) <= len(dir):
                    for i in range(len(name)):
                        if name[i] == dir[i]:
                            checkT = 1
                            truePath = reservePath + dir
                        else:
                            checkT = 0
                            break

                if checkT == 1:
                    if os.path.getmtime(reservePath + dir) > maxTime:
                        maxTime = os.path.getmtime(reservePath + dir)
                        truePath = reservePath + dir + '\\'

            trueFile = ''
            maxTimeFile = 0.01
            for file in os.listdir(truePath):
                current = truePath + file
                if os.path.getmtime(current) > maxTimeFile:
                    if current.split('.')[-1] == 'xlsb':
                        maxTimeFile = os.path.getmtime(current)
                        trueFile = current
                    if current.split('.')[-1] == 'xls':
                        if len(current) >= len(trueFile):
                            maxTimeFile = os.path.getmtime(current)
                            trueFile = current


            eng = ''
            if trueFile.split('.')[-1] == 'xls':
                eng = 'xlrd'

            if trueFile.split('.')[-1] == 'xlsb':
                eng = 'pyxlsb'

            sheetName = self.getSheetName()
            df = pd.read_excel(trueFile, engine=eng, sheet_name=sheetName)
            nameTrueFile = (trueFile.split('/')[-1]).split('.')[0] + '.xlsx'
            df.to_excel(nameTrueFile)

            xlsx = openpyxl.load_workbook(nameTrueFile, data_only=True)
            sheetName = 'Sheet1'
            sheet = xlsx[sheetName]
            sheet.delete_cols(1, 1)
            xlsx.save(nameTrueFile)
            xlsx.close()
            return nameTrueFile
        except:
            return ''

    def getChangeTime(self):
        self.t1 = os.path.getmtime(self.fileName)

    def parsingData(self):
        prevFile = ''
        while True:

            reserveFile = self.getReserveFile()

            if self.t1 != os.path.getmtime(self.fileName):
                if self.fileName.split('.')[-1] == 'xls':
                    reservePath = self.filePathReserve.get()
                    df = pd.read_excel(self.fileName, engine='xlrd', sheet_name=self.getSheetName())
                    nameTrueFile = (self.fileName.split('/')[-1]).split('.')[0] + '.xlsx'
                    patchXlsx = reservePath+nameTrueFile
                    df.to_excel(patchXlsx)
                    sheetName = 'Sheet1'
                    xlsx = openpyxl.load_workbook(patchXlsx, data_only=True)
                    sheet = xlsx.get_sheet_by_name(sheetName)
                    sheet.delete_cols(1, 1)
                    xlsx.save(patchXlsx)
                    self.readData(patchXlsx, sheetName)
                    os.remove(patchXlsx)
                else:
                    self.readData(self.fileName, self.getSheetName())
                self.getChangeTime()

            if reserveFile != prevFile:
                prevFile = reserveFile
                self.readData(reserveFile, "Sheet1")

            if stopParsingThread.is_set():
                break
            time.sleep(10)

    def readData(self,nameReadFile, sheetName):
        self.peoples = []
        extension = nameReadFile.split('.')[-1]
        if extension == 'xlsx':
            xlsx = openpyxl.load_workbook(nameReadFile, data_only=True)
            sheet = xlsx[sheetName]
            startPlace = self.dataPlace.get().split(':')[0]
            lastPlace = self.dataPlace.get().split(':')[-1]
            for cellObj in sheet[startPlace:lastPlace]:
                currentPeople = People("", "", "", "", "", "", "", "", "", "", "", "", "")
                firstColumnInd = cellObj[0].column
                currentColumn = firstColumnInd
                for cell in cellObj:
                    if currentColumn == firstColumnInd+1:
                        currentPeople.place = cell.value
                    if currentColumn == firstColumnInd + 3:
                        currentPeople.name = cell.value
                    if currentColumn == firstColumnInd + 4:
                        currentPeople.year = cell.value
                    if currentColumn == firstColumnInd + 5:
                        currentPeople.discharge = cell.value
                    if currentColumn == firstColumnInd + 6:
                        currentPeople.city = cell.value
                    if currentColumn == firstColumnInd + 9:
                        currentPeople.school = cell.value
                    if currentColumn == firstColumnInd + 10:
                        currentPeople.c1 = cell.value
                    if currentColumn == firstColumnInd + 11:
                        currentPeople.c2 = cell.value
                    if currentColumn == firstColumnInd + 12:
                        currentPeople.c3 = cell.value
                    if currentColumn == firstColumnInd + 17:
                        currentPeople.turns1 = cell.value
                    if currentColumn == firstColumnInd + 20:
                        currentPeople.turns2 = cell.value
                    if currentColumn == firstColumnInd + 21:
                        currentPeople.secBalls = cell.value
                    if currentColumn == firstColumnInd + 23:
                        currentPeople.total = cell.value
                    currentColumn += 1
                if currentPeople.total is not None:
                    self.peoples.append(currentPeople)

        for i in range(len(self.peoples)-1):
            for j in range(len(self.peoples)-i-1):
                cPeople = self.peoples[j]
                nPeople = self.peoples[j+1]
                if cPeople.getTotal() < nPeople.getTotal():
                    self.peoples[j], self.peoples[j+1] = self.peoples[j+1], self.peoples[j]

        for i in self.peoples:
            print(People.getPlace(i), " ", People.getName(i), " ", People.getYear(i), " ", People.getDischarge(i), " ",
                 People.getCity(i), " ", People.getSchool(i), " ", People.getC1(i), " ", People.getC2(i), " ", People.getC3(i), " ",
                 People.getTurns1(i), People.getTurns2(i), People.getSecBalls(i), " ", People.getTotal(i))
        print('______________________________')
