

class People(object):
    def __init__(self, place, name, year, discharge, city, school, seks, c1, c2, c3, total):
        self.place = place
        self.name = name
        self.year = year
        self.discharge = discharge
        self.city = city
        self.school = school
        self.seks = seks
        self.c1 = c1
        self.c2 = c2
        self.c3 = c3
        self.total = total

    def showData(self):
        print(self.place, " ", self.name, " ", self.year, " ", self.discharge, " ", self.city, " ", self.school, " ",
              self.c1, " ", self.c2, " ", self.c3, " ", self.seks, " ", self.total)

    def getPlace(self):
        return self.place
    def getName(self):
        return self.name
    def getYear(self):
        return self.year
    def getDischarge(self):
        return self.discharge
    def getCity(self):
        return self.city
    def getSchool(self):
        return self.school
    def getC1(self):
        return self.c1
    def getC2(self):
        return self.c2
    def getC3(self):
        return self.c3
    def getSeks(self):
        return self.seks
    def getTotal(self):
        return self.total