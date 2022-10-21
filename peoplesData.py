

class People(object):
    def __init__(self, place, name, year, discharge, city, school, total, seks, c1, c2, c3):
        self.place = place
        self.name = name
        self.year = year
        self.discharge = discharge
        self.city = city
        self.school = school
        self.total = total
        self.seks = seks
        self.c1 = c1
        self.c2 = c2
        self.c3 = c3

    def showData(self):
        print(self.place, " ", self.name, " ", self.year, " ", self.discharge, " ", self.city, " ", self.school, " ",
              self.c1, " ", self.c2, " ", self.c3, " ", self.seks, " ", self.total)

    def getTotal(self):
        return self.total