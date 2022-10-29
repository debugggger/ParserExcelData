

class People(object):
    def __init__(self, place, name, year, discharge, city, school, secBalls, c1, c2, c3, turns1, turns2, total):
        self.place = place
        self.name = name
        self.year = year
        self.discharge = discharge
        self.city = city
        self.school = school
        self.secBalls = secBalls
        self.c1 = c1
        self.c2 = c2
        self.c3 = c3
        self.turns1 = turns1
        self.turns2 = turns2
        # self.turn11 = turn11
        # self.turn12 = turn12
        # self.turn21 = turn21
        # self.turn22 = turn22
        self.total = total

    def showData(self):
        print(self.place, " ", self.name, " ", self.year, " ", self.discharge, " ", self.city, " ", self.school, " ",
              self.c1, " ", self.c2, " ", self.c3, " ", self.secBalls, " ", self.total)

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
    def getSecBalls(self):
        return self.secBalls

    def getTurns1(self):
        return self.turns1
    def getTurns2(self):
        return self.turns2
    # def getTurn11(self):
    #     return self.turn11
    # def getTurn12(self):
    #     return self.turn12
    # def getTurn21(self):
    #     return self.turn22
    # def getTurn22(self):
    #     return self.turn22
    def getTotal(self):
        return self.total