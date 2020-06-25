
class Section:

    def __init__(self, name, units, semester):
        self.name = name
        self.units = units
        self.semester = semester

    def __str__(self):
        return self.name

    def info(self):
        print('name:', self.name)
        print('units:', self.units)
        print('semester:', self.semester)