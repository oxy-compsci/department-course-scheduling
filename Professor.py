

class Professor:
    'A professor with infos'

    def __init__(self, name, max_units, can_teach, preference):
        self.name = name
        self.max_units = max_units
        self.can_teach = can_teach
        self.preference = preference

    def __str__(self):
        return self.name

    def info(self):
        print('name:', self.name)
        print('units:', self.max_units)
        print('can teach: ')
        for sec in self.can_teach:
            print(sec, end = ' ')
            print(self.can_teach[sec])
        print('prefer: ')
        for sec in self.preference:
            print(sec, end = ' ')
            print(self.preference[sec])
