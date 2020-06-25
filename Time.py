
class Time:
    def __init__(self, start, end, weekday):
        self.start = start
        self.end = end
        self.weekday = weekday

    def info(self):
        print('start:', self.start)
        print('end:', self.end)
        print('weekday:', self.weekday)

    # check if conflicts with a time
    # return true if there is a conflict
    def conflict(self, time):
        for day in range(len(self.weekday)):
            if self.weekday[day] == 1 and time.weekday[day] == 1:
                if time.start <= self.start <= time.end:
                    return True
                if time.start <= self.end <= time.end:
                    return True
                if self.start <= time.start <= self.end:
                    return True
                if self.start <= time.end <= self.end:
                    return True
        return False
