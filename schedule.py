from ortools.sat.python import cp_model
import pandas as pd

MAX_UNITS_PER_SEMESTER = 12


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


class Professor:

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
            print(sec, end=' ')
            print(self.can_teach[sec])
        print('prefer: ')
        for sec in self.preference:
            print(sec, end=' ')
            print(self.preference[sec])


# get data from excel
# read input, separate classes into sections
# check for infeasible situation, and store info into objects
def read_input(input_file):
    # get data from excel
    can_teach = pd.read_excel(input_file, sheet_name='CanTeach', index_col=0)
    course = pd.read_excel(input_file, sheet_name='Course', index_col=0)
    prof = pd.read_excel(input_file, sheet_name='Prof', index_col=0)
    prefer = pd.read_excel(input_file, sheet_name='Prefer', index_col=0)
    class_name = can_teach.columns.tolist()
    prof_name = can_teach.index.tolist()
    sem_name = course.columns.tolist()
    sem_name.remove('Unit')
    sem_name.remove('UnitSum')

    # infeasible if no prof can teach one class

    for col in class_name:
        assert sum(can_teach[col].tolist()) > 0, 'No professor can teach ' + col

    prof_max_credits = prof['MaxUnit'].tolist()
    class_sec_num = get_rows(course, 2)
    # [[(class 0)section number in sem 0, section number in sem 1],
    # [(class 0)section number in sem 0, section number in sem 1], ...]
    class_credits = course['Unit'].tolist()
    class_can_teach = get_rows(can_teach)
    # [[(prof 0) class 0, class 1, class 2,...], [(prof 2) class 0 ,...],...]
    class_prefer = get_rows(prefer)

    return class_name, class_credits, class_sec_num, prof_name, \
           prof_max_credits, class_can_teach, class_prefer, sem_name


# separate classes into sections
def create_sections(
    class_name, class_credits, class_sec_num, prof_name, prof_max_credits, class_can_teach,
    class_prefer, sem_name
):
    num_prof = len(prof_name)
    num_class = len(class_name)
    num_sem = len(sem_name)
    all_profs = range(num_prof)
    all_classes = range(num_class)
    all_sems = range(num_sem)

    # expand name and semester lists
    sec_name = []
    sec_sem = []
    for c in all_classes:
        total_sec = sum(class_sec_num[c])
        sec_count = 0
        for sem in all_sems:
            for _ in range(class_sec_num[c][sem]):
                if sec_count < total_sec:
                    semester = [0] * num_sem
                    semester[sem] = 1
                    sec_sem.append(semester)
                    sec_name.append('%s Section %i' % (class_name[c], sec_count))
                sec_count = sec_count + 1

    # expand credits lists
    sec_credits = []
    for c in all_classes:
        total_sec = sum(class_sec_num[c])
        for _ in range(total_sec):
            sec_credits.append(class_credits[c])

    # expand can_teach and prefer
    sec_can_teach = []
    sec_prefer = []
    for p in all_profs:
        can_teach = []
        prefer = []
        for c in all_classes:
            total_sec = sum(class_sec_num[c])
            for _ in range(total_sec):
                can_teach.append(class_can_teach[p][c])
                prefer.append(class_prefer[p][c])
        sec_can_teach.append(can_teach)
        sec_prefer.append(prefer)

    # check for infeasible situations
    assert sum(sec_credits) <= sum(prof_max_credits), 'Professor can not provide enough number of units'

    for section in sec_name:
        assigned = False
        for sem in all_sems:
            if section[sem] == 1:
                assigned = True
        assert assigned, section + ' is not assigned to a semester'

    return sec_name, sec_credits, sec_sem, sec_can_teach, sec_prefer


# put the information into section and professor objects
# return two lists that contains all professors and sections
def create_objects(
    sec_name, sec_credits, sec_sem, sec_can_teach, sec_prefer, prof_name, prof_max_credits
):
    all_secs = range(len(sec_name))
    all_profs = range(len(prof_name))
    sections = []
    for x in all_secs:
        c = Section(sec_name[x], sec_credits[x], sec_sem[x])
        sections.append(c)

    professors = []
    for x in all_profs:
        can_teach = {}
        prefer = {}
        for y in all_secs:
            can_teach[sections[y]] = sec_can_teach[x][y]
        for y in all_secs:
            prefer[sections[y]] = sec_prefer[x][y]
        professors.append(Professor(prof_name[x], prof_max_credits[x], can_teach, prefer))

    return professors, sections


# create model and add constraints
def create_model(professors, sections, semesters):
    num_sec = len(sections)

    # Creates the model.
    model = cp_model.CpModel()

    # Creates class variables.
    # classes[(p,c,s)]: professor 'p' teaches class 'c' in semester 's'
    classes = {}
    for p in professors:
        for c in sections:
            for s in semesters:
                classes[(p, c, s)] = model.NewBoolVar('classes_n%s%s%s' % (p.name, c.name, s))

    # hard constraints

    # Each class is assigned to exactly one professor.
    for c in sections:
        model.Add(sum(classes[(p, c, s)] for p in professors for s in semesters) == 1)

    # Professors cannot teach more than their max number of units
    # 12 units per semester max
    for p in professors:
        model.Add(
            sum(classes[(p, c, s)] * c.units for c in sections for s in semesters)
            <= p.max_units
        )
        for s in semesters:
            model.Add(
                sum(classes[(p, c, s)] * c.units for c in sections)
                <= MAX_UNITS_PER_SEMESTER
            )

    # Only schedule classes that professors can teach
    # assign semesters to each class
    model.Add(
        sum(
            classes[(p, c, s)] * p.can_teach[c]
            for p in professors
            for c in sections
            for s in semesters
        ) == num_sec
    )
    model.Add(
        sum(
            classes[(p, c, s)] * c.semester[semesters.index(s)]
            for p in professors
            for c in sections
            for s in semesters
        ) == num_sec
    )

    # soft constraints
    # assign classes according to prof preference
    model.Maximize(sum(
        classes[(p, c, s)] * p.preference[c]
        for p in professors
        for c in sections
        for s in semesters
    ))

    return model, classes


# solve and print the schedule
# return infos for timeslots scheduling
def solve_model(model, classes, professors, sections, semesters):
    # Creates the solver and solve.
    solver = cp_model.CpSolver()
    solver.Solve(model)
    assert solver.StatusName() != 'INFEASIBLE', 'PROBLEM IS INFEASIBLE'

    scheduled_classes = []
    prof_class = []
    for s in semesters:
        sem_class = []
        prof_sem = {}
        print(s)
        for p in professors:
            prof_sem[p] = []
            print(p)
            for c in sections:
                if solver.Value(classes[(p, c, s)]) == 1:
                    sem_class.append((p, c))
                    prof_sem[p].append(c)
                    if p.preference[c] == 1:
                        print('Semester', s, c.name, '(requested)')
                    else:
                        print('Semester', s, c.name, '(not requested)')
        scheduled_classes.append(sem_class)
        prof_class.append(prof_sem)
        print()
    return scheduled_classes, prof_class


def main():
    # switch input data order
    input_file = 'Testing data.xlsx'
    class_name, class_credits, class_sec_num, prof_name, \
        prof_max_credits, class_can_teach, class_prefer, sem_name = read_input(input_file)

    sec_name, sec_credits, sec_sem, sec_can_teach, sec_prefer = create_sections \
        (class_name, class_credits, class_sec_num, prof_name, prof_max_credits, class_can_teach, class_prefer, sem_name)

    professors, sections = create_objects(
        sec_name, sec_credits, sec_sem, sec_can_teach, sec_prefer, prof_name, prof_max_credits
    )

    model, classes = create_model(professors, sections, sem_name)
    scheduled_classes, prof_class = solve_model(model, classes, professors, sections, sem_name)

    # schedule time for each class
    # get input from excel and create Time objects
    # Time have start time, end time, list of 1/0 for weekdays, and list of conflicts with all time slots
    timeslots = pd.read_excel(input_file, sheet_name='Time')
    times = []
    for _, rows in timeslots.iterrows():
        info = rows.tolist()
        t = Time(info[5], info[6], info[0:5])
        times.append(t)
    for t in times:
        conflicts = {}
        for c in times:
            if t.conflict(c):
                conflicts[c] = 1
            else:
                conflicts[c] = 0
        t.conflicts = conflicts

    # run on semester 0 only for now
    sem_class = scheduled_classes[0]
    sem_prof_class = prof_class[0]

    # Creates the model.
    model = cp_model.CpModel()
    # Creates class variables.
    # c - tuple (professor, section), t - time
    time_assign = {}
    for c in sem_class:
        for t in times:
            time_assign[(c, t)] = model.NewBoolVar(
                'times_n%s%s%d' % (c[0].name, c[1].name, times.index(t))
            )

    # hard constraints

    # Each class is assigned to exactly one time slot.
    for c in sem_class:
        model.Add(sum(time_assign[(c, t)] for t in times) == 1)

    # time slots for a professor don't conflict
    '''for p in sem_prof_class:
        model.Add(sum(time_assign[((p, c1), t1)] * time_assign[((p, c2), t2)] * t1.conflicts[t2]
                        for c1 in sem_prof_class[p] for c2 in sem_prof_class[p] for t1 in times for t2 in times) == 1)'''

    # Creates the solver and solve.
    solver = cp_model.CpSolver()
    solver.Solve(model)
    assert solver.StatusName() != 'INFEASIBLE', 'PROBLEM IS INFEASIBLE'
    # print results
    for c in sem_class:
        for t in times:
            if solver.Value(time_assign[(c, t)]) == 1:
                print(c[0], c[1], t.start, t.end, t.weekday)


if __name__ == '__main__':
    main()
