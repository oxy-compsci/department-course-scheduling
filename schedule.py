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

    def __init__(self, course, section, units, semester):
        self.course = course
        self.section = section
        self.units = units
        self.semester = semester

    def __str__(self):
        return self.name

    @property
    def name(self):
        return self.semester + ' ' + self.course + ' Section ' + str(self.section)

    def info(self):
        print('course:', self.course)
        print('section:', self.section)
        print('units:', self.units)
        print('semester:', self.semester)


class Professor:

    def __init__(self, name, max_units, capabilities, preference):
        self.name = name
        self.max_units = max_units
        self.capabilities = capabilities
        self.preference = preference

    def __str__(self):
        return self.name

    def info(self):
        print('name:', self.name)
        print('units:', self.max_units)
        print('can teach: ')
        for sec in self.capabilities:
            print(sec, end=' ')
            print(self.capabilities[sec])
        print('prefer: ')
        for sec in self.preference:
            print(sec, end=' ')
            print(self.preference[sec])

    def can_teach(self, course):
        if course in self.capabilities:
            return 1
        else:
            return 0

    def prefers(self, course):
        if course in self.preference:
            return 1
        else:
            return 0


# get data from excel
# read input, separate classes into sections
# check for infeasible situation, and store info into objects
def read_input(input_file):
    # get data from excel
    can_teach_tab = pd.read_excel(input_file, sheet_name='CanTeach', index_col=0)
    prefer_tab = pd.read_excel(input_file, sheet_name='Prefer', index_col=0)
    course_tab = pd.read_excel(input_file, sheet_name='Course', index_col=0)
    prof_tab = pd.read_excel(input_file, sheet_name='Prof', index_col=0)
    time_tab = pd.read_excel(input_file, sheet_name='Time', index_col=None)

    # check that the same professors are defined across all tabs
    professors_okay = set(can_teach_tab.index) == set(prefer_tab.index) == set(prof_tab.index)
    assert professors_okay, 'some professors are missing from some tabs!'

    # check that the same courses are defined across all tabs
    courses_okay = set(can_teach_tab.columns) == set(prefer_tab.columns) == set(course_tab.index)
    assert courses_okay, 'some courses are missing from some tabs!'

    course_names = set(can_teach_tab.columns)
    professor_names = set(can_teach_tab.index)
    semesters = course_tab.columns.tolist()
    semesters.remove('Unit')
    semesters.remove('UnitSum')

    professors = {}
    for name in professor_names:
        max_units = prof_tab['MaxUnit'][name]
        capabilities = set(
            course_name for course_name in course_names
            if can_teach_tab[course_name][name] == 1
        )
        preferences = set(
            course_name for course_name in course_names
            if prefer_tab[course_name][name] == 1
        )
        professors[name] = Professor(name, max_units, capabilities, preferences)

    sections = {}
    for course_name in course_names:
        units = course_tab['Unit'][course_name]
        for semester in semesters:
            num_sections = course_tab[semester][course_name]
            for section_num in range(num_sections):
                section = Section(course_name, section_num, units, semester)
                sections[section.name] = section

    # create Time objects
    # Time have start time, end time, list of 1/0 for weekdays, and list of conflicts with all time slots
    times = []
    for _, rows in time_tab.iterrows():
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

    # error checking

    # check that every course has at least one professor who can teach it
    for course_name in course_names:
        if not any(course_name in professor.capabilities for professor in professors.values()):
            raise ValueError('No professor can teach ' + course_name)

    # check that the total number of units from professors is enough
    # TODO this may not be true in the future, if we just through all available courses into the solver
    # put in all available courses and select ones with professor preference/set if the course must be offered
    total_professor_units = sum(professor.max_units for professor in professors.values())
    total_course_units = sum(section.units for section in sections.values())
    if total_course_units > total_professor_units:
        raise ValueError('Professors can only teach {} units but there are {} course units total'.format(
            total_professor_units,
            total_course_units,
        ))

    return semesters, sections, professors, times


# create model and add constraints
def create_model(professors, sections, semesters):
    # Creates the model.
    model = cp_model.CpModel()

    # Creates class variables.
    # classes[(p,c,s)]: professor 'p' teaches class 'c'
    classes = {}
    for professor in professors.values():
        for section in sections.values():
            classes[(professor.name, section.name)] = model.NewBoolVar('{} teaches {}'.format(professor.name, section.name))

    # hard constraints

    # All sections must be assigned
    # TODO this may not be true in the future, if we just through all available courses into the solver
    model.Add(
        sum(
            classes[(prof_name, section_name)]
            for prof_name in professors
            for section_name in sections
        ) == len(sections)
    )

    # Each class is assigned to exactly one professor.
    for section_name in sections:
        model.Add(sum(classes[(prof_name, section_name)] for prof_name in professors) == 1)

    # Only schedule classes that professors can teach
    model.Add(
        sum(
            classes[(prof_name, section_name)] * professor.can_teach(section.course)
            for prof_name, professor in professors.items()
            for section_name, section in sections.items()
        ) == len(sections)
    )

    # Professors cannot teach more than their max number of units
    # 12 units per semester max
    for prof_name, professor in professors.items():
        model.Add(
            sum(
                classes[(prof_name, section_name)] * section.units
                for section_name, section in sections.items()
            ) <= professor.max_units
        )
        for semester in semesters:
            model.Add(
                sum(
                    classes[(prof_name, section_name)] * section.units
                    for section_name, section in sections.items()
                    if section.semester == semester
                ) <= MAX_UNITS_PER_SEMESTER
            )

    # soft constraints

    # assign classes according to prof preference
    model.Maximize(sum(
        classes[(prof_name, section_name)] * professor.prefers(section.course)
        for prof_name, professor in professors.items()
        for section_name, section in sections.items()
    ))

    return model, classes


def solve_model(model):
    # Creates the solver and solve.
    solver = cp_model.CpSolver()
    solver.Solve(model)
    assert solver.StatusName() != 'INFEASIBLE', 'PROBLEM IS INFEASIBLE'
    return solver


def print_results(solver, classes, professors, sections, semesters):

    # print in course-first format
    for semester in semesters:
        print(semester)
        for _, section in sorted(sections.items()):
            if section.semester != semester:
                continue
            for _, professor in sorted(professors.items()):
                if solver.Value(classes[(professor.name, section.name)]) == 1:
                    if professor.prefers(section.course):
                        print(section.name + ' assigned to ' + professor.name + ' (requested)')
                    else:
                        print(section.name + ' assigned to ' + professor.name + ' (not requested)')
                    break
        print()

    # print in professor-first format
    for semester in semesters:
        print(semester)
        for _, professor in sorted(professors.items()):
            for _, section in sorted(sections.items()):
                if section.semester != semester:
                    continue
                if solver.Value(classes[(professor.name, section.name)]) == 1:
                    if professor.prefers(section.course):
                        print(professor.name + ' will be teaching ' + section.name + ' (requested)')
                    else:
                        print(professor.name + ' will be teaching ' + section.name + ' (not requested)')
        print()
        
    # Statistics.
    print('Statistics')
    print('  - Number of requests met = %i' % solver.ObjectiveValue())
    print('  - wall time       : %f s' % solver.WallTime())


# return infos for timeslots scheduling
# list of scheduled classes tuple (professor.name, section.name) , listed by semesters
def get_scheduled_classes(solver, classes, professors, sections, semesters):
    scheduled_classes = {}
    for semester in semesters:
        profs_classeses = []
        for _, section in sections.items():
            if section.semester != semester:
                continue
            for _, professor in professors.items():
                if solver.Value(classes[(professor.name, section.name)]) == 1:
                    profs_classeses.append((professor.name, section.name))
                    break
        scheduled_classes[semester] = profs_classeses
    return scheduled_classes


def create_timetable_model(profs_classes, times):
    prof_teach = {}  # professor to list of scheduled classes
    for prof_sec in profs_classes:
        prof = prof_sec[0]
        sec = prof_sec[1]
        if prof in prof_teach:
            prof_teach[prof].append(sec)
        else:
            prof_teach[prof] = [sec]

    # Creates the model.
    model = cp_model.CpModel()
    # Creates variables.
    # c - tuple (professor, section), t - time
    time_assign = {}
    for c in profs_classes:
        for t in times:
            time_assign[(c, t)] = model.NewBoolVar(
                '{} teaches {} in timeslot {}'.format(c[0], c[1], times.index(t))
            )

    # hard constraints
    # Each class is assigned to exactly one time slot.
    for c in profs_classes:
        model.Add(sum(time_assign[(c, t)] for t in times) == 1)

    # time slots for a professor don't conflict
    for prof_name, sec_name in prof_teach.items():
        for s1 in sec_name:
            for s2 in sec_name:
                for t1 in times:
                    for t2 in times:
                        if s1 != s2:
                            content = [time_assign[((prof_name, s1), t1)], time_assign[((prof_name, s2), t2)],
                                       t1.conflicts[t2]]
                            model.AddMultiplicationEquality(0, content)


    # Minimize the overall time conflicts
    conflicts = {}
    conflict_name_template = 'course {} (timeslot {}) and course {} (timeslot {}) conflict'
    for time1 in times:
        for time2 in times:
            if not time1.conflict(time2):
                continue
            for course1 in profs_classes:
                for course2 in profs_classes:
                    if course1 == course2:
                        continue
                    # (course1, time1) && (course2, time2) -> (course1, time1, course2, time2)
                    key = (course1, time1, course2, time2)
                    # create the conflict variable
                    conflicts[key] = model.NewBoolVar(conflict_name_template.format(*key))
                    # tell the model that (course1, time1) && (course2, time2) -> AND_VAR
                    model.Add(conflicts[key] == 1).OnlyEnforceIf([
                        time_assign[(course1, time1)],
                        time_assign[(course2, time2)],
                    ])
    model.Minimize(sum(conflicts.values()))

    # Assign according to prof preference - both hard and soft, preferences over general time periods
    # (MWF/TR, morning/evening)
    # create the conflict variable

    return model, time_assign


# print the final timetable for one semester
# professor, class name, start time, end time, weekday
def print_timetable(solver, time_assign, times, profs_classes):
    for c in profs_classes:
        for t in times:
            if solver.Value(time_assign[(c, t)]) == 1:
                print(c[0], c[1], t.start, t.end, t.weekday)
    print()


def main():

    # switch input data order
    input_file = 'Testing data.xlsx'

    semesters, sections, professors, times = read_input(input_file)

    model, classes = create_model(professors, sections, semesters)

    solver = solve_model(model)

    print_results(solver, classes, professors, sections, semesters)

    scheduled_classes = get_scheduled_classes(solver, classes, professors, sections, semesters)

    for semester in semesters:
        profs_classes = scheduled_classes[semester]  # list of (professor.name, section.name)
        model, time_assign = create_timetable_model(profs_classes, times)
        solver = solve_model(model)
        print_timetable(solver, time_assign, times, profs_classes)


if __name__ == '__main__':
    main()