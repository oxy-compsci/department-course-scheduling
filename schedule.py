from ortools.sat.python import cp_model
import pandas as pd
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import copy

SHEETS_URL = "https://docs.google.com/spreadsheets/d/112IxSjwhCQmKnJdwn_UebT_lEW5CR2Q3GMzeaFJuNBg/edit?usp=sharing"
EXCEL_NAME = 'Testing data.xlsx'
MAX_UNITS_PER_SEMESTER = 12
TIMEFRAME = ['Morning', 'Afternoon', 'Evening']


class Time:
    Days_of_week = {'MWF': [1, 0, 1, 0, 1],
                    'TR': [0, 1, 0, 1, 0],
                    'MW': [1, 0, 1, 0, 0],
                    'MTWRF': [1, 1, 1, 1, 1],
                    'MF': [1, 0, 0, 0, 1],
                    'WF': [0, 0, 1, 0, 1],
                    'T': [0, 1, 0, 0, 0],
                    'W': [0, 0, 1, 0, 0],
                    'R': [0, 0, 0, 1, 0]}

    def __init__(self, start, end, weekdays, lab):
        self.start = datetime.datetime.strptime(start, '%H:%M:%S').time()
        self.end = datetime.datetime.strptime(end, '%H:%M:%S').time()
        self.weekdays = weekdays
        self.lab = lab
        self.conflicts = ()
        WEEKDAYS = 'MTWRF'
        self.days_of_week = ''.join(day_str for day_str, day_bool in zip(WEEKDAYS, weekdays) if day_bool)
        self.timeframe = 'Evening'
        if self.start <= datetime.time(17, 0, 0):
            self.timeframe = 'Afternoon'
        if self.start <= datetime.time(12, 0, 0):
            self.timeframe = 'Morning'

    def info(self):
        print('start:', self.start)
        print('end:', self.end)
        print('weekdays:', self.weekdays)

    # check if conflicts with a time
    # return true if there is a conflict
    def conflict(self, time):
        for day in range(len(self.weekdays)):
            if self.weekdays[day] == 1 and time.weekdays[day] == 1:
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

    def __init__(self, course, section, units, semester, must_offer, lab):
        self.course = course
        self.section = section
        self.units = units
        self.semester = semester
        self.must_offer = must_offer
        self.lab = lab

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

    def __init__(self, name, max_units, capabilities, preference, prefer_timeframe):
        self.name = name
        self.max_units = max_units
        self.capabilities = capabilities
        self.preference = preference
        self.prefer_timeframe = prefer_timeframe

    def __str__(self):
        return self.name

    def info(self):
        print('name:', self.name)
        print('units:', self.max_units)
        print('can teach: ')
        for sec in self.capabilities:
            print(sec, end=' ')
        print('prefer: ')
        for sec in self.preference:
            print(sec, end=' ')

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

    def prefer_time(self, time):
        return time.timeframe in self.prefer_timeframe.get(time.days_of_week, [])


# get data from google spreadsheet
# given the sheets name and the certificate file in directory
def read_ggsheets(sheets_url):
    # use creds to create a client to interact with the Google Drive API
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scopes)
    client = gspread.authorize(creds)
    # Open the sheets
    sheets = client.open_by_url(sheets_url)

    can_teach_tab = pd.DataFrame(sheets.worksheet('CanTeach').get_all_records()).set_index('')
    prefer_tab = pd.DataFrame(sheets.worksheet('Prefer').get_all_records()).set_index('')
    course_tab = pd.DataFrame(sheets.worksheet('Course').get_all_records()).set_index('')
    prof_tab = pd.DataFrame(sheets.worksheet('Prof').get_all_records()).set_index('')
    time_tab = pd.DataFrame(sheets.worksheet('Time').get_all_records())
    sheets = [can_teach_tab, prefer_tab, course_tab, prof_tab, time_tab]
    return sheets


# get data from excel
# given the excel name
def read_excel(input_file):
    can_teach_tab = pd.read_excel(input_file, sheet_name='CanTeach', index_col=0)
    prefer_tab = pd.read_excel(input_file, sheet_name='Prefer', index_col=0)
    course_tab = pd.read_excel(input_file, sheet_name='Course', index_col=0)
    prof_tab = pd.read_excel(input_file, sheet_name='Prof', index_col=0)
    time_tab = pd.read_excel(input_file, sheet_name='Time', index_col=None)
    sheets = [can_teach_tab, prefer_tab, course_tab, prof_tab, time_tab]
    return sheets


# read input, separate classes into sections
# check for infeasible situation, and store info into objects
def read_input(sheets):
    # get data from sheets
    can_teach_tab, prefer_tab, course_tab, prof_tab, time_tab = sheets

    # check that the same professors are defined across all tabs
    professors_okay = set(can_teach_tab.index) == set(prefer_tab.index) == set(prof_tab.index)
    assert professors_okay, 'some professors are missing from some tabs!'

    # check that the same courses are defined across all tabs
    courses_okay = set(can_teach_tab.columns) == set(prefer_tab.columns) == set(course_tab.index)
    assert courses_okay, 'some courses are missing from some tabs!'

    # get course names, prof names, and semester names
    course_names = set(can_teach_tab.columns)
    professor_names = set(can_teach_tab.index)
    semesters = course_tab.columns.tolist()
    semesters.remove('Unit')
    semesters.remove('Lab')
    for s in semesters:
        if "MustOffer" in s:
            semesters.remove(s)
        # check that all semesters have a MustOffer tab
        elif s + "_MustOffer" not in course_tab.columns.tolist():
            raise ValueError(s, 'semester does not have its corresponding MustOffer tab')

    # create Professor objects for each prof
    professors = {}  # {prof name : Professor}
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
        days_of_week = {}
        for days in Time.Days_of_week:
            if prof_tab[days][name] is not None:
                days_of_week[days] = str(prof_tab[days][name]).split(",")

        professors[name] = Professor(name, max_units, capabilities, preferences, days_of_week)

    # create Section objects for each section
    sections = {}  # {section name : Section}
    for course_name in course_names:
        units = course_tab['Unit'][course_name]
        lab = course_tab['Lab'][course_name]
        for semester in semesters:
            num_sections = course_tab[semester][course_name]
            must_offer = course_tab[semester + '_MustOffer'][course_name]
            section_num = 0
            for must in range(must_offer):
                section = Section(course_name, section_num, units, semester, must_offer=1, lab=lab)
                sections[section.name] = section
                section_num = section_num + 1
            for optional in range(num_sections - must_offer):
                section = Section(course_name, section_num, units, semester, must_offer=0, lab=lab)
                sections[section.name] = section
                section_num = section_num + 1

    # create Time objects for each time slots
    # Time have start time, end time, list of 1/0 for weekdays, and a set of conflicted time slots
    times = []
    for _, rows in time_tab.iterrows():
        info = rows.tolist()
        times.append(Time(start=info[5], end=info[6], weekdays=info[0:5], lab=info[7]))
    for t1 in times:
        conflicts = set(t2 for t2 in times if t1.conflict(t2))
        t1.conflicts = conflicts

    # error checking

    # for professors, it's better to not have all zeros for time slots preference

    # check that every course has at least one professor who can teach it
    for course_name in course_names:
        if not any(course_name in professor.capabilities for professor in professors.values()):
            raise ValueError('No professor can teach ' + course_name)

    # check if the required units are more than professors' total units
    total_professor_units = sum(professor.max_units for professor in professors.values())
    course_units_required = sum(section.units for section in sections.values() if section.must_offer)
    course_units_optional = sum(section.units for section in sections.values() if not section.must_offer)
    total_course_units = course_units_required + course_units_optional
    if course_units_required > total_professor_units:
        raise ValueError(course_units_required, 'are required, but professors can only teach',
                         total_professor_units, 'units.')
    # print out units information
    print('Professors can teach', total_professor_units, 'units.')
    print('There are', total_course_units, 'units form all classes.')
    print(course_units_required, 'are required,', course_units_optional, 'are optional.')

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
            classes[(professor.name, section.name)] = model.NewBoolVar(
                '{} teaches {}'.format(professor.name, section.name))

    # hard constraints

    # Each class is assigned to one professor or no professor.
    # allow the optional classes to be not assigned
    # schedule the courses that must be offered
    for section_name, section in sections.items():
        if section.must_offer:
            model.Add(sum(classes[(prof_name, section_name)] for prof_name in professors) == 1)
        else:
            model.Add(sum(classes[(prof_name, section_name)] for prof_name in professors) <= 1)

    # Only schedule classes that professors can teach
    for prof_name, professor in professors.items():
        for section_name, section in sections.items():
            model.Add(professor.can_teach(section.course) == 1).OnlyEnforceIf(classes[(prof_name, section_name)])

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
    print()


# return infos for timeslots scheduling
# list of scheduled classes tuple (professor.name, section.name) , listed by semesters
def get_semester_schedule(solver, classes, professors, sections, semesters):
    scheduled_classes = {}
    for semester in semesters:
        profs_classes = []
        for section in sections.values():
            if section.semester != semester:
                continue
            for professor in professors.values():
                if solver.Value(classes[(professor.name, section.name)]) == 1:
                    profs_classes.append((professor.name, section.name))
                    break
        scheduled_classes[semester] = profs_classes
    return scheduled_classes


# create a model for timetable scheduling
def create_timetable_model(profs_classes, professors, sections, times):
    prof_teach = {}  # professor : a list of scheduled classes
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

    # only lab sections are assign to lab time slots
    is_lab = {}
    for prof_name, sec_name in prof_teach.items():
        for s in sec_name:
            for t in times:
                # create intermediate boolean variables
                lab_key = (s, t)
                is_lab[lab_key] = model.NewBoolVar(
                    '{} section lab {}, timeslots {} lab {}'.format(s, sections[s].lab, t, t.lab))
                # variables are true when both the section and time are labs
                # or when both are not labs
                if sections[s].lab == t.lab:
                    model.Add(is_lab[lab_key] == 1)
                else:
                    model.Add(is_lab[lab_key] == 0)
                model.Add(is_lab[lab_key] == 1).OnlyEnforceIf(time_assign[(prof_name, s), t])

    # time slots for a professor don't conflict
    for prof_name, sec_name in prof_teach.items():
        for s1 in sec_name:
            for s2 in sec_name:
                if s1 != s2:
                    for t1 in times:
                        for t2 in times:
                            if t1 in t2.conflicts:
                                content = [time_assign[((prof_name, s1), t1)], time_assign[((prof_name, s2), t2)]]
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
                    model.Add(conflicts[key] == 1).OnlyEnforceIf([
                        time_assign[(course1, time1)],
                        time_assign[(course2, time2)],
                    ])

    # Maximize the number of time slots that profs prefer
    # should create a variable only if a professor prefers a time slot
    # but the solver will schedule more classes that are not preferred
    # so created a variable for all combination instead.
    prefer_time = {}
    for prof_name, sec_name in prof_teach.items():
        for s in sec_name:
            for t in times:
                key = (prof_name, s, t)
                # create the preferences variables
                prefer_time[key] = model.NewBoolVar('{} teaches course {} on timeslots {}'.format(prof_name, s, t))
                model.Add(prefer_time[key] ==
                          (professors[prof_name].prefer_time(t))).OnlyEnforceIf(
                    [time_assign[((prof_name, s), t)]]
                )

    # equalize their importance, then multiple by their weights
    model.Minimize(3 * len(prefer_time) * sum(conflicts.values())
                   - 2 * len(conflicts) * sum(prefer_time.values()))

    return model, time_assign


# print the final timetable for one semester
# professor, class name, start time, end time, weekdays
def print_semester_timetable(solver, time_assign, profs_classes, times, professors):
    for c in profs_classes:
        for t in times:
            if solver.Value(time_assign[(c, t)]) == 1:
                if professors[c[0]].prefer_time(t):
                    print(c[0], c[1], t.start, t.end, t.weekdays, 'Timeframe preferred')
                else:
                    print(c[0], c[1], t.start, t.end, t.weekdays, 'Timeframe not preferred')
    print()


# find more schedules by setting one variable
# return ones with the optimal objective_value
def find_all_schedule(model, variables):
    # find the optimal objective value by solving unrestricted variables first
    optimal_copy = copy.deepcopy(model)
    solver = solve_model(optimal_copy)
    optimal_value = solver.ObjectiveValue()

    solutions = []
    for var in variables:
        copies = copy.deepcopy(model)
        copies.Add(variables[var] == 1)
        # skip if it's infeasible
        try:
            solver = solve_model(copies)
        except AssertionError as e:
            continue
        if solver.ObjectiveValue() == optimal_value:
            solutions.append(solver)
    return solutions


def main(sheets_source):
    # schedule sections and print the result
    if sheets_source.startswith('http'):
        sheets = read_ggsheets(sheets_source)
    else:
        sheets = read_excel(sheets_source)
    semesters, sections, professors, times = read_input(sheets)
    model, classes = create_model(professors, sections, semesters)
    solver = solve_model(model)
    print_results(solver, classes, professors, sections, semesters)

    # timetable scheduling for each semester
    scheduled_classes = get_semester_schedule(solver, classes, professors, sections, semesters)

    for semester in semesters:
        profs_classes = scheduled_classes[semester]  # list of (professor.name, section.name)
        model, time_assign = create_timetable_model(profs_classes, professors, sections, times)
        solver = solve_model(model)
        print_semester_timetable(solver, time_assign, profs_classes, times, professors)


if __name__ == '__main__':
    main(SHEETS_URL)