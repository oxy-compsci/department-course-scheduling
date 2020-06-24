from ortools.sat.python import cp_model
import pandas as pd
import xlrd
from Professor import Professor
from Section import Section

Max_Unit_Per_Sem = 12


# return a list containing rows from DataFrame df, start from start_index in each row
def get_rows(df, start_index=0):
    data = []
    for index, rows in df.iterrows():
        data.append(rows.tolist()[start_index:])
    return data


def print_info(l):
    for s in l:
        s.info()
        print('-----------------------------------------------------------')



# read input, separate classes into sections
# check for infeasible situation, and store info into objects
def read_input(input):
    input_file = 'Testing data.xlsx'
    # get data from excel
    CanTeach = pd.read_excel(input_file, sheet_name='CanTeach', index_col=0)
    Course = pd.read_excel(input_file, sheet_name='Course', index_col=0)
    Prof = pd.read_excel(input_file, sheet_name='Prof', index_col=0)
    Prefer = pd.read_excel(input_file, sheet_name='Prefer', index_col=0)
    class_name = CanTeach.columns.tolist()
    prof_name = CanTeach.index.tolist()
    sem_name = Course.columns.tolist()
    sem_name.remove('Unit')
    sem_name.remove('UnitSum')

    # infeasible if no prof can teach one class
    for col in class_name:
        if sum(CanTeach[col].tolist()) == 0:
            exit('No professor can teach ' + col)

    prof_max_credits = Prof['MaxUnit'].tolist()
    class_sec_num = get_rows(Course, 2)
    # [[(class 0)section number in sem 0, section number in sem 1],
    # [(class 0)section number in sem 0, section number in sem 1], ...]
    class_credits = Course['Unit'].tolist()
    class_can_teach = get_rows(CanTeach)
    # [[(prof 0) class 0, class 1, class 2,...], [(prof 2) class 0 ,...],...]
    class_prefer = get_rows(Prefer)

    return class_name, class_credits, class_sec_num, prof_name, \
           prof_max_credits, class_can_teach, class_prefer, sem_name


def creat_sections(class_name, class_credits, class_sec_num, prof_name,
                   prof_max_credits, class_can_teach, class_prefer, sem_name):
    num_prof = len(prof_name)
    num_class = len(class_name)
    num_sem = len(sem_name)
    all_profs = range(num_prof)
    all_classes = range(num_class)
    all_sems = range(num_sem)

    # separate classes into sections
    sec_name = []
    sec_sem = []

    for c in all_classes:
        total_sec = sum(class_sec_num[c])
        sec_count = 0
        for sem in all_sems:
            for sec in range(class_sec_num[c][sem]):
                if sec_count < total_sec:
                    semester = [0] * num_sem
                    semester[sem] = 1
                    sec_sem.append(semester)
                    sec_name.append('%s Section %i' % (class_name[c], sec_count))
                sec_count = sec_count + 1

    sec_credits = []
    for c in all_classes:
        total_sec = sum(class_sec_num[c])
        for sec in range(total_sec):
            sec_credits.append(class_credits[c])

    sec_can_teach = []
    sec_prefer = []
    for p in all_profs:
        can_teach = []
        prefer = []
        for c in all_classes:
            total_sec = sum(class_sec_num[c])
            for sec in range(total_sec):
                can_teach.append(class_can_teach[p][c])
                prefer.append(class_prefer[p][c])
        sec_can_teach.append(can_teach)
        sec_prefer.append(prefer)

    # check for infeasible situations
    if sum(prof_max_credits) < sum(sec_credits):
        exit('Professor can not provide enough number of units')

    for x in range(len(sec_name)):
        assigned = False
        for sem in all_sems:
            if sec_sem[x][sem] == 1:
                assigned = True
        if not assigned:
            exit(sec_name[x] + ' is not assigned to a semester')

    return sec_name, sec_credits, sec_sem, sec_can_teach, sec_prefer



def create_objects(sec_name, sec_credits, sec_sem, sec_can_teach, sec_prefer,
                   prof_name, prof_max_credits):
    all_secs = range(len(sec_name))
    all_profs = range(len(prof_name))
    # put the information into section and professor objects
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
        p = Professor(prof_name[x], prof_max_credits[x], can_teach, prefer)
        professors.append(p)

    return professors, sections


def creat_model(professors, sections, semesters):
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
        model.Add(sum(classes[(p, c, s)] * c.units for c in sections for s in semesters)
                  <= p.max_units)
        for s in semesters:
            model.Add(sum(classes[(p, c, s)] * c.units for c in sections) <= Max_Unit_Per_Sem)

    # Only schedule classes that professors can teach
    # assign semesters to each class
    model.Add(sum(classes[(p, c, s)] * p.can_teach[c] for p in professors
                  for c in sections for s in semesters) == num_sec)
    model.Add(sum(classes[(p, c, s)] * c.semester[semesters.index(s)] for p in professors
                  for c in sections for s in semesters) == num_sec)

    # soft constraints
    # assign classes according to prof preference
    model.Maximize(sum(classes[(p, c, s)] * p.preference[c] for p in professors
                       for c in sections for s in semesters))

    return model, classes


def solve_model(model, classes, professors, sections, semesters):
    # Creates the solver and solve.
    solver = cp_model.CpSolver()
    solver.Solve(model)
    if solver.StatusName() == 'INFEASIBLE':
        exit('INFEASIBLE')

    for p in professors:
        print(p)
        for c in sections:
            for s in semesters:
                if solver.Value(classes[(p, c, s)]) == 1 and p.preference[c] == 1:
                    print('Semester', s, c.name, '(requested)')
                elif solver.Value(classes[(p, c, s)]) == 1:
                    print('Semester', s, c.name, '(not requested)')


def main():
    # switch input data order
    input_file = 'Testing data.xlsx'
    class_name, class_credits, class_sec_num, prof_name, \
    prof_max_credits, class_can_teach, class_prefer, sem_name = read_input(input_file)

    sec_name, sec_credits, sec_sem, sec_can_teach, sec_prefer = creat_sections \
        (class_name, class_credits, class_sec_num, prof_name, prof_max_credits, class_can_teach, class_prefer, sem_name)

    professors, sections = create_objects(sec_name, sec_credits, sec_sem, sec_can_teach, sec_prefer,
                   prof_name, prof_max_credits)

    model, classes = creat_model(professors, sections, sem_name)
    solve_model(model, classes, professors, sections, sem_name)


if __name__ == '__main__':
    main()
