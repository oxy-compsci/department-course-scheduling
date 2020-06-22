from __future__ import print_function
from ortools.sat.python import cp_model
import pandas as pd
import xlrd


# return a list containing rows from DataFrame df, start from start_index in each row
def get_rows(df, start_index=0):
    data = []
    for index, rows in df.iterrows():
        data.append(rows.tolist()[start_index:])
    return data



def main():
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

    prof_max_credits = Prof['MaxUnit'].tolist()
    class_sec_num = get_rows(Course, 2)
    # [[(class 0)section number in sem 0, section number in sem 1],
    # [(class 0)section number in sem 0, section number in sem 1], ...]
    class_credits = Course['Unit'].tolist()
    class_can_teach = get_rows(CanTeach)
    # [[(prof 0) class 0, class 1, class 2,...], [(prof 2) class 0 ,...],...]
    class_prefer = get_rows(Prefer)

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

    num_sec = len(sec_name)
    all_secs = range(num_sec)

    # check for infeasible situations
    if sum(prof_max_credits) < sum(sec_credits):
        exit('Professor can not provide enough number of units')

    for x in all_secs:
        assigned = False
        for sem in all_sems:
            if sec_sem[x][sem] == 1:
                assigned = True
        if not assigned:
            exit(sec_name[x] + ' is not assigned to a semester')

    for col in class_name:
        if sum(CanTeach[col].tolist()) == 0:
            exit('No professor can teach ' + col)

    # Creates the model.
    model = cp_model.CpModel()

    # Creates class variables.
    # classes[(p,c,s)]: professor 'p' teaches class 'c' in semester 's'
    classes = {}
    for p in all_profs:
        for c in all_secs:
            for s in all_sems:
                classes[(p, c, s)] = model.NewBoolVar('classes_n%id%is%i' % (p, c, s))


    # hard constraints

    # Each class is assigned to exactly one professor.
    for c in all_secs:
        model.Add(sum(classes[(p, c, s)] for p in all_profs for s in all_sems) == 1)

    # Professors cannot teach more than their max number of units
    # 12 units per semester max
    # 4 units per semester min
    for p in all_profs:
        model.Add(sum(classes[(p, c, s)] * sec_credits[c] for c in all_secs for s in all_sems)
                  <= prof_max_credits[p])
        for s in all_sems:
            model.Add(sum(classes[(p, c, s)] * sec_credits[c] for c in all_secs) <= 12)
            '''model.Add(sum(classes[(p, c, s)] * sec_credits[c] for c in all_secs) >= 4)'''

    # Only schedule classes that professors can teach
    # assign semesters to each class
    model.Add(sum(classes[(p, c, s)] * sec_can_teach[p][c] for p in all_profs
                  for c in all_secs for s in all_sems) == num_sec)
    model.Add(sum(classes[(p, c, s)] * sec_sem[c][s] for p in all_profs
                  for c in all_secs for s in all_sems) == num_sec)

    # soft constraints
    # assign classes according to prof preference
    model.Maximize(sum(classes[(p, c, s)] * sec_prefer[p][c] for p in all_profs
                       for c in all_secs for s in all_sems))

    # Creates the solver and solve.
    solver = cp_model.CpSolver()
    solver.Solve(model)
    if solver.StatusName() == 'INFEASIBLE':
        exit('INFEASIBLE')
    for p in all_profs:
        print('Prof', p)
        for s in all_sems:
            for c in all_secs:
                if solver.Value(classes[(p, c, s)]) == 1 and sec_prefer[p][c] == 1:
                    print('Semester', s, sec_name[c], '(requested)')
                elif solver.Value(classes[(p, c, s)]) == 1:
                    print('Semester', s, sec_name[c], '(not requested)')

    # Statistics.
    print()
    print('Statistics')
    print('This solution is ' + solver.StatusName())
    print('  - Number of requests met = %i' % solver.ObjectiveValue())
    print('  - wall time       : %f s' % solver.WallTime())
    # print(solver.NumConflicts())


if __name__ == '__main__':
    main()
