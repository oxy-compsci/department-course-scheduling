from __future__ import print_function
from ortools.sat.python import cp_model
import pandas as pd
import xlrd


def main():
    input_file = 'Testing data.xlsx'

    # get data from excel
    CanTeach = pd.read_excel(input_file, sheet_name='CanTeach', index_col=0)
    Course = pd.read_excel(input_file, sheet_name='Course', index_col=0)
    Prof = pd.read_excel(input_file, sheet_name='Prof', index_col=0)
    Prefer = pd.read_excel(input_file, sheet_name='Prefer', index_col=0)
    class_name = CanTeach.columns.tolist()
    prof_name = CanTeach.index.tolist()

    prof_max_credits = Prof['MaxUnit'].tolist()
    class_sem = [Course['Sem 1'].tolist(), Course['Sem 2'].tolist()]
    class_credits = Course['Unit'].tolist()
    class_can_teach = []
    for index, rows in CanTeach.iterrows():
        class_can_teach.append(rows.tolist())
    class_prefer = []
    for index, rows in Prefer.iterrows():
        class_prefer.append(rows.tolist())
    num_prof = len(prof_max_credits)
    num_class = len(class_credits)
    num_sem = len(class_sem)
    all_profs = range(num_prof)
    all_classes = range(num_class)
    all_sems = range(num_sem)

    # check for infeasible situations
    if sum(prof_max_credits) < sum(class_credits):
        exit('Professor can not provide enough number of units')
    for x in all_classes:
        if class_sem[0][x] == 0 and class_sem[1][x] == 0:
            exit(class_name[x] + ' is not assigned to a semester')
    for col in class_name:
        if sum(CanTeach[col].tolist()) == 0:
            exit('No professor can teach ' + col)

    # Creates the model.
    model = cp_model.CpModel()

    # Creates class variables.
    # classes[(p,c,s)]: professor 'p' teaches class 'c' in semester 's'
    classes = {}
    for p in all_profs:
        for c in all_classes:
            for s in all_sems:
                classes[(p, c, s)] = model.NewBoolVar('classes_n%id%is%i' % (p, c, s))


    # hard constraints

    # Each class is assigned to exactly one professor.
    for c in all_classes:
        model.Add(sum(classes[(p, c, s)] for p in all_profs for s in all_sems) == 1)

    # Professors cannot teach more than their max number of units
    # 12 units per semester
    for p in all_profs:
        model.Add(sum(classes[(p, c, s)] * class_credits[c] for c in all_classes for s in all_sems)
                  <= prof_max_credits[p])
        for s in all_sems:
            model.Add(sum(classes[(p, c, s)] * class_credits[c] for c in all_classes) <= 12)

    # Only schedule classes that professors can teach
    # assign semesters to each class
    model.Add(sum(classes[(p, c, s)] * class_can_teach[p][c] for p in all_profs
                  for c in all_classes for s in all_sems) == num_class)
    model.Add(sum(classes[(p, c, s)] * class_sem[s][c] for p in all_profs
                  for c in all_classes for s in all_sems) == num_class)

    # soft constraints
    # assign classes according to prof preference
    model.Maximize(sum(classes[(p, c, s)] * class_prefer[p][c] for p in all_profs
                       for c in all_classes for s in all_sems))

    # Creates the solver and solve.
    solver = cp_model.CpSolver()
    solver.Solve(model)
    if solver.StatusName() == 'INFEASIBLE':
        exit('INFEASIBLE')
    for p in all_profs:
        print('Prof', p)
        for s in all_sems:
            for c in all_classes:
                if solver.Value(classes[(p, c, s)]) == 1 and class_prefer[p][c] == 1:
                    print('Semester', s, 'Class', c, '(requested)')
                elif solver.Value(classes[(p, c, s)]) == 1:
                    print('Semester', s, 'Class', c, '(not requested)')

    # Statistics.
    print()
    print('Statistics')
    print('  - Number of requests met = %i' % solver.ObjectiveValue())
    print('  - wall time       : %f s' % solver.WallTime())
    # print(solver.NumConflicts())


if __name__ == '__main__':
    main()
