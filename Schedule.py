from __future__ import print_function
from ortools.sat.python import cp_model


def main():
    # user input, placeholders for now
    num_prof = 6
    num_class = 20
    num_sem = 2
    all_profs = range(num_prof)
    all_classes = range(num_class)
    all_sems = range(num_sem)
    class_credits = [4,4,2,8,4,4,2,4,4,4,4,4,2,8,4,4,2,4,4,4]
    prof_max_credits = [20,20,20,20,20,10] # check total >= total class credits

    # fill in stuff
    class_sem = [[0, 1], [1, 0], [0, 1]]  # not both semester
    # [professor][class]
    class_can_teach = [[0,0,0,0,0,0,0,1,1,0,0,0,1,0,0,1,0,0,1,1]]
    # same as model [p][c][s]
    class_requests = [[[1,1],[0,0],[0,0]],[[0,0],[0,1],[1,0]]]

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
    # consider combine them into fewer loops

    # Each class is assigned to exactly one professor.
    for c in all_classes:
        model.Add(sum(classes[(p, c, s)] for p in all_profs for s in all_sems) == 1)

    # need to add units
    for p in all_profs:
        model.Add(sum(classes[(p, c, s)] for c in all_classes for s in all_sems
                      if classes[(p,c,s)] == 1) <= 5)


    '''
    # Professors cannot teach more than their max number of units
    # 12 units per semester
    for p in all_profs:
        model.Add(sum(class_credits[c] for c in all_classes for s in all_sems
                      if classes[(p, c, s)] == 1) <= prof_max_credits[p])
        for s in all_sems:
            model.Add(sum(class_credits[c] for c in all_classes
                          if classes[(p, c, s)] == 1) <= 12)
    '''


    '''
    # Only schedule classes that professors can teach
    # assign semesters to each class
    for p in all_profs:
        for c in all_classes:
            for s in all_sems:
                model.Add(class_can_teach[p][c] == classes[(p, c, s)])
                model.Add(class_sem[c][s] == classes[(p, c, s)])
    '''


    # Creates the solver and solve.
    solver = cp_model.CpSolver()
    solver.Solve(model)
    if solver.StatusName() == 'INFEASIBLE':
        exit('INFEASIBLE')
    for c in all_classes:
        print('Class', c)
        for p in all_profs:
            for s in all_sems:
                if solver.Value(classes[(p, c, s)]) == 1:
                    print('Professor ', p, 'Semester ', s)
        print()

    # Statistics.
    print()
    print('Statistics')
    print('  - Number of shift requests met = %i' % solver.ObjectiveValue())
    print('  - wall time       : %f s' % solver.WallTime())


if __name__ == '__main__':
    main()