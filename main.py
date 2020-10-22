import xlrd

path = 'Notas.xlsx'

inputWorkbook = xlrd.open_workbook(path)
inputWorksheet = inputWorkbook.sheet_by_index(0)

print(inputWorksheet.nrows)
print(inputWorksheet.ncols)

students = []

for row in range(3, inputWorksheet.nrows):
    rm = inputWorksheet.cell_value(row,0)
    name = inputWorksheet.cell_value(row,1)
    firstHalf = inputWorksheet.cell_value(row,2)
    checkpoint1 = inputWorksheet.cell_value(row, 4)
    checkpoint2 = inputWorksheet.cell_value(row, 5)
    checkpoint3 = inputWorksheet.cell_value(row, 6)
    challengeSprint3 = inputWorksheet.cell_value(row, 8)
    challengeSprint4 = inputWorksheet.cell_value(row, 9)
    grades = {
        'rm' : rm,
        'name': name,
        'firstHalf': firstHalf,
        'checkpoint': [
            checkpoint1, 
            checkpoint2, 
            checkpoint3
        ],
        'challenge': [
            challengeSprint3, 
            challengeSprint4
        ]
    }
    students.append(grades)

noteMinimumFinal = 6

for student in students :
    print('\n')
    print(int(student['rm']), '  ', student['name'])
    print('Diciplina: Computational Thinking using Python')
    print('Semestre 1:', student['firstHalf'])
    print('Semestre 2')
    mediaCheckpoints = (student['checkpoint'][0] + student['checkpoint'][1] + student['checkpoint'][2]) / 3
    print('Checkpoints: ', mediaCheckpoints)
    mediaChallenges = (student['challenge'][0] + student['challenge'][1]) / 2
    print('Challenge: ', mediaChallenges)
    noteMediaFirstHalf = student['firstHalf'] * 0.4
    mediaCheckpointsAndChallenges = ((mediaCheckpoints + mediaChallenges) / 2) * 0.4
    noteNeedSecondHalf = (noteMinimumFinal - noteMediaFirstHalf) / 0.6
    noteNeedGlobalSolution = (noteNeedSecondHalf - mediaCheckpointsAndChallenges) / 0.6
    print('Nota Minima da Global Solution para aprovação: ', noteNeedGlobalSolution)
