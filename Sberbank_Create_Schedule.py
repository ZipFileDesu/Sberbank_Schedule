import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

#Считывание всех сотрудников
def read_all_workers(cnt_emp, cnt_drv):
    employees, drivers = [], []
    for i in range(cnt_emp):
        employees.append(data[1][i + 1])
    for i in range(cnt_drv):
        drivers.append(data[1][(cnt_emp + 1) + i])
    return employees, drivers

#Считывание дней
def read_all_days():
    days = []
    for i in range(30):
        days.append(data[i + 8][0])
    return days

#Генерация начального расписания
def generate_start_schedule():
    schedule = {}
    for i in range(len(employees)):
        schedule[employees[i]] = [days, ['' for _ in range(len(days))]]
    for i in range(len(drivers)):
        schedule[drivers[i]] = [days, ['' for _ in range(len(days))]]
    return schedule

#Ставим сначало отпуск в расписание
def set_vacation_start(vacations):
    for i in range(len(vacations)):
        worker = vacations[i][0]
        start_day = vacations[i][1]
        end_day = vacations[i][2]
        for j in range(end_day - start_day + 1):
            schedule[worker][1][start_day - 1 + j] = 'О'

#Ставим сначало выходные в расписания
def set_weekend_start(weekend):
    for i in range(len(weekend)):
        worker = weekend[i][0]
        if (len(weekend[i][1]) == 2):
            start_day = weekend[i][1][0]
            end_day = weekend[i][1][1]
            for j in range(end_day - start_day + 1):
                schedule[worker][1][start_day - 1 + j] = 'В'
        else:
            for j in range(len(weekend[i][1])):
                schedule[worker][1][weekend[i][1][j] - 1] = 'В'

data = pd.read_excel('Schedule.xls', sheet_name='График',
                     skiprows=5, index_col=None, header=None)

work_shifts = [












]

print("Enter the number of employees: ")
count_employees = int(input())
print("Enter the number of drivers: ")
count_drivers = int(input())

employees, drivers = read_all_workers(count_employees, count_drivers)
days = read_all_days()

schedule = generate_start_schedule()

vacations = [['Сотрудник 11', 1, 23],
             ['Сотрудник 15', 1, 5],
             ['Сотрудник 17', 20, 30],
             ['Сотрудник 18', 1, 30],
             ['Сотрудник 22', 16, 30],
             ['Водитель 2', 14, 30],
             ['Водитель 12', 18, 30],
             ['Водитель 17', 6, 30],
             ['Водитель 19', 16, 30],
             ['Водитель 25', 1, 23],
             ['Водитель 31', 1, 17]]

weekend = [['Сотрудник 2', [5, 8]],
           ['Сотрудник 5', [2, 4]],
           ['Сотрудник 6', [9, 13]],
           ['Водитель 3', [1, 3]],
           ['Водитель 5', [19, 22]],
           ['Водитель 17', [5]],
           ['Водитель 19', [13, 15]],
           ['Водитель 30', [7, 8, 14, 15, 21, 22, 28, 29]]

           ]
set_vacation_start(vacations)
set_weekend_start(weekend)