import math
import openpyxl as op
from openpyxl.workbook import Workbook
from openpyxl.styles import Font

# open new workbook
# wb = op.Workbook()

# load workbook
# wb= op.load_workbook('worker tracing.xls')




categories = ["Date", "Hours", "Base salary", "Tip", "Overall", "Completion", "Additional hours",  "Shabbat hours", "comments" ]

worker_list = ["Michal D", "Michal H", "Yoad", "Alon", "Rinat", "Aime","Nitai", "Gilad", "Noam", "Omri"]

income_list = ['x', "credit", "cash", "tip", "overall income"]




def add_worker(wb, name, categories):
    ws = wb.create_sheet(name)
    for column, category in enumerate(categories):
        ws.cell(1, column + 1, category)

def add_income_sheet(wb, income_list, name):
    


def find_first_empty(ws):
    current_row = 1
    while ws.cell(current_row, 1).value:
        current_row += 1
    return current_row


def calculate_base_salary(hours, isPik, shabbat_hours):
    additional_hours = max(hours-8, 0)
    if shabbat_hours:
        additional_hours = 0
    base_sum = 35 if isPik else 30
    shabbat_sum = 45
    payment_for_regular_hours = (hours - additional_hours - shabbat_hours) * base_sum + (shabbat_hours * shabbat_sum)
    payment_for_additional_hours = additional_hours * (base_sum * 1.25)
    base_salary = payment_for_regular_hours + payment_for_additional_hours
    return math.ceil(base_salary)


def calculate_completion(base_salary, tip):
    completion = 100 if base_salary < tip + 100 else base_salary - tip
    return completion

def update_sum(ws):
    current_row = find_first_empty(ws)
    sums = [0] * 9
    for i in range(2, current_row):
        for j in range(2, 9):
            sums[j-1] += ws.cell(i, j).value
    sums[0] = 'sums'
    for i, sum in enumerate(sums):
        ws.cell(current_row + 3, i+ 1, '')
        ws.cell(current_row + 4, i+ 1, sum).font = Font(bold=True)



def update_worker(ws, date, hours, tip, shabbat_hours, isPik):
    additional_hours = max(hours-8, 0)
    base_salary = calculate_base_salary(hours, isPik, shabbat_hours)
    completion = calculate_completion(base_salary, tip)
    overall = tip + completion
    worker_details = [date, hours, base_salary, tip,  overall, completion, additional_hours, shabbat_hours, "pik" if isPik else ""]
    empty_row = find_first_empty(ws)
    for column, detail in enumerate(worker_details):
        ws.cell(empty_row, column + 1, detail)
    update_sum(ws)


def new_workbook(workers_list,):
    wb = op.Workbook()
    for worker in workers_list:
        add_worker(wb, worker, categories)
    return wb


 # wb = op.Workbook()
# add_worker(wb, "Michal", categories)
wb = op.load_workbook('worker tracing.xlsx')
update_worker(wb['Alon'], "24.8.21", 7, 323, 0, False)
update_worker(wb['Michal H'], "24.8.21", 8.5, 391, 0, False)
update_worker(wb['Rinat'], "24.8.21", 4, 0, 0, True)


# add_worker(wb, "Michal", categories)
# wb = new_workbook(worker_list)

wb.save("worker tracing.xlsx")
