import sys
import openpyxl
from openpyxl.comments import Comment


def get_workbook():
    workbook = None
    if len(sys.argv) > 1:
        workbook = openpyxl.load_workbook(sys.argv[1])
        return workbook
    print("Not enough arguments please  give path to excel sheet")
    return workbook


def sort_sheet(sheet):
    max_row = sheet.max_row
    max_column = sheet.max_column
    sheet.cell(column=max_column + 1, row=1).value = "Student Total"
    for i in range(2, max_row + 1):
        count: int = 0
        for k in range(2, max_column + 1):
            if sheet.cell(column=k, row=i).value:
                count = count + sheet.cell(column=k, row=i).value
        sheet.cell(column=max_column + 1, row=i).value = count
    sheet.cell(column=1, row=max_row + 1).value = "Question Total"
    for i in range(2, max_column + 1):
        count: int = 0
        for k in range(2, max_row + 1):
            if sheet.cell(column=i, row=k).value:
                count = count + sheet.cell(column=i, row=k).value
        sheet.cell(column=i, row=max_row + 1).value = count


def main():
    workbook = get_workbook()
    if workbook:
        for sheet in workbook.sheetnames:
            sort_sheet(workbook[sheet])
        workbook.save("Sorted_Sheet.xlsx")


main()
