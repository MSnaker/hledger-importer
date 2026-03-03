from openpyxl import Workbook
from openpyxl import load_workbook
wb = load_workbook(filename = "INPUT/Balance_2018.xlsx") 

print(wb.sheetnames)
