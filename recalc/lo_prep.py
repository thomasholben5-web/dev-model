import sys, openpyxl
from openpyxl.workbook.properties import CalcProperties
wb=openpyxl.load_workbook(sys.argv[1]); wb.calculation=CalcProperties(fullCalcOnLoad=True)
wb.save(sys.argv[2]); print("prepped")
