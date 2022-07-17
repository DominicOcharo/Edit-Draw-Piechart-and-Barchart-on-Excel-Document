# Edit-Draw-Piechart-and-Barchart-on-Excel-Document
#python project

import openpyxl as xl
from openpyxl.chart import PieChart, BarChart, Reference


class DocumentEdit:
    def __init__(self, name, new_name):
        self.name = name
        self.new = new_name

    def edit_excel(self):
        wb = xl.load_workbook(self.name)
        sheet = wb['Sheet1']
        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row, 3)
            corrected_value = cell.value * 0.8
            corrected_cell = sheet.cell(row, 4)
            corrected_cell.value = corrected_value

        values = Reference(sheet,
                           max_row=sheet.max_row,
                           min_row=2,
                           max_col=4,
                           min_col=4)
        chart1 = PieChart()
        chart1.add_data(values)
        sheet.add_chart(chart1, 'a8')

        chart2 = BarChart()
        chart2.add_data(values)
        sheet.add_chart(chart2, 'e2')
        wb.save(self.new)


name = input("input the name of the excel document as (name.xlsx): ")
new_name = input("input the name of the updated excel document as (name.xlsx): ")
check = DocumentEdit(name, new_name)
check.edit_excel()
