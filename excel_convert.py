from numpy import double
import pandas as pd
import xlsxwriter
from pandas.core.frame import DataFrame

def function1():
    readexcel = pd.ExcelFile("ExcelBook.xlsx")
    df = pd.read_excel(readexcel, 0)
    dataframe = pd.DataFrame(df)

    supplier_gstin = dataframe["Supplier GSTIN"]
    sum_18 = sum_28 = sum_5 = sum_12= 0

    workbook = xlsxwriter.Workbook("workbook.xlsx")
    worksheet = workbook.add_worksheet("My Sheet")

    d = dict()
    for i in supplier_gstin:
        if i in d.keys():
            d[i]+=1
        else:
            d[i] = 1

    print(d)

    row = 1
    column = 1

    count_18 = count_28 = count_12 = count_5 = count_3 = count_0 = count_75 = 0
    for i in d.keys():
        sum_28 = sum_18 = sum_12 = sum_5 = sum_75 = sum_3 = sum_0 = 0
        for j in range(len(dataframe["Supplier GSTIN"])):
            if i == dataframe["Supplier GSTIN"][j] and dataframe["Tax Rate"][j] == 18:
                sum_18 += dataframe["Taxable Amount"][j]
                count_18 +=1
            elif i == dataframe["Supplier GSTIN"][j] and dataframe["Tax Rate"][j] == 28:
                sum_28 += dataframe["Taxable Amount"][j]
                count_28 += 1
            elif i == dataframe["Supplier GSTIN"][j] and dataframe["Tax Rate"][j] == 12:
                sum_12 += dataframe["Taxable Amount"][j]
                count_12 += 1
            elif i == dataframe["Supplier GSTIN"][j] and dataframe["Tax Rate"][j] == 5:
                sum_5 += dataframe["Taxable Amount"][j]
                count_5 += 1
            elif i == dataframe["Supplier GSTIN"][j] and dataframe["Tax Rate"][j] == 3:
                sum_3 += dataframe["Taxable Amount"][j]
                count_3 += 1
            elif i == dataframe["Supplier GSTIN"][j] and dataframe["Tax Rate"][j] == 0:
                sum_0 += dataframe["Taxable Amount"][j]
                count_0 += 1
            elif i == dataframe["Supplier GSTIN"][j] and dataframe["Tax Rate"][j] == 7.5:
                sum_75 += dataframe["Taxable Amount"][j]
                count_75 += 1
        worksheet.write(row, 3, sum_18)
        worksheet.write(row, 5, sum_28)
        worksheet.write(row, 7, sum_12)
        worksheet.write(row, 9, sum_5)
        worksheet.write(row, 11, sum_0)
        worksheet.write(row, 13, sum_3)
        worksheet.write(row, 15, sum_75)
        row+=1
        
    print("count_12 = {}, count_18 = {}, count_28 = {}, count_5 = {}".format(count_12, count_18, count_28, count_5))

    worksheet.write('A1', "Supplier GstIn")
    worksheet.write('D1', "18%")
    worksheet.write('F1', "28%")
    worksheet.write('H1', "12%")
    worksheet.write('J1', "5%")
    worksheet.write('L1', "0%")
    worksheet.write('N1', "3%")
    worksheet.write('P1', "7.5%")

    column = 0
    row = 1
    for i in d.keys():
        worksheet.write(row, column, i)
        row+=1

    workbook.close()
    print(row)

def function2():
    one = pd.ExcelFile("c://users//hp//downloads//1.xlsx") #Provide the Location of the Excel files
    two = pd.ExcelFile("c://users//hp//downloads//2.xlsx") #Provide the Location of the Excel files
    three = pd.ExcelFile("c://users//hp//downloads//3.xlsx") #Provide the Location of the Excel files
    four = pd.ExcelFile("c://users//hp//downloads//4.xlsx") #Provide the Location of the Excel files

    excel_one = pd.read_excel(one, 2)
    excel_two = pd.read_excel(two, 2)
    excel_three = pd.read_excel(three, 2)
    excel_four = pd.read_excel(four, 2)

    workbook = xlsxwriter.Workbook("ExcelBook.xlsx")
    worksheet = workbook.add_worksheet("My Sheet")

    worksheet.write("A1", "Supplier GSTIN")
    worksheet.write("E1", "Tax Rate")
    worksheet.write("H1", "Taxable Amount")

    column = 0
    row = 1
    for i in excel_four["Supplier GSTIN"]:
        worksheet.write(row, column, i)
        row += 1

    for i in excel_one["Supplier GSTIN"]:
        worksheet.write(row, column, i)
        row+=1

    for i in excel_two["Supplier GSTIN"]:
        worksheet.write(row, column, i)
        row+=1

    for i in excel_three["Supplier GSTIN"]:
        worksheet.write(row, column, i)
        row+=1


    column = 4
    row = 1
    for i in excel_four["Tax Rate"]:
        worksheet.write(row, column, i)
        row += 1

    for i in excel_one["Tax Rate"]:
        worksheet.write(row, column, i)
        row+=1

    for i in excel_two["Tax Rate"]:
        worksheet.write(row, column, i)
        row+=1

    for i in excel_three["Tax Rate"]:
        worksheet.write(row, column, i)
        row+=1


    column = 7
    row = 1
    for i in excel_four["Taxable Amount"]:
        worksheet.write(row, column, i)
        row += 1

    for i in excel_one["Taxable Amount"]:
        worksheet.write(row, column, i)
        row+=1

    for i in excel_two["Taxable Amount"]:
        worksheet.write(row, column, i)
        row+=1

    for i in excel_three["Taxable Amount"]:
        worksheet.write(row, column, i)
        row+=1

    workbook.close()

if __name__ == "__main__":
    function2()
    function1()