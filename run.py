import pandas
import os
import pyodbc 
import re
import string
import xlwt 
from xlwt import Workbook 
import xlsxwriter

import openpyxl
from openpyxl import load_workbook

count = 0

def main():
    reading(count)


def reading(count):
    num = 0
    path = os.getcwd() + "\Width Lengths.xlsx"
    data = pandas.read_excel(path, sheet_name='Sheet1')
    sheet2 = pandas.read_excel(path, sheet_name='Sheet2')

    sheet3 = load_workbook('Width Lengths.xlsx')
    sheet3.create_sheet('Sheet3')

    standardLengths = []
    for index, rows in sheet2.iterrows():
        lengths = rows['Standard']
        standardLengths.append(lengths)

    items = {}
    # count = 0
    for index, row in data.iterrows():
        key = row['ItemType']
        value = row['Width']
        # print(count)

        if not (pandas.isnull(value)):
            value = int(value)

        check1 = pandas.isnull(key)
        check2 = pandas.isnull(value)

        if check1 == True and check2 == True:
            continue

        if key == "Finish":
            num = num + 1
            boolean = False
            # count = count - 1
            count = len(items) - 1

            # loop through lengths here then remove up the count from 0 to count
            lengths = []
            increment = 0
            for elem in standardLengths:
                increment = increment + 1
                lengths.append(standardLengths.pop(0))
                print(str(elem))
                if increment == count:
                    break
                # continue
            
            if len(standardLengths) == 1:
                lengths.append(standardLengths.pop(0))
                
            # for elem in enumerate(standardLengths):
            #     increment = increment + 1
            #     if increment <= count:
            #         lengths.append(standardLengths.pop(0))
            #         print(str(elem))
            #         continue
                

            query(items,sheet3,num,lengths,count)
            continue

        concat = key, value
        print(concat)

        if check1 == False and check2 == True:
            items[key] = value
            boolean = True
           
        if boolean == True:
            items[key] = value
            # count = count + 1


# compare standard demo width output column to new output column (code already creates new column on the fly)

# if standard demo column does not equal current db column then we create sql column statement

# if standard demo column does equal current db column then we continue with logic
    

def query(items,sheet3,num,lengths,count):

    # standardLengths should be indexed according to the size of the itemProperty array (using the size value)
    # compare indices to this 
    print(lengths[0])

    conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=AND692557\SQLEXPRESS;'
                      'Database=12_SP11Test;'
                      'Trusted_Connection=yes;')

    cursor = conn.cursor()
    # count = 0
    itemProperty = []
    for key,value in items.items():
        propertyName = []
        if len(key) > 0 and pandas.isnull(value):
            itemtype = key
        
        # if key and value are not null then save properties
        if len(key) > 0 and not pandas.isnull(value):
            itemProperty.append(key)
            itemProperty.append(value)  

        print(key,value)

    # Get itemtypes source ID first 
    querySource_ID = "SELECT ID FROM innovator.ITEMTYPE WHERE NAME='" +itemtype+ "'"
    print(querySource_ID)

    cursor.execute(querySource_ID)

    # Grab the ID from response
    for row in cursor:
        row = str(row).split("'",1)
        row = row[1]
        row = row.split("',",1)
        ID = row[0]

        print(ID)

    wks = sheet3['Sheet3']
    lastRow = wks.max_row

    # Query for old width lenghts and save those before running change query
    # SELECT COLUMN_WIDTH FROM innovator.PROPERTY WHERE name='property name' AND SOURCE_ID='the id we got'
    for properties in itemProperty:
        lastRow = wks.max_row  
        lastColumn = wks.max_column
        checkRowandColumn = False
        
        if not isinstance(properties, int):
            propertyName = properties

            queryOldLengths = "SELECT COLUMN_WIDTH FROM innovator.PROPERTY WHERE NAME=" + "'" +properties+ "'" " AND SOURCE_ID=" +"'" +ID+ "'"
            cursor.execute(queryOldLengths)
            if cursor.rowcount == 0:
                print("Error in " +propertyName)
                # break
            
            for row in cursor:
                row = str(row).split(",",1)
                row = row[0].split("(",1)
                if str(row[1]) == "None":
                    row = "null"
                else:
                    row = int(row[1])         
                
                if lastColumn == 1:
                    wks.cell(row=lastRow, column=1).value = itemtype
                    wks.cell(row=lastRow, column=2).value = properties
                    wks.cell(row=lastRow, column=5).value = row
                    sheet3.save('Width Lengths.xlsx')
                    checkRowandColumn = True
                    
                if checkRowandColumn == False:
                    if num >= 2:
                        emptyRow = ""
                        wks.cell(row=lastRow+1, column=1).value = itemtype

                    wks.cell(row=lastRow+1, column=2).value = properties
                    wks.cell(row=lastRow+1, column=5).value = row
                    sheet3.save('Width Lengths.xlsx')

      
                checkRowandColumn == False

        widthLength = 0
        if isinstance(properties, int):
            widthLength = properties
        
        # UPDATE innovator.PROPERTY SET COLUMN_WIDTH='new int width' WHERE NAME='property name' AND SOURCE_ID='ID variable'
        # Setup second query using source ID and property values 
        if not widthLength == 0 and not propertyName == "":
            queryOldLengths = "UPDATE innovator.PROPERTY SET COLUMN_WIDTH=" + "'" +str(widthLength)+ "'" + " WHERE NAME=" + "'" +propertyName+ "'" +" AND SOURCE_ID=" + "'" +ID+ "'"
            # cursor.execute(queryOldLengths)
            
            wks.cell(row=lastRow, column=6).value = queryOldLengths
            sheet3.save('Width Lengths.xlsx')
            lastRow = wks.max_row
            checkRowandColumn = True

            checkRowandColumn == False
 
    print(num)
    # increment(num)
    # count = 0
               
    items.clear()
    lengths.clear()
    # return count = 0


# def increment(num):
#     if num >= 2:
#         return count = 0


main()