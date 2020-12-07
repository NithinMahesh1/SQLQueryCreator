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

def main():
    reading()


def reading():
    num = 0
    path = os.getcwd() + "\Width Lengths.xlsx"
    data = pandas.read_excel(path, sheet_name='Sheet1')

    sheet2 = load_workbook('Width Lengths.xlsx')
    sheet2.create_sheet('Sheet2')

    items = {}
    for index, row in data.iterrows():
        key = row['ItemType']
        value = row['Width']

        if not (pandas.isnull(value)):
            value = int(value)

        check1 = pandas.isnull(key)
        check2 = pandas.isnull(value)

        if check1 == True and check2 == True:
            continue

        if key == "Finish":
            num = num + 1
            boolean = False
            query(items,sheet2, num)
            continue

        concat = key, value

        if check1 == False and check2 == True:
            items[key] = value
            boolean = True
           
        if boolean == True:
            items[key] = value    

    

def query(items,sheet2,num):

    conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=AND692557\SQLEXPRESS;'
                      'Database=12_SP11;'
                      'Trusted_Connection=yes;')

    cursor = conn.cursor()
    count = 0
    itemProperty = []
    for key,value in items.items():
        propertyName = []
        if len(key) > 0 and pandas.isnull(value):
            itemtype = key
        
        # if key and value are not null then save properties
        if len(key) > 0 and not pandas.isnull(value):
            itemProperty.append(key)
            # add name here
            # propertyName = key
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

    count = 0
    wks = sheet2['Sheet2']
    lastRow = wks.max_row
    # wks.cell(row=lastRow, column=1).value = itemtype

    # Query for old width lenghts and save those before running change query
    # SELECT COLUMN_WIDTH FROM innovator.PROPERTY WHERE name='property name' AND SOURCE_ID='the id we got'
    for properties in itemProperty:
        lastRow = wks.max_row
        # count = count + 1
        # count = lastRow  
        lastColumn = wks.max_column
        checkRowandColumn = False
        print(lastColumn)
        count = count + 1
        size = len(itemProperty)
        index = size/2
        if size == 2:
            queryOldLengths = "SELECT COLUMN_WIDTH FROM innovator.PROPERTY WHERE NAME=" + "'" +properties[0]+ "'" " AND SOURCE_ID=" +"'" +ID+ "'"
            cursor.execute(queryOldLengths)

            for row in cursor:
                row = str(row).split(",",1)
                row = row[0].split("(",1)
                row = int(row[1])
                wks = sheet2['Sheet2']
                
                if lastColumn == 1:
                    wks.cell(row=lastRow, column=1).value = itemtype
                    wks.cell(row=lastRow+2, column=1).value = row
                    sheet2.save('Width Lengths.xlsx')
                    checkRowandColumn = True
                    # count = count - 1

                if checkRowandColumn == False:
                    wks.cell(row=lastRow+1, column=1).value = row
                    sheet2.save('Width Lengths.xlsx')
      
                checkRowandColumn == False
                print(row)

        
        # There are more than 1 property values
        
        if not isinstance(properties, int):
            propertyName = properties

            queryOldLengths = "SELECT COLUMN_WIDTH FROM innovator.PROPERTY WHERE NAME=" + "'" +properties+ "'" " AND SOURCE_ID=" +"'" +ID+ "'"
            cursor.execute(queryOldLengths)
            
            for row in cursor:
                row = str(row).split(",",1)
                row = row[0].split("(",1)
                if str(row[1]) == null:
                    row = "null"
                row = int(row[1])         
                
                if lastColumn == 1:
                    wks.cell(row=lastRow, column=1).value = itemtype
                    wks.cell(row=lastRow, column=2).value = properties
                    wks.cell(row=lastRow, column=4).value = row
                    sheet2.save('Width Lengths.xlsx')
                    checkRowandColumn = True
                    
                if checkRowandColumn == False:
                    if num >= 2:
                        emptyRow = ""
                        wks.cell(row=lastRow+1, column=1).value = itemtype
                        # wks.cell(row=lastRow+2, column=1).value = emptyRow

                    wks.cell(row=lastRow+1, column=2).value = properties
                    wks.cell(row=lastRow+1, column=4).value = row
                    sheet2.save('Width Lengths.xlsx')
                    # lastRow = wks.max_row
                # count = count - 1
                checkRowandColumn == False
                print(row)

        widthLength = 0
        if isinstance(properties, int):
            widthLength = properties
        
        # UPDATE innovator.PROPERTY SET COLUMN_WIDTH='new int width' WHERE NAME='property name' AND SOURCE_ID='ID variable'
        
        if not widthLength == 0 and not propertyName == "":
            queryOldLengths = "UPDATE innovator.PROPERTY SET COLUMN_WIDTH=" + "'" +str(widthLength)+ "'" + " WHERE NAME=" + "'" +propertyName+ "'" +" AND SOURCE_ID=" + "'" +ID+ "'"
            # cursor.execute(queryOldLengths)
            
            # if lastColumn == 1:
            wks.cell(row=lastRow, column=6).value = queryOldLengths
            sheet2.save('Width Lengths.xlsx')
            lastRow = wks.max_row
            checkRowandColumn = True

            # if checkRowandColumn == False:
            #     wks.cell(row=lastRow+1, column=6).value = queryOldLengths
            #     sheet2.save('Width Lengths.xlsx')
            #     lastRow = wks.max_row

            checkRowandColumn == False
            # for row in cursor:
            #     print(row)
        
        
    
    # Setup second query using source ID and property values 

    items.clear()


main()