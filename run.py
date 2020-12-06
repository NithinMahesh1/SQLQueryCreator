import pandas
import os
import pyodbc 
import re
import string

def main():
    reading()


def reading():
    path = os.getcwd() + "\Width Lengths.xlsx"
    data = pandas.read_excel(path, sheet_name='Sheet1')


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
            boolean = False
            query(items)
            continue

        concat = key, value

        if check1 == False and check2 == True:
            items[key] = value
            boolean = True
           
        if boolean == True:
            items[key] = value    

    

def query(items):

    conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=\SQLEXPRESS;'
                      'Database=12_SP11;'
                      'Trusted_Connection=yes;')

    cursor = conn.cursor()
    count = 0
    itemProperty = []
    # itemPropertyValue = []
    for key,value in items.items():
        
        if len(key) > 0 and pandas.isnull(value):
            itemtype = key
        
        # if key and value are not null then save properties
        if len(key) > 0 and not pandas.isnull(value):
            # count = count + 1
            itemProperty.append(key)
            itemProperty.append(value)
            # itemPropertyValue.append(value)
            # if count > 1:
            #     itemProperty.append(key)
            #     itemPropertyValue.append(value)

            # if its count is greater than 1 then add that property
            

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

    # Query for old width lenghts and save those before running change query
    # SELECT COLUMN_WIDTH FROM innovator.PROPERTY WHERE name='property name' AND SOURCE_ID='the id we got'
    for properties in itemProperty:
        size = len(itemProperty)
        if size == 2:
            queryOldLengths = "SELECT COLUMN_WIDTH FROM innovator.PROPERTY WHERE NAME=" + "'" +properties[0]+ "'" " AND SOURCE_ID=" +"'" +ID+ "'"
            cursor.execute(queryOldLengths)
        
        # There are more than 1 property values
        queryOldLengths = "SELECT COLUMN_WIDTH FROM innovator.PROPERTY WHERE NAME=" + "'" +properties+ "'" " AND SOURCE_ID=" +"'" +ID+ "'"
        cursor.execute(queryOldLengths)
        for row in cursor:
            print(row)
        # print("true")
    
    # Setup second query using source ID and property values 

    items.clear()


main()