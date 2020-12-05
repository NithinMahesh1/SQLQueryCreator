import pandas
import os
import sqlite3


def main():
    reading()


def reading():
    path = os.getcwd() + "\Width Lengths.xlsx"
    data = pandas.read_excel(path, sheet_name='Sheet1')

    itemtype = {}
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
            query(itemtype)
            continue

        concat = key, value

        if check1 == False and check2 == True:
            itemtype[key] = value
            boolean = True
           
        if boolean == True:
            itemtype[key] = value    


        print(concat)
        if(index == 790):
            print(itemtype)
    

def query(itemtype):
    print(itemtype)
    itemtype.clear()


main()