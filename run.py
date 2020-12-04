import pandas
import os

path = os.getcwd() + "\Width Lengths.xlsx"

data = pandas.read_excel(path, sheet_name='Sheet1')

# Gets certain columns 
# test = data.iloc[0:3,0:5]
# df_new=data.iloc[:data.loc[data.ItemType.str.contains('Finish',na=False)].index[0]]
# firstColumn = data.iloc[0:,0:1]


for index, row in data.iterrows():
    itemtype = row['ItemType']
    width = row['Width']

    if not (pandas.isnull(width)):
        width = int(width)

    # if not (pandas.isnull(itemtype) and pandas.isnull(width)):

    check1 = pandas.isnull(itemtype)
    check2 = pandas.isnull(width)

    if check1 == True and check2 == True:
        continue

    if check1 == False and check2 == False:
        concat = itemtype, width

    if check1 == False and check2 == True:
        concat = itemtype, width
 
    print(concat)

