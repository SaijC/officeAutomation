import pandas as pd
from officeAutomation.utils.utils import stringUtils as strUtils

excelFilePath = 'C:\\tools\\officeAutomation\\autofillWordDoc\\inputs\\inputData.xls'
data = pd.read_excel(excelFilePath)
pdData = pd.DataFrame(data)
pdDataList = pdData.values.tolist()
numIDX = len(pdData.index)

greenList = ['CTL-000148', 'CTL-000149', 'CTL-000150', 'CTL-000151', 'CTL-000152',
             'CTL-000153', 'CTL-000161', 'CTL-000162', 'CTL-000163', 'CTL-000164']

columnNames = list(pdData.columns.values)
greenListItems = list()

for i in pdDataList:
    if i[0] in greenList:
        greenListItems.append(i)

contextList = list()
for itemList in greenListItems:
    contextDict = dict()
    keyValuePair = zip(columnNames, itemList)
    for key, val in keyValuePair:
        if 'Unnamed' in key:
            continue
        newKeyName = strUtils().replaceSpace(inputStr=key)
        contextDict.update({newKeyName: val})
    contextList.append(contextDict)

for i in contextList:
    print(i)
# contexList = list()
# context = dict()
# for key, val in newList:
#     context.update({key: val})


# for idx in range(numIDX):
#     context = dict()
#     for columName, row in pdData.iteritems():
#         if 'Unnamed' in columName:
#             continue
#         newColumName = strUtils().replaceSpace(inputStr=columName)
#         if row[idx] in greenList:
#             context.update({newColumName: row[idx]})
#         else:
#             pass
#             # context.update({newColumName: row[idx]})
#     print(context)
