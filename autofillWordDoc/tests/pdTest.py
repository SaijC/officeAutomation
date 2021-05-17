import pandas as pd
from officeAutomation.utils.utils import stringUtils as strUtils

excelFilePath = 'C:\\tools\\officeAutomation\\autofillWordDoc\\inputs\\inputData.xls'
data = pd.read_excel(excelFilePath)
pdData = pd.DataFrame(data)
pdDataDict = pdData.to_dict()
pdDataList = pdData.values.tolist()
numIDX = len(pdData.index)

greenList = ['CTL-000148', 'CTL-000149', 'CTL-000150', 'CTL-000151', 'CTL-000152',
             'CTL-000153', 'CTL-000161', 'CTL-000162', 'CTL-000163', 'CTL-000164']

newList = list()

for i in pdDataList:
    if i[0] in greenList:
        newList.append(i)

for i in newList:
    print(i)

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
