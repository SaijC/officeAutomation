import pandas as pd
from docxtpl import DocxTemplate
from officeAutomation.autofillWordDoc.constants import constants as CONST


doc = DocxTemplate('{}\\{}'.format(CONST.TEMPLATES, 'wordTest.docx'))
data = pd.read_excel('{}\\{}'.format(CONST.INPUTSPATH, 'testData.xlsx'))
pdData = pd.DataFrame(data)

def replaceSpace(inputStr):
    if ' ' in inputStr:
        inputStr = inputStr.replace(' ', '_')
    return inputStr

def toFile(filename=''):
    print(CONST.TEMPLATES)

# for idx in range(len(pdData)):
#     context = dict()
#     for columName, row in pdData.iteritems():
#         if 'Unnamed' in columName:
#             continue
#         newColumName = replaceSpace(columName)
#         context.update({newColumName: row[idx]})
#     print(context)
#     renders = doc.render(context)
#     # doc.save('output_wordTest.docx')

toFile()
