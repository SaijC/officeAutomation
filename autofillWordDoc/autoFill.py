import pandas as pd
from docxtpl import DocxTemplate
from officeAutomation.autofillWordDoc.constants import constants as CONST


data = pd.read_excel('{}\\{}'.format(CONST.INPUTSPATH, 'testData.xlsx'))
pdData = pd.DataFrame(data)

def replaceSpace(inputStr):
    if ' ' in inputStr:
        inputStr = inputStr.replace(' ', '_')
    return inputStr

def buildContext(idx):
    context = dict()
    for columName, row in pdData.iteritems():
        if 'Unnamed' in columName:
            continue
        newColumName = replaceSpace(columName)
        context.update({newColumName: row[idx]})
    return context

for idx in range(len(pdData)):
    doc = DocxTemplate('{}\\{}'.format(CONST.TEMPLATES, 'wordTest.docx'))
    context = buildContext(idx)
    renders = doc.render(context)
    itemID = context['id']
    savePath = '{}\\{}.docx'.format(CONST.OUTPUTSPATH, itemID)
    doc.save(savePath)

