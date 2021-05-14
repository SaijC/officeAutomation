import pandas as pd
from docxtpl import DocxTemplate
from officeAutomation.autofillWordDoc.constants import constants as CONST

data = pd.read_excel(CONST.INPUTDATAFILENAME)
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
    doc = DocxTemplate(CONST.WORDTEMPLATEFILENAME)
    context = buildContext(idx)
    renders = doc.render(context)
    itemIDX = context['id']
    savePath = '{}\\{}.docx'.format(CONST.OUTPUTSPATH, itemIDX)
    doc.save(savePath)
