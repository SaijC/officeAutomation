"""
Automating excel to word
"""

import pandas as pd
from docxtpl import DocxTemplate
from officeAutomation.autofillWordDoc.constants import constants as CONST


def replaceSpace(inputStr):
    """
    Replace any space character with '_'
    :param inputStr: string
    :return: strin
    """
    if ' ' in inputStr:
        inputStr = inputStr.replace(' ', '_')
    return inputStr


def autoFill(inputExcelFile, templateWordFile):
    """
    auto fill a docx template
    :param inputExcelFile: xlsx or xls file
    :param templateWordFile: docx file
    :return: None
    """
    data = pd.read_excel(inputExcelFile)
    pdData = pd.DataFrame(data)
    numIDX = len(pdData.index)

    for idx in range(numIDX):
        doc = DocxTemplate(templateWordFile)

        context = dict()
        for columName, row in pdData.iteritems():
            if 'Unnamed' in columName:
                continue
            newColumName = replaceSpace(inputStr=columName)
            context.update({newColumName: row[idx]})
        doc.render(context)

        itemIDX = context['id']
        savePath = '{}\\{}.docx'.format(CONST.OUTPUTSPATH, itemIDX)
        doc.save(savePath)


inputExcel = '{}\\{}'.format(CONST.INPUTSPATH, 'inputData.xlsx')
wordTemplate = '{}\\{}'.format(CONST.TEMPLATES, 'wordTemplate.docx')

autoFill(inputExcelFile=inputExcel, templateWordFile=wordTemplate)
