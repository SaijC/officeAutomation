"""
Automating excel to word
"""
import pandas as pd
from docxtpl import DocxTemplate
from officeAutomation.utils.utils import StringUtils as strUtils


class AutoFill:
    def __init__(self, inputExcelFile, templateWordFile, outputPath):
        self.inputExcelFile = inputExcelFile
        self.templateWordFile = templateWordFile
        self.outputPath = outputPath

    def autoFill(self):
        """
        auto fill a docx template
        :param inputExcelFile: xlsx or xls file
        :param templateWordFile: docx file
        :return: None
        """
        data = pd.read_excel(self.inputExcelFile)
        pdData = pd.DataFrame(data)
        numIDX = len(pdData.index)

        for idx in range(numIDX):
            doc = DocxTemplate(self.templateWordFile)

            context = dict()
            for columName, row in pdData.iteritems():
                if 'Unnamed' in columName:
                    continue
                newColumName = strUtils().replaceSpace(inputStr=columName)
                context.update({newColumName: row[idx]})

            print(context)
            doc.render(context)

            itemIDX = context['id']
            savePath = '{}\\{}.docx'.format(self.outputPath, itemIDX)
            doc.save(savePath)

af = AutoFill(inputExcelFile='C:\\tools\\officeAutomation\\autofillWordDoc\\inputs\\inputData.xls',
              templateWordFile='C:\\tools\\officeAutomation\\autofillWordDoc\\templates\\wordTemplate.docx',
              outputPath='C:\\tools\\officeAutomation\\autofillWordDoc\\outputs')

af.autoFill()