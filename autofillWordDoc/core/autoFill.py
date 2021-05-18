"""
Automating excel to word
"""
from docxtpl import DocxTemplate


class AutoFill:
    def __init__(self, templateWordFile, outputPath, contextList):
        self.templateWordFile = templateWordFile
        self.contextList = contextList
        self.outputPath = outputPath

    def autoFill(self):
        """
        auto fill a docx template
        :param inputExcelFile: xlsx or xls file
        :param templateWordFile: docx file
        :return: None
        """

        for context in self.contextList:
            doc = DocxTemplate(self.templateWordFile)
            doc.render(context)
            itemIDX = context['id']
            savePath = '{}\\{}.docx'.format(self.outputPath, itemIDX)
            doc.save(savePath)
