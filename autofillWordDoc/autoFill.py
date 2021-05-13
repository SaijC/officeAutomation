from docxtpl import DocxTemplate
import pandas as pd

doc = DocxTemplate('wordTest.docx')
data = pd.read_excel('testData.xls')

context = dict(zip(data['var'], data['years']))
doc.render(context)
doc.save('output_wordTest.docx')
