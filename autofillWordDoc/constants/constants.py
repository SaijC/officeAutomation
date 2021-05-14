import os

_currentFilePath = os.path.abspath(__file__)
_currentFilePathSplit = _currentFilePath.split('\\')[:-2]

ROOTPATH = '\\'.join(_currentFilePathSplit)
OUTPUTSPATH = '{}\\{}'.format(ROOTPATH, 'outputs')
INPUTSPATH = '{}\\{}'.format(ROOTPATH, 'inputs')
TEMPLATES = '{}\\{}'.format(ROOTPATH, 'templates')

INPUTDATAFILENAME = '{}\\{}'.format(INPUTSPATH, 'inputData.xlsx')
WORDTEMPLATEFILENAME = '{}\\{}'.format(TEMPLATES, 'wordTemplate.docx')
