import os

# Folder paths
_currentFilePath = os.path.abspath(__file__)
_currentFilePathSplit = _currentFilePath.split('\\')[:-2]

ROOTPATH = '\\'.join(_currentFilePathSplit)
OUTPUTSPATH = '{}\\{}'.format(ROOTPATH, 'outputs')
INPUTSPATH = '{}\\{}'.format(ROOTPATH, 'inputs')
TEMPLATES = '{}\\{}'.format(ROOTPATH, 'templates')

# File extensions
EXCELDOC = 'Excel (*.xlsx *.xls)'
WORDDOC = 'Word *.docx'
