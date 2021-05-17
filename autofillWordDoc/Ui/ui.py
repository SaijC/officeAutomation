
"""Horizontal layout example."""

import sys
import logging
import PyQt5.QtWidgets as QtWidgets
from officeAutomation.autofillWordDoc.constants import constants as CONST
from officeAutomation.autofillWordDoc.core.autoFill import AutoFill

logging.basicConfig(level=logging.DEBUG)

class MainDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('BLA')
        # self.setMinimumSize(470, 115)
        self.tempFilePath = ''
        self.excelFilePath = ''
        self.createWidgets()
        self.createLayouts()
        self.createConnections()

    def createWidgets(self):
        self.loadTempLineEdit = QtWidgets.QLineEdit(CONST.TEMPLATES)
        self.loadTempPathBtn = QtWidgets.QPushButton('...')
        self.dividerLable1 = QtWidgets.QLabel('|')
        self.loadExcelLineEdit = QtWidgets.QLineEdit(CONST.INPUTSPATH)
        self.loadExcelPathBtn = QtWidgets.QPushButton('...')

        self.mappingQComboBox = QtWidgets.QComboBox()

        self.okBtn = QtWidgets.QPushButton('OK')
        self.cancelBtn = QtWidgets.QPushButton('Cancel')

    def createLayouts(self):
        self.layoutTop = QtWidgets.QHBoxLayout()
        self.layoutMid = QtWidgets.QVBoxLayout()
        self.layoutBot = QtWidgets.QHBoxLayout()
        self.layoutMain = QtWidgets.QVBoxLayout()

        self.layoutMain.addLayout(self.layoutTop)
        self.layoutMain.addLayout(self.layoutMid)
        self.layoutMain.addLayout(self.layoutBot)

        self.layoutTop.addWidget(self.loadTempLineEdit)
        self.layoutTop.addWidget(self.loadTempPathBtn)
        self.layoutTop.addWidget(self.dividerLable1)
        self.layoutTop.addWidget(self.loadExcelLineEdit)
        self.layoutTop.addWidget(self.loadExcelPathBtn)

        self.layoutMid.addWidget(self.mappingQComboBox)

        self.layoutBot.addWidget(self.okBtn)
        self.layoutBot.addWidget(self.cancelBtn)
        self.setLayout(self.layoutMain)

    def createConnections(self):
        self.loadTempPathBtn.clicked.connect(self.setTempfilePath)
        self.loadExcelPathBtn.clicked.connect(self.setExcelfilePath)
        self.okBtn.clicked.connect(self.doIt)
        self.cancelBtn.clicked.connect(self.close)

    def setTempfilePath(self):
        """
        pyside2 connection instructions for setPath button
        :return: None
        """
        file_path, selected_filter = QtWidgets.QFileDialog.getOpenFileName(self,
                                                                           "Select File",
                                                                           CONST.TEMPLATES,
                                                                           'Word *.docx'
                                                                           )
        if file_path:
            self.tempFilePath = file_path
            self.loadTempLineEdit.setText(file_path)

    def setExcelfilePath(self):
        """
        pyside2 connection instructions for setPath button
        :return: None
        """

        file_path, selected_filter = QtWidgets.QFileDialog.getOpenFileName(self,
                                                                           "Select File",
                                                                           CONST.INPUTSPATH,
                                                                           "Excel (*.xlsx *.xls)")
        if file_path:
            self.excelFilePath = file_path
            self.loadExcelLineEdit.setText(file_path)

    def doIt(self):
        af = AutoFill(inputExcelFile=self.excelFilePath,
                      templateWordFile=self.tempFilePath,
                      outputPath=CONST.OUTPUTSPATH)

        af.autoFill()

    def closeWindow(self):
        self.exit()

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    dlg = MainDialog()
    dlg.show()
    sys.exit(app.exec_())
