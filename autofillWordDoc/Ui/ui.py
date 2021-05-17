
"""Horizontal layout example."""

import sys
import PyQt5.QtWidgets as QtWidgets
from officeAutomation.autofillWordDoc.constants import constants as CONST

class MainDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('BLA')
        # self.setMinimumSize(470, 115)
        self.createWidgets()
        self.createLayouts()
        self.createConnections()

    def createWidgets(self):
        self.LoadTempLineEdit = QtWidgets.QLineEdit('Load Template File')
        self.loadTempPathBtn = QtWidgets.QPushButton('...')
        self.dividerLable1 = QtWidgets.QLabel('|')
        self.LoadExcelLineEdit = QtWidgets.QLineEdit('Load Excel File')
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

        self.layoutTop.addWidget(self.LoadTempLineEdit)
        self.layoutTop.addWidget(self.loadTempPathBtn)
        self.layoutTop.addWidget(self.dividerLable1)
        self.layoutTop.addWidget(self.LoadExcelLineEdit)
        self.layoutTop.addWidget(self.loadExcelPathBtn)

        self.layoutMid.addWidget(self.mappingQComboBox)

        self.layoutBot.addWidget(self.okBtn)
        self.layoutBot.addWidget(self.cancelBtn)
        self.setLayout(self.layoutMain)

    def createConnections(self):
        self.loadTempPathBtn.clicked.connect(self.setTempfilePath)
        self.loadExcelPathBtn.clicked.connect(self.setExcelfilePath)

    def setTempfilePath(self):
        """
        pyside2 connection instructions for setPath button
        :return: None
        """
        file_path, selected_filter = QtWidgets.QFileDialog.getOpenFileName(self,
                                                                           "Select File",
                                                                           CONST.ROOTPATH,
                                                                           '*.docx'
                                                                           )
        if file_path:
            self.LoadTempLineEdit.setText(file_path)

    def setExcelfilePath(self):
        """
        pyside2 connection instructions for setPath button
        :return: None
        """
        file_path, selected_filter = QtWidgets.QFileDialog.getOpenFileName(self,
                                                                           "Select File",
                                                                           CONST.ROOTPATH,
                                                                           "Excel (*.xlsx *.xls)")
        if file_path:
            self.LoadTempLineEdit.setText(file_path)

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    dlg = MainDialog()
    dlg.show()
    sys.exit(app.exec_())
