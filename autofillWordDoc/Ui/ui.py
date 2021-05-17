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
        self.setWindowTitle('Auto Fill')
        self.setMinimumSize(600, 115)
        self.tempFilePath = ''
        self.excelFilePath = ''
        self.createWidgets()
        self.createLayouts()
        self.createConnections()

    def createWidgets(self):
        self.loadExcelLable = QtWidgets.QLabel('Excel Doc')
        self.loadExcelLineEdit = QtWidgets.QLineEdit(CONST.INPUTSPATH)
        self.loadExcelPathBtn = QtWidgets.QPushButton('select')

        self.dividerLable1 = QtWidgets.QLabel('|')

        self.loadTempLable = QtWidgets.QLabel('Template Doc')
        self.loadTempLineEdit = QtWidgets.QLineEdit(CONST.TEMPLATES)
        self.loadTempPathBtn = QtWidgets.QPushButton('select')

        self.mappingQComboBox = QtWidgets.QComboBox()

        self.okBtn = QtWidgets.QPushButton('OK')
        self.cancelBtn = QtWidgets.QPushButton('Cancel')

    def createLayouts(self):

        # Create Layouts
        self.layoutTop = QtWidgets.QHBoxLayout()
        self.layoutTopLeft = QtWidgets.QVBoxLayout()
        self.layoutTopRight = QtWidgets.QVBoxLayout()
        self.layoutTopRightHBox = QtWidgets.QHBoxLayout()
        self.layoutMid = QtWidgets.QVBoxLayout()
        self.layoutBot = QtWidgets.QHBoxLayout()
        self.layoutMain = QtWidgets.QVBoxLayout()

        # TOP left layout
        self.layoutTopLeft.addWidget(self.loadExcelLable)
        self.layoutTopLeftHBox = QtWidgets.QHBoxLayout()
        self.layoutTopLeftHBox.addWidget(self.loadExcelLineEdit)
        self.layoutTopLeftHBox.addWidget(self.loadExcelPathBtn)

        self.layoutTopLeft.addLayout(self.layoutTopLeftHBox)  # add to top left layout

        # TOP right layout
        self.layoutTopRight.addWidget(self.loadTempLable)
        self.layoutTopRightHBox.addWidget(self.loadTempLineEdit)
        self.layoutTopRightHBox.addWidget(self.loadTempPathBtn)
        self.layoutTopRight.addLayout(self.layoutTopRightHBox)  # add to top right layout

        # add layout to main TOP layout
        self.layoutTop.addLayout(self.layoutTopLeft)
        self.layoutTop.addLayout(self.layoutTopRight)

        # MIDDLE layout
        self.layoutMid.addWidget(self.mappingQComboBox)

        # BOTTOM layout
        self.layoutBot.addWidget(self.okBtn)
        self.layoutBot.addWidget(self.cancelBtn)

        # add to main layout
        self.layoutMain.addLayout(self.layoutTop)
        self.layoutMain.addLayout(self.layoutMid)
        self.layoutMain.addLayout(self.layoutBot)
        self.setLayout(self.layoutMain)

    def createConnections(self):
        self.loadTempPathBtn.clicked.connect(self.setTempfilePath)
        self.loadExcelPathBtn.clicked.connect(self.setExcelfilePath)
        self.okBtn.clicked.connect(self.doIt)
        self.cancelBtn.clicked.connect(self.close)

    def setExcelfilePath(self):
        """
        pyside2 connection instructions for setPath button
        :return: None
        """

        file_path, selected_filter = QtWidgets.QFileDialog.getOpenFileName(self,
                                                                           "Select File",
                                                                           CONST.INPUTSPATH,
                                                                           CONST.EXCELDOC)
        if file_path:
            self.excelFilePath = file_path
            self.loadExcelLineEdit.setText(file_path)

    def setTempfilePath(self):
        """
        pyside2 connection instructions for setPath button
        :return: None
        """
        file_path, selected_filter = QtWidgets.QFileDialog.getOpenFileName(self,
                                                                           "Select File",
                                                                           CONST.TEMPLATES,
                                                                           CONST.WORDDOC
                                                                           )
        if file_path:
            self.tempFilePath = file_path
            self.loadTempLineEdit.setText(file_path)

    def doIt(self):
        af = AutoFill(inputExcelFile=self.excelFilePath,
                      templateWordFile=self.tempFilePath,
                      outputPath=CONST.OUTPUTSPATH)

        af.autoFill()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    dlg = MainDialog()
    dlg.show()
    sys.exit(app.exec_())
