"""Horizontal layout example."""

import sys
import pandas as pd
from PyQt5.QtCore import Qt
import PyQt5.QtWidgets as QtWidgets
from officeAutomation.autofillWordDoc.constants import constants as CONST
from officeAutomation.utils.utils import StringUtils as strUtils
from officeAutomation.autofillWordDoc.core.autoFill import AutoFill


class MainDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Auto Fill')
        self.setMinimumSize(600, 115)
        self.tempFilePath = ''
        self.excelFilePath = ''
        self.outputPath = CONST.OUTPUTSPATH
        self.createWidgets()
        self.createLayouts()
        self.createConnections()

    def createWidgets(self):
        """
        Create Widgets
        :return:
        """
        self.loadExcelLable = QtWidgets.QLabel('Excel Doc:')
        self.loadExcelLineEdit = QtWidgets.QLineEdit(CONST.INPUTSPATH)
        self.loadExcelPathBtn = QtWidgets.QPushButton('select')

        self.dividerLable1 = QtWidgets.QLabel('-' * 150)
        self.dividerLable1.setAlignment(Qt.AlignCenter)

        self.loadTempLable = QtWidgets.QLabel('Template Doc:')
        self.loadTempLineEdit = QtWidgets.QLineEdit(CONST.TEMPLATES)
        self.loadTempPathBtn = QtWidgets.QPushButton('select')

        self.dividerLable2 = QtWidgets.QLabel('-' * 150)
        self.dividerLable2.setAlignment(Qt.AlignCenter)

        self.loadOutputLable = QtWidgets.QLabel('Output Directory:')
        self.loadOutputLineEdit = QtWidgets.QLineEdit(CONST.OUTPUTSPATH)
        self.loadOutputPathBtn = QtWidgets.QPushButton('select')

        self.mappingQComboBox = QtWidgets.QComboBox()
        self.listBox = QtWidgets.QListWidget()
        self.listBox.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)

        self.okBtn = QtWidgets.QPushButton('OK')
        self.cancelBtn = QtWidgets.QPushButton('Cancel')

    def createLayouts(self):
        """
        Create Layouts
        :return:
        """
        self.layoutTop = QtWidgets.QVBoxLayout()
        self.layoutTopHBox1 = QtWidgets.QHBoxLayout()
        self.layoutTopHBox2 = QtWidgets.QHBoxLayout()
        self.layoutTopHBox3 = QtWidgets.QHBoxLayout()

        self.layoutMid = QtWidgets.QVBoxLayout()
        self.layoutBot = QtWidgets.QHBoxLayout()
        self.layoutMain = QtWidgets.QVBoxLayout()

        # TOP Layout
        self.layoutTop.addWidget(self.loadExcelLable)
        self.layoutTopHBox1.addWidget(self.loadExcelLineEdit)
        self.layoutTopHBox1.addWidget(self.loadExcelPathBtn)
        self.layoutTop.addLayout(self.layoutTopHBox1)

        self.layoutTop.addWidget(self.dividerLable1)

        self.layoutTop.addWidget(self.loadTempLable)
        self.layoutTopHBox2.addWidget(self.loadTempLineEdit)
        self.layoutTopHBox2.addWidget(self.loadTempPathBtn)
        self.layoutTop.addLayout(self.layoutTopHBox2)

        self.layoutTop.addWidget(self.dividerLable2)

        self.layoutTop.addWidget(self.loadOutputLable)
        self.layoutTopHBox3.addWidget(self.loadOutputLineEdit)
        self.layoutTopHBox3.addWidget(self.loadOutputPathBtn)
        self.layoutTop.addLayout(self.layoutTopHBox3)

        # add layout to main TOP layout
        self.layoutTop.addLayout(self.layoutTopHBox1)
        self.layoutTop.addLayout(self.layoutTopHBox2)
        self.layoutTop.addLayout(self.layoutTopHBox3)

        # MIDDLE layout
        self.layoutMid.addWidget(self.listBox)

        # BOTTOM layout
        self.layoutBot.addWidget(self.okBtn)
        self.layoutBot.addWidget(self.cancelBtn)

        # add to main layout
        self.layoutMain.addLayout(self.layoutTop)
        self.layoutMain.addLayout(self.layoutMid)
        self.layoutMain.addLayout(self.layoutBot)
        self.setLayout(self.layoutMain)

    def createConnections(self):
        """
        Create connection between widget and functions
        :return:
        """
        self.loadTempPathBtn.clicked.connect(self.setTempfilePath)
        self.loadExcelPathBtn.clicked.connect(self.setExcelfilePath)
        self.loadOutputPathBtn.clicked.connect(self.setOutputPath)
        self.okBtn.clicked.connect(self.doIt)
        self.cancelBtn.clicked.connect(self.close)

    def addToListWidget(self):
        """
        Read the input data and populate the listWidget
        :return: None
        """
        if self.excelFilePath:
            data = pd.read_excel(self.excelFilePath)
            pdData = pd.DataFrame(data)
            pdDataDict = pdData.to_dict()

            for key, value in pdDataDict['id'].items():
                self.listBox.addItem(value)

    def setExcelfilePath(self):
        """
        pyside2 connection instructions for setPath button
        :return: None
        """

        file_path, selected_filter = QtWidgets.QFileDialog.getOpenFileName(self,
                                                                           caption="Select File",
                                                                           directory=CONST.INPUTSPATH,
                                                                           filter=CONST.EXCELDOC)
        if file_path:
            self.excelFilePath = file_path
            self.loadExcelLineEdit.setText(file_path)
            self.addToListWidget()

    def setTempfilePath(self):
        """
        pyside2 connection instructions for setPath button
        :return: None
        """
        file_path, selected_filter = QtWidgets.QFileDialog.getOpenFileName(self,
                                                                           caption="Select File",
                                                                           directory=CONST.TEMPLATES,
                                                                           filter=CONST.WORDDOC
                                                                           )
        if file_path:
            self.tempFilePath = file_path
            self.loadTempLineEdit.setText(file_path)

    def setOutputPath(self):
        """
        pyside2 connection instructions for set output path button
        :return: None
        """
        file_path = QtWidgets.QFileDialog.getExistingDirectory(self,
                                                               caption="Select File",
                                                               directory=CONST.OUTPUTSPATH,
                                                               )
        if file_path:
            self.outputPath = file_path
            self.loadOutputLineEdit.setText(file_path)

    def createContextList(self, greenList):
        """
        Process pandas dataframe data to create a list of dictionaries for context
        :param greenList: list
        :return: list of dictionaries [{}, {}, {}]
        """
        excelFilePath = self.excelFilePath
        data = pd.read_excel(excelFilePath)
        pdData = pd.DataFrame(data)
        pdDataList = list(pdData.values.tolist())

        blacklist = ('Unnamed')

        if greenList:
            # if there are selected items in listWidget, create context of selected ids
            columnNames = list(pdData.columns.values)
            greenListItems = list()
            contextList = list()

            for i in pdDataList:
                if i[0] in greenList:
                    greenListItems.append(i)

            for itemList in greenListItems:
                contextDict = dict()
                keyValuePair = zip(columnNames, itemList)
                for key, val in keyValuePair:
                    if key in blacklist:
                        continue
                    newKeyName = strUtils().replaceSpace(inputStr=key)
                    contextDict.update({newKeyName: val})
                contextList.append(contextDict)
            return contextList
        else:
            # Create a context for all entries in input data
            numIDX = len(pdData.index)
            contextList = list()

            for idx in range(numIDX):
                contextDict = dict()
                for columName, row in pdData.iteritems():
                    if columName in blacklist:
                        continue
                    newColumName = strUtils().replaceSpace(inputStr=columName)
                    contextDict.update({newColumName: row[idx]})
                contextList.append(contextDict)
            return contextList

    def popup(self):
        """
        Create popup
        :return:
        """
        msg = QtWidgets.QMessageBox()
        msg.setWindowTitle("Done!")
        msg.setText("Task Finished!")
        msg.exec_()

    def resolve(self):
        """
        ensure that the necessary paths are available
        :return:
        """
        if not self.excelFilePath:
            print('Please set Excel file path!')
            self.setExcelfilePath()

        if not self.tempFilePath:
            print('Please set Word file path!')
            self.setTempfilePath()

    def doIt(self):
        """
        Run program
        :return: None
        """
        self.resolve()

        selItems = self.listBox.selectedItems()
        greenListItems = [selItems[idx].text() for idx in range(len(selItems))]
        contextList = self.createContextList(greenList=greenListItems)

        af = AutoFill(templateWordFile=self.tempFilePath,
                      outputPath=self.outputPath,
                      contextList=contextList)
        af.autoFill()
        self.popup()
