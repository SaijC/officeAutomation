import sys
import PyQt5.QtWidgets as QtWidgets
from officeAutomation.autofillWordDoc.ui.ui import MainDialog

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    dlg = MainDialog()
    dlg.show()
    sys.exit(app.exec_())
