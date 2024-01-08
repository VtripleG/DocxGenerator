import sys
import parser
from docx import Document
from PySide6.QtWidgets import QApplication, QWidget, QPushButton, QListWidget, QHBoxLayout, QVBoxLayout, QFileDialog, \
    QLineEdit, QSizePolicy, QMessageBox
from PyQt5.QtCore import Qt



class MainWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.discListOch = dict()
        self.discListZaoch = dict()
        self.searchLine = QLineEdit()
        self.searchLine.setSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Minimum)
        self.searchButton = QPushButton()
        self.searchButton.setSizePolicy(QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Minimum)
        self.searchButton.setText('Search')
        self.searchButton.clicked.connect(self.SearchButtonClicked)
        self.searchLay = QHBoxLayout()
        self.searchLay.addWidget(self.searchLine)
        self.searchLay.addWidget(self.searchButton)
        self.leftLay = QVBoxLayout()
        self.leftLay.addLayout(self.searchLay)
        self.generateButton = QPushButton()
        self.generateButton.setText('Generate docx')
        self.setMinimumSize(640, 780)
        self.leftListWidget = QListWidget()
        self.rightListWidget = QListWidget()
        self.listLay = QHBoxLayout()
        self.leftLay.addWidget(self.leftListWidget)
        self.listLay.addLayout(self.leftLay)
        self.fileLay = QHBoxLayout()
        self.ochButton = QPushButton()
        self.ochButton.setText('Ochnoe')
        self.zaochButton = QPushButton()
        self.zaochButton.setText('Zaochoe')
        self.fileLay.addWidget(self.ochButton)
        self.fileLay.addWidget(self.zaochButton)
        self.rightLay = QVBoxLayout()
        self.rightLay.addLayout(self.fileLay)
        self.rightLay.addWidget(self.rightListWidget)
        self.listLay.addLayout(self.rightLay)
        self.mainLay = QVBoxLayout()
        self.mainLay.addLayout(self.listLay)
        self.mainLay.addWidget(self.generateButton)
        self.setLayout(self.mainLay)
        self.leftListWidget.doubleClicked.connect(self.DoubleClickedOnLeftWidget)
        self.rightListWidget.doubleClicked.connect(self.DoubleClickedOnRightWidget)
        self.generateButton.clicked.connect(self.GenerateButtonClicked)
        self.ochButton.clicked.connect(self.OchButtonClicked)
        self.zaochButton.clicked.connect(self.ZaochButtonClicked)

    def DoubleClickedOnLeftWidget(self):
        self.rightListWidget.addItem(self.leftListWidget.itemFromIndex(self.leftListWidget.currentIndex()).text())
        self.leftListWidget.takeItem(self.leftListWidget.currentRow())

    def DoubleClickedOnRightWidget(self):
        self.leftListWidget.addItem(self.rightListWidget.itemFromIndex(self.rightListWidget.currentIndex()).text())
        self.rightListWidget.takeItem(self.rightListWidget.currentRow())

    def GenerateButtonClicked(self):
        filePath = QFileDialog.getExistingDirectory()
        if filePath == '':
            return
        filePath += '/'
        self.setEnabled(False)
        # self.setCursor(Qt.WaitCursor)
        for index in range(self.rightListWidget.count()):
            try:
                doc = parser.ReadDocxTemplate('./examples/RPD.docx')
                if self.rightListWidget.item(index).text() in self.discListZaoch.values():
                    fullInfOch = parser.GetFullInf(self.rightListWidget.item(index).text(),
                                                   parser.KeyFromVal(self.discListOch,
                                                                     self.rightListWidget.item(index).text()),
                                                   self.fileDataOch)
                    fullInfZaoch = parser.GetFullInf(self.rightListWidget.item(index).text(),
                                                     parser.KeyFromVal(self.discListZaoch,
                                                                       self.rightListWidget.item(index).text()),
                                                     self.fileDataZaoch)
                    doc = parser.GenerateDocxOchZ(fullInfOch, fullInfZaoch, doc)
                else:
                    fullInf = parser.GetFullInf(self.rightListWidget.item(index).text(),
                                                parser.KeyFromVal(self.discListOch,
                                                                  self.rightListWidget.item(index).text()),
                                                self.fileDataOch)
                    doc = parser.GenerateDocxOch(fullInf, doc)
                parser.SaveDocx(doc, self.rightListWidget.item(index).text(), filePath)
            except Exception as e:
                print(e)
                QMessageBox.critical(self, 'Generate docx file ERROR',
                                     f"An ERROR occurred during file generation {self.rightListWidget.item(index).text()}")
        self.rightListWidget.clear()
        QMessageBox.information(self, 'Complite', 'Generate complite!')
        self.setEnabled(True)

    def SearchButtonClicked(self):
        self.leftListWidget.clear()
        if self.searchLine.text() == '':
            for key in self.discListOch.keys():
                self.leftListWidget.addItem(self.discListOch[key])
            return
        for key in self.discListOch.keys():
            if self.searchLine.text().lower() in str(self.discListOch[key]).lower():
                self.leftListWidget.addItem(self.discListOch[key])

    def OchButtonClicked(self):
        dialog = QFileDialog()
        path = dialog.getOpenFileName(filter="plx(*.plx)")[0]
        if path == '':
            return
        self.leftListWidget.clear()
        self.rightListWidget.clear()
        self.discListZaoch.clear()
        self.fileDataOch = parser.XmlToDict(path)
        self.discListOch = parser.GetDisciplineList(self.fileDataOch)
        for key in self.discListOch.keys():
            self.leftListWidget.addItem(self.discListOch[key])

    def ZaochButtonClicked(self):
        dialog = QFileDialog()
        path = dialog.getOpenFileName(filter="plx(*.plx)")[0]
        if path == '':
            return
        self.fileDataZaoch = parser.XmlToDict(path)
        self.discListZaoch = parser.GetDisciplineList(self.fileDataZaoch)


if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
