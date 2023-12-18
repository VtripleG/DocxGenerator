#!/usr/bin/env python3
import sys
import parser
from docx import Document
from PySide6.QtWidgets import QApplication, QWidget, QPushButton, QListWidget, QHBoxLayout, QVBoxLayout, QFileDialog, QLineEdit, QSizePolicy



class MainWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.dialog = QFileDialog()
        path = self.dialog.getOpenFileName()[0]
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
        self.fileData = parser.XmlToDict(path)
        self.discList = parser.GetDisciplineList(self.fileData)
        for key in self.discList.keys():
            self.leftListWidget.addItem(self.discList[key])
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

    def DoubleClickedOnLeftWidget(self):
        self.rightListWidget.addItem(self.leftListWidget.itemFromIndex(self.leftListWidget.currentIndex()).text())
        self.leftListWidget.takeItem(self.leftListWidget.currentRow())
    def DoubleClickedOnRightWidget(self):
        self.leftListWidget.addItem(self.rightListWidget.itemFromIndex(self.rightListWidget.currentIndex()).text())
        self.rightListWidget.takeItem(self.rightListWidget.currentRow())
    def GenerateButtonClicked(self):
        filePath = QFileDialog.getExistingDirectory()+'/'
        print(filePath)
        for index in range(self.rightListWidget.count()):
            fullInf = parser.GetFullInf(self.rightListWidget.item(index).text(), parser.KeyFromVal(self.discList, self.rightListWidget.item(index).text()), self.fileData)
            doc = parser.ReadDocxTemplate('./examples/RPD.docx')
            # doc = parser.ReadDocxTemplate('./examples/RPD_backup.docx')
            doc = parser.GenerateDocxOch(fullInf, doc)
            parser.SaveDocx(doc, self.rightListWidget.item(index).text(), filePath)
        self.rightListWidget.clear()

    def SearchButtonClicked(self):
        self.leftListWidget.clear()
        if self.searchLine.text() == '':
            for key in self.discList.keys():
                self.leftListWidget.addItem(self.discList[key])
            return
        for key in self.discList.keys():
            if self.searchLine.text().lower() in str(self.discList[key]).lower():
                self.leftListWidget.addItem(self.discList[key])

if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
