#!/usr/bin/env python3
import sys
import parser
from docx import Document
from PySide6.QtWidgets import QApplication, QWidget, QPushButton, QListWidget, QHBoxLayout, QVBoxLayout, QFileDialog



class MainWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.dialog = QFileDialog()
        path = self.dialog.getOpenFileName()[0]
        print(path.__str__())
        self.generateButton = QPushButton()
        self.generateButton.setText('Generate docx')
        self.setMinimumSize(640, 480)
        self.leftListWidget = QListWidget()
        self.rightListWidget = QListWidget()
        self.listLay = QHBoxLayout()
        self.listLay.addWidget(self.leftListWidget)
        self.listLay.addWidget(self.rightListWidget)
        self.fileData = parser.XmlToDict(path)
        self.discList = parser.GetDisciplineList(self.fileData)
        for key in self.discList.keys():
            self.leftListWidget.addItem(self.discList[key])
        self.leftListWidget.doubleClicked.connect(self.DoubleClickedOnLeftWidget)
        self.rightListWidget.doubleClicked.connect(self.DoubleClickedOnRightWidget)
        self.generateButton.clicked.connect(self.GenerateButtonClicked)
        self.mainLay = QVBoxLayout()
        self.mainLay.addLayout(self.listLay)
        self.mainLay.addWidget(self.generateButton)
        self.setLayout(self.mainLay)

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
            doc = parser.GenerateDocx(fullInf, doc)
            parser.SaveDocx(doc, self.rightListWidget.item(index).text(), filePath)
        self.rightListWidget.clear()

if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
