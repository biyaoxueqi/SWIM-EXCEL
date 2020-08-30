import sys


import pandas as pd
from PySide2 import QtCore, QtWidgets
import createStatic
from createStatic import SWIM, SWIMExcel


class MyWidget(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("SWIM static")
        self.hello = ["Hallo Welt", "Hei maailma", "Hola Mundo", "中国人"]
        self.button1 = QtWidgets.QPushButton("Select SWIM issues excel")
        self.button2 = QtWidgets.QPushButton('Select ECU list excel')
        self.button3 = QtWidgets.QPushButton("click to create static files")
        self.text = QtWidgets.QLabel("Read ME:\nStep1: Select the SWIM issue excel;\nStep2: Select ECU list excel\nStep3:click to create SWIM static files")
        # self.text.setAlignment(QtCore.Qt.AlignLeft|QtCore.Qt.AlignBottom)
        self.text.setAlignment(QtCore.Qt.AlignCenter)

        self.layout = QtWidgets.QVBoxLayout()
        self.layout.addWidget(self.text)
        self.layout.addWidget(self.button1)
        self.layout.addWidget(self.button2)
        self.layout.addWidget(self.button3)
        self.setLayout(self.layout)

        self.button1.clicked.connect(self.selectTargetFile)
        self.button2.clicked.connect(self.selectECUlist)
        self.button3.clicked.connect(self.runStatic)
        # self.button3.changeEvent(self.creatStaticFile())

    def selectTargetFile(self):
        # path = QtWidgets.QFileDialog.getExistingDirectory(self, '选择文件', './')
        path1 = QtWidgets.QFileDialog.getOpenFileName(self, '选择文件', './')
        print(path1[0])
        self.targetFile = path1[0]
        # SWIMs = pd.read_excel(self.targetFile)
        # return self.targetFile
        self.text.setText("The following target file is selected:\n{0}".format(self.targetFile))


    def selectECUlist(self):
        # path = QtWidgets.QFileDialog.getExistingDirectory(self, '选择文件', './')
        path2 = QtWidgets.QFileDialog.getOpenFileName(self, '选择文件', './')
        print(path2[0])
        self.ECUlist = path2[0]
        # INFO= pd.read_excel(self.ECUlist)
        # return self.ECUlist
        self.text.setText("The following ECU list is selected:\n{0}".format(self.ECUlist))

    def runStatic(self):
         swim = SWIM(name="swim")
         swimexcel = SWIMExcel(self.targetFile,self.ECUlist)
         SWIMs= swimexcel.readSWIMs()
         INFO= swimexcel.readECUlist()
         createStatic.issueStatic = []
         createStatic.abnormalIssue = []
         createStatic.delayIssue = []
         swim.createIssuelist(SWIMs, INFO)
         swim.createStatic(INFO)
         data = pd.DataFrame.from_dict(createStatic.issueStatic, orient="columns")
         data1 = pd.DataFrame.from_dict(createStatic.abnormalIssue, orient="columns")
         data2 = pd.DataFrame.from_dict(createStatic.delayIssue, orient="columns")

         data.to_excel("Issue Static.xlsx")
         data1.to_excel("AbnormalIssue.xlsx")
         data2.to_excel("DelayIssue.xlsx")
         self.text.setText("The following excels has been generated:Issue Static,AbnormalIssue,DelayIssue")
if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    widget = MyWidget()
    widget.resize(400,300)
    widget.show()
    sys.exit(app.exec_())
