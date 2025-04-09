from PyQt6.QtWidgets import QHBoxLayout,QGroupBox,QLabel,QRadioButton,QPushButton,QSpinBox,QFormLayout,QWidget,QApplication,QFileDialog
from PyQt6.QtCore import QStringListModel
from RecordingSheet import RecordingSheet
import Writting
import xlwt
from Writting import Pass_Fail
from Writting import Levels
from Writting import Top
from Writting import Groups
class Window(QWidget):
    def __init__(self):
        super().__init__()
        with open("packages/style.css", "r") as f:
            self.setStyleSheet(f.read())
        self.write=None
        self.SavefileName='emptyName.xls'
        self.labelstatus=QLabel('No recording Sheet(SASAMS)',self)
        self.recoding = None
        self.workbook=None
        self.numSheets=0
        self.btnOption=[QRadioButton('levels',self),QRadioButton('Top',self),QRadioButton('Bottom',self),QRadioButton('Pass/Fail',self),QRadioButton('Alphabetical',self),QRadioButton('Average',self)]
        self.groupBox=QGroupBox('Group By: ',self)
        tempHBlayout=QHBoxLayout()
        for item in self.btnOption:
            tempHBlayout.addWidget(item)
            item.clicked.connect(self.setLimits)
        self.groupBox.setLayout(tempHBlayout)  
        self.btnButtons=[QPushButton('Close',self),QPushButton('Group Leaners',self),QPushButton('Select recording Sheet',self),]
        self.btnButtons[0].clicked.connect(self.close_clicked)
        self.btnButtons[1].clicked.connect(self.group_clicked)
        self.btnButtons[2].clicked.connect(self.upload_clicked)
        self.groupLabel=QLabel('',self)
        tempHBbtnLayout=QHBoxLayout()
        for item in self.btnButtons:
            tempHBbtnLayout.addWidget(item)
            
        self.numGroups=QSpinBox(self)
        self.winLayout=QFormLayout(self)
        
        self.winLayout.addRow(self.groupBox)
        self.winLayout.addRow(self.groupLabel,self.numGroups)
        self.winLayout.addRow(QLabel('Status: ',self),self.labelstatus)
        self.winLayout.addRow(tempHBbtnLayout)
        
        self.setLayout(self.winLayout)
        self.setWindowTitle('Leaner Analysis App')
        self.show()
    def setLimits(self):
        if(self.btnOption[0].isChecked()):
            self.numGroups.setEnabled(False)
            self.numSheets=7
        elif(self.btnOption[1].isChecked()):
            self.numGroups.setMinimum(1)
            self.numGroups.setMaximum(1000)
            self.groupLabel.setText("Top : ")
            self.numGroups.setEnabled(True)
            self.numSheets=1
        elif(self.btnOption[2].isChecked()):
            self.numGroups.setMinimum(1)
            self.numGroups.setMaximum(1000)
            self.groupLabel.setText("Bottom: ")
            self.numSheets=1
            self.numGroups.setEnabled(True)
        elif(self.btnOption[3].isChecked()):
            self.numGroups.setMinimum(1)
            self.numGroups.setMaximum(7)
            self.groupLabel.setText("Minimum Pass level: ")
            self.numSheets=2
        elif(self.btnOption[4].isChecked()):
            self.numGroups.setMinimum(1)
            self.numGroups.setMaximum(1000)
            self.groupLabel.setText("Number of Groups: ")
            self.numGroups.setEnabled(True)
            self.numSheets=self.numGroups.value()
        elif(self.btnOption[5].isChecked()):
            self.numGroups.setMinimum(1)
            self.numGroups.setMaximum(1000)
            self.numGroups.setEnabled(True)
            self.numSheets=self.numGroups.value()
    def close_clicked(self):
        self.close()
    def group_clicked(self):
        task=''
        if self.btnOption[0].isChecked():
            self.recoding.leaners.sortAlphabetically()
            self.write=Levels(self.SavefileName,self.recoding.leaners,self.workbook)
            task='Levels analysis'
        elif self.btnOption[1].isChecked():
            self.recoding.leaners.sortByReportMark()
            self.write=Top(self.SavefileName,self.recoding.leaners,self.numGroups.value(),'Top',self.workbook)
            task=f'Top {self.numGroups.value()}'
        elif self.btnOption[2].isChecked():
            self.recoding.leaners.sortByReportMarkDescending()
            self.write=Top(self.SavefileName,self.recoding.leaners,self.numGroups.value(),'Bottom',self.workbook)
            task=f'Bottom {self.numGroups.value()}'
        elif self.btnOption[3].isChecked():
            self.recoding.leaners.sortAlphabetically()
            self.write = Pass_Fail(self.SavefileName,self.numGroups.value(), self.recoding.leaners,self.workbook)
            task='Pass and Fail Groups'
        elif self.btnOption[4].isChecked():
            self.recoding.leaners.sortAlphabetically()
            self.write=Groups(self.SavefileName,self.numGroups.value(),self.recoding.leaners,'Alphabet',self.workbook,'Alphabetical')
            task=f'Groups alphabetically= {self.numGroups.value()} Group(s)'
        elif self.btnOption[5].isChecked():
            self.recoding.leaners.sortByAverage()
            self.write=Groups(self.SavefileName,self.numGroups.value(),self.recoding.leaners,'Average',self.workbook,'Average')
            task=f'Groups by Average= {self.numGroups.value()} Group(s)'
        self.labelstatus.setText(self.labelstatus.text()+'\n' +f"Added: {task}")

    def upload_clicked(self):
        self.reset_window()
        
        fileName, _ = QFileDialog.getOpenFileName(self, "Select an SASAASM recording sheet", "", "Excel Files (*.xls *.xlsx)")
        if fileName:
            try:
                # Initialize the RecordingSheet object with the selected file
                self.recoding = RecordingSheet(fileName)
                # Assuming addLeaners() processes the file and populates the leaners list
                self.recoding.addLeaners()
                # Output the number of learners in the console
                print(f'Number of Leaners = {len(self.recoding.leaners)}')
            except Exception as e:
                print(f"Error processing the file: {e}")
        else:
            print("No file selected")
        self.labelstatus.setText(f"{self.recoding.leaners.name}")
        self.SavefileName,_=QFileDialog.getSaveFileName(self, 'Save As', '', 'Excel Files (*.xls);;All Files (*)')

        
    def reset_window(self):
        self.workbook=Writting.reset_workbook()
        self.labelstatus.setText("No recording Sheet(SASAMS)")
        self.groupLabel.setText("")
        self.numGroups.setValue(0)
        self.numGroups.setEnabled(False)
        self.SavefileName = 'emptyName.xls'
        self.write = None
        self.recoding = None
        for btn in self.btnOption:
            btn.setAutoExclusive(False)
            btn.setChecked(False)
            btn.setAutoExclusive(True)
        
       
