import LeanerList
import learner
import xlrd
class RecordingSheet:
    def __init__(self,filename=""):
        self.file = xlrd.open_workbook(filename)
        self.worksheet=self.file.sheet_by_index(0)
        self.leaners=LeanerList.LeanerList(self.text(0,0))
    def searchHeadingRow(self):
        for x in range(self.worksheet.nrows):
            if(self.text(x,0)=='No'):
                return x
        print("Cannot Find The heading Row")
        return -1
    def searchTaskColomns(self):
        answer=[]
        row=self.searchHeadingRow()
        if(row==-1):
            return -1
        for col in range(self.worksheet.ncols):
            if self.text(row, col) in {'T1', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'T8'}:
                answer.append(col)
        return answer
    def searchWeightsRow(self):
        return 3

    def searchLeanerColomn(self):
        row=self.searchHeadingRow()
        for x in range(self.worksheet.ncols):
            if(self.text(row,x)=='Learner'):
                return x
        print('Cannot find the leaner colomn')
        return -1
    def searchLeanerRowBegin(self):
        row=self.searchHeadingRow()
        return row+1
    def genderColomn(self):
        row=self.searchHeadingRow()
        for x in range(self.worksheet.ncols):
            if(self.text(row,x)=='Gender'):
                return x
        return -1
    def isName(self, row:int, col:int):
        name=self.text(row,col)
        names=name.split(',')
        if(len(names)==2):
            return True
        else:
            return False
    def text(self, row:int, col:int):
        return self.worksheet.cell_value(row,col)
    def value(self, row: int, col: int) -> float:
        try:
            return float(self.worksheet.cell_value(row, col))
        except ValueError:
            return 0.0
        except TypeError:
            return 0.0
    def rowCount(self):
        return self.worksheet.nrows
    def colomnCount(self):
        return self.worksheet.ncols
    def addLeaners(self):
        leanerRow=self.searchLeanerRowBegin()
        learnerCol=self.searchLeanerColomn()
        taskColomn=self.searchTaskColomns()
        weightRow=self.searchWeightsRow()
        
        for row in range(leanerRow,self.rowCount()):
            if(self.isName(row,learnerCol)):
                newL=learner.Leaner(self.text(row,learnerCol))
                for mark in taskColomn:
                    newL.addMark(self.value(row,mark),self.value(weightRow+1,mark),self.value(weightRow,mark))
                self.leaners.addLeaner(newL)
