import learner
import xlrd
class LeanerList:
    def __init__(self, name="", leaners=None):
        self.name = name
        self.indexAll=name.find('All')
        self.leaners = leaners if leaners is not None else []
    def at(self, num:int):
        return self.leaners[num]
    def highest(self):
        if not self.leaners:
            return None
        ref = self.leaners[0]
        for lean in self.leaners[1:]:
            if ref.reportMark(self) < lean.reportMark(self):
                ref = lean
        return ref
    def lowest(self):
        if not self.leaners:
            return None
        ref = self.leaners[0]
        for lean in self.leaners[1:]:
            if ref.reportMark(self) > lean.reportMark(self):
                ref = lean
        return ref
    def at(self,index=0):
        return self.leaners[index]
    def addLeaner(self, lean:learner):
        for index in self.leaners:
            if index.name==lean.name:
                return
        self.leaners.append(lean)
        self.name = self.name[:self.indexAll]+f'{len(self.leaners):03}'+self.name[self.indexAll+3:]
    def setName(self,name:str):
        self.name=name
    def __str__(self):
        answer=self.name+'\n'
        for l in self.leaners:
            answer+=f'{l}\n'
        return answer
    def __len__(self):
        return len(self.leaners)
    #alphabetical check
    def sortAlphabetically(self):
        self.leaners.sort(key=lambda x: x.name)

    # Sort by average mark check
    def sortByAverage(self):
        self.leaners.sort(key=lambda x: x.average(), reverse=True)
    # Sort by report mark check
    def sortByReportMark(self):
        self.leaners.sort(key=lambda x: x.reporMark(), reverse=True)
    def sortByReportMarkDescending(self):
        self.leaners.sort(key=lambda x: x.reporMark(), reverse=False)
    def writeToFile(self,fileName:str):
        pass
    