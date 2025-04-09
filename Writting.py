import xlwt
from PyQt6.QtWidgets import QFileDialog
from LeanerList import LeanerList
from learner import Leaner

def reset_workbook():
    return xlwt.Workbook()

# Create a bold style
bold_style = xlwt.XFStyle()
font = xlwt.Font()
font.bold = True
bold_style.font = font
style = xlwt.XFStyle()
style.alignment.horz = xlwt.Alignment.HORZ_LEFT
style.alignment.vert = xlwt.Alignment.VERT_TOP

combined_style = xlwt.XFStyle()
combined_style.font = bold_style.font  # Assigning only the font from bold_style
combined_style.alignment = style.alignment  # Assigning the alignment from style

class Pass_Fail:
    def __init__(self, fileName: str, minLevel: int, leaners: LeanerList, workbook):
        self.fileName = fileName
        self.workBook = workbook
        self.sheetPass = self.workBook.add_sheet('Passed Learners')
        self.sheetFail = self.workBook.add_sheet('Fail Learners')
        pass_row = 6
        fail_row = 6

        self.sheetPass.write(0, 0, leaners.name,combined_style)
        self.sheetPass.row(0).height_mismatch = True  # Allow custom height
        self.sheetPass.row(0).height = 255*5  # Set the height to 24 points (480/20)
        self.sheetPass.write(1, 0, 'Passed Learners', bold_style)
        self.sheetPass.write(2, 0, 'Number Of Learners', bold_style)
        self.sheetPass.write(3, 0, 'Weighting', bold_style)
        self.sheetPass.write(4, 0, 'Total', bold_style)
        self.sheetPass.write(5, 0, 'Learner', bold_style)
        self.sheetPass.write(5, len(leaners.leaners[0].marks) + 1, 'Report', bold_style)
        self.sheetPass.write(5, len(leaners.leaners[0].marks) + 2, 'Level', bold_style)

        self.sheetFail.write(0, 0, leaners.name, combined_style)
        self.sheetFail.row(0).height_mismatch = True  # Allow custom height
        self.sheetFail.row(0).height = 255*5  # Set the height to 24 points (480/20)
        self.sheetFail.write(1, 0, 'Failed Learners', bold_style)
        self.sheetFail.write(2, 0, 'Number Of Learners', bold_style)
        self.sheetFail.write(3, 0, 'Weighting', bold_style)
        self.sheetFail.write(4, 0, 'Total', bold_style)
        self.sheetFail.write(5, 0, 'Learner', bold_style)
        self.sheetFail.write(5, len(leaners.leaners[0].marks) + 1, 'Report Mark', bold_style)
        self.sheetFail.write(5, len(leaners.leaners[0].marks) + 2, 'Level', bold_style)

        if leaners.leaners:
            for x in range(len(leaners.leaners[0].marks)):
                self.sheetPass.col(x).width = 256 * 20
                self.sheetPass.write(3, x + 1, str(leaners.leaners[0].weights[x]), bold_style)
                self.sheetPass.write(4, x + 1, str(leaners.leaners[0].total[x]), bold_style)
                self.sheetPass.write(5, x + 1, f'T{x + 1}', bold_style)

                self.sheetFail.write(3, x + 1, str(leaners.leaners[0].weights[x]), bold_style)
                self.sheetFail.write(4, x + 1, str(leaners.leaners[0].total[x]), bold_style)
                self.sheetFail.write(5, x + 1, f'T{x + 1}', bold_style)

        for learner in leaners.leaners:
            if learner.achieved(minLevel):
                writeLearner(self, learner, self.sheetPass, pass_row)
                pass_row += 1
            else:
                writeLearner(self, learner, self.sheetFail, fail_row)
                fail_row += 1

        self.sheetPass.col(0).width = 256 * 40
        self.sheetFail.col(0).width = 256 * 40
        self.workBook.save(self.fileName)

def writeLearner(self, Lean: Leaner, sheet, row: int):
    sheet.write(row, 0, f'{row - 5}. {Lean.name}')
    for col, mark in enumerate(Lean.marks):
        sheet.write(row, col + 1, str(mark))
    sheet.write(row, len(Lean.marks) + 1, str(Lean.reporMark()))
    sheet.write(row, len(Lean.marks) + 2, str(Lean.level()))

class Levels:
    def __init__(self, fileName: str, leaners: LeanerList, workbook):
        self.fileName = fileName
        self.workBook = workbook
        self.numLevel = [0, 0, 0, 0, 0, 0, 0]
        self.sheet = self.workBook.add_sheet('Levels Analysis')
        self.sheet.write(0, 0, leaners.name, combined_style)
        for i in range(7):
            self.sheet.write(1, i, f'Level {i+1}', bold_style)
            self.sheet.col(i).width = 256 * 20

        countLevels(self, leaners)

        for i in range(7):
            self.sheet.write(2, i, f"{self.numLevel[i]}")

        self.workBook.save(self.fileName)

def countLevels(self, leaners: LeanerList):
    for lean in leaners.leaners:
        lvl = lean.level()
        if 1 <= lvl <= 7:
            self.numLevel[lvl - 1] += 1

class Top:
    def __init__(self, fileName: str, leaners: LeanerList, top: int, sheetName: str, workbook):
        self.fileName = fileName
        self.workBook = workbook
        self.sheet = self.workBook.add_sheet(f'{sheetName} {top}')

        row = 6

        self.sheet.write(0, 0, leaners.name, combined_style)
        self.sheet.row(0).height_mismatch = True  # Allow custom height
        self.sheet.row(0).height = 255*5  # Set the height to 24 points (480/20)
        self.sheet.write(1, 0, f'{top} learners', bold_style)
        self.sheet.write(2, 0, f'out of {len(leaners)} learners', bold_style)
        self.sheet.write(3, 0, 'Weighting', bold_style)
        self.sheet.write(4, 0, 'Total', bold_style)
        self.sheet.write(5, 0, 'Learner', bold_style)
        self.sheet.write(5, len(leaners.leaners[0].marks) + 1, 'Report', bold_style)
        self.sheet.write(5, len(leaners.leaners[0].marks) + 2, 'Level', bold_style)

        if leaners.leaners:
            for x in range(len(leaners.leaners[0].marks)):
                self.sheet.col(x).width = 256 * 20
                self.sheet.write(3, x + 1, str(leaners.leaners[0].weights[x]), bold_style)
                self.sheet.write(4, x + 1, str(leaners.leaners[0].total[x]), bold_style)
                self.sheet.write(5, x + 1, f'T{x + 1}', bold_style)

        for x in range(top):
            writeLearner(self, leaners.leaners[x], self.sheet, row)
            row += 1

        self.workBook.save(self.fileName)

class Groups:
    def __init__(self, fileName: str, numGroups: int, leaners: LeanerList, groupType: str, workbook,name:str):
        self.fileName = fileName
        self.workBook = workbook
        self.sheets = []

        for x in range(numGroups):
            sheet = self.workBook.add_sheet(f'{name} Group {x + 1}')
            sheet.write(0, 0, leaners.name, combined_style)
            sheet.row(0).height_mismatch = True  # Allow custom height
            sheet.row(0).height = 255*5  # Set the height to 24 points (480/20)
            sheet.write(1, 0, f'Grouped by {groupType}', bold_style)
            sheet.write(2, 0, f'Group {x + 1}', bold_style)
            sheet.write(3, 0, 'Weighting', bold_style)
            sheet.write(4, 0, 'Total', bold_style)
            sheet.write(5, 0, 'Learner', bold_style)
            sheet.write(5, len(leaners.leaners[0].marks) + 1, 'Report', bold_style)
            sheet.write(5, len(leaners.leaners[0].marks) + 2, 'Level', bold_style)
            self.sheets.append(sheet)

        for i in range(numGroups):
            self.sheets[i].col(0).width = 256 * 40
            if leaners.leaners:
                for x in range(len(leaners.leaners[0].marks)):
                    self.sheets[i].col(x + 1).width = 256 * 20
                    self.sheets[i].write(3, x + 1, str(leaners.leaners[0].weights[x]), bold_style)
                    self.sheets[i].write(4, x + 1, str(leaners.leaners[0].total[x]), bold_style)
                    self.sheets[i].write(5, x + 1, f'T{x + 1}', bold_style)

        groupCut = int((len(leaners) / numGroups) + 1)
        row = 6
        sheetIndex = 0
        for x in range(len(leaners.leaners)):
            if x in {groupCut, groupCut * 2, groupCut * 3, groupCut * 4}:
                row -= groupCut
                sheetIndex += 1
            writeLearner(self, leaners.leaners[x], self.sheets[sheetIndex], row + x)

        self.workBook.save(self.fileName)
