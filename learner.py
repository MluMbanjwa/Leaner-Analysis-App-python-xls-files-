
class Leaner:
    def __init__(self, name="", marks=None, total=None,weights=None):
        self.name = name
        self.marks = marks if marks is not None else []
        self.weights = weights if weights is not None else []
        self.total= total if total is not None else []

    def weightedMarks(self):
        if not (len(self.marks) == len(self.total) == len(self.weights)):
            raise ValueError("All input lists must be of the same length.")
        answer = []
        for i, mark in enumerate(self.marks):
            answer.append((mark / self.total[i]) * self.weights[i])
        return answer

    def highest(self):
        return max(self.marks, default=0.0)

    def lowest(self):
        return min(self.marks, default=0.0)

    def markIndex(self, mark):
        for i, value in enumerate(self.marks):
            if mark == value:
                return i
        return -1

    def addMark(self, mark,total, weight=0):
        self.marks.append(mark)
        self.total.append(total)
        self.weights.append(weight)

    def reporMark(self):
        return round(sum(self.weightedMarks()))  # ✅ Corrected function call

    def __str__(self):
        return f"{self.name}: {self.marks} ={self.reporMark()}"
    def level(self):
        if(self.reporMark()<30):
            return 1
        elif(self.reporMark()>29 and self.reporMark()<40):
            return 2
        elif(self.reporMark()>39 and self.reporMark()<50):
            return 3
        elif(self.reporMark()>49 and self.reporMark()<60):
            return 4
        elif(self.reporMark()>59 and self.reporMark()<70):
            return 5
        elif(self.reporMark()>69 and self.reporMark()<80):
            return 6
        elif(self.reporMark()>79 and self.reporMark()<101):
            return 7
        else:
            return -1
    def average(self):
        return sum(self.marks)/len(self.marks)
    def achieved(self,level: int):
        if(self.level()>=level):
            return True
        else:
            return False
