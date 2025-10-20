class Employee:
    def __init__(self, EmpID, FName, LName, PayRate):
        self.EmpID = EmpID
        self.FName = FName
        self.LName = LName
        self.PayRate = PayRate

    def update_info(self, FName=None, LName=None, PayRate=None):
        if FName:
            self.FName = FName
        if LName:
            self.LName = LName
        if PayRate:
            self.PayRate = PayRate

    def __str__(self):
        return f"EmpID: {self.EmpID}, Name: {self.FName} {self.LName}, PayRate: {self.PayRate}"
