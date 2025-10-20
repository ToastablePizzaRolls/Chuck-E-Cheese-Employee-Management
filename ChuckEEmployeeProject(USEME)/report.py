import pandas as pd

class Report:
    @staticmethod
    def generate_report(rows):
        # rows is a list of tuples from DB: (ID, EmpID, FName, LName, PayRate)
        data = {
            "ID": [],
            "EmpID": [],
            "First Name": [],
            "Last Name": [],
            "Pay Rate": []
        }
        for row in rows:
            data["ID"].append(row[0])
            data["EmpID"].append(row[1])
            data["First Name"].append(row[2])
            data["Last Name"].append(row[3])
            data["Pay Rate"].append(row[4])
        df = pd.DataFrame(data)
        return df
    
    @staticmethod
    def export_to_excel(df):
        try:
            df.to_excel("employee_report.xlsx", index=False)
            return True
        except Exception as e:
            print("Error exporting to Excel:", e)
            return False
