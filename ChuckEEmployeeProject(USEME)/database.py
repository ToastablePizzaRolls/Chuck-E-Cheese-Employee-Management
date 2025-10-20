import pyodbc 
import os

class Database:
    def __init__(self, db_path):
        # Convert to absolute path for reliability
        self.db_path = os.path.abspath(db_path)
        self.conn = None
        self.connect()

    def connect(self):
        try:
            connection_string = (
                r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};'
                f'DBQ={self.db_path};'
            )
            self.conn = pyodbc.connect(connection_string)
            print(f"Database connected successfully to {self.db_path}")
        except Exception as e:
            print(f"Error connecting to database at {self.db_path}: {e}")
            self.conn = None

    def get_max_empid(self):
        if not self.conn:
            print("No database connection available in get_max_empid.")
            return 1000
        cursor = self.conn.cursor()
        try:
            cursor.execute("SELECT MAX(EmpID) FROM EmployeeTable")
            result = cursor.fetchone()[0]
            return result if result is not None else 1000
        except Exception as e:
            print("Error fetching max EmpID:", e)
            return 1000

    def add_employee(self, emp):
        if not self.conn:
            print("No database connection available in add_employee.")
            return False
        cursor = self.conn.cursor()
        try:
            cursor.execute(
                "INSERT INTO EmployeeTable (EmpID, FName, LName, PayRate) VALUES (?, ?, ?, ?)",
                (emp.EmpID, emp.FName, emp.LName, emp.PayRate)
            )
            self.conn.commit()
            return True
        except Exception as e:
            print("Error adding employee:", e)
            return False

    def update_employee(self, emp):
        if not self.conn:
            print("No database connection available in update_employee.")
            return False
        cursor = self.conn.cursor()
        try:
            cursor.execute(
                "UPDATE EmployeeTable SET FName=?, LName=?, PayRate=? WHERE EmpID=?",
                (emp.FName, emp.LName, emp.PayRate, emp.EmpID)
            )
            self.conn.commit()
            return True
        except Exception as e:
            print("Error updating employee:", e)
            return False

    def get_employee(self, emp_id):
        if not self.conn:
            print("No database connection available in get_employee.")
            return None
        cursor = self.conn.cursor()
        try:
            cursor.execute("SELECT ID, EmpID, FName, LName, PayRate FROM EmployeeTable WHERE EmpID=?", (emp_id,))
            row = cursor.fetchone()
            if row:
                return row
            else:
                return None
        except Exception as e:
            print("Error querying employee:", e)
            return None

    def get_all_employees(self):
        if not self.conn:
            print("No database connection available in get_all_employees.")
            return []
        cursor = self.conn.cursor()
        try:
            cursor.execute("SELECT ID, EmpID, FName, LName, PayRate FROM EmployeeTable")
            return cursor.fetchall()
        except Exception as e:
            print("Error fetching employees:", e)
            return []
