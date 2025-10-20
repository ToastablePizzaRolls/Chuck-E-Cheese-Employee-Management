import os
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter import filedialog
from database import Database
from employee import Employee
from report import Report

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Chuck E Cheese Employee Management")

        # Use absolute path for the database
        db_path = os.path.abspath(r"C:\Users\OneOf\OneDrive\Desktop\ChuckEEmployeeProject\ChuckEEmployeeProject\EmployeeSDEV120.accdb")
        self.db = Database(db_path)

        # Check if DB connection succeeded
        if not self.db.conn:
            messagebox.showerror("Database Error", f"Failed to connect to database at:\n{db_path}")
            self.root.destroy()
            return

        self.setup_ui()

    def setup_ui(self):
        # Labels and entry fields
        tk.Label(self.root, text="Employee ID:").grid(row=0, column=0, sticky=tk.W)
        self.empid_var = tk.StringVar()
        self.empid_entry = tk.Entry(self.root, textvariable=self.empid_var, state='readonly')
        self.empid_entry.grid(row=0, column=1)

        tk.Label(self.root, text="First Name:").grid(row=1, column=0, sticky=tk.W)
        self.fname_var = tk.StringVar()
        self.fname_entry = tk.Entry(self.root, textvariable=self.fname_var)
        self.fname_entry.grid(row=1, column=1)

        tk.Label(self.root, text="Last Name:").grid(row=2, column=0, sticky=tk.W)
        self.lname_var = tk.StringVar()
        self.lname_entry = tk.Entry(self.root, textvariable=self.lname_var)
        self.lname_entry.grid(row=2, column=1)

        tk.Label(self.root, text="Pay Rate:").grid(row=3, column=0, sticky=tk.W)
        self.payrate_var = tk.StringVar()
        self.payrate_entry = tk.Entry(self.root, textvariable=self.payrate_var)
        self.payrate_entry.grid(row=3, column=1)

        # Buttons
        tk.Button(self.root, text="Add Employee", command=self.add_employee).grid(row=4, column=0, pady=5)
        tk.Button(self.root, text="Query Employee", command=self.query_employee).grid(row=4, column=1, pady=5)
        tk.Button(self.root, text="Update Employee", command=self.update_employee).grid(row=5, column=0, pady=5)
        tk.Button(self.root, text="Show Report", command=self.show_report).grid(row=5, column=1, pady=5)
        tk.Button(self.root, text="Export Report to Excel", command=self.export_report).grid(row=6, column=0, columnspan=2, pady=5)

        # Report Treeview
        self.tree = ttk.Treeview(self.root, columns=("ID", "EmpID", "FName", "LName", "PayRate"), show='headings')
        for col in ("ID", "EmpID", "FName", "LName", "PayRate"):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        self.tree.grid(row=7, column=0, columnspan=2, pady=10)

        # Initialize EmpID field
        self.reset_empid()

    def reset_empid(self):
        max_empid = self.db.get_max_empid()
        self.empid_var.set(str(int(max_empid) + 1))

    def add_employee(self):
        try:
            emp = Employee(
                EmpID=int(self.empid_var.get()),
                FName=self.fname_var.get(),
                LName=self.lname_var.get(),
                PayRate=float(self.payrate_var.get())
            )
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid data.")
            return

        if self.db.add_employee(emp):
            messagebox.showinfo("Success", "Employee added successfully.")
            self.reset_empid()
            self.clear_entries()
            self.show_report()
        else:
            messagebox.showerror("Error", "Failed to add employee.")

    def query_employee(self):
        try:
            emp_id = int(self.empid_var.get())
        except ValueError:
            messagebox.showerror("Input Error", "Please enter a valid Employee ID.")
            return

        row = self.db.get_employee(emp_id)
        if row:
            _, empid, fname, lname, payrate = row
            self.empid_var.set(str(empid))
            self.fname_var.set(fname)
            self.lname_var.set(lname)
            self.payrate_var.set(str(payrate))
        else:
            messagebox.showinfo("Not Found", f"No employee found with EmpID {emp_id}.")

    def update_employee(self):
        try:
            emp = Employee(
                EmpID=int(self.empid_var.get()),
                FName=self.fname_var.get(),
                LName=self.lname_var.get(),
                PayRate=float(self.payrate_var.get())
            )
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid data.")
            return

        if self.db.update_employee(emp):
            messagebox.showinfo("Success", "Employee updated successfully.")
            self.clear_entries()
            self.show_report()
        else:
            messagebox.showerror("Error", "Failed to update employee.")

    def show_report(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        employees = self.db.get_all_employees()
        for emp in employees:
            self.tree.insert('', 'end', values=emp)

    def export_report(self):
        employees = self.db.get_all_employees()
        if not employees:
            messagebox.showinfo("No Data", "No employee data to export.")
            return

        report = Report()
        df = report.generate_report(employees)

    # Ask user where to save the Excel file
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save Employee Report As"
      )

        if not file_path:
            return  # user canceled the dialog

        try:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Export Successful", f"Report saved to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export report:\n{e}")

    def clear_entries(self):
        self.fname_var.set("")
        self.lname_var.set("")
        self.payrate_var.set("")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
