import pandas as pd
import os

class Customer:
    def __init__(self, reference, name, phone_number, initial_problem):
        self.reference = reference
        self.name = name
        self.phone_number = phone_number
        self.problems = []
        self.add_problem(initial_problem)

    def add_problem(self, problem):
        self.problems.append(problem)

    def solve_problem(self):
        self.problems.clear()
        return f"Problems of {self.name} are solved."

    def view_problems(self):
        if self.problems == []:
            return f"No problems for {self.name}."
        else:
            return f"The phone has following problems left: {self.problems}."

customers = []

def create_customer():
    reference = input("Enter customer reference: ")
    name = input("Enter customer name: ")
    phone_number = input("Enter customer phone number: ")
    initial_problem = input("Enter initial problem: ")

    customer = Customer(reference, name, phone_number, initial_problem)
    customers.append(customer)
    print(f"Customer {customer.name} has been added.")
    run_wsl()

def view_customers():
    if not customers:
        print("No customers found.")
    else:
        print("List of Customers:")
        for customer in customers:
            print(customer.name)
            print(customer.phone_number)
            print(customer.view_problems())

def log_customers_to_excel():
    if not customers:
        print("No customers to log.")
    else:
        data = {
            'Reference': [customer.reference for customer in customers],
            'Name': [customer.name for customer in customers],
            'Phone Number': [customer.phone_number for customer in customers],
            'Problems': [', '.join(customer.problems) for customer in customers]
        }
        excel_file = 'customer_log.xlsx'
        if os.path.exists(excel_file):
            # Load the existing Excel file into a DataFrame
            existing_df = pd.read_excel(excel_file)

            # Append the new data to the existing data
            updated_df = pd.concat([existing_df, pd.DataFrame(data)])

            # Write the updated DataFrame back to the Excel file
            updated_df.to_excel(excel_file, index=False)
        else:
            # If the file doesn't exist, create it with the new data
            df = pd.DataFrame(data)
            df.to_excel(excel_file, index=False)

        print("Customer data logged to customer_log.xlsx")

def run_wsl():
    command = input("Type what to do: add / view / log / exit: ")
    if command == "add":
        create_customer()
    elif command == "view":
        view_customers()
    elif command == "log":
        log_customers_to_excel()
        run_wsl()
    elif command == "exit":
        return
    else:
        run_wsl()

run_wsl()
