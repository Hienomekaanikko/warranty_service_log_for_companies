from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import os

app = Flask(__name__)

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
      return f"The phone has following problems             left: {self.problems}."

customers = []

@app.route('/')
def index():
    return render_template('index.html', customers=customers)

@app.route('/add_customer', methods=['POST'])
def add_customer():
    reference = request.form['reference']
    name = request.form['name']
    phone_number = request.form['phone_number']
    initial_problem = request.form['initial_problem']

    customer = Customer(reference, name, phone_number, initial_problem)
    customers.append(customer)
    return redirect(url_for('index'))

@app.route('/log_to_excel')
def log_to_excel():
    if not customers:
        return "No customers to log."
    
    data = {
        'Reference': [customer.reference for customer in customers],
        'Name': [customer.name for customer in customers],
        'Phone Number': [customer.phone_number for customer in customers],
        'Problems': [', '.join(customer.problems) for customer in customers]
    }
    
    excel_file = 'customer_log.xlsx'
    
    if os.path.exists(excel_file):
        existing_df = pd.read_excel(excel_file)
        updated_df = pd.concat([existing_df, pd.DataFrame(data)])
        updated_df.to_excel(excel_file, index=False)
    else:
        df = pd.DataFrame(data)
        df.to_excel(excel_file, index=False)

    return "Customer data logged to customer_log.xlsx"

if __name__ == '__main__':
    app.run()
