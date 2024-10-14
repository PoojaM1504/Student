from flask import Flask, render_template, request, redirect
import pandas as pd
import os

app = Flask(__name__)

# Path to the Excel file
EXCEL_FILE = 'student_info.xlsx'

# Check if the Excel file exists, if not create one with headers
if not os.path.exists(EXCEL_FILE):
    df = pd.DataFrame(columns=['Full Name', 'Roll Number', 'Date of Birth', 'Gender', 'Address', 'Email', 'Phone', 'Course', 'Year', 'Hobbies'])
    df.to_excel(EXCEL_FILE, index=False)

@app.route('/')
def form():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    # Get form data
    name = request.form['name']
    roll_number = request.form['rollNumber']
    dob = request.form['dob']
    gender = request.form['gender']
    address = request.form['address']
    email = request.form['email']
    phone = request.form['phone']
    course = request.form['course']
    year = request.form['year']
    hobbies = request.form['hobbies']

    # Read existing Excel file
    df = pd.read_excel(EXCEL_FILE)

    # Append new data
    new_data = {
        'Full Name': name,
        'Roll Number': roll_number,
        'Date of Birth': dob,
        'Gender': gender,
        'Address': address,
        'Email': email,
        'Phone': phone,
        'Course': course,
        'Year': year,
        'Hobbies': hobbies
    }
    df = df._append(new_data, ignore_index=True)

    # Save back to Excel file
    df.to_excel(EXCEL_FILE, index=False)

    return redirect('/')

if __name__ == '__main__':
    app.run(debug=True)
