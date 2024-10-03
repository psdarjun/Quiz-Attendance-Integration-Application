from flask import Flask, request, render_template, redirect, url_for, send_file
import pandas as pd
import os
import re

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Helper function to extract enrollment number from email address
def extract_enrollment_from_email(email):
    if isinstance(email, str):  # Check if email is a string
        match = re.search(r'(\d{4}[A-Z]{2}\d{6})', email)
        if match:
            return match.group(1)
    return None

@app.route('/')
def index():
    return render_template('upload.html')

@app.route('/process', methods=['POST'])
def process_files():
    if 'csv_files' not in request.files or 'attendance_file' not in request.files:
        return redirect(request.url)
    
    # Get the uploaded files
    files = request.files.getlist('csv_files')
    attendance_file = request.files['attendance_file']

    # Load the Btech Attendance 2024.xlsx file, skipping the first 4 rows (to start at row 5)
    attendance_filepath = os.path.join(app.config['UPLOAD_FOLDER'], attendance_file.filename)
    attendance_file.save(attendance_filepath)
    
    # Skip first 4 rows and set correct header
    attendance_df = pd.read_excel(attendance_filepath, skiprows=4)
    
    # Ensure we are working with the correct Enrollment column (3rd column)
    enrollment_column = 'Enrollment'

    if enrollment_column not in attendance_df.columns:
        return "Error: Could not find the 'Enrollment' column in the attendance file.", 400

    # Initialize a dictionary to keep track of the quizzes present count
    quizzes_present = {}

    # Process each CSV file
    for file in files:
        if file.filename.endswith('.csv'):
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            df = pd.read_csv(filepath)

            for _, row in df.iterrows():
                email = row[4]  # Assuming 'Email address' is in the 5th column
                state = row[5]  # Assuming 'State' is in the 6th column
                enrollment_number = extract_enrollment_from_email(email)

                if enrollment_number:
                    if enrollment_number not in quizzes_present:
                        quizzes_present[enrollment_number] = 0

                    # Mark as present if 'State' is not 'In progress'
                    if state != 'In progress':
                        quizzes_present[enrollment_number] += 1

    # Add the 'Quizzes Present' column to the attendance file using the 'Enrollment' column
    attendance_df['Quizzes Present'] = attendance_df[enrollment_column].apply(
        lambda x: quizzes_present.get(x, 0)
    )

    # Ensure the layout is preserved
    # Find the index of the 'Total Percentage' column
    total_percentage_index = attendance_df.columns.get_loc('Total Percentage')

    # Insert the 'Quizzes Present' column right after the 'Total Percentage' column
    attendance_df.insert(total_percentage_index + 1, 'Quizzes Present', attendance_df.pop('Quizzes Present'))

    # Save the updated attendance file
    updated_attendance_file = os.path.join(app.config['UPLOAD_FOLDER'], 'Updated_Btech_Attendance_2024.xlsx')
    attendance_df.to_excel(updated_attendance_file, index=False)

    return send_file(updated_attendance_file, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
