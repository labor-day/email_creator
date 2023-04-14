from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
import pandas as pd
import win32com.client as win32
import pythoncom
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = "uploads/"

@app.route('/', methods=['GET', 'POST'])
def generate_emails():
    if request.method == 'POST':

        # Get the uploaded Excel file
        excel_file = request.files['excel-file']
        df = pd.read_excel(excel_file)

        # Get the uploaded other files
        other_files = request.files.getlist('other-files[]')

        # Generate the emails
        for index, row in df.iterrows():
            outlook = win32.Dispatch('Outlook.Application', pythoncom.CoInitialize())
            mail = outlook.CreateItem(0x0)
            mail.To = row['Primary Contact']
            mail.Subject = row['Old Agency Name']
            mail.Body = row['PCF Market Name']

            for file in other_files:
              file.save(app.config['UPLOAD_FOLDER'] + secure_filename(file.filename))
              attachment = os.path.join(os.getcwd(), app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
              print(attachment)
              mail.Attachments.Add(attachment)

    return render_template('index.html')

if __name__ == '__main__':
    app.run()
