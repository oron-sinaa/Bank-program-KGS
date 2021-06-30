import os
import bank_program
from flask import Flask, flash, request, redirect, url_for, render_template
from werkzeug.utils import secure_filename

app = Flask(__name__)

@app.route('/', methods = ["GET","POST"])
def upload_file():
    if request.method == 'POST':
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        # If the user does not select a file, the browser submits an
        # empty file without a filename.
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if file:
            file = request.files['file']
            file.save(os.path.join("uploads_folder",file.filename))
            return render_template('kgs_home.html', message="File uploaded!", name = file.filename)

    return render_template('kgs_home.html', message="Waiting to submit...")

@app.route('/kgs_instructions')
def instructions_page():
    return render_template('kgs_instructions.html')

@app.route('/kgs_list_files')
def file_list():
    files_list = os.listdir("uploads_folder")
    return render_template('kgs_list_files.html', files_list = files_list)

# @app.route('/remarks_page/)
# def remarks_page():
#     bank_program.data_clean(import_file)
#     bank_program.export_remarks(import_file, cleaned_sheet, f_name)

@app.route('/totals_page')
def totals_page():
    bank_program.totals_sheet(export_remarks, import_file)

if __name__ == '__main__':
    app.secret_key = 'super secret key'
    app.run(debug = True)