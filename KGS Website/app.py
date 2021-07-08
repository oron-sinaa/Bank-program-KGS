import os
import bank_program
from flask import Flask, flash, request, redirect, url_for, render_template, session, send_file
from flask_wtf import Form
from werkzeug.utils import secure_filename
import pandas as pd

app = Flask(__name__)

APP_ROUTE = os.path.dirname(os.path.abspath(__file__))

@app.route('/', methods = ["GET","POST"])
def upload_file():
    if request.method == 'POST':
        file_var = request.files['file']
        # If the user does not select a file, the browser submits an empty file without a filename.
        if file_var.filename == '':
            flash('No selected file')
            return render_template('kgs_home.html')
        if file_var:
            file_var = request.files['file']
            file_var.save(os.path.join(APP_ROUTE+"\\uploads_folder", file_var.filename))
            session['name'] = str(file_var.filename)
            return render_template('kgs_home.html', message="File uploaded!", val = file_var.filename)

    return render_template('kgs_home.html', message="Waiting to submit...") 
    name = file_var.filename

@app.route('/kgs_instructions')
def instructions_page():
    try:
        file_var = session['name']
        os.path.join(APP_ROUTE+"\\uploads_folder", file_var)
        return render_template('kgs_instructions.html', val = os.path.join(APP_ROUTE+"\\uploads_folder", file_var))
    except:
        return render_template('kgs_instructions.html')

@app.route('/remarks_page/')
def remarks_page():
    try:
        file_var = session['name']
        in_file_path = os.path.join(APP_ROUTE+"\\uploads_folder", file_var)
        out_file_path = os.path.join(APP_ROUTE+"\\output_folder")
        sheet = bank_program.import_file(in_file_path)
        cleaned_sheet = bank_program.data_clean(sheet[0])
        rem_df = bank_program.generate_remarks(sheet[1], cleaned_sheet)
        f_out = bank_program.export_remarks(rem_df, str(file_var), out_file_path)
        return send_file(f_out, as_attachment=False)
        #return render_template('kgs_report.html', msg="Remarks file generated!", out_file_path = out_file_path, f_out = f_out)
    except:
        return render_template('kgs_report.html', msg="An error occured while generating remarks file! Check general instructions page.")

@app.route('/totals_page')
def totals_page():
    try:
        file_var = session['name']
        in_file_path = os.path.join(APP_ROUTE+"\\uploads_folder", file_var)
        out_file_path = os.path.join(APP_ROUTE+"\\output_folder")
        sheet = bank_program.import_file(in_file_path)
        cleaned_sheet = bank_program.data_clean(sheet[0])
        rem_df = bank_program.generate_remarks(sheet[1], cleaned_sheet)
        f_out = bank_program.totals_sheet(rem_df, str(file_var), out_file_path)
        return send_file(f_out, as_attachment=False)
    #return render_template('kgs_report.html', msg="Totals file generated!", out_file_path = out_file_path)
    except:
        return render_template('kgs_report.html', msg="An error occured while generating totals file! Check general instructions page.")

@app.route('/total_remarks')
def totals_and_remarks():
    try:
        file_var = session['name']
        in_file_path = os.path.join(APP_ROUTE+"\\uploads_folder", file_var)
        out_file_path = os.path.join(APP_ROUTE+"\\output_folder")
        sheet = bank_program.import_file(in_file_path)
        cleaned_sheet = bank_program.data_clean(sheet[0])
        rem_df = bank_program.generate_remarks(sheet[1], cleaned_sheet)
        tot_df = bank_program.totals_sheet(rem_df, str(file_var), out_file_path)[1]
        out_name = str(file_var.split(".", 1)[0]) + '-rem_totals.xlsx'
        final_f = out_file_path + '\\' + out_name
        writer = pd.ExcelWriter(final_f, engine='xlsxwriter')
        rem_df.to_excel(writer, sheet_name='Remarks')
        tot_df.to_excel(writer, sheet_name='Totals')
        writer.save()
        return send_file(writer, as_attachment=False)
    #return render_template('kgs_report.html', msg="Totals file generated!", out_file_path = out_file_path)
    except:
        return render_template('kgs_report.html', msg="An error occured while generating totals file! Check general instructions page.")

if __name__ == '__main__':
    app.secret_key = 'kgsfintechconnectkey'
    app.run(debug = True)