from KGSweb import app
import os
from KGSweb import bank_program
from flask import Flask, flash, request, redirect, url_for, render_template, session, send_file
from flask_wtf import Form
from werkzeug.utils import secure_filename
import pandas as pd

APP_ROUTE = os.path.dirname(os.path.abspath(__file__))
app.secret_key = 'kgsfintechconnectkey'

@app.route('/', methods = ["GET","POST"])
def upload_file():
    if request.method == 'POST':
        file_var = request.files['file']
        rel_party_var = request.form['rel_party_htm']
        # If the user does not select a file, the browser submits an empty file without a filename.
        if file_var.filename == '':
            flash('No selected file')
            return render_template('kgs_home.html')
        if file_var:
            file_var = request.files['file']
            file_var.save(os.path.join(APP_ROUTE+"\\uploads_folder", file_var.filename))
            session['name'] = str(file_var.filename)
            session['rel_party'] = str(rel_party_var)
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
        return render_template('kgs_report.html', msg="An error occured while generating remarks file! Check general instructions or ")

@app.route('/totals_page')
def totals_page():
    try:
        file_var = session['name']
        in_file_path = os.path.join(APP_ROUTE+"\\uploads_folder", file_var)
        out_file_path = os.path.join(APP_ROUTE+"\\output_folder")
        sheet = bank_program.import_file(in_file_path)
        cleaned_sheet = bank_program.data_clean(sheet[0])
        rem_df = bank_program.generate_remarks(sheet[1], cleaned_sheet)
        f_out = bank_program.totals_sheet(rem_df, str(file_var), out_file_path)[0]
        return send_file(f_out, as_attachment=False)
    #return render_template('kgs_report.html', msg="Totals file generated!", out_file_path = out_file_path)
    except:
        return render_template('kgs_report.html', msg="An error occured while generating totals file! Check general instructions or ")

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
        return render_template('kgs_report.html', msg="An error occured while generating totals file! Check general instructions page or ")

@app.route('/related_party')
def related_party():
    try:
        file_var = session['name']
        rel_party_var = session['rel_party']
        in_file_path = os.path.join(APP_ROUTE+"\\uploads_folder", file_var)
        out_file_path = os.path.join(APP_ROUTE+"\\output_folder")
        sheet = bank_program.import_file(in_file_path)
        cleaned_sheet = bank_program.data_clean(sheet[0])
        rem_df = bank_program.generate_remarks(sheet[1], cleaned_sheet)
        out_name = str(file_var.split(".", 1)[0]) + '-related_party.xlsx'
        final_f = out_file_path + '\\' + out_name
        related_df = bank_program.related_party(rel_party_var.split(','),rem_df)
        writer = pd.ExcelWriter(final_f, engine='xlsxwriter')
        related_df.to_excel(writer, sheet_name='Related Party')
        writer.save()
        return send_file(writer, as_attachment=False)
    except:
        return render_template('kgs_home.html', msg="An error occured while generating related party file! Check general instructions page")

@app.route('/add_features', methods = ["GET","POST"])
def add_features():
    if request.method == 'POST':
        file_list = request.files.getlist("multi_files")
        rel_party_var = request.form['file_list_names']
        # If the user does not select a file, the browser submits an empty file without a filename.
        for files in file_list:
            if files.filename == '':
                flash('Empty file name selected')
                return render_template('kgs_addfeatures.html')
        if file_list:
            name_list = []
            file_list = request.files.getlist("file_list_names")
            for files in file_list:
                files.save(os.path.join(APP_ROUTE+"\\uploads_folder", files.filename))
                name_list.append(files.filename)
            result = bank_program.search_terms(name_list, rel_party_var)
            out_name = str('variable_search_results.xlsx')
            final_f = out_file_path + '\\' + out_name
            writer = pd.ExcelWriter(final_f, engine='xlsxwriter')
            result.to_excel(writer, sheet_name='Search Results')
            writer.save()
            return send_file(writer, as_attachment=False)
            #return render_template('kgs_addfeatures.html', message="File(s) uploaded!", val = name_list)

    return render_template('kgs_addfeatures.html', message="Waiting to submit...") 

if __name__ == '__main__':
    app.secret_key = 'kgsfintechconnectkey'
    app.run(debug = True)