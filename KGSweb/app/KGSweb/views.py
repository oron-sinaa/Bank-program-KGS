"""
File:    views.py 

Author:   Mohd Aanis Noor (https://github.com/oron-sinaa) 
Company:  KG Somani & Co (https://www.kgsomani.com/) 
Date:     Started Jun 2021

Summary of File: 

This file contains flask python code that is the frontend for
bank_program.py which analyzes & generates reports from several type
provided bank document files. The code as in August 2021 is working
but still in development phase.
"""

from KGSweb import app
import os, glob
from KGSweb import bank_program
from KGSweb import BBrecon
from flask import Flask, request, redirect, url_for, render_template, session, send_file, send_from_directory, flash
from flask_wtf import Form
from werkzeug.utils import secure_filename
import pandas as pd

home_url = "http://127.0.0.1:5000/"
APP_ROUTE = os.path.dirname(os.path.abspath(__file__))
# app.secret_key = 'kgsfintechconnectkey'
XLSX_MIMETYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

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

@app.route('/kgs_instructions')
def instructions_page():
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
        out_name = str(str(file_var).split(".", 1)[0]) + '-remarks.xlsx'
        f_out = bank_program.export_remarks(rem_df, str(file_var), out_file_path)
        #return send_file(f_out, as_attachment=False, attachment_filename = out_name, mimetype = XLSX_MIMETYPE)
        return send_from_directory(out_file_path, out_name, as_attachment=True)
    # return render_template('kgs_home.html', msg="Remarks file generated!", out_file_path = out_file_path, f_out = f_out)
    except:
        return render_template('kgs_home.html', msg="An error occured while generating remarks file! Check general instructions to see specifications.")

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
        out_name = str(file_var.split(".", 1)[0]) + '-totals.xlsx'
        # return send_file(file_var, as_attachment=False, mimetype = XLSX_MIMETYPE)
        return send_from_directory(out_file_path, out_name, as_attachment=True)
    #return render_template('kgs_home.html', msg="Totals file generated!", out_file_path = out_file_path)
    except:
        return render_template('kgs_home.html', msg="An error occured while generating totals file! Check general instructions to see specifications.")

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
        # return send_file(writer, as_attachment=False, mimetype = XLSX_MIMETYPE)
        return send_from_directory(out_file_path, out_name, as_attachment=True)
    #return render_template('kgs_home.html', msg="Totals file generated!", out_file_path = out_file_path)
    except:
        return render_template('kgs_home.html', msg="An error occured while generating totals_remarks file! Check general instructions to see specifications.")

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
        # return send_file(writer, as_attachment=False, mimetype = XLSX_MIMETYPE)
        return send_from_directory(out_file_path, out_name, as_attachment=True)
    except:
        return render_template('kgs_home.html', msg="An error occured while generating related party file! Check general instructions to see specifications.")

@app.route('/purchase_sheet')
def purchase_sheet():
    try:
        file_var = session['name']
        in_file_path = os.path.join(APP_ROUTE+"\\uploads_folder", file_var)
        out_file_path = os.path.join(APP_ROUTE+"\\output_folder")
        pur_sheet = bank_program.purchase_sheet(in_file_path)
        out_name = str(file_var.split(".", 1)[0]) + '-purchase_rate_sheet.xlsx'
        final_f = out_file_path + '\\' + out_name
        writer = pd.ExcelWriter(final_f, engine='xlsxwriter')
        pur_sheet.to_excel(writer, sheet_name='Purchase Rate Sheet')
        writer.save()
        # return send_file(writer, as_attachment=False, mimetype = XLSX_MIMETYPE)
        return send_from_directory(out_file_path, out_name, as_attachment=True)
    except:
        return render_template('kgs_home.html', msg="An error occured while generating Purchase Rate sheet file! Check general instructions to see specifications.")

@app.route('/add_features', methods = ["GET","POST"])
def add_features():
    out_file_path = os.path.join(APP_ROUTE+"\\output_folder")
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
            path_with_name_list = []
            file_list = request.files.getlist("multi_files")
            for files in file_list:
                files.save(os.path.join(APP_ROUTE+"\\uploads_folder", files.filename))
                path_with_name_list.append(os.path.join(APP_ROUTE+"\\uploads_folder", files.filename))
                name_list.append(files.filename)
            result = bank_program.search_terms(path_with_name_list, rel_party_var)
            out_name = str('variable_search_results.xlsx')
            final_f = out_file_path + '\\' + out_name
            writer = pd.ExcelWriter(final_f, engine='xlsxwriter')
            result.to_excel(writer, sheet_name='Search Results')
            writer.save()
            val = name_list
            # return send_file(writer, attachment_filename=out_name, as_attachment=False, mimetype = XLSX_MIMETYPE)
            return send_from_directory(APP_ROUTE+"\\uploads_folder", out_name, as_attachment=True)
            #return render_template('kgs_addfeatures.html', message="File(s) uploaded!", val = name_list)
    return render_template('kgs_addfeatures.html', message="Waiting to submit...") 

@app.route('/sample_extraction')
def sample_extraction():
    try:
        file_var = session['name']
        in_file_path = os.path.join(APP_ROUTE+"\\uploads_folder", file_var)
        out_file_path = os.path.join(APP_ROUTE+"\\output_folder")
        df = pd.read_excel(in_file_path)
        new_df = df.sample(frac=0.05, replace=True, random_state=1)
        out_name = str(file_var.split(".", 1)[0]) + '-random_sample.xlsx'
        final_f = out_file_path + '\\' + out_name
        writer = pd.ExcelWriter(final_f, engine='xlsxwriter')
        new_df.to_excel(writer)
        writer.save()
        return send_from_directory(out_file_path, out_name, as_attachment=True)
    except:
        return render_template('kgs_home.html', msg="An error occured while generating Purchase Rate sheet file! Check general instructions to see specifications.")

@app.route('/bb_reconcilation_uploader', methods = ["GET","POST"])
def bb_reconcilation_uploader():
    if request.method == 'POST':
        stmt_file = request.files['bank_stmt_file']
        book_file = request.files['bank_book_file']
        # If the user does not select a file, the browser submits an empty file without a filename.
        if stmt_file.filename == '' or book_file.filename == '':
            flash('No selected file')
            return render_template('kgs_bbrecon.html')
        if stmt_file and book_file:
            stmt_file = request.files['bank_stmt_file']
            book_file = request.files['bank_book_file']
            stmt_file.save(os.path.join(APP_ROUTE+"\\uploads_folder", stmt_file.filename))
            book_file.save(os.path.join(APP_ROUTE+"\\uploads_folder", book_file.filename))
            session['bank_stmt_file'] = str(stmt_file.filename)
            session['bank_book_file'] = str(book_file.filename)
            return render_template('kgs_bbrecon.html', message="Files uploaded!", val1 = stmt_file.filename, val2 = book_file.filename)        
    return render_template('kgs_bbrecon.html', message="Waiting to submit...") 

@app.route('/bb_reconcilation')
def bb_reconcilation():
    try:
        stmt_file = session['bank_stmt_file']
        book_file = session['bank_book_file']
        stmt_file_path = os.path.join(APP_ROUTE+"\\uploads_folder", stmt_file)
        book_file_path = os.path.join(APP_ROUTE+"\\uploads_folder", book_file)
        out_file_path = os.path.join(APP_ROUTE+"\\output_folder")
        out_name = str(stmt_file.lstrip(".")) + '-reconciled.xlsx'
        a = BBrecon.Bank_recon_class(bank_statement=stmt_file_path, bank_book=book_file_path)
        a.prepare_df()
        a.match_stmt_with_book()
        out_file = a.out_df()
        out_file.to_excel(out_file_path+'\\'+out_name,index=False)
        return send_from_directory(out_file_path, out_name, as_attachment=True)
    except:
        return render_template('kgs_bbrecon.html', message="Some Error Occured!")

@app.route('/del_files')
def delete_files():
    dir = APP_ROUTE+"\\uploads_folder"
    filelist = glob.glob(os.path.join(dir, "*"))
    for f in filelist:
        os.remove(f)
    return redirect(home_url)
    #return render_template('kgs_home.html', message="Upload folder cleared!")

