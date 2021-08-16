"""
File:    main_app.py

Author:   Mohd Aanis Noor (https://github.com/oron-sinaa) 
Company:  KG Somani & Co (https://www.kgsomani.com/) 
Date:     Started Jun 2021

Summary of File: 

This file contains flask python code that is the frontend for
bank_program.py which analyzes & generates reports from several type
provided bank document files. The code as in August 2021 is working
but still in development phase.
"""

# in-built modules:
from flask import Flask, request, redirect, url_for, render_template, session, send_file, send_from_directory, flash
from flask_wtf import Form
from werkzeug.utils import secure_filename
import os, glob
import pandas as pd
# self modules:
import bank_program as bp
import BBrecon as bbr


# pre-requisite variables 
app = Flask(__name__)
app.secret_key = 'kgsfintechconnectkey'
HOME_URL = "http://127.0.0.1:5000/"
APP_ROUTE = os.path.dirname(os.path.abspath(__file__))
XLSX_MIMETYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
UPLOADS_FOLDER = APP_ROUTE + "\\uploads_folder\\"
OUTPUT_FOLDER = APP_ROUTE + "\\output_folder"

@app.route('/') 
def home():  
    return render_template('kgs_home.html')

@app.route('/instructions')
def instructions():
    return render_template('kgs_instructions.html')

@app.route('/remarks', methods = ["GET","POST"])
def remarks():
    if request.method == 'POST':
        uploaded_remarks_file = request.files['uploaded_remarks_file']
        # If no file selected, ask again
        if uploaded_remarks_file.filename == '':
            return render_template('kgs_remarks.html', msg1="Waiting to upload...")
        if uploaded_remarks_file:
            uploaded_remarks_file = request.files['uploaded_remarks_file']
            session['uploaded_remarks_file_name'] = str(uploaded_remarks_file.filename)
            uploaded_remarks_file.save(os.path.join(UPLOADS_FOLDER, uploaded_remarks_file.filename))
            return render_template('kgs_remarks.html', msg1="File \"", msg2="\" uploaded!", val = uploaded_remarks_file.filename)
    return render_template('kgs_remarks.html', msg1="Waiting to upload...")

@app.route('/remarks_generated')
def remarks_generated():
    try:
        uploaded_remarks_file_name = session['uploaded_remarks_file_name']
        sheet = bp.import_file(UPLOADS_FOLDER + uploaded_remarks_file_name)
        cleaned_sheet = bp.data_clean(sheet[0])
        rem_df = bp.generate_remarks(sheet[1], cleaned_sheet)
        out_name = str(str(uploaded_remarks_file_name).split(".", 1)[0]) + '-remarks.xlsx'
        f_out = bp.export_remarks(rem_df, str(uploaded_remarks_file_name), OUTPUT_FOLDER)
        return send_from_directory(OUTPUT_FOLDER, out_name, as_attachment=True)
    except:
        return render_template('kgs_remarks.html', msg1="An error occured! Check general instructions.")

@app.route('/totals', methods = ["GET","POST"])
def totals():
    if request.method == 'POST':
        uploaded_totals_file = request.files['uploaded_totals_file']
        # If no file selected, ask again
        if uploaded_totals_file.filename == '':
            return render_template('kgs_totals.html', msg1="Waiting to upload...")
        if uploaded_totals_file:
            uploaded_totals_file = request.files['uploaded_totals_file']
            session['uploaded_totals_file_name'] = str(uploaded_totals_file.filename)
            uploaded_totals_file.save(os.path.join(UPLOADS_FOLDER, uploaded_totals_file.filename))
            return render_template('kgs_totals.html', msg1="File \"", msg2="\" uploaded!", val = uploaded_totals_file.filename)
    return render_template('kgs_totals.html', msg1="Waiting to upload...")

@app.route('/totals_generated')
def totals_generated():
    try:
        uploaded_totals_file_name = session['uploaded_totals_file_name']
        sheet = bp.import_file(UPLOADS_FOLDER + uploaded_totals_file_name)
        cleaned_sheet = bp.data_clean(sheet[0])
        rem_df = bp.generate_remarks(sheet[1], cleaned_sheet)
        out_name = str(uploaded_totals_file_name.split(".", 1)[0]) + '-totals.xlsx'
        f_out = bp.totals_sheet(rem_df, str(uploaded_totals_file_name), OUTPUT_FOLDER)[0]
        return send_from_directory(OUTPUT_FOLDER, out_name, as_attachment=True)
    except:
        return render_template('kgs_totals.html', msg1="An error occured! Check general instructions.")

@app.route('/rel_party')
def rel_party():
    return render_template('kgs_rel_party.html')

@app.route('/pur_rate')
def pur_rate():
    return render_template('kgs_purchase.html')

@app.route('/samp_extraction', methods = ["GET","POST"])
def samp_extraction():
    if request.method == 'POST':
        uploaded_samples_file = request.files['uploaded_samples_file']
        # If no file selected, ask again
        if uploaded_samples_file.filename == '':
            return render_template('kgs_samp_ext.html', msg1="Waiting to upload...")
        if uploaded_samples_file:
            uploaded_samples_file = request.files['uploaded_samples_file']
            session['uploaded_samples_file_name'] = str(uploaded_samples_file.filename)
            uploaded_samples_file.save(os.path.join(UPLOADS_FOLDER, uploaded_samples_file.filename))
            return render_template('kgs_samp_ext.html', msg1="File \"", msg2="\" uploaded!", val = uploaded_samples_file.filename)
    return render_template('kgs_samp_ext.html', msg1="Waiting to upload...")

@app.route('/samples_generated')
def samples_generated():
    try:
        uploaded_samples_file_name = session['uploaded_samples_file_name']
        df = pd.read_excel(UPLOADS_FOLDER + uploaded_samples_file_name)
        new_df = df.sample(frac=0.05, replace=True, random_state=7)
        out_name = str(uploaded_samples_file_name.split(".", 1)[0]) + '-random_sample.xlsx'
        final_f = OUTPUT_FOLDER + '\\' + out_name
        writer = pd.ExcelWriter(final_f, engine='xlsxwriter')
        new_df.to_excel(writer)
        writer.save()
        return send_from_directory(OUTPUT_FOLDER, out_name, as_attachment=True)
    except:
        return render_template('kgs_samp_ext.html', msg1="An error occured! Check general instructions.")

@app.route('/search_terms', methods = ["GET","POST"])
def search_terms():
    if request.method == 'POST':
        file_list = request.files.getlist("uploaded_search_files")
        terms_to_search = request.form['terms_to_search']
        session['file_list_session'] = str(file_list)
        # If no file selected, ask again
        for files in file_list:
            if files.filename == '':
                return render_template('kgs_search_terms.html', msg="Waiting to submit...")
        if file_list:
            name_list = []
            path_with_name_list = []
            for files in file_list:
                files.save(os.path.join(UPLOADS_FOLDER, files.filename))
                path_with_name_list.append(os.path.join(UPLOADS_FOLDER, files.filename))
                name_list.append(files.filename)
            session['path_with_name_list_session'] = path_with_name_list
        return render_template('kgs_search_terms.html', msg=" file(s) uploaded!", val = len(name_list), terms=terms_to_search)
    return render_template('kgs_search_terms.html', msg=" Waiting to submit...")
            
@app.route('/search_terms_generated')
def search_terms_generated():
    # try:
    path_with_name_list = session['path_with_name_list_session']
    terms_to_search = session['file_list_session']
    result = bp.search_terms(list(path_with_name_list), str(terms_to_search))
    out_name = str('variable_search_results.xlsx')
    final_f = UPLOADS_FOLDER + '\\' + out_name
    writer = pd.ExcelWriter(final_f, engine='xlsxwriter')
    result.to_excel(writer, sheet_name='Search Results')
    writer.save()
    return send_from_directory(APP_ROUTE+"\\uploads_folder", out_name, as_attachment=True)
    # except:
    #     return render_template('kgs_search_terms.html', msg = "An error occured! Check general instructions.)

@app.route('/bbrecon', methods = ["GET","POST"])
def bbrecon():
    if request.method == 'POST':
        stmt_file = request.files['uploaded_bank_stmt_file']
        book_file = request.files['uploaded_bank_book_file']
        # If no file selected, ask again
        if stmt_file.filename == '' or book_file.filename == '':
            return render_template('kgs_bbr.html')
        if stmt_file and book_file:
            stmt_file = request.files['uploaded_bank_stmt_file']
            book_file = request.files['uploaded_bank_book_file']
            stmt_file.save(os.path.join(UPLOADS_FOLDER, stmt_file.filename))
            book_file.save(os.path.join(UPLOADS_FOLDER, book_file.filename))
            session['bank_stmt_file_name'] = str(stmt_file.filename)
            session['bank_book_file_name'] = str(book_file.filename)
            return render_template('kgs_bbr.html', message="Files uploaded!", val1 = stmt_file.filename, val2 = book_file.filename)        
    return render_template('kgs_bbr.html', message="Waiting to submit...")

@app.route('/bb_reconcilation_generated')
def bb_reconcilation_generated():
    try:
        stmt_file = session['bank_stmt_file_name']
        book_file = session['bank_book_file_name']
        stmt_file_path = os.path.join(UPLOADS_FOLDER, stmt_file)
        book_file_path = os.path.join(UPLOADS_FOLDER, book_file)
        out_name = str(stmt_file.lstrip(".")) + '-reconciled.xlsx'
        a = bbr.Bank_recon_class(bank_statement=stmt_file_path, bank_book=book_file_path)
        a.prepare_df()
        a.match_stmt_with_book()
        out_file = a.out_df()
        out_file.to_excel(OUTPUT_FOLDER +'\\'+out_name, index=False)
        return send_from_directory(OUTPUT_FOLDER, out_name, as_attachment=True)
    except:
        return render_template('kgs_bbr.html', message="Some Error Occured!")

if __name__ =='__main__':  
    app.run(debug = True) 