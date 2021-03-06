"""
File:    bank_program.py 

Author:   Mohd Aanis Noor (https://github.com/oron-sinaa) 
Company:  KG Somani & Co (https://www.kgsomani.com/) 
Date:     Started Jun 2021
Partner:  Mansi 

Summary of File: 

This file contains functions that generate several files analyzed
from provided bank documents. The code as in August 2021 is working
but still in development phase.
"""


import pandas as pd
import numpy as np
import re
from fuzzywuzzy import process, fuzz

stop_words = ['BY','FOR','OF','REVERSAL','RETURN','TRANSFER','NEFT','FROM','AGAINST','TO','DEBIT','THROUGH','CHEQUE','FOREIGN',
              'NO','RTGS','UTR','INB','JAN','JANUARY','FEB','FEBRUARY','MAR','MARCH','APR','APRIL','MAY','JUN','JUNE','JUL',
              'JULY','AUG','AUGUST','SEP','SEPT','SEPTEMBER','OCT','OCTOBER','NOV','NOVEMBER','DEC','DECEMBER','CLOSURE',
              'NRTGS','IN','MR','MRS','C','AC','FEES','CASH','WITHDRAWAL','CLG','TRF','REVERSAL','NEFT_IN','NEFT_OUT','NEFT_CHRG',
              'BILLDESK','CREDIT','TRF','TFR','TT','TR','TFRR','TF','TL','MARGIN','ETFR','B/F','T/F','BILL ID','IMPS','DR','TXT',
              'SFMS','SCBL','SBIN','ICIC','ICICI','HDFC','ORBC','MAHB','HDFC','PUNB','BARB','UTIB','XLSX','BULK','MTFR']

exclude_from_cleaning = ['786','360','365','180']

# - import spreadsheet to work on - #
def import_file(f_name):
    df = pd.read_excel (f_name, usecols = ['Particulars'])
    #add ID column and make it the index
    df.insert(0, 'ID', df.index+2)
    df = df.set_index('ID')
    #import a copy
    df_def = pd.read_excel (f_name, header=0)
    #add ID column and make it the index
    df_def.insert(0, 'ID', df_def.index+2)
    df_def = df_def.set_index('ID')
    return df, df_def, f_name

# - clean and format file data - #
def data_clean(import_file):
    df = import_file
    #remove whitespaces
    df['Particulars'] = df['Particulars'].str.strip()
    #convert everything to uppercase
    df['Particulars'] = df['Particulars'].str.upper()
    #remove stop words (pre-cleaning)
    pat = '|'.join(r"\b{}\b".format(x) for x in stop_words)
    df['Particulars'] = df['Particulars'].str.replace(pat, '', regex=True)
    #add spaces next to special characters
    df['Particulars'] = df['Particulars'].str.replace(r'([^&\w\s])'," \\1", regex=True)
    #remove special characters
    df['Particulars'] = df['Particulars'].str.replace(r'([^\w\s\&])',"", regex=True)
    #remove alphanumeric and numeric
    if df['Particulars'].any() not in exclude_from_cleaning:
        df['Particulars'] = df['Particulars'].str.replace('\w+\d+', '', regex=True)
        df['Particulars'] = df['Particulars'].str.replace('\d+', '', regex=True)
    #replace na values with single space
    df['Particulars'] = df['Particulars'].fillna(" ")
    #modify for specific keywords
    df.loc[df['Particulars'].str.contains('|'.join(['INT', 'INTEREST']), case=False), 'Particulars'] = 'Interest'
    df.loc[df['Particulars'].str.contains('|'.join(['INB', 'EOD']), case=False), 'Particulars'] = 'Interbank Transfer'
    df.loc[df['Particulars'].str.contains('GST', case=False), 'Particulars'] = 'GST Refund'
    df.loc[df['Particulars'].str.contains('SMS', case=False), 'Particulars'] = 'SMS Charges'
    df.loc[df['Particulars'].str.contains('|'.join(['SALARY', 'WAGES', 'WAGE']), case=False), 'Particulars'] = 'Salary & Wages'
    df.loc[df['Particulars'].str.contains('FOREX', case=False), 'Particulars'] = 'Foreign Currency Conversion Tax'
    df.loc[df['Particulars'].str.contains('CAR', case=False), 'Particulars'] = 'Maintainence Charges'
    df.loc[df['Particulars'].str.contains('CASH' and 'DEPOSIT', case=False), 'Particulars'] = 'Cash Deposit'
    df.loc[df['Particulars'].str.contains('WCL', case=False), 'Particulars'] = 'Repayment of WDCL'
    df.loc[df['Particulars'].str.contains('BCCALC', case=False), 'Particulars'] = 'Bccalc Recovery Charges'
    #remove stop words (post-cleaning)
    pat = '|'.join(r"\b{}\b".format(x) for x in stop_words)
    df['Particulars'] = df['Particulars'].str.replace(pat, '', regex=True)
    #remove whitespaces
    df['Particulars'] = df['Particulars'].str.strip()
    #remove multiple spaces
    df['Particulars'] = df['Particulars'].str.replace(' +', ' ', regex=True)
    #text formatting
    df['Particulars'] = df['Particulars'].str.title()
    #define all unidentified cases
    df = df.applymap(lambda x: '- Unidentified -' if (x == r'(.) ') else x)
    df = df.applymap(lambda x: '- Unidentified -' if isinstance(x, str) and ((not x) or (x.isspace()) or (len(x)==1)) else x)
    df.loc[df['Particulars'].str.contains('Does Not', case=False), 'Particulars'] = '- Unidentified -'
    return df

# - entry resolution - #
def entry_resolution(export_remarks):
    remark_df = export_remarks
    resolution_threshold = 75
    unique_remarks = remark_df['Remarks'].unique().tolist()
    #this automatically replaces very similar terms:
    result_pre = process.dedupe(unique_remarks, threshold=90, scorer=fuzz.token_sort_ratio)
    resolved_list = []
    for item in result_pre:
        result = process.extractBests(item,
                                      result_pre,
                                      scorer=fuzz.token_sort_ratio,
                                      score_cutoff=resolution_threshold,
                                      limit=None)
        if len(result)>1 and result not in resolved_list:
            resolved_list.append(result)
    resolved_list = sorted(resolved_list)
    return resolved_list

# - generate remarks sheet - #
def generate_remarks(import_file, cleaned_sheet):
    df = cleaned_sheet
    #change edited column name to 'Remarks'
    df.rename(columns = {'Particulars':'Remarks'}, inplace = True)
    remark_df = pd.merge(import_file, df, on='ID')
    #this automatically replaces very similar terms:
    unique_remarks = remark_df['Remarks'].unique().tolist()
    process.dedupe(unique_remarks, threshold=95, scorer=fuzz.token_sort_ratio) 
    return remark_df

# - export remarks sheet - #
def export_remarks(rem_df, f_name, path):
    out_name = str(f_name.split(".", 1)[0]) + '-remarks.xlsx'
    final_f = path + '\\' + out_name
    rem_df.to_excel(final_f) 
    return final_f

# --- export totals sheet --- #
def totals_sheet(export_remarks, f_name, path):
    remark_df = export_remarks
    #remove words from numerical columns
    remark_df['WITHDRAWALS'] = remark_df['WITHDRAWALS'].replace(r'([/\D+/g])',0, regex=True).astype(float)
    remark_df['DEPOSITS'] = remark_df['DEPOSITS'].replace(r'([/\D+/g])',0, regex=True).astype(float)
    #remove commmas,blanks from numerical columns
    remark_df['WITHDRAWALS'] = remark_df['WITHDRAWALS'].replace(',', '').astype(float)
    remark_df['DEPOSITS'] = remark_df['DEPOSITS'].replace(',', '').astype(float)
    remark_df['WITHDRAWALS'] = remark_df['WITHDRAWALS'].fillna(0)
    remark_df['DEPOSITS'] = remark_df['DEPOSITS'].fillna(0)
    remark_df['WITHDRAWALS'] = remark_df['WITHDRAWALS'].replace(' ', 0).astype(float)
    remark_df['DEPOSITS'] = remark_df['DEPOSITS'].replace(' ', 0).astype(float)
    #total deposits/withdrawal
    sum_dict= {}
    for record in remark_df.values:
        remark = record[remark_df.columns.get_loc("Remarks")]
        if remark not in sum_dict:
            sum_dict[remark] = {"Total withdrawal":0,"Total deposit":0}
        withdrawal = record[remark_df.columns.get_loc('WITHDRAWALS')]
        sum_dict[remark]['Total withdrawal'] += withdrawal
        deposit = record[remark_df.columns.get_loc('DEPOSITS')]
        sum_dict[remark]['Total deposit'] += deposit
    #totals dataframe
    final_totals_df = pd.DataFrame.from_dict(sum_dict, orient ='index')
    out_name = str(f_name.split(".", 1)[0]) + '-totals.xlsx'
    final_f = path + '\\' + out_name
    final_totals_df.to_excel(final_f)
    return final_f, final_totals_df
    """
    writer = pd.ExcelWriter(str(f_name.split(".", 1)[0]) + ' - processed.xlsx', engine='xlsxwriter')
    final_remark_df.to_excel(writer, sheet_name='Remarks')
    final_totals_df.to_excel(writer, sheet_name='Totals')
    writer.save()
    """

# -generates related party sheet- #
def related_party(search_list, df):
    threshold = 75 #%
    search_list = list(map(lambda x:x.title(), search_list))
    for item in search_list:
        for itemdf in df['Remarks']:
            score = fuzz.token_sort_ratio(item, itemdf)
            if score > threshold:
                df.loc[df.Remarks == itemdf, "is Related Party?"] = "Related Party"
    return df

# -generates purchase (rate) sheet- #
def purchase_sheet(f_name):
    in_sheet = pd.read_excel (f_name, sheet_name='feb purchase', usecols=[
        'Vendor Name Reclassified','Description','Unit','Quantity','Invoice Amount','Rate PU','Material Type'])
    #remove words from numerical columns
    in_sheet['Quantity'] = in_sheet['Quantity'].replace(r'([/\D+/g])',0, regex=True).astype(float)
    in_sheet['Rate PU'] = in_sheet['Rate PU'].replace(r'([/\D+/g])',0, regex=True).astype(float)
    in_sheet['Invoice Amount'] = in_sheet['Invoice Amount'].replace(r'([/\D+/g])',0, regex=True).astype(float)
    #remove commmas,blanks from numerical columns
    in_sheet['Quantity'] = in_sheet['Quantity'].replace(',', '').astype(float)
    in_sheet['Rate PU'] = in_sheet['Rate PU'].replace(',', '').astype(float)
    in_sheet['Invoice Amount'] = in_sheet['Invoice Amount'].replace(',', '').astype(float)
    in_sheet['Quantity'] = in_sheet['Quantity'].fillna(0)
    in_sheet['Rate PU'] = in_sheet['Rate PU'].fillna(0)
    in_sheet['Invoice Amount'] = in_sheet['Invoice Amount'].fillna(0)
    in_sheet['Quantity'] = in_sheet['Quantity'].replace(' ', 0).astype(float)
    in_sheet['Rate PU'] = in_sheet['Rate PU'].replace(' ', 0).astype(float)
    in_sheet['Invoice Amount'] = in_sheet['Invoice Amount'].replace(' ', 0).astype(float)
    pd.set_option('display.precision',5)
    group = in_sheet.groupby(['Description','Vendor Name Reclassified','Unit','Material Type'], as_index=True)
    # apply stats operation and append columns to result dataframe
    result = pd.DataFrame()
    result['Sum of Quantity'] = group.sum()['Quantity']
    result['Sum of Invoice Amount'] = group.sum()['Invoice Amount']
    result['Min Rate PU'] = group.min()['Rate PU']
    result['Max Rate PU'] = group.max()['Rate PU']
    result['Avg Rate PU'] = group.mean()['Rate PU']
    result['Diff Rate PU'] = result['Max Rate PU'] - result['Min Rate PU']
    return result

# -search specific terms- #
def search_terms(sheet_list, search_terms):
    # terms to search for separated by commas
    search_terms = list(map(str, search_terms.split(',')))
    # sheets to search in
    sheet_list = sheet_list
    # sheets to search in; converted into dataframe
    list_df = list()
    for i in range(len(sheet_list)):
        list_df.append(pd.read_excel(sheet_list[i]))
    del sheet_list

    frames = list()
    for df in list_df:
        for item in search_terms:
            df_logic = df[df.apply(lambda row: row.astype(str).str.contains(item, case=False).any(), axis=1)]
            frames.append(df_logic) 
    result = pd.concat(frames, ignore_index = True)
    return result