"""
File:    views.py 

Author:   Mohd Aanis Noor & Mansi Gupta 
Company:  KG Somani & Co (https://www.kgsomani.com/) 
Date:     Started Jun 2021

Summary of File: 

This module helps achieve bank book reconcilation.
The code as in August 2021 is working,
but still in development phase.
"""

import pandas as pd
import numpy as np
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

class Bank_recon_class():
    def __init__(self, bank_statement, bank_book):
        self.xls_book= bank_book
        self.xls_stmt= bank_statement
        self.xls_output= str(bank_statement).rstrip('\\') + 'v3_Output.xlsx'
        self.df_book=pd.read_excel(self.xls_book)
        self.s_date_col1="DOC DATE"
        self.s_amt1_col1="Amount"
        self.s_amt2_col1=None
        self.s_remarks_col1="Remarks"
        self.s_chq_col1="ASSIGNMENT"
        self.df_stmt=pd.read_excel(self.xls_stmt,sheet_name='Bank Statement')
        self.s_date_col2="Txn Date"
        self.s_amt1_col2="Dr Amount"
        self.s_amt2_col2="Cr Amount"
        self.s_remarks_col2="Description"
        self.s_chq_col2="Cheque No."
        self.df_out=pd.DataFrame()
        self.df_fuz=pd.DataFrame()
        self.fddf_book=self.df_book
        self.fddf_book=self.fddf_book.drop(['CODE','Bank Account','B.AREA','DOC NO','DOC TYPE','PARTY  CODE','STATION','RECEIPT/PAYMENT','Amount','CUST ASSIGNMENT','ASSIGNMENT','WORKFLOW-ID','DESCRIPTION','BKTXT'],axis=1)
        self.fddf_stmt=self.df_stmt
        self.acc=90.0

    def show_df(self):
        print(self.df_book)
        print(self.df_stmt)

    def out_df(self):
        df1 = self.df_out
        df1 = df1.append(self.df_fuz)
        df1 = self.rearrange_columns(df1)
        # df1.to_excel(self.xls_output,index=False)
        return df1

    def rearrange_columns(self,df):
        col_titles=['Txn No.','Txn Date','Description','Branch Name','Cheque No.','Dr Amount','Cr Amount',
        'Balance','DOC DATE','Party Name as per bank book','match_type','Remarks as per the bank book','Amount in bank book']
        df = df.reindex(columns=col_titles)
        return df

    def prepare_df(self):
        self.df_book['date_amount'] =''
        self.df_book['all_cols'] =''

    def match_stmt_with_book(self):
        def assign_values(ps_msg):
            try:
                dct['DOC DATE']=df_match['DOC DATE'].values[0]
            except IndexError:
                dct['DOC DATE']=None
            try:
                dct['Party Name as per bank book']=df_match['PARTY NAME'].values[0]
            except IndexError:
                dct['Party Name as per bank book']=None
            dct['match_type']=ps_msg
            try:
                dct['Remarks as per the bank book']=df_match['Remarks'].values[0]
            except IndexError:
                dct['Remarks as per the bank book']=None
            try:
                dct['Amount in bank book']=df_match[self.s_amt1_col1].values[0]
            except IndexError:
                dct['Amount in bank book']=None
            lst.append(dct)
        lst=[]
        for idx,stmt_rec in self.df_stmt.iterrows():
            # print('Processing ',idx,' of ',len(self.df_stmt))
            dct=stmt_rec.to_dict()
            #chq matching
            df_match = self.df_book.loc[self.df_book[self.s_chq_col1]==stmt_rec[self.s_chq_col2]]
            if len(df_match)>0:
                assign_values('Chq')
                continue
            #UNIQ Amount in Statement
            df_match = self.df_book.loc[self.df_book[self.s_amt1_col1]==stmt_rec[self.s_amt1_col2]]
            df_match1 = self.df_book.loc[self.df_book[self.s_amt1_col1]==stmt_rec[self.s_amt2_col2]]
            if len(df_match)==0 and len(df_match1)==0:
                # print('No Amount Match for ',idx,':',stmt_rec[self.s_amt1_col2],stmt_rec[self.s_amt2_col2])
                assign_values('Unique Amount, No match in Book')
                continue
            #UNIQ DB matching
            df_match = self.df_book.loc[self.df_book[self.s_amt1_col1]==stmt_rec[self.s_amt1_col2]]
            if len(df_match)==1:
                # print('Amount Match for ',idx,':',stmt_rec[self.s_amt1_col2],df_match[self.s_amt1_col1])
                assign_values('Match found in DB Amount')
                continue
            #UNIQ CR matching
            df_match = self.df_book.loc[self.df_book[self.s_amt1_col1]==stmt_rec[self.s_amt2_col2]]
            if len(df_match)==1:
                # print('Amount Match for ',idx,':',stmt_rec[self.s_amt2_col2],df_match[self.s_amt1_col1])
                assign_values('Match found in CR Amount')
                continue
            #Date and DB Amount
            df_match = self.df_book.loc[(self.df_book[self.s_date_col1]==stmt_rec[self.s_date_col2].replace('/','.')) \
                        & (self.df_book[self.s_amt1_col1]==stmt_rec[self.s_amt1_col2])]
            if len(df_match)>0:
                # print('Amount Match for ',idx,':',stmt_rec[self.s_date_col2],df_match[self.s_date_col1])
                assign_values('Date and DB Amount match')
                continue
            #Date and CR Amount
            df_match = self.df_book.loc[(self.df_book[self.s_date_col1]==stmt_rec[self.s_date_col2].replace('/','.')) \
                        & (self.df_book[self.s_amt1_col1]==stmt_rec[self.s_amt2_col2])]
            if len(df_match)>0:
                # print('Amount Match for ',idx,':',stmt_rec[self.s_date_col2],df_match[self.s_date_col1])
                assign_values('Date and CR Amount match')
                continue
            dct['DOC DATE']=None
            dct['Party Name as per bank book']=None
            dct['match_type']='NO MATCH'
            dct['Remarks as per the bank book']=None
            lst.append(dct)
        self.df_out = pd.DataFrame(lst)


# if __name__ == "__main__":
#     a = Bank_recon_class()
#     a.prepare_df()
#     a.show_df()
#     a.match_stmt_with_book()
#     a.out_df()


