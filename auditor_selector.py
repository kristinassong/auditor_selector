#!/usr/bin/env python

""" Automated lead auditor selector based on availability and previous experience. """

import sys
import pandas as pd
from datetime import datetime, timedelta

def parse_rec(f_history):
    # import audit records
    df = pd.read_excel(f_history,
        usecols=['Organization','Material Category','Service Category','Auditor Name'])
    # remove rows w/o Auditor Name, Material/Service Category
    df = df[df['Auditor Name'].notna()].reset_index(drop=True)
    df_mat = df[df['Material Category'].notna()].reset_index(drop=True)
    df_mat.drop('Service Category', axis=1, inplace=True)
    df_ser = df[df['Service Category'].notna()].reset_index(drop=True)
    df_ser.drop('Material Category', axis=1, inplace=True)
    return df_mat,df_ser

def check_avail(start_date,end_date,sample,pad=1):
    # DEFAULT VALUE: date padding +-1 day
    # FEATURE1: Search for auditors that are available on the potential audit date

    auditors = pd.read_excel(sample, sheet_name='auditors').auditors.tolist() # list of auditors
    df = pd.read_excel(sample,sheet_name='schedule',dtype={'start_date': datetime,'end_date':datetime}) # current schedule of auditors

    new_start = start_date - timedelta(days=pad*2)
    new_end = end_date + timedelta(days=pad*2)

    # get list of auditors that are already occupied during audit period
    selected_tmp = df[(df['start_date'] <= new_end)]
    selected = selected_tmp[(selected_tmp['end_date'] >= new_start)].reset_index(drop=True)
    aud1 = selected.auditor1.values.tolist()
    aud2 = selected.auditor2.values.tolist()
    busy = set(aud1+aud2)

    # compare initial auditor list with 'busy' auditors on the potential audit period
    result = list(set(auditors).difference(busy))
    return result

def lead_co_aud(df):
    # function to handle audit records with 2+ auditors 
    comma = df[df['Auditor'].str.contains(',')].reset_index(drop=True)
    for i in range(len(comma)):
        lst = comma.loc[i,'Auditor'].split(',')
        counter = 0
        for el in lst:
            el = el.lstrip() # remove leading white spaces in front of name
            index=0
            row = df[df['Auditor']==el]
            if len(row) >0 :
                index = row.index.values[0]
            else: # If NONE found, then add to the list
                new_info = pd.DataFrame({'Auditor':el,'Count':[0]})
                df = pd.concat([df,new_info],ignore_index=True)
                index = df[df['Auditor']==el].index.values[0]
            if counter == 0: # lead auditor
                df.loc[index,'Count'] += 1
            else: # co-auditor
                df.loc[index,'Count'] += 0.5
            counter+=1
    # remove rows with 2+ auditors
    df = df[~df['Auditor'].str.contains(',')].reset_index(drop=True)
    return df

def check_exp(rec,subtyp,prod,avail_auditors):
    df_mat, df_ser = parse_rec(rec)
    # FEATURE2: Rank all available auditors by experience (i.e. material type/ service type)
    sel = pd.DataFrame()
    if subtyp == 'Supplier':
        sel = df_mat[df_mat['Material Category'].str.contains(prod,case=False)]['Auditor Name']
    else:
        sel = df_ser[df_ser['Service Category'].str.contains(prod,case=False)]['Auditor Name']
    # Rank & count audit record per auditor
    df_rank = pd.DataFrame(sel.value_counts())
    df_rank = df_rank.reset_index()
    df_rank.columns = ['Auditor', 'Count'] # change column names

    # Split auditor data with 2+ auditors (lead & co-auditor)
    df_rank = lead_co_aud(df_rank)

    # retrieve auditor data for only those in the list
    auditor_rank = pd.DataFrame()
    for aud in avail_auditors:
        auditor_rank = pd.concat([auditor_rank,df_rank[df_rank['Auditor']==aud]],ignore_index=True)

    # return ordered list of auditors from high --> low level of experience
    auditor_rank = auditor_rank.sort_values(by = 'Count',ascending=False).reset_index(drop=True)
    return auditor_rank

def final_decision(df,opt1,opt2):
    # make final decision of lead auditor
    lead = df.loc[0,'Auditor']
    # If requested, choose co-auditor based on 1) experience rank #2 OR 
    # 2) auditor w/ no prev experience in this field
    coaud=None
    if opt1 == 'Yes':
        if opt2 == 'Experienced':
            coaud = df.loc[1,'Auditor']
        elif opt2 == 'New':
            coaud = df.loc[len(df)-1,'Auditor']
        else:
            print('ERROR.')
    return lead, coaud

def main():
    rec = sys.argv[1] # excel file containing EQUIS audit records
    sample = sys.argv[2] # excel file containing sample data (list of auditors and their current schedule)

    # font colors
    # ERROR
    CRED = '\033[91m'
    CEND = '\033[0m'
    # PROMPT
    PROMPT = '\033[36m'
    # RESULT
    RES = '\033[92m'

    # USER INPUT: company name & supplier type
    org_entry = input(PROMPT + '\nEnter Organization: ' + CEND)
    if org_entry == '':
        sys.exit(CRED + "\nERROR: Please enter the name of the organization." + CEND)

    sup_entry = input(PROMPT + '\nEnter Supplier type (Supplier/Service Provider): ' + CEND) # Supplier OR Service Provider
    if not (sup_entry == 'Supplier' or sup_entry == 'Service Provider'):
        sys.exit(CRED + "\nERROR: Please specify the proper supplier type (Supplier/Service Provider)." + CEND)

    # USER INPUT: audit date
    try:
        start_date_entry = input(PROMPT + '\nEnter audit start date (YYYY-M-D): ' + CEND)
        start_year, start_month, start_day = map(int, start_date_entry.split('-'))
        end_date_entry = input(PROMPT + '\nEnter audit end date (YYYY-M-D): ' + CEND)
        end_year, end_month, end_day = map(int, end_date_entry.split('-'))
    except ValueError as v:
        sys.exit(CRED + "\nERROR: Please indicate the audit date in YYYY-M-D format." + CEND)
    start_date = datetime(start_year, start_month, start_day)
    end_date = datetime(end_year, end_month, end_day)

    # Search for auditors that are available on the given audit date period (considering padded dates)
    pad = input(PROMPT + '\nOPTIONAL: Specify number of days before and after audit for preparation and reporting. (i.e. ' 
        + ' +- X days)\n' + 'Default '+' +- 1 day from actual audit date.\n:' + CEND)
        #+ u'\u00B1'+ ' X days)\n' + 'Default '+u'\u00B1'+' 1 day from actual audit date.\n:' + CEND)

    if pad == "": # use default value
        avail_auditors = check_avail(start_date,end_date,sample)
    elif not pad.isnumeric():
        sys.exit(CRED + "\nPlease enter a numeric value." + CEND)
    else: # padded dates specified by user
        pad = int(pad)
        avail_auditors = check_avail(start_date,end_date,sample,pad)
    str_start = str(start_year) + '-' + str(start_month) + '-' + str(start_day)
    str_end = str(end_year) + '-' + str(end_month) + '-' + str(end_day)
    print(RES + '\nAuditors available for audit from ' + str_start + ' to ' + str_end + ':\n' + CEND)
    if len(avail_auditors)==0:
        print(RES + 'NONE.' + CEND)
    else:
        for name in avail_auditors:
            print(name)

    # USER INPUT: product category
    product = input(PROMPT + '\nEnter Material/Service Category: ' + CEND)
    by_exp = check_exp(rec,sup_entry,product,avail_auditors)
    if len(by_exp)>0:
        print(by_exp)
    else:
        print(RES + 'NONE.' + CEND)

    # Make final lead auditor choice and co-auditor if requested
    opt1 = input(PROMPT + '\nRequest for co-auditor suggestion (Yes/No): ' + CEND)
    opt2 = 'No'
    if opt1 == 'Yes':
        opt2 = input(PROMPT + '\nSelect Co-Auditor Type (Experienced/New): ' + CEND)
    lead, coaud = final_decision(by_exp,opt1,opt2)
    print(RES + '\n>> Suggested Lead Auditor for ' + org_entry + ' from ' + str_start + ' to ' + str_end + ': ' + lead + CEND)
    if opt1 == 'Yes' and coaud != None:
        print(RES + '>> Suggested Co-auditor: ' + coaud + '\n' + CEND)

if __name__ == '__main__':
    main()