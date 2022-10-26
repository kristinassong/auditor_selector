#!/usr/bin/env python

from pickle import FALSE
import PySimpleGUI as sg
import auditor_selector
import pandas as pd
from datetime import datetime


sg.theme('LightBlue6')

# Define the window's contents (Initial Window)
home = [[sg.Text("Lead Auditor Selector (LAS)",font='black 12 bold')],
          [sg.Button('Start'), sg.Button('Quit')]
        ]

input = [[sg.Text("Enter Organization:")],
          [sg.Input(key='ORG')],
          [sg.Text("Enter Supplier Type (Supplier/Service Provider):")],
          [sg.Input(key='SUP')],
          [sg.Text("Enter Audit Start Date (YYYY-M-D):")],
          [sg.Input(key='START_DATE')],
          [sg.Text("Enter Audit End Date (YYYY-M-D):")],
          [sg.Input(key='END_DATE')],
          [sg.Text("OPTIONAL: Specify number of days before and after audit \nfor preparation and reporting." +  
            "(i.e. " + u'\u00B1'+ ' X days)')],
          [sg.OptionMenu(values=[0,1,2,3,4,5,6,7,8,9],default_value=1,key='PAD')],
          [sg.Button('Next'), sg.Button('Clear'), sg.Button('Quit')]
        ]

avail = [[sg.Text("Auditors available for audit:",font='black 10 bold')],
          [sg.Text(key='AVAIL',font='black 10 bold')],
          [sg.Text("Enter Material/Service Category:")],
          [sg.Input(key='CAT')],
          [sg.Text("Request for co-auditor suggestion (Yes/No):")],

          [sg.T("         "), sg.Radio('Yes', "RADIO_CO", default=False, key="RADIO"),
            sg.T("         "), sg.Radio('No', "RADIO_CO", default=True)],
          #[sg.Input(key='COAUD')],
          
          [sg.Text("If Yes, enter co-auditor type (Experienced/New):")],
          [sg.Input(key='COAUD-1')],
          [sg.Button('Submit'), sg.Button('Clear'), sg.Button('Quit')]
        ]

lead_only = [[sg.Text("Suggested Lead Auditor:")],
          [sg.Text(key='RES1')],
          [sg.Button('Quit')]
        ]

lead_co = [[sg.Text("Suggested Lead Auditor:")],
          [sg.Text(key='RES2',font='black 12 bold')],
          [sg.Text("Suggested Co-Auditor:")],
          [sg.Text(key='CO',font='black 12 bold')],
          [sg.Button('Quit')]
        ]

# Create the window
window = sg.Window('Home', home,finalize=True,size=(290, 80))

# Display and interact with the Window using an Event Loop
while True:
    event, values = window.read()
    # See if user wants to quit or window was closed
    if event == sg.WINDOW_CLOSED or event == 'Quit':
        break


    if event == 'Clear':
        for key in ['ORG','SUP','START_DATE','END_DATE','PAD']:
            window[key]('')

    if event == 'Start':
        win1 = sg.Window('User Input',input,alpha_channel=0).Finalize()
        win1.SetAlpha(1)
        window.close()
        window = win1
        # Run auditor selection program
        sample = 'sample_data.xlsx'
        rec = 'EQUIS_audit_record.xlsx'
        e1,val1 = win1.read()
        start_year, start_month, start_day = map(int, val1['START_DATE'].split('-'))
        start_date = datetime(start_year, start_month, start_day)
        end_year, end_month, end_day = map(int, val1['END_DATE'].split('-'))
        end_date = datetime(end_year, end_month, end_day)
        avail_auditors = auditor_selector.check_avail(start_date,end_date,sample)

    if event == 'Next':
        win2 = sg.Window('User Input',avail,alpha_channel=0).Finalize()
        win2['AVAIL'].update(", ".join(avail_auditors))
        win2.SetAlpha(1)
        window.close()
        window = win2
        e2,val2 = win2.read()

    if event == 'Submit':
        if val2['RADIO'] ==False:
          win3 = sg.Window('Result',lead_only,alpha_channel=0).Finalize()
          win3.SetAlpha(1)
          window.close()
          window = win3
          by_exp = auditor_selector.check_exp(rec,val1['SUP'],val2['CAT'],avail_auditors)
          lead, coaud = auditor_selector.final_decision(by_exp,'NO','New')
          window['RES1'].update(lead)
        else:
          win3 = sg.Window('Result',lead_co,alpha_channel=0).Finalize()
          win3.SetAlpha(1)
          window.close()
          window = win3
          by_exp = auditor_selector.check_exp(rec,val1['SUP'],val2['CAT'],avail_auditors)
          print(by_exp)
          lead, coaud = auditor_selector.final_decision(by_exp,'Yes',val2['COAUD-1'])
          window['RES2'].update(lead)
          window['CO'].update(coaud)    

"""
    if event == 'Submit':
        win3 = sg.Window('Third Window',lead_only,alpha_channel=0).Finalize()
        win3.SetAlpha(1)
        window.close()
        window = win3
        by_exp = auditor_selector.check_exp(rec,val1['SUP'],val2['CAT'],avail_auditors)
        lead, coaud = auditor_selector.final_decision(by_exp,'NO')#val3['COAUD'])
        window['RES1'].update(lead)
"""
# Finish up by removing from the screen
window.close()