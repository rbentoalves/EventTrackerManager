import PySimpleGUI as sg
import myprocesses as mp
import mysubprocesses as msp
import myfunctions as mf
import re
import os
import pandas as pd



def main(incidents_file = "No File", tracker_incidents_file = "No File", site_list = ["No site list"], geography = "PT",
         date = "27-03-1996"):

    sg.theme('DarkAmber')  # Add a touch of color
    if incidents_file == "No File" or tracker_incidents_file == "No File":
        sg.popup('No files or site list available, please select them', no_titlebar=True)
        incidents_file, tracker_incidents_file, site_list, geography, date = mp.choose_incidents_files()
        if incidents_file == "No File" or tracker_incidents_file == "No File":
            return None
    else:
        print("Incidents file: " + incidents_file + "\nTracker Incidents file: " + tracker_incidents_file)

    dir = os.path.dirname(incidents_file)
    reportfiletemplate = dir + '/Info&Templates/Reporting_'+ geography +'_Sites_' + 'Template.xlsx'
    general_info_path = dir +  '/Info&Templates/General Info ' + geography + '.xlsx'


    #Reset Report Template to create new report
    reportfile = mf.reset_final_report(reportfiletemplate, date, geography)

    #Read Active and Closed Events
    df_list_active, df_list_closed = msp.read_approved_incidents(incidents_file, site_list, roundto=1)
    df_tracker_active, df_tracker_closed = msp.read_approved_tracker_inc(tracker_incidents_file, roundto=1)

    #Read sunrise and sunset hours
    df_info_sunlight = pd.read_excel(incidents_file, sheet_name='Info', engine="openpyxl")
    df_info_sunlight['Time of operation start'] = df_info_sunlight['Time of operation start'].dt.round(freq='s')
    df_info_sunlight['Time of operation end'] = df_info_sunlight['Time of operation end'].dt.round(freq='s')

    #Describe Incidents
    df_list_active = mf.describe_incidents(df_list_active, df_info_sunlight, active_events=True, tracker=False)
    df_list_closed = mf.describe_incidents(df_list_closed, df_info_sunlight, active_events=False, tracker=False)
    df_tracker_active = mf.describe_incidents(df_tracker_active, df_info_sunlight, active_events=True, tracker=True)
    df_tracker_closed = mf.describe_incidents(df_tracker_closed, df_info_sunlight, active_events=False, tracker=True)
    print(df_tracker_closed.columns)

    #Add Events to Report File
    mf.add_events_to_final_report(reportfile, df_list_active, df_list_closed, df_tracker_active, df_tracker_closed)

    #-------------------------------Analysis on Components Failures---------------------------------
    #Read and update the timestamps on the anaysis dataframes
    df_incidents_analysis, df_tracker_analysis = mf.read_analysis_df_and_correct_date(reportfiletemplate, date,
                                                                                      roundto=1)
    #Analysis of components failures
    df_incidents_analysis_final = msp.analysis_component_incidents(df_incidents_analysis, site_list, df_list_closed,
                                                                   df_list_active, df_info_sunlight)
    # Analysis of tracker failures
    df_tracker_analysis_final = msp.analysis_tracker_incidents(df_tracker_analysis, df_tracker_closed,
                                                               df_tracker_active, df_info_sunlight)
    #Add Analysis to excel file
    mf.add_analysis_to_reportfile(reportfile, df_incidents_analysis_final, df_tracker_analysis_final, df_info_sunlight)


    return reportfile


if __name__ == "__main__":
    try:
        main()
    except BaseException as e:
        print('Error: ')
        print(e)
        raise
