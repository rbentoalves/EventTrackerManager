import myfunctions as mf
import mysubprocesses as msp
import PySimpleGUI as sg
import os
import re
from datetime import datetime
import pandas as pd
import timeit
import math
import datetime as dt


def dmr_create_incidents_files(Alarm_report_path,irradiance_file_path, geography, date):
    dir = os.path.dirname(Alarm_report_path)
    #print("this is dir: " + dir)

    previous_date = datetime.strptime(date, '%Y-%m-%d') - dt.timedelta(days=1)
    prev_day = str(previous_date.day) if previous_date.day >= 10 else str(0) + str(previous_date.day)
    prev_month = str(previous_date.month) if previous_date.month >= 10 else str(0) + str(previous_date.month)


    Report_template_path = dir + '/Info&Templates/Reporting_' + geography + '_Sites_Template.xlsx'
    general_info_path = dir + '/Info&Templates/General Info ' + geography + '.xlsx'
    event_tracker_path = dir + '/Event Tracker/Event Tracker ' + geography + '.xlsx'
    previous_dmr_path = dir + '/Reporting_' + geography + '_Sites_' + prev_day + '-' + prev_month + '.xlsx'

    print(previous_dmr_path)

    #READ FILES AND EXTRACT RAW DATAFRAMES
    print('Reading Daily Alarm Report...')
    df_all, incidents_file, tracker_incidents_file, irradiance_file_data, prev_active_events, prev_active_tracker_events = mf.read_Daily_Alarm_Report(Alarm_report_path, irradiance_file_path, event_tracker_path, previous_dmr_path)
    print('Daily Alarm Report read!')
    print('newfile: ' + incidents_file)
    print('newtrackerfile: ' + tracker_incidents_file)
    print(df_all)
    print('Reading trackers info...')
    df_general_info, df_general_info_calc, all_component_data = mf.read_general_info(Report_template_path, general_info_path)
    print('Trackers info read!')

    # DIVIDE RAW DATAFRAMES INTO LIST OF DATAFRAMES BY SITE
    print('Creating incidents dataframes list...')
    site_list, df_list_active, df_list_closed = msp.create_dfs(df_all,min_dur = 1, roundto=1)
    print('Incidents dataframes list created')
    print('Creating tracker dataframes...')
    df_tracker_active, df_tracker_closed = msp.create_tracker_dfs(df_all,df_general_info_calc, roundto = 1)
    print('Tracker dataframes created')
    print('Please set time of operation')
    df_info_sunlight, final_irradiance_data = msp.read_time_of_operation(irradiance_file_data,Report_template_path, withmean = False)
    #df_info_sunlight = msp.set_time_of_operation(Report_template_path, site_list, date)
    print('Removing incidents occuring after sunset')
    df_list_closed = mf.remove_after_sunset_events(site_list, df_list_closed, df_info_sunlight)
    df_list_active = mf.remove_after_sunset_events(site_list, df_list_active, df_info_sunlight, active_df = True)
    df_tracker_closed = mf.remove_after_sunset_events(site_list, df_tracker_closed,df_info_sunlight , tracker = True)
    df_tracker_active = mf.remove_after_sunset_events(site_list, df_tracker_active, df_info_sunlight, active_df = True,
                                                      tracker = True)
    print('Adding component capacities')
    # ADD CAPACITIES TO DFS
    df_list_closed = mf.complete_dataset_capacity_data(df_list_closed, all_component_data)
    df_list_active = mf.complete_dataset_capacity_data(df_list_active, all_component_data)

    # JOIN INCIDENT TABLES AND DMR TABLES
    df_list_active = mf.complete_dataset_existing_incidents(df_list_active, prev_active_events)
    df_tracker_active = pd.concat([df_tracker_active, prev_active_tracker_events])

    # CREATE INCIDENTS FILE

    print('Creating Incindents file...')
    print(incidents_file)
    mf.add_incidents_to_excel(incidents_file,site_list,df_list_active,df_list_closed,df_info_sunlight, final_irradiance_data)
    print('Incindents file created!')
    print('Creating tracker incidents file...')
    mf.add_tracker_incidents_to_excel(tracker_incidents_file,df_tracker_active,df_tracker_closed, df_general_info)
    print('Tracker incindents file created!')




    return incidents_file, tracker_incidents_file, site_list, all_component_data

def choose_incidents_files():
    sg.theme('DarkAmber')  # Add a touch of color
    # All the stuff inside your window.
    layout = [[sg.Text('Enter date of report you want to analyse', pad=((2, 10), (2, 5)))],
              [sg.CalendarButton('Choose date', target='-CAL-', format="%Y-%m-%d"),
               sg.In(key='-CAL-', text_color='black', size=(16, 1), enable_events=True, readonly=True, visible=True)],
              [sg.Text('Choose Incidents file', pad=((0, 10), (10, 2)))],
              [sg.FileBrowse(target='-FILE-'),
               sg.In(key='-FILE-', text_color='black', size=(20, 1), enable_events=True, readonly=True, visible=True)],
              [sg.Text('Choose Tracker Incidents file', pad=((0, 10), (10, 2)))],
              [sg.FileBrowse(target='-TFILE-'),
               sg.In(key='-TFILE-', text_color='black', size=(20, 1), enable_events=True, readonly=True, visible=True)],
              [sg.Text('Enter geography ', pad=((0, 10), (10, 2)))],
              [sg.Combo(['AUS', 'ES', 'USA'], size=(4, 3), readonly=True, key='-GEO-', pad=((5, 10), (2, 10)))],
              [sg.Button('Submit'), sg.Exit()]]

    # Create the Window
    window = sg.Window('Daily Monitoring Report', layout, modal = True)
    # Event Loop to process "events" and get the "values" of the inputs
    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == 'Exit':  # if user closes window or clicks exit
            window.close()
            return "No File", "No File", ["No site list"], "PT", "27-03-1996"
            break
        if event == 'Submit':
            date = values['-CAL-']  # date is string
            date_object = datetime.strptime(date, '%Y-%m-%d')
            day = date_object.day
            month = date_object.month
            year = date_object.year
            if day < 10:
                strday = str(0) + str(day)
            else:
                strday = str(day)

            if month < 10:
                strmonth = str(0) + str(month)
            else:
                strmonth = str(month)

            stryear = str(year)
            date_to_test = strday + '-' + strmonth

            #Get files name+path
            incidents_file = values['-FILE-']
            tracker_incidents_file = values['-TFILE-']

            # Get file names
            incidents_file_name = os.path.basename(incidents_file)
            tracker_incidents_file_name = os.path.basename(tracker_incidents_file)

            #Get geography and dates
            geography_incidents_file_match = re.search(r'\-\w+\.', incidents_file_name)
            geography_incidents_file = geography_incidents_file_match.group()[3:-1]
            date_incidents_file_match = re.search(r'\d\d\-\d\d', incidents_file_name)
            date_incidents_file = date_incidents_file_match.group()

            geography_tracker_file_match = re.search(r'\-\w+\.', tracker_incidents_file_name)
            geography_tracker_file = geography_tracker_file_match.group()[3:-1]
            date_tracker_incidents_file_match = re.search(r'\d\d\-\d\d', tracker_incidents_file_name)
            date_tracker_incidents_file = date_tracker_incidents_file_match.group()

            geography = values['-GEO-']

            #print(strday)
            #print(strmonth)
            #print(stryear)
            #print(date_incidents_file)
            #print(date_tracker_incidents_file)
            #print(geography)
            #print(geography_incidents_file)
            #print(geography_tracker_file)


            if not date_incidents_file  == date_to_test:
                sg.popup('Incidents file is not the correct one, you chose the file from '
                      + date_incidents_file + ' and the following date: ' + date_to_test)

            elif not date_tracker_incidents_file == date_to_test:
                sg.popup('Incidents file is not the correct one, you chose the file from '
                      + date_tracker_incidents_file + ' and the following date: ' + date_to_test)

            elif not geography == geography_incidents_file:
                sg.popup('Selected Geography ' + geography + ' does not match geography from file '
                      + geography_incidents_file)

            elif not geography == geography_tracker_file:
                sg.popup('Selected Geography ' + geography + ' does not match geography from tracker file: '
                      + geography_tracker_file)

            elif not "Incidents" in incidents_file or "Tracker_Incidents" not in tracker_incidents_file:
                sg.popup('Files are not correct. \n' + '\n' + incidents_file
                         + ' has to be like: "Incidents01-07USA.xlsx" \n' + '\n' + tracker_incidents_file
                         + ' has to be like: "Tracker_Incidents01-07USA.xlsx"')

            elif "Incidents" in incidents_file and "Tracker_Incidents" in tracker_incidents_file:
                site_list = pd.read_excel(incidents_file, sheet_name='Info', engine="openpyxl")['Site'].tolist()
                sg.popup('Submitted files are correct: \n' + '\nIncidents file: \n' + incidents_file
                         + '\n' +'\nTracker Incidents file: \n' + tracker_incidents_file + '\n' + '\nSite list: \n'
                         + str(site_list) , no_titlebar=True)
                break


    window.close()

    return incidents_file, tracker_incidents_file, site_list, geography, date

def calculate_availability_period(site, incidents,component_data, budget_pr, df_all_irradiance,df_all_export,
                                  irradiance_threshold, date_start_str, date_end_str,granularity: float = 0.25):
    print(site)
    period = date_start_str + " to " + date_end_str
    active_events = False
    recalculate_value = True
    irradiance_incidents_corrected = {}

    # Get site info --------------------------------------------------------------------------------------------
    site_info = component_data.loc[component_data['Site'] == site]
    site_capacity = float(component_data.loc[component_data['Component'] == site]['Nominal Power DC'].values)
    budget_pr_site = budget_pr.loc[site, :]

    # Get site Incidents --------------------------------------------------------------------------------------------
    site_incidents = incidents.loc[incidents['Site Name'] == site]

    # Get site irradiance & export --------------------------------------------------------------------------------------------
    df_irradiance_site = df_all_irradiance.loc[:, df_all_irradiance.columns.str.contains(site + '|Timestamp')]
    df_export_site = df_all_export.loc[:,df_all_export.columns.str.contains(site + '|Timestamp')]

    # Get irradiance poa avg column and curated -----------------------------------------------------------------------
    actual_column, curated_column, data_gaps_proportion, poa_avg_column = mf.get_actual_irradiance_column(
        df_irradiance_site)

    # Get first timestamp under analysis and df from that timestamp onwards -------------------------------------------
    stime_index = next(i for i, v in enumerate(df_irradiance_site[poa_avg_column]) if v > irradiance_threshold)
    site_start_time = df_irradiance_site['Timestamp'][stime_index]

    df_irradiance_operation_site = df_irradiance_site.loc[df_irradiance_site['Timestamp'] >= site_start_time]
    df_export_operation_site = df_export_site.loc[df_export_site['Timestamp'] >= site_start_time]

    df_irradiance_operation_site['Day'] = [datetime.strptime(str(timestamp), '%Y-%m-%d %H:%M:%S').date() for timestamp
                                           in
                                           df_irradiance_operation_site['Timestamp']]
    df_export_operation_site['Day'] = [datetime.strptime(str(timestamp), '%Y-%m-%d %H:%M:%S').date() for timestamp
                                           in
                                           df_export_operation_site['Timestamp']]


    # Defined timeframe----------------------------------------------------
    if not date_start_str == 'None' and not date_end_str == 'None':

        # Get start time analysis
        date_start_avail_analysis = datetime.strptime(date_start_str, '%Y-%m-%d').date()
        timestamp_start_avail_analysis = datetime.strptime(date_start_str + " 00:00:00", '%Y-%m-%d %H:%M:%S')

        # Get end time analysis
        date_end_avail_analysis = datetime.strptime(date_end_str, '%Y-%m-%d').date()
        date_end_str_event = str(datetime.strptime(date_end_str, '%Y-%m-%d').date() + dt.timedelta(days=1))
        timestamp_end_avail_analysis = datetime.strptime(date_end_str_event + " 00:00:00", '%Y-%m-%d %H:%M:%S')

        # Get days list under analysis
        days_list = pd.date_range(start=date_start_avail_analysis, end=date_end_avail_analysis).date

    else:

        # Get days list under analysis
        days_list = sorted(list(set(df_irradiance_operation_site['Day'].to_list())))

        # Get start and end time analysis
        timestamp_start_avail_analysis = datetime.strptime(str(df_irradiance_operation_site['Timestamp'].to_list()[0]),
                                                           '%Y-%m-%d %H:%M:%S')
        timestamp_end_avail_analysis = datetime.strptime(str(df_irradiance_operation_site['Timestamp'].to_list()[-1]),
                                                         '%Y-%m-%d %H:%M:%S')


    # Get incidents in that period --------------------------------------------------------------------------------
    print(site_incidents[['ID', 'Event Start Time', 'Event End Time']])

    relevant_incidents = site_incidents.loc[~(site_incidents['Event Start Time'] > timestamp_end_avail_analysis) & ~(
        site_incidents['Event End Time'] < timestamp_start_avail_analysis)]
    #test
    """## print(relevant_incidents)
    ## print(relevant_incidents.loc[(relevant_incidents['Site Name'] == "LSBP - Bighorn") & (relevant_incidents['Related Component'] == "Inverter 65")])
    ## print(relevant_incidents[['Related Component', 'Event Start Time','Event End Time', "Duration (h)","Active Hours (h)", 'Energy Lost (MWh)' ]])"""


    # Get irradiance of period under analysis --------------------------------------------------------------------------
    irradiance_analysis = df_irradiance_operation_site.loc[
        (df_irradiance_operation_site['Day'] >= date_start_avail_analysis) & (
                df_irradiance_operation_site['Day'] <= date_end_avail_analysis)]

    export_analysis = df_export_operation_site.loc[
        (df_export_operation_site['Day'] >= date_start_avail_analysis) & (
            df_export_operation_site['Day'] <= date_end_avail_analysis)]

    actual_column, curated_column, data_gaps_proportion, poa_avg_column = mf.get_actual_irradiance_column(
            irradiance_analysis)

    if actual_column:
        df_irradiance_event_activeperiods = irradiance_analysis.loc[
            irradiance_analysis[actual_column] > irradiance_threshold]
    else:
        df_irradiance_event_activeperiods = irradiance_analysis.loc[
            irradiance_analysis[poa_avg_column] > irradiance_threshold]

    active_hours = df_irradiance_event_activeperiods.shape[0] * granularity

    site_active_hours_daily = {
        day: df_irradiance_event_activeperiods.loc[df_irradiance_event_activeperiods['Day'] == day].shape[
                 0] * granularity for day in days_list}

    site_active_hours_df_daily = pd.DataFrame.from_dict(site_active_hours_daily, orient='index',
                                                        columns=[site + ' Active Hours (h)'])

    site_active_hours_period = df_irradiance_event_activeperiods.shape[0] * granularity
    site_active_hours_df_period = pd.DataFrame({'Active Hours (h) ' + period: [site_active_hours_period]}, index=[site])
    # site_active_hours_df_period = pd.DataFrame({'Active Hours (h) ':[df_irradiance_event_activeperiods.shape[0]*granularity]}, index = [site])

    ## print(site_active_hours_df_daily)
    ## print(site_active_hours_df_period)


    # Correct Timestamps of incidents to timeframe of analysis ---------------------------------------------------------
    for index, row in relevant_incidents.iterrows():
        try:
            if math.isnan(row['Event End Time']):
                relevant_incidents.loc[index, 'Event End Time'] = timestamp_end_avail_analysis


            if row['Event Start Time'] < timestamp_start_avail_analysis:
                relevant_incidents.loc[index, 'Event Start Time'] = timestamp_start_avail_analysis

        except TypeError:
            if row['Event End Time'] > timestamp_end_avail_analysis:
                relevant_incidents.loc[index, 'Event End Time'] = timestamp_end_avail_analysis

            if row['Event Start Time'] < timestamp_start_avail_analysis:
                ## print('LOOK HERE')
                ## print(row[['Related Component', 'Event Start Time','Event End Time']])
                relevant_incidents.loc[index, 'Event Start Time'] = timestamp_start_avail_analysis

    ## print(relevant_incidents.loc[(relevant_incidents['Site Name'] == "LSBP - Bighorn") & (relevant_incidents['Related Component'] == "Inverter 65")])

    # Get incidents to keep unaltered
    incidents_unaltered = relevant_incidents.loc[~(relevant_incidents['Event Start Time'] == timestamp_start_avail_analysis) & ~(
        relevant_incidents['Event End Time'] == timestamp_end_avail_analysis)]

    # Get corrected incidents dict (overlappers) and then calculate real active hours and losses with that info --------
    corrected_incidents_dict_period = mf.correct_incidents_irradiance_for_overlapping_parents(relevant_incidents,
                                                                                              irradiance_analysis,
                                                                                              component_data,
                                                                                              recalculate_value)
    corrected_relevant_incidents = msp.calculate_activehours_energylost_incidents(relevant_incidents,
                                                                                  irradiance_analysis,export_analysis, budget_pr,
                                                                                  corrected_incidents_dict_period,
                                                                                  active_events, recalculate_value,
                                                                                  granularity)

    # Get corrected relevant incidents to concat with unaltered ones
    corrected_relevant_incidents = corrected_relevant_incidents.loc[
        (corrected_relevant_incidents['Event Start Time'] == timestamp_start_avail_analysis) | (
            corrected_relevant_incidents['Event End Time'] == timestamp_end_avail_analysis)]


    """# Get corrected incidents dict (overlappers) and then calculate real active hours and losses with that info --------
    corrected_incidents_dict_period = mf.correct_incidents_irradiance_for_overlapping_parents(relevant_incidents,
                                                                                              irradiance_analysis,
                                                                                              component_data,
                                                                                              recalculate_value)
    corrected_relevant_incidents = msp.calculate_activehours_energylost_incidents(relevant_incidents,
                                                                                  irradiance_analysis, export_analysis,
                                                                                  budget_pr,
                                                                                  corrected_incidents_dict_period,
                                                                                  active_events, recalculate_value,
                                                                                  granularity)"""

    ## print(corrected_relevant_incidents[['Related Component', 'Event Start Time','Event End Time', "Duration (h)","Active Hours (h)", 'Energy Lost (MWh)' ]])

    # Join corrected incidents and non-corrected incidents

    corrected_relevant_incidents = pd.concat([incidents_unaltered,corrected_relevant_incidents])


    # corrected_relevant_incidents = final_relevant_incidents
    # Calculate period availability-------------------------------------------------------------------------------------
    weighted_downtime = {}
    corrected_relevant_incidents['Weighted Downtime'] = ""
    for index, row in corrected_relevant_incidents.iterrows():
        capacity = row['Capacity Related Component']
        active_hours = row['Active Hours (h)']
        failure_mode = row['Failure Mode']
        if not failure_mode == "Curtailment":
            try:
                if capacity == float(0):
                    weighted_downtime_incident = 0
                elif math.isnan(active_hours):
                    weighted_downtime_incident = 0
                elif type(active_hours) == str:
                    weighted_downtime_incident = 0
                else:
                    weighted_downtime_incident = (capacity * active_hours) / site_capacity
            except TypeError:
                weighted_downtime_incident = 0
        else:
            weighted_downtime_incident = 0

        weighted_downtime[row['ID']] = weighted_downtime_incident
        corrected_relevant_incidents.loc[index, 'Weighted Downtime'] = weighted_downtime_incident

    weighted_downtime_df = pd.DataFrame.from_dict(weighted_downtime, orient='index',
                                                  columns=['Incident weighted downtime (h)'])
    print(weighted_downtime_df)

    total_weighted_downtime = weighted_downtime_df['Incident weighted downtime (h)'].sum()
    try:
        availability_period = ((site_active_hours_period - total_weighted_downtime) / site_active_hours_period)
    except ZeroDivisionError:
        availability_period = 0




    return availability_period, site_active_hours_period, corrected_relevant_incidents

def dmr_create_report_incomplete(incidents_path, tracker_incidents_path, date):

    a=1

    return a



