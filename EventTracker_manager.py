import PySimpleGUI as sg
import myprocesses as mp
import mysubprocesses as msp
import myfunctions as mf
import re
import os
import pandas as pd
import IPython
import datetime as dt
from datetime import datetime
import openpyxl
import xlsxwriter
import timeit
import math
import matplotlib as mpl
import matplotlib.pyplot as plt
import numpy as np
from scipy.optimize import curve_fit
import calendar
import seaborn as sns
import EventTracker_functions as etf

def main():
    sg.theme('DarkAmber')  # Add a touch of color
    # All the stuff inside your window.
    layout = [[sg.Text('Welcome to the Event Tracker Manager, what do you want to do?', pad=((2, 10), (2, 5)))],
              [sg.Text('Event Tracker Management', pad=((10, 10), (5, 5)))],
              [sg.Button('Create new Event Tracker',size = (25,1), pad = ((10,0),(4,4))), sg.Push()],
              [sg.Button('Update Event Tracker',size = (25,1), pad = ((10,0),(4,4))), sg.Push() ],
              [sg.Text('Create Reports', pad=((10, 10), (5, 5)))],
              [sg.Button('Event Tracker', size=(25, 1), pad=((10, 0), (4, 4))), sg.Push()],
              [sg.Button('Underperformance Report',size = (25,1), pad = ((10,0),(4,4))), sg.Push()],
              [sg.Button('Monday.com files',size = (25,1), pad = ((10,0),(4,4))), sg.Push()],
              [sg.Push(), sg.Exit()]]

    # Create the Window
    window = sg.Window('Event Tracker Manager', layout)
    # Event Loop to process "events" and get the "values" of the inputs
    while True:
        event, values = window.read(timeout=100)

        if event == sg.WIN_CLOSED or event == 'Exit':  # if user closes window or clicks exit
            break
        if event == 'Create new Event Tracker':

            period_list = ['monthly', 'ytd']

            # Data gathering

            # <editor-fold desc="Get inputs, files necessary to analysis">
            # Get input of critical information for update, dates and file locations
            source_folder,source, geography, geopgraphy_folder, recalculate_value = etf.new_event_tracker_from_input()


            if source_folder == "None":
                continue

            """print("Start date: ", date_start, "\n End date: ", date_end, "\n ET: ", event_tracker_path,
                  "\n DMR folder: ", dmr_folder)
"""
            # Get file paths to add
            print("Looking for files to add...")
            all_irradiance_file, all_export_file, general_info_path = \
                etf.get_files_to_add(0, 0, geopgraphy_folder, geography, no_update=True)


            dest_file = geopgraphy_folder + '/Event Tracker/Event Tracker ' + geography + '.xlsx'
            folder_img = geopgraphy_folder + '/Event Tracker/images'

            """print("All Irradiance file: ", all_irradiance_file, "\n Irradiance files: ", irradiance_files,
                  "\n All Export file: ", all_export_file, "\n Export files: ", export_files,
                  "\n Report files: ", report_files,"\n General info path: ", general_info_path)
"""
            # </editor-fold>

            # <editor-fold desc="Read dump files - Irradiance & Export">
            # Update export and irradiance dump files

            df_all_export = pd.read_excel(all_export_file, engine='openpyxl')
            df_all_irradiance = pd.read_excel(all_irradiance_file, engine='openpyxl')

            df_all_export['Timestamp'] = [datetime.strptime(str(timestamp), '%Y-%m-%d %H:%M:%S') for timestamp
                                          in df_all_export['Timestamp']]

            df_all_irradiance['Timestamp'] = [datetime.strptime(str(timestamp), '%Y-%m-%d %H:%M:%S') for
                                              timestamp in df_all_irradiance['Timestamp']]
            # </editor-fold>

            # <editor-fold desc="Get site list">
            # Get site list from irradiance dataframe
            site_list = list(set([re.search(r'\[.+\]', site).group().replace('[', "").replace(']', "") for site in
                                  df_all_irradiance.loc[:,
                                  df_all_irradiance.columns.str.contains('Irradiance')].columns]))
            site_list = [mf.correct_site_name(site) for site in site_list]
            # </editor-fold>

            # <editor-fold desc="Get info dataframes necessary">
            # Get info dataframes necessary
            print("Reading general info files and creating dataframes...")
            component_data, tracker_data, fmeca_data, site_capacities, fleet_capacity, budget_irradiance, \
            budget_pr, budget_export = etf.get_general_info_dataframes(general_info_path)

            # Correct unnamed columns
            fmeca_data = fmeca_data.loc[:, ~fmeca_data.columns.str.contains('^Unnamed')]
            fmeca_data = fmeca_data.dropna(thresh=8)

            # </editor-fold>

            # <editor-fold desc="Get incidents dataframes">



            # Get incidents' dataframes
            print("Reading incident and Event Tracker files and creating dataframes...")
            df_all = pd.read_excel(source,
                                   sheet_name=['Active Events', 'Closed Events', 'Active tracker incidents',
                                               'Closed tracker incidents'], engine='openpyxl')

            df_active_eventtracker = mf.match_df_to_event_tracker(df_all['Active Events'],component_data,fmeca_data)
            df_closed_eventtracker = mf.match_df_to_event_tracker(df_all['Closed Events'],component_data,fmeca_data)
            df_active_eventtracker_trackers = mf.match_df_to_event_tracker(df_all['Active tracker incidents'],tracker_data,fmeca_data)
            df_closed_eventtracker_trackers = mf.match_df_to_event_tracker(df_all['Closed tracker incidents'],tracker_data,fmeca_data)



            # Get final dfs to add
            print("Creating pre-treatment final dataframes of the Event tracker...")
            final_df_to_add = {'Closed Events': df_closed_eventtracker,
                               'Closed tracker incidents': df_closed_eventtracker_trackers,
                               'Active Events': df_active_eventtracker,
                               'Active tracker incidents': df_active_eventtracker_trackers,
                               'FMECA': fmeca_data}
            # </editor-fold>

            # Data calculations and handling

            # Create all component incidents df
            incidents = pd.concat([final_df_to_add['Active Events'], final_df_to_add['Closed Events']])
            incidents['Event Start Time'] = [datetime.strptime(str(timestamp), '%Y-%m-%d %H:%M:%S') for timestamp
                                           in
                                           incidents['Event Start Time']]

            # Create FMECA aux tables - can be moved to file creation
            dict_fmeca_shapes = msp.create_fmeca_dataframes_for_validation(fmeca_data)

            # Correct active hours and energy loss to account for overlapping incidents
            print("Correcting overlapping events...")
            corrected_incidents_dict = mf.correct_incidents_irradiance_for_overlapping_parents(incidents,
                                                                                               df_all_irradiance,
                                                                                               component_data,
                                                                                               recalculate_value)

            # Calculate active hours and energy lost with correction for overlapping parents
            print("Creating final dataframes of the Event tracker...")
            final_df_to_add = etf.calculate_active_hours_and_energy_lost(final_df_to_add, corrected_incidents_dict,
                                                                         df_all_irradiance, df_all_export, budget_pr,
                                                                         irradiance_threshold=20)

            final_df_to_add['Closed Events']['Event End Time'] = [datetime.strptime(str(timestamp), '%Y-%m-%d %H:%M:%S') for timestamp
                                             in
                                             final_df_to_add['Closed Events']['Event End Time']]

            incidents = pd.concat([final_df_to_add['Active Events'], final_df_to_add['Closed Events']])

            incidents['Event Start Time'] = [datetime.strptime(str(timestamp), '%Y-%m-%d %H:%M:%S') for timestamp
                                             in
                                             incidents['Event Start Time']]


            # File Creation

            # <editor-fold desc="Calculate Availability & PR">
            # File Creation - step 1 get relevant availability and pr data
            print("Creating Event Tracker...")

            print("Calculating Availability in periods")

            # <editor-fold desc="Calculate availability per period">
            availability_fleet_per_period = {}
            active_hours_fleet_per_period = {}
            incidents_corrected_fleet_period_per_period = {}

            for period in period_list:
                availability_period_df, activehours_period_df, incidents_corrected_period, date_range = \
                    etf.calculate_availability_in_period(incidents, period, component_data, df_all_irradiance,
                                                         df_all_export, budget_pr, irradiance_threshold=20,
                                                         timestamp=15)

                availability_fleet_per_period[period] = availability_period_df
                active_hours_fleet_per_period[period] = activehours_period_df
                incidents_corrected_fleet_period_per_period[period] = incidents_corrected_period
            # </editor-fold>

            print("Calculating Performance KPIs in periods")
            # <editor-fold desc="Calculate site pr per period">
            performance_fleet_per_period = {}

            for period in period_list:
                incidents_period = incidents_corrected_fleet_period_per_period[period]
                availability_period = availability_fleet_per_period[period]

                data_period_df = etf.calculate_pr_in_period(incidents_period, availability_period, period,
                                                            component_data,
                                                            df_all_irradiance, df_all_export, budget_pr, budget_export,
                                                            budget_irradiance, irradiance_threshold=20, timestamp=15)

                performance_fleet_per_period[period] = data_period_df.sort_index()
            # </editor-fold>
            # </editor-fold>

            # <editor-fold desc="Create Graphs and other visuals">
            # File Creation - step 2 create graphs & visuals
            print("Creating graphs and visual aids")
            graphs = {}
            for period in period_list:
                period_graph = etf.availability_visuals(availability_fleet_per_period, period, folder_img)
                graphs[period] = period_graph
            # </editor-fold>

            # <editor-fold desc="Create file">
            # File Creation - step 3 actually create file
            print("Creating file...")

            etf.create_event_tracker_file_all(final_df_to_add, dest_file, performance_fleet_per_period, site_capacities,
                                              dict_fmeca_shapes)
            # </editor-fold>

            if dest_file:
                event, values = sg.Window('Choose an option', [[sg.Text('Process complete, open file?')],
                                                               [sg.Button('Yes'), sg.Button('Cancel')]]).read(
                    close=True)

                if event == 'Yes':
                    command = 'start "EXCEL.EXE" "' + str(dest_file) + '"'
                    os.system(command)
                    break
                else:
                    break



            sg.popup_cancel('Under development')

        if event == 'Update Event Tracker':

            period_list = ['mtd','ytd']

            # <editor-fold desc="Get inputs, files and dataframes necessary to analysis">
            #Get input of critical information for update, dates and file locations
            date_start, date_end, event_tracker_path, dmr_folder, geography,toggle_updt, recalculate_value \
                = etf.update_event_tracker_input()

            if event_tracker_path == "None":
                continue

            """print("Start date: ", date_start, "\n End date: ", date_end, "\n ET: ", event_tracker_path,
                  "\n DMR folder: ", dmr_folder)
"""
            #Get file paths to add
            print("Looking for files to add...")
            report_files, irradiance_files, export_files, all_irradiance_file, all_export_file,general_info_path = \
                etf.get_files_to_add(date_start, date_end, dmr_folder, geography)

            folder = os.path.dirname(report_files[0])
            dest_file = folder + '/Event Tracker/Event Tracker ' + geography + '_Final.xlsx'
            folder_img = folder + '/Event Tracker/images'

            """print("All Irradiance file: ", all_irradiance_file, "\n Irradiance files: ", irradiance_files,
                  "\n All Export file: ", all_export_file, "\n Export files: ", export_files,
                  "\n Report files: ", report_files,"\n General info path: ", general_info_path)
"""
            # </editor-fold>

            # <editor-fold desc="Update dump files - Irradiance & Export">
            #Update export and irradiance dump files
            if toggle_updt == True:
                print("Updating dump files of irradiance and export...")
                df_all_irradiance = mf.update_dump_file(irradiance_files, all_irradiance_file)
                df_all_export = mf.update_dump_file(export_files, all_export_file, data_type="Energy Exported")


            else:
                print("Reading dump files of irradiance and export...")
                df_all_export = pd.read_excel(all_export_file, engine='openpyxl')
                df_all_irradiance = pd.read_excel(all_irradiance_file, engine='openpyxl')

                df_all_export['Timestamp'] = [datetime.strptime(str(timestamp), '%Y-%m-%d %H:%M:%S') for timestamp
                                              in df_all_export['Timestamp']]

                df_all_irradiance['Timestamp'] = [datetime.strptime(str(timestamp), '%Y-%m-%d %H:%M:%S') for
                                                  timestamp in df_all_irradiance['Timestamp']]
            # </editor-fold>

            # <editor-fold desc="Get site list">
            #Get site list from irradiance dataframe
            site_list = list(set([re.search(r'\[.+\]', site).group().replace('[', "").replace(']', "") for site in
                                  df_all_irradiance.loc[:,
                                  df_all_irradiance.columns.str.contains('Irradiance')].columns]))
            site_list = [mf.correct_site_name(site) for site in site_list]
            # </editor-fold>

            # <editor-fold desc="Read info and get incidents dataframes">
            #Get info dataframes necessary
            print("Reading general info files and creating dataframes...")
            component_data, tracker_data, fmeca_data, site_capacities,fleet_capacity, budget_irradiance,\
                                 budget_pr, budget_export = etf.get_general_info_dataframes(general_info_path)

            #Get incidents' dataframes
            print("Reading incident and Event Tracker files and creating dataframes...")
            dfs_to_add, dfs_event_tracker, fmeca_data = msp.get_dataframes_to_add_to_EventTracker(report_files,
                                                                                                  event_tracker_path,
                                                                                                  fmeca_data,
                                                                                                  component_data,
                                                                                                  tracker_data)

            # Get final dfs to add
            print("Creating pre-treatment final dataframes of the Event tracker...")
            final_df_to_add = msp.get_final_dataframes_to_add_to_EventTracker(dfs_to_add, dfs_event_tracker, fmeca_data)

            # Create all component incidents df
            incidents = pd.concat([final_df_to_add['Active Events'], final_df_to_add['Closed Events']])

            # </editor-fold>

            # Correct active hours and energy loss to account for overlapping incidents
            print("Correcting overlapping events...")
            corrected_incidents_dict = mf.correct_incidents_irradiance_for_overlapping_parents(incidents,
                                                                                               df_all_irradiance,
                                                                                               component_data,
                                                                                               recalculate_value)

            # Create FMECA aux tables - can be moved to file creation
            dict_fmeca_shapes = msp.create_fmeca_dataframes_for_validation(fmeca_data)


            # Calculate active hours and energy lost with correction for overlapping parents
            print("Creating final dataframes of the Event tracker...")
            final_df_to_add = etf.calculate_active_hours_and_energy_lost(final_df_to_add,corrected_incidents_dict,
                                                                     df_all_irradiance, df_all_export,budget_pr,
                                                                         irradiance_threshold = 20)

            incidents = pd.concat([final_df_to_add['Active Events'], final_df_to_add['Closed Events']])




            """Update database --------------------------------------------------------------------------------------
            """


            #File Creation
            print("Creating Event Tracker...")

            # <editor-fold desc="Calculate Availability & PR">
            print("Calculating Availability in periods")

            # <editor-fold desc="Calculate availability per period">
            availability_fleet_per_period = {}
            active_hours_fleet_per_period = {}
            incidents_corrected_fleet_period_per_period = {}

            for period in period_list:
                availability_period_df, activehours_period_df, incidents_corrected_period, date_range = \
                    etf.calculate_availability_in_period(incidents, period, component_data, df_all_irradiance,
                                                    df_all_export,budget_pr,irradiance_threshold = 20,timestamp = 15)

                availability_fleet_per_period[period] = availability_period_df
                active_hours_fleet_per_period[period] = activehours_period_df
                incidents_corrected_fleet_period_per_period[period] = incidents_corrected_period
            # </editor-fold>


            print("Calculating Performance KPIs in periods")
            # <editor-fold desc="Calculate site pr per period">
            performance_fleet_per_period = {}

            for period in period_list:
                incidents_period = incidents_corrected_fleet_period_per_period[period]
                availability_period = availability_fleet_per_period[period]

                data_period_df = etf.calculate_pr_in_period(incidents_period,availability_period, period, component_data,
                                                            df_all_irradiance,df_all_export,budget_pr,budget_export,
                                                            budget_irradiance,irradiance_threshold = 20,timestamp = 15)

                performance_fleet_per_period[period] = data_period_df.sort_index()
            # </editor-fold>
            # </editor-fold>

            # <editor-fold desc="Create Graphs & Visuals">
            #File Creation - step 2 create graphs & visuals
            print("Creating graphs and visual aids")
            graphs = {}
            for period in period_list:
                period_graph = etf.availability_visuals(availability_fleet_per_period,period,folder_img)
                graphs[period] = period_graph
            # </editor-fold>

            # <editor-fold desc="Create file">
            #File Creation - step 3 actually create file
            print("Creating file...")

            etf.create_event_tracker_file_all(final_df_to_add, dest_file,performance_fleet_per_period, site_capacities,
                                  dict_fmeca_shapes)
            # </editor-fold>


            if dest_file:
                event, values = sg.Window('Choose an option', [[sg.Text('Process complete, open file?')],
                                                               [sg.Button('Yes'), sg.Button('Cancel')]]).read(close=True)

                if event == 'Yes':
                    command = 'start "EXCEL.EXE" "' + str(dest_file) + '"'
                    os.system(command)
                    break
                else:
                    break

        if event == 'Event Tracker':

            period_list = ['mtd', 'ytd']

            #Data gathering

            # <editor-fold desc="Get inputs, files necessary to analysis">
            # Get input of critical information for update, dates and file locations
            source_folder, geography,geopgraphy_folder, recalculate_value = etf.event_tracker_from_input()

            if source_folder == "None":
                continue


            """print("Start date: ", date_start, "\n End date: ", date_end, "\n ET: ", event_tracker_path,
                  "\n DMR folder: ", dmr_folder)
"""
            # Get file paths to add
            print("Looking for files to add...")
            all_irradiance_file, all_export_file, general_info_path = \
                etf.get_files_to_add(0, 0, geopgraphy_folder, geography, no_update = True)

            event_tracker_file_path = geopgraphy_folder + '/Event Tracker/Event Tracker ' + geography + '.xlsx'
            dest_file = geopgraphy_folder + '/Event Tracker/Event Tracker ' + geography + '_Final.xlsx'
            folder_img = geopgraphy_folder + '/Event Tracker/images'

            """print("All Irradiance file: ", all_irradiance_file, "\n Irradiance files: ", irradiance_files,
                  "\n All Export file: ", all_export_file, "\n Export files: ", export_files,
                  "\n Report files: ", report_files,"\n General info path: ", general_info_path)
"""
            # </editor-fold>

            # <editor-fold desc="Read dump files - Irradiance & Export">
            # Update export and irradiance dump files

            df_all_export = pd.read_excel(all_export_file, engine='openpyxl')
            df_all_irradiance = pd.read_excel(all_irradiance_file, engine='openpyxl')

            df_all_export['Timestamp'] = [datetime.strptime(str(timestamp), '%Y-%m-%d %H:%M:%S') for timestamp
                                          in df_all_export['Timestamp']]

            df_all_irradiance['Timestamp'] = [datetime.strptime(str(timestamp), '%Y-%m-%d %H:%M:%S') for
                                              timestamp in df_all_irradiance['Timestamp']]
            # </editor-fold>

            # <editor-fold desc="Get site list">
            # Get site list from irradiance dataframe
            site_list = list(set([re.search(r'\[.+\]', site).group().replace('[', "").replace(']', "") for site in
                                  df_all_irradiance.loc[:,
                                  df_all_irradiance.columns.str.contains('Irradiance')].columns]))
            site_list = [mf.correct_site_name(site) for site in site_list]
            # </editor-fold>

            # <editor-fold desc="Get info dataframes necessary">
            # Get info dataframes necessary
            print("Reading general info files and creating dataframes...")
            component_data, tracker_data, fmeca_data, site_capacities, fleet_capacity, budget_irradiance, \
            budget_pr, budget_export = etf.get_general_info_dataframes(general_info_path)
            # </editor-fold>

            # <editor-fold desc="Get incidents dataframes">
            # Get incidents' dataframes
            print("Reading incident and Event Tracker files and creating dataframes...")
            df_all = pd.read_excel(event_tracker_file_path,
                                   sheet_name=['Active Events', 'Closed Events', 'Active tracker incidents',
                                               'Closed tracker incidents'], engine='openpyxl')

            df_active_eventtracker = df_all['Active Events']
            df_closed_eventtracker = df_all['Closed Events']
            df_active_eventtracker_trackers = df_all['Active tracker incidents']
            df_closed_eventtracker_trackers = df_all['Closed tracker incidents']

            # Correct unnamed columns
            fmeca_data = fmeca_data.loc[:, ~fmeca_data.columns.str.contains('^Unnamed')]
            fmeca_data = fmeca_data.dropna(thresh=8)

            # Get final dfs to add
            print("Creating pre-treatment final dataframes of the Event tracker...")
            final_df_to_add = {'Closed Events': df_closed_eventtracker,
                                'Closed tracker incidents': df_closed_eventtracker_trackers,
                                'Active Events': df_active_eventtracker,
                                'Active tracker incidents': df_active_eventtracker_trackers,
                               'FMECA': fmeca_data}
            # </editor-fold>


            #Data calculations and handling
            # Create FMECA aux tables - can be moved to file creation
            dict_fmeca_shapes = msp.create_fmeca_dataframes_for_validation(fmeca_data)

            # Create all component incidents df
            incidents = pd.concat([final_df_to_add['Active Events'], final_df_to_add['Closed Events']])

            # Correct active hours and energy loss to account for overlapping incidents
            print("Correcting overlapping events...")
            corrected_incidents_dict = mf.correct_incidents_irradiance_for_overlapping_parents(incidents,
                                                                                               df_all_irradiance,
                                                                                               component_data,
                                                                                               recalculate_value)

            # Calculate active hours and energy lost with correction for overlapping parents
            print("Creating final dataframes of the Event tracker...")
            final_df_to_add = etf.calculate_active_hours_and_energy_lost(final_df_to_add, corrected_incidents_dict,
                                                                         df_all_irradiance, df_all_export, budget_pr,
                                                                         irradiance_threshold=20)

            incidents = pd.concat([final_df_to_add['Active Events'], final_df_to_add['Closed Events']])

            # File Creation

            # <editor-fold desc="Calculate Availability & PR">
            # File Creation - step 1 get relevant availability and pr data
            print("Creating Event Tracker...")

            print("Calculating Availability in periods")

            # <editor-fold desc="Calculate availability per period">
            availability_fleet_per_period = {}
            active_hours_fleet_per_period = {}
            incidents_corrected_fleet_period_per_period = {}

            for period in period_list:
                availability_period_df, activehours_period_df, incidents_corrected_period,date_range = \
                    etf.calculate_availability_in_period(incidents, period, component_data, df_all_irradiance,
                                                         df_all_export, budget_pr, irradiance_threshold=20,
                                                         timestamp=15)

                availability_fleet_per_period[period] = availability_period_df
                active_hours_fleet_per_period[period] = activehours_period_df
                incidents_corrected_fleet_period_per_period[period] = incidents_corrected_period
            # </editor-fold>

            print("Calculating Performance KPIs in periods")
            # <editor-fold desc="Calculate site pr per period">
            performance_fleet_per_period = {}

            for period in period_list:
                incidents_period = incidents_corrected_fleet_period_per_period[period]
                availability_period = availability_fleet_per_period[period]

                data_period_df = etf.calculate_pr_in_period(incidents_period, availability_period, period,
                                                            component_data,
                                                            df_all_irradiance, df_all_export, budget_pr, budget_export,
                                                            budget_irradiance, irradiance_threshold=20, timestamp=15)

                performance_fleet_per_period[period] = data_period_df.sort_index()
            # </editor-fold>
            # </editor-fold>

            # <editor-fold desc="Create Graphs and other visuals">
            # File Creation - step 2 create graphs & visuals
            print("Creating graphs and visual aids")
            graphs = {}
            for period in period_list:
                period_graph = etf.availability_visuals(availability_fleet_per_period, period, folder_img)
                graphs[period] = period_graph
            # </editor-fold>

            # <editor-fold desc="Create file">
            # File Creation - step 3 actually create file
            print("Creating file...")

            etf.create_event_tracker_file_all(final_df_to_add, dest_file, performance_fleet_per_period, site_capacities,
                                              dict_fmeca_shapes)
            # </editor-fold>


            if dest_file:
                event, values = sg.Window('Choose an option', [[sg.Text('Process complete, open file?')],
                                                               [sg.Button('Yes'), sg.Button('Cancel')]]).read(close=True)

                if event == 'Yes':
                    command = 'start "EXCEL.EXE" "' + str(dest_file) + '"'
                    os.system(command)
                    break
                else:
                    break

        if event == 'Underperformance Report':

            # Data gathering

            # <editor-fold desc="Get inputs, files necessary to analysis">
            # Get input of critical information for update, dates and file locations
            source_folder, geography, geopgraphy_folder, recalculate_value, period_list, level, irradiance_threshold \
                = etf.underperformance_report_input()

            #print(source_folder, "\n" , geography, "\n" , geopgraphy_folder, "\n" ,recalculate_value,"\n" , period_list)


            """print("Start date: ", date_start, "\n End date: ", date_end, "\n ET: ", event_tracker_path,
                  "\n DMR folder: ", dmr_folder)
            """
            # Get file paths to add
            print("Looking for files to add...")
            all_irradiance_file, all_export_file, general_info_path = \
                etf.get_files_to_add(0, 0, geopgraphy_folder, geography, no_update=True)

            event_tracker_file_path = geopgraphy_folder + '/Event Tracker/Event Tracker ' + geography + '.xlsx'
            folder_img = geopgraphy_folder + '/Event Tracker/Underperformance Reports/images'

            print("All Irradiance file: ", all_irradiance_file,
                  "\n All Export file: ", all_export_file,
                  "\n Event Tracker path: ", event_tracker_file_path,
                  "\n General info path: ", general_info_path)



            # </editor-fold>

            # <editor-fold desc="Read dump files - Irradiance & Export">
            # Update export and irradiance dump files

            df_all_export = pd.read_excel(all_export_file, engine='openpyxl')
            df_all_irradiance = pd.read_excel(all_irradiance_file, engine='openpyxl')

            df_all_export['Timestamp'] = [datetime.strptime(str(timestamp), '%Y-%m-%d %H:%M:%S') for timestamp
                                          in df_all_export['Timestamp']]

            df_all_irradiance['Timestamp'] = [datetime.strptime(str(timestamp), '%Y-%m-%d %H:%M:%S') for
                                              timestamp in df_all_irradiance['Timestamp']]
            # </editor-fold>

            # <editor-fold desc="Get site list">
            # Get site list from irradiance dataframe
            site_list = list(set([re.search(r'\[.+\]', site).group().replace('[', "").replace(']', "") for site in
                                  df_all_irradiance.loc[:,
                                  df_all_irradiance.columns.str.contains('Irradiance')].columns]))
            site_list = [mf.correct_site_name(site) for site in site_list]
            # </editor-fold>

            # <editor-fold desc="Get info dataframes necessary">
            # Get info dataframes necessary
            print("Reading general info files and creating dataframes...")
            component_data, tracker_data, fmeca_data, site_capacities, fleet_capacity, budget_irradiance, \
            budget_pr, budget_export = etf.get_general_info_dataframes(general_info_path)
            # </editor-fold>

            # <editor-fold desc="Get incidents dataframes from Event Tracker">
            # Get incidents' dataframes
            print("Reading incident and Event Tracker files and creating dataframes...")
            df_all = pd.read_excel(event_tracker_file_path,
                                   sheet_name=['Active Events', 'Closed Events', 'Active tracker incidents',
                                               'Closed tracker incidents'], engine='openpyxl')

            df_active_eventtracker = df_all['Active Events']
            df_closed_eventtracker = df_all['Closed Events']
            df_active_eventtracker_trackers = df_all['Active tracker incidents']
            df_closed_eventtracker_trackers = df_all['Closed tracker incidents']

            # Correct unnamed columns
            fmeca_data = fmeca_data.loc[:, ~fmeca_data.columns.str.contains('^Unnamed')]
            fmeca_data = fmeca_data.dropna(thresh=8)

            # Get final dfs to add
            print("Creating pre-treatment final dataframes of the Event tracker...")
            final_df_to_add = {'Closed Events': df_closed_eventtracker,
                               'Closed tracker incidents': df_closed_eventtracker_trackers,
                               'Active Events': df_active_eventtracker,
                               'Active tracker incidents': df_active_eventtracker_trackers,
                               'FMECA': fmeca_data}
            # </editor-fold>

            # Data calculations and handling
            # Create FMECA aux tables - can be moved to file creation
            dict_fmeca_shapes = msp.create_fmeca_dataframes_for_validation(fmeca_data)

            # Create all component incidents df
            incidents = pd.concat([final_df_to_add['Active Events'], final_df_to_add['Closed Events']])

            if level == "All":
                pass
            elif level == 'Inverter only':
                incidents = incidents.loc[(incidents['Related Component'].str.contains("INV")) | (incidents['Related Component'].str.contains("Inverter"))]
            elif level == 'Inverter level':
                incidents = incidents.loc[~(incidents['Related Component'].str.contains("CB")) | ~(
                    incidents['Related Component'].str.contains("DC"))| ~(
                    incidents['Related Component'].str.contains("String"))]

            """# Correct active hours and energy loss to account for overlapping incidents
            print("Correcting overlapping events...")
            corrected_incidents_dict = mf.correct_incidents_irradiance_for_overlapping_parents(incidents,
                                                                                               df_all_irradiance,
                                                                                               component_data,
                                                                                               recalculate_value)

            # Calculate active hours and energy lost with correction for overlapping parents
            print("Creating final dataframes of the Event tracker...")
            final_df_to_add = etf.calculate_active_hours_and_energy_lost(final_df_to_add, corrected_incidents_dict,
                                                                         df_all_irradiance, df_all_export, budget_pr,
                                                                         irradiance_threshold=20)

            incidents = pd.concat([final_df_to_add['Active Events'], final_df_to_add['Closed Events']])"""

            # File Creation

            # <editor-fold desc="Calculate Availability & PR">
            # File Creation - step 1 get relevant availability and pr data
            print("Creating Event Tracker...")

            print("Calculating Availability in periods")

            # <editor-fold desc="Calculate availability per period">
            availability_fleet_per_period = {}
            active_hours_fleet_per_period = {}
            incidents_corrected_fleet_period_per_period = {}

            for period in period_list:
                availability_period_df, activehours_period_df, incidents_corrected_period, date_range = \
                    etf.calculate_availability_in_period(incidents, period, component_data, df_all_irradiance,
                                                         df_all_export, budget_pr, irradiance_threshold=irradiance_threshold,
                                                         timestamp=15)

                availability_fleet_per_period[period] = availability_period_df
                active_hours_fleet_per_period[period] = activehours_period_df
                incidents_corrected_fleet_period_per_period[period] = incidents_corrected_period
            # </editor-fold>

            print("Calculating Performance KPIs in periods")
            # <editor-fold desc="Calculate site pr per period">
            performance_fleet_per_period = {}

            for period in period_list:
                incidents_period = incidents_corrected_fleet_period_per_period[period]
                availability_period = availability_fleet_per_period[period]

                data_period_df = etf.calculate_pr_in_period(incidents_period, availability_period, period,
                                                            component_data,
                                                            df_all_irradiance, df_all_export, budget_pr, budget_export,
                                                            budget_irradiance, irradiance_threshold=irradiance_threshold, timestamp=15)

                performance_fleet_per_period[period] = data_period_df.sort_index()
            # </editor-fold>


            # </editor-fold>

            # <editor-fold desc="Create Graphs and other visuals - not on underperformance report">
            # File Creation - step 2 create graphs & visuals
            """print("Creating graphs and visual aids")
            graphs = {}
            for period in period_list:
                period_graph = etf.availability_visuals(availability_fleet_per_period, period, folder_img)
                graphs[period] = period_graph"""
            # </editor-fold> -  -

            # <editor-fold desc="Create file">
            # File Creation - step 3 actually create file
            print("Creating file...")

            underperformance_dest_file = geopgraphy_folder + \
                                         '/Event Tracker/Underperformance Reports/Underperformance Report ' \
                                         + geography + "_" + date_range + "_" + level + '_irr' + str(irradiance_threshold) + '.xlsx'

            etf.create_underperformance_report(underperformance_dest_file,incidents_corrected_period, performance_fleet_per_period)

            # </editor-fold>

            if underperformance_dest_file:
                event, values = sg.Window('Choose an option', [[sg.Text('Process complete, open file?')],
                                                               [sg.Button('Yes'), sg.Button('Cancel')]]).read(
                    close=True)

                if event == 'Yes':
                    command = 'start "EXCEL.EXE" "' + str(underperformance_dest_file) + '"'
                    os.system(command)
                    break
                else:
                    break

            sg.popup_cancel('Under development')

            """try:
                dmr_report = dmrprocess2.main(incidents_file, tracker_incidents_file, site_list, geography, date)
                
            except NameError:
                dmr_report = dmrprocess2.main()
                #dmr_report = dmrprocess2_fromscratch.main()
            if dmr_report:
                event, values = sg.Window('Choose an option', [[sg.Text('Process complete, open file?')],
                                                               [sg.Button('Yes'), sg.Button('Cancel')]]).read(close=True)

                if event == 'Yes':
                    command = 'start "EXCEL.EXE" "' + str(dmr_report) + '"'
                    os.system(command)
                    break
                else:
                    sg.popup_cancel('User aborted')
                    break
"""

        if event == "Monday.com files":
            date_start, date_end, event_tracker_folder, geography = etf.mondaycom_file_input()

            date_list = date_list = pd.date_range(date_start, date_end, freq = 'd')
            first_timestamp = datetime.strptime(date_start + " 00:00:00", '%Y-%m-%d %H:%M:%S')
            last_timestamp = datetime.strptime(date_end + " 23:59:59", '%Y-%m-%d %H:%M:%S')
            start_of_month_timestamp = datetime.strptime(str(first_timestamp.year) + "-" + str(first_timestamp.month) + "-01" + " 00:00:00", '%Y-%m-%d %H:%M:%S')

            event_tracker_file_path = event_tracker_folder + '/Event Tracker ' + geography + '.xlsx'
            monday_folder = event_tracker_folder + '/Monday.com Updates'


            if date_start != date_end:
                monday_day_folder = monday_folder + '/' + date_start + "to" + date_end
                dest_file_new_active_events = monday_day_folder +'/New_Events_' + date_start + "to" + date_end +'.xlsx'
                dest_file_new_closed_events = monday_day_folder +'/New_Closed_Events_' + date_start + "to" + date_end + '.xlsx'
                dest_file_closed_events = monday_day_folder + '/Closed_Events_' + date_start + "to" + date_end + '.xlsx'
                dest_file_active_events = monday_day_folder +'/Active_Events_' + date_start + "to" + date_end + '.xlsx'
            else:
                monday_day_folder = monday_folder + '/' + date_start
                dest_file_new_active_events = monday_day_folder +'/New_Events_' + date_start + '.xlsx'
                dest_file_new_closed_events = monday_day_folder + '/New_Closed_Events_' + date_start + '.xlsx'
                dest_file_closed_events = monday_day_folder +'/Closed_Events_' + date_start + '.xlsx'
                dest_file_active_events = monday_day_folder +'/Active_Events_' + date_start + '.xlsx'

            # <editor-fold desc="Get incidents dataframes">
            # Get incidents' dataframes
            print("Reading incident and Event Tracker files and creating dataframes...")
            df_all = pd.read_excel(event_tracker_file_path,
                                   sheet_name=['Active Events', 'Closed Events', 'Active tracker incidents',
                                               'Closed tracker incidents'], engine='openpyxl')

            df_active_eventtracker = df_all['Active Events']
            
            df_closed_eventtracker = df_all['Closed Events']
            df_active_eventtracker_trackers = df_all['Active tracker incidents']
            df_closed_eventtracker_trackers = df_all['Closed tracker incidents']

            new_active_events = df_active_eventtracker.loc[df_active_eventtracker['Event Start Time'] >= first_timestamp]
            new_closed_events = df_closed_eventtracker.loc[df_closed_eventtracker['Event End Time'] >= first_timestamp]
            month_closed_events = df_closed_eventtracker.loc[(df_closed_eventtracker['Event Start Time'] >= start_of_month_timestamp)|~(df_closed_eventtracker['Event End Time'] >= start_of_month_timestamp)]

            file_to_export_dict = {dest_file_new_active_events: new_active_events,
                                   dest_file_new_closed_events: new_closed_events,
                                   dest_file_closed_events: month_closed_events,
                                   dest_file_active_events: df_active_eventtracker}

            try:
                os.makedirs(monday_day_folder)
            except OSError as e:
                if e.errno != errno.EEXIST:
                    raise

            for file in file_to_export_dict.keys():
                print("Creating: " + file)
                dataframe = file_to_export_dict[file]
                writer = pd.ExcelWriter(file, engine='xlsxwriter')
                workbook = writer.book
                dataframe.to_excel(writer, sheet_name='Events', index=False)
                writer.save()
                print('Done!')

    window.close()

if __name__ == "__main__":
    try:
        main()
    except BaseException as e:
        sg.popup(e, title='Error')
        if "out of bounds" in str(e):
            sg.popup("Possible errors:\n- Start/End time of events incorrect \n- Site names incorrect",
                     title='Suggested Action')
        elif str(e) == "Timestamp" :
            sg.popup("Please confirm Irradiance file is correct", title='Suggested Action')
        raise








