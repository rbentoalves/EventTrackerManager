import pandas as pd
import numpy as np
from append_df_to_excel import append_df_to_excel
from append_df_to_excel_existing import append_df_to_excel_existing
from datetime import datetime
import datetime as dt
from openpyxl import Workbook
import openpyxl
import re
import myfunctions as mf
import sys
import PySimpleGUI as sg
import statistics
import math
import matplotlib as mpl
import matplotlib.pyplot as plt
import numpy as np
from scipy.optimize import curve_fit
import calendar


#Create dfs
def create_dfs(df,min_dur: int = 15, roundto: int = 15):

    '''From a dataframe of a report containing all incidents from various sites, this script creates a dictionary
    (list with keys and values (ex: 'LSBP - Grants' : df_active_events_Grants) for ACTIVE and CLOSED EVENTS containing
    all sites present in the orginal dataframe. Independtly of the number of sites in the input dataframe

    USAGE: create_dfs(orginal_df,minimum duration of events (default = 15), auxiliary_df_for_block_identification)

    Returns site_list,df_list_active, df_list_closed'''
    sites = df['Site Name']
    corrected_sites = [mf.correct_site_name(site) for site in sites]
    df['Site Name'] = corrected_sites

    df_closed_all = mf.filter_notprod_and_duration(df, min_dur)                 # creates dataframe with closed, not producing incidents with a minimum specified duration
    df_closed_all = mf.remove_milliseconds(df_closed_all, end_time=True)        # removes milliseconds
    #append_df_to_excel('test.xlsx', df_closed_all, sheet_name='test')


    df_active_all = mf.create_active_events_df(df)                              # creates dataframe with closed, not producing incidents
    df_active_all = mf.remove_milliseconds(df_active_all)                       # removes milliseconds

    df_closed_all = mf.remove_incidents_component_type(df_closed_all, 'Feeder') #removes incidents of inverter modules
    df_active_all = mf.remove_incidents_component_type(df_active_all, 'Feeder')

    #Add capacity of each component



    site_list, df_list_active, df_list_closed = mf.create_df_list(df)

    #Remove sites that need to be ignored
    for site in site_list:
        if "HV Almochuel" in site:
            site_list.remove(site)


    for site in site_list:
        #Create active df for a given site
        df_active = df_active_all.loc[df_active_all['Site Name'] == site]
        df_active = df_active.reset_index(None, drop=True)
        df_active = mf.rounddatesactive_15m(site, df_active, freq = roundto)

        """if site == 'LSBP - Impact_Solar':
            df_active = mf.impact_inv_blocks(df_active, aux_df_blocks)"""
        df_list_active[site] = df_active

        # Create closed df for a given site
        df_closed = df_closed_all.loc[df_closed_all['Site Name'] == site]
        df_closed = df_closed.reset_index(None, drop=True)
        df_closed = mf.rounddatesclosed_15m(site, df_closed, freq = roundto)


        """if site == 'LSBP - Impact_Solar':
            df_closed = mf.impact_inv_blocks(df_closed, aux_df_blocks)"""
        df_list_closed[site] = df_closed



    return site_list,df_list_active, df_list_closed

def create_tracker_dfs(df_all,df_general_info_calc, roundto: int = 15):
    sites = df_all['Site Name']
    corrected_sites = [mf.correct_site_name(site) for site in sites]
    df_all['Site Name'] = corrected_sites

    df_tracker_closed = mf.closedtrackerdf(df_all, df_general_info_calc)
    df_tracker_active = mf.activetrackerdf(df_all, df_general_info_calc)

    df_tracker_active = mf.remove_milliseconds(df_tracker_active)
    df_tracker_closed = mf.remove_milliseconds(df_tracker_closed, end_time=True)

    df_tracker_closed = mf.remove_incidents_component_type(df_tracker_closed, 'TrackerModeEnabled', 'State')
    df_tracker_active = mf.remove_incidents_component_type(df_tracker_active, 'TrackerModeEnabled', 'State')

    df_tracker_closed = mf.rounddatesclosed_15m('Trackers',df_tracker_closed, freq = roundto)
    df_tracker_active = mf.rounddatesactive_15m('Trackers',df_tracker_active, freq = roundto)

    return df_tracker_active,df_tracker_closed

def get_dataframes_to_add_to_EventTracker(report_files,event_tracker_file_path, fmeca_data, component_data, tracker_data):
    '''From Event Tracker & files, gets all dataframes to add separated by dictionaries.
    Returns: df_to_add - dict with new dfs to add
             df_event_tracker - dict with existing dfs in tracker
             fmeca_data - Corrected for Unnamed columns and incomplete entries'''

    #Dataframes from Event Tracker

    df_all = pd.read_excel(event_tracker_file_path,
                           sheet_name=['Active Events', 'Closed Events', 'Active tracker incidents',
                                       'Closed tracker incidents'], engine='openpyxl')

    df_active_eventtracker = df_all['Active Events']
    df_closed_eventtracker = df_all['Closed Events']
    df_active_eventtracker_trackers = df_all['Active tracker incidents']
    df_closed_eventtracker_trackers = df_all['Closed tracker incidents']

    #Dataframes to add
    for report_path in report_files:
        df_active_to_add_report = pd.read_excel(report_path, sheet_name='Active Events', engine='openpyxl')
        df_closed_to_add_report = pd.read_excel(report_path, sheet_name='Closed Events', engine='openpyxl')
        df_active_to_add_trackers_report = pd.read_excel(report_path, sheet_name='Active tracker incidents',
                                                         engine='openpyxl')
        df_closed_to_add_trackers_report = pd.read_excel(report_path, sheet_name='Closed tracker incidents',
                                                         engine='openpyxl')

        try:
            df_active_reports = df_active_reports.append(df_active_to_add_report)
        except NameError:
            df_active_reports = df_active_to_add_report

        try:
            df_closed_reports = df_closed_reports.append(df_closed_to_add_report)
        except NameError:
            df_closed_reports = df_closed_to_add_report

        try:
            df_active_reports_trackers = df_active_reports_trackers.append(df_active_to_add_trackers_report)
        except NameError:
            df_active_reports_trackers = df_active_to_add_trackers_report

        try:
            df_closed_reports_trackers = df_closed_reports_trackers.append(df_closed_to_add_trackers_report)
        except NameError:
            df_closed_reports_trackers = df_closed_to_add_trackers_report

    # Reset Index
    df_active_reports.reset_index(drop=True, inplace=True)
    df_closed_reports.reset_index(drop=True, inplace=True)
    df_active_reports_trackers.reset_index(drop=True, inplace=True)
    df_closed_reports_trackers.reset_index(drop=True, inplace=True)

    # Dicts with dataframes
    df_to_add = {'Closed Events': df_closed_reports,
                 'Closed tracker incidents': df_closed_reports_trackers,
                 'Active Events': df_active_reports,
                 'Active tracker incidents': df_active_reports_trackers}

    df_event_tracker = {'Closed Events': df_closed_eventtracker,
                        'Closed tracker incidents': df_closed_eventtracker_trackers,
                        'Active Events': df_active_eventtracker,
                        'Active tracker incidents': df_active_eventtracker_trackers}

    # Correct any unnamed columns
    fmeca_data = fmeca_data.loc[:, ~fmeca_data.columns.str.contains('^Unnamed')]
    fmeca_data = fmeca_data.dropna(thresh=8)

    # Correct unnamed columns
    for sheet in df_event_tracker:
        df = df_event_tracker[sheet]
        corrected_df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        df_event_tracker[sheet] = corrected_df

    #Correct structure of df_to_add dfs to match Event Tracker
    for df_name_pair in df_to_add.items():

        df_data = df_name_pair[1]
        df_name = df_name_pair[0]
        print(df_name)
        #print(df_data)
        if "Closed" in df_name:
            active = False
        else:
            active = True

        if not "tracker" in df_name:
            component_data_effective = component_data
        else:
            component_data_effective = tracker_data

        # print(df_name, " corrected")
        df_corrected = mf.match_df_to_event_tracker(df_data, component_data_effective, fmeca_data, active=active)

        df_to_add[df_name] = df_corrected



    return df_to_add, df_event_tracker, fmeca_data

def get_final_dataframes_to_add_to_EventTracker(df_to_add, df_event_tracker, fmeca_data):
    '''From the different dataframe dictionaries available (New Reports to Add, Event Tracker info and FMECA data,
    creates dict with final dataframes to add.
    Events are verified, excludes from new additions any incident already on Event Tracker and removes from
    active sheet any closed incident.
    Returns: df_final_to_add'''

    final_df_to_add = {}
    # Choose which incidents to add to event tracker
    for sheet in df_to_add.keys():
        print("Joining new df to event tracker df - " , sheet)
        if "Closed" in sheet:
            df_toadd_id = df_to_add[sheet]['ID'].to_list()  # Get dataframe to add - Closed
            df_ET_id = df_event_tracker[sheet]['ID'].to_list()  # Get dataframe event tracker - Closed

            set_df_ET_id = set(df_ET_id)
            df_toadd_id_tokeep = [x for x in df_toadd_id if x not in set_df_ET_id]

            df_to_add[sheet] = df_to_add[sheet][df_to_add[sheet]['ID'].isin(df_toadd_id_tokeep)].reset_index(drop=True)

            # df_to_add_id = list(set(df_id) - set(df_ET_id))

        else:
            other_sheet = sheet.replace('Active', 'Closed')

            df_toadd_id = df_to_add[sheet]['ID'].to_list()  # Get active dataframe to add and all others

            df_closed_id = df_to_add[other_sheet]['ID'].to_list()
            df_ET_id = df_event_tracker[sheet]['ID'].to_list()
            df_ET_closed_id = df_event_tracker[other_sheet]['ID'].to_list()
            all_ids = df_closed_id + df_ET_id + df_ET_closed_id

            set_all_ids = set(all_ids)
            df_toadd_id_tokeep = [x for x in df_toadd_id if x not in set_all_ids]

            df_to_add[sheet] = df_to_add[sheet][df_to_add[sheet]['ID'].isin(df_toadd_id_tokeep)].reset_index(drop=True)

        # Join new events with events from event tracker and sort them
        if not df_to_add[sheet].empty:
            new_df = pd.concat([df_event_tracker[sheet], df_to_add[sheet]]).sort_values(
                by=['Site Name', 'Event Start Time', 'Related Component'], ascending=[True, False, False],
                ignore_index=True)
        else:
            new_df = df_event_tracker[sheet]
        final_df_to_add[sheet] = new_df

    # Correct final active lists to exclude already closed events
    for sheet in final_df_to_add.keys():
        print("Correcting final dfs to add - ", sheet)
        if "Active" in sheet:
            other_sheet = sheet.replace('Active', 'Closed')
            df_active = final_df_to_add[sheet]
            df_active_id = final_df_to_add[sheet]['ID'].to_list()

            df_closed = final_df_to_add[other_sheet]
            df_closed_id = final_df_to_add[other_sheet]['ID'].to_list()

            set_closed_ids = set(df_closed_id)
            df_tokeep_id = [x for x in df_active_id if x not in set_closed_ids]
            df_toremove_id = [x for x in df_active_id if x in set_closed_ids]

            for id_incident in df_toremove_id:
                index_closed = int(df_closed.loc[df_closed['ID'] == id_incident].index.values)
                index_active = int(df_active.loc[df_active['ID'] == id_incident].index.values)
                df_closed.loc[index_closed, 'Remediation'] = df_active.loc[index_active, 'Remediation']
                df_closed.loc[index_closed, 'Fault'] = df_active.loc[index_active, 'Fault']
                df_closed.loc[index_closed, 'Fault Component'] = df_active.loc[index_active, 'Fault Component']
                df_closed.loc[index_closed, 'Failure Mode'] = df_active.loc[index_active, 'Failure Mode']
                df_closed.loc[index_closed, 'Failure Mechanism'] = df_active.loc[index_active, 'Failure Mechanism']
                df_closed.loc[index_closed, 'Category'] = df_active.loc[index_active, 'Category']
                df_closed.loc[index_closed, 'Subcategory'] = df_active.loc[index_active, 'Subcategory']
                df_closed.loc[index_closed, 'Resolution Category'] = df_active.loc[index_active, 'Resolution Category']

            final_df_to_add[sheet] = final_df_to_add[sheet][
                final_df_to_add[sheet]['ID'].isin(df_tokeep_id)].reset_index(drop=True)
            final_df_to_add[other_sheet] = df_closed
        else:
            pass

        # final_df_to_add[sheet]
    for sheet, df in final_df_to_add.items():
        print("Correcting timestamps on final dfs to add - ", sheet)
        if "Active" in sheet:
            df['Event Start Time'] = [datetime.strptime(str(timestamp), '%Y-%m-%d %H:%M:%S') for timestamp in
                                      df['Event Start Time']]
            df.sort_values(by = ['ID', 'Event Start Time'], inplace = True,ascending=[True, False],ignore_index=True)
        else:
            df['Event Start Time'] = [datetime.strptime(str(timestamp), '%Y-%m-%d %H:%M:%S') for timestamp in
                                      df['Event Start Time']]
            df['Event End Time'] = [datetime.strptime(str(timestamp), '%Y-%m-%d %H:%M:%S') for timestamp in
                                    df['Event End Time']]
            df.sort_values(by=['Event Start Time', 'ID'], inplace=True, ascending=[False, False], ignore_index=True)
        final_df_to_add[sheet] = df

    final_df_to_add['FMECA'] = fmeca_data
    final_df_to_add = dict(sorted(final_df_to_add.items()))

    return final_df_to_add

def create_fmeca_dataframes_for_validation(fmeca_data):
    '''From FMECA Table creates all dataframes needed for data_validation.
    Structures Faults, Fault Component, Failure Mode, Failure Mechanism, Category and Subcategory'''

    # data validation for FMECA

    # next level is dependent on combination of previous levels
    faults_fmeca = list(set(fmeca_data['Fault'].to_list()))
    fault_component_fmeca = dict(
        (fault, list(set(fmeca_data.loc[fmeca_data['Fault'] == fault]['Fault Component'].to_list()))) for fault in
        faults_fmeca)
    failure_mode_fmeca = dict(((fault, fault_comp), list(set(
        fmeca_data.loc[(fmeca_data['Fault'] == fault) & (fmeca_data['Fault Component'] == fault_comp)][
            'Failure Mode'].to_list()))) for fault, fault_comps in fault_component_fmeca.items() for fault_comp in
                              fault_comps)
    failure_mechanism_fmeca = dict((fault_and_comp + (fail_mode,), list(set(fmeca_data.loc[(fmeca_data['Fault'] ==
                                                                                            fault_and_comp[0]) & (
                                                                                                   fmeca_data[
                                                                                                       'Fault Component'] ==
                                                                                                   fault_and_comp[
                                                                                                       1]) & (
                                                                                                   fmeca_data[
                                                                                                       'Failure Mode'] == fail_mode)][
                                                                                'Failure Mechanism'].to_list()))) for
                                   fault_and_comp, fail_modes in failure_mode_fmeca.items() for fail_mode in fail_modes)

    category_fmeca = dict((fault_and_comp_mode + (fail_mec,), list(set(fmeca_data.loc[(fmeca_data['Fault'] ==
                                                                                       fault_and_comp_mode[0]) & (
                                                                                              fmeca_data[
                                                                                                  'Fault Component'] ==
                                                                                              fault_and_comp_mode[
                                                                                                  1]) & (fmeca_data[
                                                                                                             'Failure Mode'] ==
                                                                                                         fault_and_comp_mode[
                                                                                                             2]) & (
                                                                                              fmeca_data[
                                                                                                  'Failure Mechanism'] == fail_mec)][
                                                                           'Category'].to_list()))) for
                          fault_and_comp_mode, fail_mecs in failure_mechanism_fmeca.items() for fail_mec in fail_mecs)
    subcategory_fmeca = dict((fault_and_comp_mode_mec + (cat,), list(set(fmeca_data.loc[(fmeca_data['Fault'] ==
                                                                                         fault_and_comp_mode_mec[0]) & (
                                                                                                fmeca_data[
                                                                                                    'Fault Component'] ==
                                                                                                fault_and_comp_mode_mec[
                                                                                                    1]) & (fmeca_data[
                                                                                                               'Failure Mode'] ==
                                                                                                           fault_and_comp_mode_mec[
                                                                                                               2]) & (
                                                                                                fmeca_data[
                                                                                                    'Failure Mechanism'] ==
                                                                                                fault_and_comp_mode_mec[
                                                                                                    3]) & (fmeca_data[
                                                                                                               'Category'] == cat)][
                                                                             'Subcategory'].to_list()))) for
                             fault_and_comp_mode_mec, cats in category_fmeca.items() for cat in cats)

    # Change multi-level options' keys to have all dependencies on key
    fault_component_fmeca_newkeys = dict(
        (key, key.replace(" ", "_").replace("-", "_")) for key in fault_component_fmeca.keys())
    failure_mode_fmeca_newkeys = dict(
        (key, "_".join(key).replace(" ", "_").replace("-", "_")) for key in failure_mode_fmeca.keys())
    failure_mechanism_fmeca_newkeys = dict(
        (key, "_".join(key).replace(" ", "_").replace("-", "_")) for key in failure_mechanism_fmeca.keys())
    category_fmeca_newkeys = dict(
        (key, "_".join(key).replace(" ", "_").replace("-", "_")) for key in category_fmeca.keys())
    subcategory_fmeca_newkeys = dict(
        (key, "_".join(key).replace(" ", "_").replace("-", "_")) for key in subcategory_fmeca.keys())

    fault_component_fmeca = mf.rename_dict_keys(fault_component_fmeca, fault_component_fmeca_newkeys)
    failure_mode_fmeca = mf.rename_dict_keys(failure_mode_fmeca, failure_mode_fmeca_newkeys)
    failure_mechanism_fmeca = mf.rename_dict_keys(failure_mechanism_fmeca, failure_mechanism_fmeca_newkeys)
    category_fmeca = mf.rename_dict_keys(category_fmeca, category_fmeca_newkeys)
    subcategory_fmeca = mf.rename_dict_keys(subcategory_fmeca, subcategory_fmeca_newkeys)

    # Create dfs

    df_faults_fmeca = pd.DataFrame(data={'Faults': faults_fmeca})
    df_fault_component_fmeca = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in fault_component_fmeca.items()]))
    df_failure_mode_fmeca = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in failure_mode_fmeca.items()]))
    df_failure_mechanism_fmeca = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in failure_mechanism_fmeca.items()]))
    df_category_fmeca = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in category_fmeca.items()]))
    df_subcategory_fmeca = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in subcategory_fmeca.items()]))

    dict_fmeca_shapes = {'Faults': (df_faults_fmeca, df_faults_fmeca.shape),
                         'Failure_Component': (df_fault_component_fmeca, df_fault_component_fmeca.shape),
                         'Failure_Mode': (df_failure_mode_fmeca, df_failure_mode_fmeca.shape),
                         'Failure_Mechanism': (df_failure_mechanism_fmeca, df_failure_mechanism_fmeca.shape),
                         'Category': (df_category_fmeca, df_category_fmeca.shape),
                         'Subcategory': (df_subcategory_fmeca, df_subcategory_fmeca.shape)}




    return dict_fmeca_shapes




#Read approved incidents

def read_approved_incidents(incidents_file, site_list, roundto : int = 15):
    for site in site_list:

        if "LSBP - " in site or "LSBP – " in site:
            onlysite = site[7:]
        else:
            onlysite = site

        if onlysite[-1:] == " ":
            onlysite = onlysite[:len(onlysite)-1]
        active_sheet_name = onlysite + ' Active'
        closed_sheet_name = onlysite

        active_df = pd.read_excel(incidents_file, sheet_name=active_sheet_name, engine="openpyxl")
        active_df = active_df.loc[active_df['InSolar Check'] == 'x']  # filter for checked events
        active_df = active_df.reset_index(None, drop=True)
        active_df = mf.remove_milliseconds(active_df)
        active_df = mf.rounddatesactive_15m(site, active_df, freq = roundto)

        closed_df = pd.read_excel(incidents_file, sheet_name=closed_sheet_name, engine="openpyxl")
        closed_df = closed_df.loc[closed_df['InSolar Check'] == 'x']  # filter for checked events
        closed_df = closed_df.reset_index(None, drop=True)
        closed_df = mf.remove_milliseconds(closed_df, end_time=True)
        closed_df = mf.rounddatesclosed_15m(site, closed_df, freq = roundto)
        closed_df = mf.correct_duration_of_event(closed_df)


        try:
            if site not in df_list_active.keys():
                df_list_active[site] = active_df
        except NameError:
            df_list_active = {site: active_df}

        try:
            if site not in df_list_closed.keys():
                df_list_closed[site] = closed_df
        except NameError:
            df_list_closed = {site: closed_df}

    return  df_list_active, df_list_closed

def read_approved_tracker_inc(tracker_file, roundto: int = 15):
    df_info_trackers = pd.read_excel(tracker_file, sheet_name='Trackers info', engine="openpyxl")

    df_tracker_active = pd.read_excel(tracker_file, sheet_name='Active tracker incidents', engine="openpyxl")
    df_tracker_active = df_tracker_active.loc[df_tracker_active['InSolar Check'] == 'x']  # filter for checked events
    df_tracker_active = df_tracker_active.reset_index(None, drop=True)
    df_tracker_active = mf.remove_milliseconds(df_tracker_active)
    df_tracker_active = mf.rounddatesactive_15m('Trackers', df_tracker_active, freq = roundto)

    df_tracker_closed = pd.read_excel(tracker_file, sheet_name='Closed tracker incidents', engine="openpyxl")
    df_tracker_closed = df_tracker_closed.loc[df_tracker_closed['InSolar Check'] == 'x']  # filter for checked events
    df_tracker_closed = df_tracker_closed.reset_index(None, drop=True)
    df_tracker_closed = mf.remove_milliseconds(df_tracker_closed, end_time=True)
    df_tracker_closed = mf.rounddatesclosed_15m('Trackers', df_tracker_closed, freq = roundto)


    return df_tracker_active, df_tracker_closed


#Analysis of incidents

def analysis_component_incidents(df_incidents_analysis, site_list, df_list_closed, df_list_active, df_info_sunlight):

    df_incidents_analysis = mf.fill_events_analysis_dataframe(df_incidents_analysis, df_info_sunlight)

    for site in site_list:

        index_site_array = df_info_sunlight[df_info_sunlight['Site'] == site].index.values
        index_site = int(index_site_array[0])

        df_closed_events = df_list_closed[site]
        df_active_events = df_list_active[site]

        # Add efect of closed events
        df_incidents_analysis = mf.analysis_closed_incidents(site, index_site, df_incidents_analysis, df_closed_events,
                                                             df_info_sunlight)

        # Add effect of active events
        df_incidents_analysis = mf.analysis_active_incidents(site, index_site, df_incidents_analysis, df_active_events,
                                                             df_info_sunlight)

    #print(df_incidents_analysis)
    return df_incidents_analysis

def analysis_tracker_incidents(df_tracker_analysis, df_tracker_closed, df_tracker_active, df_info_sunlight):

    df_tracker_analysis = mf.fill_events_analysis_dataframe(df_tracker_analysis, df_info_sunlight)

    # Add efect of closed events
    df_tracker_analysis = mf.analysis_closed_tracker_incidents(df_tracker_analysis, df_tracker_closed,
                                                                df_info_sunlight)

    # Add effect of active events
    df_tracker_analysis = mf.analysis_active_tracker_incidents(df_tracker_analysis, df_tracker_active,
                                                               df_info_sunlight)


    #print(df_tracker_analysis)

    return df_tracker_analysis

def calculate_activehours_energylost_incidents(df, df_all_irradiance,df_all_export, budget_pr,corrected_incidents_dict,active_events: bool = False,
                                     recalculate: bool = False,granularity: float=0.25):

    if active_events == True:
        for index, row in df.iterrows():
            site = row['Site Name']
            incident_id = row['ID']
            component = row['Related Component']
            capacity = row['Capacity Related Component']
            real_event_start_time = row['Event Start Time']
            event_start_time = row['Rounded Event Start Time']
            budget_pr_site = budget_pr.loc[site, :]

            if type(event_start_time) == str:
                row['Rounded Event Start Time'] = event_start_time = datetime.strptime(str(event_start_time), '%Y-%m-%d %H:%M:%S')

            if type(real_event_start_time) == str:
                row['Event Start Time'] = real_event_start_time = datetime.strptime(str(event_start_time), '%Y-%m-%d %H:%M:%S')

            # event_start_time = row['Event Start Time']
            if incident_id not in corrected_incidents_dict.keys():
                df_irradiance_site = df_all_irradiance.loc[:, df_all_irradiance.columns.str.contains(site + '|Timestamp')]
                df_irradiance_event = df_irradiance_site.loc[df_irradiance_site['Timestamp'] >= event_start_time]

                # Get percentages of first timestamp to account for rounding
                index_event_start_time = df_irradiance_site.loc[df_irradiance_site['Timestamp'] == event_start_time].index.values[0]
                percentage_of_timestamp_start = mf.get_percentage_of_timestamp(real_event_start_time, event_start_time)


                actual_column, curated_column,data_gaps_proportion, poa_avg_column = mf.get_actual_irradiance_column(df_irradiance_event)
                # print(actual_column)

                if actual_column == None:
                    print(component, ' on ', event_start_time, ': No irradiance available')
                    continue
                """elif 'curated' in actual_column:
                    print(component, ' on ', event_start_time,': Using curated irradiance')
                else:
                    print(component, ' on ', event_start_time,': Using poa average due to curated irradiance having over 25% of data gaps')
                    """


                data_gaps_percentage = "{:.2%}".format(data_gaps_proportion)
                print(incident_id, ' - Data Gaps percentage: ', data_gaps_percentage)

                #Correct irradiance in first timestamp to account for rounding
                df_irradiance_event.at[index_event_start_time, actual_column] = percentage_of_timestamp_start * df_irradiance_event.loc[index_event_start_time, actual_column]
                df_irradiance_event_activeperiods = df_irradiance_event.loc[df_irradiance_event[actual_column] > 20]

                duration = df_irradiance_event.shape[0] * granularity
                active_hours = df_irradiance_event_activeperiods.shape[0] * granularity
                if site == component:
                    df_export_site = df_all_export.loc[:,
                                     df_all_export.columns.str.contains(site + '|Timestamp')]
                    export_column = df_all_export.columns[df_all_export.columns.str.contains(site)].values[0]

                    df_export_event = df_export_site.loc[df_export_site['Timestamp'] >= event_start_time]

                    df_export_event.at[index_event_start_time, export_column] = percentage_of_timestamp_start * \
                                                                                df_export_event.loc[
                                                                                    index_event_start_time, export_column]

                    energy_produced = df_export_event[export_column].sum()

                    energy_lost = (sum(
                        [row_el[actual_column] * budget_pr_site.loc[str(row_el['Timestamp'].date())[:-2] + "01"] for
                         index_el, row_el in
                         df_irradiance_event_activeperiods.iterrows()]) * capacity * granularity) / 1000

                    energy_lost = (energy_lost - energy_produced)

                    if energy_lost < 0:
                        enery_lost = 0

                else:
                    energy_lost = (sum(
                        [row_el[actual_column] * budget_pr_site.loc[str(row_el['Timestamp'].date())[:-2] + "01"] for
                         index_el, row_el in
                         df_irradiance_event_activeperiods.iterrows()]) * capacity * granularity) / 1000


            else:
                data_gaps_percentage = "{:.2%}".format(corrected_incidents_dict[incident_id]['Data Gaps Proportion'])
                actual_column = corrected_incidents_dict[incident_id]['Irradiance Column']
                df_irradiance_event_raw = corrected_incidents_dict[incident_id][
                    'Irradiance Raw']
                df_irradiance_event = corrected_incidents_dict[incident_id][
                    'Corrected Irradiance Incident']
                df_cleaned_irradiance_event = corrected_incidents_dict[incident_id][
                    'Cleaned Corrected Irradiance Incident']

                print('Using Corrected Incident for: ', component, " on ", site, "with ", data_gaps_percentage, " of data gaps")

                df_irradiance_event_activeperiods = df_irradiance_event.loc[
                    df_irradiance_event[actual_column] > 20]
                df_cleaned_irradiance_event_activeperiods = df_cleaned_irradiance_event.loc[
                    df_cleaned_irradiance_event[actual_column] > 20]

                duration = df_irradiance_event_raw.shape[0] * granularity
                active_hours = df_irradiance_event_activeperiods.shape[0] * granularity
                energy_lost = (sum([row_el[actual_column] * budget_pr_site.loc[str(row_el['Timestamp'].date())[:-2] + "01"] for
                                    index_el, row_el in
                                    df_cleaned_irradiance_event_activeperiods.iterrows()]) * capacity * granularity) / 1000
                """energy_lost = (sum([row_el[actual_column] * float(budget_pr_site[row_el['Timestamp'].month].values) for
                                    index_el, row_el in
                                    df_cleaned_irradiance_event_activeperiods.iterrows()]) * capacity * granularity) / 1000
"""
            if active_hours < 0:
                active_hours = duration

            df.loc[index, 'Duration (h)'] = duration
            df.loc[index, 'Active Hours (h)'] = active_hours
            df.loc[index, 'Energy Lost (MWh)'] = energy_lost / 1000

    else:
        if recalculate == True:
            df = mf.rounddatesclosed_15m("All", df)
            for index, row in df.iterrows():
                site = row['Site Name']
                incident_id = row['ID']
                component = row['Related Component']
                capacity = row['Capacity Related Component']
                event_start_time = row['Rounded Event Start Time']
                event_end_time = row['Rounded Event End Time']
                real_event_start_time = row['Event Start Time']
                real_event_end_time = row['Event End Time']
                budget_pr_site = budget_pr.loc[site, :]



                if incident_id not in corrected_incidents_dict.keys():
                    df_irradiance_site = df_all_irradiance.loc[:,
                                         df_all_irradiance.columns.str.contains(site + '|Timestamp')]
                    df_irradiance_event = df_irradiance_site.loc[(df_irradiance_site['Timestamp'] >= event_start_time) & (
                            df_irradiance_site['Timestamp'] <= event_end_time)]
                    actual_column, curated_column, data_gaps_proportion, poa_avg_column = mf.get_actual_irradiance_column(df_irradiance_event)

                    if actual_column == None:
                        print(component, ' on ', event_start_time, ': No irradiance available')
                        continue

                    elif 'curated' in actual_column:
                        print(component, ' on ', event_start_time, ': Using curated irradiance')
                    else:
                        print(component, ' on ', event_start_time,
                              ': Using poa average due to curated irradiance having over 25% of data gaps')

                    data_gaps_percentage = "{:.2%}".format(data_gaps_proportion)
                    print(incident_id, ' - Data Gaps percentage: ', data_gaps_percentage)

                    df_irradiance_event_activeperiods = df_irradiance_event.loc[df_irradiance_event[actual_column] > 20]

                    """duration = df_irradiance_event.shape[0] * granularity
                    #duration =
                    active_hours = df_irradiance_event_activeperiods.shape[0] * granularity
                    """
                    duration = ((real_event_end_time-real_event_start_time).days * 24) + ((real_event_end_time-real_event_start_time).seconds/3600)

                    active_hours = duration - ((
                            df_irradiance_event.shape[0] - df_irradiance_event_activeperiods.shape[0]) * granularity)
                    #print(real_event_end_time, " - ", real_event_start_time, " = ", duration)
                    #print(real_event_end_time, " - ", real_event_start_time, " = ", active_hours , " - Active Hours")

                    energy_lost = (sum(
                            [row_el[actual_column] * budget_pr_site.loc[str(row_el['Timestamp'].date())[:-2] + "01"] for
                             index_el, row_el in
                             df_irradiance_event_activeperiods.iterrows()]) * capacity * granularity) / 1000

                else:
                    data_gaps_percentage = "{:.2%}".format(
                        corrected_incidents_dict[incident_id]['Data Gaps Proportion'])
                    actual_column = corrected_incidents_dict[incident_id]['Irradiance Column']

                    df_irradiance_event_raw = corrected_incidents_dict[incident_id]['Irradiance Raw']
                    df_irradiance_event = corrected_incidents_dict[incident_id]['Corrected Irradiance Incident']
                    df_cleaned_irradiance_event = corrected_incidents_dict[incident_id][
                        'Cleaned Corrected Irradiance Incident']

                    """df_irradiance_event = 
                    df_cleaned_irradiance_event ="""

                    print('Using Corrected Incident for: ', component, " on ", site, " with ", data_gaps_percentage,
                          " of data gaps")

                    df_irradiance_event_activeperiods = df_irradiance_event.loc[
                        df_irradiance_event[actual_column] > 20]
                    df_cleaned_irradiance_event_activeperiods = df_cleaned_irradiance_event.loc[
                        df_cleaned_irradiance_event[actual_column] > 20]

                    duration = ((real_event_end_time-real_event_start_time).days * 24) + ((real_event_end_time-real_event_start_time).seconds/3600)

                    active_hours = duration - (
                                df_irradiance_event.shape[0] - df_irradiance_event_activeperiods.shape[0]) * granularity

                    #print(real_event_end_time, " - ", real_event_start_time, " = ", duration)
                    #print(real_event_end_time, " - ", real_event_start_time, " = ", active_hours, " - Active Hours")

                    energy_lost = (sum(
                        [row_el[actual_column] * budget_pr_site.loc[str(row_el['Timestamp'].date())[:-2] + "01"] for
                         index_el, row_el in
                         df_cleaned_irradiance_event_activeperiods.iterrows()]) * capacity * granularity) / 1000

                if active_hours < 0:
                    active_hours = duration

                df.loc[index, 'Duration (h)'] = duration
                df.loc[index, 'Active Hours (h)'] = active_hours
                df.loc[index, 'Energy Lost (MWh)'] = energy_lost / 1000
        else:
            df_to_update = df.loc[df['Energy Lost (MWh)'].isnull()]
            df_to_update = mf.rounddatesclosed_15m("All", df_to_update)

            for index, row in df_to_update.iterrows():
                site = row['Site Name']
                incident_id = row['ID']
                component = row['Related Component']
                capacity = row['Capacity Related Component']
                event_start_time = row['Rounded Event Start Time']
                event_end_time = row['Rounded Event End Time']
                real_event_start_time = row['Event Start Time']
                real_event_end_time = row['Event End Time']
                budget_pr_site = budget_pr.loc[site, :]

                if incident_id not in corrected_incidents_dict.keys():
                    df_irradiance_site = df_all_irradiance.loc[:,
                                         df_all_irradiance.columns.str.contains(site + '|Timestamp')]

                    df_irradiance_event = df_irradiance_site.loc[
                        (df_irradiance_site['Timestamp'] >= event_start_time) & (
                            df_irradiance_site['Timestamp'] <= event_end_time)]

                    #Get percentages of first timestamp to account for rounding
                    index_event_start_time = \
                    df_irradiance_site.loc[df_irradiance_site['Timestamp'] == event_start_time].index.values[0]

                    index_event_end_time = \
                        df_irradiance_site.loc[df_irradiance_site['Timestamp'] == event_end_time].index.values[0]

                    percentage_of_timestamp_start = mf.get_percentage_of_timestamp(real_event_start_time,
                                                                                   event_start_time)
                    percentage_of_timestamp_end = mf.get_percentage_of_timestamp(real_event_end_time,
                                                                                   event_end_time)


                    #GEt actual column to work with
                    actual_column, curated_column, data_gaps_proportion, poa_avg_column = mf.get_actual_irradiance_column(
                        df_irradiance_event)

                    if actual_column == None:
                        print(component, ' on ', event_start_time, ': No irradiance available')
                        continue

                    elif 'curated' in actual_column:
                        print(component, ' on ', event_start_time, ': Using curated irradiance')
                    else:
                        print(component, ' on ', event_start_time,
                              ': Using poa average due to curated irradiance having over 25% of data gaps')

                    #Communicate data gaps percentage
                    data_gaps_percentage = "{:.2%}".format(data_gaps_proportion)
                    print(incident_id, ' - Data Gaps percentage: ', data_gaps_percentage)

                    # Correct irradiance in first timestamp to account for rounding
                    df_irradiance_event.at[index_event_start_time, actual_column] = percentage_of_timestamp_start * \
                                                                                    df_irradiance_event.loc[
                                                                                        index_event_start_time, actual_column]

                    df_irradiance_event.at[index_event_end_time, actual_column] = percentage_of_timestamp_end * \
                                                                                    df_irradiance_event.loc[
                                                                                        index_event_end_time, actual_column]

                    #Get irradiance periods over 20W/m2
                    df_irradiance_event_activeperiods = df_irradiance_event.loc[df_irradiance_event[actual_column] > 20]

                    duration = ((real_event_end_time-real_event_start_time).days * 24) + ((real_event_end_time-real_event_start_time).seconds/3600)
                    active_hours = duration - (df_irradiance_event.shape[0]-df_irradiance_event_activeperiods.shape[0]) * granularity
                    if site == component:
                        df_export_site = df_all_export.loc[:,
                                             df_all_export.columns.str.contains(site + '|Timestamp')]
                        export_column = df_all_export.columns[df_all_export.columns.str.contains(site)].values[0]

                        df_export_event = df_export_site.loc[
                            (df_export_site['Timestamp'] >= event_start_time) & (
                                df_export_site['Timestamp'] <= event_end_time)]

                        df_export_event.at[index_event_start_time, export_column] = percentage_of_timestamp_start * \
                                                                                        df_export_event.loc[
                                                                                            index_event_start_time, export_column]

                        df_export_event.at[index_event_end_time, export_column] = percentage_of_timestamp_end * \
                                                                                      df_export_event.loc[
                                                                                          index_event_end_time, export_column]

                        energy_produced = df_export_event[export_column].sum()

                        energy_lost = (sum(
                            [row_el[actual_column] * budget_pr_site.loc[str(row_el['Timestamp'].date())[:-2] + "01"] for
                             index_el, row_el in
                             df_irradiance_event_activeperiods.iterrows()]) * capacity * granularity)/ 1000

                        print("Site: ", site, "\nEnergy produced: ", energy_produced, "\nEnergy Expected: ", energy_lost)

                        energy_lost = (energy_lost - energy_produced)

                        print("Real Energy Lost: ", energy_lost)

                        if energy_lost < 0:
                            enery_lost = 0

                    else:
                        energy_lost = (sum(
                            [row_el[actual_column] * budget_pr_site.loc[str(row_el['Timestamp'].date())[:-2] + "01"] for
                             index_el, row_el in
                             df_irradiance_event_activeperiods.iterrows()]) * capacity * granularity) / 1000


                else:
                    data_gaps_percentage = "{:.2%}".format(
                        corrected_incidents_dict[incident_id]['Data Gaps Proportion'])

                    actual_column = corrected_incidents_dict[incident_id]['Irradiance Column']

                    df_irradiance_event_raw = corrected_incidents_dict[incident_id][
                        'Irradiance Raw']
                    df_irradiance_event = corrected_incidents_dict[incident_id][
                        'Corrected Irradiance Incident']
                    df_cleaned_irradiance_event = corrected_incidents_dict[incident_id][
                        'Cleaned Corrected Irradiance Incident']

                    print('Using Corrected Incident for: ', component, " on ", site, "with ", data_gaps_percentage,
                          " of data gaps")

                    df_irradiance_event_activeperiods = df_irradiance_event.loc[
                        df_irradiance_event[actual_column] > 20]
                    df_cleaned_irradiance_event_activeperiods = df_cleaned_irradiance_event.loc[
                        df_cleaned_irradiance_event[actual_column] > 20]

                    duration = ((real_event_end_time-real_event_start_time).days * 24) + ((real_event_end_time-real_event_start_time).seconds/3600)
                    active_hours = duration - (df_irradiance_event.shape[0]-df_irradiance_event_activeperiods.shape[0]) * granularity
                    energy_lost = (sum(
                        [row_el[actual_column] * budget_pr_site.loc[str(row_el['Timestamp'].date())[:-2] + "01"] for
                         index_el, row_el in
                         df_cleaned_irradiance_event_activeperiods.iterrows()]) * capacity * granularity) / 1000

                if active_hours < 0:
                    active_hours = duration

                df.loc[index, 'Duration (h)'] = duration
                df.loc[index, 'Active Hours (h)'] = active_hours
                df.loc[index, 'Energy Lost (MWh)'] = energy_lost / 1000



    return df



# Data Analysis -----------------------------------------------------------------------------

def calculate_pr_inverters(inverter_list, all_inverter_power_data_dict, site_info, general_info,
                           pr_type: str = 'raw', granularity: str = 'daily'):

    possible_prs = ['raw', 'corrected', 'corrected_DCfocus']
    possible_gran = ['daily', 'monthly']

    if pr_type not in possible_prs:
        print('Possible PR types: ' + str(possible_prs) + "\n Your input: " + str(pr_type) )
        print('Please try again. :)')
        sys.exit()

    if granularity not in possible_gran:
        print('Possible PR types: ' + str(possible_gran) + "\n Your input: " + str(granularity))
        print('Please try again. :)')
        sys.exit()

    days_under_analysis = site_info['Days']
    months_under_analysis = site_info['Months']

    if pr_type == 'raw' and granularity == 'daily':
        for inverter in inverter_list:
            print(inverter)
            power_data = all_inverter_power_data_dict[inverter]['Power Data']
            power_data['Day'] = pd.to_datetime(power_data['Timestamp']).dt.date
            power_data['Month'] = pd.to_datetime(power_data['Timestamp']).dt.month

            maxexport_capacity_ac = float(site_info['Component Info'].loc[site_info['Component Info']['Component'] == inverter][
                                    'Capacity AC'].values) * 1.001

            daily_pr_df_inverter,irradiance_column = mf.calculate_daily_raw_pr(power_data, days_under_analysis, inverter)

            try:
                df_to_add = daily_pr_df_inverter.drop(columns=irradiance_column)
                daily_pr_df = pd.concat([daily_pr_df, df_to_add], axis=1)

            except NameError:
                daily_pr_df = daily_pr_df_inverter

        return daily_pr_df

    elif pr_type == 'corrected' and granularity == 'daily':
        for inverter in inverter_list:
            # print(inverter)
            power_data = all_inverter_power_data_dict[inverter]['Power Data']
            power_data['Day'] = pd.to_datetime(power_data['Timestamp']).dt.date
            power_data['Month'] = pd.to_datetime(power_data['Timestamp']).dt.month

            maxexport_capacity_ac = float(site_info['Component Info'].loc[site_info['Component Info']['Component'] == inverter][
                                    'Capacity AC'].values) * 1.001
            print(maxexport_capacity_ac)
            daily_corrected_pr_df_inverter,irradiance_column = mf.calculate_daily_corrected_pr(power_data, days_under_analysis, inverter,
                                                                             maxexport_capacity_ac)

            try:
                corrected_df_to_add = daily_corrected_pr_df_inverter.drop(columns=irradiance_column)
                corrected_daily_pr_df = pd.concat([corrected_daily_pr_df, corrected_df_to_add], axis=1)

            except NameError:
                corrected_daily_pr_df = daily_corrected_pr_df_inverter

        return corrected_daily_pr_df


    elif pr_type == 'corrected_DCfocus' and granularity == 'daily':
        for inverter in inverter_list:
            # print(inverter)
            power_data = all_inverter_power_data_dict[inverter]['Power Data']
            power_data['Day'] = pd.to_datetime(power_data['Timestamp']).dt.date
            power_data['Month'] = pd.to_datetime(power_data['Timestamp']).dt.month

            maxexport_capacity_ac = float(site_info['Component Info'].loc[site_info['Component Info']['Component'] == inverter][
                                              'Capacity AC'].values) * 1.001

            dcfocus_corrected_df,irradiance_column = mf.calculate_daily_corrected_pr_focusDC(power_data, days_under_analysis, inverter,
                                                                                             maxexport_capacity_ac)

            try:
                dcfocus_corrected_df_to_add = dcfocus_corrected_df.drop(columns=irradiance_column)
                dcfocus_corrected_daily_pr_df = pd.concat([dcfocus_corrected_daily_pr_df, dcfocus_corrected_df_to_add],
                                                          axis=1)

            except NameError:
                dcfocus_corrected_daily_pr_df = dcfocus_corrected_df
        return dcfocus_corrected_daily_pr_df
    elif pr_type == 'raw' and granularity == 'monthly':

        for inverter in inverter_list:
            # print(inverter)
            power_data = all_inverter_power_data_dict[inverter]['Power Data']
            power_data['Day'] = pd.to_datetime(power_data['Timestamp']).dt.date
            power_data['Month'] = pd.to_datetime(power_data['Timestamp']).apply(lambda x: x.strftime('%m-%Y'))



            raw_df_month,raw_powers_df_forsite_inv,irradiance_column = mf.calculate_monthly_raw_pr(power_data, months_under_analysis, inverter)

            try:
                dcfocus_corrected_df_to_add = raw_df_month.drop(columns=irradiance_column)
                raw_daily_pr_df = pd.concat([raw_daily_pr_df, dcfocus_corrected_df_to_add],
                                                          axis=1)

            except NameError:
                raw_daily_pr_df = raw_df_month


            try:
                raw_powers_df_forsite = pd.concat(
                    [raw_powers_df_forsite, raw_powers_df_forsite_inv], axis=1)

            except NameError:
                print('Creating dataframe with all inverters')
                raw_powers_df_forsite = raw_powers_df_forsite_inv

        # Add site wide results
        ac_power_results = raw_powers_df_forsite.loc[:,
                           raw_powers_df_forsite.columns.str.contains('Inverter AC')]
        ac_power_results['Site'] = [ac_power_results.loc[i, :].sum() for i in ac_power_results.index]

        ideal_power_results = raw_powers_df_forsite.loc[:,
                                 raw_powers_df_forsite.columns.str.contains('Ideal')]
        ideal_power_results['Site'] = [ideal_power_results.loc[i, :].sum() for i in
                                          ideal_power_results.index]

        site_pr = ac_power_results['Site'] / ideal_power_results['Site']

        raw_daily_pr_df.insert(len(raw_daily_pr_df.columns) - 1, 'Site PR %',
                                       site_pr)

        return raw_daily_pr_df


    elif pr_type == 'corrected' and granularity == 'monthly':
        for inverter in inverter_list:
            # print(inverter)
            power_data = all_inverter_power_data_dict[inverter]['Power Data']
            power_data['Day'] = pd.to_datetime(power_data['Timestamp']).dt.date
            power_data['Month'] = pd.to_datetime(power_data['Timestamp']).apply(lambda x: x.strftime('%m-%Y'))

            maxexport_capacity_ac = float(site_info['Component Info'].loc[site_info['Component Info']['Component'] == inverter][
                                    'Capacity AC'].values) * 1.001

            corrected_df_month, corrected_powers_df_forsite_inv,irradiance_column = mf.calculate_monthly_corrected_pr_and_production(
                power_data, months_under_analysis, inverter, maxexport_capacity_ac)

            try:
                corrected_df_to_add_month = corrected_df_month.drop(columns=irradiance_column)
                corrected_monthly_pr_df = pd.concat(
                    [corrected_monthly_pr_df, corrected_df_to_add_month], axis=1)

            except NameError:
                corrected_monthly_pr_df = corrected_df_month

            try:
                corrected_powers_df_forsite = pd.concat(
                    [corrected_powers_df_forsite, corrected_powers_df_forsite_inv], axis=1)

            except NameError:
                print('Creating dataframe with all inverters')
                corrected_powers_df_forsite = corrected_powers_df_forsite_inv

            # Add site wide results
        ac_power_results = corrected_powers_df_forsite.loc[:,
                           corrected_powers_df_forsite.columns.str.contains('Inverter AC')]
        ac_power_results['Site'] = [ac_power_results.loc[i, :].sum() for i in ac_power_results.index]

        ideal_power_results = corrected_powers_df_forsite.loc[:,
                                 corrected_powers_df_forsite.columns.str.contains('Ideal')]
        ideal_power_results['Site'] = [ideal_power_results.loc[i, :].sum() for i in
                                          ideal_power_results.index]

        site_pr = ac_power_results['Site'] / ideal_power_results['Site']

        corrected_monthly_pr_df.insert(len(corrected_monthly_pr_df.columns) - 1, 'Site PR %',
                                               site_pr)

        return corrected_monthly_pr_df


    elif pr_type == 'corrected_DCfocus' and granularity == 'monthly':
        for inverter in inverter_list:
            # print(inverter)
            power_data = all_inverter_power_data_dict[inverter]['Power Data']
            power_data['Day'] = pd.to_datetime(power_data['Timestamp']).dt.date
            power_data['Month'] = pd.to_datetime(power_data['Timestamp']).apply(lambda x: x.strftime('%m-%Y'))

            maxexport_capacity_ac = float(site_info['Component Info'].loc[site_info['Component Info']['Component'] == inverter][
                                    'Capacity AC'].values) * 1.001

            dcfocus_corrected_df_month, dcfocus_corrected_powers_df_forsite_inv,irradiance_column = mf.calculate_monthly_corrected_pr_and_production_focusDC(
                power_data, months_under_analysis, inverter, maxexport_capacity_ac)

            try:
                dcfocus_corrected_df_to_add_month = dcfocus_corrected_df_month.drop(columns=irradiance_column)
                dcfocus_corrected_monthly_pr_df = pd.concat(
                    [dcfocus_corrected_monthly_pr_df, dcfocus_corrected_df_to_add_month], axis=1)

            except NameError:
                dcfocus_corrected_monthly_pr_df = dcfocus_corrected_df_month

            try:
                dcfocus_corrected_powers_df_forsite = pd.concat(
                    [dcfocus_corrected_powers_df_forsite, dcfocus_corrected_powers_df_forsite_inv], axis=1)

            except NameError:
                dcfocus_corrected_powers_df_forsite = dcfocus_corrected_powers_df_forsite_inv

        # Add site wide results
        ac_power_results = dcfocus_corrected_powers_df_forsite.loc[:,
                           dcfocus_corrected_powers_df_forsite.columns.str.contains('Inverter AC')]
        ac_power_results['Site'] = [ac_power_results.loc[i, :].sum() for i in ac_power_results.index]

        ideal_power_results = dcfocus_corrected_powers_df_forsite.loc[:,
                                 dcfocus_corrected_powers_df_forsite.columns.str.contains('Ideal')]
        ideal_power_results['Site'] = [ideal_power_results.loc[i, :].sum() for i in
                                          ideal_power_results.index]

        site_pr = ac_power_results['Site'] / ideal_power_results['Site']

        dcfocus_corrected_monthly_pr_df.insert(len(dcfocus_corrected_monthly_pr_df.columns) - 1, 'Site PR %',
                                                 site_pr)

        return dcfocus_corrected_monthly_pr_df,dcfocus_corrected_powers_df_forsite


    else:
        print('Combination of PR type and granularity not possible')
        sys.exit()

    return




#Transversal functions ------------------------------------------------------------------------------

def choose_period_of_analysis(granularity_avail, month_analysis: int = 1, year_analysis: int = 0):
    ''' input: option = ["mtd", "ytd", "monthly", "choose"], month_analysis, year_analysis

    output: start_date, end_date
    '''

    possible_granularity_avail = ["mtd", "ytd", "monthly", "choose"]
    current_day = (datetime.now()).day


    if current_day == 1:
        actual_date = datetime.now() - dt.timedelta(days=1)
        year = actual_date.year
        month = actual_date.month
        day = actual_date.day
    else:
        year = datetime.now().year
        month = datetime.now().month
        day = current_day

    if granularity_avail == "mtd":

        date_start_str = str(year) + "-" + str(month) + "-01"
        if day < 10 and month < 10:
            date_end_str = str(year) + "-0" + str(month) + "-0" + str(day)
        elif day < 10:
            date_end_str = str(year) + "-" + str(month) + "-0" + str(day)
        elif month < 10:
            date_start_str = str(year) + "-0" + str(month) + "-01"
            date_end_str = str(year) + "-0" + str(month) + "-" + str(day)
        else:
            date_end_str = str(year) + "-" + str(month) + "-" + str(day)

    elif granularity_avail == "ytd":

        date_start_str = str(year) + "-01-01"
        if month<10:
            month = "0" + str(month)

        if day < 10:
            date_end_str = str(year) + "-" + str(month) + "-0" + str(day)
        else:
            date_end_str = str(year) + "-" + str(month) + "-" + str(day)

    elif granularity_avail == "monthly":
        date_start_str = mf.input_date(startend="start")
        date_start = datetime.strptime(date_start_str, '%Y-%m-%d')


        if year_analysis != 0:
            year = year_analysis
        else:
            year = date_start.year

        month = date_start.month
        if month < 12:
            day_end = (dt.date(year,month+1 ,1) - dt.timedelta(days=1)).day
        else:
            day_end = 31

        if month<10 and month > 0:
            month = "0" + str(month)
        elif month >=10 and month <=12:
            pass
        else:
            print("Month chosen invalid, you chose: " + str(month) + " \n please chose a number between 1 and 12")

        date_start_str = str(year) + "-" + str(month) + "-01"
        date_end_str = str(year) + "-" + str(month) + "-" + str(day_end)


    elif granularity_avail == "choose":
        date_start_str = mf.input_date(startend="start")
        date_end_str = mf.input_date(startend="end")
        # using custom start dates
    else:
        print(
            "Invalid input from period of availability calculation. You entered: " + str(granularity_avail) + ". \n Please choose one of the following ['mtd', 'ytd', 'custom', 'choose'].")

    return date_start_str, date_end_str


"""def calculate_site_actual_and_expected_power(inverter_list, all_inverter_power_data_dict, site_info, general_info,granularity):
    possible_gran = ['daily', 'monthly']

    if granularity not in possible_gran:
        print('Possible PR types: ' + str(possible_gran) + "\n Your input: " + str(granularity))
        print('Please try again. :)')
        sys.exit()

    days_under_analysis = site_info['Days']
    months_under_analysis = site_info['Months']

    for inverter in inverter_list:
        # print(inverter)
        power_data = all_inverter_power_data_dict[inverter]['Power Data']
        power_data['Day'] = pd.to_datetime(power_data['Timestamp']).dt.date
        power_data['Month'] = pd.to_datetime(power_data['Timestamp']).dt.month"""





#Misc scripts

def set_time_of_operation(reportfiletemplate, site_list, date):

    df_info_sunlight = pd.read_excel(reportfiletemplate, sheet_name='Info', engine="openpyxl")
    ignore_site = ['HV Almochuel']

    for site in site_list:
        if "HV Almochuel" in site:
            print(site + ' was ignored')
            pass
        else:
            index_array = df_info_sunlight[df_info_sunlight['Site'] == site].index.values
            index = int(index_array[0])

            stime, etime = mf.input_time_operation_site(site, date)

            df_info_sunlight.loc[index, 'Time of operation start'] = stime
            df_info_sunlight.loc[index, 'Time of operation end'] = etime


    return df_info_sunlight

def substitute_false_closed_event(df):
    df_simple = df[['Site Name','Related Component','Component Status','Event Start Time','Event End Time']]
    header = df_simple.columns.tolist()
    data = df_simple.values.tolist()
    site = data[0][0]
    component = data[0][1]

    sg.theme('DarkAmber')
    layout = [[sg.Text('Are these events supposed to be an active event?')],
              [sg.Table(values=data,headings=header,display_row_numbers=True,auto_size_columns=False)],
              [sg.Button('Yes'),sg.Button('No'), sg.Exit()]]

    window = sg.Window('Table', layout, grab_anywhere=False)
    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == 'Exit':  # if user closes window or clicks exit
            break
        if event == 'Yes':
            timestamp = mf.input_date_and_time()
            old_events = df
            open_event = df.iloc[0:1]
            open_event['Event Start Time'] = timestamp
            open_event['Rounded Event Start Time'] = timestamp
            open_event['Event End Time'] = ''
            open_event['Rounded Event End Time'] = ''
            open_event['Comments'] = '• ' + component +' is not producing since ' + str(timestamp.day) + '/' + str(timestamp.month)
            open_event['Status of incident'] = 'Active'
            open_event['Duration (h)'] = ''
            break
        if event == 'No':
            print('Closed Events are correct for component: ' + component + ' on ' + site)
            open_event = pd.DataFrame(columns = header)
            old_events = pd.DataFrame(columns = header)
            break
    window.close()

    return open_event, old_events

def get_destcell_formula_irradiancecheck(dict_all_columns_start_cell: dict , irradiance_lossesfile_column_name : str,
                                         row_num):

    formula_start_cell = dict_all_columns_start_cell[irradiance_lossesfile_column_name]
    formula_rowindex, formula_cell_letter = mf.get_rowindex_and_columnletter(formula_start_cell)
    formula_cell = formula_cell_letter + str(row_num)
    formula = "=isblank(" + formula_cell + ")"

    dest_start_cell = dict_all_columns_start_cell['Irradiance Check']
    dest_rowindex, dest_cell_letter = mf.get_rowindex_and_columnletter(dest_start_cell)
    dest_cell = dest_cell_letter + str(row_num)

    return dest_cell, formula

def read_time_of_operation(irradiance_df,Report_template_path , withmean: bool = False):

    df_info_capacity = pd.read_excel(Report_template_path, sheet_name='Info', engine="openpyxl")

    irradiance_df = irradiance_df.loc[:, ~irradiance_df.columns.str.contains('^Unnamed')]
    irradiance_df = irradiance_df[:-1]    #removes last timestamp of dataframe, since it's always the next day

    irradiance_df['day'] = [datetime.strptime(str(timestamp), '%Y-%m-%d %H:%M:%S').day for timestamp in
                                   irradiance_df['Timestamp']]
    only_days = irradiance_df['day'].drop_duplicates().tolist()

    try:
        del df_all
        print('Existing variable deleted. All good now, starting compilation')
    except NameError:
        print('All good, starting compilation')

    irradiance_file_data_curated = irradiance_df.loc[:, irradiance_df.columns.str.contains('curated')]

    irradiance_file_data_notcurated = irradiance_df.loc[:, ~irradiance_df.columns.str.contains('curated')]
    irradiance_file_data_poaaverage = irradiance_file_data_notcurated.loc[:,
                                      irradiance_file_data_notcurated.columns.str.contains('Average')]

    irradiance_file_data_meteostation = irradiance_df.loc[:, irradiance_df.columns.str.contains('Meteo')]


    for column in irradiance_file_data_poaaverage.columns:
        dict_timeofops = {}
        dict_timeofops_seconds = {}

        only_name_site = re.search(r'\[.+\]', column).group().replace('[', "").replace(']', "")
        only_name_site = mf.correct_site_name(only_name_site)

        print(only_name_site)

        capacity = float(df_info_capacity.loc[df_info_capacity['Site'] == only_name_site]['Capacity'])

        # Get curated data
        try:
            curated_column = irradiance_file_data_curated.loc[:,irradiance_file_data_curated.columns.str.contains(only_name_site)].columns[0]
        except IndexError:
            curated_column = column

        if not column == 'Timestamp' and not column == 'day':
            data = irradiance_df[['Timestamp', curated_column, 'day']]
        else:
            continue

        backup_data = irradiance_df[['Timestamp', column, 'day']]

        # print(name_site)
        dict_timeofops['Site'] = only_name_site
        print(dict_timeofops)
        for day in only_days:
            print("Day under analysis: " + str(day) )
            data_day = data.loc[data['day'] == day].reset_index()
            entire_day = data['Timestamp'][0]
            entire_day = datetime.strptime(str(entire_day), '%Y-%m-%d %H:%M:%S').date()

            #print(entire_day)


            try:
                stime_index = next(i for i, v in enumerate(data_day[curated_column]) if v > 20)
                etime_index = next(i for i, v in reversed(list(enumerate(data_day[curated_column]))) if v > 20)

                stime = data_day['Timestamp'][stime_index]
                etime = data_day['Timestamp'][etime_index]

                # Verify Hours read------------------------------
                stime, etime = mf.verify_read_time_of_operation(only_name_site, entire_day, stime, etime)

                # -------------------------------------------------

            except StopIteration:
                print('No data on the ' + str(entire_day))
                try:
                    stime_index = next(i for i, v in enumerate(backup_data[column]) if v > 20)
                    etime_index = next(i for i, v in reversed(list(enumerate(backup_data[column]))) if v > 20)

                    stime = data_day['Timestamp'][stime_index]
                    #print(stime)
                    etime = data_day['Timestamp'][etime_index]
                    #print(etime)

                    #print('Verify 2')

                    stime, etime = mf.verify_read_time_of_operation(only_name_site, entire_day, stime, etime)

                except StopIteration:
                    print('No backup data on the ' + str(entire_day))
                    stime, etime = mf.input_time_operation_site(only_name_site, str(entire_day))

            if type(stime) == str:
                stime = datetime.strptime(stime, '%Y-%m-%d %H:%M:%S')
            if type(etime) == str:
                etime = datetime.strptime(etime, '%Y-%m-%d %H:%M:%S')


            dict_timeofops['Capacity'] = [capacity]
            dict_timeofops['Time of operation start'] = [stime]
            dict_timeofops['Time of operation end'] = [etime]

            df_timeofops = pd.DataFrame.from_dict(dict_timeofops)
            # df_timeofops = df_timeofops.set_index('Site')


            try:
                df_all = df_all.append(df_timeofops)
            except (UnboundLocalError,NameError):
                df_all = df_timeofops

    df_info_sunlight = df_all.reset_index(drop=True)
    print(df_info_sunlight)


    if withmean == True:
        df_all = df_all.set_index('Site')
        stime_columns = df_all.columns[df_all.columns.str.contains('sunrise')].tolist()
        etime_columns = df_all.columns[df_all.columns.str.contains('sunset')].tolist()

        stime_data = df_all.loc[:, stime_columns]
        etime_data = df_all.loc[:, etime_columns]

        for index, row in stime_data.iterrows():
            timestamps = row.tolist()
            timestamps_datetime = [datetime.strptime(timestamp, '%Y-%m-%d %H:%M:%S') for timestamp in timestamps if
                                   timestamp != 'No data']
            in_seconds = [(i.hour * 3600 + i.minute * 60 + i.second) for i in timestamps_datetime if i != 'No data']
            average_in_seconds = int(statistics.mean(in_seconds))
            average_in_hours = datetime.fromtimestamp(average_in_seconds - 3600).strftime("%H:%M:%S")

            df_all.loc[index, 'Mean Start Time'] = average_in_hours

        for index, row in etime_data.iterrows():
            timestamps = row.tolist()
            timestamps_datetime = [datetime.strptime(timestamp, '%Y-%m-%d %H:%M:%S') for timestamp in timestamps if
                                   timestamp != 'No data']
            in_seconds = [(i.hour * 3600 + i.minute * 60 + i.second) for i in timestamps_datetime if i != 'No data']
            average_in_seconds = int(statistics.mean(in_seconds))
            average_in_hours = datetime.fromtimestamp(average_in_seconds - 3600).strftime("%H:%M:%S")

            df_all.loc[index, 'Mean End Time'] = average_in_hours
        df_info_sunlight = df_all

    return df_info_sunlight, irradiance_file_data_notcurated







#Discontinued -----------------------------------------------------------------

def get_final_dataframes_to_add_to_EventTracker_old(df_to_add,df_event_tracker, fmeca_data):

    '''From the different dataframe dictionaries available (New Reports to Add, Event Tracker info and FMECA data,
    creates dict with final dataframes to add.
    Events are verified, excludes from new additions any incident already on Event Tracker and removes from
    active sheet any closed incident.
    Returns: df_final_to_add'''

    # Choose which incidents to add to event tracker
    for sheet in df_to_add.keys():
        print(sheet)
        if "Closed" in sheet:
            df_toadd_id = df_to_add[sheet]['ID'].to_list()  # Get dataframe to add - Closed
            df_ET_id = df_event_tracker[sheet]['ID'].to_list()  # Get dataframe event tracker - Closed

            set_df_ET_id = set(df_ET_id)
            df_toadd_id_tokeep = [x for x in df_toadd_id if x not in set_df_ET_id]

            df_to_add[sheet] = df_to_add[sheet][df_to_add[sheet]['ID'].isin(df_toadd_id_tokeep)].reset_index(drop=True)

            # df_to_add_id = list(set(df_id) - set(df_ET_id))

        else:
            other_sheet = sheet.replace('Active', 'Closed')

            df_toadd_id = df_to_add[sheet]['ID'].to_list()  # Get active dataframe to add and all others

            df_closed_id = df_to_add[other_sheet]['ID'].to_list()
            df_ET_id = df_event_tracker[sheet]['ID'].to_list()
            df_ET_closed_id = df_event_tracker[other_sheet]['ID'].to_list()
            all_ids = df_closed_id + df_ET_id + df_ET_closed_id

            set_all_ids = set(all_ids)
            df_toadd_id_tokeep = [x for x in df_toadd_id if x not in set_all_ids]

            df_to_add[sheet] = df_to_add[sheet][df_to_add[sheet]['ID'].isin(df_toadd_id_tokeep)].reset_index(drop=True)





    # Join new events with events from event tracker and sort them
    final_df_to_add = {}
    for sheet in df_to_add.keys():
        if not df_to_add[sheet].empty:
            new_df = pd.concat([df_event_tracker[sheet], df_to_add[sheet]]).sort_values(
                by=['Site Name', 'Event Start Time', 'Related Component'], ascending=[True, False, False],
                ignore_index=True)
        else:
            new_df = df_event_tracker[sheet]

        final_df_to_add[sheet] = new_df

    # Correct final active lists to exclude already closed events
    for sheet in final_df_to_add.keys():
        if "Active" in sheet:
            other_sheet = sheet.replace('Active', 'Closed')
            df_active = final_df_to_add[sheet]
            df_active_id = final_df_to_add[sheet]['ID'].to_list()

            df_closed = final_df_to_add[other_sheet]
            df_closed_id = final_df_to_add[other_sheet]['ID'].to_list()

            set_closed_ids = set(df_closed_id)
            df_tokeep_id = [x for x in df_active_id if x not in set_closed_ids]
            df_toremove_id = [x for x in df_active_id if x in set_closed_ids]

            for id_incident in df_toremove_id:
                index_closed = int(df_closed.loc[df_closed['ID'] == id_incident].index.values)
                index_active = int(df_active.loc[df_active['ID'] == id_incident].index.values)
                df_closed.loc[index_closed, 'Remediation'] = df_active.loc[index_active, 'Remediation']
                df_closed.loc[index_closed, 'Fault'] = df_active.loc[index_active, 'Fault']
                df_closed.loc[index_closed, 'Fault Component'] = df_active.loc[index_active, 'Fault Component']
                df_closed.loc[index_closed, 'Failure Mode'] = df_active.loc[index_active, 'Failure Mode']
                df_closed.loc[index_closed, 'Failure Mechanism'] = df_active.loc[index_active, 'Failure Mechanism']
                df_closed.loc[index_closed, 'Category'] = df_active.loc[index_active, 'Category']
                df_closed.loc[index_closed, 'Subcategory'] = df_active.loc[index_active, 'Subcategory']
                df_closed.loc[index_closed, 'Resolution Category'] = df_active.loc[index_active, 'Resolution Category']

            final_df_to_add[sheet] = final_df_to_add[sheet][
                final_df_to_add[sheet]['ID'].isin(df_tokeep_id)].reset_index(drop=True)
            final_df_to_add[other_sheet] = df_closed
        else:
            pass

        # final_df_to_add[sheet]
    for sheet, df in final_df_to_add.items():
        if "Active" in sheet:
            df['Event Start Time'] = [datetime.strptime(str(timestamp), '%Y-%m-%d %H:%M:%S') for timestamp in
                                   df['Event Start Time']]
        else:
            df['Event Start Time'] = [datetime.strptime(str(timestamp), '%Y-%m-%d %H:%M:%S') for timestamp in
                                       df['Event Start Time']]
            df['Event End Time'] = [datetime.strptime(str(timestamp), '%Y-%m-%d %H:%M:%S') for timestamp in
                                   df['Event End Time']]
        final_df_to_add[sheet] = df

    final_df_to_add['FMECA'] = fmeca_data
    final_df_to_add = dict(sorted(final_df_to_add.items()))


    return final_df_to_add