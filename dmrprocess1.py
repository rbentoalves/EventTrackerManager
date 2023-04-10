# http://www.PySimpleGUI.org

import PySimpleGUI as sg
import myprocesses as mp
import re
import os
def main():
    sg.theme('DarkAmber')  # Add a touch of color
    # All the stuff inside your window.
    layout = [[sg.Text('Enter date of report you want to analyse', pad=((2, 10), (2, 5)))],
              [sg.CalendarButton('Choose date', target='-CAL-', format="%Y-%m-%d"),
               sg.In(key='-CAL-', text_color='black', size=(16, 1), enable_events=True, readonly=True, visible=True)],
              [sg.Text('Choose Alarm report', pad=((0, 10), (10, 2)))],
              [sg.FileBrowse(target='-FILE-'),
               sg.In(key='-FILE-', text_color='black', size=(20, 1), enable_events=True, readonly=True, visible=True)],
              [sg.Text('Choose Irradiance file', pad=((0, 10), (10, 2)))],
              [sg.FileBrowse(target='-IRRFILE-'),
               sg.In(key='-IRRFILE-', text_color='black', size=(20, 1), enable_events=True, readonly=True, visible=True)],
              [sg.Text('Enter geography ', pad=((0, 10), (10, 2)))],
              [sg.Combo(['AUS', 'ES', 'USA'], size=(4, 3), readonly=True, key='-GEO-', pad=((5, 10), (2, 10)))],
              [sg.Button('Create Incidents List'), sg.Exit()]]

    # Create the Window
    window = sg.Window('Daily Monitoring Report', layout, modal=True)
    # Event Loop to process "events" and get the "values" of the inputs
    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == 'Exit':  # if user closes window or clicks exit
            window.close()
            return "No File", "No File", ["No site list"],"PT", "27-03-1996"
            break
        if event == 'Create Incidents List':
            date = values['-CAL-']  # date is string
            Alarm_report_path = values['-FILE-']
            irradiance_file_path = values['-IRRFILE-']

            report_name = os.path.basename(Alarm_report_path)
            geography_report_match = re.search(r'\w+?_', report_name)
            geography_report = geography_report_match.group()[:-1]
            geography = values['-GEO-']

            #print(date[:4])
            #print(date[5:7])
            #print(date[-2:])
            print(Alarm_report_path)
            print(geography)
            print(geography_report)
            if "Daily" and "Alarm" and "Report" in Alarm_report_path and geography == geography_report and\
                    "Irradiance" in irradiance_file_path :

                incidents_file, tracker_incidents_file, site_list, all_component_data = mp.dmr_create_incidents_files(Alarm_report_path,irradiance_file_path,
                                                                                        geography, date)
                sg.popup('All incident files are ready for approval', no_titlebar=True)
                window.close()
                return incidents_file, tracker_incidents_file, site_list, geography, date,all_component_data
                break
            elif not geography == geography_report:
                msg = 'Selected Geography ' + geography + ' does not match geography from report ' + geography_report
                sg.popup(msg, title = "Error with the selections")
                #print('Selected Geography ' + geography + ' does not match geography from report ' + geography_report)
            elif not "Daily" and "Alarm" and "Report" in Alarm_report_path:
                msg='File is not a Daily Alarm Report'
                sg.popup(msg, title = "Error with the selections")
                #print('File is not a Daily Alarm Report')
            elif not "Irradiance" in irradiance_file_path:
                msg = 'File is not a Irradiance file'
                sg.popup(msg, title = "Error with the selections")
                #print('File is not a Daily Alarm Report')
    window.close()

    return

if __name__ == "__main__":
    try:
        main()
    except BaseException as e:
        print('Error: ')
        print(e)
        raise


