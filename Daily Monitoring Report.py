#! path\to\interpreter\python.exe
import PySimpleGUI as sg
import dmrprocess1
import dmrprocess2
import os

def main():
    sg.theme('DarkAmber')  # Add a touch of color
    # All the stuff inside your window.
    layout = [[sg.Text('Welcome to the DMR tool, what do you want to do?', pad=((2, 10), (2, 5)))],
              [sg.Button('Create Incidents List'), sg.Button('Create final report'), sg.Exit()]]

    # Create the Window
    window = sg.Window('Daily Monitoring Report', layout)
    # Event Loop to process "events" and get the "values" of the inputs
    while True:
        event, values = window.read(timeout=100)

        if event == sg.WIN_CLOSED or event == 'Exit':  # if user closes window or clicks exit
            break
        if event == 'Create Incidents List':
            incidents_file, tracker_incidents_file, site_list, geography, date, df_component_code = dmrprocess1.main()
        if event == 'Create final report':
            try:
                dmr_report = dmrprocess2.main(incidents_file, tracker_incidents_file, site_list, geography, date)
                #dmr_report = dmrprocess2_fromscratch.main(incidents_file,tracker_incidents_file, site_list, geography, date)
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
                    break


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

