import util
import os

from tkinter import filedialog, messagebox, Toplevel

import pandas as pd
import tkinter as tk

from application import Application

from analyzer import Analyzer

#Check if running Executable or Python Script
application_path = util.CheckExecutable()

# Build the full path to the GIF file
gif_path = os.path.join(application_path, 'math.gif')

# Create an instance of Application
app = Application(gif_path)

# Open a file dialog for the user to select a file
file_paths = app.open_files()

def process_outcomes():
    """calculate daily outcomes for each day and save a sum of outcomes for each day separately"""
    # Create an instance of tc_dict
    tc_dict_instance = tc_dict.copy()

    for index, row in history_worksheet.iterrows():
        date = row['Date']
        tc_id = row['TestCaseID']
        outcome = row['Outcome']

        # Update the tc_dict instance
        tc_dict_instance[tc_id] = outcome

        # Replace the tc_dict in date_tc_outcome_dict with the updated tc_dict instance
        date_tc_outcome_dict[date] = tc_dict_instance.copy()

    #output an outcome count table
    result_dict = {}

    for date in date_tc_outcome_dict:
        result_dict[date] = {outcome: 0 for outcome in outcome_set}

        for tc_id, outcome in date_tc_outcome_dict[date].items():
            if outcome in outcome_set:
                result_dict[date][outcome] += 1
    return result_dict

for file_path in app.tk.splitlist(file_paths):

    # Get the directory and the base name of the selected file
    dir_name = os.path.dirname(file_path)
    base_name = os.path.basename(file_path)

    # Construct the output file name
    output_file_name = os.path.splitext(base_name)[0] + ' - Daily Outcomes.csv'
    output_file_path = os.path.join(dir_name, output_file_name)

    #Load the Excel file
    analyzer = Analyzer(file_path, "History_Worksheet")

    #DATA CLEANUP
    #Convert Date to datetime
    analyzer.convert_date_to_datetime('Date')

    #Rename Columns
    analyzer.rename_history_columns('SpecificRuns1WeekOutcome.Outcome.Column1.outcome', 'Outcome')
    analyzer.rename_history_columns('SpecificRuns1WeekOutcome.Outcome.Column1.testCase.id', 'TestCaseID')

    #Standardize Test Case ID Column
    analyzer.standardize_testcaseid()

    #Temp assignment for ongoing changes to OOP
    history_worksheet = analyzer.history

    history_worksheet['Outcome'] = history_worksheet['Outcome'].fillna('Active')

    history_worksheet = history_worksheet[[
        'Date',
        'Outcome', 
        'TestCaseID']]
    
    # Calculate the start and finish dates
    start_date = history_worksheet['Date'].min()
    finish_date = history_worksheet['Date'].max()

    # Generate all dates between the start and finish dates
    project_dates = pd.date_range(start_date, finish_date)



    # Load the 'Relationship Download' sheet
    relationship_download = pd.read_excel(file_path, sheet_name='Relationship Download', engine='openpyxl')

    # Filter the 'ID' column by 'Test Case' in the 'Work Item Type' column
    all_tcs = relationship_download[relationship_download['Work Item Type'] == 'Test Case']

    all_tcs['ID'] = all_tcs['ID'].astype(int)
    all_tcs = all_tcs[['ID', 'Outcome']]


    tc_dict = all_tcs.set_index('ID').assign(Outcome='Active')['Outcome'].to_dict()
    
    outcome_set = {"Active", "NotApplicable", "Blocked", "Failed", "Passed"}
    project_outcomes = []

    date_tc_outcome_dict = {}

    for date in project_dates:
        date_tc_outcome_dict[date] = tc_dict
    
    outcome_df = pd.DataFrame(process_outcomes()).T

    # TODO : Implement Team's sharepoint connectivity
    # TODO : Validate which columns get created first in the outcome_df
    # TODO : organize the analyzer around actual project dates so that empty execution dates are saved correctly using previous day's data
    

    print("print: dates dictionary \n \n")
    print(project_dates)
    print("print: all test cases dictionary \n \n")
    print(tc_dict)
    print("print: all test cases and dates dictionary \n \n")
    print(date_tc_outcome_dict)
    print("print: specific date of the dictionary")
    print(date_tc_outcome_dict[project_dates[0]])
    print("print: result dict \n \n")
    print(outcome_df)

    try:
        # Save the outcome dataframe to a CSV file
        outcome_df.to_csv(output_file_path, index_label='Date')
        #history_worksheet.to_csv('history_worksheet.csv', index=False)

        # Create a simple pop-up window that confirms the completion of the task
        messagebox.showinfo("Confirmation", "The task has been completed successfully!")
        app.destroy()  # Destroy the main window
    except PermissionError:
        messagebox.showinfo("Permission Error", "Unable to write to file. Please make sure the file is not open in another program and try again.")
        app.destroy()  # Destroy the main window



