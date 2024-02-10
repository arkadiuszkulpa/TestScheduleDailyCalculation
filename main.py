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
    tc_dict_instance = analyzer.tc_dict.copy()

    for index, row in history_worksheet.iterrows():
        date = row['Date']
        tc_id = row['TestCaseID']
        outcome = row['Outcome']

        # Update the tc_dict instance
        tc_dict_instance[tc_id] = outcome

        # Replace the tc_dict in date_tc_outcome_dict with the updated tc_dict instance
        analyzer.date_tc_outcome_dict[date] = tc_dict_instance.copy()

    #output an outcome count table
    result_dict = {}

    for date in analyzer.date_tc_outcome_dict:
        result_dict[date] = {outcome: 0 for outcome in analyzer.outcome_set}

        for tc_id, outcome in analyzer.date_tc_outcome_dict[date].items():
            if outcome in analyzer.outcome_set:
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
    analyzer = Analyzer(file_path, 'History_Worksheet', 'Relationship Download')

    #DATA CLEANUP
    #Convert Date to datetime
    analyzer.convert_date_to_datetime('Date')

    #Rename Columns
    analyzer.rename_history_columns('SpecificRuns1WeekOutcome.Outcome.Column1.outcome', 'Outcome')
    analyzer.rename_history_columns('SpecificRuns1WeekOutcome.Outcome.Column1.testCase.id', 'TestCaseID')

    #Standardize Test Case ID Column
    analyzer.standardize_columns()
    analyzer.trim_history_data()
    analyzer.generate_project_dates()

    #Temp assignment for ongoing changes to OOP
    history_worksheet = analyzer.history

    # Filter the 'ID' column by 'Test Case' in the 'Work Item Type' column

    analyzer.identify_all_tcs()
    analyzer.standardize_rel_columns()
    analyzer.trim_relationship_data()
    analyzer.set_all_active()
    analyzer.analyze_outcomes()

    relationship_download = analyzer.relationship

    outcome_df = pd.DataFrame(analyzer.analyze_outcomes()).T

    # TODO : Implement Team's sharepoint connectivity
    # TODO : Validate which columns get created first in the outcome_df
    # TODO : organize the analyzer around actual project dates so that empty execution dates are saved correctly using previous day's data
    

    # print("print: dates dictionary \n \n")
    # print(analyzer.project_dates)
    # print("print: all test cases dictionary \n \n")
    # print(analyzer.tc_dict)
    # print("print: all test cases and dates dictionary \n \n")
    # print(date_tc_outcome_dict)
    # print("print: specific date of the dictionary")
    # print(date_tc_outcome_dict[analyzer.project_dates[0]])
    # print("print: result dict \n \n")
    # print(outcome_df)

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



