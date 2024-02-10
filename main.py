import util
import os

from tkinter import filedialog, messagebox, Toplevel
import pandas as pd
import tkinter as tk

from application import Application
from analyzer import Analyzer

# Create an instance of Application
app = Application(os.path.join(util.CheckExecutable(), 'math.gif'))

# Open a file dialog for the user to select a file
file_paths = app.open_files()

for file_path in app.tk.splitlist(file_paths):

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
        outcome_df.to_csv(util.construct_output_file_path(file_path), index_label='Date')
        #history_worksheet.to_csv('history_worksheet.csv', index=False)

        # Create a simple pop-up window that confirms the completion of the task
        messagebox.showinfo("Confirmation", "The task has been completed successfully!")
        app.destroy()  # Destroy the main window
    except PermissionError:
        messagebox.showinfo("Permission Error", "Unable to write to file. Please make sure the file is not open in another program and try again.")
        app.destroy()  # Destroy the main window



