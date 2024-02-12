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

    #History Data Cleanup
    analyzer.convert_date_to_datetime('Date')
    analyzer.rename_history_columns('SpecificRuns1WeekOutcome.Outcome.Column1.outcome', 'Outcome')
    analyzer.rename_history_columns('SpecificRuns1WeekOutcome.Outcome.Column1.testCase.id', 'TestCaseID')
    analyzer.standardize_columns()
    analyzer.trim_history_data()
    analyzer.generate_project_dates()

    #Relationship Data cleanup
    analyzer.identify_all_tcs()
    analyzer.trim_relationship_data()
    analyzer.standardize_rel_columns()
    analyzer.set_all_active()
    analyzer.analyze_outcomes()

    outcome_df = pd.DataFrame(analyzer.output_complex_outcome_table()).T

    # TODO : Implement Team's sharepoint connectivity
    # TODO : Validate which columns get created first in the outcome_df
    # TODO : organize the analyzer around actual project dates so that empty execution dates are saved correctly using previous day's data

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



