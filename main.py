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
    try:
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
        # Save the outcome dataframe to a CSV file
        outcome_df.to_csv(util.construct_output_file_path(file_path), index_label='Date')
        #history_worksheet.to_csv('history_worksheet.csv', index=False)

        # Create a simple pop-up window that confirms the completion of the task
        messagebox.showinfo("Confirmation", "The task has been completed successfully!")
        app.destroy()  # Destroy the main window
    except Exception as e:
        # # If an error occurs, display a message box with the error message
        # root = tk.Tk()
        # root.withdraw()  # Hide the main window
        # messagebox.showerror("Error", str(e))
        # root.mainloop()  # Keep the terminal open until the messagebox is clicked
        # app.destroy()  # Destroy the main window
        messagebox.showerror("error", str(e))
        app.destroy()  # Destroy the main window



