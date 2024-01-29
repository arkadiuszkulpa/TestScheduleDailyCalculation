import os

import pandas as pd
import tkinter as tk

from tkinter import filedialog, messagebox, Toplevel

from application import Application

# Path to the GIF file
gif_path = 'math.gif'

# Create an instance of Application
app = Application(gif_path)

# Open a file dialog for the user to select a file
file_paths = filedialog.askopenfilenames()
if not file_paths:
    messagebox.showinfo("No file selected", "You did not select a file.")
else:
    # Create a Toplevel window
    app.text.insert(tk.END, "Processing files...\n")
    app.text.update()



for file_path in app.tk.splitlist(file_paths):

    # Get the directory and the base name of the selected file
    dir_name = os.path.dirname(file_path)
    base_name = os.path.basename(file_path)

    # Construct the output file name
    output_file_name = os.path.splitext(base_name)[0] + ' - Daily Outcomes.csv'
    output_file_path = os.path.join(dir_name, output_file_name)

    # Load the workbook
    history_worksheet = pd.read_excel(file_path, sheet_name='History_Worksheet', engine='openpyxl')

    # Convert the date column to datetime
    history_worksheet['Date'] = pd.to_datetime(history_worksheet['Date'])

    # Simplify the History_worksheet dataset to contain only Date, Outcome, Test Case ID
    history_worksheet = history_worksheet.rename(columns={
        'SpecificRuns1WeekOutcome.Outcome.Column1.outcome': 'Outcome',
        'SpecificRuns1WeekOutcome.Outcome.Column1.testCase.id': 'TestCaseID'
    })

    # Convert non-numeric values to NaN
    history_worksheet['TestCaseID'] = pd.to_numeric(history_worksheet['TestCaseID'], errors='coerce')

    # Remove rows with NaN values in 'TestCaseID' column
    history_worksheet = history_worksheet.dropna(subset=['TestCaseID'])

    # Convert 'TestCaseID' to integer
    history_worksheet['TestCaseID'] = history_worksheet['TestCaseID'].astype(int)

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

# Take an instance of the tc_dict, then for each row of history_worksheet, find the corresponding date in date_tc_outcome_dict and the corresponding ID dictionary within it, then update the tc_dict instance then replace date_tc_outcome_dict's tc_dict with the instanced tc_dict and move onto next row of history_worksheet. 
        
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

        app.text.insert(tk.END, f"TestCaseID: {tc_id}, Outcome: {outcome}\n")
        #output an outcome count table
        outcome_counts = {}

        for date in date_tc_outcome_dict:
            outcome_counts[date] = {outcome: 0 for outcome in outcome_set}

            for tc_id, outcome in date_tc_outcome_dict[date].items():
                if outcome in outcome_set:
                    outcome_counts[date][outcome] += 1

    outcome_df = pd.DataFrame(outcome_counts).T

    # print("print: history worksheet filtered columns \n \n")
    # print(history_worksheet)
    # print("print: dates dictionary \n \n")
    # print(project_dates)
    # print("print: all test cases dictionary \n \n")
    # print(tc_dict)
    # print("print: all test cases and dates dictionary \n \n")
    # print(date_tc_outcome_dict)
    # print("print: specific date of the dictionary")
    # print(date_tc_outcome_dict[project_dates[0]])

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



