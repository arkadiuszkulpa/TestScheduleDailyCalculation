import pandas as pd

# Load the workbook
xl = pd.read_excel('R4 - TA - Test Schedule R5 v2.1 - CA 4.5.xlsm', sheet_name='History_Worksheet', engine='openpyxl')

# Convert the date column to datetime
xl['Date'] = pd.to_datetime(xl['Date'])

# Filter rows where date is 26/01/2024
filtered_df = xl[xl['Date'] == '2024-01-26']

# Print the filtered data
print(filtered_df)

# Write the filtered data to a new CSV file
filtered_df.to_csv('FilteredData.csv', index=False)