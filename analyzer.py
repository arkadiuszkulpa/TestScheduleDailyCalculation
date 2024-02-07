import pandas as pd


class Analyzer():
    """Analyze Data from the Test Schedule"""

    def __init__(self, filePath, history_worksheet_name,):
        """initialize with history_Worksheet"""
        self.file_path = filePath
        self.history = pd.read_excel(filePath, 
        sheet_name=history_worksheet_name, engine='openpyxl')
        self.project_dates = []

    def trim_data(self):
        """Trim to only required columns, perform after rename"""
        self.history = self.history[[
            'Date',
            'Outcome',
            'TestCaseID'
        ]]

    def convert_date_to_datetime(self, date_column):
        """Convert date to datetime format"""
        self.history[date_column] = pd.to_datetime(self.history[date_column])

    def rename_history_columns(self, old_name, new_name):
        """Change Column name, string, string"""
        if old_name not in self.history.columns:
            raise ValueError(f"Column {old_name} not found in history")
        if new_name in self.history.columns:
            raise ValueError(f"Column {new_name} already exists in history")
        else:
            print(f"Renaming {old_name} to {new_name}")
            self.history = self.history.rename(columns={
                old_name: new_name
            })
        
    def standardize_columns(self):
        """Standardize Test Case ID"""
        self.history['TestCaseID'] = pd.to_numeric(self.history['TestCaseID'], errors='coerce')
        self.history = self.history.dropna(subset=['TestCaseID'])
        self.history['TestCaseID'] = self.history['TestCaseID'].astype(int)
        self.history['Outcome'] = self.history['Outcome'].fillna('Active')
        
    def generate_project_dates(self):
        """Generates a list range of dates from min max of the entries in history_worksheet"""
        # Calculate the start and finish dates
        start_date = self.history['Date'].min()
        finish_date = self.history['Date'].max()

        # Generate all dates between the start and finish dates
        self.project_dates = pd.date_range(start_date, finish_date)