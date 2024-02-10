import pandas as pd


class Analyzer():
    """Analyze Data from the Test Schedule"""

    def __init__(self, filePath, history_worksheet_name, relationship_worksheet_name):
        """initialize with history_Worksheet"""
        self.file_path = filePath
        self.history = pd.read_excel(filePath, 
        sheet_name=history_worksheet_name, engine='openpyxl')
        self.project_dates = []
        # Load the 'Relationship Download' sheet
        self.relationship = pd.read_excel(filePath, sheet_name=relationship_worksheet_name, engine='openpyxl')
        self.alltcs = []
        self.tc_dict = []
        self.outcome_set = sorted(list({"Active", "NotApplicable", "Blocked", "Failed", "Passed"}))
        self.project_outcomes = []
        self.date_tc_outcome_dict = {}

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

    def trim_history_data(self):
        """Trim to only required columns, perform after rename"""
        self.history = self.history[[
            'Date',
            'Outcome',
            'TestCaseID'
        ]]

    def generate_project_dates(self):
        """Generates a list range of dates from min max of the entries in history_worksheet"""
        # Calculate the start and finish dates
        start_date = self.history['Date'].min()
        finish_date = self.history['Date'].max()

        # Generate all dates between the start and finish dates
        self.project_dates = pd.date_range(start_date, finish_date)

    def identify_all_tcs(self):
        """Saves a set of all TCs in the project to alltcs"""
        self.alltcs = self.relationship[self.relationship['Work Item Type'] == 'Test Case']

    def standardize_rel_columns(self):
        """turn ID column into Int from float"""
        self.alltcs['ID'] = self.alltcs['ID'].astype(int)
    
    def trim_relationship_data(self):
        """trim the data to only relevant columns"""
        self.alltcs = self.alltcs[['ID', 'Outcome']]

    def set_all_active(self):
        """Set all test cases to Active"""
        self.tc_dict = self.alltcs.set_index('ID').assign(Outcome='Active')['Outcome'].to_dict()

    def analyze_outcomes(self):
        """Analyze Outcomes to output a dataframe"""
        for date in self.project_dates:
            self.date_tc_outcome_dict[date] = self.tc_dict
        # Create an instance of tc_dict
        tc_dict_instance = self.tc_dict.copy()

        for index, row in self.history.iterrows():
            date = row['Date']
            tc_id = row['TestCaseID']
            outcome = row['Outcome']

            # Update the tc_dict instance
            tc_dict_instance[tc_id] = outcome

            # Replace the tc_dict in date_tc_outcome_dict with the updated tc_dict instance
            self.date_tc_outcome_dict[date] = tc_dict_instance.copy()

        #output an outcome count table
        result_dict = {}

        for date in self.date_tc_outcome_dict:
            result_dict[date] = {outcome: 0 for outcome in self.outcome_set}

            for tc_id, outcome in self.date_tc_outcome_dict[date].items():
                if outcome in self.outcome_set:
                    result_dict[date][outcome] += 1
        return result_dict