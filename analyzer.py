import pandas as pd
import copy

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
        self.removedTCs = []
        self.tc_dict = []
        self.outcome_set = sorted(list({"Active", "Paused", "NotApplicable", "Blocked", "Failed", "Passed"}))
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
        print(f'identify_all_tcs: \n{self.alltcs}')

    def standardize_rel_columns(self):
        """turn ID column into Int from float"""
        self.alltcs['ID'] = pd.to_numeric(self.alltcs['ID'], errors='coerce')
        self.alltcs['TC Complexity'] = pd.to_numeric(self.alltcs['TC Complexity'], errors='coerce')
        self.alltcs['Priority'] = pd.to_numeric(self.alltcs['Priority'], errors='coerce')

        self.alltcs.dropna(subset=['ID'], inplace=True)
        self.alltcs.dropna(subset=['TC Complexity'], inplace=True)
        self.alltcs.dropna(subset=['Priority'], inplace=True)

        self.alltcs['ID'] = self.alltcs['ID'].astype(int)
        self.alltcs['TC Complexity'] = self.alltcs['TC Complexity'].astype(int)
        self.alltcs['Priority'] = self.alltcs['Priority'].astype(int)

        print(f'standardize_rel_columns: \n{self.alltcs}')
    
    def trim_relationship_data(self):
        """trim the data to only relevant columns"""
        self.alltcs = self.alltcs[['ID', 'Outcome', 'TC Complexity', 'Priority']]
        print(f'trim_relationship_data: \n{self.alltcs}')

    def set_all_active(self):
        """Set all test cases to Active"""
        self.alltcs.loc[:, 'Outcome'] = 'Active'
        self.tc_dict = self.alltcs.set_index('ID').to_dict('index')
        for date in self.project_dates:
            self.date_tc_outcome_dict[date] = copy.deepcopy(self.tc_dict)
        print(f'set_all_active - alltcs: \n{self.alltcs.iloc[0]}')
        first_item_key, first_item_value = list(self.tc_dict.items())[0]
        print(f'set_all_active - tc_dict: \n{first_item_key} , {first_item_value}')

    def analyze_outcomes(self):
        """Analyze Outcomes to output a dataframe"""
        
        # Create an instance of tc_dict
        tc_dict_instance = copy.deepcopy(self.tc_dict)

        for date in self.project_dates:
            # Filter history for entries that match the current date
            history_for_date = self.history[self.history['Date'] == date]

            for index, row in history_for_date.iterrows():
                date = row['Date']
                tc_id = row['TestCaseID']
                outcome = row['Outcome']

                # Check if tc_id is a key in tc_dict_instance
                if tc_id not in tc_dict_instance:
                    # If not, add it to self.removedTCs
                    self.removedTCs.append(tc_id)
                else:
                    # Update the 'Outcome' field of the dictionary at tc_dict_instance[tc_id]
                    tc_dict_instance[tc_id]['Outcome'] = outcome

            # Replace the tc_dict in date_tc_outcome_dict with the updated tc_dict instance
            self.date_tc_outcome_dict[date] = copy.deepcopy(tc_dict_instance)
        # first_item_key, first_item_value = list(self.date_tc_outcome_dict.items())[0]
        # print(f'analyze_outcomes - date_tc_outcome_dict: \n{first_item_key} , {first_item_value}')
    
    def output_outcome_table(self):
        """Outputs an outcome count table for each date"""
        result_dict = {}

        for date, tc_outcome_dict in self.date_tc_outcome_dict.items():
            result_dict[date] = {outcome: 0 for outcome in self.outcome_set}

            for tc_id, outcome in tc_outcome_dict.items():
                if outcome in self.outcome_set:
                    result_dict[date][outcome] += 1
        print(f'output_outcome_table: {result_dict[0]}')
        return result_dict
    
    def output_complex_outcome_table(self):
        """Outputs a complex outcome count table for each date"""
        result_dict = {}

        for date, tc_dict in self.date_tc_outcome_dict.items():
            result_dict[date] = {f"C{i}P{j} - {outcome}": 0 for i in range(1, 5) for j in range(1, 5) for outcome in self.outcome_set}
            result_dict[date].update({f"{outcome}": 0 for outcome in self.outcome_set})

            for tc_id, tc_attributes in tc_dict.items():
                print(f"tc_id: {tc_id}, type: {type(tc_attributes)}, value: {tc_attributes}")
                complexity_priority_key = f"C{tc_attributes['TC Complexity']}P{tc_attributes['Priority']}"
                outcome = tc_attributes['Outcome']

                if outcome in self.outcome_set:
                    result_dict[date][f"{complexity_priority_key} - {outcome}"] += 1
                    result_dict[date][outcome] += 1
        return result_dict