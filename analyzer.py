import pandas as pd


class Analyzer():
    """Analyze Data from the Test Schedule"""

    def __init__(self, filePath, history_worksheet_name,):
        """initialize with history_Worksheet"""
        self.file_path = filePath
        self.history = pd.read_excel(filePath, sheet_name=history_worksheet_name, engine='openpyxl')


    def convert_date_to_datetime(self):
        self.history['Date'] = pd.to_datetime(self.history['Date'])
