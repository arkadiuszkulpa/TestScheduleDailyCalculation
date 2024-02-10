import sys
import os

def CheckExecutable():
    """Check if Running Executable or Python script"""
    if getattr(sys, 'frozen', False):
        # We're running as a packaged executable
        return sys._MEIPASS
    else:
        # We're running as a normal Python script
        return os.path.dirname(os.path.abspath(__file__))

def construct_output_file_path(filePath):
    # Get the directory and the base name of the selected file
    dir_name = os.path.dirname(filePath)
    base_name = os.path.basename(filePath)

    # Construct the output file name
    output_file_name = os.path.splitext(base_name)[0] + ' - Daily Outcomes.csv'
    return os.path.join(dir_name, output_file_name)