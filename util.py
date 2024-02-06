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