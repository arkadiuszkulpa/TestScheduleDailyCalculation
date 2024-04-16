import tkinter as tk
from tkinter import filedialog, messagebox

class Application(tk.Tk):
    """define window background image as gif"""
    def __init__(self, gif_path):
        tk.Tk.__init__(self)
        self.geometry("500x500")
        self.title("ADO API Outcome Analyzer")
        self.configure(bg="black")  # Set the background color to black
        # Load the GIF
        self.frames = []

        for i in range(1000):  # Increase the range if your GIF has more frames
            try:
                frame = tk.PhotoImage(file=gif_path, format='gif -index %i' % i)
                self.frames.append(frame)
            except tk.TclError:
                break

        # Define frame_index
        self.frame_index = 0

        self.label = tk.Label(self)
        self.label.pack()

        # Create a Text widget for displaying processing information
        self.text = tk.Text(self, bg="black", fg="white")
        self.text.insert(tk.END, "This program was developed by Arkadiusz Kulpa.\n")
        self.text.insert(tk.END, "Reach out if any issues :).\n")
        self.text.pack()

        self.update_image()

    def update_image(self):
        """Update the gif image displayed in app"""
        if not self.frames:
            print("Error: No frames loaded")
            return
        if self.frame_index >= len(self.frames):
            print(f"Error: Frame index {self.frame_index} out of range")
            return
        self.current_frame = self.frames[self.frame_index]
        self.frame_index = (self.frame_index + 1) % len(self.frames)
        self.label.configure(image=self.current_frame)
        self.after(20, self.update_image)  # Increase the delay to 100 milliseconds

    def open_files(self):
        """Select Files from User Input"""
        # Open a file dialog for the user to select a file
        file_paths = filedialog.askopenfilenames()
        if not file_paths:
            messagebox.showinfo("No file selected", "You did not select a file.")
            exit()
        else:
            self.text.insert(tk.END, "Processing files...\n")
            self.text.update()
            return file_paths


