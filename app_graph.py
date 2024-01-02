import tkinter as tk
from tkinter import filedialog
from functools import partial
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
import pandas as pd

def process_data(input_file):
    # Your existing script logic here
    df = pd.read_excel(input_file)
    # ... additional processing ...

    # For demonstration purposes, let's just print some information
    print("Data processed successfully!")

def create_pie_chart_with_legend_and_extremes(input_file):
    # Your existing script logic here
    df = pd.read_excel(input_file)
    # ... additional processing ...

    # For demonstration purposes, let's just plot a simple pie chart
    group_counts = df['NoteGroup'].value_counts()
    plt.pie(group_counts, labels=group_counts.index, autopct='%1.1f%%', startangle=140)
    plt.title("Répartition des notes par groupe (Pie Chart)")

    # Display the chart
    plt.show()

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Your App Name")

        # Create GUI components
        self.label = tk.Label(root, text="Select Excel File:")
        self.label.pack()

        self.file_button = tk.Button(root, text="Browse", command=self.browse_file)
        self.file_button.pack()

        self.process_button = tk.Button(root, text="Process Data", command=self.process_data)
        self.process_button.pack()

        self.chart_button = tk.Button(root, text="Show Pie Chart", command=self.show_pie_chart)
        self.chart_button.pack()

        self.file_path = ""

    def browse_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])

    def process_data(self):
        if self.file_path:
            process_data(self.file_path)
        else:
            print("Please select a file first.")

    def show_pie_chart(self):
        if self.file_path:
            create_pie_chart_with_legend_and_extremes(self.file_path)
        else:
            print("Please select a file first.")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
