import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from scipy.signal import welch

class GLevelPSDApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("G-Level and PSD Plotter")
        self.geometry("800x600")

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill='both', expand=True)

        self.input_tab = ttk.Frame(self.notebook)
        self.glevel_tab = ttk.Frame(self.notebook)
        self.psd_tab = ttk.Frame(self.notebook)

        self.notebook.add(self.input_tab, text='Inputs')
        self.notebook.add(self.glevel_tab, text='G-Levels')
        self.notebook.add(self.psd_tab, text='PSD Plots')

        self.create_input_tab()
        self.create_glevel_tab()
        self.create_psd_tab()

        self.data = None
        self.velocity_data = None
        self.sensitivity = None
        self.sampling_freq = None
        self.selected_range = None

    def create_input_tab(self):
        # Labels and Entries for sensitivity and sampling frequency
        ttk.Label(self.input_tab, text="Sensor Sensitivity:").grid(row=0, column=0, padx=10, pady=10)
        self.sensitivity_entry = ttk.Entry(self.input_tab)
        self.sensitivity_entry.grid(row=0, column=1, padx=10, pady=10)

        ttk.Label(self.input_tab, text="Sampling Frequency:").grid(row=1, column=0, padx=10, pady=10)
        self.sampling_freq_entry = ttk.Entry(self.input_tab)
        self.sampling_freq_entry.grid(row=1, column=1, padx=10, pady=10)

        # Buttons for loading data and velocity profile
        ttk.Button(self.input_tab, text="Load CSV/Excel File", command=self.load_file).grid(row=2, column=0, columnspan=2, padx=10, pady=10)
        ttk.Button(self.input_tab, text="Load Velocity Profile", command=self.load_velocity_profile).grid(row=3, column=0, columnspan=2, padx=10, pady=10)

    def create_glevel_tab(self):
        # Placeholder for g-level plot and PSD button
        self.glevel_fig, self.glevel_ax = plt.subplots()
        self.glevel_canvas = FigureCanvasTkAgg(self.glevel_fig, master=self.glevel_tab)
        self.glevel_canvas.get_tk_widget().pack(fill='both', expand=True)

        ttk.Button(self.glevel_tab, text="Plot G-Levels", command=self.plot_glevels).pack(side=tk.LEFT, padx=10, pady=10)
        ttk.Button(self.glevel_tab, text="Plot PSD", command=self.plot_psd_from_selection).pack(side=tk.RIGHT, padx=10, pady=10)

    def create_psd_tab(self):
        # Placeholder for PSD plot
        self.psd_fig, self.psd_ax = plt.subplots()
        self.psd_canvas = FigureCanvasTkAgg(self.psd_fig, master=self.psd_tab)
        self.psd_canvas.get_tk_widget().pack(fill='both', expand=True)

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
        if file_path:
            if file_path.endswith('.csv'):
                self.data = pd.read_csv(file_path)
            elif file_path.endswith('.xlsx'):
                self.data = pd.read_excel(file_path)

    def load_velocity_profile(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
        if file_path:
            if file_path.endswith('.csv'):
                self.velocity_data = pd.read_csv(file_path)
            elif file_path.endswith('.xlsx'):
                self.velocity_data = pd.read_excel(file_path)

    def plot_glevels(self):
        if self.data is not None:
            try:
                self.sensitivity = float(self.sensitivity_entry.get())
                self.sampling_freq = float(self.sampling_freq_entry.get())
            except ValueError:
                messagebox.showerror("Input Error", "Please enter valid numbers for sensitivity and sampling frequency.")
                return

            time = self.data.iloc[:, 0]
            x = self.data.iloc[:, 1] * self.sensitivity
            y = self.data.iloc[:, 2] * self.sensitivity
            z = self.data.iloc[:, 3] * self.sensitivity

            self.glevel_ax.clear()
            self.glevel_ax.plot(time, x, label='X')
            self.glevel_ax.plot(time, y, label='Y')
            self.glevel_ax.plot(time, z, label='Z')

            if self.velocity_data is not None:
                velocity_time = self.velocity_data.iloc[:, 0]
                velocity = self.velocity_data.iloc[:, 1]
                self.glevel_ax.plot(velocity_time, velocity, label='Velocity', linestyle='--')

            self.glevel_ax.set_xlabel("Time")
            self.glevel_ax.set_ylabel("G-Levels")
            self.glevel_ax.legend()
            self.glevel_canvas.draw()

            self.notebook.select(self.glevel_tab)
        else:
            messagebox.showerror("Data Error", "Please load the data file first.")

    def plot_psd_from_selection(self):
        if self.data is not None:
            if self.selected_range is not None:
                start, end = self.selected_range
                data_subset = self.data[(self.data.iloc[:, 0] >= start) & (self.data.iloc[:, 0] <= end)]
            else:
                data_subset = self.data

            time = data_subset.iloc[:, 0]
            x = data_subset.iloc[:, 1] * self.sensitivity
            y = data_subset.iloc[:, 2] * self.sensitivity
            z = data_subset.iloc[:, 3] * self.sensitivity

            f_x, Pxx_x = welch(x, fs=self.sampling_freq)
            f_y, Pxx_y = welch(y, fs=self.sampling_freq)
            f_z, Pxx_z = welch(z, fs=self.sampling_freq)

            self.psd_ax.clear()
            self.psd_ax.semilogy(f_x, Pxx_x, label='X')
            self.psd_ax.semilogy(f_y, Pxx_y, label='Y')
            self.psd_ax.semilogy(f_z, Pxx_z, label='Z')

            self.psd_ax.set_xlabel("Frequency [Hz]")
            self.psd_ax.set_ylabel("PSD [V**2/Hz]")
            self.psd_ax.legend()
            self.psd_canvas.draw()

            self.notebook.select(self.psd_tab)
        else:
            messagebox.showerror("Data Error", "Please load the data file first.")

if __name__ == "__main__":
    app = GLevelPSDApp()
    app.mainloop()
