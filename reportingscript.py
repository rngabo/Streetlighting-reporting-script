import tkinter as tk
from tkinter import filedialog
import pandas as pd

def browse_online_rate_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        online_rate_entry.delete(0, tk.END)
        online_rate_entry.insert(0, file_path)

def browse_light_rate_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        light_rate_entry.delete(0, tk.END)
        light_rate_entry.insert(0, file_path)

def calculate_average_rates():
    online_file_path = online_rate_entry.get()
    light_file_path = light_rate_entry.get()
    
    if online_file_path and light_file_path:
        try:
            # Read the online rate file from Excel into a DataFrame
            online_df = pd.read_excel(online_file_path)
            
            # Read the light rate file from Excel into a DataFrame
            light_df = pd.read_excel(light_file_path)
            
            # Check if the online rate file contains the expected columns
            online_columns_expected = ['Timestamp', 'Online rate']  # Modify with the actual expected columns
            if not set(online_columns_expected).issubset(online_df.columns):
                result_label.config(text="Check if you have uploaded files in the right places.")
                return
            
            # Check if the light rate file contains the expected columns
            light_columns_expected = ['Timestamp', 'Light rate']  # Modify with the actual expected columns
            if not set(light_columns_expected).issubset(light_df.columns):
                result_label.config(text="Check if you have uploaded files in the right places.")
                return
            
            # Calculate the average online rate based on the specified column
            average_online_rate = online_df['Online rate'].mean()
            
            # Calculate the offline rate by considering all timestamps as 100 - online rate
            offline_rate = 100 - average_online_rate
            
            # Calculate the aggregated light percentage based on timestamp time for nighttime (6 PM to 5:59 AM)
            light_df['Hour'] = light_df['Timestamp'].dt.hour
            light_df_night = light_df[(light_df['Hour'] >= 18) | (light_df['Hour'] < 6)]
            average_light_rate_night = light_df_night['Light rate'].mean()
            
            # Calculate the aggregated light percentage based on timestamp time for daytime (6 AM to 5:59 PM)
            light_df_day = light_df[(light_df['Hour'] >= 6) & (light_df['Hour'] < 18)]
            average_light_rate_day = light_df_day['Light rate'].mean()
            
            # Calculate Off during the night by taking 100 - light rate during evening time
            off_during_night = 100 - light_df_night['Light rate'].mean()
            
            # Display all results
            result_label.config(text=f"Online Rate: {average_online_rate:.2f}%\nOffline Rate: {offline_rate:.2f}%\nLight Rate(Night): {average_light_rate_night:.2f}%\nLight Rate(Day): {average_light_rate_day:.2f}%\nOff during the night: {off_during_night:.2f}%")
        except pd.errors.ParserError as e:
            result_label.config(text="Check if you have uploaded files in the right places.")
        except Exception as e:
            result_label.config(text="Check if you have uploaded files in the right places.")
    else:
        result_label.config(text="Please select both Online Rate and Light Rate Excel files")

# Create the main application window
app = tk.Tk()
app.title("Rate Calculator")

# Set the minimum width and height (1000px minimum width)
app.minsize(1000, 400)

# Calculate the screen width and height
screen_width = app.winfo_screenwidth()
screen_height = app.winfo_screenheight()

# Calculate the x and y positions to center the window
x = (screen_width - 1000) // 2
y = (screen_height - 400) // 2

# Set the window position
app.geometry(f"1000x400+{x}+{y}")

# Create a frame for left and right sections
frame_left = tk.Frame(app)
frame_right = tk.Frame(app)
frame_left.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.BOTH, expand=True)
frame_right.pack(side=tk.RIGHT, padx=10, pady=10, fill=tk.BOTH, expand=True)

# Create labels for left and right sections
online_rate_label = tk.Label(frame_left, text="Online Rate Excel File", font=("Arial", 12))
light_rate_label = tk.Label(frame_right, text="Light Rate Excel File", font=("Arial", 12))

online_rate_label.grid(row=0, column=0, padx=5, pady=5, sticky='w')
light_rate_label.grid(row=0, column=0, padx=5, pady=5, sticky='w')

# Create entry widgets for file paths
online_rate_entry = tk.Entry(frame_left, width=30)
light_rate_entry = tk.Entry(frame_right, width=30)

online_rate_entry.grid(row=1, column=0, padx=5, pady=5)
light_rate_entry.grid(row=1, column=0, padx=5, pady=5)

# Create browse buttons for left and right sections
online_browse_button = tk.Button(frame_left, text="Browse Online Rate Excel File", command=browse_online_rate_file)
light_browse_button = tk.Button(frame_right, text="Browse Light Rate Excel File", command=browse_light_rate_file)

online_browse_button.grid(row=2, column=0, padx=5, pady=5)
light_browse_button.grid(row=2, column=0, padx=5, pady=5)

# Create a single "Calculate" button, spanning two columns, with a fixed width of 400px
calculate_button = tk.Button(app, text="Calculate Average Rates", command=calculate_average_rates, width=40)
calculate_button.pack(pady=10, padx=5, fill=tk.BOTH, expand=True)

# Create a label to display results
result_label = tk.Label(app, text="", font=("Arial", 12))
result_label.pack()

# Run the application
app.mainloop()
