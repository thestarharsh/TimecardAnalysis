#!/usr/bin/env python
# coding: utf-8

# In[1]:


file_path = r"C:\Users\Harsh talati\Dropbox\PC\Downloads\Assignment_Timecard.xlsx"


# In[2]:


import pandas as pd
from datetime import datetime, timedelta


# In[8]:


import pandas as pd
from datetime import datetime, timedelta

def analyze_timecard(file_path):
    # Convert 'Time' and 'Time Out' columns to datetime objects
    df = pd.read_excel(file_path)

    df['Time'] = pd.to_datetime(df['Time'], errors='coerce')
    df['Time Out'] = pd.to_datetime(df['Time Out'], errors='coerce')

    def calculate_time_difference(start_time, end_time):
        diff = end_time - start_time
        return diff.total_seconds() / 3600

    results_a = [] # 7 consecutive days
    results_b = [] # less than 10 hours between shifts
    results_c = [] # More than 14 hours in a single shift

    for _, row in df.iterrows():
        if pd.isna(row['Time']) or pd.isna(row['Time Out']) or pd.isna(row['Employee Name']):
            continue

        consecutive_days = 1
        total_hours = calculate_time_difference(row['Time'], row['Time Out'])
        previous_end_time = row['Time Out']

        for _, next_row in df[df['Employee Name'] == row['Employee Name']].iterrows():
            if pd.isna(next_row['Time']) or pd.isna(next_row['Time Out']):
                continue

            if (next_row['Time'] - previous_end_time).days == 1:
                consecutive_days += 1
            else:
                consecutive_days = 1

            time_between_shifts = calculate_time_difference(previous_end_time, next_row['Time'])
            if 1 < time_between_shifts < 10:
                results_b.append(f"Employee: {row['Employee Name']}, Position: {row['Position ID']} has less than 10 hours between shifts.")

            if total_hours > 14:
                results_c.append(f"Employee: {row['Employee Name']}, Position: {row['Position ID']} has worked for more than 14 hours in a single shift.")

            previous_end_time = next_row['Time Out']
            total_hours += calculate_time_difference(next_row['Time'], next_row['Time Out'])

            if consecutive_days == 7:
                results_a.append(f"Employee: {row['Employee Name']}, Position: {row['Position ID']} has worked for 7 consecutive days.")
                break

    # Print results in the specified order
    for result in results_a:
        print(result)
    for result in results_b:
        print(result)
    for result in results_c:
        print(result)

    # Write results to the output.txt file
    with open('output.txt', 'w') as output_file:
        for result in results_a + results_b + results_c:
            output_file.write(result + '\n')

if __name__ == "__main__":
    file_path = r"C:\Users\Harsh talati\Dropbox\PC\Downloads\Assignment_Timecard.xlsx"
    analyze_timecard(file_path)


# In[ ]:




