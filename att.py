import pandas as pd

# Load data from Excel
file_path = 'c:/Users/kshre/OneDrive/Desktop/attendance_analysis/sample_attendance_10_employees.xlsx'  # Change the path if necessary
df = pd.read_excel(file_path)

# Calculate percentage of present and leave days
working_days = 21
df['present_percentage'] = (df['present'] / working_days) * 100
df['leave_percentage'] = (df['leave'] / working_days) * 100

# Calculate total counts for each timeframe
time_ranges = ['0-0915', '0916-0930', '0931-1000', 'after 1000', 'before 1600', '1610-1630', '1631-1700', '1701-1715', '1716-1730', 'after 1730']
time_range_counts = {time_range: df[time_range].sum() for time_range in time_ranges}

# Save to Excel with charts using xlsxwriter
output_path = 'c:/Users/kshre/OneDrive/Desktop/attendance_analysis/attendance_with_detailed_analysis.xlsx'
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    df.to_excel(writer, index=False, sheet_name='Attendance')

    # Access the XlsxWriter workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets['Attendance']

    # Create a histogram for present percentage
    present_chart = workbook.add_chart({'type': 'column'})
    present_chart.add_series({
        'name': 'Present Percentage',
        'categories': ['Attendance', 1, 0, len(df), 0],  # Employee names
        'values': ['Attendance', 1, df.columns.get_loc('present_percentage'), len(df), df.columns.get_loc('present_percentage')],  # Present percentages
    })
    present_chart.set_title({'name': 'Present Percentage of Employees'})
    present_chart.set_x_axis({'name': 'Employee'})
    present_chart.set_y_axis({'name': 'Present Percentage'})
    worksheet.insert_chart('M2', present_chart)

    # Create a histogram for leave percentage
    leave_chart = workbook.add_chart({'type': 'column'})
    leave_chart.add_series({
        'name': 'Leave Percentage',
        'categories': ['Attendance', 1, 0, len(df), 0],  # Employee names
        'values': ['Attendance', 1, df.columns.get_loc('leave_percentage'), len(df), df.columns.get_loc('leave_percentage')],  # Leave percentages
    })
    leave_chart.set_title({'name': 'Leave Percentage of Employees'})
    leave_chart.set_x_axis({'name': 'Employee'})
    leave_chart.set_y_axis({'name': 'Leave Percentage'})
    worksheet.insert_chart('M20', leave_chart)

    # Create a bar chart for total counts in each timeframe
    time_ranges_chart = workbook.add_chart({'type': 'column'})
    time_ranges_chart.add_series({
        'name': 'Time Range Attendance',
        'categories': ['Attendance', 0, 3, 0, len(time_ranges) + 2],  # Time range names
        'values': ['Attendance', 1, 3, 1, len(time_ranges) + 2],  # Time range counts
    })
    time_ranges_chart.set_title({'name': 'Total Counts for Each Time Range'})
    time_ranges_chart.set_x_axis({'name': 'Time Range'})
    time_ranges_chart.set_y_axis({'name': 'Total Count'})
    worksheet.insert_chart('M38', time_ranges_chart)

    # Write the total counts for each timeframe to the worksheet for reference
    for row, (time_range, count) in enumerate(time_range_counts.items(), start=1):
        worksheet.write(row, len(df.columns) + 1, time_range)
        worksheet.write(row, len(df.columns) + 2, count)

print(f"Excel file with detailed analysis saved at {output_path}")
