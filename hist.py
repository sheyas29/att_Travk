import pandas as pd

# Load data from Excel
file_path = 'c:/Users/kshre/OneDrive/Desktop/attendance_analysis/sample_attendance_10_employees.xlsx'  # Change the path if necessary
df = pd.read_excel(file_path)

# Calculate percentage of present and leave days
working_days = 21
df['present_percentage'] = (df['present'] / working_days) * 100
df['leave_percentage'] = (df['leave'] / working_days) * 100

# Save to Excel with charts using xlsxwriter
output_path = 'c:/Users/kshre/OneDrive/Desktop/attendance_analysis/attendance_with_histograms.xlsx'
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

print(f"Excel file with histograms saved at {output_path}")
