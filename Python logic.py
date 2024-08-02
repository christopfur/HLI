import pandas as pd
from datetime import datetime
import os


#C:\Users\cmccullough\Desktop\Payroll Automation 1.0

# Define the file paths and sheet names - these should be updated for each file and computer where the code is executed
input_file_path = r'C:\Users\cmccullough\Desktop\Payroll Automation 1.0\Payrolls June & July 2024\HLI Payroll 6-7-24.xlsx' #First row(s) should be a header row
input_file_path_emp_mapping = r'C:\Users\cmccullough\Desktop\Payroll Automation 1.0\Payrolls\HLI 1-15-24 Payroll.xls' #First row(s) should be a header row
sheet_name_employee_mapping = 'Employee Mapping'
gl_mapping_logic_file = r'C:\Users\cmccullough\Desktop\Payroll Automation 1.0\GL Mapping Logic.xlsx' #First row(s) should be a header row
gl_mapping_logic_sheet = 'GL Category Mapping'

# Read the sheet names from the input file
xls = pd.ExcelFile(input_file_path)
sheet_names = xls.sheet_names

# Find the sheet name that contains 'PAYROLL.ALLOCATION'
sheet_name_payroll_allocation = next((name for name in sheet_names if 'PAYROLL.ALLOCATION' in name), None)

if not sheet_name_payroll_allocation:
    raise ValueError("No sheet name contains 'PAYROLL.ALLOCATION'")

# Read the input_file_path to check the first cell
df_check = pd.read_excel(input_file_path, sheet_name=sheet_name_payroll_allocation, header=None) #Disable header so .iloc can be used
# Check if the first cell contains 'Row' and if so remove/delete first row as it's a useless header row/header
if df_check.iloc[0, 0] == 'Row' or df_check.iloc[0, 1] == 'Field 0': #Added 'Field 0' since I've encountered (rare) scenairos where 'Row' isn't included in DF
    # Read the file excluding the first row
    df_payroll_allocation = pd.read_excel(input_file_path, sheet_name=sheet_name_payroll_allocation, skiprows=1)
else:
    # Read the file as it is
    df_payroll_allocation = pd.read_excel(input_file_path, sheet_name=sheet_name_payroll_allocation)

# Read other Excel files
#df_payroll_allocation = pd.read_excel(input_file_path, sheet_name=sheet_name_payroll_allocation) (logic now above due to header check)
df_employee_mapping = pd.read_excel(input_file_path_emp_mapping, sheet_name=sheet_name_employee_mapping)
df_gl_mapping_logic = pd.read_excel(gl_mapping_logic_file, sheet_name=gl_mapping_logic_sheet)


# List of columns to keep for payroll allocation
columns_to_keep_payroll_allocation = [
    'EEID', 'Company Number', 'Company Name', 'Hire Date', 'Employee ID', 'Employee Name', 
    'Payroll Number', 'Home Location', 'Home Location Desc', 'Current Job Code', 'Current Job Desc', 
    'Hourly Pay Rate', 'Total Hours Paid', 'Total Gross Wages', 'AUTO ALLOWANCE', 'Advance - Deduction ', 'CELLPHONE (NON TAXAB', 
    'GROUP TERM LIFE', 'HALF PAY', 'HOLIDAY', 'HOUSING ALLOWANCE', 'Leave with Pay', 'LTD IMPUTED INCOME', 
    'PERSONAL TIME OFF', 'REGULAR PAY', 'STD Imputed Income', 'MEDICARE - EMPLOYER', 'OASDI - EMPLOYER', 
    'FEDERAL UNEMPLOYMENT', 'CA SUTA', 'NY SUTA', 'NY DISABILITY (ER)', 'NY RE-EMPLOYMENT SER', 'ADMIN FEE', 'ADMINISTRATIVE FEE',
    'Guardian LTD Tax Cho', 'Guardian LTD Tax Cho.1','Guardian 2xSalary to', 'Southern CA Anthem H', 'Anthem PPO 250', 'Anthem PPO 500', 
    'Anthem PPO 750 - B', 'Anthem HSA 3000', 'Northern CA Kaiser H', 'Southern CA Kaiser H', 'Guardian Dental PPO ',  'Guardian Dental PPO .1',
    'Guardian STD Tax Cho.1', 'Guardian STD Tax Cho', 'VSP Vision Standard', 'HEALTH SAVINGS ACCOU', 'NON CASH', 'WIRE FEE OFFSET','BEREAVEMENT',
    'FRINGE BENEFIT','JURY DUTY','REFERRAL BONUS','HALF PAY DOUBLE TIME','Off Cycle Fee','RETRO PAY','COMMISSION SUPPLEMEN', 'TRANSIT COMMUTER REI',
    'PTO Payout', 'SEVERANCE', 'CA EMPLOYMENT TRAINI','MD SUTA', 'S1 401k MEP Annual F', 'Anthem HSA 3200'
]

# Filter the columns to keep only those that are present in the DataFrame
columns_present = [col for col in columns_to_keep_payroll_allocation if col in df_payroll_allocation.columns]

# Keep only the specified columns in payroll allocation DataFrame
df_payroll_allocation = df_payroll_allocation[columns_present]

# Ensure 'Advance - Deduction' column values are negative
if 'Advance - Deduction ' in df_payroll_allocation.columns:
    df_payroll_allocation['Advance - Deduction '] = df_payroll_allocation['Advance - Deduction '].apply(lambda x: -abs(x) if pd.notnull(x) else x)

# Convert 'Hire Date' to datetime and format it to only include the date part
df_payroll_allocation['Hire Date'] = pd.to_datetime(df_payroll_allocation['Hire Date']).dt.date

# Rename 'Company Name' in employee mapping DataFrame to 'Company Name - Emp. Mapping'
df_employee_mapping = df_employee_mapping.rename(columns={'Company Name': 'Company Name - Emp. Mapping'})

# List of columns to keep from employee mapping for the join
columns_to_keep_employee_mapping = [
    'Employee ID', 'Company Name - Emp. Mapping', 'Home Department Desc', 'GL Account Description', 'Payroll Salaries', 
    'Payroll Taxes', 'Insurance-Medical', 'Insurance-Disability', 'Insurance-Life', 
    'Insurance-Vision', 'Insurance-Dental', 'Payroll Processing', 'Bank Charges', 'Telephone', 'Accrued PTO'
]

# Filter the employee mapping DataFrame to keep only the necessary columns for the join
df_employee_mapping = df_employee_mapping[columns_to_keep_employee_mapping]

# Filter for 'Human Longevity, Inc.' in the relevant column (confirmed that other companies aren't relevant to analysis)
df_employee_mapping = df_employee_mapping[df_employee_mapping['Company Name - Emp. Mapping'] == 'Human Longevity, Inc.']

# Perform the LEFT join on 'Employee ID'
df_merged = pd.merge(df_payroll_allocation, df_employee_mapping, on='Employee ID', how='left')

for column in df_merged.columns:
    print("'" + str(column) + "'")

# Update the 'Found / Not Found' column in df_gl_mapping_logic
df_gl_mapping_logic['Found / Not Found'] = df_gl_mapping_logic['Column Name'].apply(
    lambda x: 'FOUND' if x in df_merged.columns else 'NOT FOUND'
)

# Initialize a list to store the new dataframe rows
gl_costs_data = []

# Iterate over each row in the payroll allocation DataFrame
for idx, payroll_row in df_merged.iterrows():
    Emp_ID = payroll_row['Employee ID']
    Emp_Name = payroll_row['Employee Name']
    # Iterate over each column in GL Mapping Logic
    for _, gl_row in df_gl_mapping_logic.iterrows():
        column_name = gl_row['Column Name']
        gl_category = gl_row['GL Category']
             
        if (column_name in df_merged.columns) and pd.notnull(payroll_row[column_name]):
            # Retrieve the value from the current row and matching column
            dollar_amount = payroll_row[column_name]
            # Find the corresponding GL Category column in df_merged
            if gl_category in df_merged.columns:
                GL_Code = payroll_row[gl_category]
                # If GL Category is 'Payroll Salaries' update to a sub-category (logic added 5/30/2024)
                if gl_category == 'Payroll Salaries':
                    if GL_Code == '5015-01' or GL_Code == '5015-02' or GL_Code == '5015-01/5015-02':
                        gl_category = 'Salaries-Imaging Technicians'
                    if GL_Code == '5020-02':
                        gl_category = 'Salaries-Medical Assistants'
                    if GL_Code == '6460-00':
                        gl_category = 'Salaries-Corporate'
                    if GL_Code == '6470-00':
                        gl_category = 'Salaries-Executives'
                    if GL_Code == '6490-00':
                        gl_category = 'Salaries-Sales'
                    if GL_Code == '6485-00' or GL_Code == '6485-01' or GL_Code == '6485-02' or GL_Code == '6485-01/6485-02':
                        gl_category = 'Salaries-Other'
                    #print(gl_category, "-", GL_Code)
            else:
                GL_Code = None
            
            if str(GL_Code) and '/' in str(GL_Code):
                # Split the GL Code and divide the dollar amount
                codes = str(GL_Code).split('/')
                amount = dollar_amount / len(codes)
                for code in codes:
                    gl_costs_data.append({
                        'Employee ID': Emp_ID,
                        'Employee Name' : Emp_Name,
                        'GL Category': gl_category,
                        'GL Code': code,
                        'Dollar Amount': amount
                    })
            else:
                #Exceptions for edge cases that don't fit into standardized buckets
                #Exception #1 : added 5/28/2024 for employee NATALIE ALVAREZ
                if Emp_ID == 'L96108' and column_name == 'Advance - Deduction ': 
                    gl_category = 'Accounts Receivable'
                    GL_Code = '1100-00'
                #Exception #2 : added 5/30/2024 for employee SCOTT DEAN
                if Emp_ID == 'X90046' and column_name == 'Advance - Deduction ': 
                    gl_category = 'Accounts Receivable'
                    GL_Code = '1100-00'

                # Append the data to the list
                gl_costs_data.append({
                    'Employee ID': Emp_ID,
                    'Employee Name' : Emp_Name,
                    'GL Category': gl_category,
                    'GL Code': GL_Code,
                    'Dollar Amount': dollar_amount
                })

# Create a new DataFrame from the list
df_gl_costs = pd.DataFrame(gl_costs_data)

# Create new summarized GL Costs dataframe
df_gl_costs_2 = pd.DataFrame(columns=['GL Category Code', 'GL Category', 'GL Code', 'Dollar Amount'])

# Concatenate 'GL Category' and 'GL Code' columns in 'GL Costs 1' and sum 'Dollar Amount' for each group
df_gl_costs_2['GL Category Code'] = df_gl_costs['GL Category'] + ' ' + df_gl_costs['GL Code']
df_gl_costs_2['GL Category'] = df_gl_costs['GL Category']
df_gl_costs_2['GL Code'] = df_gl_costs['GL Code']
df_gl_costs_2['Dollar Amount'] = df_gl_costs['Dollar Amount']
df_gl_costs_2 = df_gl_costs_2.groupby(['GL Category Code', 'GL Category', 'GL Code'])['Dollar Amount'].sum().reset_index()
df_gl_costs_2 = df_gl_costs_2.rename(columns={'Dollar Amount': 'Total Dollar Amount'})

#print(df_gl_costs_2.dtypes)

# Get the base name of the input file without extension
input_file_name = os.path.basename(input_file_path).split('.')[0]
Payroll_Number = df_merged['Payroll Number'][0]

# Save the merged DataFrame and df_gl_costs to an Excel file with the first row frozen and set column widths
current_date = datetime.now().strftime('%Y-%m-%d')
output_file_path = rf'C:\Users\cmccullough\Desktop\Payroll Automation 1.0\Payrolls June & July 2024\output\OUTPUT_{input_file_name}_{Payroll_Number}_{current_date}.xlsx'

with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    df_merged.to_excel(writer, index=False, sheet_name='Payroll Allocation Details')
    workbook = writer.book
    worksheet_payroll = writer.sheets['Payroll Allocation Details']

    # Add additional tab from GL Mapping Logic file
    df_gl_mapping_logic.to_excel(writer, index=False, sheet_name='GL Mapping Logic')
    worksheet_gl_mapping = writer.sheets['GL Mapping Logic']

    # Add additional tab for GL Costs 1
    df_gl_costs.to_excel(writer, index=False, sheet_name='GL Costs 1')
    worksheet_gl_costs_1 = writer.sheets['GL Costs 1']

    # Add additional tab for GL Costs 2
    df_gl_costs_2.to_excel(writer, index=False, sheet_name='GL Costs 2')
    worksheet_gl_costs_2 = writer.sheets['GL Costs 2']

    # Freeze the first row
    worksheet_payroll.freeze_panes(1, 0)
    worksheet_gl_mapping.freeze_panes(1, 0)
    worksheet_gl_costs_1.freeze_panes(1, 0)
    worksheet_gl_costs_2.freeze_panes(1, 0)

    # Add currency formatting for dollarized columns
    currency_format = workbook.add_format({'num_format': '$#,##0.00'})
    worksheet_gl_costs_2.set_column('D:D', None, currency_format)
    worksheet_gl_costs_1.set_column('D:D', None, currency_format)

    # Set the width of each column to the length of the longest string in the column
    for column in df_merged.columns:
        max_len = max(df_merged[column].astype(str).map(len).max(), len(column)) + 2  # Adding some padding
        col_idx = df_merged.columns.get_loc(column)
        worksheet_payroll.set_column(col_idx, col_idx, max_len)

    for column in df_gl_mapping_logic.columns:
        max_len = max(df_gl_mapping_logic[column].astype(str).map(len).max(), len(column)) + 2  # Adding some padding
        col_idx = df_gl_mapping_logic.columns.get_loc(column)
        worksheet_gl_mapping.set_column(col_idx, col_idx, max_len)

    for column in df_gl_costs.columns:
        max_len = max(df_gl_costs[column].astype(str).map(len).max(), len(column)) + 2  # Adding some padding
        col_idx = df_gl_costs.columns.get_loc(column)
        worksheet_gl_costs_1.set_column(col_idx, col_idx, max_len)

    for column in df_gl_costs_2.columns:
        max_len = max(df_gl_costs_2[column].astype(str).map(len).max(), len(column)) + 2  # Adding some padding
        col_idx = df_gl_costs_2.columns.get_loc(column)
        worksheet_gl_costs_2.set_column(col_idx, col_idx, max_len)

    # Apply alternating colors for the rows across the entire row range
    row_count, col_count = df_merged.shape
    worksheet_payroll.conditional_format(1, 0, row_count, col_count - 1,
                                         {'type': 'formula', 'criteria': 'MOD(ROW(), 2) = 0',
                                          'format': workbook.add_format({'bg_color': '#D3D3D3'})})
    worksheet_payroll.conditional_format(1, 0, row_count, col_count - 1,
                                         {'type': 'formula', 'criteria': 'MOD(ROW(), 2) = 1',
                                          'format': workbook.add_format({'bg_color': '#FFFFFF'})})

    # Conditional formatting for 'Found / Not Found' column in GL Mapping Logic sheet
    not_found_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    found_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
    
    found_col_index = df_gl_mapping_logic.columns.get_loc('Found / Not Found')
    worksheet_gl_mapping.conditional_format(1, found_col_index, len(df_gl_mapping_logic), found_col_index,
                                            {'type': 'cell', 'criteria': '==', 'value': '"NOT FOUND"', 'format': not_found_format})
    worksheet_gl_mapping.conditional_format(1, found_col_index, len(df_gl_mapping_logic), found_col_index,
                                            {'type': 'cell', 'criteria': '==', 'value': '"FOUND"', 'format': found_format})

print(f"DataFrame has been saved to {output_file_path}")