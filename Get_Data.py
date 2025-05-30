import gspread
import numpy as np
from oauth2client.service_account import ServiceAccountCredentials

# Set up Google Sheets API authentication
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

# Open the Google Sheet
spreadsheet = client.open('TestforJDFrags')  # Replace with your Google Sheet name
worksheet = spreadsheet.worksheet('Applicants')  # Main sheet, where D6:BH6, BW6, BX6, CX6, DW6, BY6:CW6 are
job_frags_sheet = spreadsheet.worksheet('Job Frags')  # 'Job Frags' sheet for weights

# Function to get the column index by header name
def get_column_by_header(sheet, header_name):
    """
    Find the column index based on the header name.
    """
    headers = sheet.row_values(1)  # Get the first row as headers
    try:
        return headers.index(header_name) + 1  # Adding 1 to get the correct index for gspread
    except ValueError:
        return None  # Return None if the header doesn't exist

# Function to convert value to float, handling empty or invalid values
def safe_float(val):
    """
    Convert a value to float, replacing invalid or empty values with 0.
    """
    if not val or val == '':
        return 0  # Default to 0 if the value is empty or invalid
    try:
        return float(val)
    except ValueError:
        return 0  # Default to 0 if conversion fails (non-numeric value)

# Function to get data from a specific column in the "Job Frags" sheet based on header name
def get_column_data(sheet, header_name, row_start, row_end):
    """
    Retrieve column data based on the header name and row range.
    """
    column_index = get_column_by_header(sheet, header_name)
    if column_index is None:
        print(f"Header '{header_name}' not found.")
        return []
    
    return [safe_float(cell) for cell in sheet.col_values(column_index)[row_start-1:row_end]]  # Adjust for 1-based index

def get_column1_data(sheet, start_header, end_header, start_row=6):
    # Get the headers from row 2 (assumed header row)
    headers = sheet.row_values(1)

    # Find the indexes of the start and end headers
    start_index = headers.index(start_header)
    end_index = headers.index(end_header)

    # Now we need to get the values in the column range from the start_header to the end_header, starting from row 6

    
    # Get the data for row 6 only (adjust by -1 for zero-based index)
    row_data = sheet.row_values(start_row)[start_index:end_index+1]  # Row 6 data from start_index to end_index

    # Return the data as a flat list
    return [safe_float(val) for val in row_data]


# Function to retrieve all the necessary data for calculations
def get_data_for_calculation():
    # Retrieve data for specific ranges from the main sheet and 'Job Frags'
    
    # Get values for columns like 'Conflict Averse', 'Overall Performance', etc.
    conflict_averse_data = get_column1_data(job_frags_sheet, 'Conflict Averse', 'Initiating')
    overall_performance_data = get_column1_data(job_frags_sheet, 'Overall Performance', 'Fundraising')
    flourish_data = get_column1_data(job_frags_sheet, 'Flourish', 'Humility')
    cognitive_data = safe_float(job_frags_sheet.cell(3, get_column_by_header(job_frags_sheet, 'Cognitive')).value)
    location_data = safe_float(job_frags_sheet.cell(3, get_column_by_header(job_frags_sheet, 'Location')).value)
    lji_knowledge_data = safe_float(job_frags_sheet.cell(3, get_column_by_header(job_frags_sheet, 'LJI Knowledge')).value)
    # location_data = get_column_data(job_frags_sheet, 'Location', 2, 6)
    # lji_knowledge_data = get_column_data(job_frags_sheet, 'LJI Knowledge', 2, 6)
    
    # Pulling values for the main sheet ranges: D6:AG6, AH6:AR6, AU6:BH6, BW6
    d_ag_data = worksheet.row_values(6)[3:33]  # D6:AG6
    ah_ar_data = worksheet.row_values(6)[33:44]  # AH6:AR6
    au_bh_data = worksheet.row_values(6)[46:60]  # AU6:BH6
    bw_data = worksheet.cell(6, 75).value  # BW6
    
    # Ensure empty values are handled in the row ranges
    d_ag_data = [safe_float(val) for val in d_ag_data]
    ah_ar_data = [safe_float(val) for val in ah_ar_data]
    au_bh_data = [safe_float(val) for val in au_bh_data]
    bw_data = safe_float(bw_data)
    
    # Return all retrieved data in a dictionary
    return {
        "conflict_averse": conflict_averse_data,
        "overall_performance": overall_performance_data,
        "flourish": flourish_data,
        "cognitive": cognitive_data,
        "location": location_data,
        "lji_knowledge": lji_knowledge_data,
        "d_ag": d_ag_data,
        "ah_ar": ah_ar_data,
        "au_bh": au_bh_data,
        "bw": bw_data
    }

data = get_data_for_calculation()  # Get all the data
print(data["cognitive"])  # Print the "conflict_averse" data
# print(np.dot(data["ah_ar"],data["overall_performance"]))
print(data["overall_performance"])