import openpyxl
from datetime import datetime

def process_excel_data(file_path):
    """
    Processes an Excel file to create a dictionary of scheduled maintenance.
    
    The dictionary is structured as follows:
    {
        'date': {
            'time': 'asset_number'
        }
    }

    It only includes records where the scheduled completion time is between 5 AM and 10 AM.
    
    Args:
        file_path (str): The path to the Excel file.

    Returns:
        dict: The structured dictionary of asset maintenance schedules.
    """
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active
    except FileNotFoundError:
        print(f"Error: The file '{file_path}' was not found.")
        return {}

    schedules_dict = {}

    # Iterate over the rows, starting from the second row to skip the header
    for row in sheet.iter_rows(min_row=2):
        try:
            # Get the values from the relevant columns
            asset_number = row[0].value
            scheduled_completion_date_str = str(row[3].value)

            # Skip if either value is None
            if not asset_number or not scheduled_completion_date_str:
                continue

            # Parse the date and time from the string
            # The date format in the spreadsheet is 'DD/MM/YYYY HH:MM'
            # print(type(scheduled_completion_date_str))    
            format='%d/%m/%Y %H:%M'
            while True:
                try:
                    scheduled_datetime = datetime.strptime(scheduled_completion_date_str, format)
                    
                    # print(f"Parsed date: {scheduled_datetime.date().isoformat()}")
                    break
                except ValueError:
                    #TODO: format error handling
                    #print(f"invalid date format: {scheduled_completion_date_str}")
                    # print(f"Expected formate was {format}")
                    format = '%Y-%d-%m %H:%M:%S'  # Try a different format
                    # print(f"Expected formate was Trying new format: {format}")
                    test=datetime.strptime(scheduled_completion_date_str, format)
                    # print(f"Test successful: {test.date().isoformat()}")

            # Extract the date and time components
            schedule_date = scheduled_datetime.date().isoformat()
            schedule_time = scheduled_datetime.time()
            if schedule_date not in schedules_dict:
                    schedules_dict[schedule_date] = {"REPORTED AFTER 10 AM ON THE SAME DAY":0}

            # Check if the time is between 5 AM and 10 AM (inclusive of 5:00 and up to 9:59)
            if 5 <= schedule_time.hour and schedule_time.hour <15:
        
                if asset_number not in schedules_dict[schedule_date]:
                    schedules_dict[schedule_date][asset_number] = {}
                if row[31].value and len(row[31].value) > 10:
                    
                    # print(row[31].value)
                    
                    t = 1
                else:
                    t = 0
                # Use the time as the key and asset number as the value
                schedules_dict[schedule_date][asset_number] = {"time": schedule_date+" "+schedule_time.strftime('%H:%M'), "out_time": t}
            else:
        
                schedules_dict[schedule_date]["REPORTED AFTER 10 AM ON THE SAME DAY"]+=1

        except (ValueError, IndexError) as e:
            # Handle cases where the date format is incorrect or a row is missing data
            print(f"Skipping row due to an error: {e}")
            # continue
    
    return schedules_dict

# Specify the path to your Excel file
file_path = 'input.xlsx' 

# Call the function and print the result
maintenance_schedule = process_excel_data(file_path)

# # Print the final dictionary for verification
# if maintenance_schedule:
#     import json
#     # Use json.dumps for pretty printing the dictionary
#     # print(json.dumps(maintenance_schedule, indent=4))
# else:
#     print("No matching records found or an error occurred.")