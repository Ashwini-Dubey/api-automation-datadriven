import datetime
import json # Handle JSON Data
import requests # Send API Requests
import openpyxl # Handle Excel Workbook
import pytz # Handle IST timezone
from datetime import datetime


#Load the Excel workbooks
file_path = "tests/APITestAutomation_DataDrivenFramework.xlsx"
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

# Get IST timezone
ist_timezone = pytz.timezone('Asia/Kolkata')

# Iterate over the Excel Workbook for the test cases, skipping the header row.
for row in range(2, sheet.max_row + 1):

    # Test Case Name
    test_case = sheet.cell(row, 2).value
    print(f"Executing the : {test_case}")

    # HTTP Method
    method = sheet.cell(row, 3).value
    if not method:
        print(f"Skipping row {row} because HTTP Method is missing")
        break

    # URL
    url = sheet.cell(row, 4).value

    # Header
    headers = json.loads(sheet.cell(row, 5).value or "{}")

    # Payload
    payload = json.loads(sheet.cell(row, 6).value or "{}")

    # Expected Status Code
    expected_status_value = sheet.cell(row, 7).value

    # Capture Start time
    start_time = datetime.now()

    # Execute API Request
    response = requests.request(method, url, headers=headers, json=payload)
    response_data = json.dumps(response.json())

    # Capture End time
    end_time = datetime.now()

    # Get the response time in seconds
    response_time = (end_time - start_time).total_seconds()

    # Get Current date and time in IST
    current_datetime_ist = datetime.now(ist_timezone).strftime('%Y-%m-%d %H:%M:%S')

    # Write results back to Excel

    # Date & Time (IST)
    sheet.cell(row, 1, current_datetime_ist)
    # Actual Status Code
    sheet.cell(row, 8, response.status_code)
    # Response Time
    sheet.cell(row, 9, response_time)
    # Response Body
    sheet.cell(row, 10, response_data)

    # Check for missing expected status value and compare with actual status
    if expected_status_value is not None and response.status_code == expected_status_value:
        test_result = "✅ Pass"
    else:
        test_result = "❌ Fail"

    # Write the test result to the Excel file
    sheet.cell(row, 11, test_result)

# Save the update Excel file
workbook.save(file_path)
print("✅ API Tests Completed & Results Saved in Excel")



