
# API Test Automation with Data-Driven Framework

## Overview

This Python script automates API tests using a data-driven approach. The tests are executed based on the data provided in an Excel file. The script sends HTTP requests, checks the response status, and writes the results back to the same Excel file. The script supports reading various API parameters such as method, URL, headers, payload, and expected status code.

The script performs the following actions:
1. Reads test data from an Excel file.
2. Sends API requests based on the test cases.
3. Captures the response status, response body, and response time.
4. Writes the results back into the Excel file for further analysis.
5. Handles time zones in IST format for accurate timestamping.

## Requirements

Before you run the script, ensure the following Python packages are installed:

- `requests` - For sending HTTP requests.
- `openpyxl` - For handling Excel files.
- `pytz` - For timezone handling.
- `json` - For handling JSON data.

You can install these dependencies using pip:

```bash
pip install requests openpyxl pytz
```

## Setup

1. Download the script and place it in your desired project directory.
2. Prepare an Excel file with the following structure:

| Date & Time (IST) | Test Case Name | HTTP Method | URL | Header | Payload | Expected Status Code | Actual Status Code | Response Time | Response Body | Test Result |
|-------------------|----------------|-------------|-----|--------|---------|----------------------|---------------------|----------------|----------------|-------------|
| [Date/Time]       | Test Case 1     | GET         | ... | {...}  | {...}   | 200                  | 200                 | 0.12           | {...}           | ✅ Pass     |

3. Set the `file_path` in the script to point to your Excel file, e.g., `file_path = "tests/APITestAutomation_DataDrivenFramework.xlsx"`.

## How to Run the Script

1. **Ensure that your Excel file is correctly populated** with the necessary data (HTTP method, URL, headers, payload, etc.).
2. **Run the script**:

```bash
python api_test_automation.py
```

3. The script will read each test case from the Excel file, execute the API requests, and write the results back to the Excel sheet.

4. The output will contain the following columns in the Excel file:
   - **Date & Time (IST)** - The timestamp of the test execution in IST.
   - **Actual Status Code** - The HTTP status code returned by the API.
   - **Response Time** - The time taken for the request to complete.
   - **Response Body** - The JSON response body.
   - **Test Result** - Whether the test passed (✅) or failed (❌).

## Example Output

After the tests are completed, your Excel file will be updated with the results like below:

| Date & Time (IST) | Test Case Name | HTTP Method | URL | Header | Payload | Expected Status Code | Actual Status Code | Response Time | Response Body | Test Result |
|-------------------|----------------|-------------|-----|--------|---------|----------------------|---------------------|----------------|----------------|-------------|
| 2025-02-02 12:00:00 | Test Case 1     | GET         | ... | {...}  | {...}   | 200                  | 200                 | 0.12           | {...}           | ✅ Pass     |

## Notes

- Ensure that your Excel file follows the correct structure and is accessible by the script.
- The script assumes that the response body is in JSON format. Adjust it if working with other formats.
- If an HTTP request fails, the status code will be logged, and the test result will be marked as failed.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details.
