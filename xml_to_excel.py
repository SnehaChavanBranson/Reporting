import xml.etree.ElementTree as ET
import pandas as pd
import re
from datetime import datetime
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows  # Import this function

# Parse the XML file
tree = ET.parse("results.xml")
root = tree.getroot()

# Define the namespaces (to handle the xmlns="http://www.froglogic.com/resources/schemas/xml3")
namespaces = {"": "http://www.froglogic.com/resources/schemas/xml3"}

# Initialize a list to store data
data = []


# Function to convert time string to datetime object
def convert_to_datetime(time_str):
    return datetime.strptime(time_str, "%Y-%m-%dT%H:%M:%S.%fZ")


# Iterate through each <test> node to extract relevant information
serial_number = 1  # Initialize serial number
for test in root.findall(".//test", namespaces):
    # Extract the test name from the <name> tag inside <prolog>
    test_name = test.find(".//prolog/name", namespaces).text.strip()

    # Check if the test case name contains "test case" followed by a number
    if not re.search(r"test case \d+", test_name, re.IGNORECASE):
        continue  # Skip test cases that don't match the pattern

    testcase = {}

    # Add serial number
    testcase["Serial Number"] = serial_number
    serial_number += 1  # Increment serial number for next valid test case

    testcase["Test Case Name"] = test_name
    testcase["Status"] = "Unknown"

    # Extract the start time and end time
    start_time_str = test.find(".//prolog", namespaces).attrib.get("time")
    end_time_str = test.find(".//epilog", namespaces).attrib.get("time")

    start_time = convert_to_datetime(start_time_str)
    end_time = convert_to_datetime(end_time_str)

    # Calculate total time duration (in seconds)
    total_time = (end_time - start_time).total_seconds()

    # Store the times
    testcase["Start Time"] = start_time_str
    testcase["End Time"] = end_time_str
    testcase["Total Time (seconds)"] = total_time

    # Extract logs (messages)
    logs = []
    for message in test.findall(".//message", namespaces):
        text = message.find("text", namespaces).text.strip()
        logs.append(text)

    # Check for the verification status (PASS or FAIL)
    verification = test.find(".//verification", namespaces)
    if verification is not None:
        result = verification.find(".//scriptedVerificationResult", namespaces)
        if result is not None:
            testcase["Status"] = result.attrib["type"]

    # If status is still Unknown, mark it as FAIL
    if testcase["Status"] == "Unknown":
        testcase["Status"] = "FAIL"

    # Combine all logs into one cell (as the message logs can span multiple entries)
    testcase["Logs"] = "\n".join(logs)

    # Leave the Comments section empty
    testcase["Comments"] = ""  # Comments column should remain empty

    # Append the testcase data to the list
    data.append(testcase)

# Convert the data into a pandas DataFrame
df = pd.DataFrame(data)

# Count the number of passes and fails
pass_count = df[df["Status"] == "PASS"].shape[0]
fail_count = df[df["Status"] == "FAIL"].shape[0]
total_tests = df.shape[0]

# Module Name (you can adjust this based on the filename or other criteria)
module_name = "Module_1"  # Replace with actual module name if needed

# Create a summary DataFrame
summary_data = {"Module Name": [module_name], "Total Test Cases": [total_tests], "Pass": [pass_count], "Fail": [fail_count]}

summary_df = pd.DataFrame(summary_data)

# Create a workbook using openpyxl
wb = Workbook()

# Create a worksheet for Test Cases
ws_test_cases = wb.active
ws_test_cases.title = "Test Cases"
for r in dataframe_to_rows(df, index=False, header=True):
    ws_test_cases.append(r)

# Create a worksheet for Summary
ws_summary = wb.create_sheet("Summary")
for r in dataframe_to_rows(summary_df, index=False, header=True):
    ws_summary.append(r)

# Add a Bar Chart for Pass/Fail
chart = BarChart()
chart.type = "col"
chart.title = "Pass/Fail Test Cases"
chart.style = 10
chart.x_axis.title = "Status"
chart.y_axis.title = "Test Cases"

# Write the data to cells
data = [["Pass", pass_count], ["Fail", fail_count]]
ws_summary.append(["Status", "Count"])  # Headers for the chart data
for row in data:
    ws_summary.append(row)

# Create a reference to the data for the chart
data_ref = Reference(ws_summary, min_col=2, min_row=2, max_col=2, max_row=3)
categories_ref = Reference(ws_summary, min_col=1, min_row=2, max_row=3)

chart.add_data(data_ref, titles_from_data=True)
chart.set_categories(categories_ref)

# Add the chart to the summary sheet
ws_summary.add_chart(chart, "E5")

# Save the workbook to an Excel file
wb.save("test_report_with_times.xlsx")

print("Excel file with a bar graph in the Summary sheet created successfully!")
