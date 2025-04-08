import xml.etree.ElementTree as ET
import pandas as pd
import re
from datetime import datetime
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows


tree = ET.parse("results.xml")
root = tree.getroot()

namespaces = {"": "http://www.froglogic.com/resources/schemas/xml3"}

data = []


def convert_to_datetime(time_str):
    return datetime.strptime(time_str, "%Y-%m-%dT%H:%M:%S.%fZ")


serial_number = 1  # Initialize serial number
for test in root.findall(".//test", namespaces):
    test_name = test.find(".//prolog/name", namespaces).text.strip()

    if not re.search(r"test case \d+", test_name, re.IGNORECASE):
        continue  # Skip test cases that don't match the pattern

    testcase = {}

    testcase["Serial Number"] = serial_number
    serial_number += 1  # Increment

    testcase["Test Case Name"] = test_name
    testcase["Status"] = "Unknown"

    start_time_str = test.find(".//prolog", namespaces).attrib.get("time")
    end_time_str = test.find(".//epilog", namespaces).attrib.get("time")

    start_time = convert_to_datetime(start_time_str)
    end_time = convert_to_datetime(end_time_str)

    total_time = (end_time - start_time).total_seconds()

    testcase["Start Time"] = start_time_str
    testcase["End Time"] = end_time_str
    testcase["Total Time (seconds)"] = total_time

    logs = []
    for message in test.findall(".//message", namespaces):
        text = message.find("text", namespaces).text.strip()
        logs.append(text)

    verification = test.find(".//verification", namespaces)
    if verification is not None:
        result = verification.find(".//scriptedVerificationResult", namespaces)
        if result is not None:
            testcase["Status"] = result.attrib["type"]

    # status is Unknown, mark it as FAIL
    if testcase["Status"] == "Unknown":
        testcase["Status"] = "FAIL"

    testcase["Logs"] = "\n".join(logs)

    testcase["Comments"] = ""  # Comments column should remain empty

    data.append(testcase)

df = pd.DataFrame(data)

pass_count = df[df["Status"] == "PASS"].shape[0]
fail_count = df[df["Status"] == "FAIL"].shape[0]
total_tests = df.shape[0]

module_name = "Login"

summary_data = {"Module Name": [module_name], "Total Test Cases": [total_tests], "Pass": [pass_count], "Fail": [fail_count]}

summary_df = pd.DataFrame(summary_data)

wb = Workbook()

ws_test_cases = wb.active
ws_test_cases.title = "Test Cases"
for r in dataframe_to_rows(df, index=False, header=True):
    ws_test_cases.append(r)

# Create a worksheet for Summary
ws_summary = wb.create_sheet("Summary")
for r in dataframe_to_rows(summary_df, index=False, header=True):
    ws_summary.append(r)

# Bar Chart for Pass/Fail
chart = BarChart()
chart.type = "col"
chart.title = "Pass/Fail Test Cases"
chart.style = 10
chart.x_axis.title = "Status"
chart.y_axis.title = "Test Cases"

data = [["Pass", pass_count], ["Fail", fail_count]]
ws_summary.append(["Status", "Count"])  # Headers for the chart data
for row in data:
    ws_summary.append(row)

data_ref = Reference(ws_summary, min_col=2, min_row=2, max_col=2, max_row=3)
categories_ref = Reference(ws_summary, min_col=1, min_row=2, max_row=3)

chart.add_data(data_ref, titles_from_data=True)
chart.set_categories(categories_ref)

ws_summary.add_chart(chart, "E5")

wb.save("test_report_with_times.xlsx")

print("Excel file with a bar graph in the Summary sheet created successfully!")
