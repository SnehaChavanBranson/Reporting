import xml.etree.ElementTree as ET
import pandas as pd
import re
from datetime import datetime
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

tree = ET.parse("results.xml")
root = tree.getroot()

namespaces = {"": "http://www.froglogic.com/resources/schemas/xml3"}

data = []


def convert_to_datetime(time_str):
    return datetime.strptime(time_str, "%Y-%m-%dT%H:%M:%S.%fZ")


serial_number = 1
for test in root.findall(".//test", namespaces):
    test_name = test.find(".//prolog/name", namespaces).text.strip()

    if not re.search(r"test case \d+", test_name, re.IGNORECASE):
        continue

    testcase = {}
    testcase["Serial Number"] = serial_number
    serial_number += 1
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

    if testcase["Status"] == "Unknown":
        testcase["Status"] = "FAIL"

    testcase["Logs"] = "\n".join(logs)
    testcase["Comments"] = ""
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

header_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
bold_font = Font(bold=True, color="FFFFFF")

for cell in ws_test_cases[1]:
    cell.fill = header_fill
    cell.font = bold_font

for row in ws_test_cases.iter_rows(min_row=2, min_col=7, max_col=7):
    for cell in row:
        cell.alignment = Alignment(wrap_text=True)

thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

for row in ws_test_cases.iter_rows(min_row=1, min_col=1, max_row=ws_test_cases.max_row, max_col=ws_test_cases.max_column):
    for cell in row:
        cell.border = thin_border

pass_fill = PatternFill(start_color="228B22", end_color="228B22", fill_type="solid")
fail_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
bold_font_pass_fail = Font(bold=True)

for row in ws_test_cases.iter_rows(min_row=2, min_col=3, max_col=3):
    for cell in row:
        if cell.value == "PASS":
            cell.fill = pass_fill
            cell.font = bold_font_pass_fail
        elif cell.value == "FAIL":
            cell.fill = fail_fill
            cell.font = bold_font_pass_fail

ws_summary = wb.create_sheet("Summary")
for r in dataframe_to_rows(summary_df, index=False, header=True):
    ws_summary.append(r)

chart = BarChart()
chart.type = "col"
chart.title = "Pass/Fail Test Cases"
chart.style = 10
chart.x_axis.title = "Status"
chart.y_axis.title = "Test Cases"

data = [["Pass", pass_count], ["Fail", fail_count]]
ws_summary.append(["Status", "Count"])
for row in data:
    ws_summary.append(row)

data_ref = Reference(ws_summary, min_col=2, min_row=2, max_col=2, max_row=3)
categories_ref = Reference(ws_summary, min_col=1, min_row=2, max_row=3)

chart.add_data(data_ref, titles_from_data=True)
chart.set_categories(categories_ref)

for series in chart.series:
    series.dLbls = DataLabelList()
    series.dLbls.show_val = True
    series.dLbls.number_format = "0"

ws_summary.add_chart(chart, "E5")

wb.save("test_report_with_times.xlsx")

print("Excel file with a bar graph displaying Pass and Fail test cases and numbers created successfully!")
