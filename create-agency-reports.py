from asyncore import write
from contextlib import nullcontext
from email.mime import base
import getpass
import pyodbc
import pandas as pd
import pandas.io.sql
import numpy as np
import openpyxl
import tkinter as tk
from openpyxl import load_workbook
from openpyxl.styles import Font
from datetime import datetime
from tkinter import filedialog


### GLOBAL VARIABLES ###
base_folder = ""
report_date = ""
output_folder = ""

# Log File


def translate_acg_score(row):
    score = row["ACG Risk Score"]

    if score == "0":
        return "Non User"
    elif score == "1":
        return "Healthy"
    elif score == "2":
        return "Low"
    elif score == "3":
        return "Moderate"
    elif score == "4":
        return "High"
    elif score == "5":
        return "Very High"


def write_to_log(message):

    log_line_stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_file.write(log_line_stamp + "\t" + message + "\n")


def main():

    global log_file
    global full_report
    global patient_counts

    user_name = getpass.getuser()
    base_folder = "C:\\Users\\{user_name}\\Value Network\\Value Network - Data Analytics\\HEALTHeLINK\\Clinical Reporting\\".format(
        user_name=user_name
    )

    log_file_stamp = datetime.now().strftime("%Y%m%d%H%M%S")
    log_file = open(
        base_folder + "\\Log Files\\agency_reports_log_" + log_file_stamp + ".log",
        mode="w",
    )

    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        initialdir=base_folder,
        title="Select HEALTHeLINK Clinical Report",
        filetypes=[("Excel Workbooks", "*.xlsx")],
    )

    output_folder = filedialog.askdirectory(
        initialdir=base_folder, title="Select Folder to Save Agency Reports"
    )

    report_date = input("Enter date of roster used to generate report:")

    write_to_log("Starting agency reports creation.")
    write_to_log("Report file: {report_file}".format(report_file=file_path))
    write_to_log("Output folder: {output_folder}".format(output_folder=output_folder))
    write_to_log("Report date: {report_date}".format(report_date=report_date))
    write_to_log("")

    write_to_log("Loading workbook...")
    workbook = load_workbook(filename=file_path)
    sheet = workbook["Sheet1"]

    sheet_data = sheet.values
    sheet_headers = next(sheet_data)[0:]

    full_report = pd.DataFrame(sheet_data, columns=sheet_headers)
    write_to_log(
        "{number_rows} read from workbook".format(number_rows=full_report.shape[0])
    )

    patient_counts = full_report[["Agency", "CID"]].groupby(["Agency"]).count()

    for group_name, agency_group in patient_counts:

        for row_index, row in agency_group.iterrows():

            agency_name = row["Agency"]
            patient_count = row["CID"]

            write_to_log(
                "{agency_name}: {patient_count}".format(
                    agency_name=agency_name, patient_count=patient_count
                )
            )

    log_file.close()


if __name__ == "__main__":
    main()
