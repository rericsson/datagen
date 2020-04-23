import pytest
import datetime
from os import path
from openpyxl import load_workbook

from datagen.datagen import *

def xldate_to_datetime(xldate):
    """ convert Excel dates to Python dates """
    temp = datetime.date(1899,12,30)
    delta = datetime.timedelta(days=xldate)
    return temp+delta


def test_file_create(tmpdir):
    filename = "test.xlsx"
    filepath = f"{tmpdir}/{filename}"
    worksheet = "Data"
    # number of data rows
    rows = 5
    # some test data
    projects = [ "Project1", "Project2", "Project3" ]
    sites = ["Site1", "Site2", "Site3"]
    priority_values = ["1", "2", "3", "4"]
    priority_dictionary = { "1": "Emergency", "2": "High", "3": "Medium", "4": "Low" }
    early_date = datetime.date(2020, 1, 1)
    late_date = datetime.date(2020, 1, 4)
    columns = [
                DataColumnList("Project", projects),
                DataColumnList("Site", sites),
                DataColumnCombine("WBS", previous=2, delimiter="."),
                DataColumn("Description"),
                DataColumnInteger("Estimated", 1000, 5000),
                DataColumnIntegerDelta("Actual", 25),
                DataColumnDate("Start", early_date, late_date),
                DataColumnDateDelta("Finish", 5),
                DataColumnList("Priority", priority_values),
                DataColumnDictionary("Priority Text", priority_dictionary),
                DataColumnDateIncreasing("Date", early_date, 5), # 5 rows per day
                DataColumnIntegerIncreasing("Order", 400000)
            ]
    # create a file as specified
    create_file(tmpdir, filename, worksheet, columns, rows)
    assert path.exists(filepath)
    wb = load_workbook(filename=filepath)
    sheet_range = wb[worksheet]
    # check the values in the sheet
    values = zip("ABCDEFGHIJKL", columns)
    for col, val in values:
        for row in range(1, rows + 1):
            colrow = f"{col}{row}"
            print(f"{colrow}:{sheet_range[colrow].value}")
            # first row should always contain the header value
            if row == 1:
                assert sheet_range[colrow].value == val.header()
            else:
                if col == "A":
                    assert sheet_range[colrow].value in projects
                elif col == "B":
                    assert sheet_range[colrow].value in sites
                elif col == "C":
                    assert sheet_range[colrow].value == \
                        sheet_range[f"A{row}"].value + "." + sheet_range[f"B{row}"].value
                elif col in "CD":
                    assert sheet_range[colrow].value in val.value()
                elif col == "E":
                    assert isinstance(sheet_range[colrow].value, int)
                elif col == "F":
                    current_val = sheet_range[colrow].value
                    previous_val = sheet_range[f"E{row}"].value
                    assert 750 <= current_val <= 6250
                elif col == "G":
                    date_val = xldate_to_datetime(sheet_range[colrow].value)
                    assert date_val >= early_date
                    assert date_val <= late_date
                elif col == "H":
                    date_val = xldate_to_datetime(sheet_range[colrow].value)
                    assert date_val >= early_date
                elif col == "I":
                    assert sheet_range[colrow].value in priority_values
                elif col == "J":
                    priority_key = sheet_range[f"I{row}"].value
                    assert sheet_range[colrow].value == priority_dictionary[priority_key]
                elif col == "K":
                    date_val = xldate_to_datetime(sheet_range[colrow].value)
                    assert date_val >= early_date
                    assert date_val <= late_date
                elif col == "M":
                    assert isinstance(sheet_range[colrow].value, int)

