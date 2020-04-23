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
    # some test data
    projects = [ "Project1", "Project2", "Project3" ]
    sites = ["Site1", "Site2", "Site3"]
    early_date = datetime.date(2020, 1, 1)
    late_date = datetime.date(2020, 4, 21)
    columns = [
                DataColumnList("Project", projects),
                DataColumnList("Site", sites),
                DataColumn("WBS", previous=2),
                DataColumn("Description"),
                DataColumnInteger("Cost", 1000, 5000),
                DataColumnDate("Start", early_date, late_date)
            ]
    # number of data rows
    rows = 4
    # create a file as specified
    create_file(tmpdir, filename, worksheet, columns, rows)
    assert path.exists(filepath)
    wb = load_workbook(filename=filepath)
    sheet_range = wb[worksheet]
    # check the values in the sheet
    values = zip("ABCDEF", columns)
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
                        sheet_range[f"A{row}"].value + sheet_range[f"B{row}"].value
                elif col in "CD":
                    assert sheet_range[colrow].value in val.value()
                elif col == "E":
                    assert isinstance(sheet_range[colrow].value, int)
                elif col == "F":
                    date_val = xldate_to_datetime(sheet_range[colrow].value)
                    assert date_val >= early_date
                    assert date_val <= late_date
