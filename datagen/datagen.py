import click
from datetime import date, timedelta
from random import choice, randrange
from typing import List, Tuple, Dict
import xlsxwriter


class DataColumn:

    def __init__(self, name: str):
        self.name = name

    def header(self):
        return self.name

    def value(self):
        return self.name


class DataColumnCombine(DataColumn):

    def __init__(self, name: str, previous=0, delimiter=""):
        super().__init__(name)
        self.previous = previous
        self.delimiter = delimiter

    def previous(self):
        return self.previous


class DataColumnList(DataColumn):

    def __init__(self, name: str, values: List[str]):
        super().__init__(name)
        self.values = values

    def value(self):
        return choice(self.values)

class DataColumnDictionary(DataColumn):

    def __init__(self, name: str, values: Dict):
        super().__init__(name)
        self.values = values

    def value(self, key_value: str):
        return self.values[key_value]


class DataColumnInteger(DataColumn):

    def __init__(self, name: str, lowValue: int = 0, highValue: int = 0):
        if lowValue > highValue:
            raise ValueError
        super().__init__(name)
        self.low = lowValue
        self.high = highValue

    def value(self):
        return randrange(self.low, self.high)

class DataColumnIntegerIncreasing(DataColumn):

    def __init__(self, name: str, start: int = 0):
        super().__init__(name)
        self.start = start
        self.last = 0

    def value(self):
        if self.last == 0:
            self.last = self.start
        else:
            self.last += 1
        return self.last

class DataColumnIntegerDelta(DataColumn):

    def __init__(self, name: str, delta: int):
        if delta < 0:
            raise ValueError
        super().__init__(name)
        self.delta = delta

    def value(self, last_value: int):
        high_value = int(last_value + (self.delta/100 * last_value))
        low_value = int(last_value - (self.delta/100 * last_value))
        return randrange(low_value, high_value)


class DataColumnDate(DataColumn):

    def __init__(self, name: str, lowValue: date, highValue: date):
        if lowValue > highValue:
            raise ValueError
        super().__init__(name)
        self.low = lowValue
        self.high = highValue

    def value(self):
        time_between_dates = self.high - self.low
        days_between_dates = time_between_dates.days
        random_number_days = randrange(days_between_dates)
        random_date = self.low + timedelta(days=random_number_days)
        return random_date

class DataColumnDateDelta(DataColumn):

    def __init__(self, name: str, delta: int):
        if delta < 0:
            raise ValueError
        super().__init__(name)
        self.delta = delta

    def value(self, last_value: date):
        random_days = timedelta(days=randrange(self.delta))
        return last_value + random_days

class DataColumnDateIncreasing(DataColumn):

    def __init__(self, name: str, start: date, rows_per_day: int):
        if rows_per_day <= 0:
            raise ValueError
        super().__init__(name)
        self.start = start
        self.rows_per_day = rows_per_day
        self.last_date = 0
        self.last_row = 0

    def value(self):
        self.last_row += 1
        if isinstance(self.last_date, date):
            if self.last_row % self.rows_per_day == 0:
                self.last_date += timedelta(days=1)
        else:
            self.last_date = self.start
        return self.last_date

def create_file(directory: str, filename: str, worksheet: str, columns: List[DataColumn], rows: int):
    """ create an xslx file """
    book = xlsxwriter.Workbook(f"{directory}/{filename}")
    sheet = book.add_worksheet(worksheet)
    # create a list to store the data before it is written to the sheet
    data = []
    # write the columns to the spreadsheet
    for col, datacol in enumerate(columns):
        for row in range(rows):
            if row == 0:
                data.append(datacol.header())
                sheet.write(row, col, data[-1])
            else:
                # if we are combining previous columns, need to start
                # at the corresponding value for the previous row
                # and step through by the number of rows
                if isinstance(datacol, DataColumnCombine):
                    start = -(rows * datacol.previous)
                    step = rows
                    delim = datacol.delimiter
                    data.append(delim.join(data[start::step]))
                # if we are using a delta, get the last value and apply delta
                elif isinstance(datacol,
                        (DataColumnIntegerDelta, DataColumnDateDelta, DataColumnDictionary)):
                    last_value = data[-rows]
                    current_value = datacol.value(last_value)
                    data.append(current_value)
                else:
                    data.append(datacol.value())
                # write the data to the sheet being careful with dates
                if isinstance(data[-1], date):
                    sheet.write_datetime(row, col, data[-1])
                else:
                    sheet.write(row, col, data[-1])
    book.close()

@click.command()
@click.option('--rows', default=1, help="Number of rows")
@click.option('--filename', default="datagen.xlsx")
@click.option('--directory', default='.')
def datagen(directory, filename, rows):
    """Create a xlsx file with fake data"""
    projects = [ "Project1", "Project2", "Project3" ]
    sites = ["Site1", "Site2", "Site3"]
    priority_values = ["1", "2", "3", "4"]
    priority_dictionary = { "1": "Emergency", "2": "High", "3": "Medium", "4": "Low" }
    lat_values = ["37.778549194336", "37.77717590332", "37.780332"]
    lat_dictionary = {"37.778549194336": "-122.42134094238","37.77717590332": "-122.4227142334","37.780332": "-122.418898" }
    early_date = date(2020, 1, 1)
    late_date = date(2020, 1, 8)
    columns = [
                DataColumnIntegerIncreasing("Order", 40000),
                DataColumnList("Project", projects),
                DataColumnList("Site", sites),
                DataColumnCombine("WBS", previous=2, delimiter="."),
                DataColumn("Description"),
                DataColumnInteger("Estimated", 1000, 5000),
                DataColumnIntegerDelta("Actual", 25),
                DataColumnDateIncreasing("Start", early_date, 6),
                DataColumnDateDelta("Finish", 5),
                DataColumnList("Priority", priority_values),
                DataColumnDictionary("Priority Text", priority_dictionary),
                DataColumnList("Lat", lat_values),
                DataColumnDictionary("Lon", lat_dictionary)
            ]
    create_file(directory, filename, "Data", columns, rows)


if __name__ == "__main__":
    datagen()

