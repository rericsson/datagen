import click
import yaml
import xlsxwriter
from datetime import date, timedelta
from random import choice, randrange
from typing import List, Tuple, Dict


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

    def __init__(self, name: str, low_value: int = 0, high_value: int = 0):
        if low_value > high_value:
            raise ValueError
        super().__init__(name)
        self.low = low_value
        self.high = high_value

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

    def __init__(self, name: str, low_value: date, high_value: date):
        if low_value > high_value:
            raise ValueError
        super().__init__(name)
        self.low = low_value
        self.high = high_value

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

def create_file(filename: str, worksheet: str, columns: List[DataColumn], rows: int):
    """ create an xslx file """
    if rows < 1:
        raise ValueError("Number of rows to be generated must be 1 or more")
    book = xlsxwriter.Workbook(filename)
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


@click.command(help="Generate data in an xlsx format following a defined configuration")
@click.option("--rows", default=1, type=int, help="Number of rows to generate")
@click.option("--filename", default="datagen.xlsx", help="Output filename. datagen.xlsx by default")
@click.option("--config", default="config.yaml", help="Configuration file. config.yaml by default")
def datagen(filename, rows, config):

    try:
        stream = open(config, "r")
        columns = yaml.full_load(stream)
        create_file(filename, "Data", columns, rows)
    except yaml.constructor.ConstructorError as err:
        print(f"The configuration has an incorrect specification: {format(err)}")
    except (FileNotFoundError, UnboundLocalError, ValueError, xlsxwriter.exceptions.FileCreateError) as err:
        print(f"Error: {format(err)}")

if __name__ == "__main__":
    datagen()

