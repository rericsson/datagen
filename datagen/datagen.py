import click
from datetime import date, timedelta
from random import choice, randrange
from typing import List, Tuple
import xlsxwriter


class DataColumn:

    def __init__(self, name: str, previous=0):
        self.name = name
        self.previous = previous

    def header(self):
        return self.name

    def value(self):
        return self.name

    def previous(self):
        return self.previous

class DataColumnList(DataColumn):

    def __init__(self, name: str, values: List[str]):
        super().__init__(name)
        self.values = values

    def value(self):
        return choice(self.values)

class DataColumnInteger(DataColumn):

    def __init__(self, name: str, lowValue: int = 0, highValue: int = 0):
        if lowValue > highValue:
            raise ValueError
        super().__init__(name)
        self.low = lowValue
        self.high = highValue

    def value(self):
        return randrange(self.low, self.high)

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



def create_file(directory: str, filename: str, worksheet: str, columns: List[DataColumn], rows: int):
    """ create an xslx file """
    book = xlsxwriter.Workbook(f"{directory}/{filename}")
    sheet = book.add_worksheet(worksheet)
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
                if (datacol.previous > 0):
                    start = -(rows * datacol.previous)
                    step = rows
                    data.append("".join(data[start::step]))
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
def datagen(filename, rows):
    """Create a xlsx file with fake data"""
    projects = [ "Project1", "Project2", "Project3" ]
    sites = ["Site1", "Site2", "Site3"]
    early_date = date(2020, 1, 1)
    late_date = date(2020, 4, 21)
    columns = [
                DataColumnList("Project", projects),
                DataColumnList("Site", sites),
                DataColumn("WBS", previous=2),
                DataColumn("Description"),
                DataColumnInteger("Cost", 1000, 5000),
                DataColumnDate("Start", early_date, late_date)
            ]
    create_file(".", filename,"Data", columns, rows)


if __name__ == "__main__":
    datagen()

