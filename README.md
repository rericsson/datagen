# Datagen - create demo data in an xlsx file

## Installation
These instructions are for Windows 10. The process is slightly different on macOS and Linux. 
Ensure that you have Python installed on your computer.
Datagen was created with Python 3.8 so that version is best. YMMV with other versions. 

First, download the code from GitHub:

```
git clone  https://github.com/rericsson/datagen.git
```

After you have the code, switch to the directory where it was downloaded and run:

```
pip install .
```

That will install the datagen project in your local system. If you would rather do it in a virtual environment, you can do that by doing the following before running `pip install`:

```
virtualenv venv
. venv/bin/activate
pip install .
```
## Running

After you have done the install, you can run it with:
```
python -m datagen --help
```

That will show you the options available:

```
Usage: datagen [OPTIONS]

Generate data in an xlsx format following a defined configuration

Options:
--rows INTEGER   Number of rows to generate
--filename TEXT  Output filename. datagen.xlsx by default
--config TEXT    Configuration file. config.yaml by default
--help           Show this message and exit.
```

To get started, try running: 
```
python -m datagen --config config.yaml --rows 100
```

That will generate a file called datagen.xlsx with 100 rows of data as specified in the config.yaml file. Note that the generated values have not been formatted and you should apply appropriate formatting for currency or dates. 

## Configuration

The configuration determines the columns created and their type. It is done in yaml format with an example looking like:

```
- !!python/object:__main__.DataColumnIntegerIncreasing
  last: 0
  name: Order
  start: 40000
- !!python/object:__main__.DataColumnList
  name: Project
  values:
  - Project1
  - Project2
  - Project3
- !!python/object:__main__.DataColumnList
  name: Site
  values:
  - Site1
  - Site2
  - Site3
- !!python/object:__main__.DataColumnCombine
  delimiter: .
  name: WBS
  previous: 2
- !!python/object:__main__.DataColumn
  name: Description
- !!python/object:__main__.DataColumnInteger
  high: 5000
  low: 1000
  name: Estimated
- !!python/object:__main__.DataColumnIntegerDelta
  delta: 25
  name: Actual
- !!python/object:__main__.DataColumnDateIncreasing
  last_date: 0
  last_row: 0
  name: Start
  rows_per_day: 6
  start: 2020-01-01
- !!python/object:__main__.DataColumnDateDelta
  delta: 5
  name: Finish
- !!python/object:__main__.DataColumnList
  name: Priority
  values:
  - '1'
  - '2'
  - '3'
  - '4'
- !!python/object:__main__.DataColumnDictionary
  name: Priority Text
  values:
    '1': Emergency
    '2': High
    '3': Medium
    '4': Low
- !!python/object:__main__.DataColumnList
  name: Lat
  values:
  - '37.778549194336'
  - '37.77717590332'
  - '37.780332'
- !!python/object:__main__.DataColumnDictionary
  name: Lon
  values:
    '37.77717590332': '-122.4227142334'
    '37.778549194336': '-122.42134094238'
    '37.780332': '-122.418898'
```

This shows all of the different types of columns available and how they are used. Please note that some columns are dependent on the columns appearing before them. For example, the DataColumnIntegerDelta must appear after an integer column and will create a value that is +/- the delta percentage from the previous column. The DataColumnDateDelta column is similar. The DataColumnDictionary looks up a value from the given dictionary from the previous DataColumnList value. 

To add new columns, just add a new `!!python/object` to the list of columns. If you delete a `!!python/object` entry in the configuration file, the correspoding column will no longer appear in the output. 

Note that there are some values that are necessary for object initialization and should be set to 0. These are last in DataColumnIntegerIncreasing and last_date and last_row in DataColumnDateIncreasing. A future version will improve the object loader so that these values are not required. 
