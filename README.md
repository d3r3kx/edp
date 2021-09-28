# Excel Data Processing
### Requirements
* python: 3
* pandas: 1.3.3
* openpyxl: 3.0.9
Other versions of python and libraries can also be attempted.
  
### Assumptions
1. The format of input dataframe is a table of numerical values:
   1. No `NaN` or invalid values.
   2. All columns have the same number of rows.
2. The column names are `A_DATA`, `A_DATA`, `A_DATA`, etc. If other column names are used, `main.py` needs updates
   correspondingly.

### Setup
1. Run `pip install -r requirements.txt` in command line to install the dependencies.
2. Run an example: `python main.py -i Book1.xlsx -o output.xlsx` in command line, where `Book1.xlsx` is the input path, and 
   `output.xlsx` is the output file path.

