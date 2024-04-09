# ScheduleChronogram

Generates a visual chronogram in Excel, mapping out task hours across workweeks to aid in project management.

## Prerequisites
### Python:
Please note that Python 3 must be installed on your system to use this script. If you do not have Python 3, please install it from the [official Python website](https://www.python.org/) or use your system's package manager.

### Python Libraries:
Before running this script, you must have the following Python libraries installed:
If you do not have Python 3, please install it from the official Python website or use your system's package manager.

- `pandas`
- `openpyxl`

You can install these libraries using pip with the following command:

- `pip install pandas`
- `pip install openpyxl`

## Usage Instructions

1. **Prepare the Scripts**  
   Ensure that `chronogram.py` and `chronogram.sh` are both located in the same directory.

2. **Set Script Permissions**  
   Open a terminal and navigate to the directory containing the scripts. Give executable permissions to the shell script using the command:
   - `chmod +x chronogram.sh`

3. **Execute the Script**
   Run the script by typing the following command into the terminal:
- `chronogram.sh`

4. **Enter Task Hours**
   When prompted with Add tasks hours (as comma-separated values):, input your task hours. Example input:
- 40, 40, 60, 32, 8, 24, 40, 160
  
5. **Access the Chronogram**
   After providing the input, the script will generate two files in the same directory:
- `chronogram.xlsx: An Excel file with the visual chronogram.`
- `chronogram.csv: A CSV file with the data used to generate the chronogram.`

6. **Open and view the Chronogram**
  Open the chronogram.xlsx file in Excel to view your visual chronogram.
  - For this input: 40, 40, 60, 32, 8, 24, 40, 160, the excel file would display a chronogram like this example:
![Chronogram Example](./chronogram%20excel.png)

