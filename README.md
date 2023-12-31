# Compte py

## Prerequis

To use this projet, you need to have python installed on you computer. Tested version : `Python 3.11`

You also need th following packages (use `pip install` to install thoses packages):
- csv
- openpyxl

## Usage

### Initiate the xlsx file

If you do not have the xlsx file, you can copy the *Dépenses_init.xlsx* file in the root of this projet as *Dépenses.xlsx*.
If you want, you can manualy add operations or begin with the current script

### Get the CSV files

To get the CSV files, go to your bank account, then to your *** and export it as a CSV file selecting the dates you want to export.

### Generate the excel file

Use the following command :

`python compte.py [--lcl <lcl_csv_path>, --bourso <boursorama_csv_path>, -i <input_specified_file>, -o <output_specified_file>]`

Options:

- **-i** Path to a specific input file, *./Dépenses.xslx* if not specified
- **-o** Path to a specific output file, *./Dépenses_\<timestamp\>.xslx* if not specified
- **--lcl** Path to the csv file of the LCL export
- **--bouso** Path to the csv file of the Boursorama export

## How does it works

The program read the csv file given as parameter. It will convert then into Operations with the same format.
Then it will read the xslx file and adjust the style of the existing files.
With the operations find in the xlsx file, the script will add them at the end of each month sheet corresponding to the month of the operations. If the sheet does not exist, it will create it and add the corresponding month in the Summary sheet.
It will export the new file.
It will also generate the csv file of all the operations

### Limitations

- For the moment, you need to specify in the code the period covered by the depenses in the xlsx file.
- The supported type of csv file are the following :
    - LCL
    - Boursorama
- You can only add new operations in the futur. It you import past or old operations, they will be added at the end of the corresponding page month
- The file complete only the following fields :
    - Date of the operation
    - Type of the operation if it recongnise it
    - Value of the operation
    - Description of the operation
    
    You may nee to correct the type of the operation

## Futur evolutions

- Add more banks
- Do not specify the start and end of the period that is in the excel file