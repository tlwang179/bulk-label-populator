# Label Populator
This program can be used to fill a Microsoft Word table with Excel file data. This was primarily made to fill tables for mail labels, but any table with the correct number of columns will work. If one or more columns are selected from the Excel file, each entry of the Word table will be column separated values.

## What is needed?
1. An Excel file (.xlsx)
  - All data rows must fill every column.
2. Word table template file (.docx)
  - A table with at least one row.
3. (Optional) Output file name
4. Excel column(s) to add to table cells

## How to run label_populator.py?
1. Download python3
2. pip install docx
3. pip install pandas
4. pip install xlrd
5. In the terminal, change your current directory to the folder containing label_populator.py (we'll call this parent_folder for reference)
6. (Optional) Place Excel file and Word template file in parent_folder. This will make the file path the same as the file name.
7. Run by typing the following in terminal: "python3 -e \<excelFilePath> -t \<wordTemplateFilePath> -f \<outputFileName> -a <listOfColumns>"

## Options
  Short options                     | Long options                          | Description                                 |
|-----------------------------------|---------------------------------------|---------------------------------------------|
| -h                                | --help                                | show this help message and exit             |
| -e \<excelFilePath>               | --excelFile \<excelFilePath>          | path to the excel file with the data        | 
| -t \<templateFilePath>            | --template \<templateFilePath>        | path to the label word template             |
| -f \<fileName>                    | --fileName \<fileName>                | output file name                            |
| -c \<columns>                     | --columns \<columns>                  | number of columns to fill in template       |
| -o \<font>                        | --font \<font>                        | font to fill template                       |
| -s \<fontSize>                    | --fontSize \<fontSize>                | font size to fill template                  |
| -a \<columnName> \<columnName>... | --args \<columnName> \<columnName>... | list of column names to merge in excel file |



## Example
