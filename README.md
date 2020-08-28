# Label Populator
This program can be used to fill a Microsoft Word table with Excel file data. This was primarily made to fill tables for mail labels, but any table with the correct number of columns will work. If one or more columns are selected from the Excel file, each entry of the Word table will be column-separated values.

## What is needed?
1. An Excel file (.xlsx)

&emsp;&emsp; - For a row to be classified as data, every column in that row be non-empty.

2. Word table template file (.docx)

&emsp;&emsp; - A table with at least one row. It does not need to be empty and the row count does not need to be exact.

3. (Optional) Output file name. Default name is labels.docx.
4. Excel column(s) to add to table cells

## How to run label_populator.py?
1. Download python3
2. pip install python-docx
3. pip install pandas
4. pip install xlrd
5. In the terminal, change your current directory to the folder containing label_populator.py (we'll call this current_folder for reference)
6. (Optional) Place Excel file and Word template file in current_folder. This will make the file path equal to the file name.
7. Run by using the following format terminal: "python3 label_populator.py -e \<excelFilePath> -t \<wordTemplateFilePath> -a \<listOfColumns>"

&emsp;&emsp; - Replace the brackets with your values.

&emsp;&emsp; - This template only displays the mandatory options

## Options
Mandatory options are -e, -t, and -a

|  Short options                    | Long options                          | Description                                 |
|-----------------------------------|---------------------------------------|---------------------------------------------|
| -h                                | --help                                | show this help message and exit             |
| -e \<excelFilePath>               | --excelFile \<excelFilePath>          | path to the excel file with the data        | 
| -t \<templateFilePath>            | --template \<templateFilePath>        | path to the label word template             |
| -f \<fileName>                    | --fileName \<fileName>                | output file name                            |
| -c \<columns>                     | --columns \<numberOfColumns>          | number of columns to fill in template       |
| -o \<font>                        | --font \<font>                        | font to fill template                       |
| -s \<fontSize>                    | --fontSize \<fontSize>                | font size to fill template                  |
| -a \<columnName> \<columnName>... | --args \<columnName> \<columnName>... | list of column names to merge in excel file |

## Default Values
Default values if the option is not used
|  Short options |   Default value |
|----------------|-----------------|
| -f             | labels.docx     |
| -c             | 3               |
| -o             |'Arial'          |
| -s             | 22              |


## Example
In the following example, we will first merge two columns from excel_example.xlsx ('LASTNAME' and 'FIRSTNAME') into comma-separated values. Then, we will use the template (label_template.docx) to output a new document with those comma-separated values.

**Execute:** python3 label_populator.py -e excel_example.xlsx -t label_template.docx -f output.docx -a LASTNAME FIRSTNAME

### File Structure Before Execution
```
.
├── excel_example.xlsx
├── label_populator.py
├── label_template.docx
```
#### excel_example.xlsx
<p align="center">
  <img src="/images/example_data.png" />
</p>

#### label_template.docx
<p align="center">
  <img src="/images/template.png" />
</p>

### File Structure After Execution
```
.
├── excel_example.xlsx
├── label_populator.py
├── label_template.docx
├── output.docx
```
#### output.docx
<p align="center">
  <img src="/images/output.png" />
</p>


