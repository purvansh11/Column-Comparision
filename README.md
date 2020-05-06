# Column-Comparision
Column-Comparision is a Python desktop application used for mapping 2 columns with a few similarities using fuzzywuzzy algorithm.

## Installation
Use the package manager [pip](https://pip.pypa.io/en/stable/) to install the required libraries.

```bash 
pip install pandas
pip install tkintertable
pip install pandas
pip install xlwt
pip install fuzzywuzzy
```
## Running the test
```bash
git clone https://github.com/purvansh11/Column-Comparision.git
```
- Run code_p.py
- Click Import Excel File
- Import *Mapped sku.xlsx* from the cloned folder or any other similar file 
- Close GUI window

## Output
- Navigate to the folder where git clone is done
- *Comparision_Output.xls* is the output file
  - Sheet 1 : Green Highlighted cells display the correct matches with their score
  - Sheet 2 : All correct matches from Sheet 1 collectively displayed
