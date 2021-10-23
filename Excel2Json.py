import xlrd # pip install xlrd==1.2.0
import json

workbook = xlrd.open_workbook("MetadataFormat.xlsx") # Excel File Name
sheet = workbook.sheet_by_name("NFT_Metadata_Table") # Excel Table Name
rows = sheet.nrows # Total number of rows
header = sheet.row_values(1)


# The NFT Metadata starts at 4th
for row in range(4,rows):
    _oneRow = []
    _json = {}
    _oneRow = sheet.row_values(row) # Get one row
    _oneRow = [i for i in _oneRow if i != '']  # Remove the null data

    for _col in range(1, len(_oneRow)):
        _json[header[_col]] = _oneRow[_col]
    with open(str(row - 4) + '.json', "a+", encoding="utf_8") as fa:
        fa.write(json.dumps(_json))
        fa.write("\n")


