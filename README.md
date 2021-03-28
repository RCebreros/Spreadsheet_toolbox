Spreadsheet_toolbox

Contains tools used to work with spreadsheets.

Required Python Libraries:

openpyxl
datetime
os
sys
Contains the following functions, function names are capitalized for readability purposes only:

GET_FOLDER_RECORDS ( PATH-NAME ) - Returns all spreadsheet contents in a given directory > Gathers CSV and XLML speadsheets. > Returns a 2-dimensional list. > Path-Name refers to a directory path.

PARSE_CSV (FILE) - Returns all contents from a CSV document in a list. > Works on comma-delimented spreadsheets.

PARSE_XLSX (FILE) - Returns all contents from an XLSX document in a list. > file must end in XLSX.

SHEET_TO_DICT (FILE-PATH, KEY-INDEX, INCLUDE-SHORTS= False, APPEND_REPEATS= True) - Returns a dictionary from a given XLSX spreadsheet.
> KEY-INDEX will be the row index (cell) that will be the designated key. > DThe dictionary value defauls to the entire row as a list. > If APPEND_REPEATS is set to True, we will append rows with repeated keys. > INCLUDE_SHORTS - If True, the rows shorter than the designated key will be appended as well under a key names "shortlines"

GET_COLUMN (FILE, INDEX ) - Returns a spreadsheet column in a list. > Supports CSV and XLSX.

SAVE_TO_SHEET (DATA, TITLE, PARSE=None) - Saves a 2 dimensional list to a spreadsheet. > Every element in the list/tuple becomes a row. > DATA - refers to the dimensional list. > TITLE - refers to the title it will be saved under. > PARSE DEFAULTS TO NONE- Defauls to XLSX, where individual elements are saved in individual cells, each list being a row of its own. > PARSE 'CSV' - 'CSV' can be chosen to split the list by the commas and save each split into individual cells.

DATE_OBJECT (DATE) - Returns a datetime object. > Currently excludes time entries. > Returns a datetime object based on a string date. > Could be used to compare dates agaisnt each-other.

STRING_DATE (DATE) - Return a string object. > Takes a datetime object and returns it in string format.
