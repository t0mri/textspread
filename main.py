from os import path
from re import search
from openpyxl import load_workbook
from sys import argv

EXIT_SUCCESS = 0
EXIT_ERROR_INVALID_ARGS_AMOUNT = 1
EXIT_ERROR_INVALID_ARGS = 2
EXIT_ERROR_FILE_NOT_FOUND = 3

MSG_USAGE = "Usage: main.py <file.xlsx:string> <file.txt:string> \
<column count:integer>\n"
MSG_ERROR_INVALID_ARGS = """
Error: Invalid amount of arguments
""" + MSG_USAGE
MSG_ERROR_XL_FILE_NOT_FOUND = """
Error: Spredsheet file not found
""" + MSG_USAGE
MSG_ERROR_TXT_FILE_NOT_FOUND = """
Error: Text file not found
""" + MSG_USAGE

if len(argv) != 4:
    print(MSG_ERROR_INVALID_ARGS)
    exit(EXIT_ERROR_INVALID_ARGS_AMOUNT)

XLFILEPATH = argv[1]
TXTFILEPATH = argv[2]
try:
    COLCOUNT = int(argv[3])
except ValueError:
    print(MSG_ERROR_INVALID_ARGS)
    exit(EXIT_ERROR_INVALID_ARGS)

if (search(".xlsx", XLFILEPATH) is None or
        search(".txt", TXTFILEPATH) is None):
    print(MSG_ERROR_INVALID_ARGS)
    exit(EXIT_ERROR_INVALID_ARGS)

if not path.isfile(XLFILEPATH):
    print(MSG_ERROR_XL_FILE_NOT_FOUND)
    exit(EXIT_ERROR_FILE_NOT_FOUND)

if not path.isfile(TXTFILEPATH):
    print(MSG_ERROR_TXT_FILE_NOT_FOUND)
    exit(EXIT_ERROR_FILE_NOT_FOUND)


file = open(TXTFILEPATH, "r")
lineTokens = file.readlines()

wb = load_workbook(XLFILEPATH)
ws = wb.active

for i in range(0, len(lineTokens), COLCOUNT):
    ws.append(lineTokens[i:COLCOUNT])

wb.save(XLFILEPATH)
exit(EXIT_SUCCESS)
