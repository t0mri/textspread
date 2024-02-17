from os import path
from re import search
from openpyxl import load_workbook
from sys import argv

EXIT_SUCCESS = 0
EXIT_ERROR_INVALID_ARGS_AMOUNT = 1
EXIT_ERROR_INVALID_ARGS = 2
EXIT_ERROR_FILE_NOT_FOUND = 3

ERROR_ARGS = """
Error: Invalid amount of arguments
Usage: main.py <file.xlsx> <file.txt>
"""
ERROR_XL_FILE_NOT_FOUND = """
Error: Spredsheet file not found
Usage: main.py <file.xlsx> <file.txt>
"""
ERROR_TXT_FILE_NOT_FOUND = """
Error: Text file not found
Usage: main.py <file.xlsx> <file.txt>
"""

if len(argv) != 3:
    print(ERROR_ARGS)
    exit(EXIT_ERROR_INVALID_ARGS_AMOUNT)

if search(".xlsx", argv[1]) is None or search(".txt", argv[2]) is None:
    print(ERROR_ARGS)
    exit(EXIT_ERROR_INVALID_ARGS)

XLFILEPATH = argv[1]
TXTFILEPATH = argv[2]

if not path.isfile(XLFILEPATH):
    print(ERROR_XL_FILE_NOT_FOUND)
    exit(EXIT_ERROR_FILE_NOT_FOUND)

if not path.isfile(TXTFILEPATH):
    print(ERROR_TXT_FILE_NOT_FOUND)
    exit(EXIT_ERROR_FILE_NOT_FOUND)


file = open(TXTFILEPATH, "r")
lineTokens = file.readlines()

wb = load_workbook(XLFILEPATH)
ws = wb.active

for i in range(0, len(lineTokens), 8):
    question = lineTokens[i]
    optA = lineTokens[i + 1]
    optB = lineTokens[i + 2]
    optC = lineTokens[i + 3]
    optD = lineTokens[i + 4]
    ans = lineTokens[i + 5]
    exp = lineTokens[i + 6]

    ws.append((question, optA, optB, optC, optD, ans, exp))

wb.save(XLFILEPATH)
