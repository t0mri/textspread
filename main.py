from os import path
from re import search
from openpyxl import load_workbook
from sys import argv

EXIT_SUCCESS = 0
EXIT_ERROR_INVALID_ARGS_AMOUNT = 1
EXIT_ERROR_INVALID_ARGS = 2
EXIT_ERROR_FILE_NOT_FOUND = 3

MSG_USAGE = "Usage: main.py <file.xlsx> <file.txt>\n"
MSG_ERROR_ARGS = """
Error: Invalid amount of arguments
""" + MSG_USAGE
MSG_ERROR_XL_FILE_NOT_FOUND = """
Error: Spredsheet file not found
""" + MSG_USAGE
MSG_ERROR_TXT_FILE_NOT_FOUND = """
Error: Text file not found
""" + MSG_USAGE

if len(argv) != 3:
    print(MSG_ERROR_ARGS)
    exit(EXIT_ERROR_INVALID_ARGS_AMOUNT)

if search(".xlsx", argv[1]) is None or search(".txt", argv[2]) is None:
    print(MSG_ERROR_ARGS)
    exit(EXIT_ERROR_INVALID_ARGS)

XLFILEPATH = argv[1]
TXTFILEPATH = argv[2]

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
exit(EXIT_SUCCESS)
