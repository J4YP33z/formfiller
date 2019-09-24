from pdf_filler import (
    merge_pdf_with_data,
    PdfCreationFailed,
    PdfData,
    PdfDataConfig,
    Location,
    Color,
)
import xlrd
import json
import os.path
import glob
import fnmatch


def getRelativePath(fileName: str):
    # get the path of the current folder
    fileDir = os.path.dirname(os.path.realpath(__file__))
    # get the file in a folder contained in the current folder
    path = os.path.join(fileDir, fileName)
    return path


def readDataVertical(
    fileName: str,
    sheetName: str,
    startRow: int,
    keyColumn_idx: int,
    valueColumn_idx: int,
):
    file_path = getRelativePath(fileName)
    # open workbook
    _workbook = xlrd.open_workbook(file_path)
    _worksheet = _workbook.sheet_by_name(sheetName)
    # create new dictionary
    new_dict = {}
    # add policy number at A2 cell in excel file
    new_dict["POLICY NUMBER"] = str(int(_worksheet.cell_value(0, 1)))
    # the loop start from the input start row to the last row of the sheet
    for row_idx in range(startRow, _worksheet.nrows):
        # save the key and the value into the dictionary
        new_dict[_worksheet.cell_value(row_idx, keyColumn_idx)] = str(
            _worksheet.cell_value(row_idx, valueColumn_idx)
        )
    return new_dict


def readDataHorizontal(
    fileName: str,
    sheetName: str,
    startRow_idx: int,
    startCol_idx: int,
    keyCol_idx: int,
    keyRow_idx: int,
):
    file_path = getRelativePath(fileName)
    # open workbook
    _workbook = xlrd.open_workbook(file_path)
    # open worksheet
    _worksheet = _workbook.sheet_by_name(sheetName)
    # create new dictionary
    new_dict = {}
    # add policy number at A2 cell in excel file
    new_dict["POLICY NUMBER"] = str(int(_worksheet.cell_value(0, 1)))
    for row_idx in range(startRow_idx, _worksheet.nrows):
        # get the key value for the outer dictionary
        _key = str(_worksheet.cell_value(row_idx, keyCol_idx))
        # create a dictionary for the value of the outer dictionary
        value_dict = {}
        for col_idx in range(startCol_idx, _worksheet.ncols):
            # save column name as key and  the cell as value of the inner dictionary
            value_dict[_worksheet.cell_value(keyRow_idx, col_idx)] = str(
                _worksheet.cell_value(row_idx, col_idx)
            )
        # assign value for outer dictionary
        new_dict[_key] = value_dict
    return new_dict


# JSON structure of "01 templates.json"
#     "FILENAME": [
#         {
#             "text": "<RETRIEVE DATA FROM DICTIONARIES CREATED ABOVE>",
#             "page": "<PAGE NUMBER>",
#             "x": "<X POSITION IN PDF FILE>",
#             "y": "<Y POSITION IN PDF FILE>",
#             "fontsize": "<FONT SIZE>",
#         }
#     ]
# }


def getTemplatesFromJsonFile(fileName: str):
    """
    return a dictionary of form templates
    """
    file_path = getRelativePath(fileName)
    with open(file_path, "r") as jsonFile:
        data_str = jsonFile.read()
    templates_dict = json.loads(data_str)
    return templates_dict


def getTemplatesFromXlsxFile(fileName: str):
    result = {}  # { filename : list of dictionaries }
    file_path = getRelativePath(fileName)
    wb = xlrd.open_workbook(file_path)
    for sheetIndex in range(0, len(wb.sheets())):
        ws = wb.sheet_by_index(sheetIndex)
        if ws.nrows < 4 or ws.ncols < 7:
            continue  # skip defective templates
        infoList = []
        for r in range(3, ws.nrows):
            rowDict = {}
            for c in range(1, ws.ncols):
                rowDict[ws.cell_value(2, c)] = ws.cell_value(r, c)
            infoList.append(rowDict)
        result[ws.cell_value(0, 1)] = infoList
    return result


def getAllFileFromInputFolder():
    files = []
    for file in os.listdir("src/input/"):
        if fnmatch.fnmatch(file, "*.pdf"):
            files.append(file)
    return files


def setPDFData(template: list, data_dict: list):
    dataList = []
    for row in template:
        # print(data_dict[row["INFO"]], row["PAGE"], row["X"], row["Y"], row["FONTSIZE"])
        tmpPDFdata = PdfData(
            configs=[
                PdfDataConfig(
                    text=data_dict[row["INFO"]],
                    font_size=row["FONTSIZE"],
                    location=Location(
                        int(row["X"]), int(row["Y"]), page=int(row["PAGE"] - 1)
                    ),
                )
            ]
        )
        dataList.append(tmpPDFdata)
    return dataList


# program entry point
excel_fname = "input/00 Client information.xlsx"
# create and read the insured info. start from 4th row, keyCol_idx=0, valueCol_idx=1
insured_dict = readDataVertical(excel_fname, "Info", 3, 0, 1)
insured_dict["TICK"] = "✓"
insured_dict["CROSS"] = "X"
# create owner dictionary
owner_dict = readDataVertical(excel_fname, "Info", 3, 0, 2)
# create list of dictionaries for beneficiaries
beneficiaries = []
_workbook = xlrd.open_workbook(getRelativePath(excel_fname))
_worksheet = _workbook.sheet_by_index(0)
for i in range(3, _worksheet.ncols):
    beneficiaries.append(readDataVertical(excel_fname, "Info", 3, 0, i))

# read json file
# templates_ = getTemplatesFromJsonFile("input/01 templates.json")

# read templates file
templates = getTemplatesFromXlsxFile("input/01 templates.xlsx")

# get all pdf file in input folder
files = getAllFileFromInputFolder()
# fill the form one by one
for _file in files:
    # get all fields in template by name
    field_list = templates.get(_file)
    # handle error "Not found file name in templates.json"
    if field_list == None:
        print("template for " + _file + " not found!")
        continue
    # read, merge and write info to the file
    with open(os.path.join("src/input/", _file), "rb") as original:
        try:
            # set data for form filling
            result = merge_pdf_with_data(original, setPDFData(field_list, insured_dict))
        except PdfCreationFailed as e:
            print(e)
            exit(1)
        else:
            result.seek(0)
            with open(os.path.join("src/output/", _file), "wb") as result_pdf:
                result_pdf.write(result.read())
                print(_file + " filled!")
exit(0)
