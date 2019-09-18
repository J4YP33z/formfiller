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

# TODO: open "00 Client informatiom.xlsx" with xlrd
def getRelativePath(fileName: str):
    # get the path of the current folder
    fileDir = os.path.dirname(os.path.realpath(__file__))
    # get the file in a folder contained in the current folder
    path = os.path.join(fileDir, fileName)
    return path


# TODO: from sheet 1 column B, read information for insured into a dictionar
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


# TODO: from sheet 1 column C, read information for owner into a dictionary

# TODO: from sheet 2, read each row and store data about existing insurance into a dictionary

# TODO: from sheet 3, read each row and store data about beneficiaries into a dictionary
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
    #
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


# TODO: understand structure of "01 config.json"
#
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

# TODO: read and store data from "01 config.json"
def getTemplatesFromJsonFile(fileName: str):
    """
    return a dictionary of form templates
    """
    file_path = getRelativePath(fileName)
    with open(file_path, "r") as jsonFile:
        data_str = jsonFile.read()
    templates_dict = json.loads(data_str)
    return templates_dict


# TODO: read and store the filenames of the pdfs in the input folder into a list
def getAllFileFromInputFolder():
    files = []
    for file in os.listdir("src/input/"):
        if fnmatch.fnmatch(file, "*.pdf"):
            files.append(file)
    return files


# TODO: for each item in the list created above, fill the pdf as specified in the json file
def setPDFData(template: list, data_dict: list):

    dataList = []

    # get first value of template
    _firstValue = next(iter(template))
    # create new pdfData from first value
    _generalSetting = PdfDataConfig(
        text=insured_dict[_firstValue["text"]], font_size=int(_firstValue["fontsize"])
    )
    _configs = [
        PdfDataConfig(
            location=Location(
                int(_firstValue["x"]), int(_firstValue["y"]), int(_firstValue["page"])
            )
        )
    ]
    _newPdfData = PdfData(general_setting=_generalSetting, configs=_configs.copy())
    # set first value for dataList
    dataList.append(_newPdfData)

    for field_idx in range(1, len(template)):
        field = template[field_idx]

        # check if current field exist in dataList of not
        #   if exist:
        #       add new location for the exist field
        #   else:
        #       add current field to dataList
        fieldExist = False
        for data in dataList:

            if data_dict[field["text"]] != data.general_setting.text:
                continue
            else:
                # field exist
                fieldExist = True
                _newLocation = Location(
                    int(field["x"]), int(field["y"]), page=int(field["page"])
                )
                data.configs.append(PdfDataConfig(location=_newLocation))
                break
        if fieldExist == False:
            # field not exist
            _generalSetting = PdfDataConfig(
                text=data_dict[field["text"]], font_size=int(field["fontsize"])
            )
            _configs = [
                PdfDataConfig(
                    location=Location(
                        int(field["x"]), int(field["y"]), int(field["page"])
                    )
                )
            ]
            _newPdfData = PdfData(
                general_setting=_generalSetting, configs=_configs.copy()
            )
            dataList.append(_newPdfData)
    return dataList


# SAMPLE CODE FOR PDF FILLING BELOW
if __name__ == "__main__":
    excel_fname = "input/00 Client information.xlsx"
    # create insured dictionary
    insured_dict = {}
    # read the insured info start from 4th row, keyCol_idx=0, valueCol_idx=1
    insured_dict = readDataVertical(excel_fname, "Info", 3, 0, 1)
    # create owner dictionary
    owner_dict = {}
    owner_dict = readDataVertical(excel_fname, "Info", 3, 0, 2)
    # create baneficiaries dictionary
    beneficiaries_dict = {}
    beneficiaries_dict = readDataHorizontal(excel_fname, "Beneficiaries", 2, 2, 1, 1)

    # get key value pairs from "Beneficiaries" sheet by "ENGLISH NAME"
    # 1. Get key for searching
    # 2. Get the values
    # 3. Update the insured dictionary
    searchingKey = insured_dict["ENGLISH NAME"]
    beneficiaries_info = beneficiaries_dict[searchingKey]
    insured_dict.update(beneficiaries_info)
    # read json file
    templates = getTemplatesFromJsonFile("input/templates.json")
    # get all fields in template by name
    field_list = templates["5a.  Sun Life VOT rev (adult).pdf"]
    # get all pdf file in input folder
    files = getAllFileFromInputFolder()
    print(files)
    #test commit 
    # set data for form filling
    # dataList = setPDFData(field_list, insured_dict)

    # fill the form
    # print("merging... please keep calm and wait...!")
    # source_pdf = "5a.  Sun Life VOT rev (adult).pdf"
    # with open(source_pdf, "rb") as original:
    #     try:
    #         result = merge_pdf_with_data(original, dataList)
    #     except PdfCreationFailed as e:
    #         print(e)
    #         exit(1)
    #     else:
    #         result.seek(0)
    #         with open("result.pdf", "wb") as result_pdf:
    #             result_pdf.write(result.read())
    #         print("Success!")
    #         exit(0)
    # data = [
    #     PdfData(
    #         general_setting=PdfDataConfig(
    #             text="✓", font="Courier", font_size=16, color=Color(255, 0, 0)
    #         ),
    #         configs=[
    #             PdfDataConfig(location=Location(400, 505, page=3)),
    #             PdfDataConfig(location=Location(398, 541, page=3), text="X"),
    #             PdfDataConfig(location=Location(200, 360, page=4), text="01-01-1970"),
    #         ],
    #     ),
    #     PdfData(
    #         configs=[
    #             PdfDataConfig(
    #                 text="obama",
    #                 font_size=16,
    #                 color=Color(255, 0, 255),
    #                 location=Location(260, 215, page=1),
    #             )
    #         ]
    #     ),
    #     PdfData(
    #         general_setting=PdfDataConfig(
    #             text="Ann", location=Location(260, 590, page=1)
    #         )
    #     ),
    # ]

    # source_pdf = "4. Alpadis - CRS - Controlling Person (EN-m) 2019 07 23.pdf"

    # with open(source_pdf, "rb") as original:
    #     try:
    #         result = merge_pdf_with_data(original, data)
    #     except PdfCreationFailed as e:
    #         print(e)
    #         exit(1)
    #     else:
    #         result.seek(0)
    #         with open("result.pdf", "wb") as result_pdf:
    #             result_pdf.write(result.read())
    #         print("Success!")
    #         exit(0)

