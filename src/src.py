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

# TODO: open "00 Client informatiom.xlsx" with xlrd

# TODO: from sheet 1 column B, read information for insured into a dictionary

# TODO: from sheet 1 column C, read information for owner into a dictionary

# TODO: from sheet 2, read each row and store data about existing insurance into a dictionary

# TODO: from sheet 3, read each row and store data about beneficiaries into a dictionary

# TODO: understand structure of "01 config.json"
# {
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

# TODO: read and store the filenames of the pdfs in the input folder into a list

# TODO: for each item in the list created above, fill the pdf as specified in the json file


# SAMPLE CODE FOR PDF FILLING BELOW
# if __name__ == "__main__":
#     data = [
#         PdfData(
#             general_setting=PdfDataConfig(
#                 text="✓", font="Courier", font_size=16, color=Color(255, 0, 0)
#             ),
#             configs=[
#                 PdfDataConfig(location=Location(400, 505, page=3)),
#                 PdfDataConfig(location=Location(398, 541, page=3), text="X"),
#                 PdfDataConfig(location=Location(200, 360, page=4), text="01-01-1970"),
#             ],
#         ),
#         PdfData(
#             configs=[
#                 PdfDataConfig(
#                     text="obama",
#                     font_size=16,
#                     color=Color(255, 0, 255),
#                     location=Location(260, 215, page=1),
#                 )
#             ]
#         ),
#         PdfData(
#             general_setting=PdfDataConfig(
#                 text="Ann", location=Location(260, 590, page=1)
#             )
#         ),
#     ]

#     source_pdf = "4. Alpadis - CRS - Controlling Person (EN-m) 2019 07 23.pdf"

#     with open(source_pdf, "rb") as original:
#         try:
#             result = merge_pdf_with_data(original, data)
#         except PdfCreationFailed as e:
#             print(e)
#             exit(1)
#         else:
#             result.seek(0)
#             with open("result.pdf", "wb") as result_pdf:
#                 result_pdf.write(result.read())
#             print("Success!")
#             exit(0)
