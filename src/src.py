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

print("hello world")

data = [
    PdfData(
        general_setting=PdfDataConfig(
            text="✓", font="Courier", font_size=16, color=Color(255, 0, 0)
        )
    )
]


# OLD CODE BELOW
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
