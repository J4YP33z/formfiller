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
if __name__ == "__main__":
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
########################################################################################### 
    # def getDataFromWorkSheetHorizontal(**kwargs, fileName:str, sheetName: str, startRow:int , keyColumn:int ,valueColumn:int):
    #     #open workbook
    #     _workbook = xlrd.open_workbook(fileName)
    #     #list the sheetname 
    #     _worksheetList = _workbook.sheet_names()
    #     print("List of sheets:",_worksheetList )
    #     _worksheet = _worksheetList.sheet_by_name(sheet_name)
    #     #add policy number at A2 cell in excel file
    #     kwargs['POLICY NUMBER'] = str(int(_worksheet.cell_value(0,1)))
    #     for row_idx in range(startRow,_worksheet.nrows):
    #         kwargs[_worksheet.cell_value(row_idx,keyColumn)] = str(_worksheet.cell_value(row_idx,valueColumn))
    #     print("data dictionary:",kwargs)

    # def getDataFromWorkSheetVerical(**kwargs, fileName:str, sheetName: str,startRow:str, startCol:int, keyVar:int,valueColumn:int):
    #      #open workbook
    #     _workbook = xlrd.open_workbook(fileName)
    #     #list the sheetname 
    #     _worksheetList = _workbook.sheet_names()
    #     print("List of sheets:",_worksheetList )
    #     _worksheet = _worksheetList.sheet_by_name(sheet_name)
    #     #add policy number at A2 cell in excel file
    #     kwargs['POLICY NUMBER'] = str(int(_worksheet.cell_value(0,1)))

    #     for row in range(startRow,xl_worksheet.nrows):
    #         _key[xl_worksheet.cell_value(row ,1)]={}
    #             for col in range(startCol,xl_worksheet.ncols):
    #                 _value=str(xl_worksheet.cell_value(row ,col))
    #                 print(xl_worksheet.cell_value(1 ,col))
    #                 data[_key][xl_worksheet.cell_value(1 ,col)]=_value
    #     print("data:", data)
      
    

    excel_fname="00 Client information.xlsx"

    #open workbook 
    xl_workbook = xlrd.open_workbook(excel_fname)
    #list the sheetname
    worksheets = xl_workbook.sheet_names()
    print("worksheet",worksheets)
    xl_worksheet = xl_workbook.sheet_by_name("Info")
    print("Info:",xl_worksheet)
    
    insured_dict={}
    owner_dict={}
    #add "POLICY NUMBER"
    insured_dict["POLICY NUMBER"] = str(int(xl_worksheet.cell_value(0,1)))
    owner_dict["POLICY NUMBER"] = str(int(xl_worksheet.cell_value(0,1)))
    #add others info
    for row_idx in range(3,xl_worksheet.nrows):
        #print( xl_worksheet.cell_value(row_idx,0) ,':', xl_worksheet.cell_value(row_idx,1))
        insured_dict[xl_worksheet.cell_value(row_idx,0)] = str(xl_worksheet.cell_value(row_idx,1))
        owner_dict[xl_worksheet.cell_value(row_idx,0)] = str(xl_worksheet.cell_value(row_idx,2))
        
    #print('insured_dict:',insured_dict)
    # Load json file     
    import json
    json_fname="testTemplate.json"
    with open("testTemplate.json", 'r') as jsonFile:
        data_str=jsonFile.read()

    templates_dict = json.loads(data_str)
  

    #create data to merge
    template= templates_dict['5a.  Sun Life VOT rev (adult).pdf']
    dataList=[]
    # test
    # print('template:',template)
    # print('lenght',len(template))
    # print(type(template))

    #get first value in template
    _firstValue = next(iter(template))
    print('field not exist')
    _generalSetting = PdfDataConfig(
        text=insured_dict[_firstValue['text']],
        font_size=int(_firstValue['fontsize'])
        )
    _configs=[PdfDataConfig(
        location=Location(int(_firstValue['x']),int(_firstValue['y']), int(_firstValue['page']))
    )]
    #test
    # print(int(_firstValue['x']))
    # print(int(_firstValue['y']))
    # print(int(_firstValue['page']))
    _newPdfData=PdfData(general_setting=_generalSetting, configs=_configs.copy())
    dataList.append(_newPdfData)

    # test
    # print(len(dataList))
    # print(dataList[0].general_setting.text)
    # print(dataList[0].configs[0].location.x)
    # print('template 0:',template[0])

    for field_idx in range(1,len(template)):
        field = template[field_idx]
        # test
        print(field)
        print('field:', field['text'])
        print(len(dataList))

        fieldExist = False
        for data in dataList:
            # test
            print(len(dataList))
            print('dictionary value:',insured_dict[field['text']])
            print('data value',data.general_setting.text)

            if( insured_dict[field['text']] != data.general_setting.text):
                continue
            else:
                # field exist
                fieldExist=True
                print('lenght of configs:',len(data.configs)) 
                print('field exist')
                _newLocation= Location(int(field['x']),int(field['y']), page=int(field['page']))
                data.configs.append(PdfDataConfig(location=_newLocation))
                print('lenght of configs:',len(data.configs)) 
                print(len(dataList))
                break
        if fieldExist == False:
            # field not exist   
            print('field not exist')
            print(len(dataList))
            _generalSetting = PdfDataConfig(
                text=insured_dict[field['text']],
                font_size=int(field['fontsize'])
                )
            _configs=[PdfDataConfig(
                location=Location(int(field['x']),int(field['y']), int(field['page']))
            )]
            _newPdfData=PdfData(general_setting=_generalSetting, configs=_configs.copy())
            dataList.append(_newPdfData)
            print(len(dataList))
    print(len(dataList))

    source_pdf = "5a.  Sun Life VOT rev (adult).pdf"

    with open(source_pdf, "rb") as original:
        try:
            result = merge_pdf_with_data(original, dataList)
        except PdfCreationFailed as e:
            print(e)
            exit(1)
        else:
            result.seek(0)
            with open("result.pdf", "wb") as result_pdf:
                result_pdf.write(result.read())
            print("Success!")
            exit(0)