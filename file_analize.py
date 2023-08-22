import openpyxl as xl
from openpyxl.utils import get_column_letter
from copy import copy


def sheets(path) -> list:
    """
    Возвращает названия, кол-во строк и столбцов в листах
    из excel файла, находящегося по переданному пути

    :return: Список с названиями листови количеством строк в таблице 
    :rtype: list
    """

    res = []
    wb = xl.load_workbook(path)
    for sheet in wb.sheetnames:
        res.append([sheet, wb[sheet].max_column, wb[sheet].max_row])
    return res

def create_error_table_header(error_worksheet, main_sheet) -> None:
    """
    Создает шапку таблицы для ошибок

    :return: None
    """

    for merge in list(error_worksheet.merged_cells):
        error_worksheet.unmerge_cells(range_string=str(merge))
    
    error_worksheet["A1"].value = "Ненайденные позиции из ЛСР"
    error_worksheet["A3"].value = "№ п/п"
    error_worksheet["B3"].value = "№ в ЛСР"
    error_worksheet["C3"].value = "Наименование работ"
    error_worksheet["D3"].value = "Ед.изм."
    error_worksheet["E3"].value = "Кол-во"
    error_worksheet["F3"].value = "Тип ошибки"

    for i, row in enumerate(error_worksheet["A1:F4"]):
        for j, cell in enumerate(row):
            if i==2: error_worksheet.cell(row=i+2, column=j+1).value = j+1
            error_worksheet.cell(row=i+1, column=j+1).font = copy(main_sheet["A1"].font)
            error_worksheet.cell(row=i+1, column=j+1).border = copy(main_sheet["A1"].border)
            error_worksheet.cell(row=i+1, column=j+1).fill = copy(main_sheet["A1"].fill)
            error_worksheet.cell(row=i+1, column=j+1).alignment = copy(main_sheet["A1"].alignment)

    for i, row in enumerate(main_sheet["A1:AV4"]):
            for j, cell in enumerate(row):
                if i==3: error_worksheet.cell(row=i+1, column=j+7, value=int(cell.value)+6)
                else: error_worksheet.cell(row=i+1, column=j+7, value=cell.value)
                error_worksheet.cell(row=i+1, column=j+7).font = copy(cell.font)
                error_worksheet.cell(row=i+1, column=j+7).border = copy(cell.border)
                error_worksheet.cell(row=i+1, column=j+7).fill = copy(cell.fill)
                error_worksheet.cell(row=i+1, column=j+7).number_format = copy(cell.number_format)
                error_worksheet.cell(row=i+1, column=j+7).protection = copy(cell.protection)
                error_worksheet.cell(row=i+1, column=j+7).alignment = copy(cell.alignment)
    
    for elem in main_sheet.merged_cells:
            elem=str(elem).split(':')
            elem[0] = list(xl.utils.cell.coordinate_to_tuple(elem[0]))
            elem[1] = list(xl.utils.cell.coordinate_to_tuple(elem[1]))
            error_worksheet.merge_cells(start_row=elem[0][0],
                                  start_column=elem[0][1]+6,
                                  end_row=elem[1][0],
                                  end_column=elem[1][1]+6)  
            
    error_worksheet.merge_cells("A1:F2")

    for i, width in enumerate([5, 40, 50, 13, 7, 30]): error_worksheet.column_dimensions[get_column_letter(i+1)].width = width
    for column in range(xl.utils.cell.column_index_from_string(xl.utils.cell.coordinate_from_string("A1")[0]), xl.utils.cell.column_index_from_string(xl.utils.cell.coordinate_from_string("AV4")[0])+1):
        error_worksheet.column_dimensions[get_column_letter(column+6)].width = int(main_sheet.column_dimensions[get_column_letter(column)].width)
    for i, height in enumerate([18, 36, 16]): error_worksheet.row_dimensions[i+1].height = height

    return None

def add_error_name_header(error_worksheet, main_sheet, sheet, row_num) -> None:
    """
    Создает заголовок для нового листа

    :return: None
    """
    error_worksheet[f"A{row_num+5}"].value = sheet
    error_worksheet[f"A{row_num+5}"].font = copy(main_sheet["A1"].font)
    error_worksheet[f"A{row_num+5}"].border = copy(main_sheet["A1"].border)
    error_worksheet[f"A{row_num+5}"].fill = xl.styles.fills.PatternFill(patternType='solid', fgColor=xl.styles.colors.Color(rgb='bee5b3'))
    error_worksheet[f"A{row_num+5}"].alignment = copy(main_sheet["A1"].alignment)
    error_worksheet.merge_cells(f"A{row_num+5}:F{row_num+5}")

def add_error_row(error_worksheet, main_sheet, row_num: int, data, error_type: int, paste_data=None) -> None:
    """
    Создает новую строку с записью об ошибке

    :return: None
    """
    error_encoding = {
        101: ['Ресурс не найден', 'ff6c70'],
        102: ['Отсутствует цена', 'ffe994'],
        103: ['Не совпадает ед. изм.', 'b9d7f0'], 
    }
    for i, cell in enumerate(data):
        if i!=0:
            error_worksheet.cell(row=row_num+5, column=i+1).value = copy(cell.value)
            error_worksheet.cell(row=row_num+5, column=i+1).font = copy(cell.font)
            error_worksheet.cell(row=row_num+5, column=i+1).border = copy(cell.border)
            error_worksheet.cell(row=row_num+5, column=i+1).fill = copy(cell.fill)
            error_worksheet.cell(row=row_num+5, column=i+1).alignment = copy(cell.alignment)
    
    error_worksheet.cell(row=row_num+5, column=6).value = f'Код ошибки: {error_type} : {error_encoding[error_type][0]}'
    error_worksheet.cell(row=row_num+5, column=6).font = copy(main_sheet["A1"].font)
    error_worksheet.cell(row=row_num+5, column=6).border = copy(main_sheet["A1"].border)
    error_worksheet.cell(row=row_num+5, column=6).fill = xl.styles.fills.PatternFill(patternType='solid', fgColor=xl.styles.colors.Color(rgb=error_encoding[error_type][1]))
    error_worksheet.cell(row=row_num+5, column=6).alignment = copy(main_sheet["A1"].alignment)

    if paste_data:
        for i, cell in enumerate(paste_data):
            error_worksheet.cell(row=row_num+5, column=i+7).value = copy(cell.value)
            error_worksheet.cell(row=row_num+5, column=i+7).font = copy(cell.font)
            error_worksheet.cell(row=row_num+5, column=i+7).border = copy(cell.border)
            error_worksheet.cell(row=row_num+5, column=i+7).fill = copy(cell.fill)
            error_worksheet.cell(row=row_num+5, column=i+7).alignment = copy(cell.alignment)

    return None

def insert_info(path, sheets) -> None:
    """
    Заолнение выбранных файлов найденными материалами

    :return: None
    """
    file = xl.load_workbook(path)
    main_sheet = file[sheets[0][0]]
    
    if sheets[2]==[]: error_worksheet=file.create_sheet("ОШИБКИ")
    else: error_worksheet=file[sheets[2][0]]
    create_error_table_header(error_worksheet, main_sheet)
    
    pasted_error_rows = 0
    for sheet in sheets[1]:
        add_error_name_header(error_worksheet, main_sheet, sheet, pasted_error_rows)
        pasted_error_rows+=1

        worksheet = file[sheet]
        for row in worksheet[f"A1:A{worksheet.max_row}"]:
            for cell in row:
                if cell.value=="№ п/п":
                    start_coordinate = [cell.row-2, xl.utils.cell.column_index_from_string("X")]
        for i, row in enumerate(main_sheet["B1:AV4"]):
            for j, cell in enumerate(row):
                if i==3: worksheet.cell(row=i+start_coordinate[0], column=j+start_coordinate[1], value=int(cell.value)+7)
                else: worksheet.cell(row=i+start_coordinate[0], column=j+start_coordinate[1], value=cell.value)
                worksheet.cell(row=i+start_coordinate[0], column=j+start_coordinate[1]).font = copy(cell.font)
                worksheet.cell(row=i+start_coordinate[0], column=j+start_coordinate[1]).border = copy(cell.border)
                worksheet.cell(row=i+start_coordinate[0], column=j+start_coordinate[1]).fill = copy(cell.fill)
                worksheet.cell(row=i+start_coordinate[0], column=j+start_coordinate[1]).number_format = copy(cell.number_format)
                worksheet.cell(row=i+start_coordinate[0], column=j+start_coordinate[1]).protection = copy(cell.protection)
                worksheet.cell(row=i+start_coordinate[0], column=j+start_coordinate[1]).alignment = copy(cell.alignment)
        for elem in main_sheet.merged_cells:
            elem=str(elem).split(':')
            elem[0] = list(xl.utils.cell.coordinate_to_tuple(elem[0]))
            elem[1] = list(xl.utils.cell.coordinate_to_tuple(elem[1]))
            worksheet.merge_cells(start_row=elem[0][0]+start_coordinate[0]-1,
                                  start_column=elem[0][1]+start_coordinate[1]-2,
                                  end_row=elem[1][0]+start_coordinate[0]-1,
                                  end_column=elem[1][1]+start_coordinate[1]-2)   
        for column in range(xl.utils.cell.column_index_from_string(xl.utils.cell.coordinate_from_string("B1")[0]), xl.utils.cell.column_index_from_string(xl.utils.cell.coordinate_from_string("AV4")[0])+1):
            worksheet.column_dimensions[get_column_letter(column+start_coordinate[1]-2)].width = int(main_sheet.column_dimensions[get_column_letter(column)].width)
        for row in range(xl.utils.cell.coordinate_from_string("B1")[1], xl.utils.cell.coordinate_from_string("AV4")[1]):
            worksheet.row_dimensions[row+start_coordinate[0]-1].height = int(main_sheet.row_dimensions[column].height)

        for row in worksheet[f"C5:C{worksheet.max_row}"]:
            parse_coordinate=0
            for cell in row:
                try:
                    if len(str(cell.value))>=3:
                        parse_coordinate=cell.row
                except Exception as e: pass
            if parse_coordinate: break


        for a, row in enumerate(worksheet[f"A{parse_coordinate}:E{worksheet.max_row}"]):
            bufer_data = None
            error_code = 101
            if row[2].value!=None and row[3].value!=None:
                for parse_sheet in sheets[0]:
                    parse_sheet = file[parse_sheet]
                    
                    for b, parse_row in enumerate(parse_sheet[f"D5:AV{parse_sheet.max_row}"]):
                        # если информация полностью совпадает и все цены есть
                        if str(row[2].value)==str(parse_row[0].value) and str(row[3].value)==str(parse_row[1].value) and str(parse_row[11].value)!=None and str(parse_row[25].value)!=None and str(parse_row[39].value)!=None:
                            for copy_row in parse_sheet[f"B{b+5}:AV{b+5}"]:
                                for i, copy_cell in enumerate(copy_row):
                                    worksheet.cell(row=a+parse_coordinate, column=i+start_coordinate[1]).value = copy(copy_cell.value)
                                    worksheet.cell(row=a+parse_coordinate, column=i+start_coordinate[1]).font = copy(copy_cell.font)
                                    worksheet.cell(row=a+parse_coordinate, column=i+start_coordinate[1]).border = copy(copy_cell.border)
                                    worksheet.cell(row=a+parse_coordinate, column=i+start_coordinate[1]).fill = copy(copy_cell.fill)
                                    worksheet.cell(row=a+parse_coordinate, column=i+start_coordinate[1]).number_format = copy(copy_cell.number_format)
                                    worksheet.cell(row=a+parse_coordinate, column=i+start_coordinate[1]).protection = copy(copy_cell.protection)
                                    worksheet.cell(row=a+parse_coordinate, column=i+start_coordinate[1]).alignment = copy(copy_cell.alignment)
                            bufer_data = None
                            parse_state = True
                            break
                        # если нет цены 
                        if str(row[2].value)==str(parse_row[0].value) and str(row[3].value)==str(parse_row[1].value) and (str(parse_row[11].value)==None or str(parse_row[25].value)==None or str(parse_row[39].value)==None):
                            bufer_data = parse_sheet[f"A{b+5}:AV{b+5}"][0]
                            error_code = 102
                            parse_state=False
                        #если не совпадает ед. изм.
                        if str(row[2].value)==str(parse_row[0].value) and str(row[3].value)!=str(parse_row[1].value):
                            bufer_data = parse_sheet[f"A{b+5}:AV{b+5}"][0]
                            error_code = 103
                            parse_state=False
                        else:
                            parse_state=False
                    if parse_state: break
                
                if not parse_state:
                    add_error_row(error_worksheet, main_sheet, pasted_error_rows, row, error_code, bufer_data)
                    pasted_error_rows+=1
    
    file.save(path[:-5]+" Обработанный "+path[-5:])
    return None