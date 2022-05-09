from typing import Union
import openpyxl
import consts
from pregunta import Pregunta


def fill_categories_from_excel(
    file_path: str,
    categories: dict
) -> None:
    wb = openpyxl.load_workbook(file_path)
    sheet = wb['DATOS']
    for row in sheet.rows:
        pregunta = Pregunta(row[0].value, row[1].value,
                            row[2].value, row[3].value)
        categories[pregunta.categoria].append(pregunta)


def get_value_calc_formula(
    subtipo_row,
    curr_row,
    pregunta,
    i_respuesta
) -> Union[str, int]:
    if pregunta.tipo == consts.TYPE_VEC:
        return '=IF(E{row}="+",{mas},IF(E{row}="-",{menos},0))'.format(
            row=subtipo_row,
            mas=100/pregunta.cant_respuestas*(i_respuesta+1),
            menos=100/pregunta.cant_respuestas *
            (pregunta.cant_respuestas - i_respuesta),
        )
    elif pregunta.tipo == consts.TYPE_VEX:
        return '=IF(E{row}="+",{mas},IF(E{row}="-",{menos},0))'.format(
            row=subtipo_row,
            mas=100/pregunta.cant_respuestas*(i_respuesta+1),
            menos=100/pregunta.cant_respuestas *
            (pregunta.cant_respuestas - i_respuesta),
        )
    elif pregunta.tipo == consts.TYPE_VEM:
        return '=IF(E{row}="+",{mas},IF(E{row}="-",{menos},0))'.format(
            row=subtipo_row,
            mas=100/pregunta.cant_respuestas*(i_respuesta+1),
            menos=100/pregunta.cant_respuestas *
            (pregunta.cant_respuestas - i_respuesta)
        )
    elif pregunta.tipo == consts.TYPE_BOL:
        return ('=IF(OR(AND(E{0}="+", G{1}="Si"),'
                ' AND(E{0}="-", G{1}="No")), 100, 0)').format(
            subtipo_row, curr_row)
    elif pregunta.cant_respuestas > 0:
        val = 100 / pregunta.cant_respuestas
        return '=IF(D{}="MULT",{},IF(D{}="SING",100,0))'.format(
            subtipo_row, val, subtipo_row
        )
    return 0


def set_titles(sheet) -> None:
    for row, col, title in consts.TITLES:
        sheet.cell(row=row, column=col).value = title


def merge_title_cells(sheet) -> None:
    for col in consts.COLUMNS_TO_MERGE:
        sheet.merge_cells(col)


def set_base_data(sheet, row, pregunta):
    data = [(2, pregunta.id), (3, pregunta.pregunta),
            (4, pregunta.tipo), (6, 1)]
    for col, val in data:
        sheet.cell(row=row, column=col).value = val


def merge_cell_when_various_answers(
    sheet,
    next_row,
    cant_rtas
) -> None:
    if cant_rtas > 0:
        last_row = cant_rtas + next_row-1
        for col in ['B', 'C', 'D', 'E', 'F']:
            sheet.merge_cells(
                f"{col}{next_row}:{col}{last_row}")
            sheet.merge_cells(
                f"K{next_row}:k{cant_rtas + next_row-1}")
            # Combinar las celdas de los valores indexados por pregunta
            sheet.merge_cells(
                f"J{next_row}:j{cant_rtas + next_row-1}")


def apply_formatting(sheet) -> None:
    for i, row in enumerate(sheet.rows):
        for j, cell in enumerate(row):
            cell.border = consts.THIN_BORDER
            if i in [0, 1] or j == 2:
                cell.alignment = consts.WRAPPED_ALIGNMENT
            else:
                cell.alignment = consts.ALIGNMENT


def apply_sum_formula(
    sheet,
    row,
    col,
    first_row,
    next_row
) -> None:
    _max = f'MAX(I{first_row}:I{next_row-1})'
    val = f'=IF(D{first_row}="MULT",SUM(I{first_row}:I{next_row-1}),{_max})'
    sheet.cell(row=row, column=col).value = val


def create_new_excel(file_path: str, categories) -> None:
    wb = openpyxl.Workbook()
    for key in categories:
        wb.create_sheet(key)
        sheet = wb[key]
        sheet.freeze_panes = 'A3'
        set_titles(sheet)
        merge_title_cells(sheet)
        next_row = 3
        ult_fila = 3
        for pregunta in categories[key]:
            ult_fila += pregunta.cant_respuestas
        ult_fila -= 1
        for pregunta in categories[key]:
            # Coloca el id, la pregunta y la formula para
            # el subtipo para cada pregunta
            set_base_data(sheet, next_row, pregunta)

            # Paint cell when tipo is "undefined" (consts.UNDEFINED)
            if pregunta.tipo == consts.TYPE_UNDEFINED:
                sheet.cell(row=next_row, column=4).fill = consts.UNDEFINED_FILL

            merge_cell_when_various_answers(
                sheet, next_row, pregunta.cant_respuestas)

            first_row = next_row
            for i, respuesta in enumerate(pregunta.respuestas_as_list):
                sheet.cell(row=next_row, column=1).value = pregunta.categoria
                sheet.cell(row=next_row, column=7).value = respuesta
                value_cell = sheet.cell(row=next_row, column=8)
                # Get autocalculated value for each subtipo
                value = get_value_calc_formula(
                    first_row, next_row, pregunta, i)
                value_cell.value = value
                con_op = f'=H{next_row}*F{first_row}'
                sheet.cell(row=next_row, column=9).value = con_op
                final_val = '=I%s*(100/SUM(J3:J100))' % next_row
                sheet.cell(row=next_row, column=12).value = final_val
                next_row += 1

            apply_sum_formula(sheet, first_row, 10, first_row, next_row)
            to_format = '=IF(D{}="MULT",{},1)'
            sheet.cell(row=first_row, column=11).value = to_format.format(
                first_row, pregunta.cant_respuestas)

        apply_formatting(sheet)

    del wb["Sheet"]
    wb.save(file_path)


def main() -> None:
    categories = {key: [] for key in consts.CATEGORIES}
    fill_categories_from_excel('data.xlsx', categories)
    done = False
    count = 0
    while not done:
        try:
            create_new_excel(f'result-{count}.xlsx', categories)
            done = True
        except PermissionError:
            count += 1
    print("the end")


if __name__ == '__main__':
    main()
