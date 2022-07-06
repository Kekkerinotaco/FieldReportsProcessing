import re


def manage_cells(worksheet, cell_name, row_number):
    """ При помощи регулярных выражений определяет, какой функцией нужно обработать ячейку, и производит обработку


    :param worksheet: Книга Excel, с которой берется ячейка
    :param cell_name: Имя ячейки
    :param row_number: Номер строки
    :return: Необходимое значение
    """
    if re.match(r"^N[\d]+", cell_name):
        value = process_N(worksheet, row_number)
    elif re.match(r"^X[\d]+", cell_name):
        value = process_X(worksheet, row_number)
    elif re.match(r"^AF[\d]+", cell_name):
        value = process_AF(worksheet, row_number)
    else:
        value = worksheet[cell_name].value
    return value


def process_N(worksheet, row_number):
    value = worksheet["N{}".format(row_number + 2)].value
    return value


def process_X(worksheet, row_number):
    value = worksheet["X{}".format(row_number + 3)].value
    return value


def process_AF(worksheet, row_number):
    value = worksheet["AF{}".format(row_number + 3)].value
    return value
