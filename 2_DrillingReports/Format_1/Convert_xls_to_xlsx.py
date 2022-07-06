import os
import win32com.client as win32


def preprocess_files(link):
    """ Переводит файлы в расшишении .xls в .xlsx


    :param link: Ссылка на папку с файлами для обработки
    :return:
    """
    for root, folders, files in os.walk(link):
        for file in files:
            if file.endswith("xls"):
                print("Processing file: {}".format(file))
                file_path = os.path.join(root, file)
                # Переименовываем xls в xlsx
                excel = win32.gencache.EnsureDispatch('Excel.Application')
                wb = excel.Workbooks.Open(file_path)
                wb.SaveAs(file_path + "x", FileFormat=51)
                wb.Close()
                excel.Application.Quit()
                os.remove(file_path)


def main():
    link = r""
    preprocess_files(link)


if __name__ == '__main__':
    main()
