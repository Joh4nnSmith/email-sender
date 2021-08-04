# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# from getpass import getpass
from os import path
# import smtplib

# from decouple import config
from openpyxl import load_workbook


class ExcelReader:
    def __init__(self, path_file):
        self.path_file = path_file
        self.workbook = self.load_file()
        self.n_column = self.count_columns()
        self.n_rows = self.count_rows()

    def load_file(self):
        excel_file = load_workbook(self.path_file)
        return load_workbook(self.path_file).active

    def count_rows(self):
        for nrow, cell in enumerate(self.workbook['A']):
            if cell.value is None:
                return nrow
                break

    def count_columns(self):
        return self.workbook.max_column

    def extract_data(self, index):
        data_student = []
        for row in self.workbook[index]:
            data_student.append(row.value)
        print(data_student[:2], data_student[2:])


if __name__ == '__main__':
    print('Iniciando ejecución 3')
    # pss = getpass()
    # print(pss)
    file_path = r'C:\Users\ASUS\Desktop\datos.xlsx'

    if path.isfile(file_path):
        file = ExcelReader(file_path)
        print(file.n_column)
        print(file.n_rows)
        for i in range(1, file.n_rows + 1):
            print(file.extract_data(i))
    else:
        print("Error, the file does not exists")
