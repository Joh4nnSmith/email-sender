from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from getpass import getpass
from os import path
import smtplib

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
        return data_student[:2], data_student[2:]


class Email:
    def __init__(self, transmitter, passwd, information_student):
        self.transmitter = transmitter
        self.password = passwd
        self.receiver = information_student[0]
        self.receiver_name = information_student[1]

    def create_message(self, grades):
        message = MIMEMultipart()
        message['From'] = self.transmitter
        message['To'] = self.receiver
        message['Subject'] = 'Notas [SEMESTRE-PERIODO, etc]'

        html_body = f'''
                    <html>
                        <head> </head>
                        <body>
                        Hola, {self.receiver_name}

                        <p> A continuación presento las notas del [SEMESTRE-PERIODO, etc]. La nota global se <br>
                        divide de la siguiente manera: [NOTA1] (X%) y [NOTA1] (X%). <br> <br>

                        En términos generales, sus notas son: <br>
                        <ul>
                            <li type="disc"><b>NOTA1</b>: {round(grades[1], 1)}</li>
                            <li type="disc"><b>NOTA2</b>: {round(grades[2], 1)}</li>
                        </ul>

                        La nota total del [SEMESTRE-PERIODO, etc] es: {round(grades[0], 1)} <br> <br>

                        <br> <br>
                        </p
                        </body>
                    </html>
                    '''
        message.attach(MIMEText(html_body, 'html'))
        text = message.as_string()
        return text

    def send_mail(self, text):
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()
        server.starttls()
        server.login(self.transmitter, self.password)

        server.sendmail(self.transmitter, self.receiver, text)
        server.quit()


if __name__ == '__main__':
    email_address = input('Enter your email: ')
    passwd = input('Enter your password: ')
    # passwd = getpass('Enter your password: ')
    file_path = input('Enter file path: ')

    if path.isfile(file_path):
        file = ExcelReader(file_path)
        for i in range(1, file.n_rows + 1):
            information_student, grades = file.extract_data(i)
            email = Email(email_address, passwd, information_student)
            email.send_mail(email.create_message(grades))
            print('Email send to ', information_student[1])
    else:
        print("Error, the file does not exists")
