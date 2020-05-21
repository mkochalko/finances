import PyPDF2
import pdftotext
import io
import xlsxwriter

from PyPDF2 import PdfFileReader

def extract_information(pdf_path):
    with open(pdf_path, 'rb') as f:
        file = f.read()
        memory_file = io.BytesIO(file)
        pdf = pdftotext.PDF(memory_file)
        for page in pdf:
            lines = page.splitlines()
            for line in lines:
                if line != '':
                    print(line)
                    # print('new Line')
            
def write_to_spreadsheet(file):
    workbook = xlsxwriter.Workbook(file)
    worksheet = workbook.add_worksheet()
    # worksheet.write('C5', 'test')
    worksheet.write('C7', 'it works')
    workbook.close()

if __name__ == '__main__':
    path = 'march.pdf'
    spreadsheet = 'overview_test.xlsx'
    write_to_spreadsheet(spreadsheet)
    extract_information(path)