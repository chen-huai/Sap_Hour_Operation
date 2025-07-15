from io import StringIO
import os
import re
import fitz

class PDF_Operate():
    def __init__(self, pdf_file):
        self.pdf_file = pdf_file
        self.text = ""

    def extract_text(self):
        pdf_document = fitz.open(self.pdf_file)

        for page_num in range(pdf_document.page_count):
            page = pdf_document[page_num]
            self.text += page.get_text("text")

        pdf_document.close()

    def get_text(self):
        return self.text


    # def readPdf(inputFile):
    #     text = []
    #     with pdfplumber.open(inputFile) as pdf:
    #         for page in pdf.pages:
    #             page_text = page.extract_text(x_tolerance=2)
    #
    #             page_text_list = page_text.split("\n")
    #
    #             text += page_text_list
    #     return text

    def saveAs(inputFile, outputFile):
        with open(inputFile, 'rb') as fp1:
            b1 = fp1.read()
        with open(outputFile, 'wb') as fp2:
            fp2.write(b1)

if __name__ == '__main__':
    # file = "C:\\Users\\chen-fr\\Desktop\\nb\\Lai, Shanshan-4870045361-宁波市信润贸易有限公司-Jasmine.pdf"  # 文件夹目录
    file = "C:\\Users\\chen-fr\\Desktop\\nb\\SPOOL_782910.pdf"  # 文件夹目录

    pdf_obj = PDF_Operate(file)
    pdf_obj.extract_text()
    print(pdf_obj.get_text())






