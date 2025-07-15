import pdfplumber
import re

with pdfplumber.open(r"C:\Users\chen-fr\Desktop\临时文件\invoice\Chen, Eunice-486340123-Beechfield Brands Ltd .pdf") as pdf:
# with pdfplumber.open(r"C:\Users\chen-fr\Desktop\临时文件\invoice\4841912632-.pdf") as pdf:
# with pdfplumber.open(r"C:\Users\chen-fr\Desktop\临时文件\invoice\4841912479-LIU CHEN ENTERPRISE(HONG KONG)LIMITED .pdf") as pdf:
# with pdfplumber.open(r"C:\Users\chen-fr\Desktop\临时文件\invoice\Li, Hongyan-4870059338-ALTAI S.R.L.-Martina Sartor.pdf") as pdf:
    page = pdf.pages[0]

    # 提取文本
    page_text = page.extract_text(x_tolerance=2)

    # 将文本按行分割
    page_text_list = page_text.split("\n")

    # 你可以根据需要进一步处理列表
    # 比如过滤掉空行或只保留特定格式的行
    filtered_text_list = [line for line in page_text_list if line.strip()]

    # 打印结果
    text_list = []
    for line in filtered_text_list:
        text_list.append(line)
    print(text_list)