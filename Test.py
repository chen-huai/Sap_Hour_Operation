# class Test():
#     def __init__(self):
#         self.a = 1
#
#     def op1(func,self):
#         def warraper(*args, **kwargs):
#             self.a = 0
#             func(*args, **kwargs)
#         return warraper
#
#     def op2(func,self):
#         def warraper(*args, **kwargs):
#             self.a = 1
#             func(*args, **kwargs)
#         return warraper
#
#     @op1
#     def op3(self):
#         self.a += 1
#         print(3, self.a)
#
#     @op2
#     def op4(self):
#         self.a += 1
#         print(4, self.a)
#
#     @op1
#     def op5(self):
#         self.a += 1
#         print(5, self.a)
#
#     @op2
#     def op6(self):
#         self.a += 1
#         print(6, self.a)
#
#     def op7(self):
#         print(7, self.a)
#
#
# test = Test()
# test.op3()
# test.op7()
# test.op4()
# test.op7()
# test.op5()
# test.op7()
# test.op6()
# test.op7()


import time


# def baiyu():
#     t1 = time.time()
#     print("我是攻城狮白玉")
#     time.sleep(2)
#     print("执行时间为：", time.time() - t1)
#
#
# def blog(name):
#     t1 = time.time()
#     print('进入blog函数')
#     name()
#     print('我的博客是 https://blog.csdn.net/zhh763984017')
#     print("执行时间为：", time.time() - t1)
#
#
# if __name__ == '__main__':
#     func = baiyu  # 这里是把baiyu这个函数名赋值给变量func
#     func()  # 执行func函数
#     print('------------')
#     blog(baiyu)  # 把baiyu这个函数作为参数传递给blog函数
#


# def count_time(func):
#     def wrapper():
#         t1 = time.time()
#         func()
#         print("执行时间为：", time.time() - t1)
#
#     return wrapper
#
# def baiyu():
#     print("我是攻城狮白玉")
#     time.sleep(2)
#
# if __name__ == '__main__':
#     baiyu = count_time(baiyu)  # 因为装饰器 count_time(baiyu) 返回的时函数对象 wrapper，这条语句相当于  baiyu = wrapper
#     baiyu()  # 执行baiyu()就相当于执行wrapper()


# import time
#
#
# def count_time(func):
#     def wrapper():
#         t1 = time.time()
#         func()
#         print("执行时间为：", time.time() - t1)
#
#     return wrapper
#
#
# @count_time
# def baiyu():
#     print("我是攻城狮白玉")
#     time.sleep(2)
#
#
# if __name__ == '__main__':
#     # baiyu = count_time(baiyu)  # 因为装饰器 count_time(baiyu) 返回的时函数对象 wrapper，这条语句相当于  baiyu = wrapper
#     # baiyu()  # 执行baiyu()就相当于执行wrapper()
#
#     baiyu()  # 用语法糖之后，就可以直接调用该函数了


# # try测试是否可以包含没有错误的变量
# class test():
#     def __init__(self):
#         self.a = ''
#         self.b = ''
#
#     def test1(self):
#         try:
#             self.a = 1
#             self.b = 2
#             print(self.a)
#             self.b = 'a' + '%s' % self.a
#             print(self.b)
#         except:
#             self.a
#             print(self.a)
#
#
# test = test()
# test.test1()


# import random
#
# # 生成6个不重复的红球号码
# red_balls = []
# while len(red_balls) < 6:
#     ball = random.randint(1, 33)
#     if ball not in red_balls:
#         red_balls.append(ball)
#
# # 生成1个蓝球号码
# blue_ball = random.randint(1, 16)
#
# # 打印选号结果
# print("红球号码：", sorted(red_balls))
# print("蓝球号码：", blue_ball)

# class PdfNameUpdater:
#     def __init__(self):
#         self.pdf_names = {'invoice': '', 'fapiao': ''}
#
#     def get_pdf_name(self, flag):
#         return self.pdf_names.get(flag, '')
#
#     def update_pdf_name(self, msg, flag):
#         if flag in self.pdf_names:
#             pdf_name = self.pdf_names[flag]
#             pdf_name_list = pdf_name.split(' + ')
#
#             if msg in pdf_name_list:
#                 pdf_name_list.remove(msg)
#             else:
#                 pdf_name_list.append(msg)
#
#             changed_pdf_name = ' + '.join(pdf_name_list)
#             self.pdf_names[flag] = changed_pdf_name
#             return changed_pdf_name
#         else:
#             return ''
#
# # 使用示例
# pdf_updater = PdfNameUpdater()
# invoice_name = pdf_updater.get_pdf_name('invoice')
# print("Invoice Name:", invoice_name)
#
# pdf_updater.update_pdf_name('Invoice1', 'invoice')
# pdf_updater.update_pdf_name('Invoice2', 'invoice')
# invoice_name = pdf_updater.get_pdf_name('invoice')
# print("Updated Invoice Name:", invoice_name)



# # 数据读取问题
#
# from Get_Data import *
# from Test_Data import *
# import pandas as pd
# import os
# os.chdir('C:\\Users\\chen-fr\\Desktop\\临时文件\\sap')
# #读入数据
# #读入数据
# rawData = pd.read_excel('20230331.xlsx')
#
# # rawData['合并'] = rawData['Project No.'] + '\t' + rawData['Currency']
# # test = Test()
# # combineProject = rawData.groupby(["Invoices' name (Chinese)",'CS', 'Sales', 'Currency', 'Material Code', 'Buyer(GPC)', 'Month']).apply(test.concat_func).reset_index()
#
# rawData['row_msg'] = rawData['Project No.'] + '\t' + rawData['Currency']
# test = Get_Data()
# combineProject = rawData.groupby(["Invoices' name (Chinese)",'CS', 'Sales', 'Currency', 'Material Code', 'Buyer(GPC)', 'Month']).apply(test.row_concat_func).reset_index()
#
# result = pd.merge(rawData, combineProject, on=["Invoices' name (Chinese)",'CS', 'Sales', 'Currency', 'Material Code', 'Buyer(GPC)', 'Month'],how='right')
# print(result)


import PyPDF2

class PDF_Operate():
    def __init__(self, pdf_file):
        self.pdf_file = pdf_file
        self.text = ""

    def extract_text(self):
        with open(self.pdf_file, "rb") as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                self.text += page.extract_text()

    def get_text(self):
        return self.text

# 创建PDF_Operate类的实例
pdf_operator = PDF_Operate("C:\\Users\\chen-fr\\Desktop\\nb\\SPOOL_782910.pdf")

# 提取PDF文件中的文本内容
pdf_operator.extract_text()

# 获取提取的文本内容
extracted_text = pdf_operator.get_text()

# 打印提取的文本内容
print(extracted_text)