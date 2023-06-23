import PyPDF2
import cv2
import numpy as np
import openpyxl
from pyzbar import pyzbar


def invoice_decode_qrcode(imagedata):  # 识别图片中的发票二维码
    # 解码图像数据
    image = cv2.imdecode(np.frombuffer(imagedata, np.uint8), cv2.IMREAD_COLOR)  # 将传入的字节数据转换成字节数组
    imagegray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)  # 将图像转换成灰色图像
    _, binary = cv2.threshold(imagegray, 205, 255, cv2.THRESH_BINARY)  # 图像应用阈值进行二值化 此函数返回两个参数 前一个参数舍弃
    decodedata = pyzbar.decode(binary)  # 识别二维码 成功返回数据 失败为空 列表形式
    # print(decodedata)
    if not decodedata:
        # print("空数据")
        return ""
    else:
        for result in decodedata:  # 单张图片存在多个发票的情况，将多个列表循环遍历出来
            text_list = result.data.decode('utf-8').split(',')
            return [text_list[2], text_list[3], text_list[4], text_list[5], text_list[6]]


def pdf_invoice(pdf_path):
    reader = PyPDF2.PdfReader(pdf_path)  # 创建 PDF 阅读器对象
    workbook = openpyxl.Workbook()  # 创建一个新的 Excel 工作簿
    sheet = workbook.active  # 获取默认的工作表
    # 插入表头
    headers = ['序号', '发票代码', '发票号码', '不含税金额', '开票日期', '校验码']
    sheet.append(headers)

    row = 2
    for page_num in range(len(reader.pages)):  # 遍历 PDF 页面
        page = reader.pages[page_num]  # 获取当前页面
        for image_object in page.images:  # 单张页面存在多张图像的情况 遍历出来
            data = invoice_decode_qrcode(image_object.data)
            data.insert(0, row - 1)  # 插入序号
            # 在指定行上插入数据
            for col_num, value in enumerate(data, 1):
                col_letter = sheet.cell(row=row, column=col_num).column_letter
                sheet[f'{col_letter}{row}'] = value
            row += 1

            # print(data)

    # 保存 Excel 文件
    workbook.save('invoice_info.xlsx')


if __name__ == '__main__':
    # with open("img_1.png", "rb") as image_file:
    #     image_data = image_file.read()
    #     # image_data = io.BytesIO(image_data)
    #     invoice_decode_qrcode(image_data)

    pdf_invoice("test.pdf")
