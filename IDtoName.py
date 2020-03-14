# -*- coding: utf-8 -*-

import time
import re
from selenium import webdriver
from chromedriver_py import binary_path
from xlrd import open_workbook
from xlutils.copy import copy


excel_file_path = "examples.xlsx"
sheet_name = "Sheet1"
read_col_num = 0
write_col_num = 1
web_url = "http://plants.ensembl.org/Zea_mays/Info/Index"

browser = webdriver.Chrome(executable_path=binary_path)


def read_gene_id_from_excel(excel_file_path, sheet_name, col_num):
    """
    从excel中提取ID list
    """
    gene_id_list = []
    with open_workbook(excel_file_path) as workbook:
        worksheet = workbook.sheet_by_name(sheet_name)
        rows_num = worksheet.nrows
        print(rows_num)
        data_text = worksheet.col(col_num)
        #print(rows_num, data_text)
        for row_num in range(1, rows_num):
            row_data_str = str(data_text[row_num])
            #print(row_data_str)
            gene_id = re.findall(r"Zm\w+", row_data_str)
            #print(gene_id, type(gene_id), row_num, type(row_num))
            gene_id.append(row_num)
            #print(row_num, gene_id)
            gene_id_list.append(gene_id)
        print("gene id列表：", gene_id_list)
    return gene_id_list


def login_web(web_url):
    browser.maximize_window()
    print("登录web：", web_url)
    browser.get(web_url)
    # 等待登录页面的elements加载完毕
    browser.implicitly_wait(3)


def search_gene_id(gene_id):

    # input and search gene id
    try:
        browser.find_element_by_id("q").send_keys(gene_id)
    except Exception:
        print("Search gene id fail...")
    else:
        # wait a second to jump to next page
        time.sleep(1)
        browser.find_element_by_class_name("fbutton").click()

    current_url = browser.current_url
    print("跳转到下一登录页面：", current_url)

    try:
        browser.implicitly_wait(2)
        # get gene description
        desc_text = browser.find_element_by_class_name("rhs").get_attribute("textContent")
    except Exception:
        print("Get Description fail...")
    else:
        #print("Description：", desc_text)
        return desc_text


def write_desc_to_excel(excel_file_path, desc_text,
                        input_row_num, input_col_num=write_col_num):
    rb = open_workbook(excel_file_path)
    print(rb)
    wb = copy(rb)
    wb.get_sheet(sheet_name).write(input_row_num, input_col_num, desc_text)
    wb.save(excel_file_path)


def gene_id_to_name(excel_file_path, sheet_name, read_col_num, write_col_num,
                    web_url):
    """
    读取数据，搜索ID，写入数据
    """
    gene_id_list = read_gene_id_from_excel(excel_file_path=excel_file_path,
                                           sheet_name=sheet_name, col_num=read_col_num)

    for gene_id_info in gene_id_list:
        #print(gene_id_info, type(gene_id_info))
        # 确保子列表由[gene_id, row_num]组成
        if len(gene_id_info) == 2:
            gene_id = gene_id_info[0]
            row_num = gene_id_info[1]
            #print("gene id：", gene_id, "\n", "row num：", row_num)
            # 登录网页
            login_web(web_url=web_url)
            # 获取descript信息
            desc_text = search_gene_id(gene_id=gene_id)
            # 信息写入excel
            print("将gene id:", gene_id, "的description信息：", desc_text, "\n",
                  "写入excel表的：", "行：", row_num, "列：", write_col_num)
            write_desc_to_excel(excel_file_path=excel_file_path, desc_text=desc_text,
                                input_row_num=row_num, input_col_num=write_col_num)
        else:
            print("可能没有匹配到gene id...")

    browser.quit()
    print("测试完成，退出浏览器")


if __name__ == "__main__":
    gene_id_to_name(excel_file_path=excel_file_path, sheet_name=sheet_name,
                    read_col_num=read_col_num, write_col_num=write_col_num,
                    web_url=web_url)
