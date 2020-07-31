#!/usr/bin/python
# -*- coding: UTF-8 -*-

import os
import sys

import MySQLdb
import xlrd
import xlutils
import xlwt
from xlwt import Workbook

from ReadConfig import ReadConfig
from TableModel import FieldModel
# sys.path.append(os.path.dirname(os.path.abspath('.')+"/*"))
from TableModel import TableModel

readConfig = ReadConfig()
host = readConfig.get_param("host")
user_name = readConfig.get_param("user_name")
password = readConfig.get_param("password")
db_name = readConfig.get_param("db_name")
charset = readConfig.get_param("charset")


def get_table_list(db_name):
    """ 获取数据库有哪些表名  """
    sql = """ select table_name from information_schema.tables where table_schema = '{}' """.format(db_name)
    # 打开数据库连接
    db = MySQLdb.connect(host, user_name, password, db_name, charset='utf8')
    # 使用cursor()方法获取操作游标
    cursor = db.cursor()
    try:
        # 执行SQL语句
        cursor.execute(sql)
        # 获取所有记录列表
        results = cursor.fetchall()
        table_list = []
        for row in results:
            table_name = row[0]
            table_list.append(table_name)
    except:
        print("Error: unable to fecth data")
    return table_list


def get_single_table_model(db_name, table_name):
    """
    获取单个表的字段模型

    :param db_name:
    :param table_name:
    :return:
    """
    # SQL 查询语句
    sql = """ select
               column_name 列名,
               column_type 数据类型,
               data_type 字段类型,
               character_maximum_length 长度,
               is_nullable 是否为空,
               column_default 默认值,
               column_comment 备注 
           from information_schema.columns  
           where table_schema = '{}'  
           and table_name = '{}' """.format(db_name, table_name)

    # 打开数据库连接
    db = MySQLdb.connect(host, user_name, password, db_name, charset='utf8')
    # 使用cursor()方法获取操作游标
    cursor = db.cursor()
    try:
        # 执行SQL语句
        cursor.execute(sql)
        # 获取所有记录列表
        results = cursor.fetchall()
        field_modle_list = []
        for row in results:
            column_name = row[0]
            column_type = row[1]
            data_type = row[2]
            character_maximum_length = row[3]
            is_nullable = row[4]
            column_default = row[5]
            column_comment = row[6]
            fieldModel = FieldModel(column_name, column_type, data_type, character_maximum_length,
                                    is_nullable, column_default, column_comment)
            field_modle_list.append(fieldModel)
        single_table_model = TableModel(db_name, table_name, field_modle_list)
    except:
        print("Error: unable to fecth data")
    # 关闭数据库连接
    db.close()
    return single_table_model


def get_all_table_model():
    """
    获取库里所有表的字段信息
    :return:
    """
    all_table_model_list = []
    table_name_list = get_table_list(db_name)
    for table_name in table_name_list:
        single_table_model = get_single_table_model(db_name, table_name)
        all_table_model_list.append(single_table_model)
    return all_table_model_list


def get_cell_type(colour_type_num):
    """
    根据 colour_type_num=3 is  Green ,colour_type_num=5 is  yellow
    :param colour_type_num:
    :return:
    """
    pattern = xlwt.Pattern()  # Create the Pattern
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
    pattern.pattern_fore_colour = colour_type_num  # May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green, 4 = Blue, 5 = Yellow, 6 = Magenta, 7 = Cyan, 16 = Maroon, 17 = Dark Green, 18 = Dark Blue, 19 = Dark Yellow , almost brown), 20 = Dark Magenta, 21 = Teal, 22 = Light Gray, 23 = Dark Gray, the list goes on...
    style = xlwt.XFStyle()  # Create the Pattern
    style.pattern = pattern  # Add Pattern to Style
    return style


def generate_header_sheet(sheet, single_table_model, table_cell_type_green, filed_cell_type_yellow):
    yellow_cell_type = get_cell_type(filed_cell_type_yellow)
    green_cell_type = get_cell_type(table_cell_type_green)
    sheet.write(0, 0, "表名", green_cell_type)
    sheet.write(0, 1, single_table_model.table_name)
    sheet.write(1, 0, "列名", yellow_cell_type)
    sheet.write(1, 1, "数据类型", yellow_cell_type)
    sheet.write(1, 2, "字段类型", yellow_cell_type)
    sheet.write(1, 3, "长度", yellow_cell_type)
    sheet.write(1, 4, "是否为空", yellow_cell_type)
    sheet.write(1, 5, "默认值", yellow_cell_type)
    sheet.write(1, 6, "备注", yellow_cell_type)
    return sheet


def generate_list_link_sheet(workbook):
    list_link_sheet = workbook.add_sheet('目录链接')
    list_link_sheet.write(0, 0, "表目录", get_cell_type(3))
    i = 1
    for table_name in get_table_list(db_name):
        link = 'HYPERLINK("#{0}!A1","{0}"))'.format(table_name)
        list_link_sheet.write(i, 0, xlwt.Formula(link))
        i += 1
    return list_link_sheet


def generate_workbook():
    yellow_type = 5
    green_type = 3
    workbook = Workbook()
    sheets = []
    for single_table_model in get_all_table_model():
        sheet = workbook.add_sheet(single_table_model.table_name,
                                   cell_overwrite_ok=True)  # 创建第一个sheet页 第二参数用于确认同一个cell单元是否可以重设值
        sheet = generate_header_sheet(sheet, single_table_model, green_type, yellow_type)
        field_model_list = single_table_model.field_model_list
        i = 2
        for field_model in field_model_list:
            sheet.write(i, 0, field_model.column_name)
            sheet.write(i, 1, field_model.column_type)
            sheet.write(i, 2, field_model.data_type)
            sheet.write(i, 3, field_model.character_maximum_length)
            sheet.write(i, 4, field_model.is_nullable)
            sheet.write(i, 5, field_model.column_default)
            sheet.write(i, 6, field_model.column_comment)
            i += 1
        sheets.append(sheet)

    list_link_sheet = generate_list_link_sheet(workbook)
    sheets.insert(0, list_link_sheet)
    workbook._Workbook__worksheets = sheets

    return workbook


def generate_excel(data_write, workbook):
    # 保持
    if os.path.exists(data_write):
        print('hello delete')
        os.remove(data_write)
    workbook.save(data_write)  # 保存新的excel


if __name__ == "__main__":

    root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    print(root_dir)
    data_write = os.path.join(root_dir, "data/TableModel.xls")
    workbook = generate_workbook()
    generate_excel(data_write, workbook)
