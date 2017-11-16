# coding=utf-8
import traceback

import mysql.connector
import xlwt

import setting
import json


def execute_query(sql, args):
    host, port, user, password, database = get_mysql_props()
    conn = mysql.connector.connect(host=host, port=port, user=user, password=password, database=database)
    cursor = conn.cursor()
    cursor.execute(sql, args)
    values = cursor.fetchall()
    cursor.close()
    conn.close()
    return values


def get_mysql_props():
    mysql_props = setting.MYSQL_CONFIG
    return mysql_props['host'], mysql_props['port'], \
           mysql_props['user'], mysql_props['password'], mysql_props['database']


def get_role_urls(role_id):
    args = [role_id]
    sql = 'SELECT u.* FROM bdf2_role_resource rr INNER JOIN bdf2_url u ON rr.URL_ID_=u.ID_ WHERE ROLE_ID_=%s ORDER BY u.PARENT_ID_ ASC'
    role_urls = execute_query(sql, args)
    return role_urls


def get_parent_url(url_id):
    sql = 'SELECT * FROM bdf2_url WHERE ID_=%s'
    return execute_query(sql, [url_id])


class Url:
    def __init__(self, id, name, nodes=None):
        self.id = id
        self.name = name
        self.nodes = nodes

    def get_nodes(self):
        return self.nodes

    def set_nodes(self, nodes):
        self.nodes = nodes


def view_url_tree(url_resources):
    tree = load_url_tree(None, url_resources)
    return tree


def load_url_tree(url_id, urls):
    url_list = get_urlList_by_urlId(url_id, urls)
    for url in url_list:
        url.set_nodes(load_url_tree(url.id, urls))
    return url_list


def get_urlList_by_urlId(url_id, urls):
    global u
    url_list = []
    for url in urls:
        id = url[0]
        parent_id = url[7]
        name = url[5]
        if url_id is None and parent_id is None:
            u = Url(id.encode("utf-8"), name.encode("utf-8"))
            url_list.append(u)
        elif url_id is not None and parent_id is not None and url_id.encode("utf-8") == parent_id.encode("utf-8"):
            u = Url(id.encode("utf-8"), name.encode("utf-8"))
            url_list.append(u)
    return url_list


COL_INDEX = 0
ROW_INDEX = 0


def write_url_to_excel(urls, ws):
    global ROW_INDEX, COL_INDEX
    for i, url in enumerate(urls):
        ws.write(ROW_INDEX, COL_INDEX, url.name)
        if len(url.get_nodes()) != 0:
            COL_INDEX += 1
            write_url_to_excel(url.get_nodes(), ws)
        else:
            ROW_INDEX += 1
        if len(urls) - 1 == i:
            COL_INDEX -= 1


if __name__ == "__main__":
    # 获取C001005所有角色
    sql = 'SELECT * FROM bdf2_role WHERE COMPANY_ID_=%s'
    args = ['C001005']
    roles = execute_query(sql, args)
    wb = xlwt.Workbook(encoding='utf-8')
    url_tree_list = []
    global ROW_INDEX, COL_INDEX
    for role in roles:
        # print 'id=' + val[0] + '&name=' + val[6]
        print '#######' + role[6] + '######'
        role_urls = get_role_urls(role[0])
        url_tree = view_url_tree(role_urls)
        # 生成Excel表格
        ws = wb.add_sheet(role[6].encode("utf-8"))  # 使用角色名做sheet的名称
        ws.col(0).width = 5000  # 设置行宽
        ws.col(1).width = 6500
        ws.col(2).width = 4000
        ws.col(3).width = 6000
        ROW_INDEX, COL_INDEX = 0, 0
        write_url_to_excel(url_tree, ws)
    try:
        wb.save('roleUrlView.xls')  # 切记不可xls文件打开的时候跑脚本，否则将会报错
    except IOError as e:
        print '保存文件时出错：', e
        print '请关闭Excel文件，然后重启程序'
